##############################################################################
#
#   SCOMHealth-Check.ps1
#
#   Script by: Jeremy D. Pavleck, Cool Guy & Raconteur
#
##############################################################################

$fatalErrorText = @"

  ___ _ _____ _   _      ___ ___ ___  ___  ___ 
 | __/_|_   _/_\ | |    | __| _ | _ \/ _ \| _ \
 | _/ _ \| |/ _ \| |__  | _||   |   | (_) |   /
 |_/_/ \_|_/_/ \_|____| |___|_|_|_|_\\___/|_|_\
                                               
"@

# Fatal Error Banner - Note: The font is called 'Calvin S'
$fatalErrorText2 = @"

┬┬┬┬┬  ╔═╗╔═╗╔╦╗╔═╗╦    ╔═╗╦═╗╦═╗╔═╗╦═╗  ┬┬┬┬┬
│││││  ╠╣ ╠═╣ ║ ╠═╣║    ║╣ ╠╦╝╠╦╝║ ║╠╦╝  │││││
ooooo  ╚  ╩ ╩ ╩ ╩ ╩╩═╝  ╚═╝╩╚═╩╚═╚═╝╩╚═  ooooo
"@

#Call SCOM powershell plugin and connect to Root Management Server
$SCOMMS = "SCOMMS01.Pavleck.Army"

If(Get-Module -Name OperationsManager ){
    Import-Module -Name OperationsManager -ErrorVariable errImport -ErrorAction SilentlyContinue
    If($errImport){
    Write-Host -ForegroundColor Red -BackgroundColor White -Object      
     }
    Write-Host -ForegroundColor Green -Object ""
}

Add-PSSnapin "Microsoft.EnterpriseManagement.OperationsManager.Client" -ErrorVariable errSnapin;
Set-Location "OperationsManagerMonitoring::" -ErrorVariable errSnapin;
new-managementGroupConnection -ConnectionString:$SCOMMS -ErrorVariable errSnapin;
New-PSDrive -Name: Monitoring -PSProvider: OperationsManagerMonitoring -Root: \ -ErrorAction SilentlyContinue -ErrorVariable Err
set-location $SCOMMS -ErrorVariable errSnapin;

# Create header for HTML Report
$Head = "<style>"
$Head +="BODY{background-color:#CCCCCC;font-family:Verdana,sans-serif; font-size: x-small;}"
$Head +="TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse; width: 100%;}"
$Head +="TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:green;color:white;padding: 5px; font-weight: bold;text-align:left;}"
$Head +="TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#F0F0F0; padding: 2px;}"
$Head +="</style>"

# Get status of Management Server Health and input them into report
write-host "Getting Management Health Server States" -ForegroundColor Yellow 
$ReportOutput = "To enable HTML view, click on `"This message was converted to plain text.`" and select `"Display as HTML`""
$ReportOutput += "<p><H2>Management Servers not in Healthy States</H2></p>"

$Count = Get-ManagementServer | where {$_.HealthState -ne "Success"} | Measure-Object

if($Count.Count -gt 0) { 
 $ReportOutput += Get-ManagementServer | where {$_.HealthState -ne "Success"} | select Name,HealthState,IsRootManagementServer,IsGateway | ConvertTo-HTML -fragment
} else { 
 $ReportOutput += "<p>All management servers are in healthy state.</p>"
} 

# Get status of Maintenance Mode for Root Management Server
write-host "Getting RMS Maintenance Mode" -ForegroundColor Yellow 
$RMS = Get-ManagementServer | where {$_.IsRootManagementServer -eq $True} 
$criteria = new-object Microsoft.EnterpriseManagement.Monitoring.MonitoringObjectGenericCriteria("InMaintenanceMode=1") 
$objectsInMM = (Get-ManagementGroupConnection).ManagementGroup.GetPartialMonitoringObjects($criteria) 
$is = "is not"
foreach ($MM in $objectsInMM){ 
 if($MM.Displayname -eq $RMS.Name){ 
   $is = "is"
 } 
} 

$ReportOutput += "<h2>RMS in Maintenance Mode</h2><p>"+ $RMS.Name +" "+$is+" in maintenance mode</p>"

# Get Agent Health Status and put none healthy ones into report
write-host "Getting Agent Health Status" -ForegroundColor Yellow 
$criteria = new-object Microsoft.EnterpriseManagement.Monitoring.MonitoringObjectGenericCriteria("InMaintenanceMode=1")
$objectsInMM = (Get-ManagementGroupConnection).ManagementGroup.GetPartialMonitoringObjects($criteria)
$ObjectsFound = $objectsInMM | select-object DisplayName, @{name="Object Type";expression={foreach-object {$_.GetLeastDerivedNonAbstractMonitoringClass().DisplayName}}},@{name="StartTime";expression={foreach-object {$_.GetMaintenanceWindow().StartTime.ToLocalTime()}}},@{name="EndTime";expression={foreach-object {$_.GetMaintenanceWindow().ScheduledEndTime.ToLocalTime()}}},@{name="Path";expression={foreach-object {$_.Path}}},@{name="User";expression={foreach-object {$_.GetMaintenanceWindow().User}}},@{name="Reason";expression={foreach-object {$_.GetMaintenanceWindow().Reason}}},@{name="Comment";expression={foreach-object {$_.GetMaintenanceWindow().Comment}}}

$ReportOutput += "<h2>Agents where Health State is not Green</h2>"
#$ReportOutput += Get-Agent | where {$_.HealthState -ne "Success"} | Sort-Object HealthState -descending | select Name,HealthState | ConvertTo-HTML -fragment

$Agents = Get-Agent | where {$_.HealthState -ne "Success"} | Sort-Object HealthState -descending | select Name,HealthState

$AgentTable = New-Object System.Data.DataTable "$AvailableTable"
$AgentTable.Columns.Add((New-Object System.Data.DataColumn Name,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn HealthState,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MM,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMUser,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMReason,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMComment,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMEndTime,([string])))

foreach ($Agent in $Agents)
    {
        $FoundObject = $null
	$MaintenanceModeUser = $null
	$MaintenanceModeComment = $null
	$MaintenanceModeReason = $null
	$MaintenanceModeEndTime = $null
        $FoundObject = 0
        $FoundObject = $objectsFound | ? {$_.DisplayName -match $Agent.Name -or $_.Path -match $Agent.Name}
        if ($FoundObject -eq $null)
            {
                $MaintenanceMode = "No"
                $MaintenanceObjectCount = 0
            }
        else
            {
                $MaintenanceMode = "Yes"
                $MaintenanceObjectCount = $FoundObject.Count
		$MaintenanceModeUser = (($FoundObject | Select User)[0]).User
		$MaintenanceModeReason = (($FoundObject | Select Reason)[0]).Reason
		$MaintenanceModeComment = (($FoundObject | Select Comment)[0]).Comment
		$MaintenanceModeEndTime = ((($FoundObject | Select EndTime)[0]).EndTime).ToString()
            }
        $NewRow = $AgentTable.NewRow()
        $NewRow.Name = ($Agent.Name).ToString()
        $NewRow.HealthState = ($Agent.HealthState).ToString()
        $NewRow.MM = $MaintenanceMode
	$NewRow.MMUser = $MaintenanceModeUser
        $NewRow.MMReason = $MaintenanceModeReason
        $NewRow.MMComment = $MaintenanceModeComment
        $NewRow.MMEndTime = $MaintenanceModeEndTime
        $AgentTable.Rows.Add($NewRow)
    }
    
$ReportOutput += $AgentTable | Sort-Object MM | Select Name, HealthState, MM, MMUser, MMReason, MMComment, MMEndTime | ConvertTo-HTML -fragment

# Also put into the report agents that have a state of "Not Monitored" and/or "Unavailable" - Grey Agents
$ReportOutput += "<h2>Agents where the Monitoring Class is not available</h2>"
$AgentMonitoringClass = get-monitoringclass -name "Microsoft.SystemCenter.Agent"
$NotAvailable = Get-MonitoringObject -monitoringclass:$AgentMonitoringClass | where {$_.IsAvailable -eq $false} | select DisplayName
$AvailableTable = New-Object System.Data.DataTable "$AvailableTable"
$AvailableTable.Columns.Add((New-Object System.Data.DataColumn DisplayName,([string])))
$AvailableTable.Columns.Add((New-Object System.Data.DataColumn MM,([string])))
$AvailableTable.Columns.Add((New-Object System.Data.DataColumn MMUser,([string])))
$AvailableTable.Columns.Add((New-Object System.Data.DataColumn MMReason,([string])))
$AvailableTable.Columns.Add((New-Object System.Data.DataColumn MMComment,([string])))
$AvailableTable.Columns.Add((New-Object System.Data.DataColumn MMEndTime,([string])))
foreach ($NotAvailableAgent in $NotAvailable)
    {
        $FoundObject = $null
	$MaintenanceModeUser = $null
	$MaintenanceModeComment = $null
	$MaintenanceModeReason = $null
	$MaintenanceModeEndTime = $null
        $FoundObject = 0
        $FoundObject = $objectsFound | ? {$_.DisplayName -match $NotAvailableAgent.DisplayName -or $_.Path -match $NotAvailableAgent.DisplayName}
        if ($FoundObject -eq $null)
            {
                $MaintenanceMode = "No"
                $MaintenanceObjectCount = 0
            }
        else
            {
                $MaintenanceMode = "Yes"
                $MaintenanceObjectCount = $FoundObject.Count
		$MaintenanceModeUser = (($FoundObject | Select User)[0]).User
		$MaintenanceModeReason = (($FoundObject | Select Reason)[0]).Reason
		$MaintenanceModeComment = (($FoundObject | Select Comment)[0]).Comment
		$MaintenanceModeEndTime = ((($FoundObject | Select EndTime)[0]).EndTime).ToString()
            }
        $NewRow = $AvailableTable.NewRow()
        $NewRow.DisplayName = ($NotAvailableAgent.DisplayName).ToString()
        $NewRow.MM = $MaintenanceMode
	$NewRow.MMUser = $MaintenanceModeUser
        $NewRow.MMReason = $MaintenanceModeReason
        $NewRow.MMComment = $MaintenanceModeComment
        $NewRow.MMEndTime = $MaintenanceModeEndTime
        $AvailableTable.Rows.Add($NewRow)
    }
$ReportOutput += $AvailableTable | Sort-Object MM | Select DisplayName, MM, MMUser, MMReason, MMComment, MMEndTime | ConvertTo-HTML -fragment

# Get Alerts specific to Management Servers and put them in the report
write-host "Getting Management Server Alerts" -ForegroundColor Yellow 
$ReportOutput += "<h2>Management Server Alerts</h2>"
$ManagementServers = Get-ManagementServer
foreach ($ManagementServer in $ManagementServers){ 
 $ReportOutput += "<h3>Alerts on " + $ManagementServer.ComputerName + "</h3>"
 $ReportOutput += get-alert -Criteria ("NetbiosComputerName = '" + $ManagementServer.ComputerName + "'") | where {$_.ResolutionState -ne '255' -and $_.MonitoringObjectFullName -Match 'Microsoft.SystemCenter'} | select TimeRaised,Name,Description,Severity | ConvertTo-HTML -fragment
}

write-host "Getting all alerts"
$Alerts = Get-Alert -Criteria 'ResolutionState < "255"'

# Get alerts for last 24 hours
write-host "Getting alerts for last 24 hours"
$ReportOutput += "<h2>Top 10 Alerts With Same Name - 24 hours</h2>"
$ReportOutput += $Alerts | where {$_.LastModified -le (Get-Date).addhours(-24)} | Group-Object Name | Sort-object Count -desc | select-Object -first 10 Count, Name | ConvertTo-HTML -fragment

$ReportOutput += "<h2>Top 10 Repeating Alerts - 24 hours</h2>"
$ReportOutput += $Alerts | where {$_.LastModified -le (Get-Date).addhours(-24)} | Sort-Object -desc RepeatCount | select-Object -first 10 RepeatCount, Name, MonitoringObjectPath, Description | ConvertTo-HTML -fragment

# Get the Top 10 Unresolved alerts still in console and put them into report
write-host "Getting Top 10 Unresolved Alerts With Same Name - All Time" -ForegroundColor Yellow 
$ReportOutput += "<h2>Top 10 Unresolved Alerts</h2>"
$ReportOutput += $Alerts  | Group-Object Name | Sort-object Count -desc | select-Object -first 10 Count, Name | ConvertTo-HTML -fragment

# Get the Top 10 Repeating Alerts and put them into report
write-host "Getting Top 10 Repeating Alerts - All Time" -ForegroundColor Yellow 
$ReportOutput += "<h2>Top 10 Repeating Alerts</h2>"
$ReportOutput += $Alerts | Sort -desc RepeatCount | select-object �first 10 Name, RepeatCount, MonitoringObjectPath, Description | ConvertTo-HTML -fragment

# Get list of agents still in Pending State and put them into report
write-host "Getting Agents in Pending State" -ForegroundColor Yellow 
$ReportOutput += "<h2>Agents in Pending State</h2>"
$ReportOutput += Get-AgentPendingAction | sort AgentPendingActionType | select AgentName,ManagementServerName,AgentPendingActionType | ConvertTo-HTML -fragment

# Find Overrides that have been stored in the default management pack and put them into the report
write-host "Getting Overrides in Default Management Pack" -ForegroundColor Yellow 
$ReportOutput += "<h2>Overrides in Default Management Pack</h2>"
$OverrideCount = Get-ManagementPack | where {$_.DisplayName -match "Default Management Pack"} | get-override | measure-object
if($OverrideCount.Count -gt 2){ 
 foreach ($monitor in Get-ManagementPack | where {$_.DisplayName -match "Default Management Pack"} | get-override | where {$_.monitor}) { 
  $ReportOutput += get-monitor | where {$_.Id -eq $monitor.monitor.id} | select-object DisplayName,Description | ConvertTo-HTML -fragment
  $ReportOutput += "<br />"
 } 
 foreach ($rule in Get-ManagementPack | where {$_.DisplayName -match "Default Management Pack"} | get-override | where {$_.rule}) { 
  $ReportOutput += get-rule | where {$_.Id -eq $rule.rule.id} | select-object DisplayName,Description | ConvertTo-HTML -fragment
  $ReportOutput += "<br />"
 } 
} else { 
 $ReportOutput += "<p>There are no unexpected overrides in the Default Management Pack</p>"
}

# List number of MM Objects that have been found.
$ReportOutput += "<h2>Count of objects in Maintanence Mode</h2>"
$CountTable = New-Object System.Data.DataTable "$CountTable"
$CountTable.Columns.Add((New-Object System.Data.DataColumn ObjectCount,([int])))
$NewRow = $CountTable.NewRow()
$NewRow.ObjectCount = $ObjectsFound.Count
$CountTable.Rows.Add($NewRow)
$ReportOutput += $CountTable |Select @{n='ObjectCount';e={$_.ObjectCount}}, @{Name="Date";Expression={Get-Date -Format F}} | ConvertTo-Html -fragment

# List Management Packs updated in last 24 hours
$ReportOutput += "<h2>Management Packs Updated</h2>"
$MPDates = (Get-Date).adddays(-1)
$ReportOutput += Get-ManagementPack | Where {$_.LastModified -gt $MPDates} | Select-Object DisplayName, LastModified | Sort LastModified | ConvertTo-Html -fragment

# Take all $ReportOutput and combine it with $Body to create completed HTML output
$Body = ConvertTo-HTML -head $Head -body "$ReportOutput"

#$Body | Out-File C:\users\adm.j.rydstrand\desktop\HealthCheck-11-14-2012.html

# Setup and send output as email message.
$smtpServer = "outgoing.server"
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "SCOM.REPORTS.DONOTREPLY@microsoft.com"
$msg.To.Add("jason.rydstrand@microsoft.com")
$msg.Subject = "SCOM Daily Healthcheck Report"
$msg.IsBodyHtml = 1
$msg.Body = $Body
$smtp.Send($msg)
