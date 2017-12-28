Param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]$Guid,
    [String]$Environment,
    [String]$ReportDate
)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"
$ToolsFolder = $Dir + "\AX-Tools"
$ReportFolder = $Dir + "\Reports\AX-Report\$Environment"
$LogFolder = $Dir + "\Logs\AX-Report\$Environment"
$ReportDate = $(Get-Date (Get-Date).AddDays(-1) -format MMddyyyy) #Get-Date -f MMddyyHHmm
#
Import-Module $ModuleFolder\AX-Database.psm1 -DisableNameChecking
#Import-Module $ModuleFolder\AX-HTMLReport.psm1 -DisableNameChecking
Import-Module $ModuleFolder\AX-ReportFunc.psm1 -DisableNameChecking

$Footer = "AX Report v{3} run {0} by {1}\{2}" -f (Get-Date),$env:UserDomain,$env:UserName,'2.0'
$ReportName = "AX Daily Report"

function Run-Report
{
    Run-ReportDP
    Create-ReportSummary
    Create-Report
    Save-ReportFile

}

function Create-ReportSummary
{
    $Script:AxSummary = @()
    #AxServices
    if($Script:ReportDP.AxServices.Count -eq ($Script:ReportDP.AxServices | Where {$_.Status -match 'Running'}).Count) { 
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "AOS Services"; Status = "Ok. All Services Running."; RowColor = 'Green' }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "AOS Services"; Status = "AOS Services Failure Found."; RowColor = 'Red' }
    }

    #MRP Runtime
    if($Script:ReportDP.AxMRPLogs) {
        switch -wildcard ($Script:ReportDP.AxMRPLogs) {
            {$($Script:ReportDP.AxMRPLogs.TotalTime) -eq 0} {$Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "MRP Failed or Cancelled."; RowColor = 'Red' }}
            {($($Script:ReportDP.AxMRPLogs.TotalTime) -gt 0) -and ($($Script:ReportDP.AxMRPLogs.TotalTime) -le 45)} {$Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "$($Script:ReportDP.AxMRPLogs.TotalTime) minutes."; RowColor = 'Green' }}
            {($($Script:ReportDP.AxMRPLogs.TotalTime) -gt 45) -and ($($Script:ReportDP.AxMRPLogs.TotalTime) -le 60)} {$Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "$($Script:ReportDP.AxMRPLogs.TotalTime) minutes."; RowColor = 'Yellow' }}
            Default {$Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "$($Script:ReportDP.AxMRPLogs.TotalTime) minutes."; RowColor = 'Red' }}
        }
    }
    else {
         $Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "MRP Long Run or Failed."; RowColor = 'Red'; }
    }

    #AxBatchJobs
    if($Script:ReportDP.AxBatchJobs.Count -eq 0) { 
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Batch Jobs"; Status = "Ok."; RowColor = 'Green' }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Batch Jobs"; Status = "Errors Found."; RowColor = 'Red' }
    }
    if($Script:ReportDP.AxLongBatchJobs.Count -eq 0) {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Long Batch Jobs (>15min)"; Status = "Ok."; RowColor = 'Green' }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Long Batch Jobs (>15min)"; Status = "$($Script:ReportDP.AxLongBatchJobs.Count) Jobs Found."; RowColor = 'Red' }
    }

    #AxRetailJobs
    if($Script:ReportDP.AxCDXJobs.Count -eq 0) { 
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Retail Jobs"; Status = "Ok."; RowColor = 'Green' }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Retail Jobs"; Status = "Errors Found."; RowColor = 'Red' }
    }

    #PerfmonData Color-Set
    $Green = '(($this.Counter -like "CPU Time %" -and $this.Max -le 60) -or ($this.Counter -like "Available GBytes" -and $this.Min -ge 8) -or ($this.Counter -like "Paging File %" -and $this.Max -le 35) -or ($this.Counter -like "*Buffer cache hit ratio" -and $this.Min -ge 95) -or ($this.Counter -like "*Page life expectancy" -and $this.Min -ge 6000))'
    $Yellow = '(($this.Counter -like "CPU Time %" -and $this.Max -gt 60 -and $this.Max -lt 80) -or ($this.Counter -like "Available GBytes" -and $this.Max -gt 4 -and $this.Max -lt 8) -or ($this.Counter -like "Paging File %" -and $this.Max -gt 35 -and $this.Max -lt 50) -or ($this.Counter -like "*Buffer cache hit ratio" -and $this.Min -gt 90 -and $this.Min -lt 95) -or ($this.Counter -like "*Page life expectancy" -and $this.Min -gt 1200 -and $this.Min -lt 6000))'    
    $Red = '(($this.Counter -like "CPU Time %" -and $this.Max -ge 80) -or ($this.Counter -like "Available GBytes" -and $this.Max -le 4) -or ($this.Counter -like "Paging File %" -and $this.Max -ge 50) -or ($this.Counter -like "*Buffer cache hit ratio" -and $this.Min -le 90) -or ($this.Counter -like "*Page life expectancy" -and $this.Min -le 1200))'
    
    #REMOVING INSTANCES NOT RUNNING
    $PermonDataLogsTmp = $Script:ReportDP.PermonDataLogs | Where {$_.ServerType -notmatch 'SQL' -or $_.CounterType -like 'SRV' }
    $PermonDataLogsTmp += $Script:ReportDP.PermonDataLogs | Where {(($_.Max -ne 0) -or ($_.Min -ne 0)) -and ($_.CounterType -notmatch 'SRV') -and ($_.ServerType -match 'SQL')}
    #$AXPerfmonCLR = Set-TableRowColor $PermonDataLogsTmp -Red $Red -Yellow $Yellow -Green $Green
    $Script:ReportDP | Add-Member -Name AxPerfmonCLR -Value $(Set-TableRowColor $PermonDataLogsTmp -Red $Red -Yellow $Yellow -Green $Green) -MemberType NoteProperty


    #PerfmonData
    if(((($Script:ReportDP.AxPerfmonCLR | Group RowColor | Where Name -like 'Green').Count) + (($Script:ReportDP.AxPerfmonCLR | Group RowColor | Where Name -like 'Green').Count)) -eq $Script:ReportDP.AxPerfmonCLR.Count) {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Performance Monitor"; Status = "$(($Script:ReportDP.AxPerfmonCLR | Group RowColor | Where Name -like 'Green').Count) Alerts."; RowColor = 'Green' }
    }
    elseif (((($Script:ReportDP.AxPerfmonCLR | Group RowColor | Where Name -like 'Yellow').Count) -gt 0) -and ((($Script:ReportDP.AxPerfmonCLR | Group RowColor | Where Name -like 'Red').Count) -eq 0)) {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Performance Monitor"; Status = "$((($Script:ReportDP.AxPerfmonCLR | Group RowColor | Where Name -like 'Yellow').Count)) Warnings."; RowColor = 'Yellow' }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Performance Monitor"; Status = "$((($Script:ReportDP.AxPerfmonCLR | Group RowColor | Where Name -like 'Red').Count)) Criticals and $((($Script:ReportDP.AxPerfmonCLR | Group RowColor | Where Name -like 'Yellow').Count)) Warnings."; RowColor = 'Red' }
    }
    
    #EventLogs
    if(($($Script:ReportDP.AxEventLogs | Group Id | Where Name -like 1000).Count -gt 0) -or ($($Script:ReportDP.AxEventLogs | Group Id | Where Name -like 1002).Count -gt 0)) { 
        $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
        $Query = "SELECT  A.ServerName as [Server],
                                ServerType = Case B.ServerType
		                        WHEN 'AOS' then 'AOS Server'
		                        WHEN 'RDP' then 'RDP Server'
		                        END, 
		                        Application = CASE SUBSTRING(A.MESSAGE,(CHARINDEX('Ax32',A.MESSAGE)), ((CHARINDEX('.exe',A.MESSAGE)) - (CHARINDEX('Ax32',A.MESSAGE))))
		                        WHEN 'Ax32' then 'client'
		                        WHEN 'Ax32Serv' then 'server'
		                        END, 
		                        Type = Case A.EVENTID
		                        WHEN '1000' then 'crash(es)'
		                        WHEN '1002' then 'hang(s)'
		                        END, COUNT(1) as Count
                        FROM AxReport_EventLogs A
                        CROSS JOIN AXTools_Servers B
                        WHERE A.Guid = '$Guid' 
		                        AND (A.EVENTID = '1000' or A.EVENTID = '1002') 
		                        AND A.ENTRYTYPE = 'ERROR' 
		                        AND A.MESSAGE LIKE '%AX32%' 
		                        AND A.SERVERNAME = B.SERVERNAME
                        GROUP BY A.ServerName, B.ServerType, SUBSTRING(A.MESSAGE,(CHARINDEX('Ax32',A.MESSAGE)), ((CHARINDEX('.exe',A.MESSAGE)) - (CHARINDEX('Ax32',A.MESSAGE)))), A.EVENTID
                        ORDER BY 1 DESC, COUNT DESC"
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $Cmd
        $AxCrash = New-Object System.Data.DataSet
        $Adapter.Fill($AxCrash)
        $AxCrashLog = $AxCrash.Tables[0] | Sort Server, Type
        $Script:AxSummary += New-Object PSObject -Property @{
                            Name = "Event Logs";
                            Status = "$(($AxCrashLog | Where {$_.Type -like 'Crash(es)'} | Measure-Object Count -Sum).Sum) Crash(es) and $(($AxCrashLog | Where {$_.Type -like 'Hang(s)'} | Measure-Object Count -Sum).Sum) Hang(s)"; 
                            RowColor = if((($AxCrashLog | Where {$_.Type -like 'Crash(es)'}).Count -gt 0) -and (($AxCrashLog | Where {$_.Application -like 'server'}).Count -gt 0)) {'Red'} else {'Yellow'}
                          }
        foreach($Crash in $AxCrashLog | Group Server) {
            $TmpSummary = @()
            $i = 1
            foreach($CrashRpt in $Crash.Group) {
                if($i -eq 1) {
                    $TmpSummary += "$($CrashRpt.Count) $($CrashRpt.Application) $($CrashRpt.Type)"
                }
                else {
                    $TmpSummary += "and $($CrashRpt.Count) $($CrashRpt.Application) $($CrashRpt.Type)"
                }
                $i++
            }
            $Script:AxSummary += New-Object PSObject -Property @{ 
                            Name = "`t"; 
                            Status = " -› $($Crash.Name) - $TmpSummary"; 
                            RowColor = if($TmpSummary -match ("server crash")) {'Red'} else {'Yellow'} }
        }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Event Logs"; Status = "Ok."; RowColor = 'Green' }
    }

    #SQLErrorLogs
    if(($Script:ReportDP.SQLErrorLogs.Log | Where {$_ -like 'SQL Server is starting*' }) -or 
                ($Script:ReportDP.SQLErrorLogs.Log | Where {$_ -like 'Starting up database*' }) -or 
                    ($Script:ReportDP.SQLErrorLogs.Log | Where {$_ -like 'Recovery of database*' })) { 
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "SQL Errors"; Status = "SQL Restarted."; RowColor = 'Red' }
    }
    elseif($Script:ReportDP.SQLErrorLogs.Process.Contains('Server')) {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "SQL Errors"; Status = "SQL Server Failure Found."; RowColor = 'Yellow' }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "SQL Errors"; Status = "Ok."; RowColor = 'Green' }
    }
    #Reporting Services
    if((($Script:ReportDP.SSRSErrorLogs | Where { $_.Status -notlike 'rsReportParameterValueNotSet'}).Count) -eq 0) { 
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "SSRS Errors"; Status = "Ok."; RowColor = 'Green' }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "SSRS Errors"; Status = "$((($Script:ReportDP.SSRSErrorLogs | Where { $_.Status -notlike 'rsReportParameterValueNotSet'}).Count)) Issues Found."; RowColor = 'Yellow' }
    }
}

function Create-Report
{
    #Start Report
    $Script:AXReport = @()
    $Script:AXReport += Get-HtmlOpen -TitleText ($ReportName)
    $Script:AXReport += Get-HtmlContentOpen -HeaderText "AX Daily Report"

    ###First
    ##Summary Report
    $Script:AXReport += Get-HtmlColumn1of2
    $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "Summary Information"
    $Script:AXReport += Get-HtmlContentTable($Script:AxSummary | Select Name, Status, RowColor)
    $Script:AXReport += Get-HtmlContentClose
    $Script:AXReport += Get-HtmlColumnClose
    #
    #AX Services Status
    $Script:AXReport += Get-HtmlColumn2of2
    $Green = '$this.Status -match "Running"'
    $Red = '$this.Status -match "Stopped"'
    $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "AX Services Status"
    $Script:AXReport += Get-HtmlContentTable(Set-TableRowColor $Script:ReportDP.AxServices -Red $Red -Green $Green)
    $Script:AXReport += Get-HtmlContentClose
    $Script:AXReport += Get-HtmlColumnClose
    $Script:AXReport += Get-HtmlContentClose
    #

    #MRP Status
    if($Script:ReportDP.AxMRPLogs)
    {
        $Green = '$this.TotalTime -gt 0 -and $this.TotalTime -le 45'
        $Yellow = '$this.TotalTime -gt 45 -and $this.TotalTime -le 60'
        $Red = '$this.TotalTime -eq 0 -or $this.TotalTime -gt 60'
        $AxMRPColor = Set-TableRowColor $Script:ReportDP.AxMRPLogs -Green $Green -Yellow $Yellow -Red $Red
        $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "MRP Run Status"
        $Script:AXReport += Get-HtmlContentTable($AxMRPColor)
        $Script:AXReport += Get-HtmlContentClose
    }

    #Second
    ##Perfmon Logs
    $Script:AXReport += Get-HtmlContentOpen
    $Script:AXReport += Get-HtmlColumn1of2
        $Script:AXReport += Get-HtmlContentOpen -HeaderText "Performance Monitor by Server [Total - $($Script:ReportDP.AxPerfmonCLR.Count)]" -BackgroundShade 1
        foreach ($Type in ($Script:ReportDP.AxPerfmonCLR | Group ServerType | Sort Name ) ) {
            $Script:AXReport += Get-HtmlContentOpen -HeaderText ($Type.Name + " Servers") -IsHidden -BackgroundShade 1
            foreach ($Group in ($Script:ReportDP.AxPerfmonCLR | Where-Object {$_.ServerType -match $Type.Name} | Group ServerName | Sort Name ) ) {
                $Script:AXReport += Get-HtmlContentOpen -HeaderText ($Group.Name) -IsHidden -BackgroundShade 1
                $Script:AXReport += Get-HtmlContentTable ($Group.Group | Select Counter, Max, Min, Avg, RowColor)
                $Script:AXReport += Get-HtmlContentClose
            }
            $Script:AXReport += Get-HtmlContentClose
        }
        $Script:AXReport += Get-HtmlContentClose
    $Script:AXReport += Get-HtmlColumnClose
    #
    $Script:AXReport += Get-HtmlColumn2of2
    $Script:AXReport += Get-HtmlContentOpen -HeaderText "Performance Monitor Alerts by Threshold" -BackgroundShade 1
        $AxPerfMonGrp = $Script:ReportDP.AxPerfmonCLR | Where {($_.RowColor -like 'Red') -or ($_.RowColor -like 'Yellow')} | Group RowColor | Sort Name
        $Script:AXReport += Get-HtmlContentTable ($AxPerfMonGrp.Group | Select ServerName, Counter, Max, Min, Avg, RowColor)
    $Script:AXReport += Get-HtmlContentClose
    $Script:AXReport += Get-HtmlContentClose
    $Script:AXReport += Get-HtmlColumnClose
    
    #Third
    #Event Logs Graphs
    $PieChartObject1 = New-HTMLPieChartObject
    $PieChartObject1.Title = " "
    $PieChartObject1.Size.Height = 300
    $PieChartObject1.Size.Width = 300
    $PieChartObject1.ChartStyle.ExplodeMaxValue = $true
    $PieChartObject2 = New-HTMLPieChartObject
    $PieChartObject2.Title = " "
    $PieChartObject2.Size.Height = 300
    $PieChartObject2.Size.Width = 300
    $PieChartObject2.ChartStyle.ExplodeMaxValue = $true    				
    
    $Script:AXReport += Get-HtmlContentOpen
    $Script:AXReport += Get-HtmlColumn1of2
    $Script:AXReport += Get-HtmlContentOpen -HeaderText "Event Logs by Server (Top 5)"
    $Script:AXReport += New-HTMLPieChart -PieChartObject $PieChartObject2 -PieChartData ($Script:ReportDP.AxEventLogsChart | Group ServerName | Sort Count -Descending | Select -First 5)
    $Script:AXReport += Get-HtmlContentTable ($Script:ReportDP.AxEventLogsChart | Group ServerName | Select Name, Count | Sort Count -Descending | Select -First 5)
    $Script:AXReport += Get-HtmlContentClose
    $Script:AXReport += Get-HtmlColumnClose

    $Script:AXReport += Get-HtmlColumn2of2
    $Script:AXReport += Get-HtmlContentOpen -HeaderText "Event Logs by Server" -BackgroundShade 1
    foreach ($Type in ($Script:ReportDP.AxEventLogsChart | Group ServerType | Sort Name ) ) {
        $Script:AXReport += Get-HtmlContentOpen -HeaderText ($Type.Name + " Servers") -IsHidden -BackgroundShade 1
        foreach ($Group in ($Script:ReportDP.AxEventLogsChart | Where-Object {$_.ServerType -match $Type.Name} | Group ServerName | Sort Name ) ) {
            $Script:AXReport += Get-HtmlContentOpen -HeaderText ($Group.Name) -IsHidden -BackgroundShade 1
            $Script:AXReport += Get-HtmlContentTable ($Script:ReportDP.AxEventLogs | Where {$_.ServerName -match $Group.Name} | Select LogName, Type, Id, Source, Count | Sort Count -Descending)
            $Script:AXReport += Get-HtmlContentClose
        }
        $Script:AXReport += Get-HtmlContentClose
    }
    $Script:AXReport += Get-HtmlContentClose   
    #
    $Script:AXReport += Get-HtmlColumnClose
    $Script:AXReport += Get-HtmlContentClose

    #Batch Jobs Errors
    $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "AX Batch Jobs Errors [Total - $($Script:ReportDP.AxBatchJobs.Count)]"
    $Script:AXReport += Get-HtmlContentTable ($Script:ReportDP.AxBatchJobs)
    $Script:AXReport += Get-HtmlContentClose
    #
    if($Script:ReportDP.SSRSErrorLogs.Count -gt 0) {
        $PieChartObject3 = New-HTMLPieChartObject
        $PieChartObject3.Title = " "
        $PieChartObject3.Size.Height = 300
        $PieChartObject3.Size.Width = 300
        $PieChartObject3.ChartStyle.ExplodeMaxValue = $true

        $Script:AXReport += Get-HtmlContentOpen
        $Script:AXReport += Get-HtmlColumn1of2
        $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "SSRS Error Logs [Total - $($Script:ReportDP.SSRSErrorLogs.Count)]"
        $Script:AXReport += Get-HtmlContentTable(Set-TableRowColor($Script:ReportDP.SSRSErrorLogs | Select Instance, Message, Report, Count) -Alternating)
        $Script:AXReport += Get-HtmlContentClose
        $Script:AXReport += Get-HtmlColumnClose
        #
        $Script:AXReport += Get-HtmlColumn2of2
        $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "SSRS Errors by User (Top 5)"
        $Script:AXReport += New-HTMLPieChart -PieChartObject $PieChartObject3 -PieChartData ($Script:ReportDP.SSRSUsersReport | Sort Count -Descending | Select -First 5)
        $Script:AXReport += Get-HtmlContentTable($Script:ReportDP.SSRSUsersReport | Select User, Count | Sort Count -Descending | Select -First 5)
        $Script:AXReport += Get-HtmlContentClose
        $Script:AXReport += Get-HtmlColumnClose
        $Script:AXReport += Get-HtmlContentClose
    }

    if($Script:ReportDP.AxCDXJobs.Count -gt 0) {
        #CDX Jobs Errors
        $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "CDX Jobs Errors [Total - $($Script:ReportDP.AxCDXJobs.Count)]" 
        $Script:AXReport += Get-HtmlContentTable (Set-TableRowColor $Script:ReportDP.AxCDXJobs -Alternating)
        $Script:AXReport += Get-HtmlContentClose
    }

    ##SQL Error Logs
    $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "SQL Server Error Logs [Total - $($Script:ReportDP.SQLErrorLogs.Count)]" 
    $Script:AXReport += Get-HtmlContentTable ($Script:ReportDP.SQLErrorLogs) 
    $Script:AXReport += Get-HtmlContentClose

    if($Script:ReportDP.SSRSWeek.Count -gt 0) {
        $PieChartObject4 = New-HTMLPieChartObject
        $PieChartObject4.Title = " "
        $PieChartObject4.Size.Height = 400
        $PieChartObject4.Size.Width = 400
        $PieChartObject4.ChartStyle.ExplodeMaxValue = $true    			
    
        $Script:AXReport += Get-HtmlContentOpen
        $Script:AXReport += Get-HtmlColumn1of2
        $Script:AXReport += Get-HtmlContentOpen -HeaderText "SSRS Errors 7 Days"
        $Script:AXReport += New-HTMLPieChart -PieChartObject $PieChartObject4 -PieChartData ($Script:ReportDP.SSRSWeek | Sort Date)
        $Script:AXReport += Get-HtmlContentTable ($Script:ReportDP.SSRSWeek | Select Date, Count | Sort Date -Descending)
        $Script:AXReport += Get-HtmlContentClose
        $Script:AXReport += Get-HtmlColumnClose

        $Script:AXReport += Get-HtmlColumn2of2

        $Script:AXReport += Get-HtmlColumnClose
        $Script:AXReport += Get-HtmlContentClose
    }

    #Close Report
    $Script:AXReport += Get-HtmlContentClose
    $Script:AXReport += Get-HtmlClose -Footer $Footer
}

function Save-ReportFile
{
    ##Add Summary Email Info
    $Script:AxSummary += New-Object PSObject -Property @{ Name = '**Please see the attached report for details.'; Status = ''; RowColor = 'None' }
    $AXREmail = @()
    $AXREmail += Get-SummaryOpen -TitleText ($ReportName)
    $AXREmail += Get-HtmlContentOpen -HeaderText "Summary Information"
    $AXREmail += Get-HtmlContentTable($Script:AxSummary | Select Name, Status, RowColor)
    $AXREmail += Get-HtmlContentClose
    $AXREmail += Get-SummaryClose -Footer $Footer
    #Save Summary
    $AXReportPath = Join-Path $ReportFolder ("AXReport-$ReportDate-Summary" + ".html")
    $AXREmail | Set-Content -Path $AXReportPath -Force
    #Save Report
    $AXReportPath = Join-Path $ReportFolder ("AXReport-$ReportDate" + ".mht")
    $Script:AXReport | Set-Content -Path $AXReportPath -Force
}

function Run-ReportDP
{

#function Get-AxServices
#{
    $Script:ReportDP = New-Object -TypeName System.Object   
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT ServerName, Service AS ServiceName, DisplayName, Status FROM AXReport_AxServices WHERE Guid = '$Guid'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxServices = New-Object System.Data.DataSet
    $Adapter.Fill($AxServices) | Out-Null
    $Script:ReportDP | Add-Member -Name AxServices -Value $($AxServices.Tables[0] | Select ServerName, ServiceName, DisplayName, Status) -MemberType NoteProperty
    #return $AxServices.Tables[0] | Select ServerName, ServiceName, DisplayName, Status
#}

#function Get-AxBatchJobs
#{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT HISTORYCAPTION AS [History Caption],JOBCAPTION AS [Job Caption],Status,ServerID AS Server,STARTDATETIMECST AS [Start Time(CST)],ENDDATETIMECST AS [End Time(CST)],EXECUTEDBY AS [User], LOG AS Log
                FROM AXReport_AxBatchJobs 
                WHERE Guid = '$Guid'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxBatchJobs = New-Object System.Data.DataSet
    $Adapter.Fill($AxBatchJobs)
    $Script:ReportDP | Add-Member -Name AxBatchJobs -Value $($AxBatchJobs.Tables[0] | Select 'History Caption', 'Job Caption', 'Status', @{n='Server';e={($_.SERVER -replace '01@','').Trim()}}, 'Start Time(CST)', 'End Time(CST)', 'User', 'Log') -MemberType NoteProperty
    #return $AxBatchJobs.Tables[0] | Select 'History Caption', 'Job Caption', 'Status', @{n='Server';e={($_.SERVER -replace '01@','').Trim()}}, 'Start Time(CST)', 'End Time(CST)', 'User', 'Log'
#}

#function Get-AxLongBatchJobs
#{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT Job, Count, Status, Duration, EXECUTEDBY AS [User], ServerID AS [Server]
                FROM AXReport_AxLongBatchJobs 
                WHERE Guid = '$Guid'
                ORDER BY Duration DESC"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxLongBatchJobs = New-Object System.Data.DataSet
    $Adapter.Fill($AxLongBatchJobs)
    $Script:ReportDP | Add-Member -Name AxLongBatchJobs -Value $($AxLongBatchJobs.Tables[0] | Select 'Job', 'Count', 'Status', 'Duration', 'User', 'Server') -MemberType NoteProperty
    #return $AxLongBatchJobs.Tables[0] | Select 'Job', 'Count', 'Status', 'Duration', 'User', 'Server'
#}

#function Get-AxRetailJobs
#{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT JobID, STATUSDOWNLOADSESSIONDATASTORE AS [Download Status], Message, DateRequested, DateDownloaded, DateApplied, ROWSAFFECTED as [Rows], DATAFILEOUTPUTPATH as [Path], STATUSDOWNLOADSESSION as [Session Status], DATABASE_ as [Database], Name
                FROM AXReport_AxRetailJobs 
                WHERE Guid = '$Guid'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxCDXJobs = New-Object System.Data.DataSet
    $Adapter.Fill($AxCDXJobs)
    $Script:ReportDP | Add-Member -Name AxCDXJobs -Value $($AxCDXJobs.Tables[0] | Select JobID, 'Download Status', Message, DateRequested, DateDownloaded, DateApplied, Rows, Path, 'Session Status', Database, Name) -MemberType NoteProperty
    #return $AxCDXJobs.Tables[0] | Select JobID, 'Download Status', Message, DateRequested, DateDownloaded, DateApplied, Rows, Path, 'Session Status', Database, Name

#}

#function Get-AxEventLogs
#{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT A.ServerName, B.ServerType, A.EntryType as Type, A.EventID as ID, A.Source
                FROM AXReport_EventLogs A
                CROSS JOIN AXTools_Servers B 
                WHERE Guid = '$Guid' AND A.SERVERNAME = B.SERVERNAME --AND (SOURCE LIKE '%Dynamics%' OR SOURCE LIKE '%MSSQLSERVER%' OR SOURCE LIKE 'Application%')"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxEventLogs = New-Object System.Data.DataSet
    $Adapter.Fill($AxEventLogs)
    $Script:ReportDP | Add-Member -Name AxEventLogsChart -Value $($AxEventLogs.Tables[0]) -MemberType NoteProperty -Force

    $Query = "SELECT A.ServerName, B.ServerType, A.LogName, A.EntryType as Type, A.EventID as ID, A.Source, Count(1) as Count
                FROM AXReport_EventLogs A
                CROSS JOIN AXTools_Servers B
                WHERE Guid = '$Guid' AND A.SERVERNAME = B.SERVERNAME --AND (SOURCE LIKE '%Dynamics%' OR SOURCE LIKE '%MSSQLSERVER%' OR SOURCE LIKE 'Application%')
                GROUP BY A.ServerName, B.ServerType, A.LogName, A.EntryType, A.EventID, A.Source"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxEventLogs = New-Object System.Data.DataSet
    $Adapter.Fill($AxEventLogs)
    $Script:ReportDP | Add-Member -Name AxEventLogs -Value $($AxEventLogs.Tables[0] | Select ServerName, ServerType, LogName, Type, ID, Source, Count) -MemberType NoteProperty -Force
    #return $AxEventLogs.Tables[0] | Select ServerName, ServerType, LogName, Type, ID, Source, Count
#}

#function Get-AxMRPLogs
#{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT TOP 7 STARTDATETIME as [Start Time(CST)], ENDDATETIME as [End Time(CST)], ((TIMECOPY+TIMECOVERAGE+TIMEUPDATE)/60) AS [TotalTime], Cancelled, USEDTODAYSDATE as [Todays Date], NUMOFITEMS as Items, NUMOFINVENTONHAND as OnHand, NUMOFSALESLINE as SalesLines, NUMOFPURCHLINE as PurchLines, NUMOFTRANSFERPLANNEDORDER as Transfers, NUMOFITEMPLANNEDORDER as Orders, NUMOFINVENTJOURNAL as InventJournals, LOG as Log
                FROM AXReport_AxMRP 
                WHERE Guid = '$Guid' AND REQPLANID = 'MFIS'
                ORDER BY STARTDATETIME DESC"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxMRPLogs = New-Object System.Data.DataSet
    $Adapter.Fill($AxMRPLogs)  
    $Script:ReportDP | Add-Member -Name AxMRPLogs -Value $($AxMRPLogs.Tables[0] | Select 'Start Time(CST)', 'End Time(CST)', 'TotalTime', Cancelled, Items, OnHand, SalesLines, PurchLines, Transfers, Orders, InventJournals, Log) -MemberType NoteProperty
    #return $AxMRPLogs.Tables[0] | Select 'Start Time(CST)', 'End Time(CST)', 'TotalTime', Cancelled, Items, OnHand, SalesLines, PurchLines, Transfers, Orders, InventJournals, Log
#}

#function Get-SQLErrorLogs
#{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT MAX(LOGDATE) as Date, MAX(PROCESSINFO) as Process, TEXT as Log, MAX(Server) as Server, MAX([Database]) as [Database], COUNT(TEXT) as Count
                FROM AXReport_SqlLogs
                WHERE Guid = '$Guid' AND
		                PROCESSINFO NOT LIKE 'Backup' AND PROCESSINFO NOT LIKE 'Logon' AND
		                TEXT NOT LIKE 'SQL Trace%'
			    GROUP BY TEXT
			    ORDER BY MAX(LOGDATE)"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $SQLErrorLogs = New-Object System.Data.DataSet
    $Adapter.Fill($SQLErrorLogs)
    $Script:ReportDP | Add-Member -Name SQLErrorLogs -Value $($SQLErrorLogs.Tables[0] | Select  Date, Process, Log, Server, Database, Count) -MemberType NoteProperty
    #return $SQLErrorLogs.Tables[0] | Select  Date, Process, Log, Server, Database, Count
#}

#function Get-SSRSErrorLogs
#{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT INSTANCENAME as Instance, STATUS as Message, REPORTPATH as Report, COUNT(REPORTPATH) as Count
                FROM AXReport_SRSLogs
                WHERE Guid = '$Guid'
                GROUP BY INSTANCENAME, STATUS, REPORTPATH
                ORDER BY COUNT DESC, INSTANCENAME"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $SSRSErrorLogs = New-Object System.Data.DataSet
    $Adapter.Fill($SSRSErrorLogs)
    $Script:ReportDP | Add-Member -Name SSRSErrorLogs -Value $($SSRSErrorLogs.Tables[0] | Select Instance, Message, Report, Count) -MemberType NoteProperty
    #return $SSRSErrorLogs.Tables[0] | Select Instance, Message, Report, Count
#}

#function Get-SSRSUsers
#{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT UserName as [User], COUNT(1) as Count
                FROM AXReport_SRSLogs
                WHERE Guid = '$Guid'
                GROUP BY USERNAME
                ORDER BY COUNT DESC"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $SSRSUsers = New-Object System.Data.DataSet
    $Adapter.Fill($SSRSUsers)
    $Script:ReportDP | Add-Member -Name SSRSUsers -Value $($SSRSUsers.Tables[0]) -MemberType NoteProperty
    #return $SSRSUsers.Tables[0]
#}

#function Get-SSRSWeek
#{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT TOP 7 Guid, CONVERT(date, MAX(TIMESTART)) as [Date], COUNT(1) as Count
                FROM AXReport_SRSLogs
                GROUP BY Guid
                ORDER BY 2 DESC"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $SSRSWeek = New-Object System.Data.DataSet
    $Adapter.Fill($SSRSWeek)
    $Script:ReportDP | Add-Member -Name SSRSWeek -Value $($SSRSWeek.Tables[0] | Select @{n='Date';e={($_.Date | Get-Date -Format "MM/dd/yyyy")}}, Count) -MemberType NoteProperty
    #return $SSRSWeek.Tables[0] | Select @{n='Date';e={($_.Date | Get-Date -Format "MM/dd/yyyy")}}, Count

#}

#function Get-PermonDataLogs
#{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT A.ServerName, B.ServerType, A.Counter, A.CounterType, 
		                CASE  WHEN A.COUNTER like '%Private Bytes%' THEN SUM(ROUND((MAXIMUM/1073741824),2))
			                  WHEN A.COUNTER like '%Bytes %' THEN SUM(ROUND((MAXIMUM/1024),2))
			                  WHEN A.COUNTER like '%Virtual Bytes%' THEN SUM(ROUND((MAXIMUM/1073741824),2))
			                  WHEN A.COUNTER like '%Working Set%' THEN SUM(ROUND((MAXIMUM/1073741824),2))
			                  WHEN A.COUNTER like '%Available GBytes%' THEN SUM(ROUND((MAXIMUM/1024),2))
			                  WHEN A.COUNTER like '%Total Server Memory%' THEN SUM(ROUND((MAXIMUM/1048576),2))
			                  ELSE SUM(ROUND(MAXIMUM,2))
		                END AS Max,
		                CASE  WHEN A.COUNTER like '%Private Bytes%' THEN SUM(ROUND((MINIMUM/1073741824),2))
			                  WHEN A.COUNTER like '%Bytes %' THEN SUM(ROUND((MINIMUM/1024),2))
			                  WHEN A.COUNTER like '%Virtual Bytes%' THEN SUM(ROUND((MINIMUM/1073741824),2))
			                  WHEN A.COUNTER like '%Working Set%' THEN SUM(ROUND((MINIMUM/1073741824),2))
			                  WHEN A.COUNTER like '%Available GBytes%' THEN SUM(ROUND((MAXIMUM/1024),2))
			                  WHEN A.COUNTER like '%Total Server Memory%' THEN SUM(ROUND((MAXIMUM/1048576),2))
			                  ELSE SUM(ROUND(MINIMUM,2))
		                END AS Min,
		                CASE  WHEN A.COUNTER like '%Private Bytes%' THEN SUM(ROUND((AVERAGE/1073741824),2))
			                  WHEN A.COUNTER like '%Bytes %' THEN SUM(ROUND((AVERAGE/1024),2))
			                  WHEN A.COUNTER like '%Virtual Bytes%' THEN SUM(ROUND((AVERAGE/1073741824),2))
			                  WHEN A.COUNTER like '%Working Set%' THEN SUM(ROUND((AVERAGE/1073741824),2))
			                  WHEN A.COUNTER like '%Available GBytes%' THEN SUM(ROUND((AVERAGE/1024),2))
			                  WHEN A.COUNTER like '%Total Server Memory%' THEN SUM(ROUND((AVERAGE/1048576),2))
			                  ELSE SUM(ROUND(AVERAGE,2))
		                END AS Avg,
				    COUNT(A.SERVERNAME) AS [Check]
                FROM AXReport_PerfmonData A
                CROSS JOIN AXTools_Servers B
                WHERE A.Guid = '$Guid' AND A.REPORTVIEW = 1 AND A.SERVERNAME = B.SERVERNAME
			    GROUP BY A.SERVERNAME, B.SERVERTYPE, A.COUNTER, A.COUNTERTYPE
                ORDER BY A.SERVERNAME"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $PermonDataLogs = New-Object System.Data.DataSet
    $Adapter.Fill($PermonDataLogs)
    $Script:ReportDP | Add-Member -Name PermonDataLogs -Value $($PermonDataLogs.Tables[0] | Select ServerName, ServerType, Counter, CounterType, Max, Min, Avg) -MemberType NoteProperty
    #return $PermonDataLogs.Tables[0] | Select ServerName, ServerType, Counter, CounterType, Max, Min, Avg
}

Run-Report


function Get-HtmlOpen {
<#
	.SYNOPSIS
		Get's HTML for the header of the HTML report
    .PARAMETER TitleText
		The title of the report
#>
[CmdletBinding()]
param (
	[String] $TitleText
)
	
$CurrentDate = Get-Date -format "MMM d, yyyy hh:mm tt"
$Report = @"
MIME-Version: 1.0
Content-Type: multipart/related; boundary="PART"; type="text/html"

--PART
Content-Type: text/html; charset=us-ascii
Content-Transfer-Encoding: 7bit

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<title>$($TitleText)</title>
<style type="text/css">
* {    
    margin: 0px;
    font-family: sans-serif;
    font-size: 8pt;
}

body {
    margin: 8px 5px 8px 5px; 
}

hr {
    height: 4px; 
    background-color: #337e94; 
    border: 0px;
}

table {
    table-layout: auto; 
    width: 100%;
    border-collapse: collapse;   
}

th {
    vertical-align: top; 
    text-align: left;
    padding: 2px 5px 2px 5px;
}

td {
    vertical-align: top; 
    padding: 2px 5px 2px 5px;
    border-top: 1px solid #bbbbbb;  
}

div.section {
    padding-bottom: 12px;
} 

div.header {
    border: 1px solid #bbbbbb; 
    padding: 4px 5em 0px 5px; 
    margin: 0px 0px -1px 0px;
    height: 2em; 
    width: 95%; 
    font-weight: bold ;
    color: #ffffff;
    background-color: #337e94;
}

div.content { 
    border: 1px solid #bbbbbb; 
    padding: 4px 0px 5px 11px; 
    margin: 0px 0px -1px 0px;
    width: 95%; 
    color: #000000; 
    background-color: #f9f9f9;
}

div.reportname {
    font-size: 16pt; 
    font-weight: bold;
}

div.reportdate {
    font-size: 12pt; 
    font-weight: bold;
}

div.footer {
    padding-right: 5em;
    text-align: right; 
}

table.fixed {
    table-layout: fixed; 
}

th.content { 
    border-top: 1px solid #bbbbbb; 
	width: 25%;
}

td.content { 
	width: 75%;
}

td.groupby {
	border-top: 3px double #bbbbbb;
}

.green {
	background-color: #a1cda4;
}

.yellow {
	background-color: #fffab1;
}

.red {
	background-color: #f5a085;
}

.odd {
	background-color: #D5D8DC;
}

.even {
	background-color: #F7F9F9;
}

.header {
	background-color: #616A6B; color: #F7F9F9;
}

div.column { width: 100%; float: left; overflow-y: auto; }
div.first  { border-right: 1px  grey solid; width: 49% }
div.second { margin-left: 10px;width: 49% }

</style>

<script type="text/javascript"> 
function show(obj) {
  document.getElementById(obj).style.display='block'; 
  document.getElementById("hide_" + obj).style.display=''; 
  document.getElementById("show_" + obj).style.display='none'; 
} 
function hide(obj) { 
  document.getElementById(obj).style.display='none'; 
  document.getElementById("hide_" + obj).style.display='none'; 
  document.getElementById("show_" + obj).style.display=''; 
} 
</script> 

</head>

<body onload="hide();">

<div class="section">
    <div class="ReportName">$($TitleText) - $((Get-Date).AddDays(-1) | Get-Date -Format "D")</div>
    <hr/>
</div>
"@
	Return $Report
}


function Get-HtmlClose
{
$Report = @"
<div class="section">
    <hr />
    <div class="Footer">$Footer</div>
</div>
    
</body>
</html>

--PART-- 
"@
	Write-Output $Report
}


function Get-HtmlContentOpen {
<#
	.SYNOPSIS
		Creates a section in HTML
	    .PARAMETER HeaderText
			The heading for the section
		.PARAMETER IsHidden
		    Switch parameter to define if the section can collapse
		.PARAMETER BackgroundShade
		    An int for 1 to 6 that defines background shading
#>	
Param(
	[string]$HeaderText, 
	[switch]$IsHidden, 
	[validateset(1,2,3,4,5,6)][int]$BackgroundShade
)

switch ($BackgroundShade)
{
    1 { $bgColorCode = "#F8F8F8" }
	2 { $bgColorCode = "#D0D0D0" }
    3 { $bgColorCode = "#A8A8A8" }
    4 { $bgColorCode = "#888888" }
    5 { $bgColorCode = "#585858" }
    6 { $bgColorCode = "#282828" }
    default { $bgColorCode = "#ffffff" }
}



if ($IsHidden) {
	$RandomNumber = Get-Random
	$Report = @"
<div class="section">
    <div class="header">
        <a name="$($HeaderText)">$($HeaderText)</a> (<a id="show_$RandomNumber" href="javascript:void(0);" onclick="show('$RandomNumber');" style="color: #ffffff;">Show</a><a id="hide_$RandomNumber" href="javascript:void(0);" onclick="hide('$RandomNumber');" style="color: #ffffff; display:none;">Hide</a>)
    </div>
    <div class="content" id="$RandomNumber" style="display:none;background-color:$($bgColorCode);"> 
"@	
}
else {
	$Report = @"
<div class="section">
    <div class="header">
        <a name="$($HeaderText)">$($HeaderText)</a>
    </div>
    <div class="content" style="background-color:$($bgColorCode);"> 
"@
}
	Return $Report
}

function Get-HtmlContentClose {
<#
	.SYNOPSIS
		Closes an HTML section
#>	
	$Report = @"
</div>
</div>
"@
	Return $Report
}

function Get-HtmlContentTable {
<#
	.SYNOPSIS
		Creates an HTML table from an array of objects
	    .PARAMETER ArrayOfObjects
			An array of objects
		.PARAMETER Fixed
		    fixes the html column width by the number of columns
		.PARAMETER GroupBy
		    The column to group the data.  make sure this is first in the array
#>	
param(
	[Array]$ArrayOfObjects, 
	[Switch]$Fixed, 
	[String]$GroupBy
)
	if ($GroupBy -eq '') {
		$Report = $ArrayOfObjects | ConvertTo-Html -Fragment
		$Report = $Report -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', ""
		$Report = $Report -replace "<tr>(.*)<td>Green</td></tr>","<tr class=`"green`">`$+</tr>"
		$Report = $Report -replace "<tr>(.*)<td>Yellow</td></tr>","<tr class=`"yellow`">`$+</tr>"
    	$Report = $Report -replace "<tr>(.*)<td>Red</td></tr>","<tr class=`"red`">`$+</tr>"
		$Report = $Report -replace "<tr>(.*)<td>Odd</td></tr>","<tr class=`"odd`">`$+</tr>"
		$Report = $Report -replace "<tr>(.*)<td>Even</td></tr>","<tr class=`"even`">`$+</tr>"
		$Report = $Report -replace "<tr>(.*)<td>None</td></tr>","<tr>`$+</tr>"
		$Report = $Report -replace '<th>RowColor</th>', ''

		if ($Fixed.IsPresent) {	$Report = $Report -replace '<table>', '<table class="fixed">' }
	}
	else {
		$NumberOfColumns = ($ArrayOfObjects | Get-Member -MemberType NoteProperty  | select Name).Count
		$Groupings = @()
		$ArrayOfObjects | select $GroupBy -Unique  | sort $GroupBy | foreach { $Groupings += [String]$_.$GroupBy}
		if ($Fixed.IsPresent) {	$Report = '<table class="fixed">' }
		else { $Report = "<table>" }
		$GroupHeader = $ArrayOfObjects | ConvertTo-Html -Fragment 
		$GroupHeader = $GroupHeader -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', "" -replace '<table>', "" -replace '</table>', "" -replace "<td>.+?</td>" -replace "<tr></tr>", ""
		$GroupHeader = $GroupHeader -replace '<th>RowColor</th>', ''
		$Report += $GroupHeader
		foreach ($Group in $Groupings) {
			$Report += "<tr><td colspan=`"$NumberOfColumns`" class=`"groupby`">$Group</td></tr>"
			$GroupBody = $ArrayOfObjects | where { [String]$($_.$GroupBy) -eq $Group } | select * -ExcludeProperty $GroupBy | ConvertTo-Html -Fragment
			$GroupBody = $GroupBody -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', "" -replace '<table>', "" -replace '</table>', "" -replace "<th>.+?</th>" -replace "<tr></tr>", "" -replace '<tr><td>', "<tr><td></td><td>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Green</td></tr>","<tr class=`"green`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Yellow</td></tr>","<tr class=`"yellow`">`$+</tr>"
    		$GroupBody = $GroupBody -replace "<tr>(.*)<td>Red</td></tr>","<tr class=`"red`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Odd</td></tr>","<tr class=`"odd`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Even</td></tr>","<tr class=`"even`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>None</td></tr>","<tr>`$+</tr>"
			$Report += $GroupBody
		}
		$Report += "</table>" 
	}
	$Report = $Report -replace 'URL01', '<a href="'
	$Report = $Report -replace 'URL02', '">'
	$Report = $Report -replace 'URL03', '</a>'
	
	if ($Report -like "*<tr>*" -and $report -like "*odd*" -and $report -like "*even*") {
			$Report = $Report -replace "<tr>",'<tr class="header">'
	}
	
	return $Report
}

function Get-HtmlContentText 
{
<#
	.SYNOPSIS
		Creates an HTML entry with heading and detail
	    .PARAMETER Heading
			The type of logo
		.PARAMETER Detail
		     Some additional pish
#>	
param(
	$Heading,
	$Detail
)

$Report = @"
<table><tbody>
	<tr>
	<th class="content">$Heading</th>
	<td class="content">$($Detail)</td>
	</tr>
</tbody></table>
"@
$Report = $Report -replace 'URL01', '<a href="'
$Report = $Report -replace 'URL02', '">'
$Report = $Report -replace 'URL03', '</a>'
Return $Report
}

function Set-TableRowColor {
<#
	.SYNOPSIS
		adds a row colour field to the array of object for processing with htmltable
	    .PARAMETER ArrayOfObjects
			The type of logo
		.PARAMETER Green
		     Some additional pish
		.PARAMETER Yellow
		     Some additional pish
		.PARAMETER Red
		    use $this and an expression to measure the value
		.PARAMETER Alertnating
			a switch the will define Odd and Even Rows in the rowcolor column 
#>	
Param (
	$ArrayOfObjects, 
	$Green, 
	$Yellow, 
	$Red,
	[switch]$Alternating 
) 
    if ($Alternating) {
		$ColoredArray = $ArrayOfObjects | Add-Member -MemberType ScriptProperty -Name RowColor -Value {
		if ((([array]::indexOf($ArrayOfObjects,$this)) % 2) -eq 0) {'Odd'}
		if ((([array]::indexOf($ArrayOfObjects,$this)) % 2) -eq 1) {'Even'}
		} -PassThru -Force | Select-Object *
	} else {
		$ColoredArray = $ArrayOfObjects | Add-Member -MemberType ScriptProperty -Name RowColor -Value {
			if (Invoke-Expression $Green) {'Green'} 
			elseif (Invoke-Expression $Red) {'Red'} 
			elseif (Invoke-Expression $Yellow) {'Yellow'} 
			else {'None'}
			} -PassThru -Force | Select-Object *
	}
	return $ColoredArray
}

function New-HTMLBarChartObject
{
<#
	.SYNOPSIS
		create a Bar chart object for use with Create-HTMLPieChart
#>	
	$ChartSize = New-Object PSObject -Property @{`
		Width = 500
		Height = 400
		Left = 40
		Top = 30
	}
	
	$DataDefinition = New-Object PSObject -Property @{`
		AxisXTitle = "AxisXTitle"
		AxisYTitle = "AxisYTitle"
		DrawingStyle = "Cylinder"
		DataNameColumnName = "name"
		DataValueColumnName = "count"
		
	}
	
	$ChartStyle = New-Object PSObject -Property @{`
		BackColor = [System.Drawing.Color]::Transparent
		ExplodeMaxValue = $false
		Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor	[System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	}
	
	$ChartObject = New-Object PSObject -Property @{`
		Type = "Column"
		Title = "Chart Title"
		Size = $ChartSize
		DataDefinition = $DataDefinition
		ChartStyle = $ChartStyle
	}
	return $ChartObject
}

function New-HTMLChart
{
<#
	.SYNOPSIS
		adds a row colour field to the array of object for processing with htmltable
	    .PARAMETER PieChartObject
			This is a custom object with Pie chart properties, Create-HTMLPieChartObject
		.PARAMETER PieChartData
			Required an array with the headings Name and Count.  Using Powershell Group-object on an array
		    
#>
	param (
		$ChartObject,
		$ChartData
	)
	
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
	
	#Create our chart object 
	$Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
	$Chart.Width = $ChartObject.Size.Width
	$Chart.Height = $ChartObject.Size.Height
	$Chart.Left = $ChartObject.Size.Left
	$Chart.Top = $ChartObject.Size.Top
	
	#Create a chartarea to draw on and add this to the chart 
	$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
	$Chart.ChartAreas.Add($ChartArea)
	[void]$Chart.Series.Add("Data")
	
	#Add a datapoint for each value specified in the arguments (args) 
	foreach ($value in $ChartData)
	{
		$datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $value.Count)
		$datapoint.AxisLabel = [string]$value.Name
		$Chart.Series["Data"].Points.Add($datapoint)
	}
	
	switch ($ChartObject.type) {
		"Column"	{
			$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
			$Chart.Series["Data"]["DrawingStyle"] = $ChartObject.ChartStyle.DrawingStyle
			($Chart.Series["Data"].points.FindMaxByValue())["Exploded"] = $ChartObject.ChartStyle.ExplodeMaxValue
		}
		
		"Pie" {
			$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
			$Chart.Series["Data"]["PieLabelStyle"] = $ChartObject.ChartStyle.PieLabelStyle
			$Chart.Series["Data"]["PieLineColor"] = $ChartObject.ChartStyle.PieLineColor
			$Chart.Series["Data"]["PieDrawingStyle"] = $ChartObject.ChartStyle.PieDrawingStyle
			($Chart.Series["Data"].points.FindMaxByValue())["Exploded"] = $ChartObject.ChartStyle.ExplodeMaxValue
			
		}
		default
		{
				
		}
	}
	
    #Set the title of the Chart to the current date and time 
	$Title = new-object System.Windows.Forms.DataVisualization.Charting.Title
	[Void]$Chart.Titles.Add($Title)
	$Chart.Titles[0].Text = $ChartObject.Title
	
	$tempfile = (Join-Path $env:TEMP $ChartObject.Title.replace(' ', '')) + ".png"
	#Save the chart to a file
	if ((test-path $tempfile)) { Remove-Item $tempfile -Force }
	$Chart.SaveImage($tempfile, "png")
	
	$Base64Chart = [Convert]::ToBase64String((Get-Content $tempfile -Encoding Byte))
	$HTMLCode = '<IMG SRC="data:image/gif;base64,' + $Base64Chart + '" ALT="' + $ChartObject.Title + '">'
	return $HTMLCode
	#return $tempfile
}

function New-HTMLPieChartObject {
<#
	.SYNOPSIS
		create a Pie chart object for use with Create-HTMLPieChart
#>	
	$ChartSize = New-Object PSObject -Property @{`
		Width = 350
		Height = 350 
		Left = 1
		Top = 1
	}
	
	$DataDefinition = New-Object PSObject -Property @{`
		DataNameColumnName = "Name"
		DataValueColumnName = "Count"
	}
	
	$ChartStyle = New-Object PSObject -Property @{`
		#PieLabelStyle = "Outside"
        PieLabelStyle = "Disabled"
		PieLineColor = "Black"
		PieDrawingStyle = "Concave"
		ExplodeMaxValue = $false
	}
	
	$PieChartObject = New-Object PSObject -Property @{`
		Type = "Pie"
		Title = "Chart Title"
		Size = $ChartSize
		DataDefinition = $DataDefinition
		ChartStyle = $ChartStyle
	}
	return $PieChartObject
}

function New-HTMLPieChart {
<#
	.SYNOPSIS
		adds a row colour field to the array of object for processing with htmltable
	    .PARAMETER PieChartObject
			This is a custom object with Pie chart properties, Create-HTMLPieChartObject
		.PARAMETER PieChartData
			Required an array with the headings Name and Count.  Using Powershell Group-object on an array
		    
#>
	param(
		$PieChartObject,
		$PieChartData
		)
	      
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

	#Create our chart object 
	$Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
	$Chart.Width = $PieChartObject.Size.Width
	$Chart.Height = $PieChartObject.Size.Height
	$Chart.Left = $PieChartObject.Size.Left
	$Chart.Top = $PieChartObject.Size.Top

	#Create a chartarea to draw on and add this to the chart 
	$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
	$Chart.ChartAreas.Add($ChartArea) 
	[void]$Chart.Series.Add("Data") 

	#Add a datapoint for each value specified in the arguments (args) 
	foreach ($value in $PieChartData) {
		$datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $value.Count)
		$datapoint.AxisLabel = [string]$value.Name
		$Chart.Series["Data"].Points.Add($datapoint)
	}
	
	$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
	$Chart.Series["Data"]["PieLabelStyle"] = $PieChartObject.ChartStyle.PieLabelStyle
	$Chart.Series["Data"]["PieLineColor"] = $PieChartObject.ChartStyle.PieLineColor 
	$Chart.Series["Data"]["PieDrawingStyle"] = $PieChartObject.ChartStyle.PieDrawingStyle
	($Chart.Series["Data"].points.FindMaxByValue())["Exploded"] = $PieChartObject.ChartStyle.ExplodeMaxValue
	

	#Set the title of the Chart to the current date and time 
	$Title = new-object System.Windows.Forms.DataVisualization.Charting.Title 
	[Void]$Chart.Titles.Add($Title) 
	$Chart.Titles[0].Text = $PieChartObject.Title

	$tempfile = (Join-Path $env:TEMP $PieChartObject.Title.replace(' ','') ) + ".png"
	#Save the chart to a file
	if ((test-path $tempfile)) {Remove-Item $tempfile -Force}
	$Chart.SaveImage( $tempfile  ,"png")

	$Base64Chart = [Convert]::ToBase64String((Get-Content $tempfile -Encoding Byte))
	$HTMLCode = '<IMG SRC="data:image/gif;base64,' + $Base64Chart + '" ALT="' + $PieChartObject.Title + '">'
	return $HTMLCode 
	#return $tempfile
	
}

function Get-HTMLColumn1of2
{
	$report = '<div class="first column">'
	return $report
}

function Get-HTMLColumn2of2
{
	$report = '<div class="second column">'
	return $report
}


function Get-HTMLColumnClose
{
	$report = '</div>'
	return $report
}

function Get-SummaryOpen {
[CmdletBinding()]
param (
	[String] $TitleText
)

$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<title>$($TitleText)</title>
<style type="text/css">
* {    
    margin: 0px;
    font-family: sans-serif;
    font-size: 8pt;
}

body {
    margin: 8px 5px 8px 5px; 
}

hr {
    height: 4px; 
    background-color: #337e94; 
    border: 0px;
}

table {
    table-layout: auto; 
    width: 100%;
    border-collapse: collapse;   
}

th {
    vertical-align: top; 
    text-align: left;
    padding: 2px 5px 2px 5px;
}

td {
    vertical-align: top; 
    padding: 2px 5px 2px 5px;
    border-top: 1px solid #bbbbbb;  
}

div.section {
    padding-bottom: 12px;
} 

div.header {
    border: 1px solid #bbbbbb; 
    margin: 0px 0px -1px 0px;
    height: 2em;
    width: 95%; 
    font-weight: bold ;
    color: #ffffff;
    background-color: #337e94;
}

div.content { 
    border: 1px solid #bbbbbb; 
    margin: 0px 0px -1px 0px;
    width: 95%; 
    color: #000000; 
    background-color: #f9f9f9;
}

div.reportname {
    font-size: 16pt; 
    font-weight: bold;
}

div.footer {
    padding-right: 5em;
    text-align: right; 
}

table.fixed {
    table-layout: fixed; 
}

th.content { 
    border-top: 1px solid #bbbbbb; 
	width: 25%;
}

td.content { 
	width: 75%;
}

td.groupby {
	border-top: 3px double #bbbbbb;
}

.green {
	background-color: #a1cda4;
}

.yellow {
	background-color: #fffab1;
}

.red {
	background-color: #f5a085;
}

.odd {
	background-color: #D5D8DC;
}

.even {
	background-color: #F7F9F9;
}

.header {
	background-color: #616A6B; color: #F7F9F9;
}

</style>

</head>

<div class="section">
    <div class="reportname">$($TitleText)</div>
    <hr/>
    <br>  </br>
</div>
"@
	Return $Report
}

function Get-SummaryClose
{
$Report = @"
<div class="section">
    <hr />
    <div class="footer">$Footer</div>
</div>
    
</body>
</html>
"@
	Write-Output $Report
}

##
#$Conn.Close()
#>