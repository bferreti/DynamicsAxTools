Param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]$Guid,
    [String]$Environment,
    [String]$ReportDate
)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"

Import-Module $ModuleFolder\AX-Tools.psm1 -DisableNameChecking

$ConfigFile = Load-ConfigFile
$ReportFolder = if(!$ConfigFile.Settings.General.ReportPath) { $Dir + "\Reports\AX-Report\$Environment" } else { "$($ConfigFile.Settings.General.ReportPath)\$Environment" }
$LogFolder = if(!$ConfigFile.Settings.General.LogPath) { $Dir + "\Logs\AX-Report\$Environment" } else { "$($ConfigFile.Settings.General.LogPath)\$Environment" }
$ReportDate = $(Get-Date (Get-Date).AddDays(-1) -format MMddyyyy) #Get-Date -f MMddyyHHmm
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
    if($Script:ReportDP.AxMRPLogs.Count -gt 0) {
        $MRPTotalTime = $Script:ReportDP.AxMRPLogs.TotalTime | Measure-Object  -Maximum -Average
        switch -wildcard ($MRPTotalTime) {
            {$($MRPTotalTime.Maximum) -eq 0} {$Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "$($MRPTotalTime.Count) MRP run(s) with no execution time."; RowColor = 'Red' }}
            {($($MRPTotalTime.Maximum) -gt 0) -and ($($MRPTotalTime.Maximum) -le 45)} {$Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "$($MRPTotalTime.Count) MRP run(s) - $($MRPTotalTime.Maximum) minutes."; RowColor = 'Green' }}
            #{($($Script:ReportDP.AxMRPLogs.TotalTime) -gt 45) -and ($($Script:ReportDP.AxMRPLogs.TotalTime) -le 60)} {$Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "$($Script:ReportDP.AxMRPLogs.TotalTime) minutes."; RowColor = 'Yellow' }}
            {($($MRPTotalTime.Maximum) -gt 45)} {$Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "$($MRPTotalTime.Count) MRP run(s) - $($MRPTotalTime.Maximum) minutes."; RowColor = 'Yellow' }}
            Default {$Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "MRP Failed $(if($MRPTotalTime.Maximum -gt 0){ "- $($MRPTotalTime.Maximum) minutes."} else {"."})"; RowColor = 'Red' }}
        }
    }
    else {
         $Script:AxSummary += New-Object PSObject -Property @{ Name = "MRP Status"; Status = "MRP data not found."; RowColor = 'Green'; }
    }

    #AxBatchJobs
    if($Script:ReportDP.AxBatchJobs.Count -eq 0) { 
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Batch Jobs"; Status = "Ok."; RowColor = 'Green' }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Batch Jobs"; Status = "$($Script:ReportDP.AxBatchJobs.Count) Jobs Found."; RowColor = 'Red' }
    }
    if($Script:ReportDP.AxLongBatchJobs.Count -eq 0) {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Long Batch Jobs (>15min)"; Status = "Ok."; RowColor = 'Green' }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Long Batch Jobs (>15min)"; Status = "$($Script:ReportDP.AxLongBatchJobs.Count) Jobs Found."; RowColor = 'Red' }
    }

    #AxRetailJobs
    if($Script:ReportDP.AxCDXJobs.Rows.Count -eq 0) { 
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Retail Jobs"; Status = "Ok."; RowColor = 'Green' }
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Retail Jobs"; Status = "$($Script:ReportDP.AxCDXJobs.Rows.Count) Jobs Found."; RowColor = 'Red' }
    }

    #PerfmonData Color-Set
    $Green = '(($this.Counter -like "CPU Time %" -and $this.Max -le 60) -or ($this.Counter -like "Available GBytes" -and $this.Min -ge 8) -or ($this.Counter -like "Paging File %" -and $this.Max -le 35) -or ($this.Counter -like "*Buffer cache hit ratio" -and $this.Min -ge 95) -or ($this.Counter -like "*Page life expectancy" -and $this.Min -ge 6000))'
    $Yellow = '(($this.Counter -like "CPU Time %" -and $this.Max -gt 60 -and $this.Max -lt 80) -or ($this.Counter -like "Available GBytes" -and $this.Max -gt 4 -and $this.Max -lt 8) -or ($this.Counter -like "Paging File %" -and $this.Max -gt 35 -and $this.Max -lt 50) -or ($this.Counter -like "*Buffer cache hit ratio" -and $this.Min -gt 90 -and $this.Min -lt 95) -or ($this.Counter -like "*Page life expectancy" -and $this.Min -gt 1200 -and $this.Min -lt 6000))'    
    $Red = '(($this.Counter -like "CPU Time %" -and $this.Max -ge 80) -or ($this.Counter -like "Available GBytes" -and $this.Max -le 4) -or ($this.Counter -like "Paging File %" -and $this.Max -ge 50) -or ($this.Counter -like "*Buffer cache hit ratio" -and $this.Min -le 90) -or ($this.Counter -like "*Page life expectancy" -and $this.Min -le 1200))'
    
    #REMOVING INSTANCES NOT RUNNING
    $PermonDataLogsTmp = $Script:ReportDP.PermonDataLogs | Where {$_.ServerType -notmatch 'SQL' -or $_.CounterType -like 'SRV' }
    $PermonDataLogsTmp += $Script:ReportDP.PermonDataLogs | Where {(($_.Max -ne 0) -or ($_.Min -ne 0)) -and ($_.CounterType -notmatch 'SRV') -and ($_.ServerType -match 'SQL')}
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
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $Cmd
        $AxCrash = New-Object System.Data.DataSet
        $Adapter.Fill($AxCrash)
        if($AxCrash.Tables[0].Rows.Count -gt 0) {
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
    }
    else {
        $Script:AxSummary += New-Object PSObject -Property @{ Name = "Event Logs"; Status = "No data found."; RowColor = 'Green' }
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
    $Script:AXReport += Get-HtmlOpen -TitleText ($ReportName) -AxReport
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
    $Yellow = '(New-TimeSpan ($this.StartTime) $(Get-Date)).TotalDays -lt 25'
    $Red = '$this.Status -match "Stopped"'
    $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "AX Services Status"
    $Script:AXReport += Get-HtmlContentTable(Set-TableRowColor $Script:ReportDP.AxServices -Red $Red -Green $Green -Yellow $Yellow)
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
    if($Script:ReportDP.AxBatchJobs.Count -gt 0) {
        $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "AX Batch Jobs Errors [Total - $($Script:ReportDP.AxBatchJobs.Count)]"
        $Script:AXReport += Get-HtmlContentTable ($Script:ReportDP.AxBatchJobs)
        $Script:AXReport += Get-HtmlContentClose
    }
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
        $Script:AXReport += New-HTMLPieChart -PieChartObject $PieChartObject3 -PieChartData ($Script:ReportDP.SSRSUsers | Sort Count -Descending | Select -First 5)
        $Script:AXReport += Get-HtmlContentTable($Script:ReportDP.SSRSUsers | Select User, Count | Sort Count -Descending | Select -First 5)
        $Script:AXReport += Get-HtmlContentClose
        $Script:AXReport += Get-HtmlColumnClose
        $Script:AXReport += Get-HtmlContentClose
    }

    #CDX Jobs Errors
    if($Script:ReportDP.AxCDXJobs.Rows.Count -gt 0) {
        $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "CDX Jobs Errors [Total - $($Script:ReportDP.AxCDXJobs.Rows.Count)]" 
        $Script:AXReport += Get-HtmlContentTable (Set-TableRowColor $Script:ReportDP.AxCDXJobs -Alternating)
        $Script:AXReport += Get-HtmlContentClose
    }

    ##SQL Error Logs
    $Script:AXReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "SQL Server Error Logs [Total - $($Script:ReportDP.SQLErrorLogs.Count)]" 
    $Script:AXReport += Get-HtmlContentTable ($Script:ReportDP.SQLErrorLogs) 
    $Script:AXReport += Get-HtmlContentClose

    #if($Script:ReportDP.SSRSWeek.Count -gt 0) {
    #    $PieChartObject4 = New-HTMLPieChartObject
    #    $PieChartObject4.Title = " "
    #    $PieChartObject4.Size.Height = 400
    #    $PieChartObject4.Size.Width = 400
    #    $PieChartObject4.ChartStyle.ExplodeMaxValue = $true    			
    #
    #    $Script:AXReport += Get-HtmlContentOpen
    #    $Script:AXReport += Get-HtmlColumn1of2
    #    $Script:AXReport += Get-HtmlContentOpen -HeaderText "SSRS Errors 7 Days"
    #    $Script:AXReport += New-HTMLPieChart -PieChartObject $PieChartObject4 -PieChartData ($Script:ReportDP.SSRSWeek | Sort Date)
    #    $Script:AXReport += Get-HtmlContentTable ($Script:ReportDP.SSRSWeek | Select Date, Count | Sort Date -Descending)
    #    $Script:AXReport += Get-HtmlContentClose
    #    $Script:AXReport += Get-HtmlColumnClose
    #
    #    $Script:AXReport += Get-HtmlColumn2of2
    #
    #    $Script:AXReport += Get-HtmlColumnClose
    #    $Script:AXReport += Get-HtmlContentClose
    #}

    #Close Report
    $Script:AXReport += Get-HtmlContentClose
    $Script:AXReport += Get-HtmlClose -FooterText "Guid: $($Guid)" -AxReport
}

function Save-ReportFile
{
    ##Add Summary Email Info
    $Script:AxSummary += New-Object PSObject -Property @{ Name = '**Please see the attached report for details.'; Status = ''; RowColor = 'None' }
    $AXREmail = @()
    $AXREmail += Get-HtmlOpen -TitleText ($ReportName) -AxSummary
    $AXREmail += Get-HtmlContentOpen -HeaderText "Summary Information"
    $AXREmail += Get-HtmlContentTable($Script:AxSummary | Select Name, Status, RowColor)
    $AXREmail += Get-HtmlContentClose
    $AXREmail += Get-HtmlClose -FooterText "Guid: $($Guid)" -AxSummary
    #Save Summary
    $AXReportPath = Join-Path $ReportFolder ("AXReport-$ReportDate-Summary" + ".html")
    $AXREmail | Set-Content -Path $AXReportPath -Force
    #Save Report
    $AXReportPath = Join-Path $ReportFolder ("AXReport-$ReportDate" + ".mht")
    $Script:AXReport | Set-Content -Path $AXReportPath -Force
}

function Run-ReportDP
{
    $Script:ReportDP = New-Object -TypeName System.Object
    $Script:ReportDP | Add-Member -Name ApplicationName -Value 'AX Report Script' -MemberType NoteProperty
    $Script:ReportDP | Add-Member -Name ToolsConnectionObject -Value $(Get-ConnectionString $Script:ReportDP.ApplicationName) -MemberType NoteProperty

    $Query = "SELECT ServerName, Service AS ServiceName, Name, Status, StartTime FROM AXReport_AxServices WHERE Guid = '$Guid'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxServices = New-Object System.Data.DataSet
    $Adapter.Fill($AxServices) | Out-Null
    $Script:ReportDP | Add-Member -Name AxServices -Value $($AxServices.Tables[0] | Select ServerName, ServiceName, Name, Status, @{n='UpTime';e={"$([Math]::Truncate((New-TimeSpan ($_.StartTime) $(Get-Date)).TotalDays)) day(s)"}} ) -MemberType NoteProperty

    $Query = "SELECT HISTORYCAPTION AS [History Caption],JOBCAPTION AS [Job Caption],Status,ServerID AS Server,STARTDATETIMECST AS [Start Time(CST)],ENDDATETIMECST AS [End Time(CST)],EXECUTEDBY AS [User], LOG AS Log
                FROM AXReport_AxBatchJobs 
                WHERE Guid = '$Guid'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxBatchJobs = New-Object System.Data.DataSet
    $Adapter.Fill($AxBatchJobs)
    $Script:ReportDP | Add-Member -Name AxBatchJobs -Value $($AxBatchJobs.Tables[0] | Select 'History Caption', 'Job Caption', 'Status', @{n='Server';e={($_.SERVER -replace '01@','').Trim()}}, 'Start Time(CST)', 'End Time(CST)', 'User', 'Log') -MemberType NoteProperty

    $Query = "SELECT Job, Count, Status, Duration, EXECUTEDBY AS [User], ServerID AS [Server]
                FROM AXReport_AxLongBatchJobs 
                WHERE Guid = '$Guid'
                ORDER BY Duration DESC"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxLongBatchJobs = New-Object System.Data.DataSet
    $Adapter.Fill($AxLongBatchJobs)
    $Script:ReportDP | Add-Member -Name AxLongBatchJobs -Value $($AxLongBatchJobs.Tables[0] | Select 'Job', 'Count', 'Status', 'Duration', 'User', 'Server') -MemberType NoteProperty

    $Query = "SELECT JobID, STATUSDOWNLOADSESSIONDATASTORE AS [Download Status], Message, DateRequested, DateDownloaded, DateApplied, ROWSAFFECTED as [Rows], DATAFILEOUTPUTPATH as [Path], STATUSDOWNLOADSESSION as [Session Status], DATABASE_ as [Database], Name
                FROM AXReport_AxRetailJobs 
                WHERE Guid = '$Guid'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxCDXJobs = New-Object System.Data.DataSet
    $Adapter.Fill($AxCDXJobs)
    $Script:ReportDP | Add-Member -Name AxCDXJobs -Value $($AxCDXJobs.Tables[0] | Select JobID, 'Download Status', Message, DateRequested, DateDownloaded, DateApplied, Rows, Path, 'Session Status', Database, Name) -MemberType NoteProperty

    $Query = "SELECT A.ServerName, B.ServerType, A.EntryType as Type, A.EventID as ID, A.Source
                FROM AXReport_EventLogs A
                CROSS JOIN AXTools_Servers B 
                WHERE Guid = '$Guid' AND A.SERVERNAME = B.SERVERNAME --AND (SOURCE LIKE '%Dynamics%' OR SOURCE LIKE '%MSSQLSERVER%' OR SOURCE LIKE 'Application%')"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
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
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxEventLogs = New-Object System.Data.DataSet
    $Adapter.Fill($AxEventLogs)
    $Script:ReportDP | Add-Member -Name AxEventLogs -Value $($AxEventLogs.Tables[0] | Select ServerName, ServerType, LogName, Type, ID, Source, Count) -MemberType NoteProperty -Force

    $Query = "SELECT TOP 7 ReqPlanId as [Plan], STARTDATETIME as [Start Time(CST)], ENDDATETIME as [End Time(CST)], ((TIMECOPY+TIMECOVERAGE+TIMEUPDATE)/60) AS [TotalTime], Cancelled, USEDTODAYSDATE as [Todays Date], NUMOFITEMS as Items, NUMOFINVENTONHAND as OnHand, NUMOFSALESLINE as SalesLines, NUMOFPURCHLINE as PurchLines, NUMOFTRANSFERPLANNEDORDER as Transfers, NUMOFITEMPLANNEDORDER as Orders, NUMOFINVENTJOURNAL as InventJournals, LOG as Log
                FROM AXReport_AxMRP 
                WHERE Guid = '$Guid' AND COMPLETEUPDATE = '1'
                ORDER BY STARTDATETIME DESC"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $AxMRPLogs = New-Object System.Data.DataSet
    $Adapter.Fill($AxMRPLogs)  
    $Script:ReportDP | Add-Member -Name AxMRPLogs -Value $($AxMRPLogs.Tables[0] | Select 'Plan', 'Start Time(CST)', 'End Time(CST)', 'TotalTime', Cancelled, Items, OnHand, SalesLines, PurchLines, Transfers, Orders, InventJournals, Log) -MemberType NoteProperty

    $Query = "SELECT MAX(LOGDATE) as Date, MAX(PROCESSINFO) as Process, TEXT as Log, MAX(Server) as Server, MAX([Database]) as [Database], COUNT(TEXT) as Count
                FROM AXReport_SQLLog
                WHERE Guid = '$Guid' AND
		                PROCESSINFO NOT LIKE 'Backup' AND PROCESSINFO NOT LIKE 'Logon' AND
		                TEXT NOT LIKE 'SQL Trace%'
			    GROUP BY TEXT
			    ORDER BY MAX(LOGDATE)"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $SQLErrorLogs = New-Object System.Data.DataSet
    $Adapter.Fill($SQLErrorLogs)
    $Script:ReportDP | Add-Member -Name SQLErrorLogs -Value $($SQLErrorLogs.Tables[0] | Select  Date, Process, Log, Server, Database, Count) -MemberType NoteProperty

    $Query = "SELECT INSTANCENAME as Instance, STATUS as Message, REPORTPATH as Report, COUNT(REPORTPATH) as Count
                FROM AXReport_SRSLog
                WHERE Guid = '$Guid'
                GROUP BY INSTANCENAME, STATUS, REPORTPATH
                ORDER BY COUNT DESC, INSTANCENAME"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $SSRSErrorLogs = New-Object System.Data.DataSet
    $Adapter.Fill($SSRSErrorLogs)
    $Script:ReportDP | Add-Member -Name SSRSErrorLogs -Value $($SSRSErrorLogs.Tables[0] | Select Instance, Message, Report, Count) -MemberType NoteProperty

    $Query = "SELECT UserName as [User], COUNT(1) as Count
                FROM AXReport_SRSLog
                WHERE Guid = '$Guid'
                GROUP BY USERNAME
                ORDER BY COUNT DESC"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $SSRSUsers = New-Object System.Data.DataSet
    $Adapter.Fill($SSRSUsers)
    $Script:ReportDP | Add-Member -Name SSRSUsers -Value $($SSRSUsers.Tables[0]) -MemberType NoteProperty

    $Query = "SELECT TOP 7 Guid, CONVERT(date, MAX(TIMESTART)) as [Date], COUNT(1) as Count
                FROM AXReport_SRSLog
                GROUP BY Guid
                ORDER BY 2 DESC"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $SSRSWeek = New-Object System.Data.DataSet
    $Adapter.Fill($SSRSWeek)
    $Script:ReportDP | Add-Member -Name SSRSWeek -Value $($SSRSWeek.Tables[0] | Select @{n='Date';e={($_.Date | Get-Date -Format "MM/dd/yyyy")}}, Count) -MemberType NoteProperty

    $Query = "SELECT A.ServerName, A.ServerType, A.Counter, A.CounterType, 
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
                WHERE A.Guid = '$Guid' AND A.REPORTVIEW = 1
			    GROUP BY A.SERVERNAME, A.SERVERTYPE, A.COUNTER, A.COUNTERTYPE
                ORDER BY A.SERVERNAME"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:ReportDP.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $PermonDataLogs = New-Object System.Data.DataSet
    $Adapter.Fill($PermonDataLogs)
    $Script:ReportDP | Add-Member -Name PermonDataLogs -Value $($PermonDataLogs.Tables[0] | Select ServerName, ServerType, Counter, CounterType, Max, Min, Avg) -MemberType NoteProperty
}

Run-Report