# .DISCLAIMER
#    Microsoft Corporation. All rights reserved.
#    MIT License
#    
#    Copyright (c) 2017 bferreti
#    
#    Permission is hereby granted, free of charge, to any person obtaining a copy
#    of this software and associated documentation files (the "Software"), to deal
#    in the Software without restriction, including without limitation the rights
#    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#    copies of the Software, and to permit persons to whom the Software is
#    furnished to do so, subject to the following conditions:
#    
#    The above copyright notice and this permission notice shall be included in all
#    copies or substantial portions of the Software.
#    
#    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#    LIABILITY (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS 
#    PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER 
#    PECUNIARY LOSS), WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, 
#    ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR 
#    OTHER DEALINGS IN THE SOFTWARE.
#    
#    THE SOFTWARE IS NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT PROGRAMS 
#    OR SERVICES. THE OPINIONS AND VIEWS EXPRESSED ARE THOSE OF THE AUTHOR AND DO
#    NOT NECESSARILY STATE OR REFLECT THOSE OF MICROSOFT.
#    
#    Microsoft is a registered trademark or trademarks of Microsoft Corporation in 
#    the United States and/or other countries.
#
# .NOTES
#    Author         : Bruno Ferreti
#    Prerequisite   : PowerShell for SQL Server Modules (SQLPS)
#    Copyright 2017
#

Param (
    [Parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true)]
    [String]$Environment,
    [Parameter(Position=1,Mandatory=$false,ValueFromPipeline=$true)]
    [Switch]$Rerun
)
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null

## Get PS script directory and assign to DIR variable.
$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"

Import-Module $ModuleFolder\AX-Tools.psm1 -DisableNameChecking

$Script:Configuration = Load-ConfigFile
$ReportFolder = if(!$Script:Configuration.Settings.General.ReportPath) { $Dir + "\Reports\AX-Monitor\$Environment" } else { "$($Script:Configuration.Settings.General.ReportPath)\$Environment" }
$LogFolder = if(!$Script:Configuration.Settings.General.LogPath) { $Dir + "\Logs\AX-Monitor\$Environment" } else { "$($Script:Configuration.Settings.General.LogPath)\$Environment" }
$FileDateTime = Get-Date -f yyyyMMdd-HHmm
$AutoCleanUp = [boolean]::Parse($Script:Configuration.Settings.General.AutoCleanUp)
$Debug = [boolean]::Parse($Script:Configuration.Settings.AXMonitor.Debug)

function Get-SQLMonitoring
{
    Validate-Settings
    if(Get-GRDStatus) {
         Write-ExecLog "GRD Threshold True"
        if(($Script:Settings.CPUTotal -ge $Script:Settings.CPUThold) -and ($Script:Settings.EnableGRD -eq $true)) {
            Write-ExecLog "CPU True ($($Script:Settings.CPUTotal)/GRD-$($Script:Settings.EnableGRD))"
            Get-GRDTables
            Get-AXJobs
            Get-PerfData
            Get-SQLConfig
            Get-CreateReport
            if(![String]::IsNullOrEmpty($Script:Settings.EmailProfile)) {
                Get-SendEmail
            }
        }
        else {
            Write-ExecLog "CPU False ($($Script:Settings.CPUTotal)/GRD-$($Script:Settings.EnableGRD))"
            Get-AXJobs
            Get-PerfData
            Get-SQLConfig
            Get-CreateReport
            if(![String]::IsNullOrEmpty($Script:Settings.EmailProfile)) {
                Get-SendEmail
            }
        }
    }
    else {
        if(($Script:Settings.EnableStats -gt 0) -and (Get-SQLStatisticsInterval)) {
            Write-ExecLog "Statistics True"
            Get-SQLStatistics
        }
        else {
            Write-ExecLog "Statistics False"
        }
    }
    if($Script:Settings.GRDJobs) { 
        Write-ExecLog "Check GRD Jobs $($Script:Settings.GRDJobs.Count)"
        Get-JobStatus
    }
    ## Clean log files from folder
    if($AutoCleanUp) { Do-CleanUp }
}

function Validate-Settings
{
    $Conn = Get-ConnectionString
    $Query = "SELECT * FROM AXTools_Environments                
                WHERE ENVIRONMENT = '$Environment'"
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter($Query, $Conn)
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null

    if (![String]::IsNullOrEmpty($Table.Tables))
    {
        $Script:Settings = New-Object -TypeName System.Object
        $Script:Settings | Add-Member -Name GUID -Value (([Guid]::NewGuid()).Guid) -MemberType NoteProperty
        $Script:Settings | Add-Member -Name ToolsConnection -Value $($Conn) -MemberType NoteProperty
       
        try {
            if($Table.Tables.DBUser) {
                $Query = "SELECT UserName, Password FROM [dbo].[AXTools_UserAccount] WHERE [ID] = '$($Table.Tables.DBUser)'"
                $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter($Query, $Script:Settings.ToolsConnection)
                $UserAccount = New-Object System.Data.DataSet
                $Adapter.Fill($UserAccount) | Out-Null
                $UserPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($UserAccount.Tables[0].Password | ConvertTo-SecureString)))
                $secureUserPassword = $UserPassword | ConvertTo-SecureString -AsPlainText -Force 
                $SqlCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserAccount.Tables[0].UserName, $secureUserPassword
                $Script:Settings | Add-Member -Name SqlCredential -Value $($SqlCredential) -MemberType NoteProperty
                $SqlConn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection
                $SqlConn.ServerInstance = $Table.Tables.DBServer
                $SqlConn.DatabaseName = $Table.Tables.DBName
                $SqlConn.ApplicationName = 'SQL Monitoring Script'
                $SqlServer = New-Object Microsoft.SqlServer.Management.SMO.Server($SqlConn)
                $SqlServer.ConnectionContext.ConnectAsUser = $true
                $SqlServer.ConnectionContext.ConnectAsUserPassword = $SqlCredential.GetNetworkCredential().Password
                $SqlServer.ConnectionContext.ConnectAsUserName = $SqlCredential.GetNetworkCredential().UserName
                $SqlServer.ConnectionContext.Connect()
            }
            else {
                $SqlConn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection
                $SqlConn.ServerInstance = $Table.Tables.DBServer
                $SqlConn.DatabaseName = $Table.Tables.DBName
                $SqlConn.ApplicationName = 'SQL Monitoring Script'
                $SqlServer = New-Object Microsoft.SqlServer.Management.SMO.Server($SqlConn)
                $SqlServer.ConnectionContext.Connect()
            }
            $Script:Settings | Add-Member -Name DBServer -Value $Table.Tables.DBServer -MemberType NoteProperty
            $Script:Settings | Add-Member -Name DBName -Value $Table.Tables.DBName -MemberType NoteProperty
            $Script:Settings | Add-Member -Name Description -Value $Table.Tables.Description -MemberType NoteProperty
            $Script:Settings | Add-Member -Name SQLServer -Value $($SqlServer) -MemberType NoteProperty
            $Script:Settings | Add-Member -Name NetBios -Value $(($Script:Settings.SQLServer.Information.Properties | Where-Object { $_.Name -eq 'ComputerNamePhysicalNetBIOS' }).Value) -MemberType NoteProperty
        }
        catch {
            Write-Host "Failed to connect to AX Database. $($_.Exception.Message)"
            break
        }
    
        if(![String]::IsNullOrEmpty($Table.Tables.Environment)) {
            $Script:Settings | Add-Member -Name Environment -Value $Table.Tables.Environment -MemberType NoteProperty
        }
        else {
            Write-Host 'Environment not found.'
            break
        }
    
        if(([String]::IsNullOrEmpty($Table.Tables.CPUThold)) -or ($Table.Tables.CPUThold -le 0)) {
            $Script:Settings | Add-Member -Name CPUThold -Value 65 -MemberType NoteProperty
        }
        else {
            $Script:Settings | Add-Member -Name CPUThold -Value $Table.Tables.CPUThold -MemberType NoteProperty
        }

        if(([String]::IsNullOrEmpty($Table.Tables.BlockThold)) -or ($Table.Tables.BlockThold -le 0)) {
            $Script:Settings | Add-Member -Name BlockThold -Value 15 -MemberType NoteProperty
        }
        else {
            $Script:Settings | Add-Member -Name BlockThold -Value $Table.Tables.BlockThold -MemberType NoteProperty
        }

        if(([String]::IsNullOrEmpty($Table.Tables.WaitingThold)) -or ($Table.Tables.WaitingThold -le 0)) {
            $Script:Settings | Add-Member -Name WaitingThold -Value 1800000 -MemberType NoteProperty
        }
        else {
            $Script:Settings | Add-Member -Name WaitingThold -Value $Table.Tables.WaitingThold -MemberType NoteProperty
        }


        if($Table.Tables.RunGRD -match '1') {
            $Script:Settings | Add-Member -Name EnableGRD -Value $true -MemberType NoteProperty
        }
        else {
            $Script:Settings | Add-Member -Name EnableGRD -Value $false -MemberType NoteProperty
        }

        if($Table.Tables.RunStats -match '0|1|2') {
            $Script:Settings | Add-Member -Name EnableStats -Value $Table.Tables.RunStats -MemberType NoteProperty
        }
        else {
            $Script:Settings | Add-Member -Name EnableStats -Value 0 -MemberType NoteProperty
        }

        $Script:Settings | Add-Member -Name EmailProfile -Value $Table.Tables.EmailProfile -MemberType NoteProperty
        $Table.Dispose()
    }
}

function Get-SQLStatus
{
    $SQLProcesses = $Script:Settings.SQLServer.EnumProcesses()
    if($SQLProcesses) {
        $Script:Settings | Add-Member -Name Processes -Value $SQLProcesses -MemberType NoteProperty
    }
    else {
        $Script:Settings | Add-Member -Name Processes -Value $($Script:Settings.SQLServer.EnumProcesses()) -MemberType NoteProperty
    }
    $Script:Settings | Add-Member -Name Blocking -Value $($Script:Settings.Processes | Select Spid, BlockingSpid | Where { $_.BlockingSpid -ne 0 }) -MemberType NoteProperty
    
    $HBlockers = @()
    foreach($Block in $Script:Settings.Blocking) {
        ## Moving spids to find the headblocker
        $NextSpid = ($Script:Settings.Processes | Where {$_.Spid -eq $($Block.BlockingSpid)}).Spid
        $NextBlocker = ($Script:Settings.Processes | Where {$_.Spid -eq $($Block.BlockingSpid)}).BlockingSpid
        if(($NextBlocker -eq 0) -and ($NextSpid -ge 50)) {
            #Write-Host "HeadBlocker $NextSpid"
            $HBlockers += $NextSpid
        }
    }

    if($Script:Settings.Blocking.Spid.Count -ne 0) {
        $SelectSpid = @($Script:Settings.Blocking.Spid)
        $SelectSpid += @($HBlockers | Select -Unique)
        $Script:Settings | Add-Member -Name HeadBlockers -Value $($HBlockers | Select -Unique) -MemberType NoteProperty
    }
    else{
        $SelectSpid = $Script:Settings.Processes.Spid | Where { $_ -gt 50 }
    }

    $Query = "SELECT r.start_time as start_date_time
                    , r.session_id as [spid]
                    , r.blocking_session_id as [blocker]
	                , r.status
                    , s.[host_name]
                    , CAST(s.[context_info] as varchar(128)) as [context_info]
                    , r.wait_time as [wait_time_ms]
                    , CONVERT(VARCHAR,DATEADD(ms,r.wait_time,0),114) as [wait_time]
                    , r.total_elapsed_time as [total_time_ms]
                    , CONVERT(VARCHAR,DATEADD(ms,r.total_elapsed_time,0),114) as [total_time]
                    , r.cpu_time as [cpu_time_ms]
                    , CONVERT(VARCHAR,DATEADD(ms,r.cpu_time,0),114) as [cpu_time]
                    , CAST(CAST(r.cpu_time as DEC(20,8)) / NULLIF(CAST(r.total_elapsed_time as DEC(20,8)),0) as DEC(20,8)) * 100 AS [cpu_time_perc]
                    , r.reads
                    , r.writes
                    , r.logical_reads
                    , r.wait_type
                    , DB_NAME(r.database_id) as [database]
                    , (SELECT REPLACE(REPLACE([text],char(10),''),char(13),' ') FROM sys.dm_exec_sql_text(r.[sql_handle])) as [sql_text]
	                , '0x'+CONVERT(varchar(max),r.plan_handle,2) as plan_handle
                FROM sys.dm_exec_requests as r with(nolock)
                JOIN sys.dm_exec_sessions as s with(nolock) ON r.session_id = s.session_id
                WHERE r.session_id <> @@SPID AND r.session_id > 50 AND host_name is not null
                UNION 
                SELECT s.last_request_start_time as start_date_time
                    , c.session_id as [spid]
                    , 0 as [blocker]
	                , r.status
                    , s.[host_name]
                    , CAST(s.[context_info] as varchar(128)) as [context_info]
                    , 0 as [wait_time_ms]
                    , '00:00:00:000' as [wait_time]
                    , s.total_elapsed_time as [total_time_ms]
                    , CONVERT(VARCHAR,DATEADD(ms,s.total_elapsed_time,0),114) as [total_time]
                    , s.cpu_time as [cpu_time_ms]
                    , CONVERT(VARCHAR,DATEADD(ms,s.cpu_time,0),114) as [cpu_time]
                    , CAST(CAST(s.cpu_time as DEC(20,8)) / NULLIF(CAST(s.total_elapsed_time as DEC(20,8)),0) as DEC(20,8)) * 100 AS [cpu_time_perc]
                    , s.reads
                    , s.writes
                    , s.logical_reads
                    , '' as wait_type
                    , DB_NAME(s.database_id) as [database]
                    , (SELECT REPLACE(REPLACE([text],char(10),''),char(13),' ') FROM sys.dm_exec_sql_text(c.most_recent_sql_handle)) as [sql_text]
	                , '0x'+CONVERT(varchar(max),r.plan_handle,2) as plan_handle
                FROM sys.dm_exec_connections as c with(nolock)
                JOIN sys.dm_exec_sessions as s with(nolock) on s.session_id = c.session_id
                JOIN sys.dm_exec_requests as r with(nolock) ON c.session_id = r.blocking_session_id
                WHERE c.session_id NOT IN (SELECT session_id from sys.dm_exec_requests WHERE session_id <> @@SPID AND session_id > 50)"

    $Conn = New-Object System.Data.SqlClient.SQLConnection
    $Conn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=Master;Integrated Security=True;Connect Timeout=30"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null
    $ProcessesInfo = @($Table.Tables[0])
    $Table.Dispose()

    $Script:Settings | Add-Member -Name ProcessesInfo -Value $($ProcessesInfo | Where { $_.Spid -in $SelectSpid } | Sort-Object $_.logical_reads -Descending) -MemberType NoteProperty
    
    if($Script:Settings.Blocking.Spid.Count -ne 0) {
        $Script:Settings | Add-Member -Name WaitTotal -Value $((($Script:Settings.ProcessesInfo | Measure-Object wait_time_ms -Max).Maximum)) -MemberType NoteProperty
    }
    else {
        $Script:Settings | Add-Member -Name WaitTotal -Value 0 -MemberType NoteProperty
    }

    SQL-BulkInsert AXMonitor_SQLRunningSpids @($Script:Settings.ProcessesInfo |
                                            Select @{n='Environment';e={$Script:Settings.Environment}}, 
                                            @{n='Start_Date_Time';e={$_.start_date_time}}, 
                                            @{n='SPID';e={$_.spid}}, 
                                            @{n='Blocker';e={$_.blocker}}, 
                                            @{n='Status';e={$_.status}}, 
                                            @{n='Host_Name';e={$_.host_name}}, 
                                            @{n='Context_Info';e={$_.context_info.Split('-')[0].Trim()}}, 
                                            @{n='Wait_Time_ms';e={$_.wait_time_ms}}, 
                                            @{n='Total_Time_ms';e={$_.total_time_ms}},
                                            @{n='Cpu_Time_ms';e={$_.cpu_time_ms}}, 
                                            @{n='Cpu_Time_Perc';e={$_.cpu_time_perc}}, 
                                            @{n='Reads';e={$_.reads}}, 
                                            @{n='Writes';e={$_.writes}}, 
                                            @{n='Logical_Reads';e={$_.logical_reads}}, 
                                            @{n='Wait_Type';e={$_.wait_type}}, 
                                            @{n='Database';e={$_.database}},
                                            @{n='Sql_text';e={$_.sql_text.Trim().Replace("'","''")}},
                                            @{n='Plan_Handle';e={$_.plan_handle}},
                                            @{n='GUID';e={($Script:Settings.Guid)}})
}

function Get-CPUStatus
{
    $CPUTotal = ([Math]::Round(((Get-Counter -Counter '\Processor(_total)\% Processor Time' -ComputerName $Script:Settings.NetBios -ErrorAction SilentlyContinue).CounterSamples).CookedValue,4))
    if($CPUTotal -le 0.0001) {
        $CPUTotal = ([Math]::Round(((Get-Counter -Counter '\Processor(_total)\% Processor Time' -ComputerName $Script:Settings.NetBios -SampleInterval 3 -ErrorAction SilentlyContinue).CounterSamples).CookedValue,4))
    }
    $Script:Settings | Add-Member -Name CPUTotal -Value $CPUTotal -MemberType NoteProperty
}

function Get-GRDStatus
{
    Get-CPUStatus
    Get-SQLStatus

    #Write-Host("$Step","Blocking Count - $($Script:Settings.Blocking.Spid.Count) / $($Script:Settings.BlockThold)")
    #Write-Host("$Step","Wait Time - $($Script:Settings.WaitTotal) / $($Script:Settings.WaitingThold)")
    #Write-Host("$Step","CPU% - $($Script:Settings.CPUTotal) / $($Script:Settings.CPUThold)")
    #Write-Host("$Step","GRD Flag is set to $($Script:Settings.EnableGRD)")
   
    if(($($Script:Settings.Blocking.Spid.Count) -ge $($Script:Settings.BlockThold)) -or 
        ($($Script:Settings.WaitTotal) -gt $($Script:Settings.WaitingThold)) -or 
        ($($Script:Settings.CPUTotal) -ge $($Script:Settings.CPUThold)))
    {
        $GRDRun = $true
    }
    else
    {
        $GRDRun = $false
    }
    SQL-BulkInsert AXMonitor_ExecutionLog @($Script:Settings | 
                                            Select @{n='Environment';e={$Environment}}, 
                                            @{n='CPU';e={$_.CPUTotal}}, 
                                            @{n='Blocking';e={($_.Blocking | Measure-Object).Count}}, #$foo | Measure-Object).Count
                                            @{n='Waiting';e={if(!($_.WaitTotal)){0} else{$_.WaitTotal}}}, 
                                            @{n='GRD';e={'0'}},
                                            @{n='GRDTotal';e={'0'}},
                                            @{n='Stats';e={'0'}},
                                            @{n='StatsTotal';e={'0'}},
                                            @{n='Email';e={'0'}},
                                            @{n='Report';e={''}},
                                            @{n='Log';e={"$($Script:Settings.CPUThold) | $($Script:Settings.BlockThold) | $($Script:Settings.WaitingThold) | $($Script:Settings.EnableGRD) | $($Script:Settings.EnableStats)"}},
                                            @{n='GUID';e={($_.Guid)}})
    return $GRDRun
}

function Get-AXJobs
{
    Write-ExecLog "Started Batch Jobs Status"
    $Conn = $Script:Settings.SQLServer.ConnectionContext.SqlConnectionObject
    $Query = 'SELECT Status = CASE STATUS 
		            WHEN 2 THEN ''Executing''
		            WHEN 3 THEN ''Error''
		            WHEN 7 THEN ''Cancelling''
	            END, Caption, 
	            DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()),  STARTDATETIME) AS StartDateTime, 
	            DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), ENDDATETIME) AS EndDateTime, 
	            CreatedBy 
            FROM BATCHJOB WITH (NOLOCK)
            WHERE STATUS IN (2,3,7)'
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null
    $Script:Settings | Add-Member -Name AXBatches -Value @($Table.Tables[0]) -MemberType NoteProperty
    $Table.Dispose()  

    SQL-BulkInsert AXMonitor_AXBatchJobs @($Script:Settings.AXBatches | 
                                        Select @{n='Environment';e={$Script:Settings.Environment}}, 
                                        @{n='StartDateTime';e={$_.StartDateTime}}, 
                                        @{n='EndDateTime';e={$_.EndDateTime}},
                                        @{n='Caption';e={$_.Caption}}, 
                                        @{n='Status';e={$_.Status}},
                                        @{n='CreatedBy';e={$_.CreatedBy}},
                                        @{n='GUID';e={($Script:Settings.Guid)}})

    Write-ExecLog "Started Number Sequences Status"
    $Query = 'SELECT C.NumberSequence, C.Txt, C.Format, 
                    Status = CASE B.STATUS
		                WHEN 0 THEN ''Free''
		                WHEN 1 THEN ''Active''
		                WHEN 2 THEN ''Blocked''
		                WHEN 3 THEN ''Reserved''
		            END, 
		            C.Continuous, B.SessionID, B.UserID, B.ModifiedBy, B.TransID,
		            DATEADD(MI, DATEDIFF(MI, GETUTCDATE(), GETDATE()),B.SESSIONLOGINDATETIME) AS SessionLoginDateTime, 
		            DATEADD(MI, DATEDIFF(MI, GETUTCDATE(), GETDATE()),B.MODIFIEDDATETIME) AS ModifiedDateTime
            FROM NUMBERSEQUENCETTS A WITH (NOLOCK)
            JOIN NUMBERSEQUENCELIST B WITH (NOLOCK)
                ON A.TRANSID = B.TRANSID
            JOIN NUMBERSEQUENCETABLE C WITH (NOLOCK)
                ON B.NUMBERSEQUENCEID = C.RECID'
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null
    $Script:Settings | Add-Member -Name AxNumSequences -Value @($Table.Tables[0]) -MemberType NoteProperty
    $Table.Dispose() 

    SQL-BulkInsert AXMonitor_AXNumberSequences @($Script:Settings.AxNumSequences | 
                                                Select @{n='Environment';e={$Script:Settings.Environment}}, 
                                                @{n='NumberSequence';e={$_.NumberSequence}}, 
                                                @{n='Txt';e={$_.Txt}},
                                                @{n='Format';e={$_.Format}}, 
                                                @{n='Status';e={$_.Status}},
                                                @{n='Continuous';e={$_.Continuous}},
                                                @{n='TransID';e={$_.TransID}},
                                                @{n='SessionID';e={$_.SessionID}},
                                                @{n='UserID';e={$_.UserID}},
                                                @{n='ModifiedBy';e={$_.ModifiedBy}},
                                                @{n='SessionLoginDateTime';e={$_.SessionLoginDateTime}},
                                                @{n='ModifiedDateTime';e={$_.ModifiedDateTime}},
                                                @{n='GUID';e={($Script:Settings.Guid)}})
    $Conn.Close()
}

function Get-SQLConfig
{
    Write-ExecLog "Started SQL Configuration"
    $Script:Settings | Add-Member -Name SQLInformation -Value $($Script:Settings.SQLServer.Information.Properties | Select Name, Value) -MemberType NoteProperty
    SQL-BulkInsert AXMonitor_SQLInformation @($Script:Settings.SQLInformation | 
                                                Select @{n='Environment';e={$Script:Settings.Environment}}, 
                                                @{n='Name';e={$_.Name}}, 
                                                @{n='Value';e={$_.Value.ToString()}},
                                                @{n='GUID';e={($Script:Settings.Guid)}})

    $Script:Settings | Add-Member -Name SQLConfiguration -Value $($Script:Settings.SQLServer.Configuration.Properties | Select DisplayName, Description, RunValue, ConfigValue) -MemberType NoteProperty
    SQL-BulkInsert AXMonitor_SQLConfiguration @($Script:Settings.SQLConfiguration | 
                                                    Select @{n='Environment';e={$Script:Settings.Environment}}, 
                                                    @{n='DisplayName';e={$_.DisplayName}}, 
                                                    @{n='Description';e={$_.Description}},
                                                    @{n='RunValue';e={$_.RunValue}},
                                                    @{n='ConfigValue';e={$_.ConfigValue}},
                                                    @{n='GUID';e={($Script:Settings.Guid)}})
}

function Get-PerfData
{
    Write-ExecLog "Started Perfmon Collectors"
    if(($Script:Settings.DBServer).Contains('\')) { $InstanceName = ($Script:Settings.DBServer).Split('\')[1] } else { $InstanceName = $Script:Settings.DBServer } 
    
    #-listset *

    $PerformanceCounters = '\Memory\Available MBytes', '\Processor(_total)\% Processor Time', "\MSSQL`$$InstanceName`:Buffer Manager\Buffer cache hit ratio", "\MSSQL`$$InstanceName`:Buffer Manager\Page life expectancy", 
                            "\MSSQL`$$InstanceName`:Locks(_Total)\Lock Waits/sec", "\MSSQL`$$InstanceName`:Locks(_Total)\Lock Wait Time (ms)", "\MSSQL`$$InstanceName`:Locks(_Total)\Number of Deadlocks/sec",
                            '\SQLServer:Buffer Manager\Buffer cache hit ratio', '\SQLServer:Buffer Manager\Page life expectancy', '\SQLServer:Buffer Manager\Lock Waits/sec', '\SQLServer:Buffer Manager\Lock Wait Time (ms)',
                            '\SQLServer:Buffer Manager\Number of Deadlocks/sec', '\LoicalDisk(*)\% Free Space','\LogicalDisk(*)\Free Megabytes', '\LogicalDisk(*)\Disk Transfers/sec', '\LogicalDisk(*)\Disk Reads/sec',
                            '\LogicalDisk(*)\Disk Writes/sec','\Paging File(_Total)\% Usage', '\Process(sqlservr*)\% Processor Time', '\Process(sqlservr*)\Virtual Bytes', 'Process(sqlservr*)\Working Set', 
                            '\Process(Ax32Serv*)\% Processor Time', '\Process(Ax32Serv*)\Virtual Bytes','Process(Ax32Serv*)\Working Set', '\Network Interface(*)\Bytes Total/sec', '\Network Interface(*)\Current Bandwidth',
                            '\Network Interface(*)\Bytes Received/sec', '\Network Interface(*)\Bytes Sent/sec', '\Network Interface(*)\Packets Outbound Discarded', '\Network Interface(*)\Packets Outbound Errors', 
                            '\Network Interface(*)\Output Queue Length', '\Network Interface(*)\TCP RSC Average Packet Size', '\TCPv4\Connections Established', '\TCPv4\Connections Active', '\TCPv4\Connections Passive',
                            '\TCPv4\Connection Failures', '\TCPv4\Connections Reset'
    try
    {
        $Perfmon = @()
        foreach($Counter in $PerformanceCounters)
        {
            $Perfmon += (Get-Counter -Counter $Counter -ComputerName $Script:Settings.NetBios -ErrorAction SilentlyContinue).CounterSamples | Select Path, @{n='Value';e={[Math]::Round(($_.CookedValue),2)}}, Timestamp 
        }

        $TotalMemory = Get-WmiObject -ClassName "Win32_ComputerSystem" -Namespace "root\CIMV2" -ComputerName $Script:Settings.NetBios | Measure-Object -Property TotalPhysicalMemory -Sum | Select Property, Count, Sum 
        $Script:Settings | Add-Member -Name MemoryTotal -Value ([Math]::Truncate($TotalMemory.Sum/1Gb)) -MemberType NoteProperty
        $Script:Settings | Add-Member -Name MemoryFree -Value ([Math]::Truncate((($Perfmon | Where {$_.Path -like '*memory\available mbytes'}).Value)/1024)) -MemberType NoteProperty
        $Script:Settings | Add-Member -Name MemoryLoad -Value ([Math]::Round((((($Perfmon | Where {$_.Path -like '*memory\available mbytes'}).Value) * 100) / ($TotalMemory.Sum/1Mb)),0)) -MemberType NoteProperty
    }
    catch
    {
        Write-Host("$Step","ERROR - $($_.Exception.Message)")
    }
    $Script:Settings | Add-Member -Name PerfmonData -Value $Perfmon -MemberType NoteProperty

    SQL-BulkInsert AXMonitor_PerfmonData @($Script:Settings.PerfmonData | 
                                            Select @{n='Environment';e={$Script:Settings.Environment}}, 
                                            @{n='Path';e={$_.Path}}, 
                                            @{n='Value';e={$_.Value}},
                                            @{n='Timestamp';e={$_.Timestamp}}, 
                                            @{n='GUID';e={($Script:Settings.Guid)}})
}

function Get-GRDTables
{
    if(($Script:Settings.ProcessesInfo.Spid.Count -eq 0) -and (!($Rerun))) {
        Write-ExecLog "Missing SQL Details. Starting new instance."
        Invoke-Expression "$scriptPath $($Script:Settings.Environment) -Rerun"
        Exit
    }

    Write-ExecLog "Started GRD (Guardian Defense - Processes $($Script:Settings.ProcessesInfo.Spid.Count))"

    if($Debug) { 
        $Script:Settings.Processes | Export-Csv $LogFolder\40-GRD_Processes_$($Environment)_$($FileDateTime).csv -NoTypeInformation -Append
        $Script:Settings.Blocking | Export-Csv $LogFolder\41-GRD_Blocking_$($Environment)_$($FileDateTime).csv -NoTypeInformation -Append
        $Script:Settings | Export-Csv $LogFolder\00-GRD_Settings_$($Environment)_$($FileDateTime).csv -NoTypeInformation -Append
        $Script:Settings.ProcessesInfo | Export-Csv $LogFolder\01-GRD_ProcessesInfo_$($Environment)_$($FileDateTime).csv -NoTypeInformation -Append
        $($Script:Settings.ProcessesInfo | Select-Object Sql_Text, Logical_Reads, Host_Name) | Export-Csv $LogFolder\10-GRD_SQLTextNoFilter_$($Environment)_$($FileDateTime).csv -NoTypeInformation -Append 
    }
    
    $SQLText = $Script:Settings.ProcessesInfo | 
                Where-Object {(($_.Database -eq $Script:Settings.DBName) -and 
                #($_.Sql_Text -notmatch 'FETCH*|DECLARE*|CREATE INDEX*|CREATE PROCEDURE*|UPDATE*|INSERT*|sp_*|exec*') -and
                #($_.Sql_Text -notmatch 'FETCH*|CREATE INDEX*') -and
                ($_.Status -ne 'sleeping') -and ($_.Sql_Text -notlike '') )} | 
                Select-Object Sql_Text #, Logical_Reads, Host_Name

    if($Debug) { $SQLText | Export-Csv $LogFolder\11-GRD_SQLTextFilter_$($Environment)_$($FileDateTime).csv -NoTypeInformation -Append }

    [Array]$Tables = $SQLText.Sql_Text.Split(' ') | % { (($_.Replace('[','')).Replace(']','')).Trim() } | 
                    Where-Object { 
                        ($_ -like "INVENT*") -or ($_ -like "SCANWORKX_BLOCKEDQTY") -or
                        ($_ -like "LOGISTICS*") -or ($_ -like "DIR*") -or ($_ -like "MFIDS*") -or
                        ($_ -like "CUSTT*") -or ($_ -like "CUSTS*") -or ($_ -like "VEND*") -or 
                        ($_ -like "RETAIL*") -or ($_ -like "DIMENSIONFOCUS*") -or ($_ -like "ECORESPRODUCT*") -or
                        ($_ -like "GENERALJOURNAL*") -or ($_ -like "SALES*") -or 
                        ($_ -match "GETPRODUCTS") -or ($_ -match "CUSTOMERSEARCH") -or ($_ -match "GETPARTYBYADDRESS") -or 
                        ($_ -match "GETPARTYBYCONTACT") -or ($_ -match "GETPARTYBYCUSTOMER") -or ($_ -match "GETPARTYBYLOYALTYCARD")
                    } | Select-Object -Unique

    if($Tables) {            
        if($Debug) { $Tables | Out-File $LogFolder\20-GRD_TablesFilter_$($Environment)_$($FileDateTime).txt -Append }

        if($Tables.Contains('SCANWORKX_BLOCKEDQTY')) {
            $ScanWorksSpids = $Script:Settings.ProcessesInfo | Where {$_.sql_text -match 'SCANWORKX_BLOCKEDQTY' }
            foreach($Spid in $ScanWorksSpids.Spid) {
                $ScanWorksBlock = $Script:Settings.ProcessesInfo | Where {$_.blocker -ne 0 -and $_.blocker -like $Spid}
                if(($ScanWorksBlock.Spid.Count -ge 1) -and ($($ScanWorksBlock.wait_time_ms | Measure -Average).Average -ge 300000)) {
                    Write-ExecLog "Killed Scanworks $($Spid) Count $($ScanWorksBlock.Spid.Count) Time $(($ScanWorksBlock.wait_time_ms | Measure -Average).Average)"
                    $Script:Settings.SQLServer.KillProcess($($Spid))
                } 
            }
        }

        if(($Tables.Contains('LOGISTICSCONTACTINFOVIEW')) -or 
                ($Tables.Contains('DIRPARTYPOSTALADDRESSVIEW')) -or 
                    ($Tables.Contains('LOGISTICSPOSTALADDRESSVIEW'))) {
            $Tables += 'LOGISTICSELECTRONICADDRESS'
            $Tables += 'LOGISTICSLOCATION'
            $Tables += 'LOGISTICSPOSTALADDRESS'
            $Tables += 'DIRPARTYTABLE'
            $Tables += 'DIRPARTYLOCATION'
        }

        if($Tables.Contains('DIMENSIONFOCUSBALANCECALCULATIONVIEW')) {
            $Tables += 'DIMENSIONFOCUSUNPROCESSEDTRANSACTIONS'
            $Tables += 'DIMENSIONFOCUSLEDGERDIMENSIONREFERENCE'
            $Tables += 'GENERALJOURNALACCOUNTENTRY'
            $Tables += 'DIMENSIONFOCUSBALANCE'
            $Tables += 'GENERALJOURNALENTRY'
            $Tables += 'FISCALCALENDARPERIOD'
        }

        if($Tables -match 'GETPRODUCTS') {
            $Tables += 'ax.INVENTTABLE'
            $Tables += 'ax.INVENTDIMCOMBINATION'
            $Tables += 'ax.ECORESDISTINCTPRODUCTVARIANT'
            $Tables += 'ax.ECORESPRODUCT'
            $Tables += 'ax.ECORESPRODUCTIMAGE'
            $Tables += 'ax.ECORESTEXTVALUE'
            $Tables += 'ax.ECORESPRODUCTTRANSLATION'
            $Tables += 'ax.ECORESATTRIBUTE'
            $Tables += 'ax.RETAILPUBCATALOGPRODUCT'
            $Tables += 'ax.RETAILPUBCATALOG'
            $Tables += 'ax.RETAILCHANNELTABLE'
            $Tables += 'ax.RETAILKIT'
            $Tables += 'ax.RETAILPUBPRODUCTATTRIBUTECHANNELMETADATA'
        }

        if($Tables -match 'CUSTOMERSEARCH') {
            $Tables += 'ax.DIRPARTYTABLE'
            $Tables += 'ax.DIRPARTYLOCATION'
            $Tables += 'ax.LOGISTICSELECTRONICADDRESS'
            $Tables += 'ax.LOGISTICSPOSTALADDRESS'
            $Tables += 'ax.LOGISTICSLOCATION'
            $Tables += 'ax.DIRADDRESSBOOKPARTY'
            $Tables += 'ax.RETAILSTOREADDRESSBOOK'
            $Tables += 'ax.RETAILLOYALTYCARD'
        }

        if($Tables -match 'GETPARTYBYADDRESS') {
            $Tables += 'ax.LOGISTICSPOSTALADDRESS'
            $Tables += 'ax.DIRPARTYLOCATION'
            $Tables += 'ax.DIRPARTYTABLE'
        }

        if($Tables -match 'GETPARTYBYCONTACT') {
            $Tables += 'ax.LOGISTICSELECTRONICADDRESS'
            $Tables += 'ax.DIRPARTYLOCATION'
            $Tables += 'ax.DIRPARTYTABLE'
        }

        if($Tables -match 'GETPARTYBYCUSTOMER') {
            $Tables += 'ax.CUSTTABLE'
            $Tables += 'ax.DIRADDRESSBOOKPARTY'
            $Tables += 'ax.DIRPARTYTABLE'
            $Tables += 'ax.RETAILSTOREADDRESSBOOK'
        }

        if($Tables -match 'GETPARTYBYLOYALTYCARD') {
            $Tables += 'ax.RETAILLOYALTYCARD'
        }

        if($Debug) { $Tables | Out-File $LogFolder\21-GRD_TablesViews_$($Environment)_$($FileDateTime).txt -Append }

        $Tables = $Tables | Where-Object { ($_ -notlike "*SCANWORKX_BLOCKEDQTY*") } | Select-Object -Unique
        $Tables = $Tables | Where-Object { ($_ -notlike "*VIEW") } | Select-Object -Unique
        $Tables = $Tables | Where-Object { ($_ -notlike "*CRT*") } | Select-Object -Unique

        if($Debug) { $Tables | Out-File $LogFolder\22-GRD_TablesFinal_$($Environment)_$($FileDateTime).txt -Append }

        $TablesRet = GRD-CheckTables $Tables

        if($Debug) { $TablesRet | Out-File $LogFolder\23-GRD_TablesRetCheck_$($Environment)_$($FileDateTime).txt -Append }

        if($TablesRet) {
            GRD-StartJobs $TablesRet
        }
    }
    else {
        if($Debug) { $Tables | Out-File $LogFolder\24-GRD_NoTables_$($Environment)_$($FileDateTime).txt -Append }
    }
}

function GRD-CheckTables
{
Param(
    [array]$TblArray
)
    $TablesRet = @()
    foreach($Table in $TblArray) {
        if($Table.Split('.').Count -gt 1) {
            $Schema = $Table.Split('.')[0]
            $TableName = $Table.Split('.')[1]
        }
        else {
            $Schema = 'dbo'
            $TableName = $Table.ToUpper()
        }
        if($Script:Settings.SQLServer.Databases[$Script:Settings.DBName].Tables.Contains($TableName, $Schema)) {
            $TablesRet += "$Schema.$TableName"
        }
        else {
            foreach($Schema in ($Script:Settings.SQLServer.Databases[$Script:Settings.DBName].Schemas | Where {$_.Owner -match 'dbo' -and $_.Name -notmatch 'dbo'}).Name) {
                if($Script:Settings.SQLServer.Databases[$Script:Settings.DBName].Tables.Contains($TableName, $Schema) -and (!($TablesRet.Contains("$Schema.$TableName")))) { 
                    $TablesRet += "$Schema.$TableName"
                }
            }
        }
    }
    return $TablesRet
}

function GRD-StartJobs
{
Param(
    [array]$Tables
)
    SQL-ExecUpdate "UPDATE AXMonitor_ExecutionLog SET GRD = 1, GRDTOTAL = $($Tables.Count), STATS = 0 WHERE GUID = '$($Script:Settings.Guid)'"
    $GRDExp = @()
    $Query =   "SELECT [STARTED]
                        , [FINISHED]
                        , [TABLENAME]
                        , [STATSTYPE]
                        , [GUID]
            FROM [AXMonitor_GRDLog] 
            WHERE [ENVIRONMENT] = '$Environment' AND
		            [CREATEDDATETIME] >= '$((Get-Date).AddMinutes(-60))'
                    AND [STATSTYPE] <> 'GRD'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:Settings.ToolsConnection)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null
    $GRDStats = @($Table.Tables[0])
    $Table.Dispose()
    $GRDJobRun = @()

    foreach($Table in $Tables) {
        $GRDJobTemp = New-Object -TypeName System.Object
        if(($GRDStats | Where-Object { $_.TableName -like $Table -and $_.Finished -like $null }) -and ($Script:Settings.Processes | WHERE { $_.Command -match 'UPDATE STATISTICS' })) {
            $Query = "SELECT session_id, start_time, status, text FROM sys.dm_exec_requests CROSS APPLY sys.dm_exec_sql_text(sql_handle) WHERE session_id IN ($(($Script:Settings.Processes | WHERE { $_.Command -match 'UPDATE STATISTICS' }).Spid -join ','))"
            $DataSet = $Script:Settings.SQLServer.Databases[$Script:Settings.DBName].ExecuteWithResults($Query)
            if($DataSet.Tables[0] | Where { $_.Text -match $Table}) {
                $GRDJobTemp | Add-Member -Name TableName -Value $Table -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name StatsType -Value "GRD" -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name Statement -Value "Already Running $($DataSet.Tables[0].session_id) - $($DataSet.Tables[0].text)" -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name Started -Value $DataSet.Tables[0].start_time -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name Finished -Value $(Get-Date) -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name GRDJobName -Value "SPID$($DataSet.Tables[0].session_id)_$($DataSet.Tables[0].status.ToUpper())_$(($GRDStats | Where-Object { $_.TableName -like $Table -and $_.Finished -like $null }).Guid)" -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name GRDJobRun -Value 0 -MemberType NoteProperty
                        
            }
            else {
                $GRDJobTemp | Add-Member -Name TableName -Value $Table -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name StatsType -Value 'REGULAR' -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name Statement -Value "UPDATE STATISTICS $Table" -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name Started -Value $(Get-Date) -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name GRDJobName -Value $("GRD_$($Script:Settings.Environment)_$($Table.ToUpper())`_$(Get-Date -f yyyyMMddHHmm)") -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name GRDJobRun -Value 1 -MemberType NoteProperty
            }
        }
        elseif(($GRDStats | Where-Object { $_.TableName -like $Table -and $_.StatsType -like 'REGULAR' }).Count -ge 3) {
            $GRDJobTemp | Add-Member -Name TableName -Value $Table -MemberType NoteProperty
            $GRDJobTemp | Add-Member -Name StatsType -Value 'FULLSCAN' -MemberType NoteProperty
            $GRDJobTemp | Add-Member -Name Statement -Value "UPDATE STATISTICS $Table WITH FULLSCAN" -MemberType NoteProperty
            $GRDJobTemp | Add-Member -Name Started -Value $(Get-Date) -MemberType NoteProperty
            $GRDJobTemp | Add-Member -Name GRDJobName -Value $("GRD_$($Script:Settings.Environment)_$($Table.ToUpper())`_$(Get-Date -f yyyyMMddHHmm)") -MemberType NoteProperty
            $GRDJobTemp | Add-Member -Name GRDJobRun -Value 1 -MemberType NoteProperty
        }
        else {
            $GRDJobTemp | Add-Member -Name TableName -Value $Table -MemberType NoteProperty
            $GRDJobTemp | Add-Member -Name StatsType -Value 'REGULAR' -MemberType NoteProperty
            $GRDJobTemp | Add-Member -Name Statement -Value "UPDATE STATISTICS $Table" -MemberType NoteProperty
            $GRDJobTemp | Add-Member -Name Started -Value $(Get-Date) -MemberType NoteProperty
            $GRDJobTemp | Add-Member -Name GRDJobName -Value $("GRD_$($Script:Settings.Environment)_$($Table.ToUpper())`_$(Get-Date -f yyyyMMddHHmm)") -MemberType NoteProperty
            $GRDJobTemp | Add-Member -Name GRDJobRun -Value 1 -MemberType NoteProperty
        }
        SQL-BulkInsert AXMonitor_GRDLog @($GRDJobTemp | 
                                                Select @{n='Environment';e={$Script:Settings.Environment}}, 
                                                @{n='TableName';e={$_.TableName}}, 
                                                @{n='StatsType';e={$_.StatsType}}, 
                                                @{n='Statement';e={$_.Statement}}, 
                                                @{n='JobName';e={$_.GRDJobName}},
                                                @{n='Started';e={$_.Started}},
                                                @{n='Finished';e={if($_.Finished){$_.Finished} else{$null}}},
                                                @{n='GUID';e={($Script:Settings.Guid)}})
                
        if($GRDJobTemp.GRDJobRun -eq 1) {Run-GRDStats $GRDJobTemp}
        if($Debug) { $GRDJobTemp | Export-Csv $LogFolder\30-GRD_JobsCreation_$($Environment)_$($FileDateTime).csv -NoTypeInformation -Append }
        $GRDJobRun += $GRDJobTemp
    }
    $Script:Settings | Add-Member -Name GRDJobs -Value $GRDJobRun -MemberType NoteProperty
    if($Debug) { $(Get-Job) | Export-Csv $LogFolder\31-GRD_Jobs_$($Environment)_$($FileDateTime).csv -NoTypeInformation -Append }
}

function Run-GRDStats
{
    if($Script:Settings.SqlCredential) {
        Start-Job -Credential $Script:Settings.SqlCredential -Name $($GRDJobTemp.GRDJobName) -ScriptBlock {& $args[0] $args[1] $args[2] $args[3] $args[4] $args[5]} -ArgumentList @("$ScriptDir\AX-UpdateStats.ps1"), $($Script:Settings.DBServer), $($Script:Settings.DBName), $($GRDJobTemp.TableName), $($GRDJobTemp.StatsType), $($GRDJobTemp.GRDJobName)
    }
    else {
        Start-Job -Name $($GRDJobTemp.GRDJobName) -ScriptBlock {& $args[0] $args[1] $args[2] $args[3] $args[4] $args[5]} -ArgumentList @("$ScriptDir\AX-UpdateStats.ps1"), $($Script:Settings.DBServer), $($Script:Settings.DBName), $($GRDJobTemp.TableName), $($GRDJobTemp.StatsType), $($GRDJobTemp.GRDJobName)
    }
}

function Get-JobStatus
{
    While ($(Get-Job).Count -gt 0) {
        Start-Sleep -Milliseconds 3000
        $GRDJobs = Get-Job | Where-Object { $_.State -ne 'Running' }
        foreach ($Job in $GRDJobs) {
            if($Job.State -eq 'Failed') {
                SQL-UpdateTable 'AXMonitor_GRDLog' 'FINISHED' $($Job.PSEndTime) "JOBNAME = '$($Job.Name)' AND GUID = '$($Script:Settings.Guid)'"
                SQL-UpdateTable 'AXMonitor_GRDLog' 'LOG' $((Get-Job $Job.Name | Receive-Job 2>&1).Exception.Message) "JOBNAME = '$($Job.Name)' AND GUID = '$($Script:Settings.Guid)'"
            }
            else {
                SQL-UpdateTable 'AXMonitor_GRDLog' 'FINISHED' $($Job.PSEndTime) "JOBNAME = '$($Job.Name)' AND GUID = '$($Script:Settings.Guid)'"
            }
            Remove-Job –Name $($Job.Name)
        }
        Start-Sleep -Milliseconds 2000
        $Script:Settings | Add-Member -Name Processes -Value $($Script:Settings.SQLServer.EnumProcesses() | Where { $_.Spid -gt 50 }) -MemberType NoteProperty -Force
        $GRDSpids = ($Script:Settings.Processes | WHERE { $_.Command -match 'UPDATE STATISTICS' -and $_.Database -match $Script:Settings.DBName -and $_.Host -match $env:COMPUTERNAME } | Sort CPU -Descending)
        foreach($Spid in $GRDSpids) {
            if($Script:Settings.Processes.BlockingSpid -eq $Spid.Spid ) { 
                $BlockedSpid = $Script:Settings.Processes | Where {$_.BlockingSpid -eq $Spid.Spid}
                $Query = "SELECT session_id, start_time, status, text FROM sys.dm_exec_requests CROSS APPLY sys.dm_exec_sql_text(sql_handle) WHERE session_id = $($Spid.Spid)"
                $DataSet = $Script:Settings.SqlServer.Databases[$Script:Settings.DBName].ExecuteWithResults($Query)
                Write-ExecLog "Killed $($Spid.Spid)-$($DataSet.Tables[0].Text) blocking $($BlockedSpid.Spid)-$($BlockedSpid.Host)"
                $Script:Settings.SQLServer.KillProcess($($Spid.Spid))
            }
        }

    }
}

function Get-SQLStatisticsInterval
{
    $Query =   "SELECT TOP 1 MAX(CREATEDDATETIME)
                FROM [dbo].[AXMonitor_GRDStatistics]
                WHERE [ENVIRONMENT] = '$($Script:Settings.Environment)'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:Settings.ToolsConnection)
    if([String]::IsNullOrEmpty($Cmd.ExecuteScalar())) { $CreatedDateTime = Get-Date('1/1/1900') } else { $CreatedDateTime = $Cmd.ExecuteScalar() }
    if([Math]::Truncate((New-TimeSpan ($CreatedDateTime) $(Get-Date)).TotalMinutes) -ge $Script:Configuration.Settings.AXMonitor.StatisticsCheckInterval) {
        return $true
    }
    else {
        return $false
    }
}

function Get-SQLStatistics
{
    $Conn = $Script:Settings.SQLServer.ConnectionContext.SqlConnectionObject
    $Query = "SELECT  o.name as [TableName]
                    , object_schema_name(object_id) as [Schema]
		            , si.name as [IndexName]
		            , si.indid as [IndexID]
		            , si.rowcnt as [RowsTotal]
		            , si.rowmodctr [RowsModified]
		            , cast(1.0*reserved*8/1024 as decimal(20,2)) as [SizeMB]
		            , cast(1.0*si.rowmodctr/(si.rowcnt)*100 as decimal(10)) as [PercentChange]
		            , stats_date(si.id, si.indid) as [LastUpdate]
            FROM sys.sysindexes si 
            JOIN sys.objects o ON si.id = o.object_id
            WHERE o.object_id > 10000 --displays only user tables
                   AND o.type = 'U'
                   AND si.rowcnt > 1000 
                   AND (((sqrt(1000 * si.rowcnt) < 1.0*si.rowmodctr/(si.rowcnt)) 
	               AND sqrt(1000 * si.rowcnt) > si.rowmodctr)
	               OR (sqrt(1000 * si.rowcnt) > 1.0*si.rowmodctr/(si.rowcnt)
	               AND 1.0*si.rowmodctr/(si.rowcnt) > 0.05))--tables with more than 5% changes
                   AND si.indid > 0 --exclude heaps(value 0)"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null
    $GRDStats = @($Table.Tables[0])
    $Table.Dispose()       

    SQL-BulkInsert AXMonitor_GRDStatistics @($GRDStats |  
                                                Select @{n='Environment';e={$Script:Settings.Environment}}, 
                                                @{n='TableName';e={"$($_.Schema).$($_.TableName)"}}, 
                                                @{n='IndexName';e={$_.IndexName}},
                                                @{n='IndexID';e={$_.IndexID}},
                                                @{n='RowsTotal';e={$_.RowsTotal}},
                                                @{n='RowsModified';e={$_.RowsModified}},
                                                @{n='SizeMB';e={$_.SizeMB}},
                                                @{n='PercentChange';e={$_.PercentChange}},
                                                @{n='LastUpdate';e={$_.LastUpdate}},
                                                @{n='GUID';e={($Script:Settings.Guid)}})
    
    SQL-ExecUpdate "UPDATE AXMonitor_ExecutionLog SET STATSTOTAL = $(($GRDStats | Where {$_.PercentChange -gt $Script:Configuration.Settings.AXMonitor.StatisticsPercentChange}).Count) WHERE GUID = '$($Script:Settings.Guid)'"

    $Query =   "SELECT TOP 1 MAX(CREATEDDATETIME)
                FROM [dbo].[AXMonitor_GRDLog]
                WHERE [JOBNAME] LIKE 'SQL_%' 
                AND [ENVIRONMENT] = '$($Script:Settings.Environment)'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:Settings.ToolsConnection)
    if([String]::IsNullOrEmpty($Cmd.ExecuteScalar())) { $CreatedDateTime = Get-Date('1/1/1900') } else { $CreatedDateTime = $Cmd.ExecuteScalar() }
    if([Math]::Truncate((New-TimeSpan ($CreatedDateTime) $(Get-Date)).TotalMinutes) -ge $Script:Configuration.Settings.AXMonitor.StatisticsUpdateInterval) {
        $GRDStatsCount = $true
    }
    else {
        $GRDStatsCount = $false
    }

    $GRDJobRun = @()
    if(($Script:Settings.EnableStats -eq 2) -and ($GRDStatsCount) -and ($(($GRDStats | Where {$_.PercentChange -gt $Script:Configuration.Settings.AXMonitor.StatisticsPercentChange}).Count) -gt 0) -and ($Script:Settings.CPUTotal -le $Script:Configuration.Settings.AXMonitor.StatisticsUpdateCpuMax)) {
        SQL-ExecUpdate "UPDATE AXMonitor_ExecutionLog SET STATS = '1' WHERE GUID = '$($Script:Settings.Guid)'"
        foreach($Table in $($GRDStats | Where {$_.PercentChange -gt $Script:Configuration.Settings.AXMonitor.StatisticsPercentChange} | Group Schema,TableName | Sort Count -Descending | Select -First $Script:Configuration.Settings.AXMonitor.StatisticsUpdateTop)) {
            While ($(Get-Job -state Running).Count -ge 5){
                Start-Sleep -Milliseconds 5000
            }
                $GRDJobTemp = New-Object -TypeName System.Object
                $GRDJobTemp | Add-Member -Name TableName -Value $Table.Name.Replace(', ','.') -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name StatsType -Value 'REGULAR' -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name Statement -Value "UPDATE STATISTICS $($Table.Name.Replace(', ','.'))" -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name Started -Value $(Get-Date) -MemberType NoteProperty
                $GRDJobTemp | Add-Member -Name GRDJobName -Value $("SQL_$($Script:Settings.Environment)_$($Table.Name.Replace(', ','.').ToUpper())`_$(Get-Date -f yyyyMMddHHmm)") -MemberType NoteProperty                    
                SQL-BulkInsert AXMonitor_GRDLog @($GRDJobTemp | 
                                                    Select @{n='Environment';e={$Script:Settings.Environment}}, 
                                                    @{n='TableName';e={$_.TableName}}, 
                                                    @{n='StatsType';e={$_.StatsType}}, 
                                                    @{n='Statement';e={$_.Statement}}, 
                                                    @{n='JobName';e={$_.GRDJobName.ToUpper()}},
                                                    @{n='Started';e={$_.Started}},
                                                    @{n='Finished';e={$null}},
                                                    @{n='GUID';e={($Script:Settings.Guid)}})
                Run-GRDStats $GRDJobTemp
            $GRDJobRun += $GRDJobTemp
        }
        $Script:Settings | Add-Member -Name GRDJobs -Value $GRDJobRun -MemberType NoteProperty    
    }
}

function Get-CreateReport
{
    $GRDReport = @()
    $GRDReport += Get-HtmlOpen -TitleText ("SQL Monitoring Alert $($Script:Settings.DBServer) @ $($Script:Settings.NetBios)") -SimpleHTML

    $GRDSummary = @()
    
    if($Script:Settings.CPUTotal -gt $Script:Settings.CPUThold) {
        $GRDSummary += @($Script:Settings | Select @{n='Name';e={'CPU %'}}, @{n='Value';e={$_.CPUTotal}}, @{n='RowColor';e={'Red'}})
    }
    else {
        $GRDSummary += @($Script:Settings | Select @{n='Name';e={'CPU %'}}, @{n='Value';e={$_.CPUTotal}}, @{n='RowColor';e={'None'}})
    }

    if($Script:Settings.Blocking.Count -gt $Script:Settings.BlockThold) {
        $GRDSummary += @($Script:Settings | Select @{n='Name';e={'Blocking'}}, @{n='Value';e={$_.Blocking.Spid.Count}}, @{n='RowColor';e={'Red'}})
    }
    else {
        $GRDSummary += @($Script:Settings | Select @{n='Name';e={'Blocking'}}, @{n='Value';e={$_.Blocking.Spid.Count}}, @{n='RowColor';e={'None'}})
    }

    if($Script:Settings.WaitTotal -gt $Script:Settings.WaitingThold) {
        $GRDSummary += @($Script:Settings | Select @{n='Name';e={'Waiting Time'}}, @{n='Value';e={$_.WaitTotal}}, @{n='RowColor';e={'Red'}})
    }
    else {
        $GRDSummary += @($Script:Settings | Select @{n='Name';e={'Waiting Time'}}, @{n='Value';e={$_.WaitTotal}}, @{n='RowColor';e={'None'}})
    }
    
    if($Script:Settings.HeadBlockers) {
        $GRDSummary += @($Script:Settings | Select @{n='Name';e={'Head Blockers'}}, @{n='Value';e={$_.HeadBlockers -join ', '}}, @{n='RowColor';e={'None'}})
    }
    
    $GRDSummary += @($Script:Settings | Select @{n='Name';e={'Total Memory'}}, @{n='Value';e={$_.MemoryTotal}}, @{n='RowColor';e={'None'}})
    $GRDSummary += @($Script:Settings | Select @{n='Name';e={'Free Memory'}}, @{n='Value';e={$_.MemoryFree}}, @{n='RowColor';e={'None'}})
    
    if($Script:Settings.MemoryLoad -ge 75) {
        $GRDSummary += @($Script:Settings | Select @{n='Name';e={'Memory Load %'}}, @{n='Value';e={$_.MemoryLoad}}, @{n='RowColor';e={'LightRed'}})
    }
    elseif(($Script:Settings.MemoryLoad -ge 50) -and ($Script:Settings.MemoryLoad -lt 75)) {
        $GRDSummary += @($Script:Settings | Select @{n='Name';e={'Memory Load %'}}, @{n='Value';e={$_.MemoryLoad}}, @{n='RowColor';e={'LightYellow'}})
    }
    else {
        $GRDSummary += @($Script:Settings | Select @{n='Name';e={'Memory Load %'}}, @{n='Value';e={$_.MemoryLoad}}, @{n='RowColor';e={'LightGreen'}})
    }
    
    $GRDReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "GRD Threshold"
    $GRDReport += Get-HtmlContentTable($GRDSummary | Select Name, Value, RowColor)
    $GRDReport += Get-HtmlContentClose

    if($Script:Settings.GRDJobs) {
        $GRDReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "GRD Jobs"
        $GRDReport += Get-HtmlContentTable($Script:Settings.GRDJobs | Select @{n='Table';e={$_.TableName}}, Statement, Started)
        $GRDReport += Get-HtmlContentClose
    }
    if($Script:Settings.AXBatches) {
        $GRDReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "Active AX Batch Jobs"
        $GRDReport += Get-HtmlContentTable($Script:Settings.AXBatches | Select Status, @{n='Name';e={$_.Caption}}, @{n='Started';e={$_.StartDateTime}}, CreatedBy)
        $GRDReport += Get-HtmlContentClose
    }
    if($Script:Settings.AxNumSequences) { 
        $GRDReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "Active AX Number Sequences"
        $GRDReport += Get-HtmlContentTable($Script:Settings.AxNumSequences | Select NumberSequence, Format, Status, Continuous, SessionID, UserID, ModifiedBy, TransID, @{n='SessionTime';e={$_.SessionLoginDateTime}})
        $GRDReport += Get-HtmlContentClose
    }
    if($Script:Settings.ProcessesInfo) {
        $GRDReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "Running Queries"
        $GRDReport += Get-HtmlContentTable($Script:Settings.ProcessesInfo | Select  @{n='Database';e={$_.Database}}, 
                                                                                    @{n='HostName';e={$_.Host_Name}},
                                                                                    @{n='User';e={$_.context_info.Split('-')[0].Trim()}}, 
                                                                                    @{n='Spid';e={$_.SPID}},
                                                                                    @{n='Blocker';e={$_.Blocker}},
                                                                                    @{n='Status';e={$_.Status}}, 
                                                                                    @{n='Logical Reads';e={$_.Logical_Reads}}, 
                                                                                    @{n='Reads';e={$_.Reads}}, 
                                                                                    @{n='Writes';e={$_.Writes}}, 
                                                                                    @{n='Wait Type';e={$_.Wait_Type}}, 
                                                                                    @{n='Wait Time';e={$_.Wait_Time}}, 
                                                                                    @{n='Total Time';e={$_.Total_Time}}, 
                                                                                    @{n='Cpu Time';e={$_.Cpu_Time}}, 
                                                                                    @{n='CPU %';e={$_.Cpu_Time_Perc}}, 
                                                                                    @{n='Query';e={$_.sql_text.Trim().Replace("'","''")}} | 
                                                                                    Sort-Object 'logical reads' -Descending)
        $GRDReport += Get-HtmlContentClose
    }
    if($Script:Settings.SQLConfiguration) {
        $GRDReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "SQL Configuration"
        $GRDReport += Get-HtmlContentTable($Script:Settings.SQLConfiguration | Select @{n='Name';e={$_.DisplayName}}, Description, RunValue, ConfigValue | Sort-Object DisplayName)
        $GRDReport += Get-HtmlContentClose
    }
    if($Script:Settings.SQLInformation) {
        $GRDReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "SQL Information"
        $GRDReport += Get-HtmlContentTable($Script:Settings.SQLInformation | Select Name, Value | Sort-Object Name)
        $GRDReport += Get-HtmlContentClose
    }
    if($Script:Settings.PerfmonData) {
        $GRDReport += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText "Performance Counters"
        $GRDReport += Get-HtmlContentTable($Script:Settings.PerfmonData | Select @{n='Counter';e={$_.Path}}, Value | Sort-Object Path)
        $GRDReport += Get-HtmlContentClose
    }
    $GRDReport += Get-HtmlClose -FooterText "Guid: $($Script:Settings.Guid)"
    $Script:Settings | Add-Member -Name GRDReport -Value $GRDReport -MemberType NoteProperty

    #Save HTML
    $GRDReportPath = join-path $ReportFolder ("GRD-$($Script:Settings.NetBios)-$FileDateTime" + ".html")
    $GRDReport | Set-Content -Path $GRDReportPath -Force
    $Script:Settings | Add-Member -Name GRDReportPath -Value $GRDReportPath -MemberType NoteProperty
    SQL-ExecUpdate "UPDATE AXMonitor_ExecutionLog SET REPORT = '$GRDReportPath' WHERE GUID = '$($Script:Settings.Guid)'"
}

function Get-SendEmail
{
    $Query =   "SELECT TOP 1 MAX(CREATEDDATETIME)
                    FROM AXMonitor_ExecutionLog
                    WHERE [EMAIL] = 1 AND [ENVIRONMENT] = '$($Script:Settings.Environment)'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:Settings.ToolsConnection)
    if([String]::IsNullOrEmpty($Cmd.ExecuteScalar())) { $CreatedDateTime = Get-Date('1/1/1900') } else { $CreatedDateTime = $Cmd.ExecuteScalar() }
    if([Math]::Truncate((New-TimeSpan ($CreatedDateTime) $(Get-Date)).TotalMinutes) -ge $Script:Configuration.Settings.AXMonitor.SendEmailLowRiskInterval) {
        $GRDReportChk = $true
    }
    else {
        $GRDReportChk = $false
    }

    if(($($Script:Settings.CPUTotal) -le $Script:Settings.CPUThold) -and
        ($($Script:Settings.Blocking.Spid.Count) -le $($Script:Settings.BlockThold)) -and
        (!$GRDReportChk)) {
        Write-ExecLog "Email Suppressed"
        continue
    }
    else {
        if($Script:Settings.GRDJobs.Count -ge 1) { $Subject = "[GRD] SQL Server Blocking Report $($Script:Settings.Description)" }
        elseif($Script:Settings.CPUTotal -gt $Script:Settings.CPUThold) { $Subject = "[CPU] SQL Server Blocking Report $($Script:Settings.Description)" }
        else { $Subject = "SQL Server Blocking Report $($Script:Settings.Description)" }
        Send-Email -Subject $Subject -Body $Script:Settings.GRDReport -Attachment $Script:Settings.GRDReportPath -EmailProfile $Script:Settings.EmailProfile -GUID $Script:Settings.Guid
    }
}

function Do-Cleanup
{
    $Files = Get-ChildItem -Path $ReportFolder | Where { $_.LastWriteTime -lt $((Get-Date).AddDays((-$Script:Configuration.Settings.General.RetentionDays))) -and $_.Name -like "GRD-$($Script:Settings.NetBios)*" }
    if($Files) {
        Remove-Item -Path $Files.FullName -Force
    }
    Write-ExecLog("Cleanup $($Files.Count) files, done.")  
}

function GRD-RunCheck 
{
    $blocks = $server.EnumProcesses() | Where-Object { $_.blockingspid -ne 0 }
    foreach ($block in $blocks) {
        $blockedby = $server.EnumProcesses() | Where-Object { $_.spid -eq $block.blockingspid }
        $db = $server.Databases[$blockedby.database]
        $blockingTransaction =  $db.EnumTransactions() | Where-Object { $_.spid -eq $blockedby.spid }
        $currentime = Get-Date
        $transactionBegin = $blockingTransaction.BeginTime
        $blockingSeconds = ($currentime - $transactionBegin).TotalSeconds
        if ($blockingSeconds -gt $threshhold) {
            # Could not find pure SMO for dm_exec_sql_text :|
            $sql = "SELECT top 1 text as statement FROM sys.dm_exec_requests
            CROSS APPLY sys.dm_exec_sql_text(sql_handle) WHERE session_id = $($blockedby.spid)"
            $dataset = $db.ExecuteWithResults($sql)
            $blockingstatement = $dataset.Tables[0].Rows[0].Item(0)

            $object = New-Object PSObject -Property @{
            User = $blockedby.login
            Host = $blockedby.host
            Program = $blockedby.program
            Command = $blockedby.command
            Statement = $blockingstatement
            BlockingSince = $transactionBegin
            BlockingSeconds = [Math]::Round($blockingSeconds,1)
            }
            $blockingCollection += $object
        }
    }
    $blockingCollection
}

function Write-ExecLog
{
param (
    [String]$MSg
)
    SQL-ExecUpdate "UPDATE AXMonitor_ExecutionLog SET LOG = CASE WHEN LOG = '' THEN '$Msg' ELSE LOG + ' | ' +  '$Msg' END WHERE GUID = '$($Script:Settings.Guid)'"
}

function Check-Folder {
param(
    [String]$Path
)
    if(!(Test-Path($Path))) {
        New-Item -ItemType Directory -Force -Path $Path | Out-Null
    }
}

Check-Folder $ReportFolder
Check-Folder $LogFolder

Get-SQLMonitoring
Get-Module | Where-Object {$_.ModuleType -eq 'Script'} | % { Remove-Module $_.Name }