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

[CmdletBinding()]
Param (
    [Parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true)]
    [String]$Environment,
    [Parameter(Position=1,ParameterSetName="RecycleBlg",Mandatory=$false,ValueFromPipeline=$true)]
    [Switch]$RecycleBlg
)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"

Import-Module $ModuleFolder\AX-Tools.psm1 -DisableNameChecking

$Script:Configuration = Load-ConfigFile
$ReportFolder = if(!$Script:Configuration.Settings.General.ReportPath) { $Dir + "\Reports\AX-Report\$Environment" } else { "$($Script:Configuration.Settings.General.ReportPath)\$Environment" }
$LogFolder = if(!$Script:Configuration.Settings.General.LogPath) { $Dir + "\Logs\AX-Report\$Environment" } else { "$($Script:Configuration.Settings.General.LogPath)\$Environment" }
$LogFilesDays = $Script:Configuration.Settings.General.RetentionDays
$AutoCleanUp = [boolean]::Parse($Script:Configuration.Settings.General.AutoCleanUp)

$Global:Guid = ([Guid]::NewGuid()).Guid
$Script:Settings = New-Object -TypeName System.Object
$Script:Settings | Add-Member -Name GUID -Value $Global:Guid -MemberType NoteProperty
$Script:Settings | Add-Member -Name ReportDate -Value $(Get-Date (Get-Date).AddDays(-1) -format d) -MemberType NoteProperty
$Script:Settings | Add-Member -Name Environment -Value $Environment -MemberType NoteProperty
$Script:Settings | Add-Member -Name DataCollectorName -Value $Script:Configuration.Settings.General.PerfmonCollectorName -MemberType NoteProperty
$Script:Settings | Add-Member -Name ApplicationName -Value 'AX Report Script' -MemberType NoteProperty
$Script:Settings | Add-Member -Name ToolsConnectionObject -Value $(Get-ConnectionString $Script:Settings.ApplicationName) -MemberType NoteProperty
$Script:Settings.Guid
function Get-WrkProcess
{
  switch ($psCmdlet.ParameterSetName)
  {
    'RecycleBlg' {
        $ScriptName = 'BLG Recycle'
        foreach($WrkServer in Get-WrkServers) {
            Get-PerfmonFile
        }
        break
    }
    default {
        $ScriptName = 'AX Report Script'
        Get-AxReport(Get-WrkServers)
        break
    }
  }
}

function Validate-Settings
{
    $Query = "SELECT * FROM AXTools_Environments                
                WHERE ENVIRONMENT = '$Environment'"
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter($Query, $Script:Settings.ToolsConnectionObject)
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null

    if (![string]::IsNullOrEmpty($Table.Tables))
    {
        $Script:Settings | Add-Member -Name EmailProfile -Value $Table.Tables.EmailProfile -MemberType NoteProperty
        $Script:Settings | Add-Member -Name EmailDescription -Value $Table.Tables.Description -MemberType NoteProperty
        $Script:Settings | Add-Member -Name SQLAccount -Value $Table.Tables.DBUser -MemberType NoteProperty
        if($Table.Tables.LocalAdminUser) {
            $Script:Settings | Add-Member -Name LocalAdminAccount -Value $(Get-UserCredentials $Table.Tables.LocalAdminUser) -MemberType NoteProperty
        }
    }
    else {
        Write-Host 'Environment not found.'
        break
    }
}

function Get-WrkServers
{
    Validate-Settings
    $Query = "SELECT [SERVERNAME], [SERVERTYPE] FROM [AXTools_Servers] WHERE [Environment] = '$Environment' AND [ACTIVE] = 1"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:Settings.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $Servers = New-Object System.Data.DataSet
    $Adapter.Fill($Servers) | Out-Null

    $WrkServers = @()
    foreach($Server in $Servers.Tables[0]) {
        if(Test-Connection $Server.ServerName -Count 1 -Quiet) {
            $WrkServers += $Server
        }
        else {
            Write-Log "$($Server.ServerName) - ERROR - Unable to connect to server."
        }
    }

    if($WrkServers.Count -le 0) {
        Write-Log "ERROR - No servers found for $Environment."
        Break
    }
    else {
        return $WrkServers
    }
}

function Get-AxReport
{
param(
    [array]$WrkServers
)
    Write-Log "AX Report Started ($($Script:Settings.ReportDate))"
    foreach($WrkServer in $WrkServers) {
        Write-Log "$($WrkServer.ServerName) ($($WrkServer.ServerType)) Processing."
        $Processes = Get-Process -ComputerName $WrkServer.ServerName | 
            Select Name, Id, Handles, VM, WS, PM, NPM, WorkingSet, PagedMemorySize, PrivateMemorySize, VirtualMemorySize, BasePriority, @{n='ServerName';e={$_.MachineName}}, @{n='Guid';e={$Script:Settings.Guid}}, @{n='ReportDate';e={$Script:Settings.ReportDate}}
        SQL-BulkInsert 'AXReport_RunningProcesses' $Processes
        Switch($WrkServer.ServerType) {
            'AOS' {
                Get-AXConfiguration $WrkServer
                Get-AOSServices $WrkServer
            }
            'SQL' {
                Add-SQLInstance $WrkServer.ServerName '' 'Other Database (Non-AX)'
            }
            'REG' {
                Add-SQLInstance $WrkServer.ServerName '' 'Regional Database (StoreDB)'
            }
            'SRS' {
                $RSObject = Get-WmiObject -class "MSReportServer_ConfigurationSetting" -namespace "root\Microsoft\SqlServer\ReportServer\RS_MSSQLSERVER\v13\Admin" -ComputerName $WrkServer.ServerName
                Add-SQLInstance $RSObject.DatabaseServerName $RSObject.DatabaseName 'SSRS Database'
            }
        }
        Get-EventLogs
        Get-PerfmonLogs
    }
    Get-AXLogs
    Get-SSRSLogs
    AXR-CheckJobs
    AXR-CreateReport
    AXR-CheckJobs
    if(![string]::IsNullOrEmpty($Script:Settings.EmailProfile)) {
        AXR-SendEmail
    }
    if($AutoCleanUp) { Do-CleanUp }
    Write-Log "AX Report Finished ($($Script:Settings.ReportDate))."
}

function Get-EventLogs
{
    $JobStart = Start-Job -Name "AXReport_EventLogs_$($WrkServer.ServerName)_$($WrkServer.ServerType)" -ScriptBlock { & $args[0] $args[1] $args[2] $args[3] $args[4]} -ArgumentList @("$ScriptDir\AX-EventLogs.ps1"), $WrkServer.ServerName, $Script:Settings.Guid, $Script:Settings.ReportDate, $Script:Settings.LocalAdminAccount
}

function AXR-CheckJobs
{
    While ($(Get-Job).Count -gt 0) {
        Get-Job | select id, Name, State | FT -AutoSize
        $JobsDone = Get-Job | Where-Object { $_.State -eq 'Completed' }
        foreach ($Job in $JobsDone) {
            Write-Log "$(($Job.Name).Split('_')[2]) $(($Job.Name).Split('_')[1]) Completed. Duration (min): $([Math]::Round((($Job.PSEndTime - $Job.PSBeginTime).TotalMinutes),2))"
            Remove-Job –Name $($Job.Name)
        }
    Start-Sleep -Milliseconds 1000
    }
}

function Get-AXConfiguration
{
    if($Script:Settings.LocalAdminAccount) {
        $AOSKey = Invoke-command -Computer $($WrkServer.ServerName) -Credential $Script:Settings.LocalAdminAccount { Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services\Dynamics Server' }
        foreach($AOSVersion in $AOSKey) {
            if($AOSVersion.PSChildName.Substring(0,1) -match "^[0-9]*$"){
                    Switch($AOSVersion.PSChildName.Substring(0,1)) {
                    "5" { $Version = "AX2009" }
                    "6" { $Version = "AX2012" }
                    "7" { $Version = "D365" }
                }
                $AOSInstances = Invoke-command -Computer $($WrkServer.ServerName) -Credential $Script:Settings.LocalAdminAccount -ArgumentList $AOSVersion.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") {Get-ChildItem $args[0] }
                foreach($Instance in $AOSInstances) {
                    $Current = Invoke-command -Computer $($WrkServer.ServerName) -Credential $Script:Settings.LocalAdminAccount -ArgumentList $Instance.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).Current }
                    $InstanceName = Invoke-command -Computer $($WrkServer.ServerName) -Credential $Script:Settings.LocalAdminAccount -ArgumentList $Instance.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).InstanceName }
                    $CurrentKey = "$($Instance.Name)\$Current"
                    $DBName = Invoke-command -Computer $($WrkServer.ServerName) -Credential $Script:Settings.LocalAdminAccount -ArgumentList $CurrentKey.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).Database }
                    $DBServer = Invoke-command -Computer $($WrkServer.ServerName) -Credential $Script:Settings.LocalAdminAccount -ArgumentList $CurrentKey.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).DBServer }
                    $Details = "AX Database (Version: $Version / Instance Name: $InstanceName / Configuration: $Current / SQLServer: $DBServer / Database: $DBName)"
                    Write-Log "$($WrkServer.ServerName) - $Details"
                    Add-SQLInstance $DBServer $DBName 'AX Database'
                }
            }
        }
    }
    else {
        $AOSKey = Invoke-command -Computer $($WrkServer.ServerName) { Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services\Dynamics Server' }
        foreach($AOSVersion in $AOSKey) {
            if($AOSVersion.PSChildName.Substring(0,1) -match "^[0-9]*$"){
                    Switch($AOSVersion.PSChildName.Substring(0,1)) {
                    "5" { $Version = "AX2009" }
                    "6" { $Version = "AX2012" }
                    "7" { $Version = "D365" }
                }
                $AOSInstances = Invoke-command -Computer $($WrkServer.ServerName) -ArgumentList $AOSVersion.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") { Get-ChildItem $args[0] }
                foreach($Instance in $AOSInstances) {
                    $Current = Invoke-command -Computer $($WrkServer.ServerName) -ArgumentList $Instance.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).Current }
                    $InstanceName = Invoke-command -Computer $($WrkServer.ServerName) -ArgumentList $Instance.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).InstanceName }
                    $CurrentKey = "$($Instance.Name)\$Current"
                    $DBName = Invoke-command -Computer $($WrkServer.ServerName) -ArgumentList $CurrentKey.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).Database }
                    $DBServer = Invoke-command -Computer $($WrkServer.ServerName) -ArgumentList $CurrentKey.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).DBServer }
                    $Details = "AX Database (Version: $Version / Instance Name: $InstanceName / Configuration: $Current / SQLServer: $DBServer / Database: $DBName)"
                    Write-Log "$($WrkServer.ServerName) - $Details"
                    Add-SQLInstance $DBServer $DBName 'AX Database'
                }
            }
        }
    }

 <#
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$($WrkServer.ServerName))
    $key = "SYSTEM\\CurrentControlSet\\services\\Dynamics Server\\6.0\\01"
    $regkey = $reg.opensubkey($key)
    $keyCurrent = $regKey.GetValue('Current')
    $key = "SYSTEM\\CurrentControlSet\\services\\Dynamics Server\\6.0\\01\\$keyCurrent"
    $regkey = $reg.opensubkey($key)
    $DBServer = $regKey.GetValue('DBServer')
    $DBName = $regKey.GetValue('Database')
#>
}

function Add-SQLInstance($DBServer, $DBName, $Details)
{
    $Query = "SELECT COUNT(DBServer) FROM AXReport_SqlDatabases WHERE DBServer = '$DBServer' and DBNAME = '$DBName' and Guid = '$($Script:Settings.Guid)'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:Settings.ToolsConnectionObject)
    $DBCount = $Cmd.ExecuteScalar()

    if($DBCount -gt 0) {
        $ok = $true
    }
    else {
        $ok = $false
    }
       
    if(!$ok) {
        $SQLInstance = @()
        $Server = Get-SQLObject -DBServer $DBServer -DBName $DBName -SQLAccount $Script:Settings.SQLAccount -ApplicationName $Script:Settings.ApplicationName -SQLServerObject
        $SQLTmp = New-Object -TypeName System.Object
        $SQLTmp | Add-Member -Name Environment -Value $Script:Settings.Environment -MemberType NoteProperty
        $SQLTmp | Add-Member -Name DBServer -Value $DBServer -MemberType NoteProperty
        $SQLTmp | Add-Member -Name DBName -Value $DBName -MemberType NoteProperty
        $SQLTmp | Add-Member -Name Details -Value $Details -MemberType NoteProperty
        $SQLTmp | Add-Member -Name ReportDate -Value $Script:Settings.ReportDate -MemberType NoteProperty
        $SQLTmp | Add-Member -Name Guid -Value $Script:Settings.Guid -MemberType NoteProperty
        $SQLInstance += $SQLTmp
        if($Server.IsClustered) {
            $Cluster = Get-ClusterNode -Cluster $(($DBServer.Split('\'))[0])
            foreach($Node in $Cluster) {
                if($Node.NodeName -match $($Server.Information.Properties | Where-Object { $_.Name -eq 'ComputerNamePhysicalNetBIOS' }).Value) {
                    $SQLTmp = New-Object -TypeName System.Object
                    $SQLTmp | Add-Member -Name Environment -Value $Script:Settings.Environment -MemberType NoteProperty
                    $SQLTmp | Add-Member -Name DBServer -Value ($Node.NodeName).ToUpper() -MemberType NoteProperty
                    $SQLTmp | Add-Member -Name DBName -Value '' -MemberType NoteProperty
                    $SQLTmp | Add-Member -Name Details -Value "Active-Node | $($DBServer) | $($Node.Id) | $($Node.State)" -MemberType NoteProperty
                    $SQLTmp | Add-Member -Name ReportDate -Value $Script:Settings.ReportDate -MemberType NoteProperty
                    $SQLTmp | Add-Member -Name Guid -Value $Script:Settings.Guid -MemberType NoteProperty
                    $SQLInstance += $SQLTmp
                }
                else {
                    $SQLTmp = New-Object -TypeName System.Object
                    $SQLTmp | Add-Member -Name Environment -Value $Script:Settings.Environment -MemberType NoteProperty
                    $SQLTmp | Add-Member -Name DBServer -Value ($Node.NodeName).ToUpper() -MemberType NoteProperty
                    $SQLTmp | Add-Member -Name DBName -Value '' -MemberType NoteProperty
                    $SQLTmp | Add-Member -Name Details -Value "Passive-Node | $($DBServer) | $($Node.Id) | $($Node.State)" -MemberType NoteProperty
                    $SQLTmp | Add-Member -Name ReportDate -Value $Script:Settings.ReportDate -MemberType NoteProperty
                    $SQLTmp | Add-Member -Name Guid -Value $Script:Settings.Guid -MemberType NoteProperty
                    $SQLInstance += $SQLTmp
                }
                if(!($WrkServers.ServerName -like $Node.NodeName)) { 
                    Write-Log "$($WrkServer.ServerName) - $(($Node.NodeName).ToUpper()) ($DBServer) is not set for colletion." 
                } 
            }
        }
    SQL-BulkInsert 'AXReport_SqlDatabases' $SQLInstance
    }
}

function Get-AOSServices
{
    <#
    $AOSService = Get-Service -ComputerName $WrkServer.ServerName |
        Where-Object { $_.DisplayName -like "*Microsoft Dynamics AX*" } | 
            Select  @{n='ServerName';e={$_.MachineName.Trim()}},
                    @{n='Service';e={$_.Name}},
                    @{n='DisplayName';e={$_.DisplayName}},
                    @{n='Status';e={($_.Status).ToString()}},
                    @{n='ReportDate';e={$Script:Settings.ReportDate}},
                    @{n='Guid';e={$Script:Settings.Guid}}
    #>
    $AOSServices = @()
    if($Script:Settings.LocalAdminAccount -and $WrkServer.ServerName -ne $env:COMPUTERNAME) {
        $Services = Get-WmiObject -Class Win32_Service -ComputerName $($WrkServer.ServerName) -Credential $Script:Settings.LocalAdminAccount -ea 0 | Where-Object { $_.DisplayName -like "*Microsoft Dynamics AX*" }
        if($Services) { 
            foreach($Service in $Services) {
                $AOSTemp  = New-Object -TypeName System.Object
                $AOSTemp | Add-Member -Name ServerName -Value $($WrkServer.ServerName) -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Service -Value $Service.Name -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Name -Value $Service.DisplayName -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Status -Value $Service.State -MemberType NoteProperty
                $AOSTemp | Add-Member -Name ReportDate -Value $Script:Settings.ReportDate -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Guid -Value $Script:Settings.Guid -MemberType NoteProperty
                if($Service.ProcessID -ne 0) {
                    $ProcessInfo = Get-WmiObject -Class Win32_Process -ComputerName $($WrkServer.ServerName) -Filter "ProcessID='$($Service.ProcessID)'" -Credential $Script:Settings.LocalAdminAccount -ea 0
                    $AOSTemp | Add-Member -Name StartTime -Value $($Service.ConvertToDateTime($ProcessInfo.CreationDate)) -MemberType NoteProperty
                }
                $AOSServices += $AOSTemp
            }
        }
    }
    else {
        $Services = Get-WmiObject -Class Win32_Service -ComputerName $($WrkServer.ServerName) -ea 0 | Where-Object { $_.DisplayName -like "*Microsoft Dynamics AX*" }
        if($Services) { 
            foreach($Service in $Services) {
                $AOSTemp  = New-Object -TypeName System.Object
                $AOSTemp | Add-Member -Name ServerName -Value $($WrkServer.ServerName) -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Service -Value $Service.Name -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Name -Value $Service.DisplayName -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Status -Value $Service.State -MemberType NoteProperty
                $AOSTemp | Add-Member -Name ReportDate -Value $Script:Settings.ReportDate -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Guid -Value $Script:Settings.Guid -MemberType NoteProperty
                if($Service.ProcessID -ne 0) {
                    $ProcessInfo = Get-WmiObject -Class Win32_Process -ComputerName $($WrkServer.ServerName) -Filter "ProcessID='$($Service.ProcessID)'" -ea 0
                    $AOSTemp | Add-Member -Name StartTime -Value $($Service.ConvertToDateTime($ProcessInfo.CreationDate)) -MemberType NoteProperty
                }
                $AOSServices += $AOSTemp
            }
        }
    }
    SQL-BulkInsert 'AXReport_AxServices' $AOSServices
}

function Get-AXLogs
{
    Write-Log "Quering AX Database (Batch Jobs/Retail/MRP/SQL Errors)"
    $Query = "SELECT DBServer, DBName FROM AXReport_SqlDatabases WHERE Guid = '$($Script:Settings.Guid)' AND DETAILS = 'AX Database' GROUP BY DBServer, DBName"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:Settings.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $SQLInstances = New-Object System.Data.DataSet
    $Adapter.Fill($SQLInstances)

    foreach($SQLInstance in $SQLInstances.Tables[0]) {
        $Conn = Get-SQLObject -DBServer $($SQLInstance.DBServer) -DBName 'tempdb' -SQLAccount $Script:Settings.SQLAccount -ApplicationName $Script:Settings.ApplicationName
        $Query = Get-Content $ModuleFolder\ConDrop.sql | Out-String
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Cmd.ExecuteScalar() | Out-Null        
        $Query = Get-Content $ModuleFolder\ConPeek.sql | Out-String
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Cmd.ExecuteScalar() | Out-Null
        $Query = Get-Content $ModuleFolder\ConSize.sql | Out-String
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Cmd.ExecuteScalar() | Out-Null
        $Conn.Close()

        #Batch Jobs
        $Conn = Get-SQLObject -DBServer $($SQLInstance.DBServer) -DBName $($SQLInstance.DBName) -SQLAccount $Script:Settings.SQLAccount -ApplicationName $Script:Settings.ApplicationName
        $Query = "SELECT A.CAPTION AS HISTORYCAPTION
		                ,B.CAPTION AS JOBCAPTION
		                ,STATUS = CASE A.STATUS 
			                WHEN 3 THEN 'Error'
			                WHEN 6 THEN 'Didnt run'
			                WHEN 7 THEN 'Canceling'
			                WHEN 8 THEN 'Canceled'
			                END
		                ,A.SERVERID
		                ,DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.STARTDATETIME) as STARTDATETIMECST
		                ,DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.ENDDATETIME) as ENDDATETIMECST
		                ,A.EXECUTEDBY 
		                ,([TempDB].[dbo].[CONPEEK](CAST([TempDB].[dbo].[CONPEEK](A.INFO, 2) AS varbinary(8000)), 2)) AS LOG
		                ,A.BATCHID
		                ,A.BATCHJOBID
		                ,A.BATCHJOBHISTORYID
	                FROM BATCHHISTORY A WITH(NOLOCK)
		                FULL OUTER JOIN BATCHJOB B WITH(NOLOCK) 
			                ON A.BATCHJOBID = B.RECID
                WHERE A.STATUS IN (3, 6, 7 , 8)
                AND (DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.STARTDATETIME) >= '$((Get-Date).AddDays(-1).Date)' 
                AND DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.STARTDATETIME) < '$((Get-Date).AddDays(0).Date)')
                ORDER BY STARTDATETIMECST"
        
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $Cmd
        $AXJobsDS = New-Object System.Data.DataSet
        $AXJobsCnt = $Adapter.Fill($AXJobsDS)
        if($AXJobsCnt -gt 0) {
            $AXBatch = $AXJobsDS.Tables[0] | 
                Select HISTORYCAPTION, JOBCAPTION, STATUS, @{n='SERVERID';e={($_.SERVERID -replace '01@','').Trim()}}, STARTDATETIMECST, ENDDATETIMECST, EXECUTEDBY, BATCHID,	BATCHJOBID, BATCHJOBHISTORYID, @{n='LOG';e={($_.LOG -replace '\t|\r|\n', " ").Trim()}}, @{n='Guid';e={$Script:Settings.Guid}}, @{n='ReportDate';e={$Script:Settings.ReportDate}}
            SQL-BulkInsert 'AXReport_AxBatchJobs' $AXBatch
        }

        #Long Running Jobs
        $Query = "SELECT B.CAPTION AS JOB
		            , COUNT(1) AS 'COUNT'
		            , STATUS = CASE A.STATUS 
			            WHEN 0 THEN 'Hold'
			            WHEN 1 THEN 'Waiting'
			            WHEN 2 THEN 'Executing'
			            WHEN 3 THEN 'Error'
			            WHEN 4 THEN 'Finished'
			            WHEN 5 THEN 'Ready'
			            WHEN 6 THEN 'Didnt run'
			            WHEN 7 THEN 'Canceling'
			            WHEN 8 THEN 'Canceled'
			            END
		            , A.SERVERID
		            , DATEDIFF(MINUTE, MAX(A.STARTDATETIME),MAX(A.ENDDATETIME)) AS 'DURATION'
		            , A.EXECUTEDBY
	            FROM BATCHHISTORY A WITH(NOLOCK)
		            FULL OUTER JOIN BATCHJOB B WITH(NOLOCK) 
			            ON A.BATCHJOBID = B.RECID
            WHERE (DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.STARTDATETIME) >= '$((Get-Date).AddDays(-1).Date)' AND 
            DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.STARTDATETIME) < '$((Get-Date).AddDays(0).Date)')
            GROUP BY B.CAPTION, A.STATUS, A.SERVERID, A.EXECUTEDBY 
            HAVING DATEDIFF(MINUTE, MAX(A.STARTDATETIME),MAX(A.ENDDATETIME)) > 15
            ORDER BY 5 DESC"
        
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $Cmd
        $AXLongJobsDS = New-Object System.Data.DataSet
        $AXLongJobCnt = $Adapter.Fill($AXLongJobsDS)
        if($AXLongJobCnt -gt 0) {
            $AXLongBatch = $AXLongJobsDS.Tables[0] | 
                Select JOB, COUNT, STATUS, DURATION, EXECUTEDBY, @{n='SERVERID';e={($_.SERVERID -replace '01@','').Trim()}}, @{n='Guid';e={$Script:Settings.Guid}}, @{n='ReportDate';e={$Script:Settings.ReportDate}}
            SQL-BulkInsert 'AXReport_AxLongBatchJobs' $AXLongBatch
        }

        #CDX Jobs
        $Query = "SELECT B.JOBID  
				, A.STATUS AS DATASTORESTATUS
		        , STATUSDOWNLOADSESSIONDATASTORE = CASE A.STATUS 
		                WHEN 0 THEN 'Started'
						WHEN 1 THEN 'Available'
						WHEN 2 THEN 'Requested'
						WHEN 3 THEN 'Downloaded'
						WHEN 4 THEN 'Applied'
						WHEN 5 THEN 'Canceled'
		                WHEN 6 THEN 'Create Failed'
		                WHEN 7 THEN 'Download Failed'
					    WHEN 8 THEN 'Apply Failed'
						WHEN 9 THEN 'No Data'
	                    END
				, A.MESSAGE
				, DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.DATEREQUESTED) AS DATEREQUESTED
				, DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.DATEDOWNLOADED) AS DATEDOWNLOADED
				, DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.DATEAPPLIED) AS DATEAPPLIED
		        , B.CURRENTROWVERSION
		        , B.DATAFILEOUTPUTPATH
		        , B.ROWSAFFECTED
		        , B.STATUS AS SESSIONSTATUS
				, STATUSDOWNLOADSESSION = CASE B.STATUS 
		                WHEN 0 THEN 'Started'
						WHEN 1 THEN 'Available'
						WHEN 2 THEN 'Requested'
						WHEN 3 THEN 'Downloaded'
						WHEN 4 THEN 'Applied'
						WHEN 5 THEN 'Canceled'
		                WHEN 6 THEN 'Create Failed'
		                WHEN 7 THEN 'Download Failed'
					    WHEN 8 THEN 'Apply Failed'
						WHEN 9 THEN 'No Data'
	                    END
		        , C.DATABASE_
		        , C.NAME
		        , DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.MODIFIEDDATETIME) AS MODIFIEDDATETIME
            FROM RETAILCDXDOWNLOADSESSIONDATASTORE A WITH(NOLOCK)
            JOIN RETAILCDXDOWNLOADSESSION B WITH(NOLOCK)
            ON A.SESSION_ = B.RECID
            JOIN RETAILCONNDATABASEPROFILE C WITH(NOLOCK)
            ON A.DATASTORE = C.RECID
            WHERE (A.STATUS IN ('1','5','6','7','8') OR B.STATUS IN ('5','6','7','8')) 
                AND (DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.MODIFIEDDATETIME) >= '$((Get-Date).AddDays(-1).Date)'
                AND (DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), A.MODIFIEDDATETIME) < '$((Get-Date).AddDays(0).Date)'))"

        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $Cmd
        $RetailDS = New-Object System.Data.DataSet
        $AXRetailCnt = $Adapter.Fill($RetailDS)
        if($AXRetailCnt -gt 0) {
            $AXRetail =  $RetailDS.Tables[0] | 
                Select JOBID, DATASTORESTATUS, STATUSDOWNLOADSESSIONDATASTORE, @{n='MESSAGE';e={($_.MESSAGE -replace '\t|\r|\n', " ").Trim()}}, DATEREQUESTED, DATEDOWNLOADED, DATEAPPLIED, CURRENTROWVERSION, ROWSAFFECTED, DATAFILEOUTPUTPATH, SESSIONSTATUS, STATUSDOWNLOADSESSION, DATABASE_, NAME, MODIFIEDDATETIME, @{n='Guid';e={$Script:Settings.Guid}}, @{n='ReportDate';e={$Script:Settings.ReportDate}}
            SQL-BulkInsert 'AXReport_AxRetailJobs' $AXRetail
        }
        ##MRP
        $Query = "SELECT 
	                 REQPLANID
	                ,DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), STARTDATETIME) AS STARTDATETIME
	                ,DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), ENDDATETIME) AS ENDDATETIME
	                ,CANCELLED
	                ,USEDCHILDTHREADS
	                ,MAXCHILDTHREADS
	                ,COMPLETEUPDATE
	                ,USEDTODAYSDATE
	                ,NUMOFITEMS
	                ,NUMOFINVENTONHAND
	                ,NUMOFSALESLINE
	                ,NUMOFPURCHLINE
	                ,NUMOFTRANSFERPLANNEDORDER
	                ,NUMOFITEMPLANNEDORDER
	                ,NUMOFINVENTJOURNAL
	                ,TIMECOPY
	                ,TIMECOVERAGE
	                ,TIMEUPDATE
	                ,([TempDB].[dbo].[CONPEEK](CAST([TempDB].[dbo].[CONPEEK](LOG, 2) AS varbinary(8000)), 2)) AS LOG
                FROM REQLOG WITH(NOLOCK)
                WHERE DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), STARTDATETIME) >= '$((Get-Date).AddDays(-1).Date)'
                      AND DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), STARTDATETIME) < '$((Get-Date).AddDays(0).Date)'"

        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $Cmd
        $MRPDS = New-Object System.Data.DataSet
        $MRPCnt = $Adapter.Fill($MRPDS)
        if($MRPCnt -gt 0) {
            $AXMRP = $MRPDS.Tables[0] | 
                Select REQPLANID, STARTDATETIME, ENDDATETIME, CANCELLED, USEDCHILDTHREADS, MAXCHILDTHREADS, COMPLETEUPDATE, USEDTODAYSDATE, NUMOFITEMS, NUMOFINVENTONHAND, NUMOFSALESLINE, NUMOFPURCHLINE, NUMOFTRANSFERPLANNEDORDER, NUMOFITEMPLANNEDORDER, NUMOFINVENTJOURNAL, TIMECOPY, TIMECOVERAGE, TIMEUPDATE, @{n='LOG';e={($_.LOG -replace '\t|\r|\n', " ").Trim()}}, @{n='Guid';e={$Script:Settings.Guid}}, @{n='ReportDate';e={$Script:Settings.ReportDate}}
            SQL-BulkInsert 'AXReport_AxMRP' $AXMRP
        }

        ##SQL Error Logs
        $SQLConn = Get-SQLObject -DBServer $SQLInstance.DBServer -DBName 'master' -SQLAccount $Script:Settings.SQLAccount -ApplicationName $Script:Settings.ApplicationName -SQLServerObject
        $SQLLogs = $SQLConn.ReadErrorLog() | Where-Object { ($_.LogDate -ge $((Get-Date).AddDays(-1).Date)) -and ($_.LogDate -lt $((Get-Date).AddDays(0).Date)) } |
                Select LogDate, ProcessInfo,  @{n='Text';e={($_.Text -replace '\t|\r|\n', " ").Trim()}}, @{n='Server';e={$SQLInstance.DBServer}}, @{n='Database';e={$SQLInstance.DBName}}, @{n='Guid';e={$Script:Settings.Guid}}, @{n='ReportDate';e={$Script:Settings.ReportDate}} #| Where-Object {($_.LogDate -ge $((Get-Date).AddDays(-1).Date)) }
        SQL-BulkInsert 'AXReport_SQLLog' $SQLLogs

    $Conn.Close()
    }
}

function Get-PerfmonFile
{
    try {
        $DataCollectorSet = New-Object -COM Pla.DataCollectorSet
        $DataCollectorSet.Query($($Script:Settings.DataCollectorName),$WrkServer.ServerName)
    }
    catch {
        Write-Log "ERROR - $($Script:Settings.DataCollectorName) Failed. ($($_.Exception.Message))."
        $CIMComputer = New-CIMSession -Computername $WrkServer.ServerName
        Enable-NetFirewallRule -DisplayGroup "Performance Logs and Alerts" -CimSession $CIMComputer
        Enable-NetFirewallRule -DisplayGroup "Windows Management Instrumentation (WMI)" -CimSession $CIMComputer
        Remove-CIMSession -ComputerName $($Script:Settings.DataCollectorName)
        #
        $Query = "SELECT TOP 1 [TEMPLATEXML] FROM [AXTools_PerfmonTemplates] WHERE [SERVERTYPE] = '$($WrkServer.ServerType)' and [ACTIVE] = 1 ORDER BY CREATEDDATETIME DESC"
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:Settings.ToolsConnectionObject)
        $Xml = $Cmd.ExecuteScalar()
        $DataCollectorSet.SetXml($Xml)
        $DataCollectorSet.RootPath = "%systemdrive%\PerfLogs\Admin\$($Script:Settings.DataCollectorName)"
        $DataCollectorSet.Commit($Script:Settings.DataCollectorName,$WrkServer.ServerName,0x0003) | Out-Null
    }

    if($DataCollectorSet.Status -eq 1) {
        $DataCollectorSet.Stop($false)
        Start-Sleep -Seconds 2
        $DataCollectorSet.Start($false)
    }
    else {
        try {
            $DataCollectorSet.Start($false)
            Write-Log "ERROR - $($WrkServer.ServerName) - $($Script:Settings.DataCollectorName) stopped, attempt start it."
        }
        catch {
            Write-Log "ERROR - $($WrkServer.ServerName) - $($_.Exception.Message)"
        }
    }
    if($AutoCleanUp) {
        [Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
        $Path = "\\$($WrkServer.ServerName)\C$\PerfLogs\Admin\$($Script:Settings.DataCollectorName)"
        $BlgFiles = Get-ChildItem -Path $Path | Where {$_.Extension -match '.blg' -and $_.LastWriteTime -lt $((Get-Date).AddDays(-$Script:Configuration.Settings.General.RetentionDays))}  | Sort-Object -Property LastWriteTime
        if($BlgFiles.Count -ge 5) {
            if(!(Test-Path("$Path\Temp\"))) {
                New-Item -ItemType Directory -Force -Path "$Path\Temp" | Out-Null
            }
            Move-Item $BlgFiles.FullName -Destination "$Path\Temp\" #-Force
            $FileServer = ($BlgFiles.Name | Select -First 1).Split("_")[0]
            $FileSTLog = (($BlgFiles.Name | Select -First 1).Split(" ")).Split(".")[2]
            $FileLTLog = (($BlgFiles.Name | Select -Last 1).Split(" ")).Split(".")[2]
            ## Create zip file
            [System.IO.Compression.ZipFile]::CreateFromDirectory("$Path\Temp\","$Path\$FileServer`_$FileSTLog-$FileLTLog.zip",$CompressionLevel,$false)
            ## Delete Temp Folder
            Remove-Item -Path "$Path\Temp\" -Recurse -Force
            $ZipFiles = Get-ChildItem -Path $Path | Where {$_.Extension -match '.zip'}  | Sort-Object -Property LastWriteTime
            $DestPath = (Join-Path "\\$($env:COMPUTERNAME)" $LogFolder).Replace(':','$')
            if($ZipFiles) {
                if(!(Test-Path("$DestPath\$($WrkServer.ServerName)\"))) {
                    New-Item -ItemType Directory -Force -Path "$DestPath\$($WrkServer.ServerName)" | Out-Null
                }
                Move-Item $ZipFiles.FullName -Destination "$DestPath\$($WrkServer.ServerName)"
            }
        }
    }
}

function Get-PerfmonLogs
{
    if(Test-Path "\\$($WrkServer.ServerName)\C$\PerfLogs\Admin\$($Script:Settings.DataCollectorName)\") {
        $BlgFile = Get-ChildItem -Path "\\$($WrkServer.ServerName)\C$\PerfLogs\Admin\$($Script:Settings.DataCollectorName)\" | 
            #Where-Object { $_.Extension -match '.blg' -and $_.CreationTime -lt $((Get-Date).AddDays(0).Date) } |
            Where-Object { $_.Extension -match '.blg' -and $_.CreationTime -ge $((Get-Date).AddDays(-1).Date) -and $_.CreationTime -lt $((Get-Date).AddDays(0).Date) -and $(New-TimeSpan ($_.CreationTime) ($_.LastWriteTime)).TotalMinutes -gt 5 } |
            #Where-Object { $_.Extension -match '.blg' } | 
                Sort-Object -Property CreationTime -Descending
        if($BlgFile) {
            $Paths = Import-Counter -Path $BlgFile.FullName -ListSet * -ErrorAction SilentlyContinue | % { $_.PathsWithInstances }
            $Paths += Import-Counter -Path $BlgFile.FullName -ListSet * -ErrorAction SilentlyContinue | % { $_.Counter }
            $Script:BlgCounters = @()
            foreach($Path in $Paths) {
                switch -wildcard ($Path) {
                    '*Processor(_Total)\% Processor Time*' { Add-PerfCounter $Path 'SRV' 1 }
                    '*Available MBytes*' { Add-PerfCounter $Path 'SRV' 1 }
                    '*Paging File(_Total)\% Usage' { Add-PerfCounter $Path 'SRV' 1 }
                    #
                    '*Microsoft Dynamics AX Object Server(01)*'  { Add-PerfCounter $Path 'AX' 1 }
                    '*Process(ax32serv*)\ID Process*' { Add-PerfCounter $Path 'AX' 0 }
                    '*Process(ax32serv*)\Virtual Bytes*' { Add-PerfCounter $Path 'AX' 0 }
                    '*Process(ax32serv*)\Private Bytes*' { Add-PerfCounter $Path 'AX' 0 }
                    '*Process(ax32serv*)\% Processor Time*' { Add-PerfCounter $Path 'AX' 0 }
                    '*Process(ax32serv*)\Working Set*' { Add-PerfCounter $Path 'AX' 0 }
                    #
                    '*Process(sqlservr*)\ID Process*' { Add-PerfCounter $Path 'SQL' 0 }
                    '*Process(sqlservr*)\Virtual Bytes*' { Add-PerfCounter $Path 'SQL' 0 }
                    '*Process(sqlservr*)\Private Bytes*' { Add-PerfCounter $Path 'SQL' 0 }
                    '*Process(sqlservr*)\% Processor Time*' { Add-PerfCounter $Path 'SQL' 0 }
                    '*Process(sqlservr*)\Working Set*' { Add-PerfCounter $Path 'SQL' 0 }
                    '*SQL Statistics\SQL Re-Compilations/sec*' { Add-PerfCounter $Path 'SQL' 1 }
                    '*SQL Statistics\SQL Compilations/sec*' { Add-PerfCounter $Path 'SQL' 1 }
                    '*Memory Manager\Total Server Memory*' { Add-PerfCounter $Path 'SQL' 1 }
                    '*Memory Manager\Target Server Memory*' { Add-PerfCounter $Path 'SQL' 1 }
                    '*Locks(_Total)\Number of Deadlocks/sec*' { Add-PerfCounter $Path 'SQL' 1 }
                    '*Locks(_Total)\Lock Wait Time (ms)*' { Add-PerfCounter $Path 'SQL' 1 }
                    '*Buffer Manager\Buffer cache hit ratio*' { Add-PerfCounter $Path 'SQL' 1 }
                    '*Buffer Manager\Page life expectancy*' { Add-PerfCounter $Path 'SQL' 1 }
                    '*Buffer Manager\Page life expectancy*' { Add-PerfCounter $Path 'SQL' 1 }
                    '*Latches\Latch Waits*' { Add-PerfCounter $Path 'SQL' 1 }
                    '*Latches\Total Latch Wait Time*' { Add-PerfCounter $Path 'SQL' 1 }
                    #
                    '*Web Service(*RealTimeService)\Current Connections' { Add-PerfCounter $Path 'RTS' 1 }
                    '*Web Service(*RealTimeService)\Bytes Received/sec' { Add-PerfCounter $Path 'RTS' 1 }
                    '*Web Service(*RealTimeService)\Bytes Sent/sec' { Add-PerfCounter $Path 'RTS' 1 }
                    '*WAS_W3WP(*RealTimeService)\Health Ping Reply Latency' { Add-PerfCounter $Path 'RTS' 1 }
                    '*WAS_W3WP(*RealTimeService)\Total Health Pings.' { Add-PerfCounter $Path 'RTS' 1 }
                    '*W3SVC_W3WP(*RealTimeService)\Requests / Sec' { Add-PerfCounter $Path 'RTS' 1 }
                    '*W3SVC_W3WP(*RealTimeService)\Active Requests' { Add-PerfCounter $Path 'RTS' 1 }
                    #
                    '*Web Service(*AsyncService)\Current Connections' { Add-PerfCounter $Path 'SYNC' 1 }
                    '*Web Service(*AsyncService)\Bytes Received/sec' { Add-PerfCounter $Path 'SYNC' 1 }
                    '*Web Service(*AsyncService)\Bytes Sent/sec' { Add-PerfCounter $Path 'SYNC' 1 }
                    '*WAS_W3WP(*AsyncService)\Health Ping Reply Latency' { Add-PerfCounter $Path 'SYNC' 1 }
                    '*WAS_W3WP(*AsyncService)\Total Health Pings.' { Add-PerfCounter $Path 'SYNC' 1 }
                    '*W3SVC_W3WP(*AsyncService)\Requests / Sec' { Add-PerfCounter $Path 'SYNC' 1 }
                    '*W3SVC_W3WP(*AsyncService)\Active Requests' { Add-PerfCounter $Path 'SYNC' 1 }
                    #
                    '*Web Service(*Default*)\Current Connections' { Add-PerfCounter $Path 'STO' 1 }
                    '*Web Service(*Default*)\Bytes Received/sec' { Add-PerfCounter $Path 'STO' 1 }
                    '*Web Service(*Default*)\Bytes Sent/sec' { Add-PerfCounter $Path 'STO' 1 }
                    '*WAS_W3WP(*Default*)\Health Ping Reply Latency' { Add-PerfCounter $Path 'STO' 0 }
                    '*WAS_W3WP(*Default*)\Total Health Pings.' { Add-PerfCounter $Path 'STO' 0 }
                    '*W3SVC_W3WP(*Default*)\Requests / Sec' { Add-PerfCounter $Path 'STO' 0 }
                    '*W3SVC_W3WP(*Default*)\Active Requests' { Add-PerfCounter $Path 'STO' 0 }
                    #
                    '*Terminal Services\Inactive Sessions' { Add-PerfCounter $Path 'RDP' 1 }
                    '*Terminal Services\Total Sessions' { Add-PerfCounter $Path 'RDP' 1 }
                    '*Terminal Services\Active Sessions' { Add-PerfCounter $Path 'RDP' 1 }
                    #
                    '*Process(ReportingServicesService*)\ID Process*' { Add-PerfCounter $Path 'SRS' 0 }
                    '*Process(ReportingServicesService*)\Virtual Bytes*' { Add-PerfCounter $Path 'SRS' 0 }
                    '*Process(ReportingServicesService*)\Private Bytes*' { Add-PerfCounter $Path 'SRS' 0 }
                    '*Process(ReportingServicesService*)\% Processor Time*' { Add-PerfCounter $Path 'SRS' 0 }
                    '*Process(ReportingServicesService*)\Working Set*' { Add-PerfCounter $Path 'SRS' 0 }
                    '*ReportServer:Service\Active Connections*' { Add-PerfCounter $Path 'SRS' 1 }
                    '*ReportServer:Service\Memory Pressure State*' { Add-PerfCounter $Path 'SRS' 1 }
                    '*ReportServer:Service\Memory Shrink Amount*' { Add-PerfCounter $Path 'SRS' 0 }
                    '*ReportServer:Service\Memory Shrink Notifications/sec*' { Add-PerfCounter $Path 'SRS' 0 }
                    '*ReportServer:Service\Tasks Queued*' { Add-PerfCounter $Path 'SRS' 0 }
                    '*ReportServer:Service\Errors Total*' { Add-PerfCounter $Path 'SRS' 1 }
                    '*ReportServer:Service\Errors/sec*' { Add-PerfCounter $Path 'SRS' 1 }
		            '*ReportServer:Service\Requests Disconnected*' { Add-PerfCounter $Path 'SRS' 1 }
		            '*ReportServer:Service\Requests Executing*' { Add-PerfCounter $Path 'SRS' 1 }
		            '*ReportServer:Service\Requests Not Authorized*' { Add-PerfCounter $Path 'SRS' 1 }
		            '*ReportServer:Service\Requests Rejected*' { Add-PerfCounter $Path 'SRS' 1 }
		            '*ReportServer:Service\Requests Total*' { Add-PerfCounter $Path 'SRS' 1 }
		            '*ReportServer:Service\Requests/sec*' { Add-PerfCounter $Path 'SRS' 1 }
                }
            }
            SQL-BulkInsert 'AXReport_PerfmonData' $Script:BlgCounters
        }
    }
}

function Add-PerfCounter($Path, $Type, $ReportView)
{
    $CounterData = Import-Counter -Path $BlgFile.FullName -Counter $Path -ErrorAction SilentlyContinue
    $CounterSummary = Import-Counter -Path $BlgFile.FullName -Summary

    Switch -wildcard ($Path) {
        '*MSSQL$*'{
            $NewPath = (($Path.Substring($WrkServer.ServerName.Length + 9)).Split(':'))[0] + '\' +(($Path.Substring($WrkServer.ServerName.Length + 9)).Split('\'))[1]
            if($NewPath.Contains('Server Memory (KB)')) {
                $NewPath = $NewPath.Replace('KB','GB')
            }
        }
        '*SQLServer:*' {
            $NewPath = 'SQL\' + (($Path.Substring($WrkServer.ServerName.Length + 3)).Split('\'))[1]
            if($NewPath.Contains('Server Memory (KB)')) {
                $NewPath = $NewPath.Replace('KB','GB')
            }
        }
        '*RealTimeService*' {
            $NewPath = 'RealTimeService\' + (($Path.Substring($WrkServer.ServerName.Length + 3)).Split('\'))[1]
        }
        '*AsyncService*' {
            $NewPath = 'AsyncService\' + (($Path.Substring($WrkServer.ServerName.Length + 3)).Split('\'))[1]
        }
        '*DefaultAppPool*' {
            $NewPath = 'DefaultAppPool\' + (($Path.Substring($WrkServer.ServerName.Length + 3)).Split('\'))[1]
        }
        '*Private Bytes*' {
            $NewPath = $Path.Substring($WrkServer.ServerName.Length + 3) + ' (GB)'
        }
        '*Virtual Bytes*' {
            $NewPath = $Path.Substring($WrkServer.ServerName.Length + 3) + ' (GB)'
        }
        '*Working Set*' {
            $NewPath = $Path.Substring($WrkServer.ServerName.Length + 3) + ' (GB)'
        }
        '*Available MBytes*' {
            $NewPath = 'Available GBytes'
        }
        '*Paging File(_Total)*' {
            $NewPath = 'Paging File %'
        }
        '*Processor(_Total)*' {
            $NewPath = 'CPU Time %'
        }
        '*Microsoft Dynamics AX Object Server(01)\ACTIVE SESSIONS*' {
            $NewPath = 'AX Sessions'
        }
        Default {
            $NewPath = $Path.Substring($WrkServer.ServerName.Length + 3)
        }
    }

    $tmpCounter = New-Object -TypeName System.Object
    $tmpCounter | Add-Member -Name ServerName -Value $WrkServer.ServerName -MemberType NoteProperty
    $tmpCounter | Add-Member -Name ServerType -Value $WrkServer.ServerType -MemberType NoteProperty
    $tmpCounter | Add-Member -Name CounterType -Value $Type -MemberType NoteProperty
    $tmpCounter | Add-Member -Name ReportView -Value $ReportView -MemberType NoteProperty
    $tmpCounter | Add-Member -Name Path -Value $Path.Substring($WrkServer.ServerName.Length + 3) -MemberType NoteProperty
    $tmpCounter | Add-Member -Name Maximum -Value $(($CounterData.CounterSamples | Measure-Object CookedValue -Ave -Max -Min).Maximum) -MemberType NoteProperty 
    $tmpCounter | Add-Member -Name Minimum -Value $(($CounterData.CounterSamples | Measure-Object CookedValue -Ave -Max -Min).Minimum) -MemberType NoteProperty
    $tmpCounter | Add-Member -Name Average -Value $(($CounterData.CounterSamples | Measure-Object CookedValue -Ave -Max -Min).Average) -MemberType NoteProperty
    $tmpCounter | Add-Member -Name FullPath -Value $Path -MemberType NoteProperty
    $tmpCounter | Add-Member -Name StartDateTime -Value $CounterSummary.OldestRecord -MemberType NoteProperty
    $tmpCounter | Add-Member -Name EndDateTime -Value $CounterSummary.NewestRecord -MemberType NoteProperty
    $tmpCounter | Add-Member -Name Samples -Value $CounterSummary.SampleCount -MemberType NoteProperty
    $tmpCounter | Add-Member -Name Counter -Value $NewPath -MemberType NoteProperty
    $tmpCounter | Add-Member -Name Guid -Value $Script:Settings.Guid -MemberType NoteProperty
    $tmpCounter | Add-Member -Name ReportDate -Value $Script:Settings.ReportDate -MemberType NoteProperty
    
    $Script:BlgCounters += $tmpCounter
}

function Get-SSRSLogs
{
    $Query = "SELECT DBServer, DBName FROM AXReport_SqlDatabases WHERE Guid = '$($Script:Settings.Guid)' AND DETAILS = 'SSRS Database' GROUP BY DBServer, DBName"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Script:Settings.ToolsConnectionObject)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $SRSInstances = New-Object System.Data.DataSet
    $Adapter.Fill($SRSInstances)

    Write-Log "Working on SSRS Logs."
    foreach($SRSInstance in $SRSInstances.Tables[0]) {
        $Conn = Get-SQLObject -DBServer $($SRSInstance.DBServer) -DBName $($SRSInstance.DBName) -SQLAccount $Script:Settings.SQLAccount -ApplicationName $Script:Settings.ApplicationName
        $Query = "SELECT Status
                    , InstanceName
		            , ReportPath
		            , UserName
		            , Format
		            , TimeStart
		            , TimeEnd
		            , TimeDataRetrieval
		            , TimeProcessing
		            , TimeRendering 
                FROM $($SRSInstance.DBName)..ExecutionLog2 
                WHERE Status <> 'rsSuccess' 
		            AND TimeStart >= '$((Get-Date).AddDays(-1).Date)' 
                    AND TimeStart < '$((Get-Date).AddDays(0).Date)'"

        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $Cmd
        $SSRSDS = New-Object System.Data.DataSet
        $SRSLogCnt = $Adapter.Fill($SSRSDS)
        
        if($SRSLogCnt -gt 0) {
            $SRSLogs = $SSRSDS.Tables[0] | 
                Select Status, InstanceName, ReportPath, UserName, Format, TimeStart, TimeEnd, TimeDataRetrieval, TimeProcessing, TimeRendering, @{n='Guid';e={$Script:Settings.Guid}}, @{n='ReportDate';e={$Script:Settings.ReportDate}}
            SQL-BulkInsert 'AXReport_SRSLog' $SRSLogs
        }
    $Conn.Close()
    }
}

function AXR-CreateReport
{
    Write-Log "HTML Started ($FileDateTime)."
    $JobStart = Start-Job -Name "AXReport_CreateReport" -ScriptBlock {& $args[0] $args[1] $args[2] } -ArgumentList @("$ScriptDir\AX-CreateReport.ps1"), $Script:Settings.Guid, $Script:Settings.Environment
}

function AXR-SendEmail
{
    $Subject = "AX Daily Report <$((Get-Date).AddDays(-1) | Get-Date -Format "MMM dd, yyyy")>"
    $Body = Get-Content $ReportFolder\AXReport-$(Get-Date ($Script:Settings.ReportDate) -f MMddyyyy)-Summary.html
    $Attachment = "$ReportFolder\AXReport-$(Get-Date ($Script:Settings.ReportDate) -f MMddyyyy).mht"
    Send-Email -Subject $Subject -Body $Body -Attachment $Attachment -EmailProfile $Script:Settings.EmailProfile
    Write-Log "AX Report has been Sent."
}

function Do-Cleanup
{
    $Files = Get-ChildItem -Path $ReportFolder | Where { $_.LastWriteTime -lt $((Get-Date).AddDays((-$Script:Configuration.Settings.General.RetentionDays))) }
    if($Files) {
        Remove-Item -Path $Files.FullName -Force
    }
    Write-Log("Cleanup $($Files.Count) files, done.")
}

function Check-Folder {
param(
    [string]$Path
)
    if(!(Test-Path($Path))) {
        New-Item -ItemType Directory -Force -Path $Path | Out-Null
    }
}

Check-Folder $ReportFolder
Check-Folder $LogFolder

Get-WrkProcess
Get-Module | Where-Object {$_.ModuleType -eq 'Script'} | % { Remove-Module $_.Name }