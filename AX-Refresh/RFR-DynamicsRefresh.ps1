# .DISCLAIMER
#    Microsoft Corporation. All rights reserved.
#    Do not use in Production. Sample scripts in this guide are not supported under any Microsoft standard support program or service.
#    These sample scripts are provided 'as is' without warranty of any kind expressed or implied. Microsoft disclaims all implied warranties
#    including, without limitation, any implied warranties of merchantability or fitness for a particular purpose. The entire risk 
#    arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, 
#    its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
#    (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or 
#    other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has 
#    been advised of the possibility of such damages. The opinions and views expressed in this script are those of the author and do 
#    not necessarily state or reflect those of Microsoft. Do this at your own risk. The author will not be held responsible for any 
#    damage you incur when running, modifying or carrying out these scripts.
#
# .DESCRIPTION
#    Creates customer's user acceptance testing configuration for selected tables and restore the same data after refreshing transactional data.
#    Restores previous configuration in Training or Development environment.
#
# .NOTES
#    File Name      : RFR-DynamicsRefresh.ps1
#    Author         : Bruno Ferreti
#    Prerequisite   : PowerShell for SQL Server Modules (SQLPS)
#    Copyright 2016
#

[CmdletBinding()]
param(
	[Parameter(Mandatory = $false,ValueFromPipeline = $true)]
	[string]$EnvironName,
	[Parameter(Mandatory = $false,ValueFromPipeline = $true)]
	[int]$RefreshDays,
	[Parameter(Mandatory = $false,ValueFromPipeline = $true)]
	[switch]$RestoreDB,
	[Parameter(Mandatory = $false,ValueFromPipeline = $true)]
	[switch]$RefreshOnly
)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"
$WorkFolder = $Dir + "\WorkFolder"

Import-Module $ModuleFolder\AX-Tools.psm1 -DisableNameChecking

$Global:Guid = ([guid]::NewGuid()).GUID
$Script:Settings = Import-ConfigFile -ScriptName 'AxRefresh'
$Script:Settings | Add-Member -Name Guid -Value $Global:Guid -MemberType NoteProperty

Clear-Host

function Get-Menu
{
	switch ($MenuName) {
		'Main' { Get-MainMenu }
		'Miscellaneous' { Get-MiscMenu }
		'Batch' { Get-BatchMenu }
		'Services' { Get-ServicesMenu }
		Default { Get-MainMenu }
	}
}

function Get-MainMenu
{
	$MenuName = 'Main'
	Write-Host ''
	Write-Host 'Select one of the following options:'
	Write-Host ''
	Write-Host '0. Run entire refresh script'
	Write-Host ''
	Write-Host '1. Export Environment Configuration'
	Write-Host '2. Stop AOS Services'
	Write-Host '3. Restore DB Backup (*.bak)'
	Write-Host '4. Delete Live Data ' -NoNewline
	if ($Script:Environment.Name) { "from ($($Script:Environment.Name))" } else { "" }
	Write-Host '5. Restore Environment Configuration'
	Write-Host '6. Start AOS Services'
	Write-Host ''
	Write-Host '9. Additional AX Tools'
	Write-Host ''
	if ($Script:Environment.MachineName) { Get-MenuConfig }
	Write-Host 'L - Load Environment'
	Write-Host 'Q - Quit'
	if ($Script:WarningMsg) {
		Write-Host ''
		Write-Warning $Script:WarningMsg
		Clear-Variable WarningMsg -Scope Script }
	Write-Host '════════════════════════════════════════════════════════════'
	$Prompt = Read-Host "Option"

	switch ($Prompt.ToUpper().ToString()) {
		X {
			if ($Script:Environment.Name) {
				Clear-EnvironmentData
				Get-Menu
			}
			else {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
		}
		L {
			if ($Script:Environment.Name) {
				Clear-EnvironmentData
			}
			Get-EnviromentList
			Get-Menu
		}
		Q {
			Clear-Host
			Clear-EnvironmentData
			exit
		}
		0 {
			Get-EnviromentList
			Get-BackupManager -Backup
			Get-AOSManager -Stop
			SQL-DBRestore
			SQL-CleanUpTable
			Get-BackupManager -Restore
			Get-AOSManager -Start
			Write-Host ''
			Write-Host 'Process completed.' -Fore Green
			Get-MainMenu
		}
		1 {
			Get-EnviromentList
			Get-BackupManager -Backup
			Get-Menu
		}
		2 {
			Get-EnviromentList
			Get-AOSManager -Stop
			Get-Menu
		}
		3 {
			Get-EnviromentList
			SQL-DBRestore
			Get-Menu
		}
		4 {
			Get-EnviromentList
			SQL-CleanUpTable
			Get-Menu
		}
		5 {
			Get-EnviromentList
			Get-BackupManager -Restore
			Get-Menu
		}
		6 {
			Get-EnviromentList
			Get-AOSManager -Start
			Get-Menu
		}
		9 {
			Get-MiscMenu
		}
		Default {
			if (($Prompt -notlike "[LQ01234569]") -and !($Script:Environment.MachineName)) {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
			elseif (($Prompt -notlike "[XLQ01234569]") -and ($Script:Environment.MachineName)) {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
		}
	}

}

function Get-MiscMenu
{
	$MenuName = 'Miscellaneous'
	Write-Host ''
	Write-Host 'Select one of the following options:'
	Write-Host ''
	Write-Host '1. Batch Jobs Maintenance'
	Write-Host '2. Check RecIds'
	Write-Host '3. AX Service Tools'
	Write-Host '4. Delete Environment Store'
	Write-Host '5. Reload Servers'
	Write-Host '6. Change SQL Backup Folder'
	Write-Host '7. Update GUID'
	Write-Host ''
	Write-Host ''
	if ($Script:Environment.MachineName) { Get-MenuConfig }
	Write-Host 'L - Load Environment'
	Write-Host 'R - Return'
	Write-Host 'Q - Quit'
	if ($Script:WarningMsg) {
		Write-Host ''
		Write-Warning $Script:WarningMsg
		Clear-Variable WarningMsg -Scope Script }
	Write-Host '════════════════════════════════════════════════════════════'
	$Prompt = Read-Host "Option"

	switch ($Prompt.ToUpper().ToString()) {
		X {
			if ($Script:Environment.MachineName) {
				Clear-EnvironmentData
				Get-Menu
			}
			else {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
		}
		L {
			if ($Script:Environment.Name) {
				Clear-EnvironmentData
			}
			Get-EnviromentList
			Get-Menu
		}
		R {
			Get-MainMenu
		}
		Q {
			Clear-Host
			exit
		}
		1 {
			Get-BatchMenu
		}
		2 {
			Get-EnviromentList
			Set-TableRecId
			Get-Menu
		}
		3 {
			Get-ServicesMenu
		}
		4 {
			Get-EnviromentList
			Get-StoreManager -DeleteStore
            Clear-EnvironmentData
			Get-Menu
		}
		5 {
			Get-EnviromentList
			#Get-RunningServers
            Get-ServersByVersion
			Get-Menu
		}
		6 {
			Get-EnviromentList
			Set-SQLBKPFolder
			Get-Menu
		}
		7 {
			Get-EnviromentList
			Set-NewAXGuid
			Get-Menu
		}
		Default {
			if (($Prompt -notlike "[LQR123456]") -and !($Script:Environment.MachineName)) {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
			elseif (($Prompt -notlike "[XLQR123456]") -and ($Script:Environment.MachineName)) {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
		}
	}
}

function Get-BatchMenu
{
	$MenuName = 'Batch'
	Write-Host ''
	Write-Host 'Select one of the following options:'
	Write-Host ''
	Write-Host '1. Disable all batch jobs'
	Write-Host '2. Move Batch Groups to a different server (EnableBatch is ON).'
	Write-Host '3. Clean Batch History'
	Write-Host ''
	if ($Script:Environment.MachineName) { Get-MenuConfig }
	Write-Host 'L - Load Environment'
	Write-Host 'R - Return'
	Write-Host 'Q - Quit'
	if ($Script:WarningMsg) {
		Write-Host ''
		Write-Warning $Script:WarningMsg
		Clear-Variable WarningMsg -Scope Script }
	Write-Host '════════════════════════════════════════════════════════════'
	$Prompt = Read-Host "Option"

	switch ($Prompt.ToUpper().ToString()) {
		X {
			if ($Script:Environment.MachineName) {
				Clear-EnvironmentData
				Get-Menu
			}
			else {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
		}
		L {
			if ($Script:Environment.Name) {
				Clear-EnvironmentData
			}
			Get-EnviromentList
			Get-Menu
		}
		R {
			Get-MiscMenu
		}
		Q {
			Clear-Host
			exit
		}
		1 {
			Get-EnviromentList
			Get-BatchManager -DisableJobs
			Get-Menu
		}
		2 {
			Get-EnviromentList
			Get-BatchManager -ChangeServer
			Get-Menu
		}
		3 {
			Get-EnviromentList
			Get-BatchManager -HistoryCleanup
			Get-Menu
		}
		Default {
			if (($Prompt -notlike "[LQR123]") -and !($Script:Environment.MachineName)) {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
			elseif (($Prompt -notlike "[XLQR123]") -and ($Script:Environment.MachineName)) {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
		}
	}
}

function Get-ServicesMenu
{
	$MenuName = 'Services'
	Write-Host ''
	Write-Host 'Select one of the following options:'
	Write-Host ''
	Write-Host '1. Start AOS Services'
	Write-Host '2. Stop AOS Services'
	Write-Host '3. Restart AOS Services'
	Write-Host '4. Check AOS Services Status'
	Write-Host ''
	if ($Script:Environment.MachineName) { Get-MenuConfig }
	Write-Host 'L - Load Environment'
	Write-Host 'R - Return'
	Write-Host 'Q - Quit'
	if ($Script:WarningMsg) {
		Write-Host ''
		Write-Warning $Script:WarningMsg
		Clear-Variable WarningMsg -Scope Script }
	Write-Host '════════════════════════════════════════════════════════════'
	$Prompt = Read-Host "Option"

	switch ($Prompt.ToUpper().ToString()) {
		X {
			if ($Script:Environment.MachineName) {
				Clear-EnvironmentData
				Get-Menu
			}
			else {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
		}
		L {
			if ($Script:Environment.Name) {
				Clear-EnvironmentData
			}
			Get-EnviromentList
			Get-Menu
		}
		R {
			Get-MiscMenu
		}
		Q {
			Clear-Host
			exit
		}
		1 {
			Get-AOSManager -Start
			Get-Menu
		}
		2 {
			Get-AOSManager -Stop
			Get-Menu
		}
		3 {
			Get-AOSManager -Restart
			Get-Menu
		}
		4 {
			Get-AOSManager -Status
			Get-Menu
		}
		Default {
			if (($Prompt -notlike "[LQR1234]") -and !($Script:Environment.MachineName)) {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
			elseif (($Prompt -notlike "[XLQR1234]") -and ($Script:Environment.MachineName)) {
				$Script:WarningMsg = 'Invalid Option. Retry.'
				Get-Menu
			}
		}
	}
}

function Get-MenuConfig
{
	Write-Host 'Connected to:'
	Write-Host 'Source Environment: ' -NoNewline
	Write-Host $Script:Environment.Name -Fore Yellow
	Write-Host 'Machine Name: ' -NoNewline
	Write-Host $Script:Environment.MachineName -Fore Yellow
	Write-Host 'SQL Server: ' -NoNewline
	Write-Host $Script:Environment.keyDbServer -Fore Yellow
	Write-Host 'AX Database Name: ' -NoNewline
	Write-Host $Script:Environment.keyDBName -Fore Yellow
	Write-Host ''
	Write-Host 'X - Release Environment'
}

function Get-EnviromentList
{
	if (!($Script:Environment.Name)) {
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
		$SqlQuery = "SELECT A.ENVIRONMENT as Options 
                        FROM AXTools_Environments A
                        JOIN AXRefresh_EnvironmentsExt B on A.ENVIRONMENT = B.ENVIRONMENT
                        ORDER BY A.CREATEDDATETIME"
		$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
		$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$Adapter.SelectCommand = $SqlCommand
		$RFREnvironments = New-Object System.Data.DataSet
		$EnvTotal = $Adapter.Fill($RFREnvironments)
		$SqlConn.Close()
		$Options = @()
		[array]$Options = $RFREnvironments.Tables[0] | Select-Object Options -Unique
		$Options += @{ Options = "<< Add New >>" }
		$i = 0
		Write-Host ''
		Write-Host 'Choose an Enviroment:'
		foreach ($Option in $Options) {
			$i++
			Write-Host "$i. $($Option.Options)"
		}
		do {
			$Prompt = Read-Host "Option (1/$i)"
		} while (($Prompt -notlike "[1-$i]") -and ($Prompt))

		if (!($Prompt)) {
			Clear-EnvironmentData
			$Script:WarningMsg = 'Invalid Option. Retry.'
			Get-Menu
		}
		elseif ($Options.Count -eq $Prompt) {
			New-Environment
		}
		else {
			$Script:Environment.Name = (($RFREnvironments.Tables[0] | Select-Object Options -Unique)[$Prompt - 1]).Options
			Import-Environment
		}
	}
}

function Import-Environment
{
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
	$SqlConn.Open()
	#
	$SqlQuery = “SELECT MACHINENAME FROM AXRefresh_EnvironmentsExt WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
    $Script:Environment.MachineName = $($SqlCommand.ExecuteScalar())
	#
	$SqlQuery = “SELECT DBSERVER FROM AXTools_Environments WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$Script:Environment.keyDbServer = $SqlCommand.ExecuteScalar()
	#
	$SqlQuery = “SELECT DBNAME FROM AXTools_Environments WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$Script:Environment.keyDBName = $SqlCommand.ExecuteScalar()
    #
	$SqlQuery = “SELECT DBUSER FROM AXTools_Environments WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$Script:Environment.keyDBUser = $SqlCommand.ExecuteScalar()
    #
	Test-SQLSettings $Script:Environment.keyDbServer $Script:Environment.keyDBName
	$SqlQuery = “SELECT EMAILPROFILE FROM AXTools_Environments WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$Script:Environment.EmailProfile = $SqlCommand.ExecuteScalar()
	$SqlConn.Close()
	Get-BackupManager -Check
	Get-EnvironmentServers
}

function Get-EnvironmentServers
{
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
	$SqlConn.Open()
	$SqlQuery = "SELECT SERVERNAME, ACTIVE, AOSID, INSTANCENAME, STATUS, VERSION
                    FROM AXTools_Servers WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND SERVERTYPE = 'AOS'"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$Adapter.SelectCommand = $SqlCommand
	$DSServers = New-Object System.Data.DataSet
	$Adapter.Fill($DSServers) | Out-Null
	$SqlConn.Close()

	Write-Host ''
	Write-Host 'Environment Configuration:'
	Write-Host 'SQL Server Name: ' -NoNewline; Write-Host $Script:Environment.keyDbServer -Fore Yellow
	Write-Host 'AX Database: ' -NoNewline; Write-Host $Script:Environment.keyDBName -Fore Yellow

    if($DSServers.Tables[0].DefaultView.Count -gt 0) {
        $AOSServers = @()
	    foreach($Server in $DSServers.Tables[0]) {
            if([string]::IsNullOrEmpty($Server.InstanceName)) {$Instance='01'} else {$Instance=$Server.InstanceName.Substring(0,2)}
            $AOSTemp = New-Object -TypeName System.Object
            $AOSTemp | Add-Member -Name AosId -Value $($Server.AOSID) -MemberType NoteProperty
            $AOSTemp | Add-Member -Name Instance_Name -Value $($Server.InstanceName) -MemberType NoteProperty
            $AOSTemp | Add-Member -Name Status -Value $($Server.Status) -MemberType NoteProperty
            $AOSTemp | Add-Member -Name ServerName -Value $($Server.ServerName) -MemberType NoteProperty
            if (!(Test-Connection $AOSTemp.ServerName -Count 1 -Quiet)) {
                $AOSTemp | Add-Member -Name ServerStatus -Value "Can't Reach Server" -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Active -Value '0' -MemberType NoteProperty
                $AOSTemp | Add-Member -Name UpdateFlag -Value '1' -MemberType NoteProperty
            }
            else {
                if($Server.Version -eq '7' -and $Server.InstanceName -like 'OneBox*') {
                    $Service = Get-WmiObject -Class Win32_Service -ComputerName $AOSTemp.ServerName -ea 0 | Where-Object { $_.DisplayName -like "World Wide Web Publishing Service" }
                    if(![string]::IsNullOrEmpty($Service)) {
                        $AOSTemp | Add-Member -Name Active -Value '1' -MemberType NoteProperty
                        $AOSTemp | Add-Member -Name ServiceStatus -Value "IIS:$($Service.State)" -MemberType NoteProperty
                        if($Service.State -like 'Running' -and $AOSTemp.Status -eq 0) {
                            $AOSTemp | Add-Member -Name UpdateFlag -Value '1' -MemberType NoteProperty
                        }
                        elseif($Service.State -like 'Stopped' -and $AOSTemp.Status -eq 1) {
                            $AOSTemp | Add-Member -Name UpdateFlag -Value '1' -MemberType NoteProperty
                        }
                        else {
                            $AOSTemp | Add-Member -Name UpdateFlag -Value '0' -MemberType NoteProperty
                        }
                    }
                    else {
                        $AOSTemp | Add-Member -Name ServerStatus -Value "Can't connect to service" -MemberType NoteProperty
                        $AOSTemp | Add-Member -Name Active -Value '0' -MemberType NoteProperty
                        $AOSTemp | Add-Member -Name UpdateFlag -Value '1' -MemberType NoteProperty
                    }
                }
                elseif($Server.Version -eq '7' -and $Server.InstanceName -like 'Azure*') {
                    $Service = Invoke-Command -Computer $($AOSTemp.ServerName) -ScriptBlock { 
                        Connect-ServiceFabricCluster -TimeoutSec 1 | Out-Null
                        Get-ServiceFabricApplication -ExcludeApplicationParameters | Where { $_.ApplicationName -like '*AX*' }
                    } -ErrorAction SilentlyContinue
                    if(![string]::IsNullOrEmpty($Service)) {
                        $AOSTemp | Add-Member -Name Active -Value '1' -MemberType NoteProperty
                        $AOSTemp | Add-Member -Name ServiceStatus -Value "HealthState $($Service.HealthState)" -MemberType NoteProperty
                        if($Service.State -like 'Ok' -and $AOSTemp.Status -eq 0) {
                            $AOSTemp | Add-Member -Name UpdateFlag -Value '1' -MemberType NoteProperty
                        }
                        elseif($Service.State -like 'Error' -and $AOSTemp.Status -eq 1) {
                            $AOSTemp | Add-Member -Name UpdateFlag -Value '1' -MemberType NoteProperty
                        }
                        else {
                            $AOSTemp | Add-Member -Name UpdateFlag -Value '0' -MemberType NoteProperty
                        }
                    }
                    else {
                        $AOSTemp | Add-Member -Name ServerStatus -Value "Can't connect to service" -MemberType NoteProperty
                        $AOSTemp | Add-Member -Name Active -Value '0' -MemberType NoteProperty
                        $AOSTemp | Add-Member -Name UpdateFlag -Value '1' -MemberType NoteProperty
                    }
                }
                else {
                    $Service = Get-WmiObject -Class Win32_Service -ComputerName $AOSTemp.ServerName -ea 0 | Where-Object { $_.DisplayName -like "Microsoft Dynamics AX Object Server*" -and $_.Name.Substring($_.Name.Length-2,2) -like $Instance }
                    if(![string]::IsNullOrEmpty($Service)) {
                        $AOSTemp | Add-Member -Name Active -Value '1' -MemberType NoteProperty
                        $AOSTemp | Add-Member -Name ServiceStatus -Value $($Service.State) -MemberType NoteProperty
                        if($Service.State -like 'Running' -and $AOSTemp.Status -eq 0) {
                            $AOSTemp | Add-Member -Name UpdateFlag -Value '1' -MemberType NoteProperty
                        }
                        elseif($Service.State -like 'Stopped' -and $AOSTemp.Status -eq 1) {
                            $AOSTemp | Add-Member -Name UpdateFlag -Value '1' -MemberType NoteProperty
                        }
                        else {
                            $AOSTemp | Add-Member -Name UpdateFlag -Value '0' -MemberType NoteProperty
                        }
                    }
                    else {
                        $AOSTemp | Add-Member -Name ServerStatus -Value "Can't connect to service" -MemberType NoteProperty
                        $AOSTemp | Add-Member -Name Active -Value '0' -MemberType NoteProperty
                        $AOSTemp | Add-Member -Name UpdateFlag -Value '1' -MemberType NoteProperty
                    }
                }
	        }
            $AOSServers += $AOSTemp
        }
        $DSServers.Dispose()
	    Write-Host ''
	    Write-Host 'AOS Servers/Nodes Status:'
	    $i = 1
	    $AOSServers | ForEach-Object {
            if($_.UpdateFlag -eq 1) {
				$SqlConn = New-Object System.Data.SqlClient.SqlConnection
				$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
				$SqlConn.Open()
				$SqlQuery = “UPDATE [AXTools_Servers] SET ACTIVE = '$($_.Active)', STATUS = '$(if($_.ServiceStatus -match 'Running|Ok'){'1'} else {'0'})'
                                WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND SERVERNAME = '$($_.ServerName)' AND AOSID = '$($_.AOSID)'"
				$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
				$SqlCommand.ExecuteNonQuery() | Out-Null
				$SqlConn.Close()
            }
		    Write-Host "$i. $($_.ServerName) " -Fore Yellow -NoNewline

            if(!$_.ServiceStatus) { 
                Write-Host '' -NoNewline 
            } 
            elseif($_.ServiceStatus -match 'Running|Ok') { 
                Write-Host "($($_.ServiceStatus)) " -Fore Green -NoNewline 
            } 
            else { 
                Write-Host "($($_.ServiceStatus)) " -Fore Red -NoNewline
            }

            if($_.ServerStatus) { 
                Write-Host '- ' -NoNewline 
                Write-Warning $($_.ServerStatus)
            } 
            else { 
                Write-Host '' 
            }
		    $i++
	    }
    }
    else {
	    $Script:WarningMsg = 'No servers in the database (Run 9 > 5 to refresh servers).'
    }
}

function New-Environment
{
    $Script:Environment.Name = (Read-Host "Enter Source Environment").ToUpper()
	if ($Script:Environment.Name.Length -gt 30) {
		Write-Warning "Environment name must have less than 30 chars."
		Clear-EnvironmentData
		New-Environment
	}
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
	$SqlConn.Open()
	$SqlQuery = “SELECT ENVIRONMENT FROM [AXRefresh_EnvironmentStore] WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND DELETED = 0"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$EnvCnt = $SqlCommand.ExecuteScalar()
	$SqlConn.Close() | Out-Null
	if ($EnvCnt -ge 1) {
		$Prompt = Read-Host "Found Backup for $($Script:Environment.Name). Load Configuration (Y/N)"
		switch ($Prompt.ToUpper()) {
			'Y' {
				Import-Environment
			}
			'N' {
				Clear-EnvironmentData
				New-Environment
			}
		}
	}
	else {
		Get-AOSConfiguration
	}
}

function Get-AOSConfiguration
{
    $Script:Environment.MachineName = ((Read-Host "Enter AOS Server or ServiceFabric").ToUpper())
	if (!$Script:Environment.MachineName) {
		$Script:WarningMsg = 'Invalid AOS/SF Server. Retry.'
		Clear-EnvironmentData
		Get-Menu
	}

	if (!(Test-Connection $Script:Environment.MachineName -Count 1 -Quiet)) {
		$Script:WarningMsg = 'Invalid AOS/SF Server. Retry.'
		Clear-EnvironmentData
		Get-Menu
	}

	## Connecting to AOS Server Registry
    $AOSKey = Invoke-Command -Computer $($Script:Environment.MachineName) { Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services\Dynamics Server' } -ErrorAction SilentlyContinue
    
    if([string]::IsNullOrEmpty($AOSKey)) {
        $ServiceFabric  = Invoke-Command -Computer $($Script:Environment.MachineName) -ScriptBlock { 
            Connect-ServiceFabricCluster -TimeoutSec 1 | Out-Null
            Get-ServiceFabricApplication | Where { $_.ApplicationName -like '*AX*' }
        } -ErrorAction SilentlyContinue
        if(![string]::IsNullOrEmpty($ServiceFabric)) {
            $Script:Environment.keyDBName = ($ServiceFabric.ApplicationParameters | Where { $_.Name -eq 'DataAccess_Database' }).Value
            $Script:Environment.keyDbServer = ($ServiceFabric.ApplicationParameters | Where { $_.Name -eq 'DataAccess_DbServer' }).Value
            Get-RunningApps
            $GetSqlCreds = $true
        }
    }

    if([string]::IsNullOrEmpty($ServiceFabric)) {
        $OneBox  = Invoke-Command -Computer $($Script:Environment.MachineName) -ScriptBlock { 
            Get-WebFilePath "IIS:\Sites\AOSService"
        } -ErrorVariable Err -ErrorAction SilentlyContinue
        if(!$Err -and ![string]::IsNullOrEmpty($OneBox)) {
            [Xml]$WebConfig = Get-Content "$($OneBox.FullName)\Web.Config"
            $Script:Environment.keyDbServer = ($WebConfig.Configuration.AppSettings.Add | where { $_.Key -eq 'DataAccess.DbServer' }).Value
            $Script:Environment.keyDBName = ($WebConfig.Configuration.AppSettings.Add | where { $_.Key -eq 'DataAccess.Database' }).Value
            if($Script:Environment.keyDbServer -like 'localhost') {
                $Script:Environment.keyDbServer = $Script:Environment.MachineName
            }
            Get-ServerOneBox        
        }
        else {
            Write-Host ''
            Write-Warning $Err[0]
			Clear-EnvironmentData
			Get-Menu
        }
    }

    if(!$ServiceFabric -and !$OneBox) {
	    foreach ($AOSVersion in $($AOSKey | Where { $_.PSChildName.Substring(0,1) -match "^[0-9]*$" })) {
		    switch ($AOSVersion.PSChildName.Substring(0,1)) {
			    "5" { $Version = "AX2009" }
			    "6" { $Version = "AX2012" }
		    }
		    $AOSInstances = Invoke-Command -Computer $($Script:Environment.MachineName) -ArgumentList $AOSVersion.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") { Get-ChildItem $args[0] }
            
            if($AOSInstances.Name.Count -gt 1) {
                $i = 0
		        [array]$Options = $AOSInstances
		        Write-Host ''
		        Write-Host 'Choose an instance:'
		        foreach ($Option in $Options) {
			        $i++
			        Write-Host "$i. $($Option.Name.Substring($Option.Name.Length-2,2))"
		        }
		        do {
			        $Prompt = Read-Host "Option (1/$i)"
		        } while (($Prompt -notlike "[1-$i]") -and ($Prompt))

		        if (!($Prompt)) {
			        Clear-EnvironmentData
			        $Script:WarningMsg = 'Invalid Option. Retry.'
			        Get-Menu
		        }
		        else {
			        $Instance = ($AOSInstances[$Prompt - 1]).Name
		        }
            }
            else {
                $Instance = ($AOSInstances[0]).Name
            }

            try {
                $Script:Environment.MachineInstance = $($Instance.Substring($Instance.Length-2,2))
                $Script:Environment.MachineFullName = "$($Script:Environment.MachineInstance)@$($Script:Environment.MachineName)"
			    $Current = Invoke-Command -Computer $($Script:Environment.MachineName) -ArgumentList $Instance.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).Current }
			    $CurrentKey = "$Instance\$Current"
		        $Script:Environment.keyDbServer = Invoke-Command -Computer $($Script:Environment.MachineName) -ArgumentList $CurrentKey.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).DBServer }
		        $Script:Environment.keyDBName = Invoke-Command -Computer $($Script:Environment.MachineName) -ArgumentList $CurrentKey.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).Database }
            }
            catch {
		        $Script:WarningMsg = "AX Configuration not found."
		        Clear-EnvironmentData
		        Get-Menu
            }
        }
    Get-RunningServers
    }

    try {
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
		$SqlConn.Open()
		$SqlQuery = “INSERT INTO [AXTools_Environments] (ENVIRONMENT, DESCRIPTION, DBSERVER, DBNAME)
                        VALUES('$($Script:Environment.Name)','$($Script:Environment.Name)','$($Script:Environment.keyDbServer)','$($Script:Environment.keyDBName)')"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		$SqlCommand.ExecuteNonQuery() | Out-Null
        #
		$SqlQuery = “INSERT INTO [AXRefresh_EnvironmentsExt] (ENVIRONMENT, MACHINENAME)
                        VALUES('$($Script:Environment.Name)','$($Script:Environment.MachineName)')"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		$SqlCommand.ExecuteNonQuery() | Out-Null
		$SqlConn.Close()
    }
    catch {
		$Script:WarningMsg = "AX Configuration not found."
		Clear-EnvironmentData
		Get-Menu
    }
    if($GetSqlCreds) {
        Get-ServiceFabricSQLUser
        $GetSqlCreds = $false
    }
}

function Get-ServersByVersion
{
    $AOSKey = Invoke-Command -Computer $($Script:Environment.MachineName) { Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services\Dynamics Server' } -ErrorAction SilentlyContinue
    if([string]::IsNullOrEmpty($AOSKey)) {
        $ServiceFabric  = Invoke-Command -Computer $($Script:Environment.MachineName) -ScriptBlock { 
            Connect-ServiceFabricCluster -TimeoutSec 1 | Out-Null
            Get-ServiceFabricApplication | Where { $_.ApplicationName -like '*AX*' }
        } -ErrorAction SilentlyContinue
        if(![string]::IsNullOrEmpty($ServiceFabric)) {
            Get-RunningApps
        }
    }
    if([string]::IsNullOrEmpty($ServiceFabric)) {
        $OneBox  = Invoke-Command -Computer $($Script:Environment.MachineName) -ScriptBlock { 
            Get-WebFilePath "IIS:\Sites\AOSService"
        } -ErrorVariable Err -ErrorAction SilentlyContinue
        if(!$Err -and ![string]::IsNullOrEmpty($OneBox)) {
            Get-ServerOneBox
        }
    }

    if(!$ServiceFabric -and !$OneBox) {
        Get-RunningServers
    }

}

function Get-RunningServers
{
	try {
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$($Script:Environment.keyDbServer);Database=$($Script:Environment.keyDBName);Integrated Security=True"
		$SqlConn.Open()
		$SqlQuery = "SELECT SERVERID, AOSID, INSTANCE_NAME, VERSION, STATUS
		                , TXT_STATUS = CASE STATUS 
			                WHEN 1 THEN 'Running'
			                WHEN 0 THEN 'Stopped'
			                END
                    FROM SYSSERVERSESSIONS"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$Adapter.SelectCommand = $SqlCommand
		$DSServers = New-Object System.Data.DataSet
		$Adapter.Fill($DSServers) | Out-Null
		$SqlConn.Close()
	}
	catch {
		$Script:WarningMsg = "Error conneting to SQL Server " + $($Script:Environment.keyDbServer) + " and " + $($Script:Environment.keyDBName)
		Get-Menu
	}

	Write-Host ''
	Write-Host 'Environment Configuration:'
	Write-Host 'SQL Server Name: ' -NoNewline; Write-Host $Script:Environment.keyDbServer -Fore Yellow
	Write-Host 'AX Database: ' -NoNewline; Write-Host $Script:Environment.keyDBName -Fore Yellow
    
    if($DSServers.Tables.Rows.Count -gt 0) {
        $AOSServers = @()
	    foreach($Server in $DSServers.Tables[0]) {
            $AOSTemp = New-Object -TypeName System.Object
            $AOSTemp | Add-Member -Name AosId -Value $($Server.AOSID) -MemberType NoteProperty
            $AOSTemp | Add-Member -Name Instance_Name -Value $($Server.Instance_Name) -MemberType NoteProperty
            $AOSTemp | Add-Member -Name Version -Value $($Server.Version) -MemberType NoteProperty
            $AOSTemp | Add-Member -Name Status -Value $($Server.Status) -MemberType NoteProperty
            $AOSTemp | Add-Member -Name ServerName -Value $($Server.AOSID.Substring(0,$Server.AOSID.Length-5)) -MemberType NoteProperty
            
            if (!(Test-Connection $AOSTemp.ServerName -Count 1 -Quiet)) {
                $AOSTemp | Add-Member -Name ServerStatus -Value "Can't Reach Server" -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Active -Value '0' -MemberType NoteProperty
            }
            else {
                $AOSKey = Invoke-Command -Computer ($Server.AOSID.Substring(0,$Server.AOSID.Length-5)) { Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services\Dynamics Server' }
                Invoke-Command -Computer ($Server.AOSID.Substring(0,$Server.AOSID.Length-5)) -ArgumentList $AOSKey.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") { Get-ChildItem $args[0] } | % {
                    if($Server.INSTANCE_NAME -match $_.Name.Substring($_.Name.Length-2,2)) {
                        $Instance = $_
                        Invoke-Command -Computer ($Server.AOSID.Substring(0,$Server.AOSID.Length-5)) -ArgumentList $_.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).Current } | % {
                            $CheckDbServer = Invoke-Command -Computer ($Server.AOSID.Substring(0,$Server.AOSID.Length-5)) -ArgumentList $("$Instance\$_").Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).DBServer }
                            $CheckDbName = Invoke-Command -Computer ($Server.AOSID.Substring(0,$Server.AOSID.Length-5)) -ArgumentList $("$Instance\$_").Replace("HKEY_LOCAL_MACHINE","HKLM:") { (Get-ItemProperty $args[0]).Database }
                        }
                    }
                }
                if($CheckDbServer -like $Script:Environment.keyDbServer -and $CheckDbName -like $Script:Environment.keyDbName) {
                    $AOSTemp | Add-Member -Name Active -Value '1' -MemberType NoteProperty
                    $Service = Get-WmiObject -Class Win32_Service -ComputerName $AOSTemp.ServerName -ea 0 | Where-Object { $_.DisplayName -like "Microsoft Dynamics AX Object Server*" -and $_.Name.Substring($_.Name.Length-2,2) -like $Server.Instance_Name.Substring(0,2) }
                    $AOSTemp | Add-Member -Name ServiceStatus -Value $($Service.State) -MemberType NoteProperty
                    $AOSTemp | Add-Member -Name ServiceName -Value $($Service.Name) -MemberType NoteProperty
                    $AOSTemp | Add-Member -Name ServerIp -Value ((Test-Connection $AOSTemp.ServerName -Count 1 -ErrorAction SilentlyContinue).IPV4Address).IPAddressToString -MemberType NoteProperty
                    $AOSTemp | Add-Member -Name Domain -Value (Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges -ComputerName $AOSTemp.ServerName -ErrorAction SilentlyContinue).Domain -MemberType NoteProperty
                    $AOSTemp | Add-Member -Name FQDN -Value "$($AOSTemp.ServerName).$($AOSTemp.Domain)" -MemberType NoteProperty
                }
                else {
                    $AOSTemp | Add-Member -Name ServerStatus -Value "DB Mismatch: $CheckDbServer\$CheckDbName" -MemberType NoteProperty
                    $AOSTemp | Add-Member -Name Active -Value '0' -MemberType NoteProperty
                }
	        }
            $AOSServers += $AOSTemp
        }
        $DSServers.Dispose()
	    Write-Host ''
	    Write-Host 'AOS Servers Status:'
	    $i = 1
	    $AOSServers | Sort ServerName -Unique | ForEach-Object {
		    Write-Host "$i. $($_.ServerName) " -Fore Yellow -NoNewline

            if(!$_.ServiceStatus) { 
                Write-Host '' -NoNewline
            } 
            elseif($_.ServiceStatus -like 'Running') {
                Write-Host "($($_.ServiceStatus)) " -Fore Green -NoNewline
            } 
            else { 
                Write-Host "($($_.ServiceStatus)) " -Fore Red -NoNewline
            }

            if($_.ServerStatus) {
                Write-Host '- ' -NoNewline 
                Write-Warning $($_.ServerStatus)
            } 
            else {
                Write-Host ''
            }
		    $i++
	    }
	    do {
		    if ($Script:Environment.AsJob) { $Prompt = 'Y' } else { $Prompt = Read-Host "Continue? (Y/N)" }
		    switch ($Prompt.ToUpper()) {
			    Y {
				    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
				    $SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
				    $SqlConn.Open()
				    if ($Script:Environment.HasServers) {
					    $SqlQuery = “DELETE FROM [AXTools_Servers] WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND SERVERTYPE = 'AOS' AND CREATEDDATETIME <= '$DateTime'"
					    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					    $SqlCommand.ExecuteNonQuery() | Out-Null
				    }
                    $AOSServers | Where-Object { $_.Active -eq '1' } | ForEach-Object {
					    $SqlQuery = “INSERT INTO [AXTools_Servers] (ENVIRONMENT,ACTIVE,SERVERNAME,SERVERTYPE,IP,DOMAIN,FQDN,AOSID,INSTANCENAME,VERSION,STATUS)
                                        VALUES('$($Script:Environment.Name)','$($_.Active)','$($_.ServerName)','AOS','$($_.ServerIp)','$($_.Domain)','$($_.FQDN)','$($_.AOSID)','$("$($_.INSTANCE_NAME)`@$($_.AOSID.Substring(0,$_.AOSID.Length-5))")','$($_.VERSION)','$(if($_.ServiceStatus -like 'Running'){'1'} else {'0'})')"
					    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					    $SqlCommand.ExecuteNonQuery() | Out-Null
				    }
					$SqlQuery = “UPDATE [AXRefresh_EnvironmentsExt] SET [MACHINENAME] = '$($AOSServers.ServerName | Select -First 1)' WHERE [ENVIRONMENT] = '$($Script:Environment.Name)'"
					$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					$SqlCommand.ExecuteNonQuery() | Out-Null
				    $SqlConn.Close()
			    }
			    N {
				    $Script:WarningMsg = "Canceled!"
				    Clear-EnvironmentData
				    Get-Menu
			    }
		    }
	    } while ($Prompt -notmatch "[YN]")
    }
    else {
	    $Script:WarningMsg = 'No servers to show.'
    }
}

function Get-RunningApps
{
	try {
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$($Script:Environment.keyDbServer);Database=$($Script:Environment.keyDBName);Integrated Security=True"
		$SqlConn.Open()
		$SqlQuery = "SELECT SERVERID, AOSID, INSTANCE_NAME, VERSION, STATUS
		                , TXT_STATUS = CASE STATUS 
			                WHEN 1 THEN 'Running'
			                WHEN 0 THEN 'Stopped'
			                END
                    FROM SYSSERVERSESSIONS
                    WHERE INSTANCE_NAME = 'Azure'"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$Adapter.SelectCommand = $SqlCommand
		$DSServers = New-Object System.Data.DataSet
		$Adapter.Fill($DSServers) | Out-Null
		$SqlConn.Close()
	}
	catch {
		$Script:WarningMsg = "Error conneting to SQL Server " + $($Script:Environment.keyDbServer) + " and " + $($Script:Environment.keyDBName)
		Get-Menu
	}

	Write-Host ''
	Write-Host 'Environment Configuration:'
	Write-Host 'SQL Server Name: ' -NoNewline; Write-Host $Script:Environment.keyDbServer -Fore Yellow
	Write-Host 'AX Database: ' -NoNewline; Write-Host $Script:Environment.keyDBName -Fore Yellow
    
    if($DSServers.Tables.Rows.Count -gt 0) {
        $AOSServers = @()
	    foreach($Server in $DSServers.Tables[0]) {
            $AOSTemp = New-Object -TypeName System.Object
            $AOSTemp | Add-Member -Name AOSID -Value $($Server.AOSID) -MemberType NoteProperty
            $AOSTemp | Add-Member -Name Instance_Name -Value 'Azure' -MemberType NoteProperty
            $AOSTemp | Add-Member -Name Version -Value '7' -MemberType NoteProperty
            $AOSTemp | Add-Member -Name Status -Value $($Server.Status) -MemberType NoteProperty
            $AOSTemp | Add-Member -Name ServerName -Value $(if($($Server.AOSID).Split(':').Count -gt 1) { ($Server.AOSID).Split(':')[1] } else { ($Server.AOSID).Split('.')[0] } ) -MemberType NoteProperty
            
            if (!(Test-Connection $AOSTemp.ServerName -Count 1 -Quiet)) {
                $AOSTemp | Add-Member -Name ServerStatus -Value "Can't Reach Server" -MemberType NoteProperty
                $AOSTemp | Add-Member -Name Active -Value '0' -MemberType NoteProperty
            }
            else {
                $ServiceFabric  = Invoke-Command -Computer $($AOSTemp.ServerName) -ScriptBlock { 
                    Connect-ServiceFabricCluster -TimeoutSec 1 | Out-Null
                    Get-ServiceFabricApplication | Where { $_.ApplicationName -like '*AX*' }
                } -ErrorAction SilentlyContinue
                $CheckDbName = ($ServiceFabric.ApplicationParameters | Where { $_.Name -eq 'DataAccess_Database' }).Value
                $CheckDbServer = ($ServiceFabric.ApplicationParameters | Where { $_.Name -eq 'DataAccess_DbServer' }).Value
                if($CheckDbServer -like $Script:Environment.keyDbServer -and $CheckDbName -like $Script:Environment.keyDbName) {
                    $AOSTemp | Add-Member -Name Active -Value '1' -MemberType NoteProperty
                    $Service = Invoke-Command -Computer $($AOSTemp.ServerName) -ScriptBlock { 
                        Connect-ServiceFabricCluster -TimeoutSec 1 | Out-Null
                        Get-ServiceFabricApplication -ExcludeApplicationParameters | Where { $_.ApplicationName -like '*AX*' }
                    }
                    $AOSTemp | Add-Member -Name ServiceStatus -Value $($Service.HealthState) -MemberType NoteProperty
                    $AOSTemp | Add-Member -Name ServiceName -Value $($Service.ApplicationName) -MemberType NoteProperty
                    $AOSTemp | Add-Member -Name ServerIp -Value ((Test-Connection $AOSTemp.ServerName -Count 1 -ErrorAction SilentlyContinue).IPV4Address).IPAddressToString -MemberType NoteProperty
                    $AOSTemp | Add-Member -Name Domain -Value (Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges -ComputerName $AOSTemp.ServerName -ErrorAction SilentlyContinue).Domain -MemberType NoteProperty
                    $AOSTemp | Add-Member -Name FQDN -Value "$($AOSTemp.ServerName).$($AOSTemp.Domain)" -MemberType NoteProperty
                }
                else {
                    $AOSTemp | Add-Member -Name ServerStatus -Value "DB Mismatch: $CheckDbServer\$CheckDbName" -MemberType NoteProperty
                    $AOSTemp | Add-Member -Name Active -Value '0' -MemberType NoteProperty
                }
	        }
            $AOSServers += $AOSTemp
        }
        $DSServers.Dispose()
	    Write-Host ''
	    Write-Host 'AOS Servers Status:'
	    $i = 1
	    $AOSServers | Sort ServerName -Unique | ForEach-Object {
		    Write-Host "$i. $($_.ServerName) " -Fore Yellow -NoNewline

            if(!$_.ServiceStatus) { 
                Write-Host '' -NoNewline
            } 
            elseif($_.ServiceStatus -like 'Ok') {
                Write-Host "($($_.ServiceStatus))" -Fore Green -NoNewline
            } 
            else { 
                Write-Host "($($_.ServiceStatus))" -Fore Red -NoNewline
            }

            if($_.ServerStatus) {
                Write-Host '- ' -NoNewline 
                Write-Warning "($($_.ServerStatus))"
            } 
            else {
                Write-Host ''
            }
		    $i++
	    }
	    do {
		    if ($Script:Environment.AsJob) { $Prompt = 'Y' } else { $Prompt = Read-Host "Continue? (Y/N)" }
		    switch ($Prompt.ToUpper()) {
			    Y {
				    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
				    $SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
				    $SqlConn.Open()
				    if ($Script:Environment.HasServers) {
					    $SqlQuery = “DELETE FROM [AXTools_Servers] WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND CREATEDDATETIME <= '$DateTime'"
					    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					    $SqlCommand.ExecuteNonQuery() | Out-Null
				    }
                    $AOSServers | Where-Object { $_.Active -eq '1' } | ForEach-Object {
					    $SqlQuery = “INSERT INTO [AXTools_Servers] (ENVIRONMENT,ACTIVE,SERVERNAME,SERVERTYPE,IP,DOMAIN,FQDN,AOSID,INSTANCENAME,VERSION,STATUS)
                                        VALUES('$($Script:Environment.Name)','$($_.Active)','$($_.ServerName)','AOS','$($_.ServerIp)','$($_.Domain)','$($_.FQDN)','$($_.AOSID)','$("$($_.INSTANCE_NAME)`@$($_.AOSID)")','$($_.VERSION)','$(if($_.ServiceStatus -like 'Ok'){'1'} else {'0'})')"
					    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					    $SqlCommand.ExecuteNonQuery() | Out-Null
				    }
					$SqlQuery = “UPDATE [AXRefresh_EnvironmentsExt] SET [MACHINENAME] = '$($AOSServers.ServerName | Select -First 1)' WHERE [ENVIRONMENT] = '$($Script:Environment.Name)'"
					$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					$SqlCommand.ExecuteNonQuery() | Out-Null
				    $SqlConn.Close()
                    Get-ServiceFabricNodes
			    }
			    N {
				    $Script:WarningMsg = "Canceled!"
				    Clear-EnvironmentData
				    Get-Menu
			    }
		    }
	    } while ($Prompt -notmatch "[YN]")
    }
    else {
	    $Script:WarningMsg = 'No servers to show.'
    }
}

function Get-ServerOneBox
{
	Write-Host ''
	Write-Host 'Environment Configuration:'
	Write-Host 'SQL Server Name: ' -NoNewline; Write-Host $Script:Environment.keyDbServer -Fore Yellow
	Write-Host 'AX Database: ' -NoNewline; Write-Host $Script:Environment.keyDBName -Fore Yellow
    
    $AOSTemp = New-Object -TypeName System.Object
    $AOSTemp | Add-Member -Name AOSID -Value "OneBox:$($Script:Environment.MachineName)" -MemberType NoteProperty
    $AOSTemp | Add-Member -Name Instance_Name -Value 'OneBox' -MemberType NoteProperty
    $AOSTemp | Add-Member -Name Version -Value '7' -MemberType NoteProperty
    $AOSTemp | Add-Member -Name Status -Value '1' -MemberType NoteProperty
    $AOSTemp | Add-Member -Name ServerName -Value $($Script:Environment.MachineName) -MemberType NoteProperty
    $AOSTemp | Add-Member -Name ServiceStatus -Value 'Ok' -MemberType NoteProperty
    $AOSTemp | Add-Member -Name ServerIp -Value ((Test-Connection $AOSTemp.ServerName -Count 1 -ErrorAction SilentlyContinue).IPV4Address).IPAddressToString -MemberType NoteProperty
    $AOSTemp | Add-Member -Name Domain -Value (Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges -ComputerName $AOSTemp.ServerName -ErrorAction SilentlyContinue).Domain -MemberType NoteProperty
    $AOSTemp | Add-Member -Name FQDN -Value "$($AOSTemp.ServerName).$($AOSTemp.Domain)" -MemberType NoteProperty
    $AOSTemp | Add-Member -Name Active -Value '1' -MemberType NoteProperty

    Write-Host ''
    Write-Host 'AOS Servers Status:'
    $i = 1
    $AOSTemp | ForEach-Object {
	    Write-Host "$i. $($_.ServerName) " -Fore Yellow -NoNewline
        if(!$_.ServiceStatus) { 
            Write-Host '' -NoNewline
        } 
        elseif($_.ServiceStatus -like 'Ok') {
            Write-Host "($($_.ServiceStatus))" -Fore Green -NoNewline
        } 
        else { 
            Write-Host "($($_.ServiceStatus))" -Fore Red -NoNewline
        }

        if($_.ServerStatus) {
            Write-Host '- ' -NoNewline 
            Write-Warning "($($_.ServerStatus))"
        } 
        else {
            Write-Host ''
        }
	    $i++
    }
    do {
	    if ($Script:Environment.AsJob) { $Prompt = 'Y' } else { $Prompt = Read-Host "Continue? (Y/N)" }
	    switch ($Prompt.ToUpper()) {
		    Y {
			    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
			    $SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
			    $SqlConn.Open()
			    if ($Script:Environment.HasServers) {
				    $SqlQuery = “DELETE FROM [AXTools_Servers] WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND SERVERTYPE = 'AOS' AND CREATEDDATETIME <= '$DateTime'"
				    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
				    $SqlCommand.ExecuteNonQuery() | Out-Null
			    }
                $AOSTemp | Where-Object { $_.Active -eq '1' } | ForEach-Object {
				    $SqlQuery = “INSERT INTO [AXTools_Servers] (ENVIRONMENT,ACTIVE,SERVERNAME,SERVERTYPE,IP,DOMAIN,FQDN,AOSID,INSTANCENAME,VERSION,STATUS)
                                    VALUES('$($Script:Environment.Name)','$($_.Active)','$($_.ServerName)','AOS','$($_.ServerIp)','$($_.Domain)','$($_.FQDN)','$($_.AOSID)','$($_.Instance_Name)','$($_.VERSION)','$(if($_.ServiceStatus -like 'Ok'){'1'} else {'0'})')"
				    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
				    $SqlCommand.ExecuteNonQuery() | Out-Null
			    }
			    $SqlQuery = “UPDATE [AXRefresh_EnvironmentsExt] SET [MACHINENAME] = '$($AOSTemp.ServerName | Select -First 1)' WHERE [ENVIRONMENT] = '$($Script:Environment.Name)'"
			    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
			    $SqlCommand.ExecuteNonQuery() | Out-Null
			    $SqlConn.Close()
		    }
		    N {
			    $Script:WarningMsg = "Canceled!"
			    Clear-EnvironmentData
			    Get-Menu
		    }
	    }
    } while ($Prompt -notmatch "[YN]")
}

function Get-ServiceFabricNodes
{
    [Xml]$SFNodes  = Invoke-Command -Computer $($Script:Environment.MachineName) -ScriptBlock { 
        Connect-ServiceFabricCluster -TimeoutSec 1 | Out-Null
        Get-ServiceFabricNodeConfiguration
    } -ErrorAction SilentlyContinue
    $Nodes = @()
    foreach($Node in $SFNodes.ClusterManifest.Infrastructure.WindowsServer.NodeList.ChildNodes) {
        $ServerInfo = [System.Net.Dns]::GetHostByAddress($Node.IPAddressOrFQDN)
        $NodeTemp = New-Object -TypeName System.Object
        if (!(Test-Connection $ServerInfo.HostName -Count 1 -Quiet)) {
            $NodeTemp | Add-Member -Name ServerStatus -Value "Error" -MemberType NoteProperty
            $NodeTemp | Add-Member -Name Active -Value '0' -MemberType NoteProperty
        }
        else {
            $NodeTemp | Add-Member -Name ServerStatus -Value 'Ok' -MemberType NoteProperty
            $NodeTemp | Add-Member -Name Active -Value '1' -MemberType NoteProperty
        }
        $NodeTemp | Add-Member -Name ServerName -Value $(($ServerInfo.HostName).Split('.')[0].ToUpper()) -MemberType NoteProperty
        $NodeTemp | Add-Member -Name Type -Value $(($Node.NodeTypeRef).Replace('Type','')) -MemberType NoteProperty
        $NodeTemp | Add-Member -Name ServerIp -Value $($Node.IPAddressOrFQDN) -MemberType NoteProperty
        $NodeTemp | Add-Member -Name Domain -Value (Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges -ComputerName $NodeTemp.ServerName -ErrorAction SilentlyContinue).Domain -MemberType NoteProperty
        $NodeTemp | Add-Member -Name FQDN -Value "$(($ServerInfo.HostName).ToUpper())" -MemberType NoteProperty
        $NodeTemp | Add-Member -Name AOSID -Value "$(($ServerInfo.HostName).Split('.')[0].ToUpper())" -MemberType NoteProperty
        $NodeTemp | Add-Member -Name Instance_Name -Value 'ServiceFabric Node' -MemberType NoteProperty
        $NodeTemp | Add-Member -Name Version -Value '7' -MemberType NoteProperty
        $Nodes += $NodeTemp
    }
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
	$SqlConn.Open()
    $Nodes | ForEach-Object {
		$SqlQuery = “INSERT INTO [AXTools_Servers] (ENVIRONMENT,ACTIVE,SERVERNAME,SERVERTYPE,IP,DOMAIN,FQDN,AOSID,INSTANCENAME,VERSION,STATUS)
                        VALUES('$($Script:Environment.Name)','$($_.Active)','$($_.ServerName)','$($_.Type)','$($_.ServerIp)','$($_.Domain)','$($_.FQDN)','$($_.AOSID)','$($_.INSTANCE_NAME)','$($_.VERSION)','$(if($_.ServerStatus -like 'Ok'){'1'} else {'0'})')"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		$SqlCommand.ExecuteNonQuery() | Out-Null
	}
	$SqlConn.Close()
}

function Get-ServiceFabricSQLUser
{
    $ServiceFabric  = Invoke-Command -Computer $($Script:Environment.MachineName) -ScriptBlock { 
        Connect-ServiceFabricCluster -TimeoutSec 1 | Out-Null
        Get-ServiceFabricApplication | Where { $_.ApplicationName -like '*AX*' }
    } -ErrorAction SilentlyContinue
    $SecPasswd = ConvertTo-SecureString $(Invoke-ServiceFabricDecryptText -CipherText $(($ServiceFabric.ApplicationParameters | Where { $_.Name -eq 'DataAccess_SqlPwd' }).Value)) -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ($(Invoke-ServiceFabricDecryptText -CipherText $(($ServiceFabric.ApplicationParameters | Where { $_.Name -eq 'DataAccess_SqlUser' }).Value)), $SecPasswd)
    
    if($Credential.UserName -ne $null) {
        if(($Credential.UserName).Split('\').Count -eq 1) {
            $Id = ($Credential.UserName).toUpper()
            $UserName = $Credential.UserName
        }
        else {
            $BSTRBC = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecPasswd)
            $Root = "LDAP://" + ([ADSI]"").DistinguishedName
            $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$Credential.UserName,[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTRBC))
            $Id = (($Credential.UserName).Split('\')[1]).toUpper()
            $UserName = "$(($Domain.dc).ToUpper())\$($Credential.UserName)"
        }
        $SqlQuery = "SELECT Id FROM [dbo].[AXTools_UserAccount] WHERE [Id] = '$Id'"
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$(Get-ConnectionString))
        $IdExists = $SqlCommand.ExecuteScalar()
        if([string]::IsNullOrEmpty($IdExists)) {
            $SecureStringAsPlainText = Write-EncryptedString -InputString $Credential.GetNetworkCredential().Password -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)"
            $SqlQuery = "INSERT INTO [dbo].[AXTools_UserAccount] ([ID],[USERNAME],[PASSWORD]) VALUES ('$Id','$UserName','$SecureStringAsPlainText')"
            $SqlCommand = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$(Get-ConnectionString))
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
	    $SqlQuery = “UPDATE [AXTools_Environments] SET [DBUSER] = '$Id' WHERE [ENVIRONMENT] = '$($Script:Environment.Name)'" 
	    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$(Get-ConnectionString))
	    $SqlCommand.ExecuteNonQuery() #| Out-Null
    }
}

function Deploy-DBLoginSF
{
    if(![string]::IsNullOrEmpty($Script:Environment.keyDBUser)) {
        $SqlServer = New-Object Microsoft.SqlServer.Management.Smo.Server($Script:Environment.keyDbServer)
        if($SqlServer.Logins[$Script:Environment.keyDBUser]) {
            if(!$SqlServer.Databases[$Script:Environment.keyDBName].Users[$Script:Environment.keyDBUser]) {
                $User = New-Object Microsoft.SqlServer.Management.Smo.User($SqlServer.Databases[$Script:Environment.keyDBName],$Script:Environment.keyDBUser)
                $User.Login = $Script:Environment.keyDBUser
                $User.Create()
                $Role = $SqlServer.Databases[$Script:Environment.keyDBName].Roles['db_owner']
                $Role.AddMember($Script:Environment.keyDBUser)
            }
        }
        else {
            $WarningMsg = "SQL Login ($($Script:Environment.keyDBUser)) does not exist."
        }
    }
}

function Clear-EnvironmentData
{
    Clear-Variable -Name Environment -Scope Script -ErrorAction SilentlyContinue
    $Script:Environment = New-Object -TypeName System.Object
    $Script:Environment | Add-Member -Name Name -Value $null -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name HasEnvCfg -Value $false -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name HasServers -Value $false -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name HasStore -Value $false -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name RFROk -Value $false -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name AsJob -Value $false -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name keyDbName -Value $null -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name keyDbServer -Value $null -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name keyDbUser -Value $null -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name MachineName -Value $null -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name MachineFullName -Value $null -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name MachineInstance -Value $null -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name EmailProfile -Value $null -MemberType NoteProperty -Force
    $Script:Environment | Add-Member -Name SQLBackup -Value $null -MemberType NoteProperty -Force
}

function Get-BackupManager
{
[CmdletBinding()]
param(
	[switch]$Backup,
	[switch]$Restore,
	[switch]$Check
)
	if ($Backup) {
		Get-BackupManager -Check
		if ($Script:Environment.HasStore) {
			do {
				Write-Host ''
				if ($Script:Environment.AsJob) { $Prompt = 'Y' } else { $Prompt = Read-Host "Delete $($Script:Environment.Name) Backups? (Y/N)" }
				switch ($Prompt.ToUpper()) {
					Y {
						Get-StoreManager -Delete
						Write-Host ''
						Write-Host "Exporting $($Script:Environment.Name) to Env. Store." -Fore Green
						$SrcServer = $($Script:Environment.keyDbServer)
						$SrcDatabase = $($Script:Environment.keyDBName)
						$DestServer = $($Script:Settings.DBServer)
						$DestDatabase = $($Script:Settings.DBName)
						Get-ScriptTable
						Get-BackupManager -Check
					}
					N {
						$Script:WarningMsg = "Canceled."
						Get-Menu
					}
				}
			} while ($Prompt.ToUpper() -notmatch "[YN]")
		}
		else {
			Write-Host ''
			Write-Host "Exporting $($Script:Environment.Name) to Env. Store." -Fore Green
			$SrcServer = $($Script:Environment.keyDbServer)
			$SrcDatabase = $($Script:Environment.keyDBName)
			$DestServer = $($Script:Settings.DBServer)
			$DestDatabase = $($Script:Settings.DBName)
			Get-ScriptTable
			Get-BackupManager -Check
		}
	}
	if ($Restore) {
		Write-Host ''
		Write-Host "Importing $($Script:Environment.Name) from Env. Store." -Fore Green
        $RFRTables = Get-EnvTables
		foreach ($Table in $RFRTables.Tables[0]) {
			$SrcServer = $($Script:Settings.DBServer)
			$SrcDatabase = $($Script:Settings.DBName)
			$DestServer = $($Table.SourceDbServer)
			$DestDatabase = $($Table.SourceDbName)
			$DestTable = $($Table.SourceTable)
			$Table = $($Table.TargetTable)
			Write-Host "- Source Table: $Table --> Target Table: $DestTable" -Fore Yellow
			if($Script:Settings.SqlTruncate) {
				SQL-TruncateTable
			}
			SQL-BulkInsert
		}
        if($Script:Settings.CreateDBLogin) {
            Deploy-DBLoginSF
        }
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$DestServer;Database=$DestDatabase;Integrated Security=True"
		$SqlConn.Open()
		$SqlQuery = "UPDATE BATCHJOB SET STATUS = '0'"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		$SqlCommand.ExecuteScalar() | Out-Null
		$SqlConn.Close()
		Get-BackupManager -Check
	}
	if ($Check) {
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
		$SqlConn.Open()
		$SqlQuery = "SELECT COUNT(1) FROM AXRefresh_EnvironmentStore WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND DELETED = 0"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		[boolean]$Script:Environment.HasStore = $SqlCommand.ExecuteScalar()
		#
		$SqlQuery = "SELECT COUNT(1) FROM AXTools_Servers WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		[boolean]$Script:Environment.HasServers = $SqlCommand.ExecuteScalar()
		#
		$SqlQuery = "SELECT COUNT(1) FROM AXTools_Environments A
                        JOIN AXRefresh_EnvironmentsExt B on A.ENVIRONMENT = B.ENVIRONMENT
                        WHERE A.ENVIRONMENT = '$($Script:Environment.Name)'"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		[boolean]$Script:Environment.HasEnvCfg = $SqlCommand.ExecuteScalar()
		$SqlConn.Close()
	}
}

function Get-ScriptTable
{
	foreach ($Table in $Script:Settings.SaveTables.Split(',')) {
		$DestTable = "RFR_$($Script:Environment.Name)`_$Table"
		$Server = New-Object "Microsoft.SqlServer.Management.Smo.Server" $SrcServer
		$Database = $Server.Databases[$SrcDatabase]
        Write-Host "- Source Table: $Table --> Target Table: $DestTable" -Fore Yellow
		try {
            $TableSet = $Database.Tables[$Table]
		    if ($TableSet.FileGroup) { $TableSet.FileGroup = 'PRIMARY' }
		    $Script = $TableSet.Script().Replace("CREATE TABLE [dbo].[$Table]","CREATE TABLE [dbo].[$DestTable]")
		    SQL-CreateTable $Script
		    if ($Script:Settings.SqlCompression) { SQL-CompressTable -ColumnStore } else { $Script:SqlCompression = 'None' }
		    SQL-BulkInsert
		    Get-StoreManager -Insert
        }
        catch {
            Write-Host $_.Exception.Message
        }
	}
}

function SQL-CreateTable
{
[CmdletBinding()]
param(
	$Script
)
	try {
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$DestServer;Database=$DestDatabase;Integrated Security=True"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($Script,$SqlConn)
		$SqlConn.Open()
		$SqlCommand.ExecuteNonQuery() | Out-Null
	}
	catch [System.Management.Automation.MethodInvocationException]{
		if ($Script:Settings.SqlTruncate) {
			SQL-TruncateTable
		}
	}
	catch {
		Write-Host $_.Exception.Message
	}
	$SqlConn.Close()
}

function SQL-CompressTable
{
[CmdletBinding()]
param(
	[switch]$ColumnStore,
	[switch]$Page,
	[switch]$None
)
	$Server = New-Object "Microsoft.SqlServer.Management.Smo.Server" $DestServer
	if ($Server.Edition -match 'Enterprise Edition') {
		if ($ColumnStore) {
			try {
				$DBDest = $Server.Databases[$DestDatabase]
				$DestTblSet = $DBDest.Tables[$DestTable]
				$DestIdx = New-Object ("Microsoft.SqlServer.Management.Smo.Index") ($DestTblSet,"idx_ClusteredColumnStore")
				$DestIdx.IndexType = "ClusteredColumnStoreIndex"
				$DestTblSet.Indexes.Add($DestIdx)
				$DestTblSet.Alter()
				$Script:SqlCompression = 'ColumnStore'
			}
			catch {
				SQL-CompressTable -Page
			}
		}
		if ($Page) {
			try {
				$DestTblSet.PhysicalPartitions[0].DataCompression = 'Page'
				$DestTblSet.Rebuild()
				$Script:SqlCompression = 'Page'
			}
			catch {
				SQL-CompressTable -None
			}
		}
		if ($None) {
			$Script:SqlCompression = 'None'
		}
	}
}

function SQL-TruncateTable
{
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$DestServer;Database=$DestDatabase;Integrated Security=True"
	$SqlConn.Open()
	try {
	    $SqlTruncate = "TRUNCATE TABLE [$DestTable]"
	    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlTruncate,$SqlConn)
	    $SqlCommand.ExecuteNonQuery() | Out-Null
    }
    catch {
        Write-Host $_.Exception.Message
    }
    $SqlConn.Close()
}

function SQL-BulkInsert
{
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$SrcServer;Database=$SrcDatabase;Integrated Security=True"
	$Query = "SELECT * FROM [$Table]"
	$SqlCommand = New-Object system.Data.SqlClient.SqlCommand ($Query,$SqlConn)
	$SqlConn.Open()
	[System.Data.SqlClient.SqlDataReader]$SqlReader = $SqlCommand.ExecuteReader()
	try {
		$BulkCopy = New-Object Data.SqlClient.SqlBulkCopy ("Server=$DestServer;Database=$DestDatabase;Integrated Security=True",[System.Data.SqlClient.SqlBulkCopyOptions]::KeepIdentity)
		$BulkCopy.BulkCopyTimeout = 0
		$BulkCopy.DestinationTableName = "[$DestTable]"
		$BulkCopy.WriteToServer($SqlReader) | Out-Null
	}
	catch [System.Exception]{
		Write-Host $_.Exception.Message
	}
	$SqlReader.Close()
	$BulkCopy.Close()
	$SqlConn.Close()
	$SqlConn.Dispose()
}

function Get-StoreManager
{
[CmdletBinding()]
param(
	[switch]$Insert,
	[switch]$Delete,
	[switch]$DeleteStore
)
    if($Insert) {
	    if ($Script:Settings.SqlStoreScript) { $InsertScript = $Script } else { $InsertScript = '' }
	    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
	    $SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
	    $SqlConn.Open()
	    $SqlQuery = “INSERT INTO [AXRefresh_EnvironmentStore] (ENVIRONMENT,SOURCEDBSERVER,SOURCEDBNAME,SOURCETABLE,TARGETTABLE,SQLSCRIPT,COUNT,CREATEDDATETIME,SQLCOMPRESSION,DELETED,DELETEDDATETIME)
                     VALUES('$($Script:Environment.Name)','$SrcServer','$SrcDatabase','$Table','$DestTable','$InsertScript','$($TableSet.RowCount)','$DateTime','$Script:SqlCompression', 0, 0)"
	    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	    $SqlCommand.ExecuteNonQuery() | Out-Null
	    $SqlConn.Close()
    }
    if($Delete -or $DeleteStore) {
	    if ($Script:Environment.Name) {
		    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
		    $SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
		    $SqlConn.Open()
		    $SqlQuery = "SELECT ENVIRONMENT, SOURCETABLE, TARGETTABLE FROM [AXRefresh_EnvironmentStore] WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND DELETED = 0"
		    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
		    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
		    $Adapter.SelectCommand = $SqlCommand
		    $EnvTables = New-Object System.Data.DataSet
		    $Adapter.Fill($EnvTables) | Out-Null
		    if ($Script:Environment.AsJob) { $Prompt = 'Y' } else { $Prompt = Read-Host "Confirm Delete $($($Script:Environment.Name))? (Y/N)" }
		    switch ($Prompt.ToUpper()) {
			    'Y' {
				    Write-Host ''
				    Write-Host "Deleting $($Script:Environment.Name) from Env. Store." -Fore Green
				    foreach ($Table in $EnvTables.Tables[0]) {
					    $SqlQuery = “DROP TABLE [$($Table.TARGETTABLE)]"
					    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					    $SqlCommand.ExecuteNonQuery() | Out-Null
				    }
				    $SqlQuery = "UPDATE [AXRefresh_EnvironmentStore] SET DELETED = 1, DELETEDDATETIME = '$DateTime' WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
				    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
				    $SqlCommand.ExecuteNonQuery() | Out-Null
				    #
				    $SqlQuery = “DELETE FROM [AXTools_Servers] WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND CREATEDDATETIME < '$DateTime'"
				    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
				    $SqlCommand.ExecuteNonQuery() | Out-Null
				    #
				    if ($DeleteStore) {
					    $SqlQuery = “DELETE FROM [AXTools_Environments] WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
					    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					    $SqlCommand.ExecuteNonQuery() | Out-Null
                        #
					    $SqlQuery = “DELETE FROM [AXRefresh_EnvironmentsExt] WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
					    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					    $SqlCommand.ExecuteNonQuery() | Out-Null
					    #
					    $SqlQuery = “DELETE FROM [AXTools_Servers] WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
					    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					    $SqlCommand.ExecuteNonQuery() | Out-Null
					    #
					    $SqlQuery = “DELETE FROM [AXRefresh_EnvironmentStore] WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
					    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
					    $SqlCommand.ExecuteNonQuery() | Out-Null
					    Clear-EnvironmentData
				    }
			    }
			    'N' {
				    $Script:WarningMsg = "Canceled."
                    Get-Menu
			    }
			    Default {
				    $Script:WarningMsg = "Invalid Option. Retry."
			    }
		    }
		    $SqlConn.Close()
	    }
	    else {
		    $Script:WarningMsg = "Canceled."
	    }
    }
}

function SQL-DBRestore
{
param(
	[string]$backupPath
)
	Write-Host ''
	if ($Script:Environment.AsJob) {
		$backupFilePath = $backupPath
	}
	else {
		do {
			$backupFilePath = Read-Host "SQL Backup File (Fullpath)"
		} while (-not ($backupFilePath))
	}
	if (Test-Path $backupFilePath -Include *.bak) {
		Write-Host "Restore $($Script:Environment.keyDbServer)\$($Script:Environment.keyDBName)" -Fore Green
		Write-Host "Backup File $backupFilePath" -Fore Green
		$SqlQuery = "ALTER DATABASE [$($Script:Environment.keyDBName)] " +
		                "SET SINGLE_USER WITH ROLLBACK IMMEDIATE " +
		                "RESTORE DATABASE [$($Script:Environment.keyDBName)] " +
		                "FROM DISK = '$backupFilePath' " +
		                "WITH NOUNLOAD, REPLACE, STATS = 10 "
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$($Script:Environment.keyDbServer);Database=Master;Integrated Security=True"
		$SqlConn.Open()
		$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
		$SqlCommand.CommandTimeout = 0
		try {
			Write-Host '- Restore in progress... ' -Fore Yellow -NoNewline
			$SqlCommand.ExecuteNonQuery() | Out-Null
			Write-Host 'Done.' -Fore Yellow
		}
		catch {
			$Script:WarningMsg = "It was not possible to restore! Try to restore through SSMS"
			$SqlCommand.CommandText = "ALTER DATABASE [$($Script:Environment.keyDBName)] SET MULTI_USER"
			$SqlCommand.ExecuteNonQuery() | Out-Null
			$SqlConn.Close()
			Write-Log $_.Exception.Message "Error"
			Get-Menu
		}
		$SqlConn.Close()
	}
	else {
		Write-Warning 'Incorrect path or file does not exist (*.bak).'
		SQL-DBRestore
	}
}

function SQL-CleanUpTable
{
	Write-Host ''
	Write-Warning 'Before deleting confirm you have exported all the data. DO NOT USE IN PRODUCTION.'
	Write-Host 'Cleaning Tables: ' -Fore Green
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Environment.keyDbServer);Database=$($Script:Environment.keyDBName);Integrated Security=True"
	$SqlConn.Open()

    $CurrEnvTables = Get-EnvTables
    [Array]$TruncateAll = $CurrEnvTables.Tables.SourceTable
    [Array]$TruncateAll += $Script:Settings.DeleteTables.Split(',')

	foreach ($Table in $TruncateAll) {
		Write-Host "- $Table" -Fore Yellow
	}
	do {
		if ($Script:Environment.AsJob) { $Prompt = 'Y' } else { $Prompt = Read-Host "Truncate Tables? (Y/N)" }
	} while (($Prompt -notlike "[YN]"))
	switch ($Prompt.ToUpper()) {
		'Y' {
			foreach ($Table in $TruncateAll)
			{
                ##ADD SAME LOGIC TO EXPORT IF AN ISSUE HAPPENS
                try {
				    $SqlQuery = "TRUNCATE TABLE [$Table]"
				    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
				    $SqlCommand.ExecuteNonQuery() | Out-Null
                }
                catch {
                    Write-Host $_.Exception.Message
                }
			}
			if ($Script:Settings.DataScrub) { Get-DataScrub }
		}
		'N' {
			$Script:WarningMsg = 'Canceled.'
		}
		Default {
			$Script:WarningMsg = 'Invalid Option. Retry.'
		}
	}
	$SqlConn.Close()
}

function Get-DataScrub
{
	if ($Script:Settings.ScrubTables)
	{
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$($Script:Environment.keyDbServer);Database=$($Script:Environment.keyDBName);Integrated Security=True"
		$SqlConn.Open()
		foreach ($Update in $Script:Settings.ScrubTables.Split(','))
		{
            try {
			    $TableName = "[$($Update.Split('|')[0])]"
			    $FieldName = "[$($Update.Split('|')[1])]"
			    $Value = if ($Update.Split('|')[2] -like 'NULL') { "''" } else { "'$($Update.Split('|')[2])'" }
			    $SqlQuery = "UPDATE $TableName SET $FieldName = $Value"
			    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
			    $SqlCommand.ExecuteScalar() | Out-Null
            }
            catch {
                Write-Host $_.Exception.Message
            }
		}
	}
	$SqlConn.Close()
}

function Get-AOSManager
{
[CmdletBinding()]
param(
	[switch]$Stop,
	[switch]$Start,
	[switch]$Enable,
	[switch]$Disable,
	[switch]$Restart,
	[switch]$Status
)
	$AOSServers = @()
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
    $SqlQuery = "SELECT InstanceName, ServerName, ServerType, Version FROM AXTools_Servers WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND INSTANCENAME <> '' AND ACTIVE = 1"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $SqlCommand
    $DSServers = New-Object System.Data.DataSet
    $AOSCnt = $Adapter.Fill($DSServers)
    if ($AOSCnt -gt 0) {
        $ServiceList = @()
        foreach ($AOS in $DSServers.Tables[0]) {
            $ServiceTemp = New-Object -TypeName System.Object
            if (Test-Connection $AOS.ServerName -Count 1 -Quiet) {
			    Switch ($AOS.Version) {
                    '5' {
                        $ServiceTemp | Add-Member -Name AOSName -Value ($AOS.InstanceName).Split('@')[1] -MemberType NoteProperty
                        $ServiceTemp | Add-Member -Name InstanceName -Value ($AOS.InstanceName).Split('@')[0] -MemberType NoteProperty
                        $ServiceTemp | Add-Member -Name ServiceName -Value "AOS50`$$($InstanceName)" -MemberType NoteProperty
                        $ServiceTemp | Add-Member -Name Run -Value '1' -MemberType NoteProperty
                    }
                    '6' {
                        $ServiceTemp | Add-Member -Name AOSName -Value ($AOS.InstanceName).Split('@')[1] -MemberType NoteProperty
                        $ServiceTemp | Add-Member -Name InstanceName -Value ($AOS.InstanceName).Split('@')[0] -MemberType NoteProperty
                        $ServiceTemp | Add-Member -Name ServiceName -Value "AOS60`$$($InstanceName)" -MemberType NoteProperty
                        $ServiceTemp | Add-Member -Name Run -Value '1' -MemberType NoteProperty
                    }
                    '7' {
                        if($AOSCnt -gt 1) {
                            if($AOS.InstanceName -notlike 'Azure@*') {
                                $ServiceTemp | Add-Member -Name AOSName -Value $AOS.ServerName -MemberType NoteProperty
                                $ServiceTemp | Add-Member -Name ServiceName -Value "FabricHostSvc" -MemberType NoteProperty
                                $ServiceTemp | Add-Member -Name Run -Value '1' -MemberType NoteProperty
                            }
                        }
                        else{
                            $OneBoxServ = 'DynamicsAxBatch','Microsoft.Dynamics.AX.Framework.Tools.DMF.SSISHelperService.exe','W3SVC','MR2012ProcessService'
                            $OneBoxServ | % {
                                $ServiceTemp | Add-Member -Name AOSName -Value $AOS.ServerName -MemberType NoteProperty
                                $ServiceTemp | Add-Member -Name ServiceName -Value $_ -MemberType NoteProperty
                                $ServiceTemp | Add-Member -Name Run -Value '1' -MemberType NoteProperty
                                $ServiceList += $ServiceTemp
                                $ServiceTemp = New-Object -TypeName System.Object
                            }
                        }
                    }
                    Default {
                        $ServiceTemp | Add-Member -Name AOSName -Value $AOS.ServerName -MemberType NoteProperty
                        $ServiceTemp | Add-Member -Name ServiceName -Value "" -MemberType NoteProperty
                        $ServiceTemp | Add-Member -Name Run -Value '0' -MemberType NoteProperty
                    }
                }
            }
            else {
                ## Couldn't connect to server.
                $ServiceTemp | Add-Member -Name AOSName -Value $AOS.ServerName -MemberType NoteProperty
                $ServiceTemp | Add-Member -Name ServiceName -Value "" -MemberType NoteProperty
                $ServiceTemp | Add-Member -Name Run -Value '0' -MemberType NoteProperty
            }
            if(![string]::IsNullOrEmpty($ServiceTemp)) {
                $ServiceList += $ServiceTemp
            }
        }
    }
	else {
		$Script:WarningMsg = 'Environment servers not found.'
		break
	}
    $i=0
    foreach($Srv in $ServiceList | Where { $_.Run -eq '1' }) {
	    if ($Stop) {
		    try {
			    $i++
			    Write-Host "$i. Stopping $($Srv.AOSName)... " -Fore Yellow -NoNewline
			    if ($Disable) { Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName) | Set-Service -StartupType Disabled }
		        (Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName)).Stop()
			    Write-Host 'Done.' -Fore Yellow
		    }
		    catch [Exception]{
			    #$Script:WarningMsg = "Nothing to Stop."
                Write-Host "Couldn't Stop - Service is " -NoNewline -Fore Yellow
                Write-Host $(Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName) | Select-Object Status -ExpandProperty Status) -Fore Yellow
		    }
		    catch {
			    $Script:WarningMsg = $_.Exception.Message
		    }
	    }
	    if ($Start) {
		    try {
				$i++
				Write-Host "$i. Starting $($Srv.AOSName)... " -Fore Yellow -NoNewline
				if ($Enable) { Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName) | Set-Service -StartupType Automatic }
			    (Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName)).Start()
				Start-Sleep -s 3
				Write-Host 'Done.' -Fore Yellow
		    }
		    catch [Exception]{
			    #$Script:WarningMsg = "Nothing to Start."
                Write-Host "Couldn't Start - Service is " -NoNewline -Fore Yellow
                Write-Host $(Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName) | Select-Object Status -ExpandProperty Status) -Fore Yellow
		    }
		    catch {
			    $Script:WarningMsg = $_.Exception.Message
		    }
	    }
	    if ($Restart) {
		    try {
				$i++
				Write-Host "$i. Restarting $($Srv.AOSName)... " -Fore Yellow -NoNewline
			    if ($Disable) { Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName) | Set-Service -StartupType Disabled }
		        (Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName)).Stop()
				Start-Sleep -s 5
				if ($Enable) { Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName) | Set-Service -StartupType Automatic }
			    (Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName)).Start()
				Write-Host 'Done.' -Fore Yellow
		    }
		    catch [Exception]{
			    $SrvStatus = Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName) | Select-Object Status -ExpandProperty Status
			    Write-Host "$AOSName is $SrvStatus... " -Fore Yellow -NoNewline
			    if ($SrvStatus -match 'Running') {
				    Write-Host "trying to Stop it." -Fore Yellow
				    (Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName)).Stop()
			    }
			    else {
				    Write-Host "trying to Start it." -Fore Yellow
				    (Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName)).Start()
			    }
		    }
		    catch {
			    $Script:WarningMsg = $_.Exception.Message
		    }
	    }
	    if ($Status) {
		    try {
				$i++
				Write-Host "$i. Server $($Srv.AOSName) is " -Fore Yellow -NoNewline
				Write-Host $(Get-Service -Name $($Srv.ServiceName) -ComputerName $($Srv.AOSName) | Select-Object Status -ExpandProperty Status) -Fore Yellow
		    }
		    catch [Exception]{
			    $Script:WarningMsg = "Nothing to Check."
		    }
		    catch {
			    $Script:WarningMsg = $_.Exception.Message
		    }
        }
    }
    <#
    if (($AOSRFRCnt -eq 0) -and ($AOSCnt -eq 0)) {
	    Write-Host ''
	    $ReadSrvs = Read-Host "Reload Servers (9 > 5) or Type AOS Server(s) [comma-separated]"
	    if (!($ReadSrvs)) {
		    $Script:WarningMsg = 'Invalid Option. Retry.'
		    Get-Menu
	    }
	    foreach ($Srv in $ReadSrvs.Split(',').Trim() | Select-Object -Unique) {
		    if (Test-Connection $Srv -Count 1 -Quiet) {
			    $AOSServers += $Srv.ToUpper()
		    }
		    else {
			    Write-Host ''
			    Write-Warning "$($Srv.ToUpper()) Server is unreachable."
		    }
	    }
    }
    #>
}

function Set-NewAXGuid
{
	Write-Host ''
	Write-Warning 'DO NOT USE IN PRODUCTION.'
	Write-Host 'Updating GUID to 00000000-0000-0000-0000-000000000000.' -Fore Yellow
	do {
		$Prompt = Read-Host "Confirm Update? (Y/N)"
	} while ($Prompt -notlike "[YN]")
	switch ($Prompt) {
		'Y' {
			try {
				$SqlConn = New-Object System.Data.SqlClient.SqlConnection
				$SqlConn.ConnectionString = "Server=$($Script:Environment.keyDbServer);Database=$($Script:Environment.keyDBName);Integrated Security=True"
				$SqlConn.Open()
				$SqlQuery = "UPDATE SYSSQMSETTINGS SET GLOBALGUID = '{00000000-0000-0000-0000-000000000000}'"
				$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
				$SqlCommand.ExecuteNonQuery() | Out-Null
				$SqlConn.Close()
				$Script:WarningMsg = 'Restart AOS Servers.'
			}
			catch {
				$Script:WarningMsg = $_.Exception.Message
			}
		}
		'N' {
			$Script:WarningMsg = 'Canceled.'
		}
	}
}

function Set-TableRecId
{
	Write-Host ''
	Write-Host 'Checking Table RecID' -Fore Green
    $RFRTables = Get-EnvTables
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Environment.keyDbServer);Database=$($Script:Environment.keyDBName);Integrated Security=True"
	$SqlConn.Open()
	foreach ($Table in $RFRTables.Tables[0].SourceTable) {
		Write-Host "- Table $Table... " -Fore Yellow -NoNewline
		$SqlQuery = "SELECT MAX(RECID) FROM $Table"
		$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
		$TableRecId = $SqlCommand.ExecuteScalar()
		$SqlQuery = "SELECT NEXTVAL FROM SYSTEMSEQUENCES WHERE TABID IN (SELECT TABLEID FROM SQLDICTIONARY WHERE NAME = '$Table' AND FIELDID = 0)"
		$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
		$SystemRecId = $SqlCommand.ExecuteScalar()
		if([string]::IsNullOrEmpty($SystemRecId)) { $SystemRecId = 0 }
		if([string]::IsNullOrEmpty($TableRecId)) { $TableRecId = 0 }
		if(($TableRecId -gt $SystemRecId) -and ($TableRecId -ne 0) -and ($SystemRecId -ne 0)) {
			try {
				$TableRecId++
				$SqlQuery = "UPDATE SYSTEMSEQUENCES SET NEXTVAL = $TableRecId WHERE TABID IN (SELECT TABLEID FROM SQLDICTIONARY WHERE NAME = '$Table' AND FIELDID = 0)"
				$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
				$SqlCommand.ExecuteNonQuery() | Out-Null
				$Script:WarningMsg = 'Restart AOS Servers.'
				Write-Host 'Updated.' -Fore Red
			}
			catch {
				Write-Host 'Failed.' -Fore Red
			}
		}
		else {
			Write-Host 'Done.' -Fore Yellow
		}
	}
    $SqlConn.Close()
}

function Get-BatchManager
{
[CmdletBinding()]
param(
	[string]$Action,
    [switch]$DisableJobs,
    [switch]$ChangeServer,
    [switch]$HistoryCleanup
)
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Environment.keyDbServer);Database=$($Script:Environment.keyDBName);Integrated Security=True"
	$SqlConn.Open()
    
    if($DisableJobs) {
	    Write-Host ''
	    do {
		    $Prompt = Read-Host "Delete BatchServerGroup? (Y/N)"
	    } while ($Prompt -notlike "[YN]")
	    switch ($Prompt) {
		    'Y' {
			    Write-Host 'Removing Batch Server Groups...' -Fore Yellow -NoNewline
			    $SqlQuery = 'TRUNCATE TABLE BATCHSERVERGROUP'
			    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
			    $SqlCommand.ExecuteNonQuery() | Out-Null
			    Write-Host " Done." -Fore Yellow
		    }
		    'N' {
			    $Script:WarningMsg = 'Canceled.'
			    break
		    }
	    }
    }
	if($ChangeServer) {
		$SqlQuery = "SELECT SERVERID, ENABLEBATCH FROM SYSSERVERCONFIG ORDER BY ENABLEBATCH DESC, RECID DESC"
		$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCommand
		$BatchServers = New-Object System.Data.DataSet
		$SqlAdapter.Fill($BatchServers) | Out-Null

		$SqlQuery = "SELECT GROUP_ FROM BATCHGROUP"
		$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCommand
		$BatchGroups = New-Object System.Data.DataSet
		$SqlAdapter.Fill($BatchGroups) | Out-Null

		if ($BatchGroups.Tables[0].Rows.Count -eq 0) {
			$SqlQuery = "INSERT INTO BATCHGROUP VALUES('','Empty batch group','2012-05-08 00:03:00.000','61435','-AOS-',0,5637144576)"
			$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
			$SqlCommand.ExecuteNonQuery() | Out-Null
			$BatchGroups.Tables[0].Rows.Add('')
		}
		Write-Host ''
		Write-Host 'Moving Batch Jobs to Batch Server:' -Fore Green
		$i = 0
		foreach ($Server in $BatchServers.Tables[0]) {
			if ($Server.ENABLEBATCH -eq "1")
			{
				$i++
				Write-Host "$i. $(($Server.ServerId).Substring(3,($Server.ServerId).length-3)) " -Fore Yellow -NoNewline
				try {
					$SvcStatus = Get-Service -Name "AOS60`$01" -ComputerName $Server.ServerId.Substring(3) -ErrorAction SilentlyContinue
				}
				catch {
					$Script:WarningMsg = $_.Exception.Message
				}
				if ($SvcStatus.Status –eq "Running”) {
					Write-Host '- Running' -Fore Green
				}
				else {
					Write-Host '- Stopped' -Fore Red
				}
			}
		}
		do {
			$Prompt = Read-Host "Option (1/$i)"
		} while (($Prompt -notlike "[1-$i]") -and ($Prompt))

		if ($Prompt) {
			$ServerId = ($BatchServers.Tables[0])[$Prompt - 1].ServerId
			Write-Host ''
			$SqlQuery = "TRUNCATE TABLE BatchServerGroup"
			$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
			$SqlCommand.ExecuteNonQuery() | Out-Null

			foreach ($Group in $BatchGroups.Tables[0]) {
				Write-Host "- Creating group $($Group.GROUP_) at $($ServerId.Substring(3))" -Fore Yellow
				$SqlQuery = "INSERT INTO BATCHSERVERGROUP (GROUPID, SERVERID, CREATEDDATETIME, CREATEDBY, RECVERSION, RECID) " +
				"VALUES ('$($Group.GROUP_)','$ServerId','$DateTime', 'Admin', 1, $(Get-NextRecId 'BATCHSERVERGROUP'))"
				$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
				$SqlCommand.ExecuteNonQuery() | Out-Null
			}

			$SqlQuery = "SELECT GROUPID, CAPTION, RECID FROM BATCH A WHERE NOT EXISTS (SELECT * FROM BATCHGROUP B WHERE A.GROUPID = B.GROUP_)"
			$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCommand
			$BatchTasks = New-Object System.Data.DataSet
			$TaskCnt = $SqlAdapter.Fill($BatchTasks)

			if ($TaskCnt -gt 0) {
				Write-Host 'Inconsistency on Batch Tasks and Batch Groups.'
				Write-Host 'Batchs:'
				$i = 0
				foreach ($Task in $BatchTasks.Tables[0]) {
					$i++
					Write-Host "$i. $($Task.Caption) - Group $($Task.GroupId)"
				}
				do {
					$Prompt = Read-Host "Would you like to update them to 'Empty Batch Group'? (Y/N)"
				} while ($Prompt.ToUpper() -notmatch "[YN]")

				switch ($Prompt.ToUpper()) {
					'Y' {
						foreach ($UpdTask in $BatchTasks.Tables[0]) {
							$SqlQuery = "UPDATE BATCH SET GROUPID = '' WHERE RECID = '$($UpdTask.RecId)'"
							$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
							$SqlCommand.ExecuteNonQuery() | Out-Null
						}
					}
					'N' {
						$Script:WarningMsg = 'Canceled.'
						continue
					}
				}
			}
			$Script:WarningMsg = 'Restart AOS Servers.'
		}
		else {
			$Script:WarningMsg = 'Canceled.'
		}
	}
    if($HistoryCleanup) {
	    Write-Host ''
	    $defaultValue = Get-Date
	    ($defaultValue,(Read-Host "Delete all Batch History date or [Enter] for [$(Get-Date -Date $defaultValue -Format d)]")) -match '\S' | ForEach-Object { $delDate = $ret = $_ }
	    $DateFormat = 'mm/dd/yyyy' # hh HH mm ss dd yyyy
	    try {
		    [datetime]::TryParseExact(
			    $delDate,
			    $DateFormat,
			    [System.Globalization.DateTimeFormatInfo]::InvariantInfo,
			    [System.Globalization.DateTimeStyles]::None,
			    [ref]$ret) | Out-Null
		    Write-Host ''
		    Write-Host "Delete Batch History tables. Cut-off Date as $delDate" -Fore Green

		    Write-Host '- Deleting BatchConstraintsHistory... ' -Fore Yellow -NoNewline
		    $SqlQuery = "DELETE FROM BATCHCONSTRAINTSHISTORY 
                            WHERE BATCHCONSTRAINTSHISTORY.BATCHID IN (
                                SELECT BATCHID FROM BATCHHISTORY BH 
                                JOIN BATCHJOBHISTORY BJH ON BH.BATCHJOBHISTORYID = BJH.RECID 
                                WHERE BJH.STATUS IN (3,4,8) AND DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), BJH.CREATEDDATETIME) <= '$delDate'
                                )"
		    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
		    $SqlCommand.CommandTimeout = 0
		    $SqlCommand.ExecuteNonQuery() | Out-Null
		    Write-Host 'Done.' -Fore Yellow

		    Write-Host '- Deleting BatchHistory... ' -Fore Yellow -NoNewline
		    $SqlQuery = "DELETE FROM BATCHHISTORY
                            WHERE BATCHHISTORY.BATCHJOBHISTORYID IN (
                                SELECT RECID FROM BATCHJOBHISTORY BJH 
                                WHERE BJH.STATUS IN (3,4,8) AND DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), BJH.CREATEDDATETIME) <= '$delDate'
                                )"
		    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
		    $SqlCommand.CommandTimeout = 0
		    $SqlCommand.ExecuteNonQuery() | Out-Null
		    Write-Host 'Done.' -Fore Yellow

		    Write-Host '- Deleting BatchJobHistory... ' -Fore Yellow -NoNewline
		    $SqlQuery = "DELETE FROM BATCHJOBHISTORY 
                            WHERE BATCHJOBHISTORY.STATUS IN (3,4,8) AND DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), BATCHJOBHISTORY.CREATEDDATETIME) <= '$delDate'"
		    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
		    $SqlCommand.CommandTimeout = 0
		    $SqlCommand.ExecuteNonQuery() | Out-Null
		    Write-Host 'Done.' -Fore Yellow
	    }
	    catch {
		    $Script:WarningMsg = "Invalid Date --> $($_.Exception.Message)"
	    }
    }
	$SqlConn.Close()
}

function Get-EnvTables
{
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
	$SqlConn.Open()
	$SqlQuery = "SELECT ENVIRONMENT, SOURCEDBSERVER, SOURCEDBNAME, SOURCETABLE, TARGETTABLE, COUNT, CREATEDDATETIME FROM AXRefresh_EnvironmentStore WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND DELETED = 0"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$Adapter.SelectCommand = $SqlCommand
	$RFRTables = New-Object System.Data.DataSet
	$Adapter.Fill($RFRTables) | Out-Null
	$SqlConn.Close()
    return $RFRTables
}

function Get-NextRecId
{
[CmdletBinding()]
param(
	[string]$TableName
)
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Environment.keyDbServer);Database=$($Script:Environment.keyDBName);Integrated Security=True"
	$SqlConn.Open()
	$SqlQuery = "SELECT TABLEID FROM SQLDICTIONARY WHERE NAME = '$TableName' AND FIELDID = 0"
	$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
	$AxTableID = $SqlCommand.ExecuteScalar()
	$SqlQuery = "DECLARE @recId bigint; EXEC [dbo].[sp_GetNextRecId] $AxTableID, @recId = @recId OUTPUT; SELECT  @recId"
	$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
	$NextRecID = $SqlCommand.ExecuteScalar()
	$SqlConn.Close()

	return $NextRecID
}

function Start-AsJob
{
	$Script:Environment.AsJob = $true
	Import-Environment

	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
	$SqlConn.Open()
	$SqlQuery = “SELECT TOP 1 CREATEDDATETIME FROM AXRefresh_EnvironmentStore WHERE ENVIRONMENT = '$($Script:Environment.Name)' ORDER BY CREATEDDATETIME DESC"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$StoreDate = $SqlCommand.ExecuteScalar()
	$SqlConn.Close()

	if ($StoreDate -like '') {
		Write-Log 'Enviroment Configuration does not exist.' 'Error'
		exit
	}
	elseif ($RestoreDB) {
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
		$SqlConn.Open()
		$SqlQuery = “SELECT [BKPFOLDER] FROM [AXRefresh_EnvironmentsExt] WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		$SQLBackup = $SqlCommand.ExecuteScalar()
		$SqlConn.Close()
		if ($SQLBckPath -notlike '') {
			if (Test-Path $SQLBckPath) {
				$Script:Environment.SQLBackup = Get-ChildItem -Path $SQLBckPath | Select-Object -First 1 | Where-Object { $_.Extension -match '.bak' } | Sort-Object -Property CreationTime -Descending
				Get-AOSManager -Stop
				SQL-DBRestore $Script:Environment.SQLBackup.FullName
				SQL-CleanUpTable
				Get-BackupManager -Restore
				Get-AOSManager -Start
				$Script:Environment.RFROk = $true
			}
			else {
				Write-Log 'Incorrect Path.' 'Error'
			}
		}
	}
	elseif ($RefreshOnly) {
		Get-AOSManager -Stop
		Start-Sleep -Seconds 5
		SQL-CleanUpTable
		Get-BackupManager -Restore
		Get-AOSManager -Start
	}
	elseif ($RefreshDays -ge 1) {
		if (($StoreDate - $((Get-Date).AddDays($RefreshDays * -1))).Days -eq 0) {
			Get-BackupManager -Backup
		}
		else {
			Write-Log "Environment Date: $StoreDate" 'Warn'
		}
	}
	else {
		Write-Log 'Incorrect Parameters.' 'Error'
	}
}

function Test-SQLSettings
{
[CmdletBinding()]
param(
	[string]$ServerName,
	[string]$DBName
)
	try {
		$Server = New-Object "Microsoft.SqlServer.Management.Smo.Server" "$ServerName"
		if (!($Server.Databases.Name.ToUpper().Contains($DBName.ToUpper()))) {
			Write-Warning 'Database does not exist.'
			#exit
		}
	}
	catch {
		Write-Warning "Failed to connect to $ServerName or User doesn't have access."
		exit
	}

	try {
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$ServerName;Database=$DBName;Integrated Security=True"
		$SqlConn.Open()
		$SqlConn.Close()
		Get-BackupManager -Check
	}
	catch {
		$ErrMsg = $_.Exception
		while ($ErrMsg.InnerException) {
			$ErrMsg = $ErrMsg.InnerException
			if (($ErrMsg.Message).Contains('Login failed')) {
				$Script:WarningMsg = "Login failed to $DBName."
				$SqlFailed = $true
				break
			}
		}
	}
}

function Test-Servers
{
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
	$SqlQuery = "SELECT ServerName FROM AXTools_Servers WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$Adapter.SelectCommand = $SqlCommand
	$DSServers = New-Object System.Data.DataSet
	$AOSCnt = $Adapter.Fill($DSServers)
	if ($AOSCnt) {
		$AOSServers = @()
		$AOSServers = $DSServers.Tables[0] | Select-Object ServerName -ExpandProperty ServerName -Unique
	}
	if ($AOSServers) {
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
		$SqlConn.Open()
		foreach ($Server in $AOSServers) {
			$ServerIp = Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue
			$Computer = Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges -ComputerName $Server -ErrorAction SilentlyContinue
			if ($Computer.Domain) { $FQDN = $Server + "." + $Computer.Domain } else { $FQDN = '' }
			$Prompt = Read-Host "Would you like to enable $Server $($Computer.Domain)?"
			if ($Prompt -like 'Y') {
				$Active = '1'
			}
			else {
				$Active = '0'
			}
			$SqlQuery = "UPDATE [AXTools_Servers] SET ACTIVE = '$Active', IP = '$(($ServerIp.IPV4Address).IPAddressToString)', DOMAIN = '$($Computer.Domain)', FQDN = '$($FQDN)' WHERE ENVIRONMENT = '$($Script:Environment.Name)' AND ServerName = '$Server'"
			$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
			$SqlCommand.ExecuteNonQuery() | Out-Null
		}
		$SqlConn.Close()
	}
}

function Set-SQLBKPFolder
{
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
	$SqlConn.Open()
	$SqlQuery = “SELECT [BKPFOLDER] FROM [AXRefresh_EnvironmentsExt] WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$SQLBackup = $SqlCommand.ExecuteScalar()
	$SqlConn.Close()
	if ($SQLBackup -notlike '') {
		Write-Host ' '
		Write-Host "Current Folder set to $SQLBackup"
	}
	else {
		Write-Host ' '
		Write-Host "Backup folder not set."
	}
	Write-Host ' '
	$NewBkpFolder = Read-Host "Please, enter the new folder path"

	if ($NewBkpFolder -and (Test-Path $NewBkpFolder)) {
		$SqlConn = New-Object System.Data.SqlClient.SqlConnection
		$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
		$SqlConn.Open()
		$SqlQuery = “UPDATE [AXRefresh_EnvironmentsExt] SET [BKPFOLDER] = '$NewBkpFolder' WHERE ENVIRONMENT = '$($Script:Environment.Name)'"
		$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
		$SQLBackup = $SqlCommand.ExecuteScalar()
		$SqlConn.Close()
	}
	else {
		$Script:WarningMsg = 'Invalid Folder Path.'
	}
}

function Get-SendEmail
{
    $Subject = "$(if($Script:Environment.RFROk){"SUCCESS "} else {"FAILED "})Environment $($Script:Environment.Name) has been refreshed on server $($Script:Environment.MachineFullName)."
	$Body = "Script executed by $env:userdomain\$env:username on $env:ComputerName. Parameters: $Paramlist. $(if($Script:Environment.SQLBackup) {"SQL Backup: $($Script:Environment.SQLBackup.FullName) from $($Script:Environment.SQLBackup.CreationTime)."})"
    Send-Email -Subject $Subject -Body $Body -EmailProfile $Script:Environment.EmailProfile -GUID $Script:Settings.Guid
}

function Write-Log
{
param(
	[Parameter(Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
	[ValidateNotNullOrEmpty()]
	[string]$Message,
	[Parameter(Mandatory = $false)]
	[ValidateSet("Error","Warn","Info")]
	[string]$Level = "Info"
)
	switch ($Level) {
		'Error' {
			$LogText = "ERROR: $(if($($Script:Environment.Name)) {"($($Script:Environment.Name)) "})$Message"
		}
		'Warn' {
			$LogText = "WARNING: $(if($($Script:Environment.Name)) {"($($Script:Environment.Name)) "})$Message"
		}
		'Info' {
			$LogText = "INFO: $(if($($Script:Environment.Name)) {"($($Script:Environment.Name)) "})$Message"
		}
	}
	$SqlConn = New-Object System.Data.SqlClient.SqlConnection
	$SqlConn.ConnectionString = "Server=$($Script:Settings.DBServer);Database=$($Script:Settings.DBName);Integrated Security=True"
	$SqlConn.Open()
	$SqlQuery = “INSERT INTO [AXTools_ExecutionLog] (Log, Guid) VALUES ('$($LogText.Replace("'","''"))','$($Script:Settings.Guid)')"
	$SqlCommand = New-Object System.Data.SqlClient.SQLCommand ($SqlQuery,$SqlConn)
	$SqlCommand.ExecuteNonQuery() | Out-Null
	$SqlConn.Close()
}

function Start-Refresh
{
	[Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
	Test-SQLSettings $Script:Settings.DBServer $Script:Settings.DBName
	Clear-EnvironmentData
	$DateTime = Get-Date
	Write-Log ("Refresh Script has started - $ScriptPath")
	Write-Log ("User: $env:userdomain\$env:username")
	Write-Log ("Active Connections: $((Get-WmiObject Win32_LoggedOnUser | Select Antecedent -Unique | Where { $_.Antecedent.ToString().Split('"')[1] -like $env:userdomain } | % { “{0}\{1}” -f $_.Antecedent.ToString().Split('"')[1], $_.Antecedent.ToString().Split('"')[3] }) -join ', ')")
	if ($EnvironName)
	{
        $Script:Environment.Name = $EnvironName
		Write-Log ("Script Parameters: $Paramlist")
		Start-AsJob
	}
	else {
		Write-Host ''
		Write-Host 'Environment Refresh Tool'
		Write-Host '════════════════════════════════════════════════════════════'
		Get-MainMenu
	}
}

foreach ($Param in $PSBoundParameters.GetEnumerator()) {
	$Paramlist += ("{0}: {1} | " -f $Param.Key,$Param.Value)
}

if ($Paramlist) {
	$Paramlist = $Paramlist.Substring(0,$Paramlist.Length - 3)
}

Start-Refresh

if($Script:Settings.SendEmail) {
    if($Script:Environment.EmailProfile) { 
        Get-SendEmail
    }
    else {
        Write-Log 'There is no email profile on this enviroment.'
    }
}
Write-Log ("Refresh Script has finished.")