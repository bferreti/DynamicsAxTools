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
Param (    
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
    [String]$Environment,
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
    [Int]$RefreshDays,
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
    [Switch]$RestoreDB,
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
    [Switch]$RefreshOnly
)
Clear-Host 
$ParamDBServer = 'AAFESDEV-1' #Change SQL Server Name. Server\Instance, (local)
$ParamDBName = 'DynamicsAxTools' #Change DB Name (if you have changed during creation)
$ParamSqlTruncate = $true
$ParamSqlCompression = $true
$ParamSqlStoreScript = $false
$ParamDataScrub = $false
$SendEmail = $false
$AXVersion = '2012' #'2012' or '2009'
#
$Script:RFRTables =   @('DMFDATASOURCE','DMFDATASOURCEPROPERTIES','AFMPLUPARAMETERS','AFMPLUERRHANDLINGSETUP','DMFDEFINITIONGROUPENTITY','SYSEMAILPARAMETERS','AFMRETAILPARAMETERSEXT','AIFCHANNEL','AIFPORT','AIFSERVICE','AIFWEBSITES',
			            'BATCHGROUP','BATCHSERVERCONFIG','BATCHSERVERGROUP','BIANALYSISSERVER','BIANALYSISSERVICESDATABASE','BICONFIGURATION','DMFPARAMETERS','DOCUPARAMETERS','DOCUTYPE','EPGLOBALPARAMETERS',
			            'EPWEBSITEPARAMETERS','HCMEMPLOYMENT','HCMPERSONIMAGE','HCMPOSITION','HCMPOSITIONUNIONAGREEMENT','HCMPOSITIONWORKERASSIGNMENT','HCMWORKER','HCMWORKERPRIMARYPOSITION','INVENTLOCATIONDEFAULTLOCATION',
			            'RETAILCDXDATAGROUP','RETAILCDXFILESTORAGEPROVIDER','RETAILCDXSCHEDULEDATAGROUP','RETAILCDXSCHEDULERINTERVAL','RETAILCHANNELPROFILE','RETAILCHANNELPROFILEPROPERTY','RETAILCHANNELTABLE','RETAILCONNAOSPROFILE',
			            'RETAILCONNCONNECTIONPROFILE','RETAILCONNDATABASEPROFILE','RETAILCONNPARAMETERS','RETAILCONNSCHEDULE','RETAILCONNSTORECONNECTPROFILE','RETAILDEVICE','RETAILDEVICETYPES','RETAILHARDWAREPROFILE',
			            'RETAILINTERNALORGANIZATION','RETAILPARAMETERS','RETAILPOSITIONPOSPERMISSION','RETAILPOSPERMISSIONGROUP','RETAILPOSTHEME','RETAILPUBCHANNELATTRIBUTE','RETAILSHAREDBINGPARAMETERS','RETAILSHAREDPARAMETERS',
			            'RETAILTERMINALTABLE','RETAILTHEMEPALLET','RETAILTILLLAYOUT','RETAILTILLLAYOUTSTAFF','RETAILTILLLAYOUTSTORE','RETAILTILLLAYOUTZONE','RETAILTILLLAYOUTZONEREFERENCE','RETAILTRANSACTIONSERVICEPROFILE',
			            'SECURITYUSERROLE','SRSSERVERS','SYSBCPROXYUSERACCOUNT','SYSCLUSTERCONFIG','SYSSERVERCONFIG','SYSUSERINFO','SYSVERSIONCONTROLPARAMETERS','USERGROUPLIST','USERINFO')

$Script:RFRDelTables = @('SYSSERVERSESSIONS','SYSCLIENTSESSIONS','SYSLASTVALUE','SYSEXCEPTIONTABLE','AIFMESSAGELOG')

$Script:RFRUpdates = @('Address,Email,NULL','VendTable,Email,vendor@contoso.com','PurchTable,Email,vendor@contoso.com','CustTable,Email,customer@contoso.com')
#
$ScriptPath = $MyInvocation.MyCommand.Path
$Dir = Split-Path $ScriptPath
$WorkFolder = $Dir + "\WorkFolder"
#
function Get-Menu
{
    Switch($MenuName) {
        'Main'           { Get-MainMenu }
        'Miscellaneous'  { Get-MiscMenu }
        'Batch'          { Get-BatchMenu }
        'Services'       { Get-ServicesMenu }
         Default         { Get-MainMenu } 
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
    if($Script:RFREnviron) {"from ($Script:RFREnviron)"} else {""}
    Write-Host '5. Restore Environment Configuration'
    Write-Host '6. Start AOS Services'
    Write-Host ''  
    Write-Host '9. Additional AX Tools'
    Write-Host ''
    if($Script:MachineName) { Get-MenuConfig }
    Write-Host 'L - Load Environment'
    Write-Host 'Q - Quit'
    if($Script:WarningMsg) { 
        Write-Host ''
        Write-Warning $Script:WarningMsg
        Clear-Variable WarningMsg -Scope Script }
    Write-Host '════════════════════════════════════════════════════════════'
    $Prompt = Read-Host "Option"

    Switch($Prompt.ToUpper().ToString())  {
        X  {
            if($Script:RFREnviron) {
                RFR-VariableWiper
                Get-Menu
            }
            else {
                $Script:WarningMsg = 'Invalid Option. Retry.'
                Get-Menu
            }
        }
        L { 
            if($Script:RFREnviron) {
                RFR-VariableWiper
            }
            Get-RFREnviroments
            Get-Menu
        }
        Q {
            Clear-Host
            RFR-VariableWiper
            Exit
        }
        0 {
            Get-RFREnviroments
            RFR-StoreManager -Backup
            AOS-ServiceManager -Stop 
            SQL-DBRestore
            SQL-CleanUpTable
            RFR-StoreManager -Restore
            AOS-ServiceManager -Start
            Write-Host ''
            Write-Host 'Process completed.' -Fore Green
            Get-MainMenu
        }
        1 {
            Get-RFREnviroments
            RFR-StoreManager -Backup
            Get-Menu
        }
        2 {
            Get-RFREnviroments
            AOS-ServiceManager -Stop
            Get-Menu 
        }
        3 {
            Get-RFREnviroments
            SQL-DBRestore
            Get-Menu
        }
        4 {
            Get-RFREnviroments
            SQL-CleanUpTable
            Get-Menu
        }
        5 {
            Get-RFREnviroments
            RFR-StoreManager -Restore
            Get-Menu
        }
        6 {
            Get-RFREnviroments
            AOS-ServiceManager -Start
            Get-Menu 
        }
        9 {
            Get-MiscMenu
        }
        Default {
            if(($Prompt -notlike "[LQ01234569]") -and !($Script:MachineName)) {
                $Script:WarningMsg = 'Invalid Option. Retry.'
                Get-Menu
            }
            elseif(($Prompt -notlike "[XLQ01234569]") -and ($Script:MachineName)) {
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
    Write-Host '5. Validate Servers'
    Write-Host '6. Change SQL Backup Folder'
    Write-Host '7. Update GUID'
    Write-Host ''
    Write-Host ''
    if($Script:MachineName) { Get-MenuConfig }
    Write-Host 'L - Load Environment'
    Write-Host 'R - Return'
    Write-Host 'Q - Quit'
    if($Script:WarningMsg) { 
        Write-Host ''
        Write-Warning $Script:WarningMsg
        Clear-Variable WarningMsg -Scope Script }
    Write-Host '════════════════════════════════════════════════════════════'
    $Prompt = Read-host "Option" 

    Switch ($Prompt.ToUpper().ToString()) {
        X  {
            if($Script:MachineName) {
                RFR-VariableWiper
                Get-Menu
            }
            else {
                $Script:WarningMsg = 'Invalid Option. Retry.'
                Get-Menu
            }
        }
        L { 
            Get-RFREnviroments
            Get-Menu
        }
        R {
            Get-MainMenu
        }
        Q {
            Clear-Host
            Exit
        }
        1 {
            Get-BatchMenu
        }
        2 {
            Get-RFREnviroments
            RFR-RecIdCheck
            Get-Menu
        }
        3 {
            Get-ServicesMenu
        }
        4 {
            Get-RFREnviroments
            RFR-DeleteStore -HardDelete
            Get-Menu
        }
        5 {
            Get-RFREnviroments
            Validate-Servers
            Get-Menu
        }
        6 {
            Get-RFREnviroments
            Update-SQLBKPFolder
            Get-Menu
        }
        7 {
            Get-RFREnviroments
            RFR-UpdateGUID
            Get-Menu
        }
        Default {
            if(($Prompt -notlike "[LQR123456]") -and !($Script:MachineName)) {
                $Script:WarningMsg = 'Invalid Option. Retry.'
                Get-Menu
            }
            elseif(($Prompt -notlike "[XLQR123456]") -and ($Script:MachineName)) {
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
    if($Script:MachineName) { Get-MenuConfig }
    Write-Host 'L - Load Environment'
    Write-Host 'R - Return'
    Write-Host 'Q - Quit'
    if($Script:WarningMsg) { 
        Write-Host ''
        Write-Warning $Script:WarningMsg
        Clear-Variable WarningMsg -Scope Script }
    Write-Host '════════════════════════════════════════════════════════════'
    $Prompt = Read-host "Option"

    Switch ($Prompt.ToUpper().ToString())  {
        X  {
            if($Script:MachineName) {
                RFR-VariableWiper
                Get-Menu
            }
            else {
                $Script:WarningMsg = 'Invalid Option. Retry.'
                Get-Menu
            }
        }
        L { 
            Get-RFREnviroments
            Get-Menu
        }
        R {
            Get-MiscMenu
        }
        Q {
            Clear-Host
            Exit
        }
        1 {
            Get-RFREnviroments
            RFR-BatchJobs 'DisableJobs'
            Get-Menu
        }
        2 {
            Get-RFREnviroments
            RFR-BatchJobs 'ChangeServer'
            Get-Menu
        }
        3 {
            Get-RFREnviroments
            RFR-BatchHistory
            Get-Menu
        }
        Default {
            if(($Prompt -notlike "[LQR123]") -and !($Script:MachineName)) {
                $Script:WarningMsg = 'Invalid Option. Retry.'
                Get-Menu
            }
            elseif(($Prompt -notlike "[XLQR123]") -and ($Script:MachineName)) {
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
    if($Script:MachineName) { Get-MenuConfig }
    Write-Host 'L - Load Environment'
    Write-Host 'R - Return'
    Write-Host 'Q - Quit'
    if($Script:WarningMsg) { 
        Write-Host ''
        Write-Warning $Script:WarningMsg
        Clear-Variable WarningMsg -Scope Script }
    Write-Host '════════════════════════════════════════════════════════════'
    $Prompt = Read-host "Option"

    Switch ($Prompt.ToUpper().ToString()) {
        X  {
            if($Script:MachineName) {
                RFR-VariableWiper
                Get-Menu
            }
            else {
                $Script:WarningMsg = 'Invalid Option. Retry.'
                Get-Menu
            }
        }
        L { 
            Get-RFREnviroments
            Get-Menu
        }
        R {
            Get-MiscMenu
        }
        Q {
            Clear-Host
            Exit
        }
        1 {
            AOS-ServiceManager -Start
            Get-Menu
        }
        2 {
            AOS-ServiceManager -Stop
            Get-Menu
        }
        3 {
            AOS-ServiceManager -Restart
            Get-Menu
        }
        4 {
            AOS-ServiceManager -Status
            Get-Menu
        }
        Default {
            if(($Prompt -notlike "[LQR1234]") -and !($Script:MachineName)) {
                $Script:WarningMsg = 'Invalid Option. Retry.'
                Get-Menu
            }
            elseif(($Prompt -notlike "[XLQR1234]") -and ($Script:MachineName)) {
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
    Write-Host $Script:RFREnviron -Fore Yellow
    Write-Host 'Machine Name: ' -NoNewline
    Write-Host $Script:MachineName -Fore Yellow
    Write-Host 'SQL Server: ' -NoNewline
    Write-Host $Script:keyDbServer -Fore Yellow
    Write-Host 'AX Database Name: ' -NoNewline
    Write-Host $Script:keyDbName -Fore Yellow
    Write-Host ''
    Write-Host 'X - Release Environment'
}

function Get-EnvConfiguration
{
    $Script:RFREnviron = (Read-Host "Enter Source Environment").ToUpper()
    if($Script:RFREnviron.Length -gt 30) {
        Write-Warning "Environment name must have less than 30 chars."
        RFR-VariableWiper
        Get-EnvConfiguration
    }
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
    $SqlConn.Open()
    $SqlQuery = “SELECT ENVIRONMENT FROM [AX_EnvironmentStore] WHERE ENVIRONMENT = '$Script:RFREnviron' AND DELETED = 0"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
    $EnvCnt = $SqlCommand.ExecuteScalar()
    $SqlConn.Close() | Out-Null
    if($EnvCnt -ge 1) {
        $Prompt = Read-Host "Found Backup for $Script:RFREnviron. Load Configuration (Y/N)"
        Switch($Prompt.ToUpper()) {
            'Y' {
                RFR-LoadEnvironment
            }
            'N' {
                RFR-VariableWiper
                Get-EnvConfiguration
                }
        }
    }
    else {
        Get-AOSConfiguration
    }
}

function Get-AOSConfiguration
{
    Write-Host 'For more than one instance, add the instance number (02@SERVER)' -fore Yellow
    $Script:MachineName = (Read-Host "Enter AOS Server Name").ToUpper()

    if($Script:MachineName) { 
        if($Script:MachineName.Split('@').Count -eq 1) {
            $Script:MachineFullName = "01@$Script:MachineName"
            $Script:MachineInstance = '01'
        }
        elseif($Script:MachineName.Split('@').Count -eq 2) {
            $Script:MachineFullName = $Script:MachineName
            $Script:MachineInstance = $Script:MachineName.Split('@')[0]
            $Script:MachineName = $Script:MachineName.Split('@')[1]
        }
    }
    else {
        $Script:WarningMsg = 'Invalid AOS Server. Retry.'
        RFR-VariableWiper
        Get-Menu
    }
 
    if(!(Test-Connection $Script:MachineName -Count 1 -Quiet)) {
        $Script:WarningMsg = 'Invalid AOS Server. Retry.'
        RFR-VariableWiper
        Get-Menu
    }
    <#else {
        $Script:WarningMsg = 'Invalid. Please retry.'
        RFR-VariableWiper
        Get-Menu
    }#>
    ## Connecting to AOS Server Registry
    try {
        if($AXVersion -eq '2012') {
            $keyVersion = '6.0'
        }
        else {
            $keyVersion = '5.0'
        }
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$Script:MachineName)
        $key = "SYSTEM\\CurrentControlSet\\services\\Dynamics Server\\$keyVersion\\$Script:MachineInstance"
        $regkey = $reg.opensubkey($key)
        $keyCurrent = $regKey.GetValue('Current')
        $key = "SYSTEM\\CurrentControlSet\\services\\Dynamics Server\\$keyVersion\\$Script:MachineInstance\\$keyCurrent"
        $regkey = $reg.opensubkey($key)
        $Script:keyDbServer = $regKey.GetValue('DbServer')
        $Script:keyDbName = $regKey.GetValue('Database')
        Get-RunningServers
        #
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
        $SqlConn.Open()
        $SqlQuery = “INSERT INTO [AX_Environments] (ENVIRONMENT, DESCRIPTION, MACHINENAME, DBSERVER, DBNAME)
                        VALUES('$Script:RFREnviron','$Script:RFREnviron','$Script:MachineName','$Script:keyDbServer','$Script:keyDbName')"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        $SqlCommand.ExecuteNonQuery()  | Out-Null
        $SqlConn.Close()
    }
    catch {
        $Script:WarningMsg = "AOS Server not found."
        RFR-VariableWiper
        Get-Menu
    }
}

function Get-RunningServers
{
    try {
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=$Script:keyDbName;Integrated Security=True"
        $SqlConn.Open()
        $SqlQuery = "SELECT SERVERID, AOSID, INSTANCE_NAME, VERSION, STATUS
		                , TXT_STATUS = CASE STATUS 
			                WHEN 1 THEN 'Running'
			                WHEN 0 THEN 'Stopped'
			                END
                    FROM SYSSERVERSESSIONS"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $SqlCommand
        $DSServers = New-Object System.Data.DataSet
        $Adapter.Fill($DSServers) | Out-Null
        $SqlConn.Close()
    }
    catch {
        $Script:WarningMsg = "Error conneting to SQL Server " + $Script:keyDbServer + " and " + $Script:keyDbName
        Get-Menu
    }
        
    Write-Host ''
    Write-Host 'Environment Configuration:'
    Write-Host 'SQL Server Name: ' -NoNewLine; Write-Host $Script:keyDbServer -Fore Yellow
    Write-Host 'AX Database: ' -NoNewLine; Write-Host $Script:keyDbName -Fore Yellow
    Write-Host ''    Write-Host 'AOS Servers Status:'    $i = 1
    $DSServers.Tables[0] | foreach {
        Write-Host "$i. $($_.AOSID.substring(0,$_.AOSID.Length-5)) - $($_.TXT_STATUS)" -Fore Yellow
        $i++
    }
    Do {
        if($Script:IsJob) {$Prompt = 'Y'} else { $Prompt = Read-host "Continue? (Y/N)" } 
        Switch ($Prompt.ToUpper())  {
            Y {
                $SqlConn = New-Object System.Data.SqlClient.SqlConnection
                $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
                $SqlConn.Open()
                if($Script:HasServers) {
                    $SqlQuery = “DELETE FROM [AX_Servers] WHERE ENVIRONMENT = '$Script:RFREnviron' AND CREATEDDATETIME <= '$DateTime'"
                    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
                    $SqlCommand.ExecuteNonQuery()  | Out-Null
                }
                $DSServers.Tables[0] | foreach {
                    if($Script:MachineFullName -like "$($_.INSTANCE_NAME)@$($_.AOSID.substring(0,$_.AOSID.Length-5))") {
                        $Active = '1'
                    }
                    else {
                        $Active = '0'
                    }
                    $SqlQuery = “INSERT INTO [AX_Servers] (ENVIRONMENT, SERVERNAME, SERVERTYPE, ACTIVE, SERVERID, AOSID, INSTANCE, VERSION, STATUS)
                                    VALUES('$Script:RFREnviron','$($_.AOSID.substring(0,$_.AOSID.Length-5))','AOS','$Active','$($_.ServerID)','$($_.AOSID)','$("$($_.INSTANCE_NAME)`@$($_.AOSID.substring(0,$_.AOSID.Length-5))")','$($_.VERSION)','$($_.STATUS)')"
                    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
                    $SqlCommand.ExecuteNonQuery()  | Out-Null
                }
                $SqlConn.Close()
                $DSServers.Dispose()
            }
            N {
                $DSServers.Dispose()
                $Script:WarningMsg =  "Canceled!"
                RFR-VariableWiper
                Get-Menu
            }
        }
    } While ($Prompt -notmatch "[YN]")
}

function RFR-VariableWiper
{
    $Script:HasEnvCfg = $false
    $Script:HasServers = $false
    $Script:HasStore = $false
    $Script:RFROk = $false
    $ClearVar = 'keyDatabase',
                'keyDbName',
                'keyDbServer',
                'MachineName',
                'RFREnviron',
                'SQLQuery',
                'RFR_Env_Tables',
                'AvailEnviroments',
                'MachineFullName',
                'MachineInstance',
                'SQLBackup'
    foreach($Var in $ClearVar) {
        Clear-Variable -Name $Var -Scope Script -ErrorAction SilentlyContinue
    }
}

function Get-RFREnviroments
{
    if(!($Script:RFREnviron)) {
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
        $SqlQuery = "SELECT ENVIRONMENT as Options FROM AX_Environments WHERE MACHINENAME <> ''"
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $SqlCommand
        $RFREnvironments = New-Object System.Data.DataSet
        $EnvTotal = $Adapter.Fill($RFREnvironments)
        $SqlConn.Close()
        $Options = @()
        [array]$Options = $RFREnvironments.Tables[0] | Select Options -Unique
        $Options += @{ Options = "<< Add New >>" }
        $i = 0
        Write-Host ''
        Write-Host 'Choose an Enviroment:'
        foreach ($Option in $Options)  {
            $i++
            Write-Host "$i. $($Option.Options)"
        }
        Do {
            $Prompt = Read-Host "Option (1/$i)"
        } While (($Prompt -notlike "[1-$i]") -and ($Prompt))
    
        if(!($Prompt)) {
            RFR-VariableWiper
            $Script:WarningMsg =  'Invalid Option. Retry.'
            Get-Menu
        }
        elseif($Options.Count -eq $Prompt) {
            Get-EnvConfiguration
        }
        else {
            $Script:RFREnviron = (($RFREnvironments.Tables[0] | Select Options -Unique)[$Prompt-1]).Options
            RFR-LoadEnvironment
        }
    }
}

function RFR-LoadEnvironment
{
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
    $SqlConn.Open()
    #
    $SqlQuery = “SELECT MACHINENAME FROM AX_Environments WHERE ENVIRONMENT = '$Script:RFREnviron'"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
    $Script:MachineName = $SqlCommand.ExecuteScalar()
    #
    $SqlQuery = “SELECT DBSERVER FROM AX_Environments WHERE ENVIRONMENT = '$Script:RFREnviron'"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
    $Script:keyDbServer = $SqlCommand.ExecuteScalar()
    #
    $SqlQuery = “SELECT DBNAME FROM AX_Environments WHERE ENVIRONMENT = '$Script:RFREnviron'"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
    $Script:keyDbName = $SqlCommand.ExecuteScalar()
    $SqlConn.Close()
    RFR-SQLValidation $Script:keyDbServer $Script:keyDbName
    RFR-StoreManager -Check
    Get-RunningServers
}

function RFR-StoreManager
{
[CmdletBinding()]
param (
    [Switch] $Backup,
    [Switch] $Restore,
    [Switch] $Check
)
    if($Backup) {
        RFR-StoreManager -Check
        if($Script:HasStore) {
            Do {
                Write-Host ''
                if($Script:IsJob) { $Prompt = 'Y' } else { $Prompt = Read-host "Delete $Script:RFREnviron Backups? (Y/N)" }
                Switch ($Prompt.ToUpper()) {
                    Y {
                        RFR-DeleteStore
                        Write-Host ''
                        Write-Host "Exporting $Script:RFREnviron to Env. Store." -Fore Green 
                        $SrcServer = $Script:keyDbServer
                        $SrcDatabase = $Script:keyDbName
                        $DestServer = $ParamDBServer
                        $DestDatabase = $ParamDBName
                        Get-ScriptTable
                        RFR-StoreManager -Check
                    }
                    N {
                        $Script:WarningMsg = "Canceled."
                        Get-Menu
                    }
                }
            } While ($Prompt.ToUpper() -notmatch "[YN]")
        }
        else {
            Write-Host ''
            Write-Host "Exporting $Script:RFREnviron to Env. Store." -Fore Green 
            $SrcServer = $Script:keyDbServer
            $SrcDatabase = $Script:keyDbName
            $DestServer = $ParamDBServer
            $DestDatabase = $ParamDBName
            Get-ScriptTable
            RFR-StoreManager -Check
        }
    }
    if($Restore) {
        Write-Host ''
        Write-Host "Importing $Script:RFREnviron from Env. Store." -Fore Green
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
        $SqlConn.Open()
        $SqlQuery = "SELECT ENVIRONMENT, SOURCEDBSERVER, SOURCEDBNAME, SOURCETABLE, RFRTABLENAME, COUNT, CREATEDDATETIME FROM AX_ENVIRONMENTSTORE WHERE ENVIRONMENT = '$Script:RFREnviron' AND DELETED = 0"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $SqlCommand
        $RFRTables = New-Object System.Data.DataSet
        $Adapter.Fill($RFRTables) | Out-Null
        $SqlConn.Close()
        foreach($Table in $RFRTables.Tables[0]) {
            $SrcServer = $ParamDBServer
            $SrcDatabase = $ParamDBName
            $DestServer = $Table.SourceDbServer
            $DestDatabase = $Table.SourceDbName
            $DestTable = $Table.SourceTable
            $Table = $Table.RFRTableName
            Write-Host "- Source Table: $Table --> Target Table: $DestTable" -Fore Yellow
            If ($ParamSqlTruncate) {
                SQL-TruncateTable
            }
            SQL-BulkInsert
        }
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$DestServer;Database=$DestDatabase;Integrated Security=True"
        $SqlConn.Open()
        $SqlQuery = "UPDATE BATCHJOB SET STATUS = '0'"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        $SqlCommand.ExecuteScalar() | Out-Null
        $SqlConn.Close()
        RFR-StoreManager -Check
    }
    if($Check) {
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
        $SqlConn.Open()
        $SqlQuery = "SELECT COUNT(1) FROM AX_ENVIRONMENTSTORE WHERE ENVIRONMENT = '$Script:RFREnviron' AND DELETED = 0"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        [boolean]$Script:HasStore = $SqlCommand.ExecuteScalar()
        #
        $SqlQuery = "SELECT COUNT(1) FROM AX_SERVERS WHERE ENVIRONMENT = '$Script:RFREnviron'"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        [boolean]$Script:HasServers = $SqlCommand.ExecuteScalar()
        #
        $SqlQuery = "SELECT COUNT(1) FROM AX_ENVIRONMENTS WHERE ENVIRONMENT = '$Script:RFREnviron'"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        [boolean]$Script:HasEnvCfg = $SqlCommand.ExecuteScalar()
        $SqlConn.Close()
    }
}

function Get-ScriptTable
{
    foreach ($Table in $Script:RFRTables)  { 
        $DestTable = "RFR_$Script:RFREnviron`_$Table"
        $Server = New-Object "Microsoft.SqlServer.Management.Smo.Server" $SrcServer
        $Database = $Server.Databases[$SrcDatabase]
        $TableSet = $Database.Tables[$Table]
        if($TableSet.FileGroup) { $TableSet.FileGroup = 'PRIMARY' }
        Write-Host "- Source Table: $Table --> Target Table: $DestTable" -Fore Yellow
        $Script = $TableSet.Script().Replace("CREATE TABLE [dbo].[$Table]","CREATE TABLE [dbo].[$DestTable]")
        SQL-CreateTable $Script
        if($ParamSqlCompression) { SQL-CompressTable -ColumnStore }
        else { $Script:SQLCompress = 'None' }
        SQL-BulkInsert
        RFR-InsertStore
    }
}

function SQL-CreateTable
{
[CmdletBinding()]
param (
	$Script
)
    try {
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$DestServer;Database=$DestDatabase;Integrated Security=True"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($Script, $SqlConn)  
        $SqlConn.Open()
        $SqlCommand.ExecuteNonQuery() | Out-Null
    }
    catch [System.Management.Automation.MethodInvocationException] {
        If ($ParamSqlTruncate) {
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
param (
    [Switch] $ColumnStore,
    [Switch] $Page,
    [Switch] $None
)
    $Server = New-Object "Microsoft.SqlServer.Management.Smo.Server" $DestServer
    if($Server.Edition -match 'Enterprise Edition') {
        if($ColumnStore) {
            try {
                $DBDest = $Server.Databases[$DestDatabase]
                $DestTblSet = $DBDest.Tables[$DestTable]        
                $DestIdx = New-Object ("Microsoft.SqlServer.Management.Smo.Index")($DestTblSet, "idx_ClusteredColumnStore") 
                $DestIdx.IndexType = "ClusteredColumnStoreIndex"
                $DestTblSet.Indexes.Add($DestIdx)
                $DestTblSet.Alter()
                $Script:SqlCompression = 'ColumnStore'
            }
            catch {
                SQL-CompressTable -Page
            }
        }
        if($Page) {
            try {
                $DestTblSet.PhysicalPartitions[0].DataCompression = 'Page'
                $DestTblSet.Rebuild()
                $Script:SqlCompression = 'Page'
            }
            catch {
                SQL-CompressTable -None
            }
        }
        if($None) {
            $Script:SqlCompression = 'None'
        }
    }
}

function SQL-TruncateTable
{
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$DestServer;Database=$DestDatabase;Integrated Security=True"
    $SqlConn.Open()
    $SqlTruncate = "TRUNCATE TABLE [$DestTable]"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlTruncate,$SqlConn)
    $SqlCommand.ExecuteNonQuery() | Out-Null
    $SqlConn.Close()
}

function SQL-BulkInsert
{
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$SrcServer;Database=$SrcDatabase;Integrated Security=True"
    $Query = "SELECT * FROM [$Table]"
    $SqlCommand = New-Object system.Data.SqlClient.SqlCommand($Query, $SqlConn)  
    $SqlConn.Open()
    [System.Data.SqlClient.SqlDataReader] $SqlReader = $SqlCommand.ExecuteReader()
    try {
        $BulkCopy = New-Object Data.SqlClient.SqlBulkCopy("Server=$DestServer;Database=$DestDatabase;Integrated Security=True", [System.Data.SqlClient.SqlBulkCopyOptions]::KeepIdentity)
        $BulkCopy.BulkCopyTimeout = 0
        $BulkCopy.DestinationTableName = "[$DestTable]"
        $BulkCopy.WriteToServer($SqlReader)  | Out-Null
    }
    catch [System.Exception] {
        Write-Host $_.Exception.Message
    }
    $SqlReader.Close()
    $BulkCopy.Close()
    $SqlConn.Close()
    $SqlConn.Dispose()
}

function RFR-InsertStore
{
    if($ParamSqlStoreScript) { $InsertScript = $Script } else { $InsertScript = '' }
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
    $SqlConn.Open()
    $SqlQuery = “INSERT INTO [AX_EnvironmentStore] (ENVIRONMENT,SOURCEDBSERVER,SOURCEDBNAME,SOURCETABLE,RFRTABLENAME,SQLSCRIPT,COUNT,CREATEDDATETIME,SQLCOMPRESSION,DELETED,DELETEDDATETIME)
                 VALUES('$Script:RFREnviron','$SrcServer','$SrcDatabase','$Table','$DestTable','$InsertScript','$($TableSet.RowCount)','$DateTime','$Script:SqlCompression', 0, 0)"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
    $SqlCommand.ExecuteNonQuery() | Out-Null
    $SqlConn.Close()
}

function RFR-DeleteStore
{
[CmdletBinding()]
param (
	[Switch] $HardDelete
)
    if($Script:RFREnviron) {
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
        $SqlConn.Open()
        $SqlQuery = "SELECT ENVIRONMENT, SOURCETABLE, RFRTABLENAME FROM [AX_EnvironmentStore] WHERE ENVIRONMENT = '$Script:RFREnviron' AND DELETED = 0"
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $SqlCommand
        $EnvTables = New-Object System.Data.DataSet
        $Adapter.Fill($EnvTables) | Out-Null
        if($Script:isJob) {$Prompt = 'Y' } else { $Prompt = Read-Host "Confirm Delete $($Script:RFREnviron)? (Y/N)" }
        Switch($Prompt.ToUpper()) {
            'Y' {
                Write-Host ''
                Write-Host "Deleting $Script:RFREnviron from Env. Store." -Fore Green
                foreach($Table in $EnvTables.Tables[0] ) {
                    $SqlQuery = “DROP TABLE [$($Table.RFRTABLENAME)]"
                    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
                    $SqlCommand.ExecuteNonQuery() | Out-Null
                }
                $SqlQuery = "UPDATE [AX_EnvironmentStore] SET DELETED = 1, DELETEDDATETIME = '$DateTime' WHERE ENVIRONMENT = '$Script:RFREnviron'"
                $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
                $SqlCommand.ExecuteNonQuery() | Out-Null
                #
                $SqlQuery = “DELETE FROM [AX_Servers] WHERE ENVIRONMENT = '$Script:RFREnviron' AND CREATEDDATETIME < '$DateTime'"
                $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
                $SqlCommand.ExecuteNonQuery() | Out-Null
                #
                if($HardDelete) {
                    $SqlQuery = “DELETE FROM [AX_Environments] WHERE ENVIRONMENT = '$Script:RFREnviron'"
                    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
                    $SqlCommand.ExecuteNonQuery() | Out-Null
                    #
                    $SqlQuery = “DELETE FROM [AX_Servers] WHERE ENVIRONMENT = '$Script:RFREnviron'"
                    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
                    $SqlCommand.ExecuteNonQuery() | Out-Null
                    #
                    $SqlQuery = “DELETE FROM [AX_EnvironmentStore] WHERE ENVIRONMENT = '$Script:RFREnviron'"
                    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
                    $SqlCommand.ExecuteNonQuery() | Out-Null
                    RFR-VariableWiper
                }
            }
            'N' {
                $Script:WarningMsg = "Canceled."
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

function SQL-DBRestore
{
param (
    [string]$backupPath
)
    Write-Host ''
    if($Script:isJob) {
        $backupFilePath = $backupPath
    }
    else {
        Do {
            $backupFilePath = Read-Host "SQL Backup File (Fullpath)"
        } While (-not($backupFilePath))
    }
    if(Test-Path $backupFilePath -Include *.bak) {
        Write-Host "Restore $Script:keyDbServer\$Script:keyDbName" -Fore Green
        Write-Host "Backup File $backupFilePath" -Fore Green
        $SqlQuery = "ALTER DATABASE [$Script:keyDbName] " +
                     "SET SINGLE_USER WITH ROLLBACK IMMEDIATE " +
                     "RESTORE DATABASE [$Script:keyDbName] " +
                     "FROM DISK = '$backupFilePath' " +
                     "WITH NOUNLOAD, REPLACE, STATS = 10 "
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=Master;Integrated Security=True"
        $SqlConn.Open()
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
        $SqlCommand.CommandTimeout = 0
        try {
            Write-Host '- Restore in progress... ' -Fore Yellow -NoNewline
            $SqlCommand.ExecuteNonQuery() | Out-Null
            Write-Host 'Done.' -Fore Yellow
        }
        catch {
            $Script:WarningMsg = "It was not possible to restore! Try to restore through SSMS"
            $SqlCommand.CommandText="ALTER DATABASE [$Script:keyDbName] SET MULTI_USER"
            $SqlCommand.Executenonquery() | Out-Null
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
    $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=$Script:keyDbName;Integrated Security=True"
    $SqlConn.Open()
    $TruncateAll = $Script:RFRTables + $Script:RFRDelTables
    foreach($Table in $TruncateAll) {
        Write-Host "- $Table" -Fore Yellow
    }
    Do {
        if($Script:IsJob) { $Prompt = 'Y'} else { $Prompt = Read-Host "Truncate Tables? (Y/N)" }
    } While (($Prompt -notlike "[YN]"))
    Switch($Prompt.ToUpper()) {
        'Y' {
                foreach($Table in $TruncateAll)
                {
                    $SqlQuery = "TRUNCATE TABLE [$Table]"
                    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
                    $SqlCommand.ExecuteNonQuery() | Out-Null
                    #$SqlQuery
                }
                if($ParamDataScrub) { RFR-DataScrub }
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

function RFR-DataScrub
{
    if($Script:RFRUpdates)
    {
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=$Script:keyDbName;Integrated Security=True"
        $SqlConn.Open()
        foreach($Update in $Script:RFRUpdates)
        {
            $TableName = "[$($Update.Split(',')[0])]"
            $FieldName = "[$($Update.Split(',')[1])]"
            $Value = if($Update.Split(',')[2] -like 'NULL') { "''" } else { "'$($Update.Split(',')[2])'" }
            $SqlQuery = "UPDATE $TableName SET $FieldName = $Value"
            $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
            $SqlCommand.ExecuteScalar() | Out-Null
            #$SqlQuery 
        }
    }
    $SqlConn.Close()
}

function Csv-ExportFile
{
    Write-Host ''
    Write-Host 'Exporting CSV files... ' -Fore Yellow -NoNewline
    Invoke-SqlCmd -InputFile "$Dir\sp_GenerateInsertScript.sql" -ServerInstance $Script:keyDbServer -Database $Script:keyDbName
    ## Create WorkFolder
    if(Test-Path -Path "$WorkFolder\$Script:RFREnviron") {
        Remove-Item "$WorkFolder\$Script:RFREnviron\*"
    }
    else {
        New-Item -ItemType directory -Path "$WorkFolder\$Script:RFREnviron" | Out-Null
    }
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=$Script:keyDbName;Integrated Security=True"
    $SqlConn.Open()
    foreach($Table in $Script:RFRTables) {
        $SqlQuery = "SELECT COUNT(1) FROM [$Table]"
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
        $TotalCount += $SqlCommand.ExecuteScalar()
    }
    foreach($Table in $Script:RFRTables) {
        $SqlQuery = "sp_GenerateInsertScript $Table"
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
        
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $SqlCommand
        $DSInserts = New-Object System.Data.DataSet
        $i = $Adapter.Fill($DSInserts)
        $Count += $i
        $DSInserts.Tables[0] | Export-CSV "$WorkFolder\$Script:RFREnviron\$Table.csv" -Force -NoTypeInformation
        Write-Progress -Activity "Creating CSV files...” -status "Exporting $Table” -PercentComplete (($Count / $TotalCount) * 100)
    }
    Write-Host 'Done.' -Fore Yellow
    $SqlConn.Close()
 }

function Csv-ImportFile
{
    Write-Host ''
    Write-Host 'Importing CSV files... ' -Fore Yellow -NoNewline
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=$Script:keyDbName;Integrated Security=True"
    $SqlConn.Open()
    $AllTables = $Script:RFRTables + $Script:RFRDelTables
    if(!(Test-Path "$WorkFolder\$Script:RFREnviron\")) {
        $Script:WarningMsg = "Environment CSV Directory ($WorkFolder\$Script:RFREnviron\) does not exist."
        Get-Menu
    }
    else {
        $MissingFile = $null
        foreach($Table in $Script:RFRTables) {
            if(!(Test-Path "$WorkFolder\$Script:RFREnviron\$Table.csv")) {
                $MissingFile += "$Table.csv "
            }
        }
        if($MissingFile) {
            $Script:WarningMsg = "Environment CSV Files ($($MissingFile.Trim().Replace(" ",", "))) do not exist."
            Get-Menu
        }
    }
    foreach($Table in $AllTables) {
        $SqlQuery = “TRUNCATE TABLE [$Table]"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        $SqlCommand.ExecuteNonQuery() | Out-Null
        if(Test-Path "$WorkFolder\$Script:RFREnviron\$Table.csv") {
            $TotalCount += (Import-CSV "$WorkFolder\$Script:RFREnviron\$Table.csv" | Measure-Object).Count
        }
    }
    $count = 0
    $i = 0
    foreach($Table in $Script:RFRTables) {
        $TableCsv = Import-CSV "$WorkFolder\$Script:RFREnviron\$Table.csv" | Select Column1 -ExpandProperty Column1
        foreach($Row in $Tablecsv)
        {
            $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($Row, $SqlConn)
            $i = $SqlCommand.ExecuteNonQuery()
            $Count += $i
            Write-Progress -Activity "Importing CSV files...” -Status "Importing $Table” -PercentComplete (($Count / $TotalCount) * 100)
        }
    }
    $SqlConn.Close()
    Write-Host 'Done.' -Fore Yellow
}

function AOS-ServiceManager
{
[CmdletBinding()]
param (
    [Switch] $Stop,
    [Switch] $Start,
    [Switch] $Enable,
    [Switch] $Disable,
    [Switch] $Restart,
    [Switch] $Status
)
    $AOSServers = @()
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
    try {
        $SqlQuery = "SELECT INSTANCE FROM AX_Servers WHERE ENVIRONMENT = '$Script:RFREnviron' AND ACTIVE = 1"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $SqlCommand
        $DSServers = New-Object System.Data.DataSet
        $AOSCnt = $Adapter.Fill($DSServers)
        if($AOSCnt) {
            $AOSServers = $DSServers.Tables[0] | Select INSTANCE -ExpandProperty INSTANCE
        }
    }
    catch {
        $AOSCnt = 0
    }
    try {
        $SqlQuery = "SELECT ServerID FROM RFR_$Script:RFREnviron`_SYSSERVERCONFIG" 
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $Adapter.SelectCommand = $SqlCommand
        $RFRServers = New-Object System.Data.DataSet
        $AOSRFRCnt = $Adapter.Fill($RFRServers)
        $SqlConn.Close()
        #
        if($AOSRFRCnt) {
            $AOSServers = $RFRServers.Tables[0] | Select ServerID -ExpandProperty ServerID
        }
    }
    catch {
        $AOSRFRCnt = 0
    }
    if(($AOSRFRCnt -eq 0) -and ($AOSCnt -eq 0)) {
        Write-Host ''
        $ReadSrvs = Read-Host "Type AOS Server(s) [comma-separated]"
        if(!($ReadSrvs)) {
            $Script:WarningMsg = 'Invalid Option. Retry.'
            Get-Menu
        }
        foreach($Srv in $ReadSrvs.Split(',').Trim() | Select -Unique) {
            if(Test-Connection $Srv -Count 1 -Quiet) {
                $AOSServers += $Srv.ToUpper()
            }
            else {
                Write-Host ''
                Write-Warning "$($Srv.ToUpper()) Server is unreachable."
            }
        }
    }
    if($AOSServers) {
        foreach($AOS in $AOSServers) {
            $AOSName = $AOS.Split('@')[1]
            $InstanceName = $AOS.Split('@')[0]
            if(!(Test-Connection $AOSName -Count 1 -Quiet)) {
                #Removes AOS Server from Array
                $AOSServers = $AOSServers[1..($AOSServers.Length-$AOSServers.IndexOf($AOS))]
            }
        }
    }
    else {
        $Script:WarningMsg = 'Environment servers not found.'
        Break
    }
    if($AXVersion -eq '2012') {
        $serviceVersion = 'AOS60'
    }
    else {
        $serviceVersion = 'AOS50'
    }
    Write-Host ''
    $i = 0
    if($Stop) {
        try {
            Write-Host 'AOS Service Stop' -Fore Green
            foreach($AOS in $AOSServers) {
                $AOSName = $AOS.Split('@')[1]
                $InstanceName = $AOS.Split('@')[0]
                $i++
                Write-Host "$i. Stopping $($AOS)... " -Fore Yellow -NoNewline
                if($Disable) { Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName | Set-Service -StartupType Disabled }
                (Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $($AOSName)).Stop()
                Write-Host 'Done.' -Fore Yellow
            }
        }
        catch [Exception] {
            $Script:WarningMsg = "Nothing to Stop." 
        }                    
        catch {
            $Script:WarningMsg = $_.Exception.Message
        }
    }
    if($Start) {
        try {
            Write-Host 'AOS Service Start' -Fore Green
            foreach($AOS in $AOSServers) {
                $AOSName = $AOS.Split('@')[1]
                $InstanceName = $AOS.Split('@')[0]
                $i++
                Write-Host "$i. Starting $AOS... " -Fore Yellow -NoNewline
                if($Enable) { Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName | Set-Service -StartupType Automatic }
                (Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName).Start()
                Start-Sleep -s 3
                Write-Host 'Done.' -Fore Yellow
            }
        }
        catch [Exception] {
            $Script:WarningMsg = "Nothing to Start." 
        }                    
        catch {
            $Script:WarningMsg = $_.Exception.Message
        }
    }
    if($Restart) {
        try {
            Write-Host 'AOS Service Restart' -Fore Green
            foreach($AOS in $AOSServers) {
                $AOSName = $AOS.Split('@')[1]
                $InstanceName = $AOS.Split('@')[0]
                $i++
                Write-Host "$i. Restarting $AOS... " -Fore Yellow -NoNewline
                if($Disable) { Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName | Set-Service -StartupType Disabled }
                (Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName).Stop()
                Start-Sleep -s 5
                if($Enable) { Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName | Set-Service -StartupType Automatic }
                (Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName).Start()
                Write-Host 'Done.' -Fore Yellow
            }
        }
        catch [Exception] {
            $AOSServ = Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName | Select Status -ExpandProperty Status
            Write-Host "$AOSName is $AOSServ... " -Fore Yellow -NoNewline
            if($AOSServ -match 'Running') {
                Write-Host "trying to Stop it." -Fore Yellow
                (Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName).Stop()
            }
            else {
                Write-Host "trying to Start it." -Fore Yellow
                (Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName).Start()                    
            }
        }                    
        catch {
            $Script:WarningMsg = $_.Exception.Message
        }
    }
    if($Status) {
        try {
            Write-Host 'AOS Status Check' -Fore Green
            foreach($AOS in $AOSServers) {
                $AOSName = $AOS.Split('@')[1]
                $InstanceName = $AOS.Split('@')[0]
                $i++
                Write-Host "$i. Server $AOS is " -Fore Yellow -NoNewline
                Write-Host $(Get-Service -Name "$serviceVersion`$$($InstanceName)" -ComputerName $AOSName | Select Status -ExpandProperty Status) -Fore Yellow
            }
        }
        catch [Exception] {
            $Script:WarningMsg = "Nothing to Check." 
        }                    
        catch {
            $Script:WarningMsg = $_.Exception.Message
        }
    }
}

function RFR-UpdateGUID
{
    Write-Host ''
    Write-Warning 'DO NOT USE IN PRODUCTION.'
    Write-Host 'Updating GUID to 00000000-0000-0000-0000-000000000000.' -Fore Yellow
    Do {
        $Prompt = Read-Host "Confirm Update? (Y/N)"
    } While ($Prompt -notlike "[YN]")
    Switch($Prompt) {
        'Y' {
            try {
                $SqlConn = New-Object System.Data.SqlClient.SqlConnection
                $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=$Script:keyDbName;Integrated Security=True"
                $SqlConn.Open()
                $SqlQuery = "UPDATE SYSSQMSETTINGS SET GLOBALGUID = '{00000000-0000-0000-0000-000000000000}'"
                $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
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
 
function RFR-RecIdCheck
{
    Write-Host ''
    Write-Host 'Checking Table RecID' -Fore Green
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=$Script:keyDbName;Integrated Security=True"
    $SqlConn.Open()
    foreach($Table in $Script:RFRTables) {
        Write-Host "- Table $Table... " -Fore Yellow -NoNewline
        $SqlQuery = "SELECT MAX(RECID) FROM $Table"
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
        $TableRecId = $SqlCommand.ExecuteScalar()

        $SqlQuery = "SELECT NEXTVAL FROM SYSTEMSEQUENCES WHERE TABID IN (SELECT TABLEID FROM SQLDICTIONARY WHERE NAME = '$Table' AND FIELDID = 0)"
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
        $SystemRecId = $SqlCommand.ExecuteScalar()
        if($TableRecId.GetType().Name -match 'DBnull') { $TableRecId = 0 }
        if(($TableRecId -gt $SystemRecId) -and ($TableRecId -ne 0)) {
            try {
                $TableRecId++
                $SqlQuery = "UPDATE SYSTEMSEQUENCES SET NEXTVAL = $TableRecId WHERE TABID IN (SELECT TABLEID FROM SQLDICTIONARY WHERE NAME = '$Table' AND FIELDID = 0)"
                $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
                $SqlCommand.ExecuteNonQuery() | Out-Null
                $Script:WarningMsg = 'Restart AOS Servers.'
                Write-Host 'Updated.' -Fore Yellow
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

function RFR-BatchJobs
{
[CmdletBinding()]
param (
	[String] $Action
)
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=$Script:keyDbName;Integrated Security=True"
    $SqlConn.Open()
   
    Switch($Action) {
        DisableJobs {
            Write-Host ''
            Do {
                $Prompt = Read-Host "Delete BatchServerGroup? (Y/N)"
            } While ($Prompt -notlike "[YN]")
            Switch($Prompt) {
                'Y' {
                    Write-Host 'Removing Batch Server Groups...' -Fore Yellow -NoNewline
                    $SqlQuery = 'TRUNCATE TABLE BATCHSERVERGROUP'
                    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
                    $SqlCommand.ExecuteNonQuery() | Out-Null
                    Write-Host " Done." -Fore Yellow
                }
                'N' {
                    $Script:WarningMsg = 'Canceled.'
                    Break
                }
            }
        }
        ChangeServer {
            $SqlQuery = "SELECT SERVERID, ENABLEBATCH FROM SYSSERVERCONFIG ORDER BY ENABLEBATCH DESC, RECID DESC"
            $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
            $SqlAdapter.SelectCommand = $SqlCommand
            $BatchServers = New-Object System.Data.DataSet
            $SqlAdapter.Fill($BatchServers) | Out-Null

            $SqlQuery = "SELECT GROUP_ FROM BATCHGROUP"
            $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
            $SqlAdapter.SelectCommand = $SqlCommand
            $BatchGroups = New-Object System.Data.DataSet
            $SqlAdapter.Fill($BatchGroups) | Out-Null 

            if($BatchGroups.Tables[0].Rows.Count -eq 0) {
                $SqlQuery = "INSERT INTO BATCHGROUP VALUES('','Empty batch group','2012-05-08 00:03:00.000','61435','-AOS-',0,5637144576)"
                $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
                $SqlCommand.ExecuteNonQuery() | Out-Null
                $BatchGroups.Tables[0].Rows.Add('')
            }
            Write-Host ''
            Write-Host 'Moving Batch Jobs to Batch Server:' -Fore Green
            $i = 0         
            foreach ($Server in $BatchServers.Tables[0]) {
                if($Server.ENABLEBATCH -eq "1")
                {
                    $i++
                    Write-Host "$i. $(($Server.ServerId).Substring(3,($Server.ServerId).length-3)) " -Fore Yellow -NoNewline
                    try{
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
            Do {
                $Prompt = Read-Host "Option (1/$i)"
            } While (($Prompt -notlike "[1-$i]") -and ($Prompt))

            if($Prompt) {
                $ServerId = ($BatchServers.Tables[0])[$Prompt-1].ServerId
                Write-Host ''
                $SqlQuery = "TRUNCATE TABLE BatchServerGroup"
                $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
                $SqlCommand.ExecuteNonQuery() | Out-Null

                foreach($Group in $BatchGroups.Tables[0]) {
                    Write-Host "- Creating group $($Group.GROUP_) at $($ServerId.Substring(3))" -Fore Yellow
                    $SqlQuery = "INSERT INTO BATCHSERVERGROUP (GROUPID, SERVERID, CREATEDDATETIME, CREATEDBY, RECVERSION, RECID) " +
                                    "VALUES ('$($Group.GROUP_)','$ServerId','$DateTime', 'Admin', 1, $(RFR-NextRecId 'BATCHSERVERGROUP'))"
                    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
                    $SqlCommand.ExecuteNonQuery() | Out-Null
                }

                $SqlQuery = "SELECT GROUPID, CAPTION, RECID FROM BATCH A WHERE NOT EXISTS (SELECT * FROM BATCHGROUP B WHERE A.GROUPID = B.GROUP_)"
                $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
                $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
                $SqlAdapter.SelectCommand = $SqlCommand
                $BatchTasks = New-Object System.Data.DataSet
                $TaskCnt = $SqlAdapter.Fill($BatchTasks)

                if($TaskCnt -gt 0) {
                    Write-Host 'Inconsistency on Batch Tasks and Batch Groups.'
                    Write-Host 'Batchs:'
                    $i = 0
                    foreach ($Task in $BatchTasks.Tables[0]) {
                            $i++
                            Write-Host "$i. $($Task.Caption) - Group $($Task.GroupId)"
                    }
                    do {
                        $Prompt = Read-Host "Would you like to update them to 'Empty Batch Group'? (Y/N)"
                    } While ($Prompt.ToUpper() -notmatch "[YN]")

                    Switch($Prompt.ToUpper()) {
                        'Y' {
                            foreach ($UpdTask in $BatchTasks.Tables[0]) {
                                $SqlQuery = "UPDATE BATCH SET GROUPID = '' WHERE RECID = '$($UpdTask.RecId)'"
                                $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
                                $SqlCommand.ExecuteNonQuery() | Out-Null
                            }
                        }
                        'N'{
                            $Script:WarningMsg = 'Canceled.'
                            Continue
                        }
                    }
                }
                $Script:WarningMsg = 'Restart AOS Servers.'
            }
            else {
                $Script:WarningMsg = 'Canceled.'
            }
        }
        Default {
            $Script:WarningMsg = 'Incorrect call.'
        }
    }
    $SqlConn.Close()
}

function RFR-NextRecId
{
[CmdletBinding()]
param (
	[String] $TableName
)
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=$Script:keyDbName;Integrated Security=True"
    $SqlConn.Open()
    $SqlQuery = "SELECT TABLEID FROM SQLDICTIONARY WHERE NAME = '$TableName' AND FIELDID = 0"
    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
    $AxTableID = $SqlCommand.ExecuteScalar()
    $SqlQuery = "DECLARE @recId bigint; EXEC [dbo].[sp_GetNextRecId] $AxTableID, @recId = @recId OUTPUT; SELECT  @recId"
    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
    $NextRecID = $SqlCommand.ExecuteScalar()
    $SqlConn.Close()
    
    return $NextRecID
}

function RFR-BatchHistory
{
    Write-Host ''
    $defaultValue = Get-Date
    ($defaultValue,(Read-Host "Delete all Batch History date or [Enter] for [$(Get-Date -Date $defaultValue -Format d)]")) -match '\S' | % {$delDate = $ret = $_}
    $DateFormat = 'mm/dd/yyyy' # hh HH mm ss dd yyyy
    try {
        [datetime]::TryParseexact(
        $delDate, 
        $DateFormat,
        [system.Globalization.DateTimeFormatInfo]::InvariantInfo,
        [system.Globalization.DateTimeStyles]::None,
        [ref]$ret) | Out-Null
        Write-Host ''
        Write-Host "Delete Batch History tables. Cut-off Date as $delDate" -Fore Green
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$Script:keyDbServer;Database=$Script:keyDbName;Integrated Security=True"
        $SqlConn.Open()

        Write-Host '- Deleting BatchConstraintsHistory... ' -Fore Yellow -NoNewline
        $SqlQuery = "DELETE FROM BATCHCONSTRAINTSHISTORY WHERE BATCHCONSTRAINTSHISTORY.BATCHID IN " +
                     "(SELECT BATCHID FROM BATCHHISTORY BH JOIN BATCHJOBHISTORY BJH " +
                     "ON BH.BATCHJOBHISTORYID = BJH.RECID " +
                     "WHERE BJH.STATUS IN (3,4,8) " +
                     "AND DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), BJH.CREATEDDATETIME) <= '$delDate')"
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
        $SqlCommand.CommandTimeout = 0
        $SqlCommand.ExecuteNonQuery() | Out-Null
        Write-Host 'Done.' -Fore Yellow

        Write-Host '- Deleting BatchHistory... ' -Fore Yellow -NoNewline
        $SqlQuery = "DELETE FROM BATCHHISTORY WHERE BATCHHISTORY.BATCHJOBHISTORYID " +
                     "IN (SELECT RECID FROM BATCHJOBHISTORY BJH " +
                     "WHERE BJH.STATUS IN (3,4,8) " +
                     "AND DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), BJH.CREATEDDATETIME) <= '$delDate')"
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
        $SqlCommand.CommandTimeout = 0
        $SqlCommand.ExecuteNonQuery() | Out-Null
        Write-Host 'Done.' -Fore Yellow

        Write-Host '- Deleting BatchJobHistory... ' -Fore Yellow -NoNewline
        $SqlQuery = "DELETE FROM BATCHJOBHISTORY WHERE BATCHJOBHISTORY.STATUS " +
                     "IN (3,4,8) " +
                     "AND DATEADD(mi, DATEDIFF(mi, GETUTCDATE(), GETDATE()), BATCHJOBHISTORY.CREATEDDATETIME) <= '$delDate'"
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery, $SqlConn)
        $SqlCommand.CommandTimeout = 0
        $SqlCommand.ExecuteNonQuery() | Out-Null
        Write-Host 'Done.' -Fore Yellow
        $SqlConn.Close()
    }
    catch { 
        $Script:WarningMsg = "Invalid Date --> $($_.Exception.Message)"
    }
}

function RFR-RunAsJob
{
    $Script:IsJob = $true
    RFR-LoadEnvironment

    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
    $SqlConn.Open()
    $SqlQuery = “SELECT TOP 1 CREATEDDATETIME FROM AX_EnvironmentStore WHERE ENVIRONMENT = '$Script:RFREnviron' ORDER BY CREATEDDATETIME DESC"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
    $StoreDate = $SqlCommand.ExecuteScalar()
    $SqlConn.Close()

    if($StoreDate -like '') {
        Write-Log 'Enviroment Configuration does not exist.' 'Error'
        Exit
    }    
    elseif($RestoreDB) {
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
        $SqlConn.Open()
        $SqlQuery = “SELECT [BKPFOLDER] FROM [AX_Environments] WHERE ENVIRONMENT = '$Script:RFREnviron'"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        $SQLBackup = $SqlCommand.ExecuteScalar()
        $SqlConn.Close()
        if($SQLBckPath -notlike '') {
            if(Test-Path $SQLBckPath) {
                $Script:SQLBackup = Get-ChildItem -Path $SQLBckPath | Select -First 1 | Where-Object { $_.Extension -match '.bak' } | Sort-Object -Property CreationTime -Descending
                AOS-ServiceManager -Stop
                SQL-DBRestore $Script:SQLBackup.FullName
                SQL-CleanUpTable
                RFR-StoreManager -Restore
                AOS-ServiceManager -Start
                $Script:RFROk = $true
            }
            else {
                Write-Log 'Incorrect Path.' 'Error'
            }
        }
    }
    elseif($RefreshOnly) {
        AOS-ServiceManager -Stop
        Start-Sleep -Seconds 5
        SQL-CleanUpTable
        RFR-StoreManager -Restore
        AOS-ServiceManager -Start
    }
    elseif($RefreshDays -ge 1) {
        if(($StoreDate - $((Get-Date).AddDays($RefreshDays * -1))).Days -eq 0) {
            RFR-StoreManager -Backup
        }
        else {
            Write-Log "Environment Date: $StoreDate" 'Warn'
        }
    }
    else {
        Write-Log 'Incorrect Parameters.' 'Error'
    }
}

function RFR-SQLValidation
{
[CmdletBinding()]
param (
	[String] $ServerName,
    [String] $DBName
)
    try {
        $Server = New-Object "Microsoft.SqlServer.Management.Smo.Server" "$ServerName"
        if(!($Server.Databases.Name.ToUpper().Contains($DBName.ToUpper()))) {
                Write-Warning 'Database does not exist.'
                Exit
        }
    }
    catch {
        Write-Warning "Failed to connect to $ServerName or User doesn't have access."
        Exit
    }

    try {
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$ServerName;Database=$DBName;Integrated Security=True"
        $SqlConn.Open()
        $SqlConn.Close()
        RFR-StoreManager -Check
    }
    catch {
        $ErrMsg = $_.Exception
        While($ErrMsg.InnerException ) {
            $ErrMsg = $ErrMsg.InnerException
            if(($ErrMsg.Message).Contains('Login failed')) {
                $Script:WarningMsg = "Login failed to $DBName."
                $SqlFailed = $true
                Break
            }
        }
    }
}

function Validate-Servers
{
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
    $SqlQuery = "SELECT ServerName FROM AX_Servers WHERE ENVIRONMENT = '$Script:RFREnviron'"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $SqlCommand
    $DSServers = New-Object System.Data.DataSet
    $AOSCnt = $Adapter.Fill($DSServers)
    if($AOSCnt) {
        $AOSServers = @()
        $AOSServers = $DSServers.Tables[0] | Select ServerName -ExpandProperty ServerName -Unique
    }
    if($AOSServers) {
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
        $SqlConn.Open()
        foreach($Server in $AOSServers) {
            $ServerIp = Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue
            $Computer = Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges -ComputerName $Server -ErrorAction SilentlyContinue
            if($Computer.Domain) { $FQDN = $Server + "." + $Computer.Domain } else { $FQDN = '' }
            $Prompt = Read-Host "Would you like to enable $Server $($Computer.Domain)?"
            if($Prompt -like 'Y') {
                $Active = '1'
            }
            else {
                $Active = '0'
            }
            $SqlQuery = "UPDATE [AX_Servers] SET ACTIVE = '$Active', IP = '$(($ServerIp.IPV4Address).IPAddressToString)', DOMAIN = '$($Computer.Domain)', FQDN = '$($FQDN)' WHERE ENVIRONMENT = '$Script:RFREnviron' AND ServerName = '$Server'"
            $SqlCommand = New-Object System.Data.SqlClient.SqlCommand($SqlQuery,$SqlConn)
            $SqlCommand.ExecuteNonQuery() | Out-Null
        }
    $SqlConn.Close()
    }
}

function Update-SQLBKPFolder
{
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
    $SqlConn.Open()
    $SqlQuery = “SELECT [BKPFOLDER] FROM [AX_Environments] WHERE ENVIRONMENT = '$Script:RFREnviron'"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
    $SQLBackup = $SqlCommand.ExecuteScalar()
    $SqlConn.Close()
    if($SQLBackup -notlike '') {
        Write-Host ' '
        Write-Host "Current Folder set to $SQLBackup"
    }
    else {
        Write-Host ' '
        Write-Host "Backup folder not set."
    }
    Write-Host ' '
    $NewBkpFolder = Read-Host "Please, enter the new folder path"

    if($NewBkpFolder -and (Test-Path $NewBkpFolder)) {
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
        $SqlConn.Open()
        $SqlQuery = “UPDATE [AX_Environments] SET [BKPFOLDER] = '$NewBkpFolder' WHERE ENVIRONMENT = '$Script:RFREnviron'"
        $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
        $SQLBackup = $SqlCommand.ExecuteScalar()
        $SqlConn.Close()
    }
    else {
        $Script:WarningMsg = 'Invalid Folder Path.'
    }

}

function Send-Email
{
    Import-Module $Dir\AX-StringCrypto.psm1

    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
    $SqlQuery = “SELECT * FROM AX_EmailProfile WHERE PROFILEID = 'RFR'"
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter($SqlQuery, $SqlConn)
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null
    
    if (![string]::IsNullOrEmpty($Table.Tables))
    {
        $EmailSettings = New-Object -TypeName System.Object
        $EmailSettings | Add-Member -Name SMTPServer -Value $($Table.Tables.ConnectionInfo.Split(',')[0]) -MemberType NoteProperty
        $EmailSettings | Add-Member -Name SMTPPort -Value $($Table.Tables.ConnectionInfo.Split(',')[1]) -MemberType NoteProperty
        $EmailSettings | Add-Member -Name SMTPUserName -Value $(Read-EncryptedString -InputString $Table.Tables.ConnectionInfo.Split(',')[2] -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)") -MemberType NoteProperty
        $EmailSettings | Add-Member -Name SMTPPassword -Value $(Read-EncryptedString -InputString $Table.Tables.ConnectionInfo.Split(',')[3] -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)") -MemberType NoteProperty
        $EmailSettings | Add-Member -Name SMTPSSL -Value $($Table.Tables.ConnectionInfo.Split(',')[4]) -MemberType NoteProperty
        $EmailSettings | Add-Member -Name SMTPTo -Value $Table.Tables.To -MemberType NoteProperty
        $EmailSettings | Add-Member -Name SMTPCC -Value $Table.Tables.CC -MemberType NoteProperty
        $Table.Dispose()
    }

    #Message Parameters
    $SMTPMessage = New-Object System.Net.Mail.MailMessage
    $SMTPMessage.From = 'AX Refresh Job <ken.anderson@tdwilliamson.com>'
    if($EmailSettings.SMTPTo) {$EmailSettings.SMTPTo.Split(';') | % { $SMTPMessage.To.Add($_.Trim()) }} else { break } #add log for no To line
    if($EmailSettings.SMTPCC) {$EmailSettings.SMTPCC.Split(';') | % { $SMTPMessage.CC.Add($_.Trim()) }}
    $SMTPMessage.Subject = "$(if($Script:RFROk){"SUCCESS "} else {"FAILED "})Environment $Script:RFREnviron has been refreshed on server $Script:MachineFullName."
    #$SMTPMessage.IsBodyHtml = $true
    $SMTPMessage.Body = "Script executed by $env:userdomain\$env:username on $env:ComputerName. Parameters: $Paramlist. $(if($Script:SQLBackup) {"SQL Backup: $($Script:SQLBackup.FullName) from $($Script:SQLBackup.CreationTime)."})"
    #Create Message
    $SMTPClient = New-Object System.Net.Mail.SmtpClient($EmailSettings.SMTPServer,$EmailSettings.SMTPPort)
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($EmailSettings.SMTPUserName,$EmailSettings.SMTPPassword)
    #Send Email
    if($EmailSettings.SMTPSSL -like 'True') {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
    }
    $SMTPClient.Send($SMTPMessage)
}

function Write-Log
{
Param(
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Message,
    [Parameter(Mandatory=$false)]
    [ValidateSet("Error","Warn","Info")]
    [string]$Level="Info"
)
    switch ($Level) {
    'Error' {
        $LogText = "ERROR: $(if($Script:RFREnviron) {"($Script:RFREnviron) "})$Message"
        }
    'Warn' {
        $LogText = "WARNING: $(if($Script:RFREnviron) {"($Script:RFREnviron) "})$Message"
        }
    'Info' {
        $LogText = "INFO: $(if($Script:RFREnviron) {"($Script:RFREnviron) "})$Message"
        }
    }
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True"
    $SqlConn.Open()
    $SqlQuery = “INSERT INTO [AX_ExecutionLog] (LOG) VALUES ('$($LogText.Replace("'","''"))')"
    $SqlCommand = New-Object System.Data.SqlClient.SQLCommand($SqlQuery,$SqlConn)
    $SqlCommand.ExecuteNonQuery() | Out-Null
    $SqlConn.Close()
}

function Get-RFRStart
{
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
    RFR-SQLValidation $ParamDBServer $ParamDBName
    RFR-VariableWiper
    $DateTime = Get-Date
    Write-Log ("Refresh Script has started - $ScriptPath")
    Write-Log ("User: $env:userdomain\$env:username")
    Write-Log ("Active Connections: $((Get-WmiObject Win32_LoggedOnUser | Select Antecedent -Unique | Where { $_.Antecedent.ToString().Split('"')[1] -like $env:userdomain } | % { “{0}\{1}” -f $_.Antecedent.ToString().Split('"')[1], $_.Antecedent.ToString().Split('"')[3] }) -join ', ')")
    if($Environment)
    {
        $Script:RFREnviron = $Environment
        Write-Log ("Script Parameters: $Paramlist")
        RFR-RunAsJob
    }
    else {
        Write-Host ''
        Write-Host 'Environment Refresh Tool'
        Write-Host '════════════════════════════════════════════════════════════'
        Get-MainMenu
    }
}

foreach($Param in $PSBoundParameters.GetEnumerator()) {
  $Paramlist += ("{0}: {1} | " -f $Param.Key,$Param.Value)
}

if($Paramlist ) {
    $Paramlist = $Paramlist.Substring(0,$Paramlist.Length -3)
}

Get-RFRStart
if($SendEmail) {Send-Email}
Write-Log ("Refresh Script has finished.")