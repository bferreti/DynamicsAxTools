$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$SettingsFolder = $Dir + "\AX-Tools"

# Import settings from config file
#[xml]$ConfigFile = Get-Content "$SettingsFolder\Settings.xml"

function Load-ConfigFile
{
    if(Test-Path "$SettingsFolder\Settings.xml") {
        # Import settings from config file
        [xml]$ConfigFile = Get-Content "$SettingsFolder\Settings.xml"
    }
    else {
        Write-Warning "Configuration file does not exists."
    }

    return $ConfigFile
}

function Get-ConnectionString {
    $ParamDBServer = $ConfigFile.Settings.Database.DBServer
    $ParamDBName = $ConfigFile.Settings.Database.DBName
    return "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True;Connect Timeout=5"
}

function SQL-BulkInsert
{
param (
    [String]$Table,
    [Array]$Data
)
    #Write-Host "Insert Table: $Table"
    $DataTable = New-Object Data.DataTable   
    $First = $true  
    foreach($Object in $Data) 
    { 
        $DataReader = $DataTable.NewRow()   
        foreach($Property in $($Object.PsObject.Get_Properties() | Where-Object { $_.Name -notmatch "(PSCOMPUTERNAME|RUNSPACEID|PSSHOWCOMPUTERNAME)"})) 
        {   
            if ($First) 
            {   
                $Col =  New-Object Data.DataColumn   
                $Col.ColumnName = $Property.Name.ToString()   
			    $ValueExists = Get-Member -InputObject $Property -Name Value
			    if($ValueExists)
                { 
                    if($Property.Value -isnot [System.DBNull] -and $Property.Value -ne $null) {
                        $Col.DataType = [System.Type]::GetType("$($Property.TypeNameOfValue)")
                    } 
                } 
                $DataTable.Columns.Add($Col) 
            }
            $DataReader.Item($Property.Name) = $Property.Value 
        }   
        $DataTable.Rows.Add($DataReader)   
        $First = $false
    }
    #$DataTable.Columns.Remove('PSCOMPUTERNAME')
    #$DataTable.Columns.Remove('RUNSPACEID')
    #$DataTable.Columns.Remove('PSSHOWCOMPUTERNAME')
    #Write-Output @(,($DataTable))
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
    $BCopy = New-Object ("System.Data.SqlClient.SqlBulkCopy") $Conn
    $BCopy.DestinationTableName = "[dbo].[$Table]"
    foreach ($Col in $DataTable.Columns) {
        $ColumnMap = New-Object ("Data.SqlClient.SqlBulkCopyColumnMapping") $Col.ColumnName,($Col.ColumnName).ToUpper()
        [Void]$BCopy.ColumnMappings.Add($ColumnMap)
        #$ColumnMap
    }
    $BCopy.WriteToServer($DataTable)
    $Conn.Close()
}

function Write-EncryptedString {
param ([String]$InputString, [String]$DTKey)

if (($args -contains '-?') -or (-not $InputString) -or (-not $DTKey)) {
return
}
	$Rfc2898 = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($DTKey,32)
	$Salt = $Rfc2898.Salt
	$AESKey = $Rfc2898.GetBytes(32)
	$AESIV = $Rfc2898.GetBytes(16)
	$Hmac = New-Object System.Security.Cryptography.HMACSHA1(,$Rfc2898.GetBytes(20))
	$AES = New-Object Security.Cryptography.RijndaelManaged
	$AESEncryptor = $AES.CreateEncryptor($AESKey, $AESIV)
	$InputDataStream = New-Object System.IO.MemoryStream
	if ($Compress) { $InputEncodingStream = (New-Object System.IO.Compression.GZipStream($InputDataStream, 'Compress', $True)) }
	else { $InputEncodingStream = $InputDataStream }
	$StreamWriter = New-Object System.IO.StreamWriter($InputEncodingStream, (New-Object System.Text.Utf8Encoding($true)))
	$StreamWriter.Write($InputString)
	$StreamWriter.Flush()
	$InputData = $InputDataStream.ToArray()
	$EncryptedEncodedInputString = $AESEncryptor.TransformFinalBlock($InputData, 0, $InputData.Length)
	$AuthCode = $Hmac.ComputeHash($EncryptedEncodedInputString)
	$OutputData = New-Object Byte[](52 + $EncryptedEncodedInputString.Length)
	[Array]::Copy($Salt, 0, $OutputData, 0, 32)
	[Array]::Copy($AuthCode, 0, $OutputData, 32, 20)
	[Array]::Copy($EncryptedEncodedInputString, 0, $OutputData, 52, $EncryptedEncodedInputString.Length)
	$OutputDataAsString = [Convert]::ToBase64String($OutputData)
}

function Read-EncryptedString {
param ([String]$InputString, [String]$DTKey)

if (($args -contains '-?') -or (-not $InputString) -or (-not $DTKey -and -not $InputString.StartsWith('-----BEGIN PGP MESSAGE-----'))) {
return
}
    # Decrypt with custom algo
	$InputData = [Convert]::FromBase64String($InputString)
	$Salt = New-Object Byte[](32)
	[Array]::Copy($InputData, 0, $Salt, 0, 32)
	$Rfc2898 = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($DTKey,$Salt)
	$AESKey = $Rfc2898.GetBytes(32)
	$AESIV = $Rfc2898.GetBytes(16)
	$Hmac = New-Object System.Security.Cryptography.HMACSHA1(,$Rfc2898.GetBytes(20))
	$AuthCode = $Hmac.ComputeHash($InputData, 52, $InputData.Length - 52)
	if (Compare-Object $AuthCode ($InputData[32..51]) -SyncWindow 0) {
		throw 'Checksum failure.'
	}
	$AES = New-Object Security.Cryptography.RijndaelManaged
	$AESDecryptor = $AES.CreateDecryptor($AESKey, $AESIV)
	$DecryptedInputData = $AESDecryptor.TransformFinalBlock($InputData, 52, $InputData.Length - 52)
	$DataStream = New-Object System.IO.MemoryStream($DecryptedInputData, $false)
	if ($DecryptedInputData[0] -eq 0x1f) {
		$DataStream = New-Object System.IO.Compression.GZipStream($DataStream, 'Decompress')
	}
	$StreamReader = New-Object System.IO.StreamReader($DataStream, $true)
	$StreamReader.ReadToEnd()
}


function Write-Log
{
param (
    [string]$LogData
)
    $Script:Settings
    $TLogStamp = (Get-Date -DisplayHint Time)
    $ExecLog = New-Object -TypeName System.Object
    $ExecLog | Add-Member -Name CreatedDateTime -Value $TLogStamp -MemberType NoteProperty
    $ExecLog | Add-Member -Name Guid -Value $Global:Guid -MemberType NoteProperty
    $ExecLog | Add-Member -Name Log -Value $LogData.Trim() -MemberType NoteProperty
    SQL-BulkInsert 'AXTools_ExecutionLog' $ExecLog
}

function UpdateMsiStatus
{
param (
    [string]$Status
)
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
    $Query = "UPDATE [dbo].[AXInstallStatus] SET Status = '$Status' WHERE GUID = '$Guid'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function SQL-UpdateTable
{
param (
    [String]$Table,
    [String]$Set,
    [String]$Value,
    [String]$Where
)
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
    $Query = "UPDATE [dbo].[$Table] SET [$Set] = '$Value' WHERE $Where"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function SQL-ExecUpdate
{
param (
    [String]$Query
)
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
    #$Query = "UPDATE [dbo].[$Table] SET [$Set] = '$Value' WHERE $Where"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function SQL-WriteLog
{
param (
    [String]$Log
)
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
    $Query = "INSERT INTO AXTools_ExecutionLogs VALUES('$(Get-Date)', '', 'AXMonitor', '$Log')"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}