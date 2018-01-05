$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"

function Load-ConfigFile
{
    if(Test-Path "$ModuleFolder\AX-Settings.xml") {
        # Import settings from config file
        [xml]$ConfigFile = Get-Content "$ModuleFolder\AX-Settings.xml"
    }
    else {
        Write-Warning "Configuration file does not exists."
    }

    return $ConfigFile
}

function Get-ConnectionString 
{
[CmdletBinding()]
param (
    [String]$ApplicationName
)
    if($ApplicationName -eq '') { $ApplicationName = 'AX Powershell Script' }
    $ConfigFile = Load-ConfigFile
    $ParamDBServer = $ConfigFile.Settings.Database.DBServer
    $ParamDBName = $ConfigFile.Settings.Database.DBName
    $ParamUserName = $ConfigFile.Settings.Database.Impersonation.UserName
    $ParamPassword = $ConfigFile.Settings.Database.Impersonation.Password
    if($ParamUserName) {
        $UserPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($ParamPassword | ConvertTo-SecureString)))
        $secureUserPassword = $UserPassword | ConvertTo-SecureString -AsPlainText -Force 
        $SqlCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $ParamUserName, $secureUserPassword
        $SqlConn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection
        $SqlConn.ServerInstance = $ParamDBServer
        $SqlConn.DatabaseName = $ParamDBName
        $SqlConn.ApplicationName = $ApplicationName
        $SqlServer = New-Object Microsoft.SqlServer.Management.SMO.Server($SqlConn)
        $SqlServer.ConnectionContext.ConnectAsUser = $true
        $SqlServer.ConnectionContext.ConnectAsUserPassword = $SqlCredential.GetNetworkCredential().Password
        $SqlServer.ConnectionContext.ConnectAsUserName = $SqlCredential.GetNetworkCredential().UserName
        try {
            $SqlServer.ConnectionContext.Connect()
            return $SqlServer.ConnectionContext.SqlConnectionObject
        }
        catch {
            Write-Host "Failed to connect to AXTools Database. $($_.Exception.Message)"
            break
        }
    }
    else {
        $SqlConn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection
        $SqlConn.ServerInstance = $ParamDBServer
        $SqlConn.DatabaseName = $ParamDBName
        $SqlConn.ApplicationName = $ApplicationName
        $SqlServer = New-Object Microsoft.SqlServer.Management.SMO.Server($SqlConn)
        try {
            $SqlServer.ConnectionContext.Connect()
            return $SqlServer.ConnectionContext.SqlConnectionObject
        }
        catch {
            Write-Host "Failed to connect to AXTools Database. $($_.Exception.Message)"
            break
        }
    }
}

function Get-SQLObject
{
[CmdletBinding()]
param (
    [Switch]$SQLServerObject,
    [String]$SQLAccount,
    [String]$DBServer,
    [String]$DBName,
    [String]$ApplicationName
)
    if($ApplicationName -eq '') { $ApplicationName = 'AX Powershell Script' }
    try {
        if($SQLAccount) {
            $Query = "SELECT Password FROM [dbo].[AXTools_UserAccount] WHERE [USERNAME] = '$($SQLAccount)'"
            $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$(Get-ConnectionString))
            $UserPassword = $Cmd.ExecuteScalar()
            $UserPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($UserPassword | ConvertTo-SecureString)))
            $secureUserPassword = $UserPassword | ConvertTo-SecureString -AsPlainText -Force 
            $SqlCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $SQLAccount, $secureUserPassword
            $SqlConn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection
            $SqlConn.ServerInstance = $DBServer
            $SqlConn.DatabaseName = $DBName
            $SqlConn.ApplicationName = $ApplicationName
            $SqlServer = New-Object Microsoft.SqlServer.Management.SMO.Server($SqlConn)
            $SqlServer.ConnectionContext.ConnectAsUser = $true
            $SqlServer.ConnectionContext.ConnectAsUserPassword = $SqlCredential.GetNetworkCredential().Password
            $SqlServer.ConnectionContext.ConnectAsUserName = $SqlCredential.GetNetworkCredential().UserName
            $SqlServer.ConnectionContext.Connect()
        }
        else {
            $SqlConn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection
            $SqlConn.ServerInstance = $DBServer
            $SqlConn.DatabaseName = $DBName
            $SqlConn.ApplicationName = $ApplicationName
            $SqlServer = New-Object Microsoft.SqlServer.Management.SMO.Server($SqlConn)
            $SqlServer.ConnectionContext.Connect()
        }
    }
    catch {
        Write-Host "Failed to connect to AX Database. $($_.Exception.Message)"
        break
    }

    if($SQLServerObject) {
        return $SqlServer
    }
    else {
        return $SqlServer.ConnectionContext.SqlConnectionObject
    }
}

function SQL-BulkInsert
{
[CmdletBinding()]
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
    $Conn = Get-ConnectionString #$Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    #$Conn.Open()
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
[CmdletBinding()]
param (
    [String]$InputString, 
    [String]$DTKey
)
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
[CmdletBinding()]
param (
    [String]$InputString, 
    [String]$DTKey
)
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
[CmdletBinding()]
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
[CmdletBinding()]
param (
    [string]$Status
)
    $Conn = Get-ConnectionString #$Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    #$Conn.Open()
    $Query = "UPDATE [dbo].[AXInstallStatus] SET Status = '$Status' WHERE GUID = '$Guid'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function SQL-UpdateTable
{
[CmdletBinding()]
param (
    [String]$Table,
    [String]$Set,
    [String]$Value,
    [String]$Where
)
    $Conn = Get-ConnectionString #$Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    #$Conn.Open()
    $Query = "UPDATE [dbo].[$Table] SET [$Set] = '$Value' WHERE $Where"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function SQL-ExecUpdate
{
[CmdletBinding()]
param (
    [String]$Query
)
    $Conn = Get-ConnectionString #$Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    #$Conn.Open()
    #$Query = "UPDATE [dbo].[$Table] SET [$Set] = '$Value' WHERE $Where"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function SQL-WriteLog
{
[CmdletBinding()]
param (
    [String]$Log
)
    $Conn = Get-ConnectionString #$Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    #$Conn.Open()
    $Query = "INSERT INTO AXTools_ExecutionLogs VALUES('$(Get-Date)', '', 'AXMonitor', '$Log')"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function Send-Email
{
[CmdletBinding()]
param (
    [String]$Subject,
    [Object]$Body,
    [String]$Attachment,
    [String]$EmailProfile,
    [String]$Guid
)
    $Conn = Get-ConnectionString #$Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT * FROM [AXTools_EmailProfile] AS A
                JOIN [AXTools_UserAccount] AS B ON A.UserID = B.ID
                WHERE A.ID = '$EmailProfile'"
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter($Query, $Conn)
    $Table = New-Object System.Data.DataSet
    $Adapter.Fill($Table) | Out-Null

    if (![string]::IsNullOrEmpty($Table.Tables))
    {
        $SMTPServer = $Table.Tables.SMTPServer
        $SMTPPort = $Table.Tables.SMTPPort
        $SMTPUserName = $Table.Tables.UserName
        $SMTPPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($Table.Tables.Password | ConvertTo-SecureString)))
        $SMTPSSL = $Table.Tables.SMTPSSL
        $SMTPFrom = $Table.Tables.From
        $SMTPTo = $Table.Tables.To
        $SMTPCC = $Table.Tables.CC
        $SMTPBCC = $Table.Tables.BCC
        $Table.Dispose()
    }
    else {
        break
    }

    #Message Parameters
    $SMTPMessage = New-Object System.Net.Mail.MailMessage
    $SMTPMessage.From = $SMTPFrom
    if(-not [System.DBNull]::Value.Equals($SMTPTo)) {$SMTPTo.Split(';') | % { $SMTPMessage.To.Add($_.Trim()) }} else { break }
    if(-not [System.DBNull]::Value.Equals($SMTPCC)) {$SMTPCC.Split(';') | % { $SMTPMessage.CC.Add($_.Trim()) }}
    if(-not [System.DBNull]::Value.Equals($SMTPBCC)) {$SMTPBCC.Split(';') | % { $SMTPMessage.Bcc.Add($_.Trim()) }}
    $SMTPMessage.Subject = $Subject
    $SMTPMessage.IsBodyHtml = $true
    $SMTPMessage.Body = $Body

    #Attachemnts
    if($Attachment) {
        $AttachmentFile = New-Object System.Net.Mail.Attachment($Attachment)
        $SMTPMessage.Attachments.Add($AttachmentFile)
    }

    #Create Message
    $SMTPClient = New-Object System.Net.Mail.SmtpClient($SMTPServer,$SMTPPort)
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPUserName,$SMTPPassword)

    #Send Email
    if($SMTPSSL -eq 1) {
        $SMTPClient.EnableSsl = $true
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
    }

    try
    {
        $SMTPClient.Send($SMTPMessage)
        $Sent = '1'
        $Log = ''
        SQL-ExecUpdate "UPDATE AXMonitor_ExecutionLog SET EMAIL = '1' WHERE GUID = '$Guid'"
    }
    catch
    {
        $Sent = '0'
        $Log = $_.Exception
    }

    SQL-BulkInsert AXTools_EmailLogs @($SMTPMessage | Select @{n='Sent';e={[int]$Sent}}, 
                                        @{n='EmailProfile';e={$EmailProfile}},
                                        @{n='Subject';e={$Subject}},
                                        @{n='Body';e={$Body.ToString()}},
                                        @{n='Attachment';e={$Attachment}}, 
                                        @{n='Log';e={$Log}},
                                        @{n='Guid';e={$Guid}})

    if($Attachment) {
        $AttachmentFile.Dispose()
    }  
}