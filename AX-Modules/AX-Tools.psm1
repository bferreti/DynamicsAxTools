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
    if($ApplicationName -eq '') { $ApplicationName = 'Ax Powershell Tools' }
    $ConfigFile = Load-ConfigFile
    $ParamDBServer = $ConfigFile.Settings.Database.DBServer
    $ParamDBName = $ConfigFile.Settings.Database.DBName
    $ParamUserName = $ConfigFile.Settings.Database.UserName
    $ParamPassword = $ConfigFile.Settings.Database.Password
    if($ParamUserName) {
        $UserPassword = Read-EncryptedString -InputString $ParamPassword -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)"
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

function Get-UserCredentials
{
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]$Account
)
    [System.Management.Automation.Credential()]$UserCreds = [System.Management.Automation.PSCredential]::Empty
    if($Account) {
        $Query = "SELECT UserName, Password FROM [dbo].[AXTools_UserAccount] WHERE [ID] = '$($Account)'"
        $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter($Query,$(Get-ConnectionString))
        $UserAcct = New-Object System.Data.DataSet
        $Adapter.Fill($UserAcct) | Out-Null
        $UserPwd = Read-EncryptedString -InputString $UserAcct.Tables.Password -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)"
        $secureUserPwd = $UserPwd | ConvertTo-SecureString -AsPlainText -Force 
        $UserCreds = New-Object System.Management.Automation.PSCredential -ArgumentList $UserAcct.Tables.UserName, $secureUserPwd
        return $UserCreds
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
    if($ApplicationName -eq '') { $ApplicationName = 'Ax Powershell Tools' }
    try {
        if($SQLAccount) {
            $SqlCredential = Get-UserCredentials $($SQLAccount)
            #$Query = "SELECT UserName, Password FROM [dbo].[AXTools_UserAccount] WHERE [ID] = '$($SQLAccount)'"
            #$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter($Query,$(Get-ConnectionString))
            #$UserAccount = New-Object System.Data.DataSet
            #$Adapter.Fill($UserAccount) | Out-Null
            #$UserPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($UserAccount.Tables.Password | ConvertTo-SecureString)))
            #$secureUserPassword = $UserPassword | ConvertTo-SecureString -AsPlainText -Force 
            #$SqlCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserAccount.Tables.UserName, $secureUserPassword
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
    #Write-Output @(,($DataTable))
    $Conn = Get-ConnectionString
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
    [String]$DTKey,
    [Switch]$Compress
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
	if ($Compress) { $InputEncodingStream = (New-Object System.IO.Compression.GZipStream($InputDataStream, ([IO.Compression.CompressionMode]::Compress), $True)) }
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
    return $OutputDataAsString
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
        $SMTPPassword = Read-EncryptedString -InputString $($Table.Tables.Password) -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)"
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

    SQL-BulkInsert AXTools_EmailLog @($SMTPMessage | Select @{n='Sent';e={[int]$Sent}}, 
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

function Get-HtmlOpen {
<#
	.SYNOPSIS
		Header HTML for report
    .PARAMETER TitleText
		The title of the report
    .PARAMETER SimpleHTML
		CSS with basic formatting
    .PARAMETER AxReport
		CSS for AX Report formatting (mht)
    .PARAMETER AxSummary
		CSS for AX Report Summary email
#>
[CmdletBinding()]
param (
	[String]$Title,
	[Switch]$SimpleHTML,
    [Switch]$AxReport,
    [Switch]$AxSummary
)
	
$CurrentDate = Get-Date -format "MMM d, yyyy hh:mm tt"

if($SimpleHTML) {
$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>$($Title)</title>
      <style type=text/css>
      *{font-family:Segoe UI Symbol;margin-top:4px;margin-bottom:4px}
       body{margin:8px 5px}
       h1{color:#000;font-size:18pt;text-align:left;text-decoration:underline}
       h2{color:#000;font-size:16pt;text-align:left;text-decoration:underline}
       h3{color:#000;font-size:14pt;text-align:left;text-decoration:underline}
       hr{background:#337e94;height:4px}
       table{border:1px solid #000033;border-collapse:collapse;margin:0px;margin-left:10px}
       td{border:1px solid #000033;font-size:8pt;font-weight:550;padding-left:3px;padding-right:15px;}
       th{background:#337e94;border:1px solid #000033;color:#FFF;font-size:9pt;font-weight:bold;margin:0px;padding:2px;text-align:center}
       table.fixed{table-layout:fixed}
       tr:hover{background:#808080}
       div.header{color:black;font-size:12pt;font-weight:bold;background-color:transparent;margin-bottom:4px;margin-top:12px}
       div.footer{padding-right:5em;text-align:right;font-size:9pt;padding:2px}
       div.reportdate{font-size:12pt;font-weight:bold}
       div.reportname{font-size:16pt;font-weight:bold}
       div.section{width:auto}
       .header{background:#616a6b;color:#f7f9f9}
       .odd{background:#d5d8dc}
       .even{background:#f7f9f9}
        .green {background-color:#a1cda4;}
        .yellow {background-color:#fffab1;}
        .red {background-color:#FF0000;}
        .orange {background-color:#FFA500}
        .lightred {background-color:#FFA39F}
        .lightyellow {background-color:#FFFFA9}
        .lightgreen {background-color:#D7FFCC}
       .none{background:#FFF}
       </style>
</head>
<div class="section">
    <div class="ReportName">$($Title)</div>
    <hr/>
</div>
"@
}

elseif($AxReport) {
$Report = @"
MIME-Version: 1.0
Content-Type: multipart/related; boundary="PART"; type="text/html"

--PART
Content-Type: text/html; charset=us-ascii
Content-Transfer-Encoding: 7bit

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head><title>$($Title)</title>
<style type="text/css">
    * {margin: 0px;font-family: sans-serif;font-size: 8pt;}
    body {margin: 8px 5px 8px 5px;}
    hr {height: 4px;background-color: #337e94;border: 0px;}
    table {table-layout: auto;width: 100%;border-collapse: collapse;}
    th {vertical-align: top;text-align: left;padding: 2px 5px 2px 5px;}
    td {vertical-align: top;padding: 2px 5px 2px 5px;border-top: 1px solid #bbbbbb;}
    div.section {padding-bottom: 12px;} 
    div.header {border: 1px solid #bbbbbb;padding: 4px 5em 0px 5px;margin: 0px 0px -1px 0px;height: 2em;width: 95%;font-weight: bold;color: #ffffff;background-color: #337e94;}
    div.content {border: 1px solid #bbbbbb;padding: 4px 0px 5px 11px;margin: 0px 0px -1px 0px;width: 95%;color: #000000;background-color: #f9f9f9;}
    div.reportname {font-size: 16pt;font-weight: bold;}
    div.repordate {font-size: 12pt;font-weight: bold;}
    div.footer {padding-right: 5em;text-align: right;}
    table.fixed {table-layout: fixed;}
    th.content {border-top: 1px solid #bbbbbb;width: 25%;}
    td.content {width: 75%;}
    td.groupby {border-top: 3px double #bbbbbb;}
    .green {background-color: #a1cda4;}
    .yellow {background-color: #fffab1;}
    .red {background-color: #f5a085;}
    .odd {background-color: #D5D8DC;}
    .even {background-color: #F7F9F9;}
    .header {background-color: #616A6B; color: #F7F9F9;}
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
    <div class="ReportName">$($Title) - $((Get-Date).AddDays(-1) | Get-Date -Format "D")</div>
    <hr/>
</div>
"@
}

elseif($AxSummary) {
$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"> 
<html>
<head><title>$($Title)</title> 
    <style type="text/css"> * {margin: 0px;font-family: sans-serif;font-size: 8pt;}
    body {margin: 8px 5px 8px 5px;}
    hr {height: 4px;background-color: #337e94;border: 0px;}
    table {table-layout: auto;width: 100%;border-collapse: collapse;}
    th {vertical-align: top;text-align: left;padding: 2px 5px 2px 5px;}
    td {vertical-align: top;padding: 2px 5px 2px 5px;border-top: 1px solid #bbbbbb;}
    div.section {padding-bottom: 12px;}
    div.header {border: 1px solid #bbbbbb;margin: 0px 0px -1px 0px;height: 2em;width: 95%;font-weight: bold ;color: #ffffff;background-color: #337e94;}
    div.content {border: 1px solid #bbbbbb;margin: 0px 0px -1px 0px;width: 95%;color: #000000;background-color: #f9f9f9;}
    div.reportname {font-size: 16pt;font-weight: bold;}
    div.footer {padding-right: 5em;text-align: right;}
    table.fixed {table-layout: fixed;}
    th.content {border-top: 1px solid #bbbbbb;width: 25%;}
    td.content {width: 75%;}td.groupby {border-top: 3px double #bbbbbb;}
    .green {background-color: #a1cda4;}
    .yellow {background-color: #fffab1;}
    .red {background-color: #f5a085;}
    .odd {background-color: #D5D8DC;}
    .even {background-color: #F7F9F9;}
    .header {background-color: #616A6B;color: #F7F9F9;}
    </style> 
</head>

<div class="section"> 
    <div class="reportname">$($Title)</div> 
    <hr/>
    <br></br> 
</div>
"@
}

else {
$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>$($Title)</title>
    <style type="text/css">*{font:8pt sans-serif;margin:0px}
    body{margin:8px 5px 8px 5px}
    hr{background:#337e94;border:0px;height:4px}
    table{border-collapse:collapse;table-layout:auto;width:100%}
    td{border-top:1px solid #bbbbbb;padding:2px 5px 2px 5px;vertical-align:top}
    th{padding:2px 5px 2px 5px;text-align:left;vertical-align:top}
    div.column{float:left;overflow-y:auto;width:100%}
    div.content{background:#f9f9f9;border:1px solid #bbbbbb;color:#000;margin:0px 0px -1px 0px;padding:4px 0px 5px 11px;width:95%}
    div.first{border-right:1px grey solid;width:49%}
    div.footer{padding-right:5em;text-align:right}
    div.header{background:#337e94;border:1px solid #bbbbbb;color:#fff;font-weight:bold;height:2em;margin:0px 0px -1px 0px;padding:4px 5em 0px 5px;width:95%}
    div.reportdate{font-size:12pt;font-weight:bold}
    div.reportname{font-size:16pt;font-weight:bold}
    div.second{margin-left:10px;width:49%}
    div.section{padding-bottom:12px}
    table.fixed{table-layout:fixed}
    td.content{width:75%}
    td.groupby{border-top:3px double #bbbbbb}
    th.content{border-top:1px solid #bbbbbb;width:25%}
    .header{background:#616A6B;color:#F7F9F9}
    .odd{background:#d5d8dc}
    .even{background:#f7f9f9}
    .green {background-color: #a1cda4;}
    .yellow {background-color: #fffab1;}
    .red {background-color: #FF0000;}
    .orange {background-color:#FFA500}
    .lightred {background-color:#FFA39F}
    .lightyellow {background-color:#FFFFA9}
    .lightgreen {background-color:#D7FFCC}
    .none{background:#FFF}
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
    <div class="ReportName">$($Title) - $((Get-Date).AddDays(-1) | Get-Date -Format "D")</div>
    <hr/>
</div>
"@
}
	Return $Report
}

function Get-HtmlClose
{
<#
	.SYNOPSIS
		Close HTML for report
    .PARAMETER FooterTxt
		The footer of the report
    .PARAMETER AxReport
		CSS for AX Report formatting (mht)
    .PARAMETER AxSummary
		CSS for AX Report Summary email
#>
Param(
	[string]$Footer,
    [Switch]$AxReport,
    [Switch]$AxSummary    	
)

$Footer = "Date: {0} | UserName: {1}\{2} | {3}" -f $(Get-Date),$env:UserDomain,$env:UserName,$Footer

if($AxReport) {
$Report = @"
<div class="section">
    <hr />
    <div class="Footer">$Footer</div>
</div>
    
</body>
</html>

--PART-- 
"@
}

elseif($AxSummary) {
$Report = @"
<div class="section">
    <hr />
    <div class="footer">$Footer</div>
</div>
    
</body>
</html>
"@
}

else {
$Report = @"
<div class="section">
    <hr />
    <div class="Footer">$Footer</div>
</div></div></div>
    
</body>
</html>

"@
}
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
	[string]$Header, 
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
	$JavaScriptRdm = Get-Random
	$Report = @"
<div class="section">
    <div class="header">
        <a name="$($Header)">$($Header)</a> (<a id="show_$JavaScriptRdm" href="javascript:void(0);" onclick="show('$JavaScriptRdm');" style="color: #ffffff;">Show</a><a id="hide_$JavaScriptRdm" href="javascript:void(0);" onclick="hide('$JavaScriptRdm');" style="color: #ffffff; display:none;">Hide</a>)
    </div>
    <div class="content" id="$JavaScriptRdm" style="display:none;background-color:$($bgColorCode);"> 
"@	
}
else {
	$Report = @"
<div class="section">
    <div class="header">
        <a name="$($Header)">$($Header)</a>
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

function Get-HtmlAddNewLine {
<#
	.SYNOPSIS
		Add new line
#>	
	$Report = @"
<br>
"@
	Return $Report
}

function Get-HtmlContentTable {
<#
	.SYNOPSIS
		Creates an HTML table from an array of objects
	    .PARAMETER ObjectArray
			An array of objects
		.PARAMETER Fixed
		    fixes the html column width by the number of columns
		.PARAMETER GroupBy
		    The column to group the data. make sure this is first in the array
		.PARAMETER Title
		    Title
		.PARAMETER Style
		    Style
#>	
param(
	[Array]$ObjectArray, 
	[Switch]$Fixed, 
	[String]$GroupBy,
    [String]$Title,
    [String]$Style
)
	if ($GroupBy -eq '') {
		if($Title) { $Report = "<h2>$Title</h2>" }
        $ReportHtml = $ObjectArray | ConvertTo-Html -Fragment
		$ReportHtml = $ReportHtml -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', ""
		$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Green</td></tr>","<tr class=`"green`">`$+</tr>"
		$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Yellow</td></tr>","<tr class=`"yellow`">`$+</tr>"
    	$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Red</td></tr>","<tr class=`"red`">`$+</tr>"
        $ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Orange</td></tr>","<tr class=`"orange`">`$+</tr>"
        $ReportHtml = $ReportHtml -replace "<tr>(.*)<td>LightRed</td></tr>","<tr class=`"lightred`">`$+</tr>"
        $ReportHtml = $ReportHtml -replace "<tr>(.*)<td>LightGreen</td></tr>","<tr class=`"lightgreen`">`$+</tr>"
        $ReportHtml = $ReportHtml -replace "<tr>(.*)<td>LightYellow</td></tr>","<tr class=`"lightyellow`">`$+</tr>"
		$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Odd</td></tr>","<tr class=`"odd`">`$+</tr>"
		$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Even</td></tr>","<tr class=`"even`">`$+</tr>"
		$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>None</td></tr>","<tr>`$+</tr>"
		$ReportHtml = $ReportHtml -replace '<th>RowColor</th>', ''
        
        $Report += $ReportHtml

		if ($Fixed.IsPresent) {	$Report = $Report -replace '<table>', '<table class="fixed">' }
        if ($Style) { $Report = $Report -replace '<table>', "<table class=""$Style"">" }
	}
	else {
		$NumberOfColumns = ($ObjectArray | Get-Member -MemberType NoteProperty  | select Name).Count
		$Groupings = @()
		$ObjectArray | select $GroupBy -Unique  | sort $GroupBy | foreach { $Groupings += [String]$_.$GroupBy}
		if ($Fixed.IsPresent) {	$Report = '<table class="fixed">' }
		else { $Report = "<table>" }
		$GroupHeader = $ObjectArray | ConvertTo-Html -Fragment 
		$GroupHeader = $GroupHeader -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', "" -replace '<table>', "" -replace '</table>', "" -replace "<td>.+?</td>" -replace "<tr></tr>", ""
		$GroupHeader = $GroupHeader -replace '<th>RowColor</th>', ''
		$Report += $GroupHeader
		foreach ($Group in $Groupings) {
			$Report += "<tr><td colspan=`"$NumberOfColumns`" class=`"groupby`">$Group</td></tr>"
			$GroupBody = $ObjectArray | where { [String]$($_.$GroupBy) -eq $Group } | select * -ExcludeProperty $GroupBy | ConvertTo-Html -Fragment
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
			Heading text or picture
		.PARAMETER Detail
		     Some additional info
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

function Set-RowColor {
<#
	.SYNOPSIS
		Adds a RowColor field to each row in the array
	    .PARAMETER ObjectArray
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
	$ObjectArray, 
	$Green, 
	$Yellow, 
	$Red,
	[switch]$Alternating 
) 
    if ($Alternating) {
		$ColoredArray = $ObjectArray | Add-Member -MemberType ScriptProperty -Name RowColor -Value {
		if ((([array]::indexOf($ObjectArray,$this)) % 2) -eq 0) {'Odd'}
		if ((([array]::indexOf($ObjectArray,$this)) % 2) -eq 1) {'Even'}
		} -PassThru -Force | Select-Object *
	} else {
		$ColoredArray = $ObjectArray | Add-Member -MemberType ScriptProperty -Name RowColor -Value {
			if (Invoke-Expression $Green) {'Green'} 
			elseif (Invoke-Expression $Yellow) {'Yellow'}
			elseif (Invoke-Expression $Red) {'Red'} 
			else {'None'}
			} -PassThru -Force | Select-Object *
	}
	return $ColoredArray
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
}

function Get-HTMLColumn1of2
{
<#
	.SYNOPSIS
#>
	$report = '<div class="first column">'
	return $report
}

function Get-HTMLColumn2of2
{
<#
	.SYNOPSIS
#>
	$report = '<div class="second column">'
	return $report
}

function Get-HTMLColumnClose
{
<#
	.SYNOPSIS
#>
	$report = '</div>'
	return $report
}