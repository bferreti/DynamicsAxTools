﻿# .DISCLAIMER
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

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"

function Import-ConfigFile
{
param(
    [string]$ScriptName
)
    if($ScriptName) { $ScriptName = "$ScriptName|General" } else { $ScriptName = "General" } 
    if(Test-Path "$ModuleFolder\AX-Settings.xml") {
        # Import settings from config file
        [xml]$ConfigFile = Get-Content "$ModuleFolder\AX-Settings.xml"
        $PSObject = New-Object PSObject
        foreach ($Object in @($ConfigFile.DynamicsAxTools)) {
            foreach ($Property in @($Object.Setting | Where {$_.Module -Match "$ScriptName"})) {
                if($Property.Value -match 'true') {
                    $PSObject | Add-Member NoteProperty $Property.Key $([boolean]$true)
                }
                elseif($Property.Value -match 'false') {
                    $PSObject | Add-Member NoteProperty $Property.Key $([boolean]$false)
                }
                else {
                    $PSObject | Add-Member NoteProperty $Property.Key $Property.Value
                }
            }
        }
    }
    else {
        Write-Warning "Configuration file does not exists."
    }
    return $PSObject
}

function Check-Folder
{
param(
    [string]$Path
)
    if(!(Test-Path($Path))) {
        New-Item -ItemType Directory -Force -Path $Path | Out-Null
    }
}

function Get-ConnectionString 
{
[CmdletBinding()]
param (
    [String]$ApplicationName
)
    if($ApplicationName -eq '') { $ApplicationName = 'Ax Powershell Tools' }
    $ConfigFile =  Import-ConfigFile
    $ParamDBServer = $ConfigFile.DBServer
    $ParamDBName = $ConfigFile.DBName
    $ParamUserName = $ConfigFile.UserName
    $ParamPassword = $ConfigFile.Password
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
        $Query = "SELECT TOP 1 UserName, Password FROM [dbo].[AXTools_UserAccount] WHERE [ID] = '$($Account)'"
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
			$SqlServer.ConnectionContext.StatementTimeout = 0
            $SqlServer.ConnectionContext.Connect()
        }
        else {
            $SqlConn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection
            $SqlConn.ServerInstance = $DBServer
            $SqlConn.DatabaseName = $DBName
            $SqlConn.ApplicationName = $ApplicationName
            $SqlServer = New-Object Microsoft.SqlServer.Management.SMO.Server($SqlConn)
			$SqlServer.ConnectionContext.StatementTimeout = 0
            $SqlServer.ConnectionContext.Connect()
        }
    }
    catch {
        Write-Host "Failed to connect to AX Database. $($_.Exception.Message)"
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
    }
    $BCopy.WriteToServer($DataTable)
    $Conn.Close()
}

function Write-EncryptedString
{
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

function Read-EncryptedString
{
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
    $TLogStamp = (Get-Date -DisplayHint Time)
    $ExecLog = New-Object -TypeName System.Object
    $ExecLog | Add-Member -Name CreatedDateTime -Value $TLogStamp -MemberType NoteProperty
    $ExecLog | Add-Member -Name Guid -Value $(if($Global:Guid) {$Global:Guid} else {'0'}) -MemberType NoteProperty
    $ExecLog | Add-Member -Name Log -Value $LogData.Trim() -MemberType NoteProperty
    SQL-BulkInsert 'AXTools_ExecutionLog' $ExecLog
}

function UpdateMsiStatus
{
[CmdletBinding()]
param (
    [string]$Status
)
    $Conn = Get-ConnectionString
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
    $Conn = Get-ConnectionString
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
    $Conn = Get-ConnectionString
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function Set-SQLUpdate
{
[CmdletBinding()]
param (
    [String]$Query
)
    $Conn = Get-ConnectionString 
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
    $Conn = Get-ConnectionString
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
    $Conn = Get-ConnectionString
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
    if(-not [string]::IsNullOrEmpty($SMTPTo)) {$SMTPTo.Split(';') | % { $SMTPMessage.To.Add($_.Trim()) }} else { break }
    if(-not [string]::IsNullOrEmpty($SMTPCC)) {$SMTPCC.Split(';') | % { $SMTPMessage.CC.Add($_.Trim()) }}
    if(-not [string]::IsNullOrEmpty($SMTPBCC)) {$SMTPBCC.Split(';') | % { $SMTPMessage.Bcc.Add($_.Trim()) }}
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
        $Log = $_.Exception.Message.ToString()
    }

    SQL-BulkInsert AXTools_EmailLog @($SMTPMessage | Select @{n='Sent';e={[int]$Sent}}, 
                                        @{n='EmailProfile';e={$EmailProfile}},
                                        @{n='Subject';e={$Subject}},
                                        @{n='Body';e={[String]$Body}},
                                        @{n='Attachment';e={$Attachment}}, 
                                        @{n='Log';e={$Log}},
                                        @{n='Guid';e={$Guid}})

    if($Attachment) {
        $AttachmentFile.Dispose()
    }  
}

function Get-HtmlOpen
{
[CmdletBinding()]
param (
	[String]$Title,
	[Switch]$SimpleHTML,
    [Switch]$AxReport,
    [Switch]$AxSummary
)
	
$CurDate = Get-Date -f "MMM d, yyyy hh:mm tt"

if($SimpleHTML) {
$ReportHtml = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
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
$ReportHtml = @"
MIME-Version: 1.0
Content-Type: multipart/related; boundary="PART"; type="text/html"

--PART
Content-Type: text/html; charset=us-ascii
Content-Transfer-Encoding: 7bit

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
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
$ReportHtml = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"> 
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
$ReportHtml = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
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
	return $ReportHtml
}

function Get-HtmlClose
{
Param(
	[string]$Footer,
    [Switch]$AxReport,
    [Switch]$AxSummary    	
)

$Footer = "Date: {0} | UserName: {1}\{2} | {3}" -f $(Get-Date),$env:UserDomain,$env:UserName,$Footer

if($AxReport) {
$ReportHtml = @"
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
$ReportHtml = @"
<div class="section">
    <hr />
    <div class="footer">$Footer</div>
</div>
    
</body>
</html>
"@
}

else {
$ReportHtml = @"
<div class="section">
    <hr />
    <div class="Footer">$Footer</div>
</div></div></div>
    
</body>
</html>

"@
}
	Write-Output $ReportHtml
}

function Get-HtmlContentOpen
{
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
	$ReportHtml = @"
<div class="section">
    <div class="header">
        <a name="$($Header)">$($Header)</a> (<a id="show_$JavaScriptRdm" href="javascript:void(0);" onclick="show('$JavaScriptRdm');" style="color: #ffffff;">Show</a><a id="hide_$JavaScriptRdm" href="javascript:void(0);" onclick="hide('$JavaScriptRdm');" style="color: #ffffff; display:none;">Hide</a>)
    </div>
    <div class="content" id="$JavaScriptRdm" style="display:none;background-color:$($bgColorCode);"> 
"@	
}
else {
	$ReportHtml = @"
<div class="section">
    <div class="header">
        <a name="$($Header)">$($Header)</a>
    </div>
    <div class="content" style="background-color:$($bgColorCode);"> 
"@
}
	return $ReportHtml
}

function Get-HtmlContentClose
{
	$ReportHtml = @"
</div>
</div>
"@
	return $ReportHtml
}

function Get-HtmlAddNewLine
{
	$ReportHtml = @"
<br>
"@
	return $ReportHtml
}

function Get-HtmlContentTable
{
param(
	[Array]$ObjectArray, 
	[Switch]$Fixed, 
	[String]$GroupBy,
    [String]$Title,
    [String]$Style
)
	if ($GroupBy -eq '') {
		if($Title) { $ReportHtml = "<h2>$Title</h2>" }
        $ReportHtmlArr = $ObjectArray | ConvertTo-Html -Fragment
		$ReportHtmlArr = $ReportHtmlArr -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', ""
		$ReportHtmlArr = $ReportHtmlArr -replace "<tr>(.*)<td>Green</td></tr>","<tr class=`"green`">`$+</tr>"
		$ReportHtmlArr = $ReportHtmlArr -replace "<tr>(.*)<td>Yellow</td></tr>","<tr class=`"yellow`">`$+</tr>"
    	$ReportHtmlArr = $ReportHtmlArr -replace "<tr>(.*)<td>Red</td></tr>","<tr class=`"red`">`$+</tr>"
        $ReportHtmlArr = $ReportHtmlArr -replace "<tr>(.*)<td>Orange</td></tr>","<tr class=`"orange`">`$+</tr>"
        $ReportHtmlArr = $ReportHtmlArr -replace "<tr>(.*)<td>LightRed</td></tr>","<tr class=`"lightred`">`$+</tr>"
        $ReportHtmlArr = $ReportHtmlArr -replace "<tr>(.*)<td>LightGreen</td></tr>","<tr class=`"lightgreen`">`$+</tr>"
        $ReportHtmlArr = $ReportHtmlArr -replace "<tr>(.*)<td>LightYellow</td></tr>","<tr class=`"lightyellow`">`$+</tr>"
		$ReportHtmlArr = $ReportHtmlArr -replace "<tr>(.*)<td>Odd</td></tr>","<tr class=`"odd`">`$+</tr>"
		$ReportHtmlArr = $ReportHtmlArr -replace "<tr>(.*)<td>Even</td></tr>","<tr class=`"even`">`$+</tr>"
		$ReportHtmlArr = $ReportHtmlArr -replace "<tr>(.*)<td>None</td></tr>","<tr>`$+</tr>"
		$ReportHtmlArr = $ReportHtmlArr -replace '<th>RowColor</th>', ''
        $ReportHtml += $ReportHtmlArr

		if ($Fixed.IsPresent) {	$ReportHtml = $ReportHtml -replace '<table>', '<table class="fixed">' }
        if ($Style) { $ReportHtml = $ReportHtml -replace '<table>', "<table class=""$Style"">" }
	}
	else {
		$NumberOfColumns = ($ObjectArray | Get-Member -MemberType NoteProperty  | select Name).Count
		$Groupings = @()
		$ObjectArray | select $GroupBy -Unique  | sort $GroupBy | foreach { $Groupings += [String]$_.$GroupBy}
		if ($Fixed.IsPresent) {	$ReportHtml = '<table class="fixed">' }
		else { $ReportHtml = "<table>" }
		$GroupHeader = $ObjectArray | ConvertTo-Html -Fragment 
		$GroupHeader = $GroupHeader -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', "" -replace '<table>', "" -replace '</table>', "" -replace "<td>.+?</td>" -replace "<tr></tr>", ""
		$GroupHeader = $GroupHeader -replace '<th>RowColor</th>', ''
		$ReportHtml += $GroupHeader
		foreach ($Group in $Groupings) {
			$ReportHtml += "<tr><td colspan=`"$NumberOfColumns`" class=`"groupby`">$Group</td></tr>"
			$GroupBody = $ObjectArray | where { [String]$($_.$GroupBy) -eq $Group } | select * -ExcludeProperty $GroupBy | ConvertTo-Html -Fragment
			$GroupBody = $GroupBody -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', "" -replace '<table>', "" -replace '</table>', "" -replace "<th>.+?</th>" -replace "<tr></tr>", "" -replace '<tr><td>', "<tr><td></td><td>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Green</td></tr>","<tr class=`"green`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Yellow</td></tr>","<tr class=`"yellow`">`$+</tr>"
    		$GroupBody = $GroupBody -replace "<tr>(.*)<td>Red</td></tr>","<tr class=`"red`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Odd</td></tr>","<tr class=`"odd`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Even</td></tr>","<tr class=`"even`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>None</td></tr>","<tr>`$+</tr>"
			$ReportHtml += $GroupBody
		}
		$ReportHtml += "</table>" 
	}
	$ReportHtml = $ReportHtml -replace 'URL01', '<a href="'
	$ReportHtml = $ReportHtml -replace 'URL02', '">'
	$ReportHtml = $ReportHtml -replace 'URL03', '</a>'
	
	if ($ReportHtml -like "*<tr>*" -and $ReportHtml -like "*odd*" -and $ReportHtml -like "*even*") {
			$ReportHtml = $ReportHtml -replace "<tr>",'<tr class="header">'
	}
	
	return $ReportHtml
}

function Get-HtmlContentText 
{
param(
	$Heading,
	$Detail
)
$ReportHtml = @"
<table><tbody>
	<tr>
	<th class="content">$Heading</th>
	<td class="content">$($Detail)</td>
	</tr>
</tbody></table>
"@
    $ReportHtml = $ReportHtml -replace 'URL01', '<a href="'
    $ReportHtml = $ReportHtml -replace 'URL02', '">'
    $ReportHtml = $ReportHtml -replace 'URL03', '</a>'
    return $ReportHtml
}

function Set-RowColor
{
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
param (
	$ChartObject,
	$ChartData
)
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
	$Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
	$Chart.Width = $ChartObject.Size.Width
	$Chart.Height = $ChartObject.Size.Height
	$Chart.Left = $ChartObject.Size.Left
	$Chart.Top = $ChartObject.Size.Top
	$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
	$Chart.ChartAreas.Add($ChartArea)
	[void]$Chart.Series.Add("Data")
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
	}
	$Title = new-object System.Windows.Forms.DataVisualization.Charting.Title
	[Void]$Chart.Titles.Add($Title)
	$Chart.Titles[0].Text = $ChartObject.Title
	$tempfile = (Join-Path $env:TEMP $ChartObject.Title.replace(' ', '')) + ".png"
	if ((test-path $tempfile)) { Remove-Item $tempfile -Force }
	$Chart.SaveImage($tempfile, "png")
	$Base64Chart = [Convert]::ToBase64String((Get-Content $tempfile -Encoding Byte))
	$HTMLCode = '<IMG SRC="data:image/gif;base64,' + $Base64Chart + '" ALT="' + $ChartObject.Title + '">'
	return $HTMLCode
}

function New-HTMLPieChartObject
{
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

function New-HTMLPieChart
{
param(
    $PieChartObject,
    $PieChartData
)
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
	$Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
	$Chart.Width = $PieChartObject.Size.Width
	$Chart.Height = $PieChartObject.Size.Height
	$Chart.Left = $PieChartObject.Size.Left
	$Chart.Top = $PieChartObject.Size.Top
	$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
	$Chart.ChartAreas.Add($ChartArea) 
	[void]$Chart.Series.Add("Data") 
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
	$Title = new-object System.Windows.Forms.DataVisualization.Charting.Title 
	[Void]$Chart.Titles.Add($Title) 
	$Chart.Titles[0].Text = $PieChartObject.Title
	$tempfile = (Join-Path $env:TEMP $PieChartObject.Title.replace(' ','') ) + ".png"
	if ((test-path $tempfile)) {Remove-Item $tempfile -Force}
	$Chart.SaveImage( $tempfile  ,"png")
	$Base64Chart = [Convert]::ToBase64String((Get-Content $tempfile -Encoding Byte))
	$HTMLCode = '<IMG SRC="data:image/gif;base64,' + $Base64Chart + '" ALT="' + $PieChartObject.Title + '">'
	return $HTMLCode 
}

function Get-HTMLColumn1of2
{
	$ReportHtml = '<div class="first column">'
	return $ReportHtml
}

function Get-HTMLColumn2of2
{
	$ReportHtml = '<div class="second column">'
	return $ReportHtml
}

function Get-HTMLColumnClose
{
	$ReportHtml = '</div>'
	return $ReportHtml
}

function Test-AosServices
{
[CmdletBinding()]
param (
	[String]$AxEnvironment,
	[String]$AosInstance,
    [String]$ServerType,
	[Switch]$Start
)
    $Query = "SELECT SERVERNAME FROM AXTools_Servers WHERE ENVIRONMENT = '$AXEnvironment' AND SERVERTYPE IN ($ServerType) AND ACTIVE = '1'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$(Get-ConnectionString))
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $Servers = New-Object System.Data.DataSet
    $TotalServers = $Adapter.Fill($Servers)

    $Stopped = 0
    $Running = 0

    if($TotalServers -gt 0) {
        $Query = "SELECT LocalAdminUser FROM [AXTools_Environments] WHERE ENVIRONMENT = '$AXEnvironment'"
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$(Get-ConnectionString))
        $LocalAdminAccount = $Cmd.ExecuteScalar()
        if(![String]::IsNullOrEmpty($LocalAdminAccount)) { $LocalAdminAccount = Get-UserCredentials -Account $LocalAdminAccount }

        foreach($Server in $Servers.Tables[0]) {
            if(Test-Connection $Server.ServerName -Count 1 -Quiet) {
                if(![String]::IsNullOrEmpty($LocalAdminAccount) -and $Server.ServerName -ne $env:COMPUTERNAME) {
                    $Services = Get-WmiObject -Class Win32_Service -ComputerName $($Server.ServerName) -Credential $LocalAdminAccount -ea 0 | Where-Object { $_.DisplayName -like "Microsoft Dynamics AX Object Server*" }
                    if($Services) { 
                        foreach($Service in $Services) {
                            if($Service.State -match 'Stopped') {
                                if($Start) { $Service.StartService() }
                                Write-Log "ERROR: AOS Check $($Service.SystemName) | $($Service.Name) | $($Service.State)"
                                $Stopped++
                            }
                            else {
                                $Running++
                            }
                        }
                    }
                }
                else {
                    $Services = Get-WmiObject -Class Win32_Service -ComputerName $($Server.ServerName) -ea 0 | Where-Object { $_.DisplayName -like "Microsoft Dynamics AX Object Server*" }
                    if($Services) { 
                        foreach($Service in $Services) {
                            if($Service.State -match 'Stopped') {
                                if($Start) { $Service.StartService() }
                                Write-Log "ERROR: AOS Check $($Service.SystemName) | $($Service.Name) | $($Service.State)"
                                $Stopped++
                            }
                            else {
                                $Running++
                            }
                        }
                    }
                }
            }
            else {
                $Stopped++
                Write-Log "ERROR: AOS Check - Server unavailable $($Server.ServerName)."
            }
        }
        Write-Log "AOS Check - Servers $TotalServers - Running $Running | Failed $Stopped."
    }
    else {
        Write-Log "ERROR: AOS Check environment $AXEnvironment not found."
    }
}

function Test-PerfmonSet
{
[CmdletBinding()]
param (
	[String]$AxEnvironment,
	[Switch]$Start
)
    $SettingsXml = Import-ConfigFile -ScriptName 'AxReport'
    $Query = "SELECT SERVERNAME FROM AXTools_Servers WHERE ENVIRONMENT = '$AXEnvironment' AND ACTIVE = '1'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$(Get-ConnectionString))
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $Servers = New-Object System.Data.DataSet
    $TotalServers = $Adapter.Fill($Servers)
    if($TotalServers -gt 0) {
        $Query =   "SELECT LocalAdminUser FROM [AXTools_Environments] WHERE ENVIRONMENT = '$AXEnvironment'"
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$(Get-ConnectionString))
        $LocalAdminAccount = $Cmd.ExecuteScalar()
        if(![String]::IsNullOrEmpty($LocalAdminAccount)) { $LocalAdminAccount = Get-UserCredentials -Account $LocalAdminAccount }
        foreach($Server in $Servers.Tables[0]) {
            if(Test-Connection $Server.ServerName -Count 1 -Quiet) {
                if(![String]::IsNullOrEmpty($LocalAdminAccount) -and $Server.ServerName -ne $env:COMPUTERNAME) {
                    Invoke-Command -ComputerName $Server.ServerName -Credential $LocalAdminAccount -ArgumentList $($SettingsXml.PerfmonName), $Server.ServerName, $Start -ScriptBlock {
                        Param($DataCollectorName, $ServerName, $Start)
                        try {
                            $DataCollectorSet = New-Object -COM Pla.DataCollectorSet
                            $DataCollectorSet.Query("$DataCollectorName", $ServerName)
                            if($DataCollectorSet.Status -eq 0) {
                                if($Start) { $DataCollectorSet.Start($false) }
                                $Msg = "ERROR: Perfmon Check $ServerName | $DataCollectorName | Stopped."
                            }
                        }
                        catch {
                            $Msg = "ERROR: Perfmon $($ServerName) - $($_.exception.message)"
                        }
                        return $Msg
                    } -OutVariable Msg
                    if($Msg) { Write-Log $Msg }
                }
                else {
                    Invoke-Command -ComputerName $Server.ServerName -ArgumentList $($SettingsXml.PerfmonName), $Server.ServerName, $Start -ScriptBlock {
                        Param($DataCollectorName, $ServerName, $Start)
                        try {
                            $DataCollectorSet = New-Object -COM Pla.DataCollectorSet
                            $DataCollectorSet.Query("$DataCollectorName", $ServerName)
                            if($DataCollectorSet.Status -eq 0) {
                                if($Start) { $DataCollectorSet.Start($false) }
                                $Msg = "ERROR: Perfmon Check $ServerName | $DataCollectorName | Stopped."
                            }
                        }
                        catch {
                            $Msg = "ERROR: Perfmon $($ServerName) - $($_.exception.message)"
                        }
                        return $Msg
                    } -OutVariable Msg
                    if($Msg) { Write-Log $Msg }
                }
            }
        }
        Write-Log "Perfmon Check Completed. Env. $AXEnvironment."
    }
    else {
        Write-Log "ERROR: Perfmon Check environment $AXEnvironment not found."
    }
}

function New-Popup 
{
<#
.Synopsis
Display a Popup Message
.Description
This command uses the Wscript.Shell PopUp method to display a graphical message
box. You can customize its appearance of icons and buttons. By default the user
must click a button to dismiss but you can set a timeout value in seconds to 
automatically dismiss the popup. 

The command will write the return value of the clicked button to the pipeline:
  OK     = 1
  Cancel = 2
  Abort  = 3
  Retry  = 4
  Ignore = 5
  Yes    = 6
  No     = 7

If no button is clicked, the return value is -1.
.Example
PS C:\> new-popup -message "The update script has completed" -title "Finished" -time 5

This will display a popup message using the default OK button and default 
Information icon. The popup will automatically dismiss after 5 seconds.
.Outputs
Null   = -1
OK     = 1
Cancel = 2
Abort  = 3
Retry  = 4
Ignore = 5
Yes    = 6
No     = 7
#>

Param (
[Parameter(Position=0,Mandatory=$True,HelpMessage="Enter a message for the popup")]
[ValidateNotNullorEmpty()]
[string]$Message,
[Parameter(Position=1,Mandatory=$True,HelpMessage="Enter a title for the popup")]
[ValidateNotNullorEmpty()]
[string]$Title,
[Parameter(Position=2,HelpMessage="How many seconds to display? Use 0 require a button click.")]
[ValidateScript({$_ -ge 0})]
[int]$Time=0,
[Parameter(Position=3,HelpMessage="Enter a button group")]
[ValidateNotNullorEmpty()]
[ValidateSet("OK","OKCancel","AbortRetryIgnore","YesNo","YesNoCancel","RetryCancel")]
[string]$Buttons="OK",
[Parameter(Position=4,HelpMessage="Enter an icon set")]
[ValidateNotNullorEmpty()]
[ValidateSet("Stop","Question","Exclamation","Information" )]
[string]$Icon="Information"
)
    Switch ($Buttons) {
        "OK"               {$ButtonValue = 0}
        "OKCancel"         {$ButtonValue = 1}
        "AbortRetryIgnore" {$ButtonValue = 2}
        "YesNo"            {$ButtonValue = 4}
        "YesNoCancel"      {$ButtonValue = 3}
        "RetryCancel"      {$ButtonValue = 5}
    }

    #set an integer value for Icon type
    Switch ($Icon) {
        "Stop"        {$iconValue = 16}
        "Question"    {$iconValue = 32}
        "Exclamation" {$iconValue = 48}
        "Information" {$iconValue = 64}
    }

    #create the COM Object
    Try {
        $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
        #Button and icon type values are added together to create an integer value
        $Answer = $wshell.Popup($Message,$Time,$Title,$ButtonValue+$iconValue)
        return $Answer
    }
    Catch {
        #You should never really run into an exception in normal usage
        Write-Warning "Failed to create Wscript.Shell COM object"
        Write-Warning $_.exception.message
    }
}