﻿param (
    [Switch]$Impersonation
)
[Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
[Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"
$ToolsFolder = $Dir + "\AX-Tools"
Import-Module $ModuleFolder\AX-Tools.psm1 -DisableNameChecking

[System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty

function Validate-User
{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "SELECT UserName FROM [dbo].[AXTools_UserAccount] WHERE [USERNAME] = '$UserName'"
    $Conn.Open()
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $UserExists = $Cmd.ExecuteScalar()
    $Conn.Close()

    if($UserExists) {
        return $true
    }
    else {
        return $false
    }
}

function Delete-User
{
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "DELETE FROM [dbo].[AXTools_UserAccount] WHERE [USERNAME] = '$UserName'"
    $Conn.Open()
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery() | Out-Null
}

function Insert-User
{
    $SecureStringAsPlainText = $Credential.Password | ConvertFrom-SecureString
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Query = "INSERT INTO [dbo].[AXTools_UserAccount] ([ID],[USERNAME],[PASSWORD])
                VALUES ('$ID','$UserName','$SecureStringAsPlainText')"
    $Conn.Open()
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery() | Out-Null
    $Conn.Close()

    if($Impersonation) {
        [xml]$ConfigFile = Get-Content "$ModuleFolder\AX-Settings.xml"
        $ConfigFile.Settings.Database.Impersonation.UserName = $UserName
        $ConfigFile.Settings.Database.Impersonation.Password = $SecureStringAsPlainText.ToString()
        $ConfigFile.Save("$ModuleFolder\AX-Settings.xml")
    }
}

$Credential = Get-Credential -Message "Credentials <DOMAIN\Username or user@emailserver.com>" -ErrorAction SilentlyContinue

if ($Credential.UserName -ne $null ) {
    $BSTRBC = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
    $Root = "LDAP://" + ([ADSI]"").distinguishedName
    $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$Credential.UserName,[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTRBC))
    if ($Domain.Name -eq $null)
    {
        Write-Host "This is not a domain account."
        $ID = "$($Credential.UserName.Split('@')[0])"
        $UserName = "$($Credential.UserName)"
        if(Validate-User) {
            write-host "Username already exists. Updating current credentials."
            Delete-User
            Insert-User
        }
        else {
            Insert-User
        }

    }
        else
    {
        write-host "Domain successfully authenticated."
        $ID = "$($Credential.UserName.Split('\')[1].ToUpper())"
        $UserName = "$((([ADSI]`"").dc).toUpper())\$($Credential.UserName.Split('\')[1])"

        if(Validate-User) {
            write-host "Username already exists. Updating current credentials."
            Delete-User
            Insert-User
        }
        else {
            Insert-User
        }
    }
}
else {
    Write-Warning 'Canceled.'
}
<#
$VerbosePreference = "SilentlyContinue"

$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString

$PlainPassword = "P@ssw0rd"
$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force

$UserName = "Domain\User"
$Credentials = New-Object System.Management.Automation.PSCredential `
      -ArgumentList $UserName, $SecurePassword 

$PlainPassword = $Credentials.GetNetworkCredential().Password

$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

$SecureStringAsPlainText = $SecurePassword | ConvertFrom-SecureString 

$SecureString = $SecureStringAsPlainText  | ConvertTo-SecureString 

#>