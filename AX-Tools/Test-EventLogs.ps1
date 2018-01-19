Param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string]$ServerName,
    [System.Management.Automation.PSCredential]$Credential
)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"

Import-Module C:\Users\Administrator\Documents\GitHub\DynamicsAxTools\AX-Modules\AX-Tools.psm1 -DisableNameChecking

function Get-EventLogs
{
    try
    {
        if($Credential) {
            $EventLogs = Get-WinEvent –FilterHashtable @{LogName = 'Application', 'System'; Level = 2, 3; StartTime=$((Get-Date).AddDays(-1).Date)} -ComputerName $ServerName -Credential $Credential | 
                    Select @{n='LogName';e={$_.LogName}}, @{n='EntryType';e={($_.LevelDisplayName).ToString()}}, @{n='EventID';e={$_.ID}}, @{n='Source';e={$_.ProviderName}}, @{n='TimeGenerated';e={$_.TimeCreated}},  @{n='Message';e={$_.Message -replace '\t|\r|\n|  ', " "}},@{n='FQDN';e={$_.MachineName}}, @{n='ServerName';e={$ServerName}}, @{n='Guid';e={$Guid}}, @{n='ReportDate';e={$ReportDate}}
        }
        else {
            $EventLogs = Get-WinEvent –FilterHashtable @{LogName = 'Application', 'System'; Level = 2, 3; StartTime=$((Get-Date).AddDays(-1).Date)} -ComputerName $ServerName | 
                    Select @{n='LogName';e={$_.LogName}}, @{n='EntryType';e={($_.LevelDisplayName).ToString()}}, @{n='EventID';e={$_.ID}}, @{n='Source';e={$_.ProviderName}}, @{n='TimeGenerated';e={$_.TimeCreated}},  @{n='Message';e={$_.Message -replace '\t|\r|\n|  ', " "}},@{n='FQDN';e={$_.MachineName}}, @{n='ServerName';e={$ServerName}}, @{n='Guid';e={$Guid}}, @{n='ReportDate';e={$ReportDate}}
        }

        SQL-BulkInsert 'AXReport_EventLogs' $EventLogs

    }
    catch
    {
        Write-Log "$ServerName - ERROR - EventLogs: $($_.Exception.Message)"
        #$_.Exception.Message | Out-File C:\Users\Administrator\Documents\GitHub\DynamicsAxTools\AX-Report\Joberror.txt -Append
    }
}

Get-EventLogs

