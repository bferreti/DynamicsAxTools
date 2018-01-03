Param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string]$ServerName,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string]$Guid,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string]$ReportDate
)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"

Import-Module $ModuleFolder\AX-Tools.psm1 -DisableNameChecking

$EventLogName = 'Application', 'System'

function Get-EventLogs
{
    try
    {
        foreach($LogName in $EventLogName) {
            #Write-Log $ServerName "EventLogs |-> Started $LogName"
            $EventLogs = Get-EventLog -Computername $ServerName -LogName $LogName -EntryType Warning, Error -After $((Get-Date).AddDays(-1).Date) |
                Select @{n='LogName';e={$LogName}}, @{n='EntryType';e={($_.EntryType).ToString()}}, EventID, Source, TimeGenerated,  @{n='Message';e={$_.Message -replace '\t|\r|\n', " "}},@{n='FQDN';e={$_.MachineName}}, @{n='ServerName';e={$ServerName}}, @{n='Guid';e={$Guid}}, @{n='ReportDate';e={$ReportDate}}
            #Write-Log $ServerName "EventLogs |-> Total $($LogName): $($EventLogs.Count) records."
            SQL-BulkInsert 'AXReport_EventLogs' $EventLogs
        }
    }
    catch
    {
        Write-Log "$ServerName - ERROR - EventLogs: {0}" -f $_.Exception.Message
    }
}

Get-EventLogs