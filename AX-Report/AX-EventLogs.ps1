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

Param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string]$ServerName,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string]$Guid,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string]$ReportDate,
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
    [System.Management.Automation.PSCredential]$Credentials
)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"

Import-Module $ModuleFolder\AX-Tools.psm1 -DisableNameChecking

function Get-EventLogs
{
    Write-Log "Running EvenLogs job for $ServerName. RunAs - $($Credentials.UserName)"
    if($Credentials) {
        try {
            $EventLogs = Get-WinEvent –FilterHashtable @{LogName = 'Application', 'System'; Level = 2, 3; StartTime=$((Get-Date).AddDays(-1).Date)} -ComputerName $ServerName -Credential $Credentials | 
                    Select @{n='LogName';e={$_.LogName}}, @{n='EntryType';e={($_.LevelDisplayName).ToString()}}, @{n='EventID';e={$_.ID}}, @{n='Source';e={$_.ProviderName}}, @{n='TimeGenerated';e={$_.TimeCreated}},  @{n='Message';e={$_.Message -replace '\t|\r|\n|  ', " "}},@{n='FQDN';e={$_.MachineName}}, @{n='ServerName';e={$ServerName}}, @{n='Guid';e={$Guid}}, @{n='ReportDate';e={$ReportDate}}
        }
        catch [System.Diagnostics.Eventing.Reader.EventLogException]{
	        $CIMComputer = New-CimSession -ComputerName $ServerName
	        Enable-NetFirewallRule -DisplayGroup "Remote Event Log Management" -CimSession $CIMComputer
	        Remove-CimSession -ComputerName $ServerName
            #
            $EventLogs = Get-WinEvent –FilterHashtable @{LogName = 'Application', 'System'; Level = 2, 3; StartTime=$((Get-Date).AddDays(-1).Date)} -ComputerName $ServerName -Credential $Credentials | 
                    Select @{n='LogName';e={$_.LogName}}, @{n='EntryType';e={($_.LevelDisplayName).ToString()}}, @{n='EventID';e={$_.ID}}, @{n='Source';e={$_.ProviderName}}, @{n='TimeGenerated';e={$_.TimeCreated}},  @{n='Message';e={$_.Message -replace '\t|\r|\n|  ', " "}},@{n='FQDN';e={$_.MachineName}}, @{n='ServerName';e={$ServerName}}, @{n='Guid';e={$Guid}}, @{n='ReportDate';e={$ReportDate}}
        }
        catch {
            $Exception = $_.Exception.Message
        }
    }
    else {
        try {
            $EventLogs = Get-WinEvent –FilterHashtable @{LogName = 'Application', 'System'; Level = 2, 3; StartTime=$((Get-Date).AddDays(-1).Date)} -ComputerName $ServerName | 
                Select @{n='LogName';e={$_.LogName}}, @{n='EntryType';e={($_.LevelDisplayName).ToString()}}, @{n='EventID';e={$_.ID}}, @{n='Source';e={$_.ProviderName}}, @{n='TimeGenerated';e={$_.TimeCreated}},  @{n='Message';e={$_.Message -replace '\t|\r|\n|  ', " "}},@{n='FQDN';e={$_.MachineName}}, @{n='ServerName';e={$ServerName}}, @{n='Guid';e={$Guid}}, @{n='ReportDate';e={$ReportDate}}
        }
        catch [System.Diagnostics.Eventing.Reader.EventLogException]{
	        $CIMComputer = New-CimSession -ComputerName $ServerName
	        Enable-NetFirewallRule -DisplayGroup "Remote Event Log Management" -CimSession $CIMComputer
	        Remove-CimSession -ComputerName $ServerName
            #
            $EventLogs = Get-WinEvent –FilterHashtable @{LogName = 'Application', 'System'; Level = 2, 3; StartTime=$((Get-Date).AddDays(-1).Date)} -ComputerName $ServerName | 
                Select @{n='LogName';e={$_.LogName}}, @{n='EntryType';e={($_.LevelDisplayName).ToString()}}, @{n='EventID';e={$_.ID}}, @{n='Source';e={$_.ProviderName}}, @{n='TimeGenerated';e={$_.TimeCreated}},  @{n='Message';e={$_.Message -replace '\t|\r|\n|  ', " "}},@{n='FQDN';e={$_.MachineName}}, @{n='ServerName';e={$ServerName}}, @{n='Guid';e={$Guid}}, @{n='ReportDate';e={$ReportDate}}
        }
        catch {
            $Exception = $_.Exception.Message
        }
    }
    SQL-BulkInsert 'AXReport_EventLogs' $EventLogs
    #
    if($Exception) { Write-Log "$ServerName - ERROR - EventLogs: $Exception" }
}

Get-EventLogs



