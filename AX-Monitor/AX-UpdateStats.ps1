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
    [String]$SQLInstance,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]$AXDatabase,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]$Table,
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
    [String]$StatsType,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]$GRDJobName,
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
    [String]$SQLUsername
)
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"

Import-Module $ModuleFolder\AX-Tools.psm1 -DisableNameChecking

$Table = ($Table.Replace('[','')).Replace(']','')

try
{
    if($Table.Split('.').Count -gt 1) {
        $Schema = $Table.Split('.')[0]
        $Table = $Table.Split('.')[1]
    }

    if($SQLUsername) {
        $Server = Get-SQLObject -DBServer $SQLInstance -DBName $AXDatabase -SQLAccount $SQLUsername -ApplicationName 'Ax Powershell Tools (SQL STATS)' -SQLServerObject
    }
    else {
        $Server = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $SQLInstance
        $Server.ConnectionContext.ApplicationName = 'Ax Powershell Tools (SQL STATS)'
    }
    $Server.ConnectionContext.StatementTimeout = 0
    $Db = $Server.Databases["$AXDatabase"]
    if($Schema) { $Db.DefaultSchema = $Schema }
    
    if($StatsType -match 'FullScan') {
        #$Server.ConnectionContext.StatementTimeout = 900
        $Db.Tables["$Table"].UpdateStatistics('All','FullScan')
    }
    else {
        $Server.ConnectionContext.StatementTimeout = 810 #270
        $Db.Tables["$Table"].UpdateStatistics()        
    }
}
catch
{
    $Msg = "ERROR - {0}" -f $_.Exception #.Message
    SQL-UpdateTable 'AXMonitor_GRDLog' 'LOG' $($Msg) "JOBNAME = '$GRDJobName'"
}


$JobSettings = Load-ScriptSettings -ScriptName 'AxMonitor'
if([boolean]::Parse($JobSettings.Debug)) {
    $Environment = $GRDJobName.Split('_')[1]
    if(![String]::IsNullOrEmpty($JobSettings.LogFolder)) {
        $JobSettings.LogFolder = Join-Path $JobSettings.LogFolder $Environment
    }
    else {
        $JobSettings.LogFolder = Join-Path $Dir "Logs\$Environment"
    }

    "$GRDJobName - $SQLInstance - $AXDatabase - $Table - $StatsType `r`n $($SQLUsername.UserName) `r`n $Msg" | Out-File $(Join-Path $JobSettings.LogFolder "$GRDJobName.txt")
}

$Server.ConnectionContext.Disconnect()