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
    [String]$GRDJobName
)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null

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

    $Server = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $SQLInstance
    $Db = $Server.Databases["$AXDatabase"]
    if($Schema) { $Db.DefaultSchema = $Schema }
    $Server.ConnectionContext.StatementTimeout = 0
    
    if($StatsType -match 'FullScan') {
        #$Server.ConnectionContext.StatementTimeout = 900
        $Db.Tables["$Table"].UpdateStatistics('All','FullScan')
    }
    else {
        $Server.ConnectionContext.StatementTimeout = 540 #270
        $Db.Tables["$Table"].UpdateStatistics()        
    }
}
catch
{
    $Msg = "ERROR - {0}" -f $_.Exception.Message
    SQL-UpdateTable 'AXMonitor_GRDLog' 'LOG' $($Msg) "JOBNAME = '$GRDJobName'"
}