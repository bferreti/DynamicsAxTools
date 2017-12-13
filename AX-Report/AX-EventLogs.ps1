Param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]
    $ServerName,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]
    $ServerType,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]
    $FileDateTime
)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"
$ToolsFolder = $Dir + "\AX-Tools"
$ReportFolder = $Dir + "\Reports\AX-Report"
$LogFolder = $Dir + "\Logs\AX-Report"

Import-Module $ModuleFolder\AX-Database.psm1 -DisableNameChecking
Import-Module $ModuleFolder\AX-HTMLReport.psm1 -DisableNameChecking
Import-Module $ModuleFolder\AX-SendEmail.psm1 -DisableNameChecking

$EventLogName = 'Application', 'System'

function Get-EventLogs
{
    try
    {
        foreach($LogName in $EventLogName) {
            #Write-Log $ServerName "EventLogs |-> Started $LogName"
            $EventLogs = Get-EventLog -Computername $ServerName -LogName $LogName -EntryType Warning, Error -After $((Get-Date).AddDays(-1).Date) |
                Select MachineName, @{n='LogName';e={$LogName}}, @{n='EntryType';e={($_.EntryType).ToString()}}, EventID, Source, TimeGenerated,  @{n='Message';e={$_.Message -replace '\t|\r|\n', " "}}
            #Write-Log $ServerName "EventLogs |-> Total $($LogName): $($EventLogs.Count) records."
            SQL-InsertDB 'AXReportEventLogs' $EventLogs
        }
    }
    catch
    {
        Write-Log "$ServerName - ERROR - EventLogs: {0}" -f $_.Exception.Message
    }
}

function SQL-InsertDB($Table, $Data)
{
    $CreatedDateTime = Get-Date -f G
    if($Table | Select-String 'AXReportSSRSLogs|AXReportSQLServerLogs|AXReportCDXJobs|AXReportBatchJobs|AXReportMRP|AXReportLongBatchJobs') {
        $Data = $Data | Select *, @{n='CreatedDateTime';e={$CreatedDateTime}}, @{n='ReportID';e={$FileDateTime}}
    }
    else {
        $Data = $Data | Select *, @{n='ServerName';e={$ServerName}}, @{n='ServerType';e={$ServerType}}, @{n='CreatedDateTime';e={$CreatedDateTime}}, @{n='ReportID';e={$FileDateTime}}
    }
    SQL-BulkInsert $Table $Data
}

function SQL-BulkInsert($Table, $Data)
{
    $DataTable = New-Object Data.DataTable   
    $First = $true  
    foreach($Object in $Data) 
    { 
        $DataReader = $DataTable.NewRow()   
        foreach($Property in $Object.PsObject.Get_Properties()) 
        {   
            if ($First) 
            {   
                $Col =  new-Object Data.DataColumn   
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
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
    $BCopy = New-Object ("System.Data.SqlClient.SqlBulkCopy") $Conn
    $BCopy.DestinationTableName = "dbo.$Table"
    foreach ($Col in $DataTable.Columns) {
        $ColumnMap = New-Object ("Data.SqlClient.SqlBulkCopyColumnMapping") $Col.ColumnName,($Col.ColumnName).ToUpper()
        [Void]$BCopy.ColumnMappings.Add($ColumnMap)
    }
    $BCopy.WriteToServer($DataTable)
    $Conn.Close()
}

function Write-Log($LogData)
{
    $TLogStamp = (Get-Date -DisplayHint Time)
    $ExecLog = New-Object -TypeName System.Object
    $ExecLog | Add-Member -Name CreatedDateTime -Value $TLogStamp -MemberType NoteProperty
    $ExecLog | Add-Member -Name ReportID -Value $FileDateTime -MemberType NoteProperty
    $ExecLog | Add-Member -Name ScriptName -Value 'AX Report' -MemberType NoteProperty
    #$ExecLog | Add-Member -Name ServerName -Value $LogStep -MemberType NoteProperty
    $ExecLog | Add-Member -Name Log -Value $LogData -MemberType NoteProperty
    SQL-BulkInsert 'AXTools_ExecutionLogs' $ExecLog
}

Get-EventLogs