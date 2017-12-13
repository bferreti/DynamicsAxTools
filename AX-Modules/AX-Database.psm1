$ParamDBServer = 'HPFAXSQL01' #Change SQL Server Name. Server\Instance, (local)
$ParamDBName = 'DynamicsAXTools' #Change DB Name (if you have changed during creation)

function Get-ConnectionString {
    return "Server=$ParamDBServer;Database=$ParamDBName;Integrated Security=True;Connect Timeout=5"
}

<#
function SQL-InsertDB
{
param (
    [String]$Table,
    [Array]$Data
)
    $CreatedDateTime = Get-Date -f G
    if($Table | Select-String 'AXReportSSRSLogs|AXReportSQLServerLogs|AXReportCDXJobs|AXReportBatchJobs|AXReportMRP|AXReportLongBatchJobs') {
        $Data = $Data | Select *, @{n='CreatedDateTime';e={$CreatedDateTime}}, @{n='ReportID';e={$FileDateTime}}
    }
    else {
        $Data = $Data | Select *, @{n='ServerName';e={$WrkServer.ServerName}}, @{n='ServerType';e={$WrkServer.ServerType}}, @{n='CreatedDateTime';e={$CreatedDateTime}}, @{n='ReportID';e={$FileDateTime}}
    }
    SQL-BulkInsert $Table $Data
} #>

function SQL-BulkInsert
{
param (
    [String]$Table,
    [Array]$Data
)
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
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
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

function Write-Log
{
param (
    [String]$LogData
)
    $TLogStamp = (Get-Date -DisplayHint Time)
    $ExecLog = New-Object -TypeName System.Object
    $ExecLog | Add-Member -Name CreatedDateTime -Value $TLogStamp -MemberType NoteProperty
    $ExecLog | Add-Member -Name ReportID -Value $FileDateTime -MemberType NoteProperty
    $ExecLog | Add-Member -Name ScriptName -Value $ScriptName -MemberType NoteProperty
    #$ExecLog | Add-Member -Name ServerName -Value $LogStep -MemberType NoteProperty
    $ExecLog | Add-Member -Name Log -Value $LogData.Trim() -MemberType NoteProperty
    SQL-BulkInsert 'AXTools_ExecutionLogs' $ExecLog
}

function UpdateMsiStatus
{
param (
    [string]$Status
)
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
    $Query = "UPDATE [dbo].[AXInstallStatus] SET Status = '$Status' WHERE GUID = '$Guid'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function SQL-UpdateTable
{
param (
    [String]$Table,
    [String]$Set,
    [String]$Value,
    [String]$Where
)
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
    $Query = "UPDATE [dbo].[$Table] SET [$Set] = '$Value' WHERE $Where"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function SQL-ExecUpdate
{
param (
    [String]$Query
)
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
    #$Query = "UPDATE [dbo].[$Table] SET [$Set] = '$Value' WHERE $Where"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}

function SQL-WriteLog
{
param (
    [String]$Log
)
    $Conn = New-Object System.Data.SqlClient.SQLConnection(Get-ConnectionString)
    $Conn.Open()
    $Query = "INSERT INTO AXTools_ExecutionLogs VALUES('$(Get-Date)', '', 'AXMonitor', '$Log')"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery()
    $Conn.Close()
}