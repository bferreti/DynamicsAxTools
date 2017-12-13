Param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]
    $AXEnvironment
)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null

$ScriptName = 'AX Perfmon Check'
$DataCollectorName = 'AxPerfmon'
$FileDateTime = 0

function Get-WrkServers
{
    #Write-Log "`t" "AX Perfmon Check Started."
        
    $Conn = New-Object System.Data.SqlClient.SQLConnection
    $Conn.ConnectionString = "Server=UDBSQCR3-MAX\MAX;Database=DynamicsAXTools;Integrated Security=True;Connect Timeout=5"
    $Query = "SELECT SERVERNAME, SERVERTYPE FROM AXServers WHERE ENVIRONMENT = '$AXEnvironment' AND ACTIVE = '1'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $Servers = New-Object System.Data.DataSet
    $TotalServers = $Adapter.Fill($Servers)
    $Conn.Close()

    $RunningServers = 0
    $FailedServers = @()

    $WrkServers = @()
    foreach($Server in $Servers.Tables[0]) {
        if(Test-Connection $Server.ServerName -Count 1 -Quiet) {
            $WrkServers += $Server
        }
        else {
            Write-Log "$($Server.ServerName) | ERROR - Server unavailable."
        }
    }
    if($WrkServers) {
        foreach($WrkServer in $WrkServers) {
            #Write-Log "$($WrkServer.ServerName)" "$($WrkServer.ServerType) collection started."
            try {
                $DataCollectorSet = New-Object -COM Pla.DataCollectorSet
                $DataCollectorSet.Query("$DataCollectorName",$WrkServer.ServerName)
                if($DataCollectorSet.Status -eq 0) {
                    $FailedServers += $WrkServer.ServerName
                    $DataCollectorSet.Start($false)
                }
                else {
                    $RunningServers++
                } 
            }
            catch {
                #$_.exception.message
                Write-Log "$($WrkServer.ServerName) - ERROR $($_.exception.message)"
            }
        }
    }
    else {
        Write-Log "ERROR - Selecting environment failed."
    }

    Write-Log "Total Servers - $TotalServers - Running $RunningServers - Failed $($FailedServers.Count) $(if($FailedServers) {($($FailedServers -join ', '))})"
}

function Write-Log($LogData)
{
    $TLogStamp = (Get-Date -DisplayHint Time)
    $ExecLog = New-Object -TypeName System.Object
    $ExecLog | Add-Member -Name CreatedDateTime -Value $TLogStamp -MemberType NoteProperty
    $ExecLog | Add-Member -Name ReportID -Value $FileDateTime -MemberType NoteProperty
    $ExecLog | Add-Member -Name ScriptName -Value $ScriptName -MemberType NoteProperty
    #$ExecLog | Add-Member -Name ServerName -Value $LogStep -MemberType NoteProperty
    $ExecLog | Add-Member -Name Log -Value $LogData.Trim() -MemberType NoteProperty
    SQL-BulkInsert 'AXTools_ExecutionLogs' $ExecLog
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
    $Conn = New-Object System.Data.SqlClient.SQLConnection
    $Conn.ConnectionString = "Server=UDBSQCR3-MAX\MAX;Database=DynamicsAXTools;Integrated Security=True;Connect Timeout=5"
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

Get-WrkServers

#Write-Host $datacollectorset.Server "|" $datacollectorset.Name "|" $datacollectorset.Status "|" $datacollectorset.Duration "|" $datacollectorset.SerialNumber "|" $datacollectorset.OutputLocation
<#
foreach($server in $Servers)
{

    try {
        $datacollectorset = new-object -COM Pla.DataCollectorSet
        $datacollectorset.Query("$DataCollectorName",$server.Name)
        Write-Host $datacollectorset.Server "|" $datacollectorset.Name "|" $datacollectorset.Status "|" $datacollectorset.Duration "|" $datacollectorset.SerialNumber "|" $datacollectorset.OutputLocation
    }
    catch {
        $msg += " | " + $_.exception.message
    }

    Write-Host "Starting Data Collector $DataCollectorName on "$server.Name""
    try {
        $datacollectorset = new-object -COM Pla.DataCollectorSet
        $datacollectorset.Query("$DataCollectorName",$server.Name)
        $datacollectorset.start($false)
    }
    catch {
        $msg = $_.exception.message
    }
    <#
    Write-Host "Deleting Data Collector $DataCollectorName on "$server.Name""
    try {
        $datacollectorset = new-object -COM Pla.DataCollectorSet
        $datacollectorset.Query($DataCollectorName, $server.Name)
        $datacollectorset.delete()
    }
    catch {
        $_.exception.message
    }

    Write-Host "Creating Data Collector $DataCollectorName on "$server.Name""
    $datacollectorset = new-object -COM Pla.DataCollectorSet
    switch($server.Type)
    {
        ("AOS")
        {
        $xml = Get-Content $DirTemplates\AxPerfmon_AOS.xml
        $datacollectorset.SetXml($xml)
        $datacollectorset.RootPath = "%systemdrive%\PerfLogs\Admin\$DataCollectorName"
            try {
                $datacollectorset.Commit($DataCollectorName , $server.Name , 0x0003) | Out-Null
            }
            catch {
                $_.exception.message
                $CIMComputer = New-CIMSession -Computername $server.Name
                Set-NetFirewallRule -DisplayGroup "Performance Logs and Alerts" -Profile Domain -Enabled True -CimSession $CIMComputer
                $datacollectorset.Commit($DataCollectorName , $server.Name , 0x0003) | Out-Null
            }
        }
        ("RDP")
        {
            $xml = Get-Content $DirTemplates\AxPerfmon_RDP.xml
            $datacollectorset.SetXml($xml)
            $datacollectorset.RootPath = "%systemdrive%\PerfLogs\Admin\$DataCollectorName"
            try {
                $datacollectorset.Commit($DataCollectorName , $server.Name , 0x0003) | Out-Null
            }
            catch {
                $_.exception.message
                $CIMComputer = New-CIMSession -Computername $server.Name
                Set-NetFirewallRule -DisplayGroup "Performance Logs and Alerts" -Profile Domain -Enabled True -CimSession $CIMComputer
                $datacollectorset.Commit($DataCollectorName , $server.Name , 0x0003) | Out-Null
            }
        }
        ("SQL")
        {
            $xml = Get-Content $DirTemplates\AxPerfmon_SQL.xml
            $datacollectorset.SetXml($xml)
            $datacollectorset.RootPath = "%systemdrive%\PerfLogs\Admin\$DataCollectorName"
            try {
                $datacollectorset.Commit($DataCollectorName , $server.Name , 0x0003) | Out-Null
            }
            catch {
                $_.exception.message
                $CIMComputer = New-CIMSession -Computername $server.Name
                Set-NetFirewallRule -DisplayGroup "Performance Logs and Alerts" -Profile Domain -Enabled True -CimSession $CIMComputer
                $datacollectorset.Commit($DataCollectorName , $server.Name , 0x0003) | Out-Null
            }
        }
        ("SRS")
        {
            $xml = Get-Content $DirTemplates\AxPerfmon_SRS.xml
            $datacollectorset.SetXml($xml)
            $datacollectorset.RootPath = "%systemdrive%\PerfLogs\Admin\$DataCollectorName" 
            try {
                $datacollectorset.Commit($DataCollectorName , $server.Name , 0x0003) | Out-Null
            }
            catch {
                $_.exception.message
                $CIMComputer = New-CIMSession -Computername $server.Name
                Set-NetFirewallRule -DisplayGroup "Performance Logs and Alerts" -Profile Domain -Enabled True -CimSession $CIMComputer
                $datacollectorset.Commit($DataCollectorName , $server.Name , 0x0003) | Out-Null
            }
        }
        ("IIS")
        {
            $xml = Get-Content $DirTemplates\AxPerfmon_IIS.xml
            $datacollectorset.SetXml($xml)
            $datacollectorset.RootPath = "%systemdrive%\PerfLogs\Admin\$DataCollectorName" 
            try {
                $datacollectorset.Commit($DataCollectorName , $server.Name , 0x0003) | Out-Null
            }
            catch {
                $_.exception.message
                $CIMComputer = New-CIMSession -Computername $server.Name
                Set-NetFirewallRule -DisplayGroup "Performance Logs and Alerts" -Profile Domain -Enabled True -CimSession $CIMComputer
                $datacollectorset.Commit($DataCollectorName , $server.Name , 0x0003) | Out-Null
            }
        }
        default {Write-Output "Invalid Type"} 
    }#>
#}
