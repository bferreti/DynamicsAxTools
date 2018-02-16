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

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [String]$AXEnvironment,
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
    [Array]$ServerType = 'AOS',
    [Switch]$Start
)

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"
Import-Module $ModuleFolder\AX-Tools.psm1 -DisableNameChecking


if($ServerType.Split(',').Count -gt 0) {
    $ServerType.Split(',').Trim() | % { [Array]$ServerTypeQuery += "'$_'" }
}

[String]$ServerType = ($ServerTypeQuery -join ',')

if($Start) {
    Test-AosServices -AxEnvironment $AXEnvironment -ServerType $ServerType -Start
}
else {
    Test-AosServices -AxEnvironment $AXEnvironment -ServerType $ServerType
}

Get-Module | Where-Object {$_.ModuleType -eq 'Script'} | % { Remove-Module $_.Name }



<#
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null

#$FileDateTime = Get-Date -f MMddyyHHmm
$ScriptName = 'AOS Check'
$FileDateTime = 0

function Get-WrkServers
{
    #Write-Log "`t" "AX Perfmon Check Started."
        
    $Conn = New-Object System.Data.SqlClient.SQLConnection
    $Conn.ConnectionString = "Server=UDBSQCR3-MAX\MAX;Database=DynamicsAXTools;Integrated Security=True;Connect Timeout=5"
    $Query = "SELECT SERVERNAME, SERVERTYPE FROM AXServers WHERE ENVIRONMENT = '$AXEnvironment' AND SERVERTYPE = '$ServerType' AND ACTIVE = '1'" 
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $Adapter.SelectCommand = $Cmd
    $Servers = New-Object System.Data.DataSet
    $TotalServers = $Adapter.Fill($Servers)
    $Conn.Close()

    $Stopped = @()
    $Running = 0
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
            if((Get-Service -Name "AOS60`$01" -ComputerName $WrkServer.ServerName).Status -match 'Stopped') {
                    $Stopped += $WrkServer.ServerName
                    (Get-Service -Name "AOS60`$01" -ComputerName $WrkServer.ServerName).Start()
                }
                else {
                    $Running++
                    Write-Host "$($WrkServer.ServerName) is Running."
                }
        }
    }
    else {
        Write-Log "`t" "ERROR - Selecting environment failed."
    }

    Write-Log "Total Servers - $TotalServers - Running $Running - Failed $($Stopped.Count) $(if($Stopped) {($($Stopped -join ', '))})"
}

function AOSCheck_Deprecated
{
    foreach($Server in $Servers) {
        if((Get-Service -Name "AOS60`$01" -ComputerName $Server).Status -match 'Stopped') {
            $StoppedServers += $Server
            (Get-Service -Name "AOS60`$01" -ComputerName $Server).Start()
        }
        else {
            Write-Host "$Server is Running."
        }
    }
    if($StoppedServers) {
        Write-Log "Stopped AOS Servers - $($StoppedServers -join ', ')"
    }
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
#>