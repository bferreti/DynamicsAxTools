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
    [String]$Environment,
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
    Test-AosServices -AxEnvironment $Environment -ServerType $ServerType -Start
}
else {
    Test-AosServices -AxEnvironment $Environment -ServerType $ServerType
}

Get-Module | Where-Object {$_.ModuleType -eq 'Script'} | % { Remove-Module $_.Name }