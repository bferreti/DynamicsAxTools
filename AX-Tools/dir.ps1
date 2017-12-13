$Scriptpath = $MyInvocation.MyCommand.Path
$Dir = Split-Path $ScriptPath 
$RootDir = Split-Path $Dir

$SplitPath = $Dir.Split('\')[-2]
