function Get-HtmlOpen {
<#
	.SYNOPSIS
		Get's HTML for the header of the HTML report
    .PARAMETER TitleText
		The title of the report
#>
[CmdletBinding()]
param (
	[String]$TitleText,
	[Switch]$SimpleHTML,
    [Switch]$AxReport,
    [Switch]$AxSummary
)
	
$CurrentDate = Get-Date -format "MMM d, yyyy hh:mm tt"

if($SimpleHTML) {
$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>$($TitleText)</title>
      <style type=text/css>
      *{font-family:Segoe UI Symbol;margin-top:4px;margin-bottom:4px}
       body{margin:8px 5px}
       h1{color:#000;font-size:18pt;text-align:left;text-decoration:underline}
       h2{color:#000;font-size:16pt;text-align:left;text-decoration:underline}
       h3{color:#000;font-size:14pt;text-align:left;text-decoration:underline}
       hr{background:#337e94;height:4px}
       table{border:1px solid #000033;border-collapse:collapse;margin:0px;margin-left:10px}
       td{border:1px solid #000033;font-size:8pt;font-weight:550;padding-left:3px;padding-right:15px;}
       th{background:#337e94;border:1px solid #000033;color:#FFF;font-size:9pt;font-weight:bold;margin:0px;padding:2px;text-align:center}
       table.fixed{table-layout:fixed}
       tr:hover{background:#808080}
       div.header{color:black;font-size:12pt;font-weight:bold;background-color:transparent;margin-bottom:4px;margin-top:12px}
       div.footer{padding-right:5em;text-align:right;font-size:9pt;padding:2px}
       div.reportdate{font-size:12pt;font-weight:bold}
       div.reportname{font-size:16pt;font-weight:bold}
       div.section{width:auto}
       .header{background:#616a6b;color:#f7f9f9}
       .odd{background:#d5d8dc}
       .even{background:#f7f9f9}
        .green {background-color:#a1cda4;}
        .yellow {background-color:#fffab1;}
        .red {background-color:#FF0000;}
        .orange {background-color:#FFA500}
        .lightred {background-color:#FFA39F}
        .lightyellow {background-color:#FFFFA9}
        .lightgreen {background-color:#D7FFCC}
       .none{background:#FFF}
       </style>
</head>
<div class="section">
    <div class="ReportName">$($TitleText)</div>
    <hr/>
</div>
"@
}

elseif($AxReport) {
$Report = @"
MIME-Version: 1.0 Content-Type: multipart/related;boundary="PART";type="text/html" 
--PART Content-Type: text/html;charset=us-ascii Content-Transfer-Encoding: 7bit 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"> 
<html><head> 
<title>$($TitleText)</title> 
    <style type="text/css"> * {margin: 0px;font-family: sans-serif;font-size: 8pt;}
    body {margin: 8px 5px 8px 5px;}
    hr {height: 4px;background-color: #337e94;border: 0px;}
    table {table-layout: auto;width: 100%;border-collapse: collapse;}
    th {vertical-align: top;text-align: left;padding: 2px 5px 2px 5px;}
    td {vertical-align: top;padding: 2px 5px 2px 5px;border-top: 1px solid #bbbbbb;}
    div.section {padding-bottom: 12px;}
    div.header {border: 1px solid #bbbbbb;padding: 4px 5em 0px 5px;margin: 0px 0px -1px 0px;height: 2em;width: 95%;font-weight: bold ;color: #ffffff;background-color: #337e94;}
    div.content {border: 1px solid #bbbbbb;padding: 4px 0px 5px 11px;margin: 0px 0px -1px 0px;width: 95%;color: #000000;background-color: #f9f9f9;}
    div.reportname {font-size: 16pt;font-weight: bold;}
    div.reportdate {font-size: 12pt;font-weight: bold;}
    div.footer {padding-right: 5em;text-align: right;}
    table.fixed {table-layout: fixed;}
    th.content {border-top: 1px solid #bbbbbb;width: 25%;}
    td.content {width: 75%;}
    td.groupby {border-top: 3px double #bbbbbb;}
    .green {background-color: #a1cda4;}
    .yellow {background-color: #fffab1;}
    .red {background-color: #f5a085;}
    .odd {background-color: #D5D8DC;}
    .even {background-color: #F7F9F9;}
    .header {background-color: #616A6B;color: #F7F9F9;}
    div.column {width: 100%;float: left;overflow-y: auto;}
    div.first {border-right: 1px grey solid;width: 49% }
    div.second {margin-left: 10px;width: 49% }
</style> 

<script type="text/javascript"> 
function show(obj) {
  document.getElementById(obj).style.display='block'; 
  document.getElementById("hide_" + obj).style.display=''; 
  document.getElementById("show_" + obj).style.display='none'; 
} 
function hide(obj) { 
  document.getElementById(obj).style.display='none'; 
  document.getElementById("hide_" + obj).style.display='none'; 
  document.getElementById("show_" + obj).style.display=''; 
} 
</script> 
</head>
<body onload="hide();">

<div class="section">
    <div class="ReportName">$($TitleText) - $((Get-Date).AddDays(-1) | Get-Date -Format "D")</div>
    <hr/>
</div>
"@
}

elseif($AxSummary) {
$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"> 
<html>
<head><title>$($TitleText)</title> 
    <style type="text/css"> * {margin: 0px;font-family: sans-serif;font-size: 8pt;}
    body {margin: 8px 5px 8px 5px;}
    hr {height: 4px;background-color: #337e94;border: 0px;}
    table {table-layout: auto;width: 100%;border-collapse: collapse;}
    th {vertical-align: top;text-align: left;padding: 2px 5px 2px 5px;}
    td {vertical-align: top;padding: 2px 5px 2px 5px;border-top: 1px solid #bbbbbb;}
    div.section {padding-bottom: 12px;}
    div.header {border: 1px solid #bbbbbb;margin: 0px 0px -1px 0px;height: 2em;width: 95%;font-weight: bold ;color: #ffffff;background-color: #337e94;}
    div.content {border: 1px solid #bbbbbb;margin: 0px 0px -1px 0px;width: 95%;color: #000000;background-color: #f9f9f9;}
    div.reportname {font-size: 16pt;font-weight: bold;}
    div.footer {padding-right: 5em;text-align: right;}
    table.fixed {table-layout: fixed;}
    th.content {border-top: 1px solid #bbbbbb;width: 25%;}
    td.content {width: 75%;}td.groupby {border-top: 3px double #bbbbbb;}
    .green {background-color: #a1cda4;}
    .yellow {background-color: #fffab1;}
    .red {background-color: #f5a085;}
    .odd {background-color: #D5D8DC;}
    .even {background-color: #F7F9F9;}
    .header {background-color: #616A6B;color: #F7F9F9;}
    </style> 
</head>

<div class="section"> 
    <div class="reportname">$($TitleText)</div> 
    <hr/>
    <br></br> 
</div>
"@
}

else {
$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>$($TitleText)</title>
    <style type="text/css">*{font:8pt sans-serif;margin:0px}
    body{margin:8px 5px 8px 5px}
    hr{background:#337e94;border:0px;height:4px}
    table{border-collapse:collapse;table-layout:auto;width:100%}
    td{border-top:1px solid #bbbbbb;padding:2px 5px 2px 5px;vertical-align:top}
    th{padding:2px 5px 2px 5px;text-align:left;vertical-align:top}
    div.column{float:left;overflow-y:auto;width:100%}
    div.content{background:#f9f9f9;border:1px solid #bbbbbb;color:#000;margin:0px 0px -1px 0px;padding:4px 0px 5px 11px;width:95%}
    div.first{border-right:1px grey solid;width:49%}
    div.footer{padding-right:5em;text-align:right}
    div.header{background:#337e94;border:1px solid #bbbbbb;color:#fff;font-weight:bold;height:2em;margin:0px 0px -1px 0px;padding:4px 5em 0px 5px;width:95%}
    div.reportdate{font-size:12pt;font-weight:bold}
    div.reportname{font-size:16pt;font-weight:bold}
    div.second{margin-left:10px;width:49%}
    div.section{padding-bottom:12px}
    table.fixed{table-layout:fixed}
    td.content{width:75%}
    td.groupby{border-top:3px double #bbbbbb}
    th.content{border-top:1px solid #bbbbbb;width:25%}
    .header{background:#616A6B;color:#F7F9F9}
    .odd{background:#d5d8dc}
    .even{background:#f7f9f9}
    .green {background-color: #a1cda4;}
    .yellow {background-color: #fffab1;}
    .red {background-color: #FF0000;}
    .orange {background-color:#FFA500}
    .lightred {background-color:#FFA39F}
    .lightyellow {background-color:#FFFFA9}
    .lightgreen {background-color:#D7FFCC}
    .none{background:#FFF}
</style>

<script type="text/javascript"> 
function show(obj) {
  document.getElementById(obj).style.display='block'; 
  document.getElementById("hide_" + obj).style.display=''; 
  document.getElementById("show_" + obj).style.display='none'; 
} 
function hide(obj) { 
  document.getElementById(obj).style.display='none'; 
  document.getElementById("hide_" + obj).style.display='none'; 
  document.getElementById("show_" + obj).style.display=''; 
} 
</script> 
</head>
<body onload="hide();">

<div class="section">
    <div class="ReportName">$($TitleText) - $((Get-Date).AddDays(-1) | Get-Date -Format "D")</div>
    <hr/>
</div>
"@
}
	Return $Report
}

function Get-HtmlClose
{
Param(
	[string]$FooterText,
    [Switch]$AxReport,
    [Switch]$AxSummary    	
)

$Footer = "Date: {0} | UserName: {1}\{2} | {3}" -f $(Get-Date),$env:UserDomain,$env:UserName,$FooterText

if($AxReport) {
$Report = @"
<div class="section">
    <hr />
    <div class="Footer">$Footer</div>
</div>
    
</body>
</html>

--PART-- 
"@
}

elseif($AxSummary) {
$Report = @"
<div class="section">
    <hr />
    <div class="footer">$Footer</div>
</div>
    
</body>
</html>
"@
}

else {
$Report = @"
<div class="section">
    <hr />
    <div class="Footer">$Footer</div>
</div></div></div>
    
</body>
</html>

"@
}
	Write-Output $Report
}

function Get-HtmlContentOpen {
<#
	.SYNOPSIS
		Creates a section in HTML
	    .PARAMETER HeaderText
			The heading for the section
		.PARAMETER IsHidden
		    Switch parameter to define if the section can collapse
		.PARAMETER BackgroundShade
		    An int for 1 to 6 that defines background shading
#>	
Param(
	[string]$HeaderText, 
	[switch]$IsHidden, 
	[validateset(1,2,3,4,5,6)][int]$BackgroundShade
)

switch ($BackgroundShade)
{
    1 { $bgColorCode = "#F8F8F8" }
	2 { $bgColorCode = "#D0D0D0" }
    3 { $bgColorCode = "#A8A8A8" }
    4 { $bgColorCode = "#888888" }
    5 { $bgColorCode = "#585858" }
    6 { $bgColorCode = "#282828" }
    default { $bgColorCode = "#ffffff" }
}
if ($IsHidden) {
	$RandomNumber = Get-Random
	$Report = @"
<div class="section">
    <div class="header">
        <a name="$($HeaderText)">$($HeaderText)</a> (<a id="show_$RandomNumber" href="javascript:void(0);" onclick="show('$RandomNumber');" style="color: #ffffff;">Show</a><a id="hide_$RandomNumber" href="javascript:void(0);" onclick="hide('$RandomNumber');" style="color: #ffffff; display:none;">Hide</a>)
    </div>
    <div class="content" id="$RandomNumber" style="display:none;background-color:$($bgColorCode);"> 
"@	
}
else {
	$Report = @"
<div class="section">
    <div class="header">
        <a name="$($HeaderText)">$($HeaderText)</a>
    </div>
    <div class="content" style="background-color:$($bgColorCode);"> 
"@
}
	Return $Report
}

function Get-HtmlContentClose {
<#
	.SYNOPSIS
		Closes an HTML section
#>	
	$Report = @"
</div>
</div>
"@
	Return $Report
}

function Get-HtmlAddNewLine {
<#
	.SYNOPSIS
		Add new line
#>	
	$Report = @"
<br>
"@
	Return $Report
}

function Get-HtmlContentTable {
<#
	.SYNOPSIS
		Creates an HTML table from an array of objects
	    .PARAMETER ArrayOfObjects
			An array of objects
		.PARAMETER Fixed
		    fixes the html column width by the number of columns
		.PARAMETER GroupBy
		    The column to group the data.  make sure this is first in the array
#>	
param(
	[Array]$ArrayOfObjects, 
	[Switch]$Fixed, 
	[String]$GroupBy,
    [String]$Title,
    [String]$Style
)
	if ($GroupBy -eq '') {
		if($Title) { $Report = "<h2>$Title</h2>" }
        $ReportHtml = $ArrayOfObjects | ConvertTo-Html -Fragment
		$ReportHtml = $ReportHtml -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', ""
		$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Green</td></tr>","<tr class=`"green`">`$+</tr>"
		$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Yellow</td></tr>","<tr class=`"yellow`">`$+</tr>"
    	$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Red</td></tr>","<tr class=`"red`">`$+</tr>"
        $ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Orange</td></tr>","<tr class=`"orange`">`$+</tr>"
        $ReportHtml = $ReportHtml -replace "<tr>(.*)<td>LightRed</td></tr>","<tr class=`"lightred`">`$+</tr>"
        $ReportHtml = $ReportHtml -replace "<tr>(.*)<td>LightGreen</td></tr>","<tr class=`"lightgreen`">`$+</tr>"
        $ReportHtml = $ReportHtml -replace "<tr>(.*)<td>LightYellow</td></tr>","<tr class=`"lightyellow`">`$+</tr>"
		$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Odd</td></tr>","<tr class=`"odd`">`$+</tr>"
		$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>Even</td></tr>","<tr class=`"even`">`$+</tr>"
		$ReportHtml = $ReportHtml -replace "<tr>(.*)<td>None</td></tr>","<tr>`$+</tr>"
		$ReportHtml = $ReportHtml -replace '<th>RowColor</th>', ''
        
        $Report += $ReportHtml

		if ($Fixed.IsPresent) {	$Report = $Report -replace '<table>', '<table class="fixed">' }
        if ($Style) { $Report = $Report -replace '<table>', "<table class=""$Style"">" }
	}
	else {
		$NumberOfColumns = ($ArrayOfObjects | Get-Member -MemberType NoteProperty  | select Name).Count
		$Groupings = @()
		$ArrayOfObjects | select $GroupBy -Unique  | sort $GroupBy | foreach { $Groupings += [String]$_.$GroupBy}
		if ($Fixed.IsPresent) {	$Report = '<table class="fixed">' }
		else { $Report = "<table>" }
		$GroupHeader = $ArrayOfObjects | ConvertTo-Html -Fragment 
		$GroupHeader = $GroupHeader -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', "" -replace '<table>', "" -replace '</table>', "" -replace "<td>.+?</td>" -replace "<tr></tr>", ""
		$GroupHeader = $GroupHeader -replace '<th>RowColor</th>', ''
		$Report += $GroupHeader
		foreach ($Group in $Groupings) {
			$Report += "<tr><td colspan=`"$NumberOfColumns`" class=`"groupby`">$Group</td></tr>"
			$GroupBody = $ArrayOfObjects | where { [String]$($_.$GroupBy) -eq $Group } | select * -ExcludeProperty $GroupBy | ConvertTo-Html -Fragment
			$GroupBody = $GroupBody -replace '<col/>', "" -replace '<colgroup>', "" -replace '</colgroup>', "" -replace '<table>', "" -replace '</table>', "" -replace "<th>.+?</th>" -replace "<tr></tr>", "" -replace '<tr><td>', "<tr><td></td><td>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Green</td></tr>","<tr class=`"green`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Yellow</td></tr>","<tr class=`"yellow`">`$+</tr>"
    		$GroupBody = $GroupBody -replace "<tr>(.*)<td>Red</td></tr>","<tr class=`"red`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Odd</td></tr>","<tr class=`"odd`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>Even</td></tr>","<tr class=`"even`">`$+</tr>"
			$GroupBody = $GroupBody -replace "<tr>(.*)<td>None</td></tr>","<tr>`$+</tr>"
			$Report += $GroupBody
		}
		$Report += "</table>" 
	}
	$Report = $Report -replace 'URL01', '<a href="'
	$Report = $Report -replace 'URL02', '">'
	$Report = $Report -replace 'URL03', '</a>'
	
	if ($Report -like "*<tr>*" -and $report -like "*odd*" -and $report -like "*even*") {
			$Report = $Report -replace "<tr>",'<tr class="header">'
	}
	
	return $Report
}

function Get-HtmlContentText 
{
<#
	.SYNOPSIS
		Creates an HTML entry with heading and detail
	    .PARAMETER Heading
			The type of logo
		.PARAMETER Detail
		     Some additional pish
#>	
param(
	$Heading,
	$Detail
)

$Report = @"
<table><tbody>
	<tr>
	<th class="content">$Heading</th>
	<td class="content">$($Detail)</td>
	</tr>
</tbody></table>
"@
$Report = $Report -replace 'URL01', '<a href="'
$Report = $Report -replace 'URL02', '">'
$Report = $Report -replace 'URL03', '</a>'
Return $Report
}

function Set-TableRowColor {
<#
	.SYNOPSIS
		adds a row colour field to the array of object for processing with htmltable
	    .PARAMETER ArrayOfObjects
			The type of logo
		.PARAMETER Green
		     Some additional pish
		.PARAMETER Yellow
		     Some additional pish
		.PARAMETER Red
		    use $this and an expression to measure the value
		.PARAMETER Alertnating
			a switch the will define Odd and Even Rows in the rowcolor column 
#>	
Param (
	$ArrayOfObjects, 
	$Green, 
	$Yellow, 
	$Red,
	[switch]$Alternating 
) 
    if ($Alternating) {
		$ColoredArray = $ArrayOfObjects | Add-Member -MemberType ScriptProperty -Name RowColor -Value {
		if ((([array]::indexOf($ArrayOfObjects,$this)) % 2) -eq 0) {'Odd'}
		if ((([array]::indexOf($ArrayOfObjects,$this)) % 2) -eq 1) {'Even'}
		} -PassThru -Force | Select-Object *
	} else {
		$ColoredArray = $ArrayOfObjects | Add-Member -MemberType ScriptProperty -Name RowColor -Value {
			if (Invoke-Expression $Green) {'Green'} 
			elseif (Invoke-Expression $Red) {'Red'} 
			elseif (Invoke-Expression $Yellow) {'Yellow'} 
			else {'None'}
			} -PassThru -Force | Select-Object *
	}
	return $ColoredArray
}

function New-HTMLBarChartObject
{
<#
	.SYNOPSIS
		create a Bar chart object for use with Create-HTMLPieChart
#>	
	$ChartSize = New-Object PSObject -Property @{`
		Width = 500
		Height = 400
		Left = 40
		Top = 30
	}
	
	$DataDefinition = New-Object PSObject -Property @{`
		AxisXTitle = "AxisXTitle"
		AxisYTitle = "AxisYTitle"
		DrawingStyle = "Cylinder"
		DataNameColumnName = "name"
		DataValueColumnName = "count"
		
	}
	
	$ChartStyle = New-Object PSObject -Property @{`
		BackColor = [System.Drawing.Color]::Transparent
		ExplodeMaxValue = $false
		Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor	[System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	}
	
	$ChartObject = New-Object PSObject -Property @{`
		Type = "Column"
		Title = "Chart Title"
		Size = $ChartSize
		DataDefinition = $DataDefinition
		ChartStyle = $ChartStyle
	}
	return $ChartObject
}

function New-HTMLChart
{
<#
	.SYNOPSIS
		adds a row colour field to the array of object for processing with htmltable
	    .PARAMETER PieChartObject
			This is a custom object with Pie chart properties, Create-HTMLPieChartObject
		.PARAMETER PieChartData
			Required an array with the headings Name and Count.  Using Powershell Group-object on an array
		    
#>
	param (
		$ChartObject,
		$ChartData
	)
	
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
	
	#Create our chart object 
	$Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
	$Chart.Width = $ChartObject.Size.Width
	$Chart.Height = $ChartObject.Size.Height
	$Chart.Left = $ChartObject.Size.Left
	$Chart.Top = $ChartObject.Size.Top
	
	#Create a chartarea to draw on and add this to the chart 
	$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
	$Chart.ChartAreas.Add($ChartArea)
	[void]$Chart.Series.Add("Data")
	
	#Add a datapoint for each value specified in the arguments (args) 
	foreach ($value in $ChartData)
	{
		$datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $value.Count)
		$datapoint.AxisLabel = [string]$value.Name
		$Chart.Series["Data"].Points.Add($datapoint)
	}
	
	switch ($ChartObject.type) {
		"Column"	{
			$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
			$Chart.Series["Data"]["DrawingStyle"] = $ChartObject.ChartStyle.DrawingStyle
			($Chart.Series["Data"].points.FindMaxByValue())["Exploded"] = $ChartObject.ChartStyle.ExplodeMaxValue
		}
		
		"Pie" {
			$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
			$Chart.Series["Data"]["PieLabelStyle"] = $ChartObject.ChartStyle.PieLabelStyle
			$Chart.Series["Data"]["PieLineColor"] = $ChartObject.ChartStyle.PieLineColor
			$Chart.Series["Data"]["PieDrawingStyle"] = $ChartObject.ChartStyle.PieDrawingStyle
			($Chart.Series["Data"].points.FindMaxByValue())["Exploded"] = $ChartObject.ChartStyle.ExplodeMaxValue
			
		}
		default
		{
				
		}
	}
	
    #Set the title of the Chart to the current date and time 
	$Title = new-object System.Windows.Forms.DataVisualization.Charting.Title
	[Void]$Chart.Titles.Add($Title)
	$Chart.Titles[0].Text = $ChartObject.Title
	
	$tempfile = (Join-Path $env:TEMP $ChartObject.Title.replace(' ', '')) + ".png"
	#Save the chart to a file
	if ((test-path $tempfile)) { Remove-Item $tempfile -Force }
	$Chart.SaveImage($tempfile, "png")
	
	$Base64Chart = [Convert]::ToBase64String((Get-Content $tempfile -Encoding Byte))
	$HTMLCode = '<IMG SRC="data:image/gif;base64,' + $Base64Chart + '" ALT="' + $ChartObject.Title + '">'
	return $HTMLCode
	#return $tempfile
}

function New-HTMLPieChartObject {
<#
	.SYNOPSIS
		create a Pie chart object for use with Create-HTMLPieChart
#>	
	$ChartSize = New-Object PSObject -Property @{`
		Width = 350
		Height = 350 
		Left = 1
		Top = 1
	}
	
	$DataDefinition = New-Object PSObject -Property @{`
		DataNameColumnName = "Name"
		DataValueColumnName = "Count"
	}
	
	$ChartStyle = New-Object PSObject -Property @{`
		#PieLabelStyle = "Outside"
        PieLabelStyle = "Disabled"
		PieLineColor = "Black"
		PieDrawingStyle = "Concave"
		ExplodeMaxValue = $false
	}
	
	$PieChartObject = New-Object PSObject -Property @{`
		Type = "Pie"
		Title = "Chart Title"
		Size = $ChartSize
		DataDefinition = $DataDefinition
		ChartStyle = $ChartStyle
	}
	return $PieChartObject
}

function New-HTMLPieChart {
<#
	.SYNOPSIS
		adds a row colour field to the array of object for processing with htmltable
	    .PARAMETER PieChartObject
			This is a custom object with Pie chart properties, Create-HTMLPieChartObject
		.PARAMETER PieChartData
			Required an array with the headings Name and Count.  Using Powershell Group-object on an array
		    
#>
	param(
		$PieChartObject,
		$PieChartData
		)
	      
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

	#Create our chart object 
	$Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
	$Chart.Width = $PieChartObject.Size.Width
	$Chart.Height = $PieChartObject.Size.Height
	$Chart.Left = $PieChartObject.Size.Left
	$Chart.Top = $PieChartObject.Size.Top

	#Create a chartarea to draw on and add this to the chart 
	$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
	$Chart.ChartAreas.Add($ChartArea) 
	[void]$Chart.Series.Add("Data") 

	#Add a datapoint for each value specified in the arguments (args) 
	foreach ($value in $PieChartData) {
		$datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $value.Count)
		$datapoint.AxisLabel = [string]$value.Name
		$Chart.Series["Data"].Points.Add($datapoint)
	}
	
	$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
	$Chart.Series["Data"]["PieLabelStyle"] = $PieChartObject.ChartStyle.PieLabelStyle
	$Chart.Series["Data"]["PieLineColor"] = $PieChartObject.ChartStyle.PieLineColor 
	$Chart.Series["Data"]["PieDrawingStyle"] = $PieChartObject.ChartStyle.PieDrawingStyle
	($Chart.Series["Data"].points.FindMaxByValue())["Exploded"] = $PieChartObject.ChartStyle.ExplodeMaxValue
	

	#Set the title of the Chart to the current date and time 
	$Title = new-object System.Windows.Forms.DataVisualization.Charting.Title 
	[Void]$Chart.Titles.Add($Title) 
	$Chart.Titles[0].Text = $PieChartObject.Title

	$tempfile = (Join-Path $env:TEMP $PieChartObject.Title.replace(' ','') ) + ".png"
	#Save the chart to a file
	if ((test-path $tempfile)) {Remove-Item $tempfile -Force}
	$Chart.SaveImage( $tempfile  ,"png")

	$Base64Chart = [Convert]::ToBase64String((Get-Content $tempfile -Encoding Byte))
	$HTMLCode = '<IMG SRC="data:image/gif;base64,' + $Base64Chart + '" ALT="' + $PieChartObject.Title + '">'
	return $HTMLCode 
}

function Get-HTMLColumn1of2
{
	$report = '<div class="first column">'
	return $report
}

function Get-HTMLColumn2of2
{
	$report = '<div class="second column">'
	return $report
}

function Get-HTMLColumnClose
{
	$report = '</div>'
	return $report
}