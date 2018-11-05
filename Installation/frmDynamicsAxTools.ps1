[void][System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework')
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
#[void][System.Reflection.Assembly]::LoadFrom('C:\Users\bferreti\Documents\Elysium.dll')
#[void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')
#[void][System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration')

$Scriptpath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptPath
$Dir = Split-Path $ScriptDir
$ModuleFolder = $Dir + "\AX-Modules"

Import-Module $ModuleFolder\AX-Tools.psm1 -DisableNameChecking

$inputXML = @"
<Window x:Class="DynamicsAxTools.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DynamicsAxTools"
        mc:Ignorable="d"
        Title="DynamicsAXTools" Height="500" Width="800">
    <Grid>
        <TabControl x:Name="tabControl" Margin="10,76,10,10" Height="300">
            <TabItem Header="Environments" TabIndex="10">
                <Grid>
                    <Rectangle Fill="#FFEFEFF1" Height="30" Margin="13,10,14,0" Stroke="Black" VerticalAlignment="Top"/>
                    <Button x:Name="btnEnvNew" Content="New" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnEnvEdit" Content="Edit" HorizontalAlignment="Left" Margin="90,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnEnvSave" Content="Save" HorizontalAlignment="Left" Margin="160,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnEnvDelete" Content="Delete" HorizontalAlignment="Left" Margin="230,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnEnvtestSQL" Content="Test DB Conn" HorizontalAlignment="Left" Margin="300,15,0,0" VerticalAlignment="Top" Width="100" IsEnabled="False"/>
                    <Button x:Name="btnEnvReports" Content="Reports" HorizontalAlignment="Left" Margin="615,15,0,0" VerticalAlignment="Top" Width="60" IsEnabled="False"/>
                    <Button x:Name="btnEnvLogs" Content="Logs" HorizontalAlignment="Left" Margin="680,15,0,0" VerticalAlignment="Top" Width="60" IsEnabled="False"/>
                    <Rectangle Fill="#FFEFEFF1" HorizontalAlignment="Left" Height="192" Margin="13,50,0,0" Stroke="Black" VerticalAlignment="Top" Width="737"/>
                    <ComboBox x:Name="cbxEnvEnvironment" HorizontalAlignment="Left" Margin="95,59,0,0" VerticalAlignment="Top" Width="223" IsEditable="False" DisplayMemberPath="ENVIRONMENT"/>
                    <TextBox x:Name="txtEnvEnvironment" HorizontalAlignment="Left" Height="24" Margin="91,85,0,0" VerticalAlignment="Top" Width="183" IsEnabled="False"/>
                    <ComboBox x:Name="cbxEnvEmail" HorizontalAlignment="Left" Margin="320,86,0,0" VerticalAlignment="Top" Width="125" DisplayMemberPath="ID" IsEnabled="False"/>
                    <Label x:Name="lblEnvLocalUser" Content="Local User" HorizontalAlignment="Left" Margin="450,84,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxEnvLocalUser" HorizontalAlignment="Left" Margin="519,86,0,0" VerticalAlignment="Top" Width="125" DisplayMemberPath="ID" IsEnabled="False"/>
                    <CheckBox x:Name="chkEnvRefresh" Content="AX Refresh" HorizontalAlignment="Left" Margin="656,90,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxEnvDBUser" HorizontalAlignment="Left" Margin="78,123,0,0" VerticalAlignment="Top" Width="125" DisplayMemberPath="ID" IsEnabled="False"/>
                    <TextBox x:Name="txtEnvDBServer" HorizontalAlignment="Left" Height="24" Margin="286,122,0,0" VerticalAlignment="Top" Width="175" IsEnabled="False"/>
                    <TextBox x:Name="txtEnvDBName" HorizontalAlignment="Left" Height="24" Margin="535,123,0,0" VerticalAlignment="Top" Width="175" IsEnabled="False"/>
                    <ComboBox x:Name="cbxEnvDBStats" HorizontalAlignment="Left" Margin="130,150,0,0" VerticalAlignment="Top" Width="200" DisplayMemberPath="Value" SelectedValuePath="Name" IsEnabled="False"/>
                    <CheckBox x:Name="chkEnvGRD" Content="Enable GRD" HorizontalAlignment="Left" Margin="22,190,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                    <TextBox x:Name="txtEnvCPU" HorizontalAlignment="Left" Height="24" Margin="109,209,0,0" VerticalAlignment="Top" Width="75" IsEnabled="False"/>
                    <TextBox x:Name="txtEnvBlocking" HorizontalAlignment="Left" Height="24" Margin="305,209,0,0" VerticalAlignment="Top" Width="75" />
                    <TextBox x:Name="txtEnvWaiting" HorizontalAlignment="Left" Height="24" Margin="495,209,0,0" VerticalAlignment="Top" Width="75" IsEnabled="False"/>
                    <Label x:Name="lblEnvDBServer" Content="SQL Server" HorizontalAlignment="Left" Margin="214,121,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblEnvDBName" Content="DB Name" HorizontalAlignment="Left" Margin="470,121,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblEnvDBUser" Content="SQL User" HorizontalAlignment="Left" Margin="15,121,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblEnvDBStats" Content="Check DB Statistics" HorizontalAlignment="Left" Margin="15,148,0,0" VerticalAlignment="Top"/>
                    <Separator HorizontalAlignment="Left" Height="18" Margin="18,108,0,0" VerticalAlignment="Top" Width="725"/>
                    <Label x:Name="lblEnvCPU" Content="CPU Threshold" HorizontalAlignment="Left" Margin="16,208,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblEnvBlocking" Content="Blocking Threshold" HorizontalAlignment="Left" Margin="189,208,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblEnvWaiting" Content="Waiting Threshold" HorizontalAlignment="Left" Margin="385,208,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblEnvName" Content="Environment" HorizontalAlignment="Left" Margin="15,57,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblEnvDescription" Content="Description" HorizontalAlignment="Left" Margin="15,84,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblEnvEmail" Content="Email" HorizontalAlignment="Left" Margin="279,84,0,0" VerticalAlignment="Top"/>
                    <Separator HorizontalAlignment="Left" Height="18" Margin="19,172,0,0" VerticalAlignment="Top" Width="725"/>
                </Grid>
            </TabItem>
            <TabItem Header="Users" TabIndex="20">
                <Grid>
                    <Rectangle Fill="#FFEFEFF1" Height="30" Margin="13,10,14,0" Stroke="Black" VerticalAlignment="Top"/>
                    <Button x:Name="btnUsrNew" Content="New" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnUsrDelete" Content="Delete" HorizontalAlignment="Left" Margin="90,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnUsrTest" Content="Test" HorizontalAlignment="Left" Margin="160,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Rectangle Fill="#FFEFEFF1" Height="65" Margin="13,50,14,0" Stroke="Black" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxUsrID" HorizontalAlignment="Left" Margin="77,58,0,0" VerticalAlignment="Top" Width="180" DisplayMemberPath="ID"/>
                    <Label x:Name="lblUsrID" Content="User ID" HorizontalAlignment="Left" Margin="23,56,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblUsrUsername" Content="Username" HorizontalAlignment="Left" Margin="23,83,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtUsrUsername" HorizontalAlignment="Left" Height="22" Margin="91,85,0,0" VerticalAlignment="Top" Width="320" IsEnabled="False"/>
                </Grid>
            </TabItem>
            <TabItem Header="Emails" TabIndex="30">
                <Grid>
                    <Rectangle Fill="#FFEFEFF1" HorizontalAlignment="Left" Height="30" Margin="13,10,0,0" Stroke="Black" VerticalAlignment="Top" Width="737"/>
                    <Button x:Name="btnEmlNew" Content="New" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnEmlEdit" Content="Edit" HorizontalAlignment="Left" Margin="90,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnEmlSave" Content="Save" HorizontalAlignment="Left" Margin="160,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnEmlDelete" Content="Delete" HorizontalAlignment="Left" Margin="230,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnEmlTest" Content="Test Email" HorizontalAlignment="Left" Margin="300,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Rectangle Fill="#FFEFEFF1" Height="208" Margin="13,50,0,0" Stroke="Black" VerticalAlignment="Top" HorizontalAlignment="Left" Width="737"/>
                    <Label x:Name="lblEmlID" Content="ID" HorizontalAlignment="Left" Margin="19,54,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxEmlID" HorizontalAlignment="Left" Margin="46,56,0,0" VerticalAlignment="Top" Width="180" DisplayMemberPath="ID"/>
                    <Label x:Name="lblEmlSMTP" Content="Email Server" HorizontalAlignment="Left" Margin="19,80,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlSMTP" HorizontalAlignment="Left" Height="24" Margin="97,82,0,0" VerticalAlignment="Top" Width="224" IsEnabled="False"/>
                    <Label x:Name="lblEmlSMTPPort" Content="Port" HorizontalAlignment="Left" Margin="326,81,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlSMTPPort" HorizontalAlignment="Left" Height="24" Margin="363,83,0,0" VerticalAlignment="Top" Width="75" IsEnabled="False"/>
                    <CheckBox x:Name="chkEmlSSL" Content="Use SSL" HorizontalAlignment="Left" Margin="450,86,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                    <Label x:Name="lblEmlUserId" Content="User" HorizontalAlignment="Left" Margin="19,108,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxEmlUserID" HorizontalAlignment="Left" Margin="58,110,0,0" VerticalAlignment="Top" Width="180" IsEnabled="False" DisplayMemberPath="ID"/>
                    <Label x:Name="lblEmlFrom" Content="From" HorizontalAlignment="Left" Margin="19,134,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlFrom" HorizontalAlignment="Left" Height="24" Margin="61,135,0,0" VerticalAlignment="Top" Width="301" IsEnabled="False"/>
                    <Label x:Name="lblEmlTo" Content="To" HorizontalAlignment="Left" Margin="19,161,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlTo" HorizontalAlignment="Left" Height="24" Margin="47,162,0,0" VerticalAlignment="Top" Width="498" IsEnabled="False"/>
                    <Label x:Name="lblEmlCC" Content="CC" HorizontalAlignment="Left" Margin="19,188,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlCC" HorizontalAlignment="Left" Height="24" Margin="49,189,0,0" VerticalAlignment="Top" Width="496" IsEnabled="False"/>
                    <Label x:Name="lblEmlBCC" Content="BCC" HorizontalAlignment="Left" Margin="19,215,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlBCC" HorizontalAlignment="Left" Height="24" Margin="56,216,0,0" VerticalAlignment="Top" Width="489" IsEnabled="False"/>
                </Grid>
            </TabItem>
            <TabItem Header="Task Scheduler" TabIndex="40">
                <Grid>
                    <Rectangle Fill="#FFEFEFF1" Height="30" Margin="13,10,14,0" Stroke="Black" VerticalAlignment="Top"/>
                    <Button x:Name="btnTskNew" Content="New" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnTskEdit" Content="Edit" HorizontalAlignment="Left" Margin="90,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnTskDelete" Content="Delete" HorizontalAlignment="Left" Margin="160,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnTskSave" Content="Save" HorizontalAlignment="Left" Margin="230,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnTskDisable" Content="Disable" HorizontalAlignment="Left" Margin="300,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnTskEnable" Content="Enable" HorizontalAlignment="Left" Margin="370,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Rectangle Fill="#FFEFEFF1" HorizontalAlignment="Left" Height="214" Margin="13,50,0,0" Stroke="Black" VerticalAlignment="Top" Width="737" RenderTransformOrigin="0.5,0.5">
                        <Rectangle.RenderTransform>
                            <TransformGroup>
                                <ScaleTransform/>
                                <SkewTransform AngleX="0.591"/>
                                <RotateTransform/>
                                <TranslateTransform X="0.99"/>
                            </TransformGroup>
                        </Rectangle.RenderTransform>
                    </Rectangle>
                    <Label x:Name="lblTskName" Content="Environment" HorizontalAlignment="Left" Margin="15.023,53.47,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxTskEnvironment" HorizontalAlignment="Left" Margin="93,55.47,0,0" VerticalAlignment="Top" Width="100" IsEditable="False" DisplayMemberPath="ENVIRONMENT"/>
                    <Label x:Name="lblTskTaskName" Content="Task" HorizontalAlignment="Left" Margin="196.293,53.47,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxTskTaskName" HorizontalAlignment="Left" Margin="228.676,55.47,0,0" VerticalAlignment="Top" Width="125" DisplayMemberPath="Value" SelectedValuePath="Name"/>
                    <TextBox x:Name="txtTskInterval" HorizontalAlignment="Left" Height="22" Margin="409.186,55.45,0,0" VerticalAlignment="Top" Width="40" IsEnabled="False"/>
                    <Label x:Name="lblTskInterval" Content="Interval:" HorizontalAlignment="Left" Margin="357.342,53.47,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblTskTimeSpan" Content="At:" HorizontalAlignment="Left" Margin="451.291,53.47,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtTskTime" HorizontalAlignment="Left" Height="22" Margin="474.62,55.45,0,0" VerticalAlignment="Top" Width="55" IsEnabled="False"/>
                    <Label x:Name="lblTskUserId" Content="Run as" HorizontalAlignment="Left" Margin="532.283,53.47,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxTskUserID" HorizontalAlignment="Left" Margin="580.621,55.47,0,0" VerticalAlignment="Top" Width="160" IsEnabled="False" DisplayMemberPath="ID"/>
                    <ListView x:Name="lstCurrJobs" Height="180" Margin="16.5,82.659,17.5,0" VerticalAlignment="Top" Width="732">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Environment" DisplayMemberBinding ="{Binding Environment}" Width="85"/>
                                <GridViewColumn Header="Name" DisplayMemberBinding ="{Binding Name}" Width="85"/>
                                <GridViewColumn Header="Repetition" DisplayMemberBinding ="{Binding Interval}" Width="75"/>
                                <GridViewColumn Header="Daily" DisplayMemberBinding ="{Binding DaysInterval}" Width="75"/>
                                <GridViewColumn Header="At" DisplayMemberBinding ="{Binding At}" Width="75"/>
                                <GridViewColumn Header="User" DisplayMemberBinding ="{Binding User}" Width="75"/>
                                <GridViewColumn Header="Status" DisplayMemberBinding ="{Binding State}" Width="75"/>
                                <GridViewColumn Header="Next Run" DisplayMemberBinding ="{Binding NextRunTime}" Width="75"/>
                                <GridViewColumn Header="Last Run" DisplayMemberBinding ="{Binding LastRunTime}" Width="75"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Grid>
            </TabItem>
            <TabItem Header="Services" TabIndex="50">
                <Grid>
                </Grid>
            </TabItem>
            <TabItem Header="Perfmon" TabIndex="60">
                <Grid>
                </Grid>
            </TabItem>            
            <TabItem Header="Settings" TabIndex="70">
                <Grid>
                    <DataGrid x:Name="dgXMLSettings" HorizontalAlignment="Left" Height="210" Margin="13,50,0,0" VerticalAlignment="Top" Width="737" AutoGenerateColumns="False">
                        <DataGrid.Columns>
                            <DataGridTextColumn Binding="{Binding Key}" Header="Parameter" Width="200" CanUserResize="True"/>
                            <DataGridTextColumn Binding="{Binding Value}" Header="Value" Width="Auto" CanUserResize="True"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <Rectangle Fill="#FFEFEFF1" Height="30" Margin="13,10,0,0" Stroke="Black" VerticalAlignment="Top" HorizontalAlignment="Left" Width="737"/>
                    <Button x:Name="btnSetSave" Content="Save" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top" Width="65"/>
                </Grid>
            </TabItem>
            <TabItem Header="Database" TabIndex="80">
                <Grid>
                    <Rectangle Fill="#FFEFEFF1" Height="30" Margin="13,10,14,0" Stroke="Black" VerticalAlignment="Top"/>
                    <Button x:Name="btnDBCreate" Content="Create" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnDBDrop" Content="Drop" HorizontalAlignment="Left" Margin="90,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnDBTestConn" Content="Test DB Conn" HorizontalAlignment="Left" Margin="160,15,0,0" VerticalAlignment="Top" Width="100" IsEnabled="False"/>
                    <Rectangle Fill="#FFEFEFF1" HorizontalAlignment="Left" Height="214" Margin="13,50,0,0" Stroke="Black" VerticalAlignment="Top" Width="737"/>
                    <Label x:Name="lblDBServer" Content="DBServer" HorizontalAlignment="Left" Margin="16,104,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtDBServer" HorizontalAlignment="Left" Height="24" Margin="79,105,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="190" Background="White"/>
                    <Label x:Name="lblDBName" Content="DBName" HorizontalAlignment="Left" Margin="16,133,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtDBName" HorizontalAlignment="Left" Height="24" Margin="79,134,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="190" Background="White"/>
                    <Label x:Name="lblDBReportPath" Content="Reports Folder" HorizontalAlignment="Left" Margin="16,188.733,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtDBReportPath" HorizontalAlignment="Left" Height="24" Margin="110,189.735,0,0" TextWrapping="Wrap" IsEnabled="False" VerticalAlignment="Top" Width="475" Background="White"/>
                    <Label x:Name="lblDBLogPath" Content="Logs Folder" HorizontalAlignment="Left" Margin="16,216.733,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtDBLogPath" HorizontalAlignment="Left" Height="24" Margin="110,218.735,0,0" TextWrapping="Wrap" IsEnabled="False" VerticalAlignment="Top" Width="475" Background="White"/>
                    <Label x:Name="lblDBStatus" Content="Database Connection:" HorizontalAlignment="Left" Margin="16,54,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblDBCurrent" Content="Connection Failed" HorizontalAlignment="Left" Margin="140,54,0,0" VerticalAlignment="Top" Foreground="#FFFF0606" FontWeight="Bold"/>
                    <Separator HorizontalAlignment="Left" Height="25.265" Margin="20,76,0,0" VerticalAlignment="Top" Width="722"/>
                    <Separator HorizontalAlignment="Left" Height="25.265" Margin="20,160,0,0" VerticalAlignment="Top" Width="722"/>
                </Grid>
            </TabItem>
        </TabControl>
        <Image x:Name="image" HorizontalAlignment="Left" Height="68" Margin="13,10,0,0" VerticalAlignment="Top" Width="71" Source="C:\Users\Administrator\Pictures\D365Tools.png"/>
        <StatusBar Height="25" VerticalAlignment="Bottom" Width="Auto">
            <StatusBar.ItemsPanel>
                <ItemsPanelTemplate>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto" />
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="Auto" />
                            <ColumnDefinition Width="100" />
                            <ColumnDefinition Width="Auto" />
                        </Grid.ColumnDefinitions>
                    </Grid>
                </ItemsPanelTemplate>
            </StatusBar.ItemsPanel>
            <Separator Grid.Column="0"/>
            <StatusBarItem Grid.Column="1">
                <TextBlock Name="lblWarning"/>
            </StatusBarItem>
            <Separator Grid.Column="2"/>
            <StatusBarItem Grid.Column="3">
                <TextBlock Name="lblControl2"/>
            </StatusBarItem>
            <Separator Grid.Column="4"/>
        </StatusBar>
    </Grid>
</Window>
"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[xml]$XAML = $inputXML

#Read XAML
$Reader=(New-Object System.Xml.XmlNodeReader $XAML)
try{$Form=[Windows.Markup.XamlReader]::Load($Reader)} catch{Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged properties (PowerShell cannot process them)"; Throw}
$XAML.SelectNodes("//*[@Name]") | %{<#"trying item $($_.Name)";#> try {Set-Variable -Name "Wpf$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global -ErrorAction Stop} catch{Throw}}
 
function Get-FormVariables{
    if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
    #write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
    #Get-Variable WPF*
}

Get-FormVariables

function Get-DataSources
{
    Get-EnvironmentsDB
    Get-UsersDB
    Get-EmailsDB
    Get-TasksList
    Get-SettingsXML    
}

function Set-DataSources
{
    $WpfcbxEnvEnvironment.ItemsSource = $Script:EnvironmentDB.Tables[0].DefaultView
    $WpfcbxTskEnvironment.ItemsSource = $Script:EnvironmentDB.Tables[0].DefaultView
    $WpfcbxEnvEmail.ItemsSource = $Script:EmailsDB.Tables[0].DefaultView
    $WpfcbxEmlID.ItemsSource = $Script:EmailsDB.Tables[0].DefaultView
    $WpfcbxEnvLocalUser.ItemsSource = $Script:UsersDB.Tables[0].DefaultView
    $WpfcbxEnvDBUser.ItemsSource = $Script:UsersDB.Tables[0].DefaultView
    $WpfcbxUsrID.ItemsSource = $Script:UsersDB.Tables[0].DefaultView
    $WpfcbxEmlUserID.ItemsSource = $Script:UsersDB.Tables[0].DefaultView
    $WpfcbxTskUserID.ItemsSource = $Script:UsersDB.Tables[0].DefaultView
    $WpfdgXMLSettings.ItemsSource = $Script:SettingsXML.Tables[0].DefaultView
    $WpflstCurrJobs.ItemsSource = @($Script:Tasks)
}

function Get-EnvironmentsDB
{
    $SqlConn = Get-ConnectionString	
	$SqlQuery = "SELECT A.ENVIRONMENT,A.DESCRIPTION,A.DBSERVER,A.DBNAME,A.DBUSER,A.CPUTHOLD,A.BLOCKTHOLD,A.WAITINGTHOLD,A.RUNGRD,A.RUNSTATS,A.EMAILPROFILE,A.LOCALADMINUSER,
                    CASE WHEN B.ENVIRONMENT IS NOT NULL
                      THEN 1
                      ELSE 0
                    END AS AXREFRESH
                    FROM AXTools_Environments A
                    LEFT JOIN AXRefresh_EnvironmentsExt B ON A.ENVIRONMENT = B.ENVIRONMENT"
	$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
	$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$Adapter.SelectCommand = $SqlCommand
	$Script:EnvironmentDB = New-Object System.Data.DataSet
	$Adapter.Fill($Script:EnvironmentDB) | Out-Null
    $WpfcbxEnvEnvironment.ItemsSource = $Script:EnvironmentDB.Tables[0].DefaultView
    $WpfcbxTskEnvironment.ItemsSource = $Script:EnvironmentDB.Tables[0].DefaultView
}

function Get-SettingsXML
{
    $Script:SettingsXML = New-Object System.Data.Dataset
    $Null = $Script:SettingsXML.ReadXml("$ModuleFolder\AX-Settings.xml")
    $WpfdgXMLSettings.ItemsSource = $Script:SettingsXML.Tables[0].DefaultView
}

function Get-UsersDB
{
    $SqlConn = Get-ConnectionString	
	$SqlQuery = "SELECT [ID],[USERNAME] FROM AXTools_UserAccount"
	$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
	$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$Adapter.SelectCommand = $SqlCommand
	$Script:UsersDB = New-Object System.Data.DataSet
	$Adapter.Fill($Script:UsersDB) | Out-Null
    #[void]$Users.Tables[0].Rows.Add('',$null)
    if(![string]::IsNullOrEmpty($Script:UsersDB)) {
        $WpfcbxEnvLocalUser.ItemsSource = $Script:UsersDB.Tables[0].DefaultView
        $WpfcbxEnvDBUser.ItemsSource = $Script:UsersDB.Tables[0].DefaultView
        $WpfcbxUsrID.ItemsSource = $Script:UsersDB.Tables[0].DefaultView
        $WpfcbxEmlUserID.ItemsSource = $Script:UsersDB.Tables[0].DefaultView
        $WpfcbxTskUserID.ItemsSource = $Script:UsersDB.Tables[0].DefaultView
    }
}

function Get-EmailsDB
{
    $SqlConn = Get-ConnectionString	
	$SqlQuery = "SELECT [ID],[USERID],[SMTPSERVER],[SMTPPORT],[SMTPSSL],[FROM],[TO],[CC],[BCC] FROM AXTools_EmailProfile"
	$SqlCommand = New-Object System.Data.SqlClient.SqlCommand ($SqlQuery,$SqlConn)
	$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$Adapter.SelectCommand = $SqlCommand
	$Script:EmailsDB = New-Object System.Data.DataSet
	$Adapter.Fill($Script:EmailsDB) | Out-Null
    #[void]$Emails.Tables[0].Rows.Add('',$null,$null,$null,$null,$null,$null,$null,$null)
    $WpfcbxEnvEmail.ItemsSource = $Script:EmailsDB.Tables[0].DefaultView
    $WpfcbxEmlID.ItemsSource = $Script:EmailsDB.Tables[0].DefaultView
}

function Get-TasksList
{
    try {
        $Script:Tasks = Get-ScheduledTask -TaskPath \DynamicsAxTools\ -ErrorAction SilentlyContinue | 
                        Select @{'n'='Environment'; 'e'={$_.TaskName.Split('-')[0].Trim()}}, 
                        @{'n'='Name'; 'e'={$_.TaskName.Split('-')[1].Trim()}} ,  
                        @{'n'='Interval'; 'e'={if([string]::IsNullOrEmpty($_.Triggers.Repetition.Interval)) { '0' }
                                                    else{ $(Get-TasksInterval $_.Triggers.Repetition.Interval)}}}, 
                        @{'n'='User'; 'e'={$_.Principal.UserId}}, 
                        @{'n'='DaysInterval'; 'e'={$_.Triggers.DaysInterval}}, 
                        @{'n'='At'; 'e'={if(![string]::IsNullOrEmpty($_.Triggers.DaysInterval)) {([datetime]$_.Triggers.StartBoundary).ToShortTimeString()}}}, 
                        State,
                        @{'n'='NextRunTime'; 'e'={(Get-ScheduledTaskInfo -InputObject $_).NextRunTime}},
                        @{'n'='LastRunTime'; 'e'={(Get-ScheduledTaskInfo -InputObject $_).LastRunTime}}, 
                        TaskName, 
                        TaskPath
    }
    catch {
        $Script:Tasks = @()
    }
    #return $Tasks
    $WpflstCurrJobs.ItemsSource = @($Script:Tasks)
}

function Get-TasksFolder {
    $ScheduleObject = New-Object -ComObject Schedule.Service
    $ScheduleObject.Connect()
    try {
        $Folder = $ScheduleObject.GetFolder("\DynamicsAxTools")
    }
    catch [System.Runtime.InteropServices.COMException] {
        $WpflblWarning.Text = "Connected to Tasks Folder."
    }
    catch [System.IO.FileNotFoundException] {
        $RootFolder = $ScheduleObject.GetFolder("\")
        $RootFolder.CreateFolder("DynamicsAxTools")
        $WpflblWarning.Text = "Tasks Folder Created."
    }
}

function Get-EmailClear
{
    $WpfbtnEmlEdit.IsEnabled = $False
    $WpfbtnEmlSave.IsEnabled = $False
    $WpfbtnEmlDelete.IsEnabled = $False
    $WpfbtnEmlTest.IsEnabled = $False
    $WpftxtEmlSMTP.Clear()
    $WpftxtEmlSMTPPort.Clear()
    $WpfchkEmlSSL.IsChecked = $false
    $WpfcbxEmlUserID.SelectedItem = $null
    $WpftxtEmlFrom.Clear()
    $WpftxtEmlTo.Clear()
    $WpftxtEmlCC.Clear()
    $WpftxtEmlBCC.Clear()
}

function Get-UsersClear
{
    $WpftxtUsrUsername.Clear()
    $WpfbtnUsrDelete.IsEnabled = $false
    $WpfbtnUsrTest.IsEnabled = $false
}

function Get-TasksClear
{
    $WpfbtnTskEdit.IsEnabled = $false
    $WpfbtnTskDelete.IsEnabled = $false
    $WpfbtnTskSave.IsEnabled = $false
    $WpfbtnTskDisable.IsEnabled = $false
    $WpfbtnTskEnable.IsEnabled = $false
    $WpfcbxTskEnvironment.SelectedItem = $null
    $WpfcbxTskTaskName.SelectedItem = $null
    $WpftxtTskInterval.Clear()
    $WpftxtTskTime.Clear()
    $WpfcbxTskUserID.Clear()
}

function Get-TabItemClear
{
    switch($WpftabControl.SelectedIndex) {
        '0' {
            Get-EnvironmentsDB
            $WpfcbxEnvEnvironment.IsEditable = $false
            $WpfbtnEnvEdit.IsEnabled = $false
            $WpfbtnEnvDelete.IsEnabled = $false
            $WpfbtnEnvSave.IsEnabled = $false
            $WpfbtnEnvtestSQL.IsEnabled = $false
            $WpfbtnEnvReports.IsEnabled = $false
            $WpfbtnEnvLogs.IsEnabled = $false
            $WpftxtEnvEnvironment.IsEnabled = $false
            $WpfchkEnvRefresh.IsEnabled = $false
            $WpftxtEnvDBServer.IsEnabled = $false
            $WpftxtEnvDBName.IsEnabled = $false
            $WpfchkEnvGRD.IsEnabled = $false
            $WpftxtEnvCPU.IsEnabled = $false
            $WpftxtEnvBlocking.IsEnabled = $false
            $WpftxtEnvWaiting.IsEnabled = $false
            $WpfcbxEnvEnvironment.SelectedItem = $null
            $WpftxtEnvEnvironment.Clear()
            $WpfcbxEnvEmail.SelectedItem = $null
            $WpfcbxEnvLocalUser.SelectedItem = $null
            $WpfchkEnvRefresh.IsChecked = $false
            $WpfcbxEnvDBUser.SelectedItem = $null
            $WpftxtEnvDBServer.Clear()
            $WpftxtEnvDBName.Clear()
            $WpfcbxEnvDBStats.SelectedItem = $null
            $WpfchkEnvGRD.IsChecked = $false
            $WpftxtEnvCPU.Clear()
            $WpftxtEnvBlocking.Clear()
            $WpftxtEnvWaiting.Clear()
        }        
        '1' {
            Get-UsersDB
            $WpftxtUsrUsername.Clear()
            $WpfbtnUsrDelete.IsEnabled = $false
            $WpfbtnUsrTest.IsEnabled = $false
        }
        '2' {
            Get-EmailsDB
            $WpfcbxEmlID.IsEditable = $false
            $WpfbtnEmlEdit.IsEnabled = $false
            $WpfbtnEmlSave.IsEnabled = $false
            $WpfbtnEmlDelete.IsEnabled = $false
            $WpfbtnEmlTest.IsEnabled = $false
            $WpfchkEmlSSL.IsChecked = $false
            $WpftxtEmlSMTP.IsEnabled = $false
            $WpftxtEmlSMTPPort.IsEnabled = $false
            $WpfcbxEmlUserID.IsEnabled = $false
            $WpfchkEmlSSL.IsEnabled = $false
            $WpftxtEmlFrom.IsEnabled = $false
            $WpftxtEmlTo.IsEnabled = $false
            $WpftxtEmlCC.IsEnabled = $false
            $WpftxtEmlBCC.IsEnabled = $false
            $WpftxtEmlSMTP.Clear()
            $WpftxtEmlSMTPPort.Clear()
            $WpfcbxEmlUserID.SelectedItem = $null
            $WpftxtEmlFrom.Clear()
            $WpftxtEmlTo.Clear()
            $WpftxtEmlCC.Clear()
            $WpftxtEmlBCC.Clear()
        }
        '3' {
            Get-TasksList
            $WpfbtnTskEdit.IsEnabled = $false
            $WpfbtnTskDelete.IsEnabled = $false
            $WpfbtnTskSave.IsEnabled = $false
            $WpfbtnTskDisable.IsEnabled = $false
            $WpfbtnTskEnable.IsEnabled = $false
            $WpfcbxTskEnvironment.SelectedItem = $null
            $WpfcbxTskTaskName.SelectedItem = $null
            $WpftxtTskInterval.Clear()
            $WpftxtTskTime.Clear()
            $WpfcbxTskUserID.SelectedItem = $null
        }
        '4' {

        }
        '5' {

        }
        '6' {
            Get-SettingsXML
        }
        '7' {
            $Srv = New-Object ('Microsoft.SqlServer.Management.SMO.Server') $WpftxtDBServer.Text
            if($Srv.Databases | Where { $_.Name -eq $WpftxtDBName.Text }) {
                $WpftabControl.Items[0].IsEnabled = $true
                $WpftabControl.Items[1].IsEnabled = $true
                $WpftabControl.Items[2].IsEnabled = $true
                $WpftabControl.Items[3].IsEnabled = $true
                $WpftabControl.Items[4].IsEnabled = $true
                $WpftabControl.Items[5].IsEnabled = $true
                $WpftabControl.Items[6].IsEnabled = $true
                $WpflblDBCurrent.Content = 'Connection Successful'
                $WpflblDBCurrent.Foreground = '#00802b'
                $WpfbtnDBCreate.IsEnabled = $false
                $WpfbtnDBDrop.IsEnabled = $true
                $WpfbtnDBTestConn.IsEnabled = $true
                $WpftxtDBReportPath.Text = ((Import-ConfigFile).ReportFolder)
                $WpftxtDBLogPath.Text = ((Import-ConfigFile).LogFolder)
                Get-SettingsXML
            }
            else {
                $WpftabControl.Items[0].IsEnabled = $false
                $WpftabControl.Items[1].IsEnabled = $false
                $WpftabControl.Items[2].IsEnabled = $false
                $WpftabControl.Items[3].IsEnabled = $false
                $WpftabControl.Items[4].IsEnabled = $false
                $WpftabControl.Items[5].IsEnabled = $false
                $WpftabControl.Items[6].IsEnabled = $false
                $WpflblDBCurrent.Content = 'Connection Failed'
                $WpflblDBCurrent.Foreground = '#FFFF0606'
                $WpftxtDBServer.Clear()
                $WpftxtDBName.Clear()
                $WpftxtDBServer.IsEnabled = $true
                $WpftxtDBName.IsEnabled = $true
                $WpfbtnDBCreate.IsEnabled = $true
                $WpfbtnDBDrop.IsEnabled = $false
                $WpfbtnDBTestConn.IsEnabled = $false
                $WpftxtDBReportPath.Text = ((Import-ConfigFile).ReportFolder)
                $WpftxtDBLogPath.Text = ((Import-ConfigFile).LogFolder)
            }
        }
    }
}

function Get-DisableAll
{
    $WpftxtEmlSMTP.IsEnabled = $false
    $WpftxtEmlSMTPPort.IsEnabled = $false
    $WpfcbxEmlUserID.IsEnabled = $false
    $WpfchkEmlSSL.IsEnabled = $false
    $WpftxtEmlBCC.IsEnabled = $false
    $WpftxtEmlCC.IsEnabled = $false
    $WpftxtEmlFrom.IsEnabled = $false
    $WpftxtEmlSMTP.IsEnabled = $false
    $WpftxtEmlSMTPPort.IsEnabled = $false
    $WpftxtEmlTo.IsEnabled = $false
    $WpftxtEnvDBName.IsEnabled = $false
    $WpfcbxEnvDBStats.IsEnabled = $false
    $WpfcbxEnvDBUser.IsEnabled = $false
    $WpfcbxEnvEmail.IsEnabled = $false
    $WpfcbxEnvLocalUser.IsEnabled = $false
    $WpfchkEnvRefresh.IsEnabled = $false
    $WpftxtEnvBlocking.IsEnabled = $false
    $WpftxtEnvCPU.IsEnabled = $false
    $WpftxtEnvDBServer.IsEnabled = $false
    $WpftxtEnvEnvironment.IsEnabled = $false
    $WpftxtEnvWaiting.IsEnabled = $false
    $WpfchkEnvGRD.IsEnabled = $false
    #$WpfcbxTskEnvironment.IsEnabled = $false
    #$WpfcbxTskInterval.IsEnabled = $false
    #$WpfcbxTskTaskName.IsEnabled = $false
    #$WpfckcTskRunAs.IsEnabled = $false
    #$WpflstCurrJobs.IsEnabled = $false
    #$WpftxtTskTimeSpan.IsEnabled = $false
    $WpftxtUsrUsername.IsEnabled = $false
}

function Validate-User
{
    $Query = "SELECT UserName FROM [dbo].[AXTools_UserAccount] WHERE [USERNAME] = '$UserName'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $UserExists = $Cmd.ExecuteScalar()
    if([string]::IsNullOrEmpty($UserExists)) {
        return $false
    }
    else {
        return $true
    }
}

function Delete-User
{
    $Conn = Get-ConnectionString
    $Query = "DELETE FROM [dbo].[AXTools_UserAccount] WHERE [USERNAME] = '$UserName'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery() | Out-Null
    $Conn.Close()
}

function Delete-Email
{
[CmdletBinding()]
param (
    [String]$Id
)
    $Conn = Get-ConnectionString
    $Query = "DELETE FROM [dbo].[AXTools_EmailProfile] WHERE [ID] = '$Id'"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery() | Out-Null
    $Conn.Close()
}

function Delete-Environment
{
[CmdletBinding()]
param (
    [String]$Id
)
	$Conn = Get-ConnectionString
    $Query = "SELECT COUNT(1) FROM AXRefresh_EnvironmentStore WHERE ENVIRONMENT = '$($Id)' AND DELETED = 0"
	$Cmd = New-Object System.Data.SqlClient.SQLCommand ($Query,$Conn)
	[boolean]$HasStore = $Cmd.ExecuteScalar()
    if($HasStore) {
        New-Popup -Message "There is a backup for $Id." -Title "Canceled" -Buttons OK -Icon Stop
        $WpflblWarning.Text = "Canceled."
    }
    else {
        $Query = “DELETE FROM [AXTools_Environments] WHERE ENVIRONMENT = '$Id'"
	    $Cmd = New-Object System.Data.SqlClient.SQLCommand ($Query,$Conn)
	    $Cmd.ExecuteNonQuery() | Out-Null
        #
	    $Query = “DELETE FROM [AXRefresh_EnvironmentsExt] WHERE ENVIRONMENT = '$Id'"
	    $Cmd = New-Object System.Data.SqlClient.SQLCommand ($Query,$Conn)
	    $Cmd.ExecuteNonQuery() | Out-Null
	    #
	    $Query = “DELETE FROM [AXTools_Servers] WHERE ENVIRONMENT = '$Id'"
	    $Cmd = New-Object System.Data.SqlClient.SQLCommand ($Query,$Conn)
	    $Cmd.ExecuteNonQuery() | Out-Null
	    #
	    $Query = “DELETE FROM [AXRefresh_EnvironmentStore] WHERE ENVIRONMENT = '$Id'"
	    $Cmd = New-Object System.Data.SqlClient.SQLCommand ($Query,$Conn)
	    $Cmd.ExecuteNonQuery() | Out-Null


        $WpflblWarning.Text = "Completed."
    }
    $Conn.Close()
}

function Insert-User
{
    $SecureStringAsPlainText = Write-EncryptedString -InputString $Credential.GetNetworkCredential().Password -DTKey "$((Get-WMIObject Win32_Bios).PSComputerName)-$((Get-WMIObject Win32_Bios).SerialNumber)"
    $Query = "INSERT INTO [dbo].[AXTools_UserAccount] ([ID],[USERNAME],[PASSWORD])
                VALUES ('$ID','$UserName','$SecureStringAsPlainText')"
    $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
    $Cmd.ExecuteNonQuery() | Out-Null
    if(![string]::IsNullOrEmpty($RunAs)) {
        [xml]$ConfigFile = Get-Content "$ModuleFolder\AX-Settings.xml"
        $ConfigFile.Settings.Database.UserName = $UserName
        $ConfigFile.Settings.Database.Password = $SecureStringAsPlainText.ToString()
        $ConfigFile.Save("$ModuleFolder\AX-Settings.xml")
    }
}

function Get-TasksInterval
{
param(
    $Interval
)
    if( $Interval -match 'PT(\d+)(.*)$' )
    {
        $modifier = $Matches[1]
        $unit = $Matches[2]
        $hour = 0
        $minute = 0
        $second = 0
        switch($unit) {
            'H' { $hour = $modifier }
            'M' { $minute = $modifier }
        }
        $timespan = New-Object 'TimeSpan' $hour,$minute,$second
        return $timespan.TotalMinutes
    }
}

function Get-TaskChanges
{
param(
    [string]$TaskName
)
    $RegisteredTask = Get-ScheduledTask -TaskName $TaskName
    #Check Interval
    if((![string]::IsNullOrEmpty($WpftxtTskInterval.Text)) -and ((Get-TasksInterval $RegisteredTask.Triggers.Repetition.Interval) -ne $WpftxtTskInterval.Text)) {
            $RegisteredTask.Triggers[0].Repetition.Interval = "PT$($WpftxtTskInterval.Text)M"
            $RegisteredTask.Settings[0].ExecutionTimeLimit = "PT$($WpftxtTskInterval.Text)M"
    }
    #Check User/Password
    if($WpfcbxTskUserID.SelectedIndex -eq -1) {
        $Credential = Get-Credential -Credential "$env:userdomain\$env:username"
    }
    else {
        $Credential = Get-UserCredentials -Account $WpfcbxTskUserID.SelectedItem.Id
        if(($RegisteredTask.Principal.Id).ToUpper() -ne $Credential.Username) {
            $RegisteredTask.Principal[0].Id = $Credential.Username
            $RegisteredTask.Principal[0].UserId = $Credential.Username                    
        }
    }
    #Check Start Time
    if((![string]::IsNullOrEmpty($WpftxtTskTime.Text)) -and ([DateTime]::Parse($RegisteredTask.Triggers.StartBoundary).ToShortTimeString()) -ne ([DateTime]::Parse($WpftxtTskTime.Text).ToShortTimeString())) {
        $RegisteredTask.Triggers[0].StartBoundary = [DateTime]::Parse($WpftxtTskTime.Text)
    }
    #Change Task
    $RegisteredTask.Description = "$($RegisteredTask.Description) `r`nChanged: $(Get-Date -Format G) - $env:USERDOMAIN\$env:USERNAME"
    $RegisteredTask | Set-ScheduledTask -User $Credential.Username -Password $Credential.GetNetworkCredential().Password
}

function Get-ChkTaskExists
{
param(
    [string]$TaskName
)
    if(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        return $true
    }
    else {
        return $false
    }
}

function Get-TaskSettings
{
param(
    [string]$TaskName
)
    $RegisteredTask = Get-ScheduledTask -TaskName $TaskName
}

function Get-ListFilter
{
    $Script:View = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WpflstCurrJobs.ItemsSource)
    if($WpfcbxTskEnvironment.SelectedIndex -ne -1) { $FilterEnvironment = $WpfcbxTskEnvironment.SelectedItem.Environment } else { $FilterEnvironment = '' }
    if($WpfcbxTskTaskName.SelectedIndex -ne -1) { $FilterTask = $WpfcbxTskTaskName.SelectedItem.Value } else { $FilterTask = '' }
    $Script:View.Filter = {param ($Item) ($Item.Environment -match $FilterEnvironment) -and ($Item.Name -match $FilterTask)}
    $Script:View.Refresh()
}

function Check-UserPassword
{
param(
    [System.Management.Automation.PSCredential]$Credential
)
    $BSTRSetup = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
    $UserPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTRSetup)
    $Root = "LDAP://" + ([ADSI]"").distinguishedName
    $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$Credential.UserName,$UserPassword)
    if ($Domain.Name -eq $null) {
        $WpflblWarning.Text = "$($Credential.UserName) - Authentication failed."
    }
    else {
        $WpflblWarning.Text = "$($Credential.UserName) - Successfully authenticated."
    }
}

#===========================================================================
# Form Selection Control
#===========================================================================

$WpfcbxEnvEnvironment.Add_SelectionChanged({
    if($WpfcbxEnvEnvironment.SelectedIndex -ne -1) {
        if($WpftxtEnvDBServer.Text -ne '' -and $WpftxtEnvDBName.Text -ne ''){
            $WpfbtnEnvtestSQL.IsEnabled = $true
        }
        else {
            $WpfbtnEnvtestSQL.IsEnabled = $false
        }
        $WpfbtnEnvReports.IsEnabled = $true
        $WpfbtnEnvLogs.IsEnabled = $true
        $WpfbtnEnvEdit.IsEnabled = $true
        $WpfbtnEnvDelete.IsEnabled = $true
        $WpftxtEnvEnvironment.Text = $WpfcbxEnvEnvironment.SelectedItem["Description"]
        $WpfcbxEnvEmail.SelectedItem = $WpfcbxEnvEmail.Items | Where { $_.ID -eq $WpfcbxEnvEnvironment.SelectedItem["EmailProfile"] }
        $WpfcbxEnvLocalUser.SelectedItem = $WpfcbxEnvLocalUser.Items | Where { $_.ID -eq $WpfcbxEnvEnvironment.SelectedItem["LocalAdminUser"] }
        if(($WpfcbxEnvEnvironment.SelectedItem["AxRefresh"] -eq 1)) { $WpfchkEnvRefresh.IsChecked = $true }
        $WpftxtEnvDBServer.Text = $WpfcbxEnvEnvironment.SelectedItem["DBServer"]
        $WpfcbxEnvDBUser.SelectedItem = $WpfcbxEnvDBUser.Items | Where { $_.ID -eq $WpfcbxEnvEnvironment.SelectedItem["DBUser"] }
        $WpftxtEnvDBName.Text = $WpfcbxEnvEnvironment.SelectedItem["DBName"]
        $WpfcbxEnvDBStats.SelectedItem = $WpfcbxEnvDBStats.Items | Where { $_.Name -eq $WpfcbxEnvEnvironment.SelectedItem["RunStats"] }
        if(($WpfcbxEnvEnvironment.SelectedItem["RunGrd"] -eq 1)) { $WpfchkEnvGRD.IsChecked = $true }
        $WpftxtEnvCPU.Text = $WpfcbxEnvEnvironment.SelectedItem["CpuThold"]
        $WpftxtEnvBlocking.Text = $WpfcbxEnvEnvironment.SelectedItem["BlockThold"]
        $WpftxtEnvWaiting.Text = $WpfcbxEnvEnvironment.SelectedItem["WaitingThold"]
    }
})

$WpfcbxUsrID.Add_SelectionChanged({
    if($WpfcbxUsrID.SelectedIndex -ne -1) {
        $WpftxtUsrUsername.Text = $WpfcbxUsrID.SelectedItem["UserName"]
        $WpfbtnUsrDelete.IsEnabled = $true
        $WpfbtnUsrTest.IsEnabled = $true
    }
    else {
        Get-UsersClear
    }
})

$WpfcbxEmlID.Add_SelectionChanged({
    if($WpfcbxEmlID.SelectedIndex -ne -1) {
        $WpftxtUsrUsername.Text = $WpfcbxEmlID.SelectedItem["Username"]
        $WpftxtEmlSMTP.Text = $WpfcbxEmlID.SelectedItem["SmtpServer"]
        $WpftxtEmlSMTPPort.Text = $WpfcbxEmlID.SelectedItem["SmtpPort"]
        if(($WpfcbxEmlID.SelectedItem["SmtpSSL"] -eq 1)) { $WpfchkEmlSSL.IsChecked = $true }
        $WpfcbxEmlUserID.SelectedItem = $WpfcbxEmlUserID.Items | Where { $_.ID -eq $WpfcbxEmlID.SelectedItem["UserId"] }
        $WpftxtEmlFrom.Text = $WpfcbxEmlID.SelectedItem["From"]
        $WpftxtEmlTo.Text = $WpfcbxEmlID.SelectedItem["To"]
        $WpftxtEmlCC.Text = $WpfcbxEmlID.SelectedItem["CC"]
        $WpftxtEmlBCC.Text = $WpfcbxEmlID.SelectedItem["BCC"]
        $WpfbtnEmlEdit.IsEnabled = $true
        $WpfbtnEmlDelete.IsEnabled = $true
        $WpfbtnEmlTest.IsEnabled = $true
    }
    else {
        Get-EmailClear
    }
})

$WpflstCurrJobs.Add_SelectionChanged({
    if($WpflstCurrJobs.SelectedIndex -eq -1) {
        $WpfbtnTskDisable.IsEnabled = $false
        $WpfbtnTskEnable.IsEnabled = $false
        $WpfbtnTskDelete.IsEnabled = $false
        $WpfbtnTskEdit.IsEnabled = $false
        $WpfbtnTskSave.IsEnabled = $false
    }
    else {
        if($WpflstCurrJobs.SelectedItem.State -like 'Disabled') {
            $WpfbtnTskDisable.IsEnabled = $false
            $WpfbtnTskEnable.IsEnabled = $true
            $WpfbtnTskDelete.IsEnabled = $true
            $WpfbtnTskEdit.IsEnabled = $true

        }
        elseif($WpflstCurrJobs.SelectedItem.State -like 'Ready') {
            $WpfbtnTskDisable.IsEnabled = $true
            $WpfbtnTskEnable.IsEnabled = $false
            $WpfbtnTskDelete.IsEnabled = $true
            $WpfbtnTskEdit.IsEnabled = $true    
        }
    }
})

$WpfcbxTskEnvironment.Add_SelectionChanged({
    Get-ListFilter
})

$WpfcbxTskTaskName.Add_SelectionChanged({
    Get-ListFilter
    if($WpfbtnTskSave.IsEnabled -eq $true) {
        if($WpfcbxTskTaskName.SelectedItem.Value -match "AX Monitor|Check AOS|Check Perfmon") {
            $WpftxtTskInterval.IsEnabled = $true
            $WpftxtTskTime.IsEnabled = $false
            $WpfcbxTskUserID.IsEnabled = $true
        }
        else {
            $WpftxtTskInterval.IsEnabled = $false
            $WpftxtTskTime.IsEnabled = $true
            $WpfcbxTskUserID.IsEnabled = $true
        }
    }
})

#===========================================================================
# Form Button New Click
#===========================================================================

$WpfbtnEnvNew.Add_Click({
    $WpfcbxEnvEnvironment.IsEditable = $true
    $WpfbtnEnvSave.IsEnabled = $true
    $WpfbtnEnvtestSQL.IsEnabled = $false
    $WpfchkEnvRefresh.IsChecked = $false
    $WpfchkEnvGRD.IsChecked= $false
    $WpfcbxEnvEnvironment.SelectedIndex = -1
    $WpfcbxEnvEmail.SelectedIndex = -1
    $WpfcbxEnvLocalUser.SelectedIndex = -1
    $WpfcbxEnvDBUser.SelectedIndex = -1
    $WpfcbxEnvDBStats.SelectedIndex = -1
    $WpftxtEnvEnvironment.Clear()
    $WpftxtEnvDBServer.Clear()
    $WpftxtEnvDBName.Clear()
    $WpftxtEnvCPU.Clear()
    $WpftxtEnvBlocking.Clear()
    $WpftxtEnvWaiting.Clear()
    $WpfchkEnvRefresh.IsEnabled = $false
    $WpfchkEnvGRD.IsEnabled = $false
    $WpfcbxEnvEmail.IsEnabled = $true
    $WpfcbxEnvLocalUser.IsEnabled = $true
    $WpfcbxEnvDBUser.IsEnabled = $true
    $WpfcbxEnvDBStats.IsEnabled = $true
    $WpftxtEnvEnvironment.IsEnabled = $true
    $WpftxtEnvDBServer.IsEnabled = $true
    $WpftxtEnvDBName.IsEnabled = $true
    $WpftxtEnvCPU.IsEnabled = $true
    $WpftxtEnvBlocking.IsEnabled = $true
    $WpftxtEnvWaiting.IsEnabled = $true
    $WpfchkEnvRefresh.IsEnabled = $true
    $WpfchkEnvGRD.IsEnabled= $true
})

$WpfbtnUsrNew.Add_Click({
    $WpfcbxUsrID.SelectedIndex = -1
    $WpftxtUsrUsername.Clear()
    [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty
    $Credential = Get-Credential -Message "<DOMAIN\Username> OR <user@emailserver.com>" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if($Credential.UserName -ne $null) {
        $Conn = Get-ConnectionString
        $BSTRBC = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
        $Root = "LDAP://" + ([ADSI]"").distinguishedName
        $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$Credential.UserName,[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTRBC))
        if ($Domain.Name -eq $null) {
            $WpflblWarning.Text = "This is not a domain account."
            $ID = "$($Credential.UserName.Split('@')[0].toUpper())"
            $UserName = "$($Credential.UserName)"
            if(Validate-User) {
                Delete-User
                Insert-User
                $WpflblWarning.Text = "Username already exists. Updating current credentials."
            }
            else {
                Insert-User
                $WpflblWarning.Text = "Completed."
            }
        }
        else {
            $WpflblWarning.Text =  "Domain successfully authenticated."
            $ID = "$($Credential.UserName.Split('\')[1].ToUpper())"
            $UserName = $Credential.UserName.ToUpper()
            if(Validate-User) {
                Delete-User
                Insert-User
                $WpflblWarning.Text = "Username already exists. Updating current credentials."
            }
            else {
                Insert-User
                
            }
        }
        Get-UsersDB
        $WpfcbxUsrID.SelectedItem = $WpfcbxUsrID.ItemsSource | Where { $_.Username -eq $Credential.UserName }
    }
    else {
        Get-TabItemClear
        $WpflblWarning.Text = "Canceled." 
    }
})

$WpfbtnEmlNew.Add_Click({
    $WpfbtnEmlSave.IsEnabled = $True
    $WpfcbxEmlID.SelectedIndex = -1
    $WpfcbxEmlID.IsEditable = $True
    $WpftxtEmlCC.Clear()
    $WpftxtEmlFrom.Clear()
    $WpftxtEmlSMTP.Clear()
    $WpftxtEmlSMTPPort.Clear()
    $WpftxtEmlTo.Clear()
    $WpfchkEmlSSL.IsChecked = $false
    $WpftxtEmlSMTP.Clear()
    $WpftxtEmlSMTPPort.Clear()
    $WpfcbxEmlUserID.SelectedIndex = -1
    $WpftxtEmlBCC.Clear()
    $WpftxtEmlCC.IsEnabled = $True
    $WpftxtEmlFrom.IsEnabled = $True
    $WpftxtEmlSMTP.IsEnabled = $True
    $WpftxtEmlSMTPPort.IsEnabled = $True
    $WpftxtEmlTo.IsEnabled = $True
    $WpfchkEmlSSL.IsEnabled = $True
    $WpftxtEmlSMTP.IsEnabled = $True
    $WpftxtEmlSMTPPort.IsEnabled = $True
    $WpfcbxEmlUserID.IsEnabled = $True
})

$WpfbtnTskNew.Add_Click({
    $WpfcbxTskEnvironment.SelectedIndex = -1
    $WpfcbxTskTaskName.SelectedIndex = -1
    $WpflstCurrJobs.SelectedIndex = -1
    $WpfbtnTskDelete.IsEnabled = $false
    $WpfbtnTskDisable.IsEnabled = $false
    $WpfbtnTskEnable.IsEnabled = $false
    $WpfbtnTskSave.IsEnabled = $true
})

#===========================================================================
# Form Button Edit Click
#===========================================================================

$WpfbtnEnvEdit.Add_Click({
    if($WpfcbxEnvEnvironment.SelectedIndex -ne -1) {
        $WpfbtnEnvSave.IsEnabled = $true
        $WpfchkEnvRefresh.IsEnabled = $false
        $WpfchkEnvGRD.IsEnabled = $true
        $WpfcbxEnvEmail.IsEnabled = $true
        $WpfcbxEnvLocalUser.IsEnabled = $true
        $WpfcbxEnvDBUser.IsEnabled = $true
        $WpfcbxEnvDBStats.IsEnabled = $true
        $WpftxtEnvEnvironment.IsEnabled = $true
        $WpftxtEnvDBServer.IsEnabled = $true
        $WpftxtEnvDBName.IsEnabled = $true
        $WpftxtEnvCPU.IsEnabled = $true
        $WpftxtEnvBlocking.IsEnabled = $true
        $WpftxtEnvWaiting.IsEnabled = $true
    }
    else {
        $WpflblWarning.Text = "Nothing to Edit."
    }
})

$WpfbtnEmlEdit.Add_Click({
    if($WpfcbxEmlID.SelectedIndex -ne -1) {
        $WpfbtnEmlSave.IsEnabled = $True
        $WpftxtEmlBCC.IsEnabled = $True
        $WpftxtEmlCC.IsEnabled = $True
        $WpftxtEmlFrom.IsEnabled = $True
        $WpftxtEmlSMTP.IsEnabled = $True
        $WpftxtEmlSMTPPort.IsEnabled = $True
        $WpftxtEmlTo.IsEnabled = $True
        $WpfchkEmlSSL.IsEnabled = $True
        $WpftxtEmlSMTP.IsEnabled = $True
        $WpftxtEmlSMTPPort.IsEnabled = $True
        $WpfcbxEmlUserID.IsEnabled = $True
    }
    else {
        $WpflblWarning.Text = "Nothing to Edit."
    }
})

$WpfbtnTskEdit.Add_Click({
    $WpfbtnTskSave.IsEnabled = $true
    $WpfcbxTskEnvironment.SelectedItem = $WpfcbxTskEnvironment.Items | Where { $_.Environment -eq $WpflstCurrJobs.SelectedItem.Environment }
    $WpfcbxTskTaskName.SelectedItem = $WpfcbxTskTaskName.Items | Where { $_.Value -eq $WpflstCurrJobs.SelectedItem.Name }
    $TaskUser = Get-ScheduledTask -TaskName $WpflstCurrJobs.SelectedItem.TaskName
    if($WpfcbxTskUserID.Items | Where { $_.UserName -eq $TaskUser.Principal.Id }) {
        $WpfcbxTskUserID.SelectedItem = $WpfcbxTskUserID.Items | Where { $_.UserName -eq $TaskUser.Principal.Id }
    }
})

#===========================================================================
# Form Button Save Click
#===========================================================================

$WpfbtnEnvSave.Add_Click({
    if($WpfcbxEnvEnvironment.SelectedIndex -eq -1) {
        $Conn = Get-ConnectionString
        $Query = "INSERT INTO [dbo].[AXTools_Environments] ([ENVIRONMENT],[DESCRIPTION],[DBSERVER],[DBNAME],[DBUSER],[CPUTHOLD],[BLOCKTHOLD],[WAITINGTHOLD],[RUNGRD],[RUNSTATS],[EMAILPROFILE],[LOCALADMINUSER])
                  VALUES ('$($WpfcbxEnvEnvironment.Text)','$($WpftxtEnvEnvironment.Text)','$($WpftxtEnvDBServer.Text)','$($WpftxtEnvDBName.Text)','$($WpfcbxEnvDBUser.Text)','$($WpftxtEnvCPU.Text)','$($WpftxtEnvBlocking.Text)','$($WpftxtEnvWaiting.Text)','$(if($WpfchkEnvGRD.IsChecked){'1'} else{'0'})','$($WpfcbxEnvDBStats.SelectedItem.Name)','$($WpfcbxEnvEmail.Text)','$($WpfcbxEnvLocalUser.Text)')"
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Cmd.ExecuteNonQuery() | Out-Null

        if($WpfchkEnvRefresh.IsChecked) {
            $Query = "INSERT INTO [dbo].[AXRefresh_EnvironmentsExt] ([ENVIRONMENT]) VALUES ('$($WpfcbxEnvEnvironment.Text)')"
            $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
            $Cmd.ExecuteNonQuery() | Out-Null
        }
        $Conn.Close()
        $WpflblWarning.Text = "Saved."
        $WpfbtnEnvSave.IsEnabled = $false
        $WpfcbxEnvEnvironment.IsEditable = $false
        Get-TabItemClear
        $WpfcbxEnvEnvironment.SelectedIndex = ($WpfcbxEnvEnvironment.Items.Count - 1)
    }
    else {
        $Original = New-Object PSObject
        $Original | Add-Member NoteProperty 'ENVIRONMENT' $WpfcbxEnvEnvironment.SelectedItem.ENVIRONMENT
        $Original | Add-Member NoteProperty 'DESCRIPTION' $WpfcbxEnvEnvironment.SelectedItem.DESCRIPTION
        $Original | Add-Member NoteProperty 'DBSERVER' $WpfcbxEnvEnvironment.SelectedItem.DBSERVER
        $Original | Add-Member NoteProperty 'DBNAME' $WpfcbxEnvEnvironment.SelectedItem.DBNAME
        $Original | Add-Member NoteProperty 'DBUSER' $WpfcbxEnvEnvironment.SelectedItem.DBUSER
        $Original | Add-Member NoteProperty 'CPUTHOLD' $WpfcbxEnvEnvironment.SelectedItem.CPUTHOLD
        $Original | Add-Member NoteProperty 'BLOCKTHOLD' $WpfcbxEnvEnvironment.SelectedItem.BLOCKTHOLD
        $Original | Add-Member NoteProperty 'WAITINGTHOLD' $WpfcbxEnvEnvironment.SelectedItem.WAITINGTHOLD
        $Original | Add-Member NoteProperty 'RUNGRD' $WpfcbxEnvEnvironment.SelectedItem.RUNGRD
        $Original | Add-Member NoteProperty 'RUNSTATS' $WpfcbxEnvEnvironment.SelectedItem.RUNSTATS
        $Original | Add-Member NoteProperty 'EMAILPROFILE' $WpfcbxEnvEnvironment.SelectedItem.EMAILPROFILE
        $Original | Add-Member NoteProperty 'LOCALADMINUSER' $WpfcbxEnvEnvironment.SelectedItem.LOCALADMINUSER

        $Changed = New-Object PSObject
        $Changed | Add-Member NoteProperty 'ENVIRONMENT' $WpfcbxEnvEnvironment.Text
        $Changed | Add-Member NoteProperty 'DESCRIPTION' $WpftxtEnvEnvironment.Text
        $Changed | Add-Member NoteProperty 'DBSERVER' $WpftxtEnvDBServer.Text
        $Changed | Add-Member NoteProperty 'DBNAME' $WpftxtEnvDBName.Text
        $Changed | Add-Member NoteProperty 'DBUSER' $WpfcbxEnvDBUser.Text
        $Changed | Add-Member NoteProperty 'CPUTHOLD' $WpftxtEnvCPU.Text
        $Changed | Add-Member NoteProperty 'BLOCKTHOLD' $WpftxtEnvBlocking.Text
        $Changed | Add-Member NoteProperty 'WAITINGTHOLD' $WpftxtEnvWaiting.Text
        $Changed | Add-Member NoteProperty 'RUNGRD' $(if($WpfchkEnvGRD.IsChecked){'1'} else{'0'})
        $Changed | Add-Member NoteProperty 'RUNSTATS' $WpfcbxEnvDBStats.SelectedItem.Name
        $Changed | Add-Member NoteProperty 'EMAILPROFILE' $WpfcbxEnvEmail.Text
        $Changed | Add-Member NoteProperty 'LOCALADMINUSER' $WpfcbxEnvLocalUser.Text

        $Properties = $Original.PsObject.Properties.Name
        $ObjectChange = $false
        $Conn = Get-ConnectionString
        foreach ($Property in $Properties) {
            $ValueChange = Compare-Object $Original $Changed -Property $Property
            if($ValueChange.Count -eq 2 -and $ValueChange[0] -ne $ValueChange[1]) {
                $ObjectChange = $true
                $Query = "UPDATE [AXTools_Environments] SET $Property = '$(($ValueChange | Where { $_.SideIndicator -eq '=>' }).$Property)' WHERE ENVIRONMENT = '$($WpfcbxEnvEnvironment.Text)'"
                $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
                $Cmd.ExecuteNonQuery() | Out-Null
            }
        }
        if($ObjectChange) {
            $WpflblWarning.Text = "Saved."
            $CurrentIndex = $WpfcbxEnvEnvironment.SelectedIndex
            Get-TabItemClear
            $WpfcbxEnvEnvironment.SelectedIndex = $CurrentIndex
        }
        else {
            $CurrentIndex = $WpfcbxEnvEnvironment.SelectedIndex
            $WpflblWarning.Text = "Nothing to save."
            Get-TabItemClear
            $WpfcbxEnvEnvironment.SelectedIndex = $CurrentIndex
        }
        $Conn.Close()
        $WpfbtnEnvSave.IsEnabled = $false
    }
})

$WpfbtnEmlSave.Add_Click({
    if($WpfcbxEmlID.SelectedIndex -eq -1) {
        $Conn = Get-ConnectionString
        $Query = "INSERT INTO [dbo].[AXTools_EmailProfile] ([ID],[USERID],[SMTPSERVER],[SMTPPORT],[SMTPSSL],[FROM],[TO],[CC],[BCC])
                    VALUES ('$($WpfcbxEmlID.Text)','$($WpfcbxEmlUserID.SelectedItem["Id"])','$($WpftxtEmlSMTP.Text)','$($WpftxtEmlSMTPPort.Text)','$(if($WpfchkEmlSSL.IsChecked){'1'} else{'0'})','$($WpftxtEmlFrom.Text)','$($WpftxtEmlTo.Text)','$($WpftxtEmlCC.Text)','$($WpftxtEmlBCC.Text)')"
        $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
        $Cmd.ExecuteNonQuery() | Out-Null
        $Conn.Close()
        $WpflblWarning.Text = "Saved."
        Get-TabItemClear
        $WpfcbxEmlID.SelectedIndex = ($WpfcbxEmlID.Items.Count - 1)
    }
    else {
        $Original = New-Object PSObject
        $Original | Add-Member NoteProperty 'ID' $WpfcbxEmlID.SelectedItem.ID
        $Original | Add-Member NoteProperty 'USERID' $WpfcbxEmlID.SelectedItem.USERID
        $Original | Add-Member NoteProperty 'SMTPSERVER' $WpfcbxEmlID.SelectedItem.SMTPSERVER
        $Original | Add-Member NoteProperty 'SMTPPORT' $WpfcbxEmlID.SelectedItem.SMTPPORT
        $Original | Add-Member NoteProperty 'SMTPSSL' $WpfcbxEmlID.SelectedItem.SMTPSSL
        $Original | Add-Member NoteProperty 'FROM' $WpfcbxEmlID.SelectedItem.FROM
        $Original | Add-Member NoteProperty 'TO' $WpfcbxEmlID.SelectedItem.TO
        $Original | Add-Member NoteProperty 'CC' $WpfcbxEmlID.SelectedItem.CC
        $Original | Add-Member NoteProperty 'BCC' $WpfcbxEmlID.SelectedItem.BCC

        $Changed = New-Object PSObject
        $Changed | Add-Member NoteProperty 'ID' $WpfcbxEmlID.Text
        $Changed | Add-Member NoteProperty 'USERID' $WpfcbxEmlUserID.Text
        $Changed | Add-Member NoteProperty 'SMTPSERVER' $WpftxtEmlSMTP.Text
        $Changed | Add-Member NoteProperty 'SMTPPORT' $WpftxtEmlSMTPPort.Text
        $Changed | Add-Member NoteProperty 'SMTPSSL' $(if($WpfchkEmlSSL.IsChecked){'1'} else{'0'})
        $Changed | Add-Member NoteProperty 'FROM' $WpftxtEmlFrom.Text
        $Changed | Add-Member NoteProperty 'TO' $WpftxtEmlTo.Text
        $Changed | Add-Member NoteProperty 'CC' $WpftxtEmlCC.Text
        $Changed | Add-Member NoteProperty 'BCC' $WpftxtEmlBCC.Text

        $Properties = $Original.PsObject.Properties.Name
        $ObjectChange = $false
        $Conn = Get-ConnectionString
        foreach ($Property in $Properties) {
            $ValueChange = Compare-Object $Original $Changed -Property $Property
            if($ValueChange.Count -eq 2 -and $ValueChange[0] -ne $ValueChange[1]) {
                $ObjectChange = $true
                $Query = "UPDATE [AXTools_EmailProfile] SET [$Property] = '$(($ValueChange | Where { $_.SideIndicator -eq '=>' }).$Property)' WHERE ID = '$($WpfcbxEmlID.Text)'"
                $Cmd = New-Object System.Data.SqlClient.SqlCommand($Query,$Conn)
                $Cmd.ExecuteNonQuery() | Out-Null
            }
        }
        if($ObjectChange) {
            $WpflblWarning.Text = "Saved."
            $CurrentIndex = $WpfcbxEmlID.SelectedIndex
            Get-TabItemClear
            $WpfcbxEmlID.SelectedIndex = $CurrentIndex
        }
        else {
            $WpflblWarning.Text = "Nothing to save."
            $CurrentIndex = $WpfcbxEmlID.SelectedIndex
            Get-TabItemClear
            $WpfcbxEmlID.SelectedIndex = $CurrentIndex
        }
        $Conn.Close()
        $WpfbtnEmlSave.IsEnabled = $false
    }
})

$WpfbtnTskSave.Add_Click({
    if(Get-ChkTaskExists "$($WpfcbxTskEnvironment.Text) - $($WpfcbxTskTaskName.Text)") {
        $Answer = New-Popup -Message "Do you want to update $($WpfcbxTskEnvironment.Text) - $($WpfcbxTskTaskName.Text)?" -Title "Alert" -Buttons YesNo -Icon Question
        if($Answer -eq 6) {
            Get-TaskChanges "$($WpfcbxTskEnvironment.Text) - $($WpfcbxTskTaskName.Text)"
            Get-TabItemClear
        }
        else {
            $WpflblWarning.Text = "Canceled."
            Get-TabItemClear
        }
    }
    else {
        if($WpfcbxTskUserID.SelectedIndex -eq -1) {
            $Credential = Get-Credential -Credential "$env:userdomain\$env:username"
        }
        else {
            $Credential = Get-UserCredentials -Account $WpfcbxTskUserID.SelectedItem.Id
        }
        Check-UserPassword $Credential
        Switch($WpfcbxTskTaskName.Text) {
            'AX Monitor' {
                $TaskName = "$($WpfcbxTskEnvironment.Text) - $($WpfcbxTskTaskName.Text)"
                $TaskRunAsUser = $Credential.UserName
                $TaskRunAsPassword = $Credential.GetNetworkCredential().Password             
                $PowershellFilePath = 'Powershell.exe'
                $ScriptFilePath = "$Dir\AX-Monitor\AX-SQLMonitor.ps1"
                $ScriptParameters = $WpfcbxTskEnvironment.Text
                $Action = New-ScheduledTaskAction -Execute $PowershellFilePath -Argument "-NonInteractive -NoLogo -NoProfile -File $ScriptFilePath $ScriptParameters"
                if([System.Environment]::OSVersion.Version.Major -ge 10) {
                    $Trigger = New-ScheduledTaskTrigger -At $(Get-Date) -Once -RepetitionInterval (New-TimeSpan -Minute $($WpftxtTskInterval.Text))
                    $Trigger.ExecutionTimeLimit = 'PT0S'
                }
                else {
                    $Trigger = New-ScheduledTaskTrigger -At $(Get-Date) -Once -RepetitionInterval (New-TimeSpan -Minute $($WpftxtTskInterval.Text)) -RepetitionDuration $([System.TimeSpan]::MaxValue)
                }
                $Settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit $(New-TimeSpan -Minute $($WpftxtTskInterval.Text)) -MultipleInstances Parallel
                $Principals = New-ScheduledTaskPrincipal -RunLevel Highest -LogonType Password -UserId $TaskRunAsUser -Id $TaskRunAsUser
                $Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principals -Description "Created: $(Get-Date -Format G) - $env:USERDOMAIN\$env:USERNAME"
                Get-TasksFolder
                Register-ScheduledTask -TaskPath '\DynamicsAxTools' -TaskName $TaskName -InputObject $Task -User $TaskRunAsUser -Password $TaskRunAsPassword
                Get-TabItemClear
            }
            'AX Report' {
                $TaskName = "$($WpfcbxTskEnvironment.Text) - $($WpfcbxTskTaskName.Text)"
                $TaskRunAsUser = $Credential.UserName
                $TaskRunAsPassword = $Credential.GetNetworkCredential().Password              
                $PowershellFilePath = 'Powershell.exe'
                $ScriptFilePath = "$Dir\AX-Report\AX-ReportManager.ps1"
                $ScriptParameters = "-Environment $($WpfcbxTskEnvironment.Text)"
                $Action = New-ScheduledTaskAction -Execute $PowershellFilePath -Argument "-NonInteractive -NoLogo -NoProfile -File $ScriptFilePath $ScriptParameters"
                $Trigger = New-ScheduledTaskTrigger -Daily -At $(([DateTime]::Parse($WpftxtTskTime.Text)).ToShortTimeString()) -DaysInterval 1
                $Settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit $(New-TimeSpan -Hours 2) -MultipleInstances Queue
                $Principals = New-ScheduledTaskPrincipal -RunLevel Highest -LogonType Password -UserId $TaskRunAsUser -Id $TaskRunAsUser
                $Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principals -Description "Created: $(Get-Date -Format G) - $env:USERDOMAIN\$env:USERNAME"
                Get-TasksFolder
                Register-ScheduledTask -TaskPath '\DynamicsAxTools' -TaskName $TaskName -InputObject $Task -User $TaskRunAsUser -Password $TaskRunAsPassword
                Get-TabItemClear
            }
            'Recycle Perfmon' {
                $TaskName = "$($WpfcbxTskEnvironment.Text) - $($WpfcbxTskTaskName.Text)"
                $TaskRunAsUser = $Credential.UserName
                $TaskRunAsPassword = $Credential.GetNetworkCredential().Password       
                $PowershellFilePath = 'Powershell.exe'
                $ScriptFilePath = "$Dir\AX-Report\AX-ReportManager.ps1"
                $ScriptParameters = "-Environment $($WpfcbxTskEnvironment.Text) -RecycleBlg"
                $Action = New-ScheduledTaskAction -Execute $PowershellFilePath -Argument "-NonInteractive -NoLogo -NoProfile -File $ScriptFilePath $ScriptParameters"
                $Trigger = New-ScheduledTaskTrigger -Daily -At $(([DateTime]::Parse($WpftxtTskTime.Text)).ToShortTimeString()) -DaysInterval 1
                $Settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit $(New-TimeSpan -Hours 2) -MultipleInstances Queue
                $Principals = New-ScheduledTaskPrincipal -RunLevel Highest -LogonType Password -UserId $TaskRunAsUser -Id $TaskRunAsUser
                $Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principals -Description "Created: $(Get-Date -Format G) - $env:USERDOMAIN\$env:USERNAME"
                Get-TasksFolder
                Register-ScheduledTask -TaskPath '\DynamicsAxTools' -TaskName $TaskName -InputObject $Task -User $TaskRunAsUser -Password $TaskRunAsPassword
                Get-TabItemClear
            }
            'Check AOS' {
                $TaskName = "$($WpfcbxTskEnvironment.Text) - $($WpfcbxTskTaskName.Text)"
                $TaskRunAsUser = $Credential.UserName
                $TaskRunAsPassword = $Credential.GetNetworkCredential().Password   
                $PowershellFilePath = 'Powershell.exe'
                $ScriptFilePath = "$Dir\AX-Tools\AX-AOSCheck.ps1"
                $ScriptParameters = "-Environment $($WpfcbxTskEnvironment.Text) -Start"
                $Action = New-ScheduledTaskAction -Execute $PowershellFilePath -Argument "-NonInteractive -NoLogo -NoProfile -File $ScriptFilePath $ScriptParameters"
                if([System.Environment]::OSVersion.Version.Major -ge 10) {
                    $Trigger = New-ScheduledTaskTrigger -At $(Get-Date) -Once -RepetitionInterval (New-TimeSpan -Minute $($WpftxtTskInterval.Text))
                    $Trigger.ExecutionTimeLimit = 'PT0S'
                }
                else {
                    $Trigger = New-ScheduledTaskTrigger -At $(Get-Date) -Once -RepetitionInterval (New-TimeSpan -Minute $($WpftxtTskInterval.Text)) -RepetitionDuration $([System.TimeSpan]::MaxValue)
                }
                $Settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit $(New-TimeSpan -Minute $($WpftxtTskInterval.Text)) -MultipleInstances Parallel
                $Principals = New-ScheduledTaskPrincipal -RunLevel Highest -LogonType Password -UserId $TaskRunAsUser -Id $TaskRunAsUser
                $Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principals -Description "Created: $(Get-Date -Format G) - $env:USERDOMAIN\$env:USERNAME"
                Get-TasksFolder
                Register-ScheduledTask -TaskPath '\DynamicsAxTools' -TaskName $TaskName -InputObject $Task -User $TaskRunAsUser -Password $TaskRunAsPassword
                Get-TabItemClear
            }
            'Check Perfmon' {
                $TaskName = "$($WpfcbxTskEnvironment.Text) - $($WpfcbxTskTaskName.Text)"
                $TaskRunAsUser = $Credential.UserName
                $TaskRunAsPassword = $Credential.GetNetworkCredential().Password   
                $PowershellFilePath = 'Powershell.exe'
                $ScriptFilePath = "$Dir\AX-Tools\AX-PerfmonCheck.ps1"
                $ScriptParameters = "-Environment $($WpfcbxTskEnvironment.Text) -Start"
                $Action = New-ScheduledTaskAction -Execute $PowershellFilePath -Argument "-NonInteractive -NoLogo -NoProfile -File $ScriptFilePath $ScriptParameters"
                if([System.Environment]::OSVersion.Version.Major -ge 10) {
                    $Trigger = New-ScheduledTaskTrigger -At $(Get-Date) -Once -RepetitionInterval (New-TimeSpan -Minute $($WpftxtTskInterval.Text))
                    $Trigger.ExecutionTimeLimit = 'PT0S'
                }
                else {
                    $Trigger = New-ScheduledTaskTrigger -At $(Get-Date) -Once -RepetitionInterval (New-TimeSpan -Minute $($WpftxtTskInterval.Text)) -RepetitionDuration $([System.TimeSpan]::MaxValue)
                }
                $Settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit $(New-TimeSpan -Minute $($WpftxtTskInterval.Text)) -MultipleInstances Parallel
                $Principals = New-ScheduledTaskPrincipal -RunLevel Highest -LogonType Password -UserId $TaskRunAsUser -Id $TaskRunAsUser
                $Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principals -Description "Created: $(Get-Date -Format G) - $env:USERDOMAIN\$env:USERNAME"
                Get-TasksFolder
                Register-ScheduledTask -TaskPath '\DynamicsAxTools' -TaskName $TaskName -InputObject $Task -User $TaskRunAsUser -Password $TaskRunAsPassword
                Get-TabItemClear
            }
            Default {
                $WpflblWarning.Text = "Incorrect Task Name $($WpfcbxTskTaskName.Text)"
            }
        }
    }
})

$WpfbtnSetSave.Add_Click({
    $dsCheck = New-Object System.Data.Dataset
    $Null = $dsCheck.ReadXml("$ModuleFolder\AX-Settings.xml")
    [xml]$ConfigFile = Get-Content "$ModuleFolder\AX-Settings.xml"
    $i=0
    foreach($Row in $Script:SettingsXML.Tables[0]) {
        if($Row.Value -notlike ($dsCheck.Tables.Rows[$i]).Value) {
            $Node = $ConfigFile.DynamicsAxTools.Setting | where {$_.Key -eq $($Row.Key)}
            $Node.Value = $Row.Value
        }
        $i++
    }
    $ConfigFile.Save("$ModuleFolder\AX-Settings.xml")
})

#===========================================================================
# Form Button Delete Click
#===========================================================================

$WpfbtnEnvDelete.Add_Click({
    if($WpfcbxEnvEnvironment.SelectedIndex -ne -1) {
        $Answer = New-Popup -Message "Do you want to continue?" -Title "Alert" -Buttons YesNo -Icon Question
        if($Answer -eq 6) {
            Delete-Environment $WpfcbxEnvEnvironment.SelectedItem["ENVIRONMENT"]
            Get-TabItemClear
            $WpflblWarning.Text = "Completed."
        }
        else {
            $WpflblWarning.Text = "Canceled."
        }
    }
})

$WpfbtnUsrDelete.Add_Click({
    if($WpfcbxUsrID.SelectedIndex -ne -1) {
        $Answer = New-Popup -Message "Do you want to continue?" -Title "Alert" -Buttons YesNo -Icon Question
        if($Answer -eq 6) {
            Delete-User $WpfcbxUsrID.SelectedItem["UserName"]
            $WpfcbxUsrID.SelectedItem = $null
            Get-TabItemClear
            $WpflblWarning.Text = "Completed."
        }
        else {
            $WpflblWarning.Text = "Canceled."
        }
    }
})

$WpfbtnEmlDelete.Add_Click({
    if($WpfcbxEmlID.SelectedIndex -ne -1) {
        $Answer = New-Popup -Message "Do you want to continue?" -Title "Alert" -Buttons YesNo -Icon Question
        if($Answer -eq 6) {
            Delete-Email $WpfcbxEmlID.SelectedItem["Id"]
            $WpfcbxEmlID.SelectedItem = $null
            $WpflblWarning.Text = "Completed."
            Get-TabItemClear
        }
        else {
            $WpflblWarning.Text = "Canceled."
        }
    }
})

$WpfbtnTskDelete.Add_Click({
    $Answer = New-Popup -Message "Do you want to delete $($WpflstCurrJobs.SelectedItem.TaskName)?" -Title "Alert" -Buttons YesNo -Icon Question
    if($Answer -eq 6) {
        Unregister-ScheduledTask -TaskPath $WpflstCurrJobs.SelectedItem.TaskPath -TaskName $WpflstCurrJobs.SelectedItem.TaskName -Confirm:$false
        Get-TabItemClear
    }
    else {
        $WpflblWarning.Text = "Canceled."
    }
    Get-TasksClear
})

#===========================================================================
# Form Environment Func.
#===========================================================================

$WpfbtnEnvReports.Add_Click({
    if(Test-Path "$(($Script:SettingsXML.Tables[0] | Where { $_.Key -eq 'ReportFolder' }).Value)\$($WpfcbxEnvEnvironment.SelectedItem.ENVIRONMENT)") {
        Invoke-Item "$(($Script:SettingsXML.Tables[0] | Where { $_.Key -eq 'ReportFolder' }).Value)\$($WpfcbxEnvEnvironment.SelectedItem.ENVIRONMENT)"
    }
    elseif(Test-Path "$Dir\Reports\$Environment") {
        Invoke-Item "$Dir\Reports\$Environment"
    }
    else {
        $WpflblWarning.Text = 'Folder not found.'
    }
})

$WpfbtnEnvLogs.Add_Click({
    if(Test-Path "$(($Script:SettingsXML.Tables[0] | Where { $_.Key -eq 'LogFolder' }).Value)\$($WpfcbxEnvEnvironment.SelectedItem.ENVIRONMENT)") {
        Invoke-Item "$(($Script:SettingsXML.Tables[0] | Where { $_.Key -eq 'LogFolder' }).Value)\$($WpfcbxEnvEnvironment.SelectedItem.ENVIRONMENT)"
    }
    elseif(Test-Path "$Dir\Logs\$Environment") {
        Invoke-Item "$Dir\Logs\$Environment"
    }
    else {
        $WpflblWarning.Text = 'Folder not found.'
    }
})

$WpftxtEnvDBServer.Add_LostFocus({
    if($WpftxtEnvDBServer.Text -ne '' -and $WpftxtEnvDBName.Text -ne ''){
        $WpfbtnEnvtestSQL.IsEnabled = $true
    }
    else {
        $WpfbtnEnvtestSQL.IsEnabled = $false
    }
})

$WpftxtEnvDBName.Add_TextChanged({
    if($WpftxtEnvDBServer.Text -ne '' -and $WpftxtEnvDBName.Text -ne ''){
        $WpfbtnEnvtestSQL.IsEnabled = $true
    }
    else {
        $WpfbtnEnvtestSQL.IsEnabled = $false
    }
})

$WpfbtnEnvTestSQL.Add_Click({
    $WpflblWarning.Text = "Connecting to $($WpftxtEnvDBServer.Text) - $($WpftxtEnvDBName.Text)..."
    if([string]::IsNullOrEmpty($WpfcbxEnvDBUser.Text)) {
        $SqlServer = Get-SQLObject -DBServer $WpftxtEnvDBServer.Text -DBName $WpftxtEnvDBName.Text
    }
    else {
        $SqlServer = Get-SQLObject -SQLAccount $WpfcbxEnvDBUser.Text -DBServer $WpftxtEnvDBServer.Text -DBName $WpftxtEnvDBName.Text
    }
    if($SqlServer.State -eq 'Open') {
        $SqlServer.Close()
        $WpflblWarning.Text = "$($WpftxtEnvDBServer.Text) - $($WpftxtEnvDBName.Text) connection successful."
    }
    else {
		$WpflblWarning.Text = "Failed to connect to server $($WpftxtEnvDBServer.Text) - $($WpftxtEnvDBName.Text)."
        New-Popup -Message "Failed to connect to $($WpftxtEnvDBServer.Text)" -Title "Error" -Buttons OK -Icon Stop
        $WpftxtEnvDBServer.Clear()
        $WpftxtEnvDBName.Clear()
        $WpftxtEnvDBServer.Focus()
    }
})

#===========================================================================
# Form User Func.
#===========================================================================

$WpfbtnUsrTest.Add_Click({
    if($WpfcbxUsrID.SelectedIndex -eq -1) {
        $WpflblWarning.Text = "User not selected."
    }
    else {  
        [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty
        $Credential = Get-UserCredentials $WpfcbxUsrID.SelectedItem['ID']
        if ($Credential.UserName -ne $null) {
            $BSTRBC = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
            $Root = "LDAP://" + ([ADSI]"").distinguishedName
            $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$Credential.UserName,[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTRBC))
            if ($Domain.Name -eq $null) {
                $WpflblWarning.Text = "Authentication failed or this is not a domain account."
            }
            else {
                $WpflblWarning.Text = "Domain successfully authenticated."
            }
        }
    }
})

#===========================================================================
# Form Email Func.
#===========================================================================

$WpfbtnEmlTest.Add_Click({
    if($WpfcbxEmlID.SelectedIndex -ne -1) {
        Send-Email -Subject "Test Email" -Body "Test Email from DynamicsAxTools / $($WpfcbxEmlID.SelectedItem['Id']) / $($WpfcbxEmlID.SelectedItem['SMTPSERVER']) / $($WpfcbxEmlID.SelectedItem['SMTPPORT'])" -EmailProfile $WpfcbxEmlID.SelectedItem["Id"]
        $WpflblWarning.Text = "Email sent."
    }
})

#===========================================================================
# Form Tasks Func.
#===========================================================================

$WpfbtnTskEnable.Add_Click({
    Enable-ScheduledTask -TaskPath \DynamicsAxTools\ -TaskName $WpflstCurrJobs.SelectedValue.TaskName
    $SelIdx = [ordered]@{'ListView'="$($WpflstCurrJobs.SelectedIndex)";'TaskName'="$($WpfcbxTskTaskName.SelectedIndex)";'EnvName'="$($WpfcbxTskEnvironment.SelectedIndex)"}
    Get-TasksList
    $WpfcbxTskEnvironment.SelectedIndex = $SelIdx[2]
    $WpfcbxTskTaskName.SelectedIndex = $SelIdx[1]
    $WpflstCurrJobs.SelectedIndex = $SelIdx[0]
})

$WpfbtnTskDisable.Add_Click({
    Disable-ScheduledTask -TaskPath \DynamicsAxTools\ -TaskName $WpflstCurrJobs.SelectedValue.TaskName
    $SelIdx = [ordered]@{'ListView'="$($WpflstCurrJobs.SelectedIndex)";'TaskName'="$($WpfcbxTskTaskName.SelectedIndex)";'EnvName'="$($WpfcbxTskEnvironment.SelectedIndex)"}
    Get-TasksList
    $WpfcbxTskEnvironment.SelectedIndex = $SelIdx[2]
    $WpfcbxTskTaskName.SelectedIndex = $SelIdx[1]
    $WpflstCurrJobs.SelectedIndex = $SelIdx[0]
})

$WpftxtTskTime.Add_LostFocus({
    if(![string]::IsNullOrEmpty($WpftxtTskTime.Text)) {
        try {
            $ShortTime = ([DateTime]::Parse($WpftxtTskTime.Text)).ToShortTimeString()
        }
        catch {
            $WpftxtTskTime.Clear()
            $WpflblWarning.Text = "Invalid Time."
        }
    }
})

$WpftxtTskInterval.Add_LostFocus({
    try {
        [ValidateRange(1,59)]$min = $WpftxtTskInterval.Text
    }
    catch {
        $WpftxtTskInterval.Clear()
        $WpflblWarning.Text = "Invalid Interval. Values between 1 and 59."
    }
})

#===========================================================================
# Form Database Func.
#===========================================================================

$WpfbtnDBTestConn.Add_Click({
    $WpflblWarning.Text = "Connecting to $($WpftxtDBServer.Text) - $($WpftxtDBName.Text)..."
    $SqlServer = Get-ConnectionString
    if($SqlServer.State -eq 'Open') {
        $SqlServer.Close()
        $WpflblWarning.Text = "$($WpftxtDBServer.Text) - $($WpftxtDBName.Text) connection successful."
    }
    else {
		$WpflblWarning.Text = "Failed to connect to server $($WpftxtDBServer.Text) - $($WpftxtDBName.Text)."
        New-Popup -Message "Failed to connect to $($WpftxtDBServer.Text)" -Title "Error" -Buttons OK -Icon Stop
    }
})

$WpfbtnDBDrop.Add_Click({
    $Answer = New-Popup -Message "Do you want to drop $($WpftxtDBName.Text)?" -Title "Alert" -Buttons YesNo -Icon Question
    if($Answer -eq 6) {
        $Server = ((Import-ConfigFile).DbServer)
        $Database = ((Import-ConfigFile).DbName)
        $Srv = New-Object ('Microsoft.SqlServer.Management.SMO.Server') $Server
        $Srv.KillAllProcesses($Database)
        $Srv.Databases[$Database].Drop()
        Get-TabItemClear
    }
    else {
        $WpflblWarning.Text = "Canceled."
    }
})

$WpfbtnDBCreate.Add_Click({
    if(!$WpftxtDBServer.Text){ 
        New-Popup -Message "Database server cannot be empty!" -Title "Alert" -Buttons OK -Icon Stop
        $WpftxtDBServer.Clear()
        $WpftxtDBServer.Focus()
        $WpflblWarning.Text = "Canceled."
    }
    elseif(!$WpftxtDBName.Text){
        New-Popup -Message "Database name cannot be empty!" -Title "Alert" -Buttons OK -Icon Stop
        $WpftxtDBName.Clear()
        $WpftxtDBName.Focus()
        $WpflblWarning.Text = "Canceled."
    }
    else {
        try {
            $Srv = New-Object ('Microsoft.SqlServer.Management.SMO.Server') $WpftxtDBServer.Text
            #if($Srv.Databases | Where { $_.Name -eq $WpftxtDBName.Text }) {
            if($Srv.Status -ne 'Online') {
                New-Popup -Message "Failed to connect to server $($WpftxtDBServer.Text)." -Title "Alert" -Buttons OK -Icon Stop
                $SqlServer.Close()
                $WpflblWarning.Text = "Canceled."
            }
            elseif($Srv.Databases[$WpftxtDBName.Text]) {
                New-Popup -Message "Database already exists!" -Title "Alert" -Buttons OK -Icon Stop
                $WpflblWarning.Text = "Canceled."
            }
            else {
                $Db = New-Object ('Microsoft.SqlServer.Management.SMO.Database') ($Srv, $WpftxtDBName.Text)
                $Db.RecoveryModel = 'Simple'
                $Db.Create()
                Invoke-Sqlcmd -InputFile "$ModuleFolder\DynamicsAxNew.sql" -ServerInstance $WpftxtDBServer.Text -Database $WpftxtDBName.Text -ErrorAction 'Stop' -QueryTimeout 1800
                #
                [xml]$ConfigFile = Get-Content "$ModuleFolder\AX-Settings.xml"
                $($ConfigFile.DynamicsAxTools.Setting | where {$_.Key -eq 'DbServer'}).Value = $WpftxtDBServer.Text
                $($ConfigFile.DynamicsAxTools.Setting | where {$_.Key -eq 'DbName'}).Value = $WpftxtDBName.Text
                $ConfigFile.Save("$ModuleFolder\AX-Settings.xml")
                $WpflblWarning.Text = "Done."
                Get-TabItemClear
            }
        }
        catch {
            $WpflblWarning.Text = 'Canceled.' #$error[0]
        }
    }
})

#===========================================================================
# Form Esc. Key
#===========================================================================

$Form.Add_KeyDown{
param
(
    [Parameter(Mandatory)][Object]$Sender,
    [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$e
)
    if($e.Key -eq 'Escape')
    {
        if($WpfbtnEnvSave.IsEnabled -eq $true) {
            $WpflblWarning.Text = "Canceled."
            $WpfcbxEnvEnvironment.IsEditable = $false
            $WpfbtnEnvSave.IsEnabled = $false
            Get-TabItemClear
            if($WpfcbxEnvEnvironment.SelectedIndex -eq -1) {
                $WpfbtnEnvtestSQL.IsEnabled = $false
                $WpfchkEnvRefresh.IsChecked = $false
                $WpfchkEnvGRD.IsChecked= $false
                $WpfcbxEnvEnvironment.SelectedIndex = -1
                $WpfcbxEnvEmail.SelectedIndex = -1 
                $WpfcbxEnvLocalUser.SelectedIndex = -1
                $WpfcbxEnvDBUser.SelectedIndex = -1
                $WpfcbxEnvDBStats.SelectedIndex = -1
                $WpftxtEnvEnvironment.Clear()
                $WpftxtEnvDBServer.Clear()
                $WpftxtEnvDBName.Clear()
                $WpftxtEnvCPU.Clear()
                $WpftxtEnvBlocking.Clear()
                $WpftxtEnvWaiting.Clear()
            }
            else {
                $CurrentIndex = $WpfcbxEnvEnvironment.SelectedIndex
                $WpfcbxEnvEnvironment.SelectedIndex = -1
                $WpfcbxEnvEnvironment.SelectedIndex = $CurrentIndex
            }
        }
        if($WpfbtnEmlSave.IsEnabled -eq $true) {
            $WpflblWarning.Text = "Canceled."
            $WpfcbxEmlID.IsEditable = $false
            $WpfbtnEmlSave.IsEnabled = $false
            Get-TabItemClear
            if($WpfcbxEmlID.SelectedIndex -eq -1) {
                $WpfcbxEmlID.SelectedIndex = -1
                $WpfcbxEmlUserID.SelectedIndex = -1
                $WpfchkEmlSSL.IsChecked = $false
                $WpftxtEmlCC.Clear()
                $WpftxtEmlFrom.Clear()
                $WpftxtEmlSMTP.Clear()
                $WpftxtEmlSMTPPort.Clear()
                $WpftxtEmlTo.Clear()
                $WpftxtEmlSMTP.Clear()
                $WpftxtEmlSMTPPort.Clear()
                $WpftxtEmlBCC.Clear()
            }
            else {
                $CurrentIndex = $WpfcbxEmlID.SelectedIndex
                $WpfcbxEmlID.SelectedIndex = -1
                $WpfcbxEmlID.SelectedIndex = $CurrentIndex
            }
        }
        if($WpfbtnTskSave.IsEnabled -eq $false -and $WpftabControl.SelectedIndex -eq 3) {
            $WpfcbxTskEnvironment.SelectedIndex = -1
            $WpfcbxTskTaskName.SelectedIndex = -1
        }
        if($WpfbtnTskSave.IsEnabled -eq $true) {
            $WpfcbxTskEnvironment.SelectedIndex = -1
            $WpfcbxTskTaskName.SelectedIndex = -1
            $WpfbtnTskSave.IsEnabled = $false
            $WpftxtTskInterval.Clear()
            $WpftxtTskInterval.IsEnabled = $false
            $WpftxtTskTime.Clear()
            $WpftxtTskTime.IsEnabled = $false
            $WpfcbxTskUserID.SelectedIndex = -1
            $WpfcbxTskUserID.IsEnabled = $false
        }
    }    
}

#===========================================================================
# Form Load
#===========================================================================

$Form.Add_Loaded({
    $WpftxtDBServer.Text = ((Import-ConfigFile).DbServer)
    $WpftxtDBName.Text = ((Import-ConfigFile).DbName)
    $Srv = New-Object ('Microsoft.SqlServer.Management.SMO.Server') $WpftxtDBServer.Text
    if($Srv.Status -eq 'Online' -and $Srv.Databases[$WpftxtDBName.Text]) {
        Get-UsersDB
        Get-EmailsDB
        Get-TasksList
        Get-SettingsXML 
        Get-TabItemClear
        $DBStats = [ordered]@{0="No";1="Log Statistics only";2="Log and Update Statistics"}
        $WpfcbxEnvDBStats.ItemsSource = $DBStats
        $SchedTasks = [ordered]@{0="AX Monitor";1="AX Report";2="Check AOS";3="Recycle Perfmon";4="Check Perfmon"}
        $WpfcbxTskTaskName.ItemsSource = $SchedTasks
        $WpfdgXMLSettings.CanUserAddRows = $false
        $WpfdgXMLSettings.CanUserDeleteRows = $false
        $WpflblDBCurrent.Content = 'Connection Successful'
        $WpflblDBCurrent.Foreground = '#00802b'
        $WpfbtnDBCreate.IsEnabled = $false
        $WpfbtnDBDrop.IsEnabled = $true
        $WpfbtnDBTestConn.IsEnabled = $true
        $WpftxtDBServer.IsEnabled = $false
        $WpftxtDBName.IsEnabled = $false
        $WpftxtDBReportPath.Text = ((Import-ConfigFile).ReportFolder)
        $WpftxtDBLogPath.Text = ((Import-ConfigFile).LogFolder)
    }
    else {
        $WpftabControl.Items[0].IsEnabled = $false
        $WpftabControl.Items[1].IsEnabled = $false
        $WpftabControl.Items[2].IsEnabled = $false
        $WpftabControl.Items[3].IsEnabled = $false
        $WpftabControl.Items[4].IsEnabled = $false
        $WpftabControl.Items[5].IsEnabled = $false
        $WpftabControl.Items[6].IsEnabled = $false
        $WpftabControl.SelectedIndex = 7
        $WpfbtnDBCreate.IsEnabled = $true
        $WpfbtnDBDrop.IsEnabled = $false
        $WpfbtnDBTestConn.IsEnabled = $false
        $WpftxtDBServer.Clear()
        $WpftxtDBName.Clear()
    }
})

#===========================================================================
# Form Logo
#===========================================================================

$base64 = "iVBORw0KGgoAAAANSUhEUgAABLAAAAQTCAMAAABDdyC9AAAAJ1BMVEUXJE8XJE8XJE8XJE8YJU91fJWU
m6/////FydTs7fEXJE9PWXkrNl4fCGXHAAAABXRSTlPcQJIZAHDtCVIAACAASURBVHja7d3plptIwgZh
EEuB4f6vd7z22K5yCRC5P/HnW05XWe0ZxXkzlJK6EQAKofNXAICwAICwABAWABAWABAWAMICAMICAMIC
QFgAQFgAQFgACAsACAsACAsAYeE7PXADnkmEFcVXX4AbYCzCivJ35qmGO/DkIywDCyYWYeE/BgMLd02s
wfOJsALz8DzDXTw8nwgr8MDyLMN9mFiEZWDBxCIsKO7Q3QmrrL8vTzHc2t09pwjLwIKJRVhQ3KG7E5bi
Dt0dhGVgwcQiLAMLMLEIS3GH7g7CuufvylMLIfAkJCwHQjgUEpbiDujuhGVgwcQCYRlYMLEIS3EHdHfC
yhRXGhAUVxsIy8CCiUVYijuguxOW4g7dHYRlYMHEIizFHdDdCUtxh+4OwjKwYGIRVuXF3cBCnImluxOW
4g7dnbBcaQBux8QiLAMLJhZhKe6A7k5YrjSg4e7uGUdYBhZMLMJS3AHdnbAUd+juICwDCyYWYSnugO5O
WIo7dHcQloEFE4uwFHdAdycsxR26OwjLwIKJRViKO6C7E5biDt2dsGBgwcQirHKLu4GFlBNLdycsxR26
O2G50gDcjolFWAYWTCzCUtwB3Z2wXGlAw93d85CwDCyYWISluAO6O2Ep7tDdQVgGFkwswjKwABOLsBR3
6O4gLFcakC+enoTlQAiHQsJS3AHdnbAMLJhYICwDCyYWYSnugO5OWK40oEVcbSAsAwsmFmEp7oDuTliK
O3R3EJaBBROLsBR3QHcnLMUdujthwcCCiUVYBRV3Awv5TizdnbAUd+juhOVKA3A7JhZhGVgwsQhLcQd0
d8JypQENd3fPUsIysGBiEZbiDujuhKW4Q3cnLAMLMLEIS3EHdHfCUtyhuxOWgQWYWISluAO6O2Ep7tDd
CcvAAkwswlLcAd2dsBR36O6EZWABJhZh5VHcDSyUNrEGwlLcAd2dsFxpAO5mICwDCzCxCEtxB3R3wnKl
Ae12d8IysAATi7AUd0B3JyzFHbo7YRlYgIlFWAYWYGIRluIO3Z2wXGkA8qUjLAdCwKGQsBR3QHcnLAML
JhZhGViAiUVYijuguxOWKw1okJ6wDCzAxCIsxR3Q3QlLcYfuTlgGFmBiEZbiDujuhKW4Q3cnLAMLMLEI
K01xN7BQ18QaCEtxB3R3wnKlAbibgbAMLMDEIizFHdDdCcuVBrTb3QnLwAJMLMJS3AHdnbAUd+juhGVg
ASYWYSnugO5OWIo7dHfCMrAAE4uwFHdAdycsxR26O2EZWICJRViKO6C7E5biDt2dsAwswMQiLAMLMLHa
FZbiDt2dsFxpALJiICwDCzCxCEtxB3R3wlLcobsTloEFmFiEpbgDujthKe7Q3QnLwAJMLMIysAATq2Vh
Ke7Q3QnLlQYgVzrCciAEHAoJS3EHdHfCMrBgYhGWgQWYWISluAO6e8vCcqUBzdITloEFmFiEpbgDunvj
wlLcobsTloEFmFiEpbgDunurwlLcobsTloEFmFiEdXNxN7BgYg2EpbgDujthudIA3M1AWAYWYGIRluIO
6O4tCktxB350d8IysAATi7AUd0B3b05YijtQb3fvDCzAxCIsxR3Q3QlLcQda7e6dgQWYWISluAO6O2Ep
7kCr3b0zsAATi7AUd0B3JyzFHWi1u3cGFmBiEZaBBZhYhKW4A61292qE5UoD8E8GwjKwABOLsBR3QHev
W1iKO9BAd+8MLMDEIizFHdDdCUtxB1rt7p2BBZhYhGVgASYWYSnuQKvdvQZhudIAHKAjLAdCwKGQsBR3
QHevUFgGFtDMxOoMLMDEIizFHdDdCcuVBuAqPWEZWICJRViKO6C7VyQsxR1oqrt3BhZgYhGW4g7o7oSl
uAOtdvfOwAJMLMIKX9wNLODKxBoIS3EHSuFBWK40AMUwEJaBBZhYhKW4A3fTE5YrDUApdIRlYAEmFmEp
7sDtDISluAOl8CAsAwswsQhLcQfupiMsxR0ohp6wDCzAxCIsxR24mwdhKe5AMQyEZWABJhZhKe7A3fSE
pbgDpdARloEFmFiEZWABzU+s0h6w4g7cyIOwXGkAimEgLAMLMLEIS3EH7qYnLMUdKIWOsAwswMRqXViK
OxCAgbAUd6AUHoRlYAEmVtPCMrCA1idWOcJS3IFA9ITlSgNQCh1hORACDoWtCktxBwIyEJaBBZhYTQrL
wAJMrGKEpbgDQekIy5UGoBh6wjKwABOrOWEp7kBwHoSluAPFMBCWgQWYWE0JS3EHotATluIOlEJHWAYW
YGK1I6zBwAJiTayBsBR3oBQehOVKA1AMA2EZWICJ1YSwFHcgKj1hudIAlEJHWAYWYGLVLyzFHYjOQFiK
O1AKD8IysAATq3JhKe5AAjrCUtyBYugJy8ACTKyKhaW4A4l4EJbi3gR7lB9BaAbCMrDqt9W8zud/al4n
zjKxCheW4l4c2/L29raeds/+9afe5s3fX170hKW41zyupvXtO9PZn5x+/JyZlRcdYRlY9Y6r+e0X69mf
Xf/7UTPLxCpVWAZWSeNqefuNk9bZfv/ZxcwysUoUluJejq7m9e0PlnM/P//506uZlQsPwnKloTb+HFc/
ODWS9vc/v0z+XrNgICwDq+px9TNFnTLeR79hnZ0MTaxyhKW4l8C2vH3MqZsN6z9+iZmVAT1hKe6VjKvp
X6Y5d7Nh+/dvMbOS0xGWgVXFuJrfPuPEzYbPf9EiwJtY2QtLcc9+Xr094fDE2te7fhOCMBCW4l4+yxPN
/HazYdu2+RvLN779L9P222yanvyi1d91Wh6EZWBVcCR8NrG+Omn/Kqrln0l9+SquA+ab/V2bWJkLy8Aq
gGcnuWV59k98/6fmZ/+E7G5iZS4sxb0EprcoLP6mdffMheVKQxHZfY0iLMk9PR1hORCWzxzDV5K7Q2Hm
wlLcS5lYMYQluevumQvLwCqFJYKwJHcTK29hGVjFsEnuJlbzwlLcy2GV3HX3xoXlSkNBTJJ7M/SEZWAV
n91Xyd3EalpYintRzJK77t6ysBR32V1y192LEZaBVZavgh8JTSwTK2NhKe5FFawIV91XH9+XDz1hKe7m
le5eCh1hGVjFEuWdhN9DlmOhiZWlsAYDq5zj4PIWDcfCbCbWQFiKu+OgY6HuXqCwXGkohuktMo6FmTAQ
loElXzGWiVWcsBR3+UrI0t2LEZbizleMlX93JywDi68Yy8QqS1iKO18xlu5ejLAUd75iLN29FGEZWHzF
WCZWMcJS3PmKsXT3UoSluPPVQWO5j6W7pxeWgVUC1++LLss8T9tP5nlerr+zxw1SEyu5sBT3en21fDXV
R79tm+Zr1vIppLp7YmEp7iVw4f2D67x9uof26crU8k7o5rt7Z2DhCac/n+GrrQ793vPDzXcVtj6x0gpL
cS8huJ/01XL85bx9Wr1UWBp9w8JS3LNbU+//X8u5cXWyjG/Li+F9t7ri0rUrLAMrO9Z3xplC6uq7ceYX
MtY331ldDU2szsDCb8//d2e6M18/ePWbuU6trO3diVKKb2hipfzDFffsWN4tpSXKNakTXf+/+6P7f5cj
3M+Ky6NNYbnSkB37u3Q+R8rhJ86Fy7tVpmJFZmhSWAZWdvxhjXXajx8IX7+FfvxcuP396uLqP7lmJlY6
YSnu+fH3wWw+KJH1jo1zeGSt7y7Ky+6R6RsUluKeHVe/EWe5SRjT1fcZes9OZLr2hGVg5cfFj2SYb4ve
l7/zUHZvZWKlEpbinh97+q86vfoxNm42xGZoTFiKe35c+0yGe1+i2689CNk9No+2hGVgZTiw1vS+ykSb
yHZidQYWfjBlIopZdjexMhOW4p4hSy7DZpbdS6BvSFiuNOTHls9BbJbdC6BrR1gOhBky5yOJK68V+oqK
Ng6FKYSluGfIntOoudL/ZffoDI0Iy8DKkPPJPeSX2Fy4QSq7NzGxOgML31jzOoNdeMnSGwpbmFgJhKW4
Z8iWmx8W2T1/uhaE5UpDjiy56eFCxpLdo9M3ICwDK0P2/F6U22R3EysDYSnuOTJnGIxOPyZvKIzPo3Zh
Ke5Zsmb4ktz5Q6HsHp+hcmEZWDkyZXlLc3KzwcRKLCzFPUuWPF+QW2T3/OmrFpbiniN7pm+D2dxsyJ+u
ZmEZWFky5/py3CK7m1gJhTUYWFmyLVkOrLMTa50dCVNMrKFaYSnu2SprzvPgdcKk375GESl41CosVxpy
7ljzmt3AOvFC4eJKQzqGSoVlYOXNtORXtldnQRMrkbAU9zpOhlG3zOwsWAB9lcJS3Ks4Gca9nbk7CxZA
V6OwDKxCmLN6i/Hi5pWJlUJYinsxKSur2+QTYZXAUJ2wFPc6Flbs9+vt3j5YAo/ahGVgFcOS14dOPTkT
+s+rrYnVKe74k8w+wmX2kTIl0NUlLMW9GPbM3q23+ZjRIuirEpaBVQxbbpFbdTexYgtLcS+HKbdBs9xV
3eflEz4X33L9R5vhUY+wFPeCmHNLRre9arlc/z1eqTzAUI2wDKyCWHJ7UW6662VCwqpgYnWKO+56WqeI
aoSVD30lwlLcS2LN7rl51817wgpMV4ewDKyiyO81ufWmqkZYFUyszsDC8edmkltPC2GZWPGEpbgXxZbf
vXLCKoZH+cJypYGwXmS+6ZRKWOEZiheWgUVYhGVilSIsxb0qYX0hLML6lL5wYSnuhEVYDdGVLSwDi7AI
y8QqRViKO2ERVmMMBQtLcScswmqMR7nCMrAIi7BMrGKEZWDVJizXGggr7cTqFHcQFmHdSV+osFxpIKw7
8NacwujKFJYDYZF48zNhZXwo7BR3HH9u+ngZwjrCUKCwDKwKheUD/Agr7cTqDCzc9bROUdV8RHJbE6tT
3HGiGCV4QNNd3+xKWPHoShOWKw2l4mu+COsG+sKEZWDVKaySv0iVsCqYWJ3ijhPJqOSvqiesmDxKEpbi
Xquw1swez5kzKmFFZShIWAZWwdx1iyDKEZWwGptYneKOM81oyuvhvH0hrFzpixGW4l4yc1bPzv3ttodD
WHHpShGWgVU0U1Znwum25k5YNUysAMIaDKyahRX5TLgQVrkTayhCWIp70bp6YojIT88nJ8K3t3XaCStX
HiUIy5WGctnn9e0pUc+E8/PHs847YWXKUICwDKxS2Q7YIfbd0fXQQ1o2wmpkYnWKO46dBf9bNBEn1vR2
9EEdOBkSVnz67IWluFd7FkyQ3Zfjj+r5yZCw4tPlLiwDq96zYPy352ynHtezkyFhVTCxOsUd53wVcWIt
Jx/YSljZMWQtLMW9yIV1qxfSPa4nJiWsFDxyFpaBVSbrSTHMeT6sJzcuCKuCidUp7jj+YlzUFwrPPqpn
HiWsJHT5CktxL5Q9x4l1+kE9+6gZwkpDn62wDKxSmW92Q5LH9EwchFXBxOoUd1zI2+EPhacf0tMXLwkr
EY88haW4F8zyltmh8PyBcHVxNFeGLIVlYBXM6cAd+lB42qDPFUpYFUysTnHHN9a8DoXnBfr8UyQIKxl9
hsJS3IvmdOIO+jw9H7AOPBzCSkaXn7AMrLLZ397yyVjnA9aR9wsRVgUTqzOwcDEaBXtP4X7hsaxfCKuF
idUp7rh6CgtlrAu+OjL3CCshj7yE5UpD+ay5GGu+8ECOfHAzYaVkyEpYBlb5XPLElMnjOCINwqpgYnWK
O36Wo7csjHXJV4duhRFWUvqMhKW4NzuxbjbWfu1BHPqELsJKSpePsAysKtguueLW2w1XXh88bE3CqmBi
dYo7frFeNNZtd963i4/g2HclElZihkyEpbhnPY2OL6Dpoi6Wm4w1XfXVsX9FwkrMIw9hGVi1COvvC+br
fPCEtt7xTujD+Wp9J7aNsFqZWJ2BRVj/8YcylunEe2ReH1nHj4Nf9bQtF4xBWBVMrE5xJ6yPftG8nTsl
vjiyTrw6OP/8gfXsC5WElZw+A2G50lCPsH49p3/7FuUlSsk6Udv/+1ib/dfJ8OgH3RBWcrr0wnIgrElY
06+z4D+7VoiXC7czlxm2dz84fyGsZg6FneJOWL+x/joLHvzlf1f6PbCu/v63+XYy3AmrHIbEwjKw6hLW
9v7Zf+ru+WllndPVB3LYD8czwqpgYnUGFmG98kT/4Pl7PL/vZ29evfS5zIRVwcTqFHfCeqaVs/c513k7
ZKvz78PZvhBW4XQpheVKQwvCuvImw6/O2u+21atvtCasLOgTCsvAakJYF9+zs8zTR4to36bl2ntwXvw3
IawKJlanuBPWc6596Mv3Z/Myz/P2k6//67Je/1VfCKsGHqmEpbg3I6wXjHUbL7/7h7AyYUgkLAOrHWF9
WYr3FWHVMLE6xZ2wjrAnNtYNXzRNWLnQJxGW4p4F+/bsQ2D+Eb+LMtYdH2BDWLnQpRCWgZXBtDr+etsy
bwUb65YP3CKsCibWZWENBlbiZTXNZ6+JL9NeprHu+UhTwspnYg3RhaW4p7XVRXm85qy5ZF8RVkY8YgvL
lYaQJ70nnXx6aeo8ea/f/Jke5oJ9RVg5MUQWloEV0FfrZ6/s7fP6chH65CMV5s8FMcX31W1fI0ZYFUys
TnHP0Ff/fpZuN3WkfyX4+ZkitjWyr+77olbCyok+qrAU97C++oc0phtt8eHJcH4+arao6f2WlwcJK0O6
mMIysEL76iNpTDePm/fKmo8cw/a5vHxFWLVMrE5xz9NX76QxBTiL/aWs+WA4mtbS8hVh1dLdO8U9U1/9
+XQNdRD7Pb/Ph10R50bWncdBwqqku3cGVq6++k0aAU9h6/yhr56tmwgja96/EJaJdYOwFPc4vvpPGmHt
8PNcOJ87j4UeWXfPK8Kqo7t3inu+vvohjfDnr29bZj5bkLawj2n6Qli6+y3CMrCi+eqbNGIU7nWbTzfv
wK8VroRlYt0iLMU9oq++PnFTfqjLZxfuQ//ZjoS6+x3CUtyj+ioxc6qBFUIDhFVBd+8MLL66Yqzwj3cn
LBPrZWEp7k356p/xe0p6HiWsZrt7p7jz1QVjRbg5urqHpbu/KiwDqzVffWysLeG4I6yWJ1ZnYPHVeXFE
ef/zSlgm1mvCUtwb9NUHxtrj/LneS6i7vyQsVxqa9NV7Y0X6gJmFsFpgCCYsA6tNX73bOs8e8zwfiPLL
/NR7O2GZWNeFpbi3Kqy/3on87E7D99f39u2rtj78l1uXZd62Q+IjLN39urAU90aN9fcnJywnPLNv2zbN
P5m+/h/7WfMRlu5+SVgGVqPG+ttX+20nuf3Zv/pEWCbWRWH5qufsjfXtrPWLZVkC+erpN42deAJ/WrHW
WcNqZGINAYSluGdsrPVXFfpTLNs/KtJrvvr+kJdbbiN8MtZe+45qwqq0u3euNJRurPXzp/Y+vfTNq//8
5M9/zqz1DonM93+6DGFlzHC7sAysLI21Hnpmb5ed9eknFU/Ly93pw/f4rLePK8KqZWJ1invBxlqOy+Ha
1+48+2T1D0x48pW99//eyxbo75iwMqa/WViKe37GOtmk9/Mz68A3Qex/f4TzybtTU9jQTliF0N0rLAfC
7Ix14Zl9VlkHv7nmzw+EP7mP/rjZsEwh/4YJq4JDYae4F2msi0Pk1BccHv+mrd9MePrZO8cYV4SVP8ON
wjKwYnC8Mr2QeU58adipJ+KvRnZ6I+2hbjEQVo0TqzOwcuHwpw6vrx2cji+5c3/O95m1XvPIvEX4Cyas
CiZWp7hnwn7UI8urU+TwufD0W/mm5cLblbc1/LgirALo7hKWKw0ZHQjXO7r00ZG1tPM3TFjp6W8SloEV
gS14vbpSsibCIqysJlanuJd0ILzvhbQ5zKGQsHCdxx3CUtxjcMwfd36g3RT9TyQsPGG4QVgGVjYHwnsP
aMdC1kZYhJXRxOoU9+yfS3fm9tPGWgiLsOLRvywsxT2TgbXev3UOGWsjLMKKRveqsAysXAZWCHEcMdZC
WISVz8TqfC5yeqZkFwyOGGsiLMKKN7GGl4SluMcgoTWmhiYWYZXA4xVhudKQScEKd7tgaqdiEVYRDC8I
y8DKo2CFfErMzUwswqpgYnWKe/4DK+h98wPv0tkIi7Di0V8WluIegzmxMJ6/K2gmLMKKR3dVWAZWDPa3
1L54nrF2wiKsPCZWp7jnPrCWtM/kem42EFYpDJeEpbhHYU0/b54eClfCIqyIPK4Iy8CKwpZDQJqayO6E
VcHE6hT3vE+Ea/rnci3ZnbCKoTsvLMU9jxNhnG2ztXAmJKxy6E8Ly8DKwhSxngxzA2dCwqpgYnWKe9Yn
wlii2Bs4ExJWQTzOCUtxz+NEGO+5MNd/JiSskhhOCcvAyuNEGO8kttd/JiSsCiZWp7hnfCJccnk6V3Em
JKyi6E8IS3HPwxIxb5hv1X9kA2EVRXdcWAZWHuewuOHoSU4jLMLKYGJ1Bla+CSvuMWyqPWIRVgUTq1Pc
801YcR2x1x6xCKswHseE5UpDHk+h6FcJlsqfloRVGsMhYRlY0chr00yV38QirAomVqe4p2PLqxrtlX+K
H2EVR39AWIp7JsKKP2nWuqs7YRVH91xYBlbAM9f8J7lFo89fA1j+evQTYRFW9IllYeXyjMngc4m3U49v
ISzCSr+wTKxchBX/CLYTFmEV17C8SpiJsBI8wJWwCCsjHu5hlSOsxQMkrMYZ3HQvR1gpbpbPhEVYWQ8s
7yUkrN/YCIuwsqHzaQ0lCWsjLMJqmt7nYRHW5+yERVhZDyyfOJqtsJI8QsIirKwHls90J6zfWQmLsPLg
4VtzCOvWR0hYhBWQwfcSFiWsxSMkLAPLNz8TFmERVvac/uZn3Z2wCIuwUtGfFpaJRViERViZDaxPhKW7
ExZhEVYShgvC0t0Ji7AIKwWP8YqwTKyUOnCtgbAMrFPC0t0Ji7AIKz79eE1YuntzwvLWHMJKzudOGk2s
PIW1ERZhGVinhKW7NyYsHy9DWMl5jNeFpbunE9ZEWITVIsMLwjKx0gnLRyQTloF1VliD7p5KWAthEVZ7
dMNLwtLdkwlrzf0BEhZh3U8/viYsVxtS+eBtj/8AfZEqYSUeWOOrwjKxUgnLV9UTloF1Wli6eyphxa/u
cyxhbZ8S6kcJK3se4+vCcrUhGFNmEWuJJdDl+h3/UD9KWDkw3CAsEyvVESz21dE92sMhLFwbWEeEpbsn
MkTsiLVFexGAsPABh2R04J/R3UOxZvVEmKOdUAkLH9DfJCwTK40iYl9siKdPwsK1gXVMWLp7IKacXieM
+GAIC+8ZbhOW7p6mGsV9nXCJ9xIAYeEdj/E+YZlYSU5hUV8n3CN+oCBh4drAOigsEytNxFoqfSiEhWsD
66iwdPck3Shidt/XiHcsCAt/cVhEB/85VxuSnMPiZfc5pjoJC3/R3ywsh8IwLJlMrGcDa/1CWISV/EB4
Qli6e5IzYayJNUd9HISFPxluF5aJleRMGOmFwmcD6+aHQVi4NrBOCMvESnImjPNseDaw1i+ERVgZDKwT
wtLdk5wJo7wFeov8IAgLv9OPIYTlakOSM+Eaobs//WTBnbAIKxinJDSaWGl5+jGf4Z8PU+yHQFi4NrBO
CUt3T3IcC97dnxb324+lhIX/8xhDCUt3D8Ga+lD49EC4fiEswgrGEExYJlaSA1ngZ8Tzr56YCYuw8hhY
J4Xli6CTnMiCXh99fiS9f+ERFn7RDQGFpbunmTgB7zZsKXRJWPhFP4YUlqsNISbWc2GtocL7gXkX4P2M
hIVfA2sMKywTK83ECmSs/cB3uwY4jxIWrg2s08LS3dNMrDAvFR7xVYgPjCAs/OAxhhaWqw1pJtbbEl0c
4YI/YeEHQ3BhmViJStL9xjq0r4IsO8LCtYF1QVi6e6KJdXfHOuSrMDcqCAvfuGKf8z+iuwdgjW6sY74K
c8uesPCNPoqwTKwAbG+HuO8+1nZIkYHex0hYuDawLglLd4/+HL79hDYd++OWL4RFWKEYIglLdw/Avh5U
yB1ntH0+9oeFets1YeFKcb8qLBMrAAc3z9v6+rHw4HEw3HsYCQvXBtY1YenuCQ+FL4+so/Mq4DORsHCl
uF8Wlu6e8FD4dWS9sny243/MRliEFYqr5hlNrOIm1gsu2U/8IcsXwiKsvAbWVWHp7vczv51huaKs/dyf
MRMWYQXiMcYVlu6e1ldXlLWd/iNEd8IKxBBZWCZWal99f8HwRH6flgt/wkxYhJXTwLouLJ+WnNxX341y
bGZt83rx9xMWYd1PN0QXlu6eg6++v2T4bGddtpU3PxNWGPoxvrBcbcjCVz+eMvP08dLat3l58Xf7PCzC
un1gjSmEZWLl4qufz5tv2tp+iOvr/5zmeVnv+L0+IpmwshlYrwhLd8/KV+HwrTmElUdxf1FYrjY04Stf
pEpY9zIkEpaJ1Yav7jcWYRlYKYSlu7/O9PbWnrEIS3FPIizd/WX2pQBf3f39F4SluCcRlonVhLFu/74e
wjKw0ghLd2/AWPd/IyJhKe5phKW712+sAN/gSliKeyJhmVi1G2vxVfWElc3AellYJlblxlp88zNh5TOw
XheW7l61sRZfpEpY2RT3W4TlakPFxgrjK8JqlT4DYTkU1musxfcSElZOB8JbhKW7BzTWOq0RxPSPPyWU
rwhLcU8oLBMrnLHWLcb0+iqmj778K5ivCMvASiksEyuUsX58m1fgkfXjD3lvrHC+IiwDK6WwdPdAxvr1
7YNBR9b800t/GyugrwhLcU8qLFcbghjrt29L3UIp67fvCvvTWCF9RVgtcpNqRhMrU2P9+e3OQc6F6/T7
H/G7sYL6irAMrLTC0t3vN9bf30a/z2tQXf1hrLC+IizFPbGwdPe7jbW+/xqce5X10bew/jJWYF8RluKe
WFgm1s3GWj/82q77lPVuXf1urNC+IiwDK7WwfBH0rcZa//mFztNyb2r/wFjBfUVYzdENmQlLd7/TWOtn
X0C/vTiz1vkzIW1reF8RVnP0Y27CcrXhRmNtT/6B6zNrnZ/97j28rwiruYE15icsEyuq06YLX+v83FaR
ICwDK7mwdPfYbPOJobUu05bNIyesMr+OZQAADHtJREFUtniMOQrL1YYUh8f5wNJa5oxkRVjtMWQpLBMr
2fHw0+fTlN8jJiwDKwNh6e6pzoafPp82wiKslNzrmDt/me5OWISFv+izFZaJRViEhYAD62Zh6e6ERVj4
gyFjYenuhEVY+J3HmLOwTCzCIiwEG1h3C0t3JyzCwv/px7yFpbsTFmHhF/f7ZTSxCIuwCKuMgXW/sHR3
wiIs/OAx5i8s3Z2wCAvfGQoQlolFWISFMAMrhLB8WjJhERZu/FzksMLS3QmLsBCguAcSlqsNhEVYCOOW
0cQqnn3+jJ2wCKuWgRVGWLo7CKt1HmM5wnK1AYTVOENBwjKxQFgGVjnC0t1BWC0TTCyBfq/uDsJqmL4w
YZlYICwDqxxh6e4grGYZihOW7g7CapXHWJ6wTCwQloFVjLBMLBCWgVWOsHR3EFaLBJVKwN/tagMIq0H6
QoXlUAjCciAsR1i6OwirOYZihWVigbAMrHKEZWKBsAysYoSlu4Ow2qIfSxaWqw0grJYILpTRxAJhEVYZ
Ayu4sHR3EFY7PMbShaW7g7CaYSheWCYWCMvAKkdYvggahNUG3VCBsHR3EFYb9GMNwnK1AYTVxMAa6xCW
iQXCMrCKEZbuDsKqn8dYi7BcbQBhVc9QjbBMLBCWgVWOsHR3EFbdxDJJnD9GdwdhVU1flbBMLBCWgVWO
sHR3EFbFDJUJS3cHYdXLY6xNWCYWCMvAKkZYujsIq1b6sT5h6e4grDqJaZHRxAJhEVYZAyumsHR3EFaN
PMY6haW7g7AqZKhUWCYWCMvAKkdYPi0ZhFUb3VCtsHR3EFZt9GO9wnK1AYRV2cAaaxaWiQXCMrCKEZbu
DsKqicdYt7BcbQBhVcRQubBMLBCWgVWOsHR3EFYtJNBH9D9RdwdhVULfgLBMLBCWgVWOsHR3wiKsKhia
EJbuTliEVQOPsQ1hmViERVgGVjHCMrEIi7AMrHKEpbsTFmGVTiJ1JPlTXW0gLMIqnL4hYTkUEhZhORCW
Iyzd/R3bp0+KOdCPzp/+6BboRwlLcS9LWCYWYRGWgVWOsEwswiIsA6sYYenuhEVYins5wnK1gbAIq1QS
amM0sQiLsAirjIGVUFi6O2ERluJejrB0d8IiLMW9GGGZWIRFWAZWOcLyRdCERVgFFvehUWHp7oRFWIp7
OcJytYGwCKu4gTW2KywTi7AIy8AqRli6O2ERluJejrBcbSAswiqKoWlhmViERVgGVjnC0t0Ji7AU93KE
pbsTFmEp7sUIy8QiLMIysMoRlu5OWISluBcjLN2dsAhLcS9HWCYWYRGWgVWMsHR3wiIsxb0cYenuhEVY
ins5wjKxCIuwDKxihKW7ExZhKe7lCEt3JyzCUtyLEZaJRViEZWCVIyyflkxYhJV1cR8IS3cnLMJS3IsU
lqsNhEVYGQ+skbBMLMIiLAOrUGE1390Ji7AU94KENRAWYRFWngyEZWIRFmEZWAULq/HuTliEpbgXJaye
sAiLsBT3UoTV9sQiLMIysMoS1kBYhEVYinspwmq6uxMWYSnuhQlrICzCIiwDqxRhtTyxCIuwDKzShNVw
dycswlLcixNWT1iERVj50BOWQyFhEZYDYSXCGgiLsAhLcS9FWM1OLMIiLAOrQGENhEVYhGVglSKsVrs7
YRGW4l6isBq92kBYhJUZWcphNLEIi7AIq4yBlaWw2uzuhEVYinuZwhoIi7AIS3EvRVhNTizCIiwDq1Bh
tfhF0IRFWFkV94GwdHfCIizFvUJhNXi1gbAIK6eBNRKWiUVYhGVgVSms9ro7YRGW4l6usJq72rBvn7FX
9aOf/uT2paofLZOBsEwswMCqVliNfxE0oLgXJazef3EAxb0UYZlYgIFVjrAG/9UB4jMQlu4OFMJjJCwT
CzCwKheW7g7Eph8JS3cHyiB3I4wmFoAyBlbuwtLdgZg8RsLS3YFCGAjLxAIMrEaENejuQCS6gbB0d6AQ
+pGwXG0AChlYI2GZWICB1ZCwdHcgBo+RsFxtAAphICwTCzCwGhOW7g6EphAVFPEodXcgMD1hmViAgdWg
sHR3ICgDYenuQCE8RsIysQADq0lhmViAgVWOsHR3IBQFaaCYR+pqAxCInrAcCgEHwoaFpbsDQRgIy8QC
DKymhWViAW0PrKKEpbsD99OPhOVqA1AGhSlgNLEAA4uwdHcgex4jYenuQCEMhGViAQYWYY2+CBq4k24g
LN0dKIR+JCxXG4BCBtZIWCYWYGARlu4O3MpjJCxXG4BCGAjLxAIMLMLS3YFbKfO5X+KD1t2Bl+kJy8QC
DCzC0t2BexkIS3cHCuExEpaJBRhYhKW7A7fSj4SluwNlUO7zfjSxAAOLsHR3IE8eI2Hp7kAhDIRlYgEG
FmH9e2Lp7sAFuoGwdHegEPqRsFxtAAoZWCNhmViAgUVYujtwK4+RsFxtAAphICwTCzCwCEt3BxT3ioSl
uwOtFPcKhGViAe0MrPKFZWIBzQys8oWluwONFPcqhOVqA3CQgbBMLMDAIizdHVDcKxSW7g60UNwrEZZD
IdDEgbASYenuQAPFvRZhmVhAEwOrEmGZWEALA6sSYenuQP3FvR5hudoAfEY1T/TRxAIMLMLS3QHFnbB0
d6DF4l6TsEwsoPqBVZGwfBE08I/iPhCW7g4o7oTlagNw78AaCcvEAgwswtLdAcW9CWG52gC8YyAsEwsw
sAhLdwcU92aEpbsD1Rb3+oRlYgEVD6zqhKW7A7UW9wqFpbsDtRb3GoVlYgHVDqz6hKW7A5UW9yqFpbsD
dRb3OoVlYgGVDqwahaW7A1UW90qFpbsDNRb3SoVlYgF1Dqw6heXTkoFuICzdHVDcCcvVBuDegTUSlokF
GFiEpbsDinvDwnK1AU0zEJaJBRhYhKW7A4p748LS3aG4E5aJBRhYhGViAQZWu8LS3aG4E5arDUDWDIRl
YgEGFmHp7oDiTli6OxR3wnIoBBwICUt3BxT3xoVlYsHAIiwTCzCwCEt3BxT3hoXlagPaoYGn82hiAQYW
YenugOJOWLo70GJxb0NYJhYMLMIqaGLp7mihuA+EpbsDijthudoA3DuwRsIysQADi7B0d0BxJyxXG9Am
A2GZWICBRVi6O6C4E5buDsWdsEwswMAiLN0dUNwJS3eH4k5YJhZgYBGW7g4o7oSlu0NxJywTCzCwCEt3
BxR3wtLdobgTlokFGFiEFWhi6e6oq7gPhKW7A4o7YbnaANw7sEbCMrEAA4uwdHdAcScsVxvQJgNhmViA
gUVYujuguBOW7g7FnbBMLMDAIiwTCzCwCEt3h+JOWK42APkxEJaJBRhYhKW7A4o7YenuUNwJy6EQcCAk
LN0dUNwJy8SCgUVYJhZgYBGW7g4o7oTlagNqp+kn7WhiAQYWYenugOJOWLo7FHfCMrEAA4uwMp5YujtK
K+4DYenugOJOWK42APcOrJGwTCzAwCIs3R1Q3AnL1Qa0yUBYJhZgYBGW7g4o7oSlu0NxJywTCzCwCEt3
BxR3wtLdobgTlokFGFiEpbsDijth6e5Q3AnLxAIMLMLS3QHFnbB0dyjuhGViAQYWYWU0sXR35FvcDSzC
0t2huBOWqw3AzQPLs5OwTCwYWISluwOKO2G52oBGUdwJy8SCgUVYujuguBOW7g7FHYRlYsHAIiwTCzCw
CEt3h+JOWHC1ARniSgNhmVgwsAhLdwcUd8LS3aG4g7AcCuFASFi6O6C4E5aJBQOLsGBiwcAiLN0dUNwJ
y9UGVI+nJmGZWDCwCEt3BxR3wtLdobiDsEwsGFiEVdHE0t2RsrgbWISlu0NxJyxXG4CbB5bnH2GZWDCw
CEt3BxR3wnK1AY2iuBOWiQUDi7B0d0BxJyzdHYo7CMvEgoFFWLo7oLgTlu4OxR2EZWLBwCIs3R1Q3AlL
d4fiTlgwsWBgEZbuDijuhKW7Q3EnLJhYMLAIq/SJpbsjTnE3sAhLd4fiTliuNgA3DyzPNMIysWBgEZbu
DijuhOVqAxpFcScsEwsGFmHp7oDiTli6OxR3EJaJBQOLsEwswMAiLN0dijsI6zquNiAYrjQQlokFA4uw
dHdAcScs3R2KOwjLoRAOhISluwOKO2GZWDCwQFgmFgwswtLdAcWdsLL8+/IMw514AhKWiQUDi7Cgu0Nx
JyzdHYo7CMvEgoFFWI1PLN0ddxV3A4uwdHco7oSF//+deabhloHluURYJhYMLMLC78YCbsAzibAAEBYA
EBYAEBYAwgIAwgIAwgJAWABAWABAWAAICwAICwAICwBhAQBhAQBhASiL/wGc6ZHma9h6sAAAAABJRU5E
rkJggg=="

# Create a streaming image by streaming the base64 string to a bitmap streamsource
$Bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$Bitmap.BeginInit()
$Bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
$Bitmap.EndInit()
$Bitmap.Freeze()
# This is the pic in the middle
$WpfImage.Source = $Bitmap
$WpflblControl2.Text = $((Get-Date).ToShortTimeString())

#===========================================================================
# Make PowerShell Disappear and handle GUI close
#===========================================================================
<#
# When Exit is clicked, close everything and kill the PowerShell process
$menuitem.add_Click({
	$notifyicon.Visible = $false
	$window.Close()
	Stop-Process $pid
 })

$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
#>

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null
[System.GC]::Collect()


#$WpflstCurrJobs | Get-member Add* -MemberType Method -force
#<TextBlock Text="{Binding ElementName=comboBox1, Path=SelectedItem}"/>