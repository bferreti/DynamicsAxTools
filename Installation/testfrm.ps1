[void][System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework')
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
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
                    <Button x:Name="btnEnvEdit" Content="Edit" HorizontalAlignment="Left" Margin="90,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnEnvSave" Content="Save" HorizontalAlignment="Left" Margin="160,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnEnvDelete" Content="Delete" HorizontalAlignment="Left" Margin="230,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnEnvtestSQL" Content="Test DB Conn" HorizontalAlignment="Left" Margin="300,15,0,0" VerticalAlignment="Top" Width="100" IsEnabled="False"/>
                    <Button x:Name="btnEnvReports" Content="Reports" HorizontalAlignment="Left" Margin="610,15,0,0" VerticalAlignment="Top" Width="60" IsEnabled="False"/>
                    <Button x:Name="btnEnvLogs" Content="Logs" HorizontalAlignment="Left" Margin="680,15,0,0" VerticalAlignment="Top" Width="60" IsEnabled="False"/>
                    <Rectangle Fill="#FFEFEFF1" HorizontalAlignment="Left" Height="192" Margin="13,50,0,0" Stroke="Black" VerticalAlignment="Top" Width="737"/>
                    <ComboBox x:Name="cbxEnvEnvironment" HorizontalAlignment="Left" Margin="95,59,0,0" VerticalAlignment="Top" Width="223" IsEditable="False" DisplayMemberPath="ENVIRONMENT"/>
                    <TextBox x:Name="txtEnvEnvironment" HorizontalAlignment="Left" Height="24" Margin="91,85,0,0" VerticalAlignment="Top" Width="183" />
                    <ComboBox x:Name="cbxEnvEmail" HorizontalAlignment="Left" Margin="320,86,0,0" VerticalAlignment="Top" Width="125" DisplayMemberPath="ID"/>
                    <Label x:Name="lblEnvLocalUser" Content="Local User" HorizontalAlignment="Left" Margin="450,84,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxEnvLocalUser" HorizontalAlignment="Left" Margin="519,86,0,0" VerticalAlignment="Top" Width="125" DisplayMemberPath="ID" />
                    <CheckBox x:Name="chkEnvRefresh" Content="AX Refresh" HorizontalAlignment="Left" Margin="656,90,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxEnvDBUser" HorizontalAlignment="Left" Margin="78,123,0,0" VerticalAlignment="Top" Width="125" DisplayMemberPath="ID"/>
                    <TextBox x:Name="txtEnvDBServer" HorizontalAlignment="Left" Height="24" Margin="286,122,0,0" VerticalAlignment="Top" Width="175"/>
                    <TextBox x:Name="txtEnvDBName" HorizontalAlignment="Left" Height="24" Margin="535,123,0,0" VerticalAlignment="Top" Width="175"/>
                    <ComboBox x:Name="cbxEnvDBStats" HorizontalAlignment="Left" Margin="130,150,0,0" VerticalAlignment="Top" Width="200" DisplayMemberPath="Value" SelectedValuePath="Name"/>
                    <CheckBox x:Name="chkEnvGRD" Content="Enable GRD" HorizontalAlignment="Left" Margin="22,190,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEnvCPU" HorizontalAlignment="Left" Height="24" Margin="109,209,0,0" VerticalAlignment="Top" Width="75"/>
                    <TextBox x:Name="txtEnvBlocking" HorizontalAlignment="Left" Height="24" Margin="305,209,0,0" VerticalAlignment="Top" Width="75"/>
                    <TextBox x:Name="txtEnvWaiting" HorizontalAlignment="Left" Height="24" Margin="495,209,0,0" VerticalAlignment="Top" Width="75"/>
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
                    <Button x:Name="btnUsrDelete" Content="Delete" HorizontalAlignment="Left" Margin="90,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnUsrTest" Content="Test" HorizontalAlignment="Left" Margin="160,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Rectangle Fill="#FFEFEFF1" Height="65" Margin="13,50,14,0" Stroke="Black" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxUsrID" HorizontalAlignment="Left" Margin="77,58,0,0" VerticalAlignment="Top" Width="180" DisplayMemberPath="ID"/>
                    <Label x:Name="lblUsrID" Content="User ID" HorizontalAlignment="Left" Margin="23,56,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblUsrUsername" Content="Username" HorizontalAlignment="Left" Margin="23,83,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtUsrUsername" HorizontalAlignment="Left" Height="22" Margin="91,85,0,0" VerticalAlignment="Top" Width="320"/>
                </Grid>
            </TabItem>
            <TabItem Header="Emails" TabIndex="30">
                <Grid>
                    <Rectangle Fill="#FFEFEFF1" HorizontalAlignment="Left" Height="30" Margin="13,10,0,0" Stroke="Black" VerticalAlignment="Top" Width="737"/>
                    <Button x:Name="btnEmlNew" Content="New" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnEmlEdit" Content="Edit" HorizontalAlignment="Left" Margin="90,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnEmlSave" Content="Save" HorizontalAlignment="Left" Margin="160,15,0,0" VerticalAlignment="Top" Width="65" IsEnabled="False"/>
                    <Button x:Name="btnEmlDelete" Content="Delete" HorizontalAlignment="Left" Margin="230,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnEmlTest" Content="Test Email" HorizontalAlignment="Left" Margin="300,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Rectangle Fill="#FFEFEFF1" Height="208" Margin="13,50,0,0" Stroke="Black" VerticalAlignment="Top" HorizontalAlignment="Left" Width="737"/>
                    <Label x:Name="lblEmlID" Content="ID" HorizontalAlignment="Left" Margin="19,54,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxEmlID" HorizontalAlignment="Left" Margin="46,56,0,0" VerticalAlignment="Top" Width="180" DisplayMemberPath="ID"/>
                    <Label x:Name="lblEmlSMTP" Content="Email Server" HorizontalAlignment="Left" Margin="19,80,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlSMTP" HorizontalAlignment="Left" Height="24" Margin="97,82,0,0" VerticalAlignment="Top" Width="224"/>
                    <Label x:Name="lblEmlSMTPPort" Content="Port" HorizontalAlignment="Left" Margin="326,81,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlSMTPPort" HorizontalAlignment="Left" Height="24" Margin="363,83,0,0" VerticalAlignment="Top" Width="75"/>
                    <CheckBox x:Name="chkEmlSSL" Content="Use SSL" HorizontalAlignment="Left" Margin="450,86,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblEmlUserId" Content="User" HorizontalAlignment="Left" Margin="19,108,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxEmlUserID" HorizontalAlignment="Left" Margin="58,110,0,0" VerticalAlignment="Top" Width="180" IsEnabled="False" DisplayMemberPath="ID"/>
                    <Label x:Name="lblEmlFrom" Content="From" HorizontalAlignment="Left" Margin="19,134,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlFrom" HorizontalAlignment="Left" Height="24" Margin="61,135,0,0" VerticalAlignment="Top" Width="301" />
                    <Label x:Name="lblEmlTo" Content="To" HorizontalAlignment="Left" Margin="19,161,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlTo" HorizontalAlignment="Left" Height="24" Margin="47,162,0,0" VerticalAlignment="Top" Width="498" />
                    <Label x:Name="lblEmlCC" Content="CC" HorizontalAlignment="Left" Margin="19,188,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlCC" HorizontalAlignment="Left" Height="24" Margin="49,189,0,0" VerticalAlignment="Top" Width="496" />
                    <Label x:Name="lblEmlBCC" Content="BCC" HorizontalAlignment="Left" Margin="19,215,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtEmlBCC" HorizontalAlignment="Left" Height="24" Margin="56,216,0,0" VerticalAlignment="Top" Width="489" />
                </Grid>
            </TabItem>
            <TabItem Header="Task Scheduler" TabIndex="40">
                <Grid>
                    <Rectangle Fill="#FFEFEFF1" Height="30" Margin="13,10,14,0" Stroke="Black" VerticalAlignment="Top"/>
                    <Button x:Name="btnTskNew" Content="New" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnTskDelete" Content="Delete" HorizontalAlignment="Left" Margin="90,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnTskSave" Content="Save" HorizontalAlignment="Left" Margin="160,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnTskDisable" Content="Disable" HorizontalAlignment="Left" Margin="230,15,0,0" VerticalAlignment="Top" Width="65"/>
                    <Button x:Name="btnTskEnable" Content="Enable" HorizontalAlignment="Left" Margin="300,15,0,0" VerticalAlignment="Top" Width="65"/>
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
                    <ComboBox x:Name="cbxTskTaskName" HorizontalAlignment="Left" Margin="297,57,0,0" VerticalAlignment="Top" Width="121" DisplayMemberPath="Value" SelectedValuePath="Name"/>
                    <Label x:Name="lblTskTaskName" Content="Task" HorizontalAlignment="Left" Margin="259,55,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxTskEnvironment" HorizontalAlignment="Left" Margin="97,57,0,0" VerticalAlignment="Top" Width="157" IsEditable="False" DisplayMemberPath="ENVIRONMENT"/>
                    <Label x:Name="lblTskName" Content="Environment" HorizontalAlignment="Left" Margin="15,55,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="cbxTskInterval" HorizontalAlignment="Left" Margin="477,57,0,0" VerticalAlignment="Top" Width="93"/>
                    <Label x:Name="lblTskInterval" Content="Interval" HorizontalAlignment="Left" Margin="423,55,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblTskTimeSpan" Content="At" HorizontalAlignment="Left" Margin="575,55,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtTskTimeSpan" HorizontalAlignment="Left" Height="22" Margin="597,57,0,0" Text="HHAM/PM" VerticalAlignment="Top" Width="69"/>
                    <CheckBox x:Name="ckcTskRunAs" Content="Run As" HorizontalAlignment="Left" Margin="676,60,0,0" VerticalAlignment="Top"/>
                    <ListView x:Name="lstCurrJobs" HorizontalAlignment="Left" Height="164" Margin="17,98,0,0" VerticalAlignment="Top" Width="732">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Environment" DisplayMemberBinding ="{Binding Environment}"/>
                                <GridViewColumn Header="Name" DisplayMemberBinding ="{Binding Name}"/>
                                <GridViewColumn Header="IntervalMin" DisplayMemberBinding ="{Binding Interval}"/>
                                <GridViewColumn Header="DaysInterval" DisplayMemberBinding ="{Binding DaysInterval}"/>
                                <GridViewColumn Header="At" DisplayMemberBinding ="{Binding At}"/>
                                <GridViewColumn Header="User" DisplayMemberBinding ="{Binding User}"/>
                                <GridViewColumn Header="Status" DisplayMemberBinding ="{Binding State}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Grid>
            </TabItem>
            <TabItem Header="Tools" TabIndex="50" IsEnabled="False">
                <Grid Background="#FFE5E5E5"/>
            </TabItem>
            <TabItem Header="Settings" TabIndex="60">
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

$Form.Add_Loaded({
    $TasksNew = Get-ScheduledTask | Select @{'n'='Environment'; 'e'={$_.TaskName.Split('-')[0].Trim()}}, @{'n'='Name'; 'e'={$_.TaskName.Split('-')[1].Trim()}} ,  @{'n'='Interval'; 'e'={$(Get-Interval $_.Triggers.Repetition.Interval)}}, @{'n'='User'; 'e'={$_.Principal.UserId}}, @{'n'='DaysInterval'; 'e'={$_.Triggers.DaysInterval}}, @{'n'='At'; 'e'={([datetime]$_.Triggers.StartBoundary).ToShortTimeString()}}, State, TaskName, TaskPath
    $wpflstCurrJobs.ItemsSource = @($TasksNew) #| Select Environment, Name, Interval, DaysInterval, At, User, State
    #$wpflistView.IsReadOnly = $true
})

#$WpflstCurrJobs | Get-member Add* -MemberType Method -force
#<TextBlock Text="{Binding ElementName=comboBox1, Path=SelectedItem}"/>

#===========================================================================
# Shows the form
#===========================================================================
#write-host "To show the form, run the following" -ForegroundColor Cyan

#$Form.Activate()
#$Form.Topmost = $true  
#$Form.Topmost = $false 
#$Form.Focus()         
$Form.ShowDialog() | out-null
[System.GC]::Collect()