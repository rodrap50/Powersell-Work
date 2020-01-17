Add-Type -AssemblyName PresentationFramework

[XML]$XAML = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:MyWorkstation"
        Title="My Workstation" Height="400" Width="550" ResizeMode="CanMinimize">
        <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="271"/>
            <ColumnDefinition Width="271"/>
        </Grid.ColumnDefinitions>
        <Label Content="My Workstation Information" HorizontalAlignment="Left" Margin="116,21,0,0" VerticalAlignment="Top" FontSize="24" Grid.ColumnSpan="2" RenderTransformOrigin="0.66,0.561" Height="42" Width="309"/>
        <Label Content="Workstation Id:" HorizontalAlignment="Left" Margin="67,68,0,0" VerticalAlignment="Top"/>
        <Label Content="Workstation OS Name:" HorizontalAlignment="Left" Margin="67,99,0,0" VerticalAlignment="Top"/>
        <Label Content="Workstation OS Version:" HorizontalAlignment="Left" Margin="67,130,0,0" VerticalAlignment="Top"/>
        <Label Content="Last Restart Time:" HorizontalAlignment="Left" Margin="67,161,0,0" VerticalAlignment="Top"/>
        <Label Content="Current Uptime:" HorizontalAlignment="Left" Margin="67,192,0,0" VerticalAlignment="Top"/>
        <Label Content="Current Logged On User:" HorizontalAlignment="Left" Margin="67,223,0,0" VerticalAlignment="Top"/>
        <Label Content="Current Domain:" HorizontalAlignment="Left" Margin="67,254,0,0" VerticalAlignment="Top"/>
        <Label Content="Current IP Address:" HorizontalAlignment="Left" Margin="67,285,0,0" VerticalAlignment="Top"/>
        <Label Name="Label1" Content="" Grid.Column="1" HorizontalAlignment="Left" Margin="10,68,0,0" VerticalAlignment="Top" Width="246" FontSize="18"/>
        <Label Name="Label2" Content="" Grid.Column="1" HorizontalAlignment="Left" Margin="10,99,0,0" VerticalAlignment="Top" Width="246" FontSize="18"/>
        <Label Name="Label3" Content="" Grid.Column="1" HorizontalAlignment="Left" Margin="10,130,0,0" VerticalAlignment="Top" Width="246" FontSize="18"/>
        <Label Name="Label4" Content="" Grid.Column="1" HorizontalAlignment="Left" Margin="10,161,0,0" VerticalAlignment="Top" Width="246" FontSize="18"/>
        <Label Name="Label5"  Content="" Grid.Column="1" HorizontalAlignment="Left" Margin="10,192,0,0" VerticalAlignment="Top" Width="246" FontSize="18"/>
        <Label Name="Label6"  Content="" Grid.Column="1" HorizontalAlignment="Left" Margin="10,223,0,0" VerticalAlignment="Top" Width="246" FontSize="18"/>
        <Label Name="Label7"  Content="" Grid.Column="1" HorizontalAlignment="Left" Margin="10,254,0,0" VerticalAlignment="Top" Width="246" FontSize="18"/>
        <Label Name="Label8"  Content="" Grid.Column="1" HorizontalAlignment="Left" Margin="10,285,0,0" VerticalAlignment="Top" Width="246" FontSize="18"/>
        <Border BorderBrush="Black" BorderThickness="1" Grid.Column="1" HorizontalAlignment="Left" Height="243" Margin="0,68,0,0" VerticalAlignment="Top" Width="1"/>
        <Border BorderBrush="Black" BorderThickness="1" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="31" Margin="16,68,0,0" VerticalAlignment="Top" Width="511"/>
        <Border BorderBrush="Black" BorderThickness="1,0,1,1" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="31" Margin="16,94,0,0" VerticalAlignment="Top" Width="511"/>
        <Border BorderBrush="Black" BorderThickness="1,0,1,1" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="31" Margin="16,125,0,0" VerticalAlignment="Top" Width="511"/>
        <Border BorderBrush="Black" BorderThickness="1,0,1,1" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="31" Margin="16,156,0,0" VerticalAlignment="Top" Width="511"/>
        <Border BorderBrush="Black" BorderThickness="1,0,1,1" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="31" Margin="16,187,0,0" VerticalAlignment="Top" Width="511"/>
        <Border BorderBrush="Black" BorderThickness="1,0,1,1" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="31" Margin="16,218,0,0" VerticalAlignment="Top" Width="511"/>
        <Border BorderBrush="Black" BorderThickness="1,0,1,1" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="31" Margin="16,249,0,0" VerticalAlignment="Top" Width="511"/>
        <Border BorderBrush="Black" BorderThickness="1,0,1,1" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="31" Margin="16,280,0,0" VerticalAlignment="Top" Width="511"/>

    </Grid>
</Window>
"@

#Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
} catch {
    Write-Warning $_.Exception
    throw
}

# Create variables based on form control names.
# Variable will be named as 'var_<control name>'

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)"
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}
Get-Variable var_*

$now = get-date
$startTime = [Management.ManagementDateTimeConverter]::ToDateTime((get-WMIObject Win32_OperatingSystem).lastbootuptime)
$uptime = $now - $startTime

$IPAddress = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.Ipaddress.length -gt 1} 


$var_Label1.Content = $env:COMPUTERNAME
$var_Label2.Content = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
$var_Label3.Content = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseID
$var_Label4.Content = $startTime
$var_Label5.Content = "$($uptime.days) Days, $($uptime.Hours) Hours, $($uptime.Minutes) Minutes, $($uptime.Seconds) Seconds"
$var_Label6.Content = $env:USERNAME 
$var_Label7.Content = $env:UserDnsDomain
$var_Label8.Content = $IpAddress.Ipaddress[0] 
<#
$var_txtResults.Text = ""
$var_txtResults.Text = $var_txtResults.Text  + "Workstation ID: $env:COMPUTERNAME`n"
$var_txtResults.Text = $var_txtResults.Text  + "Workstation OS Name: $((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName)`n"
$var_txtResults.Text = $var_txtResults.Text  + "Workstation OS Version: $((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseID)`n"
$var_txtResults.Text = $var_txtResults.Text  + "Last Restart Time: $startTime  `n"
$var_txtResults.Text = $var_txtResults.Text  + "Current Uptime = $($uptime.days) Days, $($uptime.Hours) Hours, $($uptime.Minutes) Minutes, $($uptime.Seconds) Seconds`n"
$var_txtResults.Text = $var_txtResults.Text  + "Current Logged On User: $env:USERNAME  `n"
$var_txtResults.Text = $var_txtResults.Text  + "Current Domain: $env:UserDnsDomain `n"
$var_txtResults.Text = $var_txtResults.Text  + "Current IP Address: $($IpAddress.Ipaddress[0])  `n"
#>


$Null = $window.ShowDialog()

