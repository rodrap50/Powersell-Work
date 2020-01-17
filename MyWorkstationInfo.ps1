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
        <TextBox Name="txtResults" HorizontalAlignment="Left" Height="273" Margin="10,86,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="522" IsEnabled="False" Grid.ColumnSpan="2" FontSize="18"/>
        <Label Content="My Workstation Information" HorizontalAlignment="Left" Margin="116,21,0,0" VerticalAlignment="Top" FontSize="24" Grid.ColumnSpan="2" RenderTransformOrigin="0.66,0.561" Height="42" Width="309"/>
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

$var_txtResults.Text = ""
$var_txtResults.Text = $var_txtResults.Text  + "Workstation ID: $env:COMPUTERNAME`n"
$var_txtResults.Text = $var_txtResults.Text  + "Workstation OS Name: $((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName)`n"
$var_txtResults.Text = $var_txtResults.Text  + "Workstation OS Version: $((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseID)`n"
$var_txtResults.Text = $var_txtResults.Text  + "Last Restart Time: $startTime  `n"
$var_txtResults.Text = $var_txtResults.Text  + "Current Uptime = $($uptime.days) Days, $($uptime.Hours) Hours, $($uptime.Minutes) Minutes, $($uptime.Seconds) Seconds`n"
$var_txtResults.Text = $var_txtResults.Text  + "Current Logged On User: $env:USERNAME  `n"
$var_txtResults.Text = $var_txtResults.Text  + "Current Domain: $env:UserDnsDomain `n"
$var_txtResults.Text = $var_txtResults.Text  + "Current IP Address: $($IpAddress.Ipaddress[0])  `n"



$Null = $window.ShowDialog()

