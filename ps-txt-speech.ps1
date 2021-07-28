Add-Type -AssemblyName PresentationFramework 
 
[xml]$xaml =  
@" 
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
        Title="Text To Voice" Height="202" Width="420" WindowStyle="ToolWindow" Background="White"> 
    <Grid> 
        <Label Content="Convert Text to Voice" Height="28" HorizontalAlignment="Left" Margin="122,12,0,0" Name="label1" VerticalAlignment="Top" Foreground="Blue" FontWeight="Bold" /> 
        <Label Content="Your Text" Height="28" HorizontalAlignment="Left" Margin="12,46,0,0" Name="label2" VerticalAlignment="Top" FontWeight="Bold" Foreground="Blue" /> 
        <TextBox Height="23" HorizontalAlignment="Left" Margin="90,49,0,0" Name="textBox1" VerticalAlignment="Top" Width="283" Text="Just For Learning!!!" /> 
        <Button Content="Listen" Height="23" HorizontalAlignment="Left" Margin="90,91,0,0" Name="Listen" VerticalAlignment="Top" Width="75" /> 
        <Button Content="Close" Height="23" HorizontalAlignment="Left" Margin="195,91,0,0" Name="Close" VerticalAlignment="Top" Width="75" /> 
        <Button Content="Save" Height="23" HorizontalAlignment="Left" Margin="298,91,0,0" Name="Save" VerticalAlignment="Top" Width="75" /> 
         
    </Grid> 
</Window> 
 
"@ 
 
 
$speak = 
{ 
Add-Type -AssemblyName System.Speech 
$Synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer 
$synthesizer.Rate = "7" 
$synthesizer.Speak($Textbox.Text) 
} 
 
 
$wavfile =  
{ 
$path = "C:\Temp\texttowav.wav" 
Add-Type -AssemblyName System.Speech 
$synthesizer=New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer 
$synthesizer.Rate = "-10" 
$synthesizer.SetOutputToWaveFile($wav) 
$synthesizer.Speak($Textbox.Text) 
$synthesizer.SetOutputToDefaultAudioDevice() 
Invoke-Item $wav 
} 
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
$Window=[Windows.Markup.XamlReader]::Load( $reader ) 
 
 
$Textbox = $Window.Findname("textBox1")  
$Listen = $Window.FindName("Listen") 
$Save = $Window.FindName("Save") 
$Close = $Window.FindName("Close") 
$listen.Add_MouseEnter({$Window.Cursor = [System.Windows.Input.Cursors]::Hand}) 
$listen.Add_MouseLeave({$Window.Cursor = [System.Windows.Input.Cursors]::Arrow}) 
$Listen.add_Click($speak) 
$Save.add_Click($wavfile) 
 
$Close.Add_Click({$window.Close()}) 
 
$Window.ShowDialog() | Out-Null 