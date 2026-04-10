# Hourly Quran
$scriptDir    = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path -Parent }
$quranPath    = Join-Path $scriptDir "docs/quran.json"
$audioBase    = Join-Path $scriptDir "audio"
$recitersFile = Join-Path $scriptDir "docs/reciters.json"
$khatmaFile   = Join-Path $scriptDir "configs/khitmaProgress.json"
$settingsFile = Join-Path $scriptDir "configs/settings.json"
$logFile      = Join-Path $scriptDir "logs/debug.log"

function Get-Reciters {
    if (-not (Test-Path $recitersFile)) { Log "reciters.json not found at $recitersFile"; return @() }
    try { return @(Get-Content $recitersFile -Encoding UTF8 | ConvertFrom-Json) }
    catch { Log "reciters.json parse error: $_"; return @() }
}
function Save-Reciters($list) { $list | ConvertTo-Json -Depth 3 | Out-File -FilePath $recitersFile -Encoding UTF8 }
function Set-ReciterStatus([string]$url, [string]$status) {
    $list = Get-Reciters
    foreach ($r in $list) { if ($r.url -eq $url) { $r.status = $status; break } }
    Save-Reciters $list; Log "Reciter status updated: $url -> $status"
}
function Log($msg) {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$ts  $msg"
    $retries = 0
    while ($retries -lt 5) {
        try {
            $stream = [System.IO.File]::Open($logFile, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
            $writer.WriteLine($line)
            $writer.Close()
            $stream.Close()
            break
        } catch {
            $retries++
            Start-Sleep -Milliseconds 50
        }
    }
}
function Get-ReciterDisplayName($rObj) {
    if ($rObj -eq $null) { return "" }
    if ($rObj.PSObject.Properties['nameAr'] -and -not [string]::IsNullOrWhiteSpace($rObj.nameAr)) { return $rObj.nameAr }
    return $rObj.name
}
function Get-Settings {
    $defaults = [PSCustomObject]@{ mode="equalBoth"; percentage=33; timerMultiplier=1.0; verseCount=5; fullSurah=$false; reciter="none" }
    if (Test-Path $settingsFile) {
        try {
            $raw = Get-Content $settingsFile -Encoding UTF8 | ConvertFrom-Json
            if ($raw.PSObject.Properties['mode'])            { $defaults.mode            = $raw.mode }
            if ($raw.PSObject.Properties['percentage'])      { $defaults.percentage      = [int]$raw.percentage }
            if ($raw.PSObject.Properties['timerMultiplier']) { $defaults.timerMultiplier = [double]$raw.timerMultiplier }
            if ($raw.PSObject.Properties['verseCount'])      { $defaults.verseCount      = [int]$raw.verseCount }
            if ($raw.PSObject.Properties['fullSurah'])       { $defaults.fullSurah       = [bool]$raw.fullSurah }
            if ($raw.PSObject.Properties['reciter'])         { $defaults.reciter         = [string]$raw.reciter }
        } catch { Log "Settings load error: $_" }
    } else { $defaults | ConvertTo-Json | Out-File -FilePath $settingsFile -Encoding UTF8; Log "Created default settings.json" }
    return $defaults
}
function Save-Settings($s) {
    $s | ConvertTo-Json | Out-File -FilePath $settingsFile -Encoding UTF8
    Log "Settings saved: mode=$($s.mode) pct=$($s.percentage) mult=$($s.timerMultiplier) verses=$($s.verseCount) fullSurah=$($s.fullSurah) reciter=$($s.reciter)"
}
Log "=== SCRIPT STARTED === args: $($args -join ' ')"

# =============================================================================
# MODE: PLAY
# =============================================================================
if ($args -contains '-play') {
    Log "MODE: -play"
    $playPidFile = Join-Path $env:TEMP "quran_play.pid"
    try {
        $audioFiles = $args[$args.IndexOf('-play') + 1] -split ';'
        Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
public class MCI {
    [DllImport("winmm.dll", CharSet=CharSet.Auto)]
    public static extern int mciSendString(string cmd,string ret,int retLen,IntPtr hwnd);
    public static int GetDurationMs(string alias) {
        string r=new string(' ',128);
        mciSendString("status "+alias+" length",r,128,IntPtr.Zero);
        r=r.Trim('\0').Trim(); int ms; return int.TryParse(r,out ms)?ms:0;
    }
}
"@
        # Write PID so popup can stop us
        [string]$PID | Out-File -FilePath $playPidFile -Encoding ASCII -Force
        Log "Play PID=$PID written to $playPidFile"
        $index = 0
        foreach ($file in $audioFiles) {
            $file = $file.Trim()
            if (-not (Test-Path $playPidFile)) { Log "Stop signal detected"; break }
            Log "Playing: $file  exists=$(Test-Path $file)"
            if (Test-Path $file) {
                $alias = "track$index"
                [MCI]::mciSendString("open `"$file`" type mpegvideo alias $alias",$null,0,[IntPtr]::Zero) | Out-Null
                [MCI]::mciSendString("play $alias",$null,0,[IntPtr]::Zero) | Out-Null
                $timeout=0; $durationMs=0
                while ($durationMs -eq 0 -and $timeout -lt 30) { Start-Sleep -Milliseconds 100; $durationMs=[MCI]::GetDurationMs($alias); $timeout++ }
                $elapsed=0; $waitMs=if($durationMs -gt 0){$durationMs+300}else{8000}
                while ($elapsed -lt $waitMs) {
                    if (-not (Test-Path $playPidFile)) { Log "Stop signal during track"; break }
                    Start-Sleep -Milliseconds 200; $elapsed+=200
                }
                [MCI]::mciSendString("stop $alias",$null,0,[IntPtr]::Zero) | Out-Null
                [MCI]::mciSendString("close $alias",$null,0,[IntPtr]::Zero) | Out-Null
                $index++
                if (-not (Test-Path $playPidFile)) { break }
            }
        }
    } catch { Log "PLAY ERROR: $_" }
    finally { if (Test-Path $playPidFile) { Remove-Item $playPidFile -Force -EA SilentlyContinue } }
    Log "MODE: -play done"; exit
}

# =============================================================================
# MODE: STOP PLAY
# =============================================================================
if ($args -contains '-stopplay') {
    Log "MODE: -stopplay"
    $playPidFile = Join-Path $env:TEMP "quran_play.pid"
    try {
        if (Test-Path $playPidFile) {
            $pidStr = (Get-Content $playPidFile -Encoding ASCII -Raw).Trim()
            Log "Stopping play PID=$pidStr"
            Remove-Item $playPidFile -Force -EA SilentlyContinue
            if ($pidStr -match '^\d+$') {
                $proc = Get-Process -Id ([int]$pidStr) -EA SilentlyContinue
                if ($proc) { Stop-Process -Id ([int]$pidStr) -Force -EA SilentlyContinue; Log "Killed PID $pidStr" }
            }
        } else { Log "No play PID file found" }
    } catch { Log "STOPPLAY ERROR: $_" }
    exit
}
# =============================================================================
# MODE: ADVANCE-KHATMA
# =============================================================================
if ($args -contains '-advancekhatma') {
    Log "MODE: -advancekhatma"
    try {
        $quranData = Get-Content $quranPath -Encoding UTF8 | ConvertFrom-Json
        $settings  = Get-Settings; $vc = [int]$settings.verseCount
        if (Test-Path $khatmaFile) {
            $raw = Get-Content $khatmaFile -Encoding UTF8 | ConvertFrom-Json
            $progress = [PSCustomObject]@{
                surahIndex=[int]$raw.surahIndex; verseIndex=[int]$raw.verseIndex
                nextMode=if($raw.PSObject.Properties['nextMode']){[string]$raw.nextMode}else{'khatma'}
            }
        } else { $progress=[PSCustomObject]@{surahIndex=0;verseIndex=0;nextMode="khatma"} }
        $si=[int]$progress.surahIndex; $vi=[int]$progress.verseIndex
        Log "Before advance: surah=$si verse=$vi (advancing $vc)"
        $vi+=$vc
        while ($si -lt $quranData.Count -and $vi -ge $quranData[$si].verses.Count) { $vi-=$quranData[$si].verses.Count; $si++ }
        if ($si -ge $quranData.Count) { $si=0; $vi=0 }
        $progress.surahIndex=$si; $progress.verseIndex=$vi
        Log "After advance: surah=$si verse=$vi"
        $progress | ConvertTo-Json | Out-File -FilePath $khatmaFile -Encoding UTF8
    } catch { Log "ADVANCEKHATMA ERROR: $_" }
    exit
}

# =============================================================================
# MODE: SETTINGS WINDOW
# =============================================================================
if ($args -contains '-settings') {
    Log "MODE: -settings"
    try {
        $settings = Get-Settings
        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName PresentationCore
        Add-Type -AssemblyName WindowsBase

        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="إعدادات القرآن الكريم" Width="450" SizeToContent="Height"
        WindowStartupLocation="CenterScreen" FlowDirection="RightToLeft"
        Background="Transparent" WindowStyle="None" AllowsTransparency="True"
        ResizeMode="CanMinimize" Topmost="True">
    <Border CornerRadius="14" BorderThickness="1" BorderBrush="#3a6bc4" Name="RootBorder">
        <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                <GradientStop Color="#F70d1b3e" Offset="0"/><GradientStop Color="#F7071428" Offset="1"/>
            </LinearGradientBrush>
        </Border.Background>
        <StackPanel Margin="28,24,28,28">
            <!-- Header -->
            <Grid Margin="0,0,0,22">
                <TextBlock Text="الإعدادات" FontSize="24" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial"
                           Foreground="#8ab4f8" FontWeight="Bold" HorizontalAlignment="Right"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Left">
                    <Button Name="BtnMinimize" Content="─" Background="Transparent" BorderThickness="0"
                            Foreground="#6a94d8" FontSize="17" Cursor="Hand" Padding="6,0" Margin="0,0,4,0"/>
                    <Button Name="BtnX" Content="✕" Background="Transparent" BorderThickness="0"
                            Foreground="#6a94d8" FontSize="17" Cursor="Hand" Padding="6,0"/>
                </StackPanel>
            </Grid>
            <!-- Mode -->
            <TextBlock Text="وضع العرض" FontSize="17" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" Foreground="#8ab4f8" Margin="0,0,0,9"/>
            <ComboBox Name="CmbMode" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" FontSize="14"
                      Background="White" Foreground="Black" BorderBrush="#3a6bc4" Padding="8,5">
                <ComboBoxItem Tag="onlyKhatma"  Content="ختمة فقط"                             Foreground="Black"/>
                <ComboBoxItem Tag="onlyRandom"  Content="عشوائي فقط"                           Foreground="Black"/>
                <ComboBoxItem Tag="equalBoth"   Content="تناوب متساوٍ (ختمة ثم عشوائي)"       Foreground="Black"/>
                <ComboBoxItem Tag="percentage"  Content="ختمة + احتمال آيات عشوائية"           Foreground="Black"/>
            </ComboBox>
            <!-- Percentage -->
            <StackPanel Name="PctPanel" Margin="0,15,0,0">
                <TextBlock Text="نسبة احتمال العشوائي %" FontSize="16" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" Foreground="#8ab4f8" Margin="0,0,0,9"/>
                <Grid>
                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="58"/></Grid.ColumnDefinitions>
                    <Slider Name="SldPct" Grid.Column="0" Minimum="0" Maximum="100" TickFrequency="1" IsSnapToTickEnabled="True" Foreground="#3a6bc4" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBlock Name="LblPct" Grid.Column="1" FontSize="17" Foreground="White" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" VerticalAlignment="Center" TextAlignment="Center"/>
                </Grid>
            </StackPanel>
            <!-- Reciter -->
            <TextBlock Text="القارئ" FontSize="17" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" Foreground="#8ab4f8" Margin="0,20,0,9"/>
            <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <ComboBox Grid.Column="0" Name="CmbReciter" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" FontSize="14"
                          Background="White" Foreground="Black" BorderBrush="#3a6bc4" Padding="8,5" Margin="0,0,9,0"/>
                <Button Grid.Column="1" Name="BtnDownload" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial"
                        FontSize="13" Foreground="White" BorderThickness="0" Padding="10,7" Cursor="Hand" VerticalAlignment="Center">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Name="Bd" Background="{TemplateBinding Background}" CornerRadius="7" Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#2a5298"/></Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
            </Grid>
            <!-- Verse Count -->
            <StackPanel Name="VersePanel" Margin="0,20,0,0">
                <TextBlock Text="عدد الآيات لكل إشعار" FontSize="17" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" Foreground="#8ab4f8" Margin="0,0,0,9"/>
                <Grid>
                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="58"/></Grid.ColumnDefinitions>
                    <Slider Name="SldVerses" Grid.Column="0" Minimum="1" Maximum="300" TickFrequency="1" IsSnapToTickEnabled="True" Foreground="#3a6bc4" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBlock Name="LblVerses" Grid.Column="1" FontSize="17" Foreground="White" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" VerticalAlignment="Center" TextAlignment="Center"/>
                </Grid>
            </StackPanel>
            <!-- Full Surah checkbox -->
            <CheckBox Name="ChkFullSurah" Margin="0,14,0,0" Foreground="White" Cursor="Hand" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" FontSize="16">
                <TextBlock Text="عرض السورة كاملة (يُعطّل عدد الآيات)" Foreground="White" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" FontSize="16"/>
            </CheckBox>
            <!-- Timer multiplier -->
            <TextBlock Text="معامل وقت القراءة" FontSize="17" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" Foreground="#8ab4f8" Margin="0,20,0,9"/>
            <TextBlock FontSize="13" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" Foreground="#6a94d8" Margin="0,0,0,10" TextWrapping="Wrap"
                       Text="1.0 = الوقت الافتراضي  |  2.0 = ضعف الوقت  |  0.5 = نصف الوقت"/>
            <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="58"/></Grid.ColumnDefinitions>
                <Slider Name="SldMult" Grid.Column="0" Minimum="0.25" Maximum="3.0" TickFrequency="0.05" IsSnapToTickEnabled="True" Foreground="#3a6bc4" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBlock Name="LblMult" Grid.Column="1" FontSize="17" Foreground="White" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" VerticalAlignment="Center" TextAlignment="Center"/>
            </Grid>
            <!-- Save -->
            <Button Name="BtnSave" Content="حفظ الإعدادات" Margin="0,20,0,0"
                    FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" FontSize="14"
                    Background="#1a4a9e" Foreground="White" BorderThickness="0" Padding="0,7" Cursor="Hand">
                <Button.Template>
                    <ControlTemplate TargetType="Button">
                        <Border Name="Bd" Background="{TemplateBinding Background}" CornerRadius="9" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#2a5298"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Button.Template>
            </Button>
        </StackPanel>
    </Border>
</Window>
"@

        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
        $win    = [System.Windows.Markup.XamlReader]::Load($reader)
        $win.Tag = $PSCommandPath
        $win.FindName("RootBorder").Add_MouseLeftButtonDown({ param($s,$e) $win.DragMove() })

        $cmbMode=$win.FindName("CmbMode"); $cmbReciter=$win.FindName("CmbReciter")
        $chkFullSurah=$win.FindName("ChkFullSurah"); $versePanel=$win.FindName("VersePanel")
        $sldPct=$win.FindName("SldPct"); $lblPct=$win.FindName("LblPct")
        $sldMult=$win.FindName("SldMult"); $lblMult=$win.FindName("LblMult")
        $sldVerse=$win.FindName("SldVerses"); $lblVerse=$win.FindName("LblVerses")
        $pctPanel=$win.FindName("PctPanel"); $btnSave=$win.FindName("BtnSave")
        $btnX=$win.FindName("BtnX"); $btnMin=$win.FindName("BtnMinimize"); $btnDl=$win.FindName("BtnDownload")

        $btnMin.Add_Click({ $win.WindowState=[System.Windows.WindowState]::Minimized })
        $btnX.Add_Click({ $win.Close() })

        $nv=New-Object System.Windows.Controls.ComboBoxItem
        $nv.Content="🔇  بدون صوت"; $nv.Tag="none"; $nv.Foreground=[System.Windows.Media.Brushes]::Black
        $cmbReciter.Items.Add($nv) | Out-Null

        $recitersData=Get-Reciters
        foreach ($r in $recitersData) {
            $si2=switch($r.status){'full'{"✅"}default{"⬇"}}
            $dn=Get-ReciterDisplayName $r
            $item=New-Object System.Windows.Controls.ComboBoxItem
            $item.Content="$si2  $dn"; $item.Tag=$r.url; $item.Foreground=[System.Windows.Media.Brushes]::Black
            $cmbReciter.Items.Add($item) | Out-Null
        }

        $sel0=$settings.reciter; if([string]::IsNullOrEmpty($sel0)){$sel0="none"}
        $ok=$false
        foreach($item in $cmbReciter.Items){if($item.Tag -eq $sel0){$cmbReciter.SelectedItem=$item;$ok=$true;break}}
        if(-not $ok){$cmbReciter.SelectedIndex=0}
        $prevUrl=$sel0

        $updateDlBtn={
            $s2=$cmbReciter.SelectedItem
            if($s2 -eq $null -or $s2.Tag -eq "none"){$btnDl.Visibility='Collapsed';return}
            $btnDl.Visibility='Visible'
            $r2=$recitersData|Where-Object{$_.url -eq $s2.Tag}|Select-Object -First 1
            if($r2 -eq $null){return}
            switch($r2.status){
                'full' {$btnDl.Content="✅ مكتمل";$btnDl.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x10,0x45,0x15))}
                default{$btnDl.Content="⬇ تحميل"; $btnDl.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x10,0x3a,0x70))}
            }
        }
        & $updateDlBtn

        $cmbReciter.Add_SelectionChanged({
            & $updateDlBtn
        })

        $btnDl.Add_Click({
            $u2=$cmbReciter.SelectedItem.Tag
            if($u2 -and $u2 -ne "none"){Start-Process powershell.exe -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($win.Tag)`" -download `"$u2`""}
        })

        foreach($item in $cmbMode.Items){if($item.Tag -eq $settings.mode){$cmbMode.SelectedItem=$item;break}}
        if($cmbMode.SelectedItem -eq $null){$cmbMode.SelectedIndex=3}

        $sldPct.Value=$settings.percentage;       $lblPct.Text="$($settings.percentage)%"
        $sldMult.Value=$settings.timerMultiplier; $lblMult.Text="$($settings.timerMultiplier)x"
        $sldVerse.Value=$settings.verseCount;     $lblVerse.Text="$($settings.verseCount)"
        $chkFullSurah.IsChecked=$settings.fullSurah
        $pctPanel.Visibility=if($settings.mode -eq 'percentage'){'Visible'}else{'Collapsed'}
        $versePanel.IsEnabled=-not $settings.fullSurah
        $versePanel.Opacity=if($settings.fullSurah){0.35}else{1.0}

        $cmbMode.Add_SelectionChanged({$pctPanel.Visibility=if($cmbMode.SelectedItem.Tag -eq 'percentage'){'Visible'}else{'Collapsed'}})
        $chkFullSurah.Add_Checked({$versePanel.IsEnabled=$false;$versePanel.Opacity=0.35})
        $chkFullSurah.Add_Unchecked({$versePanel.IsEnabled=$true;$versePanel.Opacity=1.0})
        $sldPct.Add_ValueChanged({$lblPct.Text="$([int]$sldPct.Value)%"})
        $sldMult.Add_ValueChanged({$lblMult.Text="$([Math]::Round($sldMult.Value,2))x"})
        $sldVerse.Add_ValueChanged({$lblVerse.Text="$([int]$sldVerse.Value)"})

        $btnSave.Add_Click({
            $cu=$cmbReciter.SelectedItem.Tag; if(-not $cu){$cu="none"}
            if($cu -ne "none"){
                $r4=$recitersData|Where-Object{$_.url -eq $cu}|Select-Object -First 1
                if($r4 -ne $null -and $r4.status -ne 'full'){
                    foreach($item in $cmbReciter.Items){if($item.Tag -eq $prevUrl){$cmbReciter.SelectedItem=$item;break}}
                    $btnSave.Content="⚠ القارئ غير محمَّل، تم الرجوع للسابق"
                    $win.Dispatcher.InvokeAsync({Start-Sleep -Milliseconds 2000;$btnSave.Content="حفظ الإعدادات"},[System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
                    return
                }
            }
            $ns=[PSCustomObject]@{mode=$cmbMode.SelectedItem.Tag;percentage=[int]$sldPct.Value;timerMultiplier=[Math]::Round($sldMult.Value,2);verseCount=[int]$sldVerse.Value;fullSurah=[bool]$chkFullSurah.IsChecked;reciter=$cu}
            Save-Settings $ns; $btnSave.Content="✅ تم الحفظ"
            $win.Dispatcher.InvokeAsync({Start-Sleep -Milliseconds 1200;$win.Close()},[System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
        })
        $win.ShowDialog() | Out-Null
    } catch { Log "SETTINGS ERROR: $_`n$($_.ScriptStackTrace)" }
    exit
}

# =============================================================================
# MODE: DOWNLOAD MANAGER  (Start | Pause/Resume | Cancel)
# =============================================================================
if ($args -contains '-download') {
    Log "MODE: -download"
    $dlUrl = $args[$args.IndexOf('-download') + 1]
    Log "Download url key: $dlUrl"
    try {
        $reciters=Get-Reciters
        $reciterObj=$reciters|Where-Object{$_.url -eq $dlUrl}|Select-Object -First 1
        if(-not $reciterObj){Log "Reciter not found: $dlUrl";exit}

        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName PresentationCore
        Add-Type -AssemblyName WindowsBase

        $reciterName   = $reciterObj.name
        $reciterNameAr = Get-ReciterDisplayName $reciterObj
        $reciterUrl    = $reciterObj.url
        $reciterStatus = $reciterObj.status
        $safeName      = $reciterName -replace '[\\/:*?"<>|]','_'
        $destDir       = Join-Path $audioBase $safeName
        $zipUrl2       = "https://everyayah.com/data/$reciterUrl/000_versebyverse.zip"
        $zipDest       = Join-Path $destDir "000_versebyverse.zip"
        $logFilePath   = $logFile
        $recitersFilePath = $recitersFile

        Log "=== DOWNLOAD INIT ===  name=$reciterName  ar=$reciterNameAr  url=$reciterUrl  folder=$safeName  dest=$destDir  zip=$zipUrl2"

        $xamlDl = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="تحميل القرآن" Width="500" SizeToContent="Height"
        WindowStartupLocation="CenterScreen" FlowDirection="RightToLeft"
        Background="Transparent" WindowStyle="None" AllowsTransparency="True"
        ResizeMode="CanMinimize" Topmost="True">
    <Border CornerRadius="14" BorderThickness="1" BorderBrush="#3a6bc4" Name="DlRootBorder">
        <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                <GradientStop Color="#F70d1b3e" Offset="0"/><GradientStop Color="#F7071428" Offset="1"/>
            </LinearGradientBrush>
        </Border.Background>
        <StackPanel Margin="26,22,26,26">
            <Grid Margin="0,0,0,18">
                <TextBlock Name="TitleLbl" FontSize="19" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial"
                           Foreground="#8ab4f8" FontWeight="Bold" HorizontalAlignment="Right" TextWrapping="Wrap"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Left">
                    <Button Name="BtnMinDl" Content="─" Background="Transparent" BorderThickness="0"
                            Foreground="#6a94d8" FontSize="16" Cursor="Hand" Padding="6,0" Margin="0,0,4,0"/>
                    <Button Name="BtnCloseDl" Content="✕" Background="Transparent" BorderThickness="0"
                            Foreground="#6a94d8" FontSize="16" Cursor="Hand" Padding="6,0"/>
                </StackPanel>
            </Grid>
            <TextBlock Name="StatusLbl" FontSize="17" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial"
                       Foreground="#c8d8ff" TextWrapping="Wrap" Margin="0,0,0,16"
                       TextAlignment="Right" MinHeight="44" FlowDirection="LeftToRight"/>
            <Border Name="ProgContainer" CornerRadius="5" Background="#0d2240" Height="30" Margin="0,0,0,18">
                <Grid>
                    <Border Name="ProgFill" CornerRadius="5" HorizontalAlignment="Left" Width="0">
                        <Border.Background>
                            <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                                <GradientStop Color="#1a6abf" Offset="0"/><GradientStop Color="#3d9be8" Offset="1"/>
                            </LinearGradientBrush>
                        </Border.Background>
                    </Border>
                    <TextBlock Name="ProgLbl" FontSize="14" Foreground="White"
                               HorizontalAlignment="Center" VerticalAlignment="Center" FlowDirection="LeftToRight"/>
                </Grid>
            </Border>
            <StackPanel Orientation="Horizontal" FlowDirection="LeftToRight">
                <Button Name="BtnStart" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" FontSize="13"
                        Foreground="White" BorderThickness="0" Padding="12,7" Cursor="Hand" Margin="0,0,8,0">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Name="Bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#2a5298"/></Trigger>
                                <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
                <Button Name="BtnPause" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" FontSize="13"
                        Foreground="White" BorderThickness="0" Padding="12,7" Cursor="Hand"
                        Margin="0,0,8,0" IsEnabled="False" Background="#1a5a2a">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Name="Bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
                                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#2a7a3a"/></Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
                <Button Name="BtnCancel" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" FontSize="13"
                        Foreground="White" BorderThickness="0" Padding="12,7" Cursor="Hand"
                        Background="#5a1a1a" IsEnabled="False">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Name="Bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
                                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#8b2222"/></Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
            </StackPanel>
        </StackPanel>
    </Border>
</Window>
"@

        $rdr2=$([System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlDl)))
        $dlWin=[System.Windows.Markup.XamlReader]::Load($rdr2)
        $dlWin.FindName("DlRootBorder").Add_MouseLeftButtonDown({param($s,$e)$dlWin.DragMove()})

        $titleLbl =$dlWin.FindName("TitleLbl");  $statusLbl=$dlWin.FindName("StatusLbl")
        $progFill =$dlWin.FindName("ProgFill");  $progLbl  =$dlWin.FindName("ProgLbl")
        $progCont =$dlWin.FindName("ProgContainer")
        $btnStart =$dlWin.FindName("BtnStart");  $btnPause =$dlWin.FindName("BtnPause")
        $btnCancel=$dlWin.FindName("BtnCancel"); $btnCloseDl=$dlWin.FindName("BtnCloseDl")
        $btnMinDl =$dlWin.FindName("BtnMinDl")

        $btnMinDl.Add_Click({$dlWin.WindowState=[System.Windows.WindowState]::Minimized})
        $titleLbl.Text="تحميل: $reciterNameAr"

        $dlState=[PSCustomObject]@{
            IsRunning  =$false
            IsPaused   =$false
            CancelToken=[System.Threading.CancellationTokenSource]::new()
            PauseEvent =[System.Threading.ManualResetEventSlim]::new($true)
        }
        $ui=@{ Win=$dlWin; Status=$statusLbl; Fill=$progFill; Lbl=$progLbl; Cont=$progCont
               BtnStart=$btnStart; BtnPause=$btnPause; BtnCancel=$btnCancel }

        $applyBtnState={
            param([string]$state)
            switch($state){
                'idle'   {
                    $ui.BtnStart.Content="⬇ تحميل";$ui.BtnStart.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x10,0x3a,0x70));$ui.BtnStart.IsEnabled=$true
                    $ui.BtnPause.Content="⏸ إيقاف";$ui.BtnPause.IsEnabled=$false
                    $ui.BtnCancel.Content="✕ إلغاء";$ui.BtnCancel.IsEnabled=$false
                }
                'running'{
                    $ui.BtnStart.IsEnabled=$false
                    $ui.BtnPause.Content="⏸ إيقاف";$ui.BtnPause.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x1a,0x5a,0x2a));$ui.BtnPause.IsEnabled=$true
                    $ui.BtnCancel.Content="✕ إلغاء";$ui.BtnCancel.IsEnabled=$true
                }
                'paused' {
                    $ui.BtnStart.IsEnabled=$false
                    $ui.BtnPause.Content="▶ متابعة";$ui.BtnPause.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x70,0x50,0x00));$ui.BtnPause.IsEnabled=$true
                    $ui.BtnCancel.Content="✕ إلغاء";$ui.BtnCancel.IsEnabled=$true
                }
                'done'   {
                    $ui.BtnStart.Content="↺ إعادة التحميل";$ui.BtnStart.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x20,0x20,0x60));$ui.BtnStart.IsEnabled=$true
                    $ui.BtnPause.IsEnabled=$false;$ui.BtnCancel.IsEnabled=$false
                }
                'error'  {
                    $ui.BtnStart.Content="↺ إعادة المحاولة";$ui.BtnStart.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x10,0x3a,0x70));$ui.BtnStart.IsEnabled=$true
                    $ui.BtnPause.IsEnabled=$false;$ui.BtnCancel.IsEnabled=$false
                }
            }
        }

        $dlScript={
            param($ui,$zipUrl2,$zipDest,$destDir,$reciterUrl,$reciterName,$recitersFilePath,$logFilePath,$dlState,$applyBtnState)
            function BgLog($m){
                $ts=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                $line="$ts  [DL] $m"
                $retries=0
                while($retries -lt 5){
                    try{
                        $fs=[System.IO.File]::Open($logFilePath,[System.IO.FileMode]::Append,[System.IO.FileAccess]::Write,[System.IO.FileShare]::ReadWrite)
                        $wr=New-Object System.IO.StreamWriter($fs,[System.Text.Encoding]::UTF8)
                        $wr.WriteLine($line);$wr.Close();$fs.Close();break
                    }catch{$retries++;Start-Sleep -Milliseconds 50}
                }
            }
            function UI([scriptblock]$sb){$ui.Win.Dispatcher.Invoke([Action]$sb)}
            function UpdStatus([string]$msg,[double]$pct){
                UI {
                    $ui.Status.Text=$msg
                    if($pct -ge 0){$w=[Math]::Max(0,$ui.Cont.ActualWidth*$pct/100);$ui.Fill.Width=$w;$ui.Lbl.Text="$([int]$pct)%"}
                }
            }
            function SaveStatus([string]$uk,[string]$st){
                try{
                    $lst=Get-Content $recitersFilePath -Encoding UTF8|ConvertFrom-Json
                    foreach($r in $lst){if($r.url -eq $uk){$r.status=$st;break}}
                    $lst|ConvertTo-Json -Depth 3|Out-File -FilePath $recitersFilePath -Encoding UTF8
                    BgLog "Status saved: $uk -> $st"
                }catch{BgLog "SaveStatus error: $_"}
            }

            BgLog "=== THREAD START  url=$zipUrl2  dest=$destDir"
            UpdStatus "جارٍ إنشاء المجلد" -1
            $ws=$null;$rs2=$null;$rsp=$null

            try{
                if(-not(Test-Path $destDir)){New-Item -ItemType Directory -Path $destDir -Force|Out-Null;BgLog "Dir created"}else{BgLog "Dir exists"}
                UpdStatus "جارٍ الاتصال بالخادم" -1
                BgLog "Building request: $zipUrl2"
                [System.Net.ServicePointManager]::SecurityProtocol=[System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
                $req=[System.Net.HttpWebRequest]::CreateHttp($zipUrl2)
                $req.UserAgent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                $req.Timeout=30000;$req.ReadWriteTimeout=300000;$req.AllowAutoRedirect=$true;$req.KeepAlive=$false
                BgLog "GetResponse..."
                $rsp=$req.GetResponse()
                if($rsp -eq $null){throw "No response from server"}
                $statusCode=[int]$rsp.StatusCode
                BgLog "HTTP $statusCode $($rsp.StatusDescription)"
                $total=$rsp.ContentLength
                BgLog "Content-Length=$total ($([Math]::Round($total/1048576,1)) MB)"
                try{$ws=[System.IO.File]::Create($zipDest);BgLog "File open: $zipDest"}catch{BgLog "File open error: $_";throw}
                $rs2=$rsp.GetResponseStream()
                $buf=New-Object byte[] 65536;$dl=[long]0;$lastPct=-1;$lastLogMb=0
                BgLog "Read loop start..."
                UI{& $applyBtnState 'running'}

                while($true){
                    if($dlState.CancelToken.IsCancellationRequested){BgLog "Cancelled at $dl bytes";break}
                    if($dlState.IsPaused){
                        BgLog "Paused at $dl bytes"
                        while($dlState.IsPaused){
                            if($dlState.CancelToken.IsCancellationRequested){break}
                            $dlState.PauseEvent.Wait(200) | Out-Null
                        }
                        if($dlState.CancelToken.IsCancellationRequested){BgLog "Cancelled while paused";break}
                        BgLog "Resumed at $dl bytes"
                    }
                    $rd=0
                    try{$rd=$rs2.Read($buf,0,$buf.Length)}catch{BgLog "Read error: $_";throw}
                    if($rd -le 0){BgLog "EOF total=$dl bytes";break}
                    $ws.Write($buf,0,$rd);$dl+=$rd
                    $cm=[Math]::Floor($dl/(10*1048576))
                    if($cm -gt $lastLogMb){$lastLogMb=$cm;BgLog "Downloaded: $([Math]::Round($dl/1048576,1)) MB"}
                    if($total -gt 0){
                        $p=[Math]::Round($dl/$total*100,1)
                        if([int]$p -ne $lastPct){
                            $lastPct=[int]$p
                            $mb=[Math]::Round($dl/1048576.0,1);$tot=[Math]::Round($total/1048576.0,1)
                            UpdStatus "$mb MB / $tot MB" $p
                        }
                    }else{UpdStatus "$([Math]::Round($dl/1048576.0,1)) MB" -1}
                }

                $ws.Flush();$ws.Close();$rs2.Close();$rsp.Close();$ws=$null;$rs2=$null;$rsp=$null
                BgLog "Write complete: $dl bytes"

                if($dlState.CancelToken.IsCancellationRequested){
                    if(Test-Path $zipDest){Remove-Item $zipDest -Force -EA SilentlyContinue}
                    SaveStatus $reciterUrl 'missing'
                    UpdStatus "تم الإلغاء" 0
                    UI{& $applyBtnState 'idle'}
                    return
                }
                if(Test-Path $zipDest){$fsz=(Get-Item $zipDest).Length;BgLog "Zip on disk: $fsz bytes";if($fsz -lt 1000000){BgLog "WARNING small zip"}}
                else{BgLog "ERROR: zip missing after write!";throw "Zip missing"}

                BgLog "Extracting..."
                UpdStatus "جارٍ فك الضغط" 100
                try{Add-Type -AssemblyName System.IO.Compression.FileSystem;[System.IO.Compression.ZipFile]::ExtractToDirectory($zipDest,$destDir);BgLog "Extracted"}
                catch{BgLog "Extract error: $_";throw}
                $cnt=(Get-ChildItem -Path $destDir -Filter "*.mp3" -EA SilentlyContinue).Count
                BgLog "MP3s: $cnt"
                Remove-Item $zipDest -Force -EA SilentlyContinue;BgLog "Zip deleted"
                SaveStatus $reciterUrl 'full'
                UpdStatus "✅ اكتمل التحميل" 100
                UI{
                    & $applyBtnState 'done'
                    $refreshed = @(Get-Content $recitersFilePath -Encoding UTF8 | ConvertFrom-Json)
                    $ui.Win.Dispatcher.Invoke([Action]{
                    })
                }
                BgLog "=== DOWNLOAD COMPLETE for $reciterName ==="

            }catch{
                $em=$_.Exception.Message
                $statusDetails=""
                if($_.Exception -is [System.Net.WebException] -and $_.Exception.Response){
                    $statusDetails=" | HTTP $([int]$_.Exception.Response.StatusCode) $($_.Exception.Response.StatusDescription)"
                }
                BgLog "=== EXCEPTION: $em$statusDetails`nStack: $($_.ScriptStackTrace)"
                try{if($ws){$ws.Close()}}catch{}
                try{if($rs2){$rs2.Close()}}catch{}
                try{if($rsp){$rsp.Close()}}catch{}
                if(Test-Path $zipDest){Remove-Item $zipDest -Force -EA SilentlyContinue}
                SaveStatus $reciterUrl 'missing'
                UpdStatus "خطأ: $em$statusDetails" 0
                UI{& $applyBtnState 'error'}
            }finally{
                $dlState.IsRunning=$false
                BgLog "Thread exit"
            }
        }

        & $applyBtnState 'idle'
        if($reciterStatus -eq 'full'){$statusLbl.Text="✅ مكتمل — الملفات موجودة";&$applyBtnState 'done'}
        else{$statusLbl.Text="الملفات غير موجودة — اضغط تحميل للبدء"}

        $btnStart.Add_Click({
            if($dlState.IsRunning){return}
            $dlState.IsRunning=$true;$dlState.IsPaused=$false
            $dlState.CancelToken=[System.Threading.CancellationTokenSource]::new()
            $dlState.PauseEvent.Set()
            Log "Starting download runspace for: $reciterName"
            $rs3=[System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
            $rs3.ApartmentState=[System.Threading.ApartmentState]::STA
            $rs3.ThreadOptions=[System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
            $rs3.Open()
            $ps3=[System.Management.Automation.PowerShell]::Create();$ps3.Runspace=$rs3
            $ps3.AddScript($dlScript).AddArgument($ui).AddArgument($zipUrl2).AddArgument($zipDest).AddArgument($destDir).AddArgument($reciterUrl).AddArgument($reciterName).AddArgument($recitersFilePath).AddArgument($logFile).AddArgument($dlState).AddArgument($applyBtnState) | Out-Null
            $hdl=$ps3.BeginInvoke()
            $mt=New-Object System.Windows.Threading.DispatcherTimer;$mt.Interval=[TimeSpan]::FromSeconds(1)
            $mt.Add_Tick({
                if($hdl.IsCompleted){
                    $mt.Stop();Log "DL runspace done"
                    if($ps3.HadErrors){foreach($er in $ps3.Streams.Error){Log "Runspace err: $er"}}
                    try{$ps3.EndInvoke($hdl)}catch{Log "EndInvoke: $_"}
                    try{$rs3.Close()}catch{}
                    $refreshed2=Get-Reciters
                    $rObj2=$refreshed2|Where-Object{$_.url -eq $reciterUrl}|Select-Object -First 1
                    if($rObj2 -ne $null -and $rObj2.status -eq 'full'){
                        $statusLbl.Text="✅ اكتمل التحميل"
                    }
                }
            })
            $mt.Start()
        })

        $btnPause.Add_Click({
            if(-not $dlState.IsRunning){return}
            if($dlState.IsPaused){
                $dlState.IsPaused=$false;$dlState.PauseEvent.Set()
                $statusLbl.Text="جارٍ التحميل";&$applyBtnState 'running';Log "Resumed"
            }else{
                $dlState.IsPaused=$true;$dlState.PauseEvent.Reset()
                $statusLbl.Text="⏸ متوقف مؤقتاً — اضغط متابعة للاستمرار";&$applyBtnState 'paused';Log "Paused"
            }
        })

        $btnCancel.Add_Click({
            Log "Cancel clicked"
            $dlState.IsPaused=$false;$dlState.PauseEvent.Set();$dlState.CancelToken.Cancel()
            $statusLbl.Text="جارٍ الإلغاء";$btnCancel.IsEnabled=$false;$btnPause.IsEnabled=$false
        })

        $btnCloseDl.Add_Click({
            Log "DL close"
            $dlState.IsPaused=$false;$dlState.PauseEvent.Set();$dlState.CancelToken.Cancel()
            $dlWin.Close()
        })

        $dlWin.ShowDialog() | Out-Null
    } catch { Log "DOWNLOAD MODE ERROR: $_`n$($_.ScriptStackTrace)" }
    exit
}

# =============================================================================
# MODE: POPUP FROM FILE
# =============================================================================
if ($args -contains '-popupfile') {
    Log "MODE: -popupfile"
    $payloadFile = $args[$args.IndexOf('-popupfile') + 1]
    Log "Payload file: $payloadFile  exists=$(Test-Path $payloadFile)"
    if (-not (Test-Path $payloadFile)) { Log "PAYLOAD FILE MISSING"; exit }
    $payloadJson = Get-Content $payloadFile -Encoding UTF8 -Raw
    & $PSCommandPath -popup $payloadJson
    exit
}

# =============================================================================
# MODE: POPUP
# =============================================================================
if ($args -contains '-popup') {
    Log "MODE: -popup"
    try {
        $payloadJson=$args[$args.IndexOf('-popup')+1]
        $payload=$payloadJson|ConvertFrom-Json
        $scriptPath=$payload.scriptPath; $headerTitle=$payload.headerTitle; $verseText=$payload.verseText
        $audioList=$payload.audioList; $surahFile=$payload.surahFile
        $isKhatma=[bool]$payload.isKhatma; $readSeconds=[int]$payload.readSeconds; $hasAudio=[bool]$payload.hasAudio
        Log "headerTitle=$headerTitle  isKhatma=$isKhatma  readSeconds=$readSeconds  hasAudio=$hasAudio"

        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName PresentationCore
        Add-Type -AssemblyName WindowsBase

        $headerTitleXml=[System.Security.SecurityElement]::Escape($headerTitle)
        $khatmaRowVis=if($isKhatma){'Visible'}else{'Collapsed'}

        $noAudioXmlStr = '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="تلاوة" Width="440" SizeToContent="Height" WindowStartupLocation="CenterScreen" FlowDirection="RightToLeft" Background="Transparent" WindowStyle="None" AllowsTransparency="True" ResizeMode="NoResize" Topmost="True"><Border CornerRadius="12" BorderThickness="1" BorderBrush="#3a6bc4"><Border.Background><LinearGradientBrush StartPoint="0,0" EndPoint="0,1"><GradientStop Color="#F70d1b3e" Offset="0"/><GradientStop Color="#F7071428" Offset="1"/></LinearGradientBrush></Border.Background><StackPanel Margin="28,26,28,26"><TextBlock Text="&#x1F507;" FontSize="34" Foreground="White" HorizontalAlignment="Center" Margin="0,0,0,14"/><TextBlock FontSize="17" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" Foreground="#c8d8ff" TextWrapping="Wrap" TextAlignment="Center" Margin="0,0,0,8">لم تقم باختيار قارئ بعد</TextBlock><TextBlock FontSize="15" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" Foreground="#8ab4f8" TextWrapping="Wrap" TextAlignment="Center" Margin="0,0,0,22">اذهب إلى الإعدادات واختر قارئك المفضل&#x0a;سيتطلب ذلك تحميل ملفات التلاوة الصوتية للقارئ</TextBlock><StackPanel Orientation="Horizontal" HorizontalAlignment="Center" FlowDirection="LeftToRight"><Button Name="BtnGoSettings" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" FontSize="13" Foreground="White" BorderThickness="0" Padding="12,7" Cursor="Hand" Background="#1a4a9e" Margin="0,0,10,0"><Button.Template><ControlTemplate TargetType="Button"><Border Name="Bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#2a5298"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Button.Template>&#x2699; فتح الإعدادات</Button><Button Name="BtnDismiss" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial" FontSize="13" Foreground="White" BorderThickness="0" Padding="12,7" Cursor="Hand" Background="#2a2a3a"><Button.Template><ControlTemplate TargetType="Button"><Border Name="Bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#44445a"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Button.Template>إغلاق</Button></StackPanel></StackPanel></Border></Window>'

        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="&#x062A;&#x0646;&#x0628;&#x064A;&#x0647; &#x0627;&#x0644;&#x0642;&#x0631;&#x0622;&#x0646;"
        Width="478" SizeToContent="Height" MaxHeight="770"
        WindowStartupLocation="Manual" FlowDirection="RightToLeft"
        Background="Transparent" Foreground="White" Topmost="True"
        ResizeMode="NoResize" WindowStyle="None" AllowsTransparency="True"
        ShowActivated="False" Focusable="False">
    <Window.Resources>
        <Style x:Key="Btn" TargetType="Button">
            <Setter Property="Foreground" Value="White"/><Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10,2"/><Setter Property="FontSize" Value="15"/>
            <Setter Property="FontFamily" Value="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial"/><Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Name="Bd" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#2a5298"/></Trigger></Style.Triggers>
        </Style>
    </Window.Resources>
    <Border Name="MainBorder" CornerRadius="14" BorderThickness="1" BorderBrush="#3a6bc4">
        <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                <GradientStop Color="#F70d1b3e" Offset="0"/><GradientStop Color="#F7071428" Offset="1"/>
            </LinearGradientBrush>
        </Border.Background>
        <Border.Effect><DropShadowEffect BlurRadius="28" ShadowDepth="8" Color="#88000033" Opacity="0.0"/></Border.Effect>
        <Grid Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Image Name="WatermarkImg" Grid.Row="0" Grid.RowSpan="3" HorizontalAlignment="Center" VerticalAlignment="Center"
                   Width="260" Height="260" Opacity="0.058" RenderOptions.BitmapScalingMode="HighQuality" IsHitTestVisible="False"/>
            <!-- FIXED HEADER: Auto title + * basmala + Auto countdown + Auto close -->
            <Grid Grid.Row="0" Margin="0,0,0,10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="$headerTitleXml" FontSize="16" Foreground="#8ab4f8"
                           FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial"
                           VerticalAlignment="Center" TextWrapping="NoWrap" TextAlignment="Right" Margin="0,0,10,0"/>
                <TextBlock Grid.Column="1"
                           Text="&#x0628;&#x0650;&#x0633;&#x0652;&#x0645;&#x0650; &#x0627;&#x0644;&#x0644;&#x0651;&#x064E;&#x0647;&#x0650; &#x0627;&#x0644;&#x0631;&#x0651;&#x064E;&#x062D;&#x0652;&#x0645;&#x064E;&#x0670;&#x0646;&#x0650; &#x0627;&#x0644;&#x0631;&#x0651;&#x064E;&#x062D;&#x0650;&#x064A;&#x0645;&#x0650;"
                           FontSize="18" FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial"
                           Foreground="#c8d8ff" TextAlignment="Center" VerticalAlignment="Center"
                           HorizontalAlignment="Stretch"/>
                <TextBlock Grid.Column="2" Name="CountdownLabel" FontSize="11" Foreground="#6a94d8"
                           VerticalAlignment="Center" Margin="0,0,6,0" FlowDirection="LeftToRight"/>
                <Button Grid.Column="3" Name="BtnClose" Content="&#x2715;" Style="{StaticResource Btn}"
                        Background="Transparent" FontSize="15" Padding="8,4" Foreground="#8ab4f8"/>
            </Grid>
            <Grid Grid.Row="1" Margin="0,0,0,14">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="10"/><ColumnDefinition Width="6"/><ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Canvas Grid.Column="0" Name="ScrollTrackCanvas" VerticalAlignment="Stretch" IsHitTestVisible="False">
                    <Rectangle Name="ScrollTrack" Width="4" RadiusX="2" RadiusY="2" Canvas.Left="3">
                        <Rectangle.Fill><SolidColorBrush Color="#1a3a52" Opacity="0.6"/></Rectangle.Fill>
                    </Rectangle>
                    <Rectangle Name="ScrollThumb" Width="4" RadiusX="2" RadiusY="2" Canvas.Left="3">
                        <Rectangle.Fill><SolidColorBrush Color="#3d6e96" Opacity="0.85"/></Rectangle.Fill>
                    </Rectangle>
                </Canvas>
                <ScrollViewer Grid.Column="2" Name="VerseScroller" MaxHeight="770"
                              VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Disabled"
                              CanContentScroll="False" PanningMode="VerticalOnly" PanningDeceleration="0.001" PanningRatio="1">
                    <TextBlock Name="VerseBlock" FontSize="27"
                               FontFamily="Scheherazade New, KFGQPC Uthmanic Script HAFS, Traditional Arabic, Arial"
                               TextWrapping="Wrap" LineHeight="56" Foreground="White" TextAlignment="Right"/>
                </ScrollViewer>
            </Grid>
            <Grid Grid.Row="2">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Image Grid.Column="0" Name="QuranIcon" Width="28" Height="28" Margin="0,0,10,0"
                       VerticalAlignment="Center" RenderOptions.BitmapScalingMode="HighQuality"/>
                <StackPanel Grid.Column="2" Orientation="Horizontal" FlowDirection="LeftToRight">
                    <Button Name="BtnSettings" Content="⚙ الإعدادات" Style="{StaticResource Btn}"
                            Background="#1a3260" Padding="8,2" Margin="0,0,6,0" Foreground="White"/>
                    <Button Name="BtnPlay" Content="🔊 تلاوة" Style="{StaticResource Btn}"
                            Background="#1a4a9e" Margin="0,0,6,0"/>
                    <Button Name="BtnSurah" Content="📖 السورة" Style="{StaticResource Btn}"
                            Background="#0d2a5e" Margin="0,0,6,0"/>
                    <Button Name="BtnRead" Content="✔️ قرأتُ" Style="{StaticResource Btn}"
                            Background="#145214" Visibility="$khatmaRowVis"/>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

        $reader=[System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
        $window=[System.Windows.Markup.XamlReader]::Load($reader)
        Log "Window loaded OK"

        $btnClose=$window.FindName("BtnClose"); $btnPlay=$window.FindName("BtnPlay")
        $btnSurah=$window.FindName("BtnSurah"); $btnRead=$window.FindName("BtnRead")
        $btnSettings=$window.FindName("BtnSettings"); $cntLabel=$window.FindName("CountdownLabel")
        $screen=[System.Windows.SystemParameters]::WorkArea

        $slideIn={
            $ft=$screen.Bottom-$window.ActualHeight-16
            $window.Left=$screen.Right-$window.ActualWidth-16;$window.Top=$ft+60;$window.Opacity=0
            for($i=1;$i-le 18;$i++){$e=1-[Math]::Pow(1-($i/18),3);$window.Top=$ft+60*(1-$e);$window.Opacity=$e;Start-Sleep -Milliseconds 18}
            $window.Top=$ft;$window.Opacity=1
        }
        $slideOut={
            $st=$window.Top
            for($i=1;$i-le 14;$i++){$e=[Math]::Pow($i/14,3);$window.Top=$st+60*$e;$window.Opacity=1-$e;Start-Sleep -Milliseconds 16}
            $window.Close()
        }
        $window.Left=-9999;$window.Top=-9999

        $cd=[PSCustomObject]@{SecondsLeft=$readSeconds}
        $countdownTimer=New-Object System.Windows.Threading.DispatcherTimer
        $countdownTimer.Interval=[TimeSpan]::FromSeconds(1)
        $countdownTimer.Add_Tick({$cd.SecondsLeft--;$cntLabel.Text="$($cd.SecondsLeft)s";if($cd.SecondsLeft -le 0){$countdownTimer.Stop()}})

        $mouseOver=[PSCustomObject]@{Active=$false;ScrollKilled=$false}
        $ss=[PSCustomObject]@{Phase=0;Ticks=0;Sv=$null;Total=0.0;Offset=0.0;TargetStep=0.0;CurrentStep=0.0}
        $masterTimer=New-Object System.Windows.Threading.DispatcherTimer
        $masterTimer.Interval=[TimeSpan]::FromMilliseconds(16)
        $masterTimer.Add_Tick({
            $ss.Ticks++
            if($ss.Phase -eq 0){
                if($ss.Ticks -ge 10){$ss.Sv=$window.FindName("VerseScroller");if($ss.Sv -ne $null -and $ss.Sv.ScrollableHeight -gt 0){$ss.Phase=2;$ss.Ticks=0}else{$masterTimer.Stop()}}
            }elseif($ss.Phase -eq 2){
                if($ss.Ticks -ge 1250){$ss.Total=$ss.Sv.ScrollableHeight;$ss.TargetStep=$ss.Total/300.0;$ss.CurrentStep=0.0;$ss.Offset=$ss.Sv.VerticalOffset;$ss.Phase=3}
            }elseif($ss.Phase -eq 3){
                if($mouseOver.ScrollKilled){$masterTimer.Stop();return}
                $ss.CurrentStep+=($ss.TargetStep-$ss.CurrentStep)*0.08;$ss.Offset+=$ss.CurrentStep
                $ss.Sv.ScrollToVerticalOffset($ss.Offset);& $updateScrollbar
                if($ss.Offset -ge $ss.Total){$masterTimer.Stop()}
            }
        })

        $wh=[PSCustomObject]@{TargetOffset=0.0;Sv=$null}
        $sb=[PSCustomObject]@{Track=$null;Thumb=$null;Canvas=$null}
        $updateScrollbar={
            if($sb.Track -eq $null -or $wh.Sv -eq $null){return}
            $th=$sb.Track.Height;if($th -le 0){return}
            $tot=$wh.Sv.ScrollableHeight;if($tot -le 0){$sb.Thumb.Visibility='Collapsed';return}
            $sb.Thumb.Visibility='Visible'
            $thumbH=[Math]::Max(20,$th*$wh.Sv.ViewportHeight/$wh.Sv.ExtentHeight);$sb.Thumb.Height=$thumbH
            [System.Windows.Controls.Canvas]::SetTop($sb.Thumb,($wh.Sv.VerticalOffset/$tot)*($th-$thumbH))
        }
        $wheelTimer=New-Object System.Windows.Threading.DispatcherTimer
        $wheelTimer.Interval=[TimeSpan]::FromMilliseconds(16)
        $wheelTimer.Add_Tick({
            $cur=$wh.Sv.VerticalOffset;$diff=$wh.TargetOffset-$cur
            if([Math]::Abs($diff) -lt 0.5){$wh.Sv.ScrollToVerticalOffset($wh.TargetOffset);$wheelTimer.Stop()}
            else{$wh.Sv.ScrollToVerticalOffset($cur+$diff*0.18)}
            & $updateScrollbar
        })

        $window.Add_ContentRendered({
            Log "ContentRendered"
            $cntLabel.Text="${readSeconds}s";$wh.Sv=$window.FindName("VerseScroller")
            $sb.Track=$window.FindName("ScrollTrack");$sb.Thumb=$window.FindName("ScrollThumb");$sb.Canvas=$window.FindName("ScrollTrackCanvas")
            $sb.Canvas.Add_SizeChanged({$sb.Track.Height=$sb.Canvas.ActualHeight;[System.Windows.Controls.Canvas]::SetTop($sb.Track,0);& $updateScrollbar})
            $wh.Sv.Add_ScrollChanged({& $updateScrollbar})
            $wh.Sv.Add_PreviewMouseWheel({
                param($s2,$e2);$e2.Handled=$true;$mouseOver.ScrollKilled=$true;$masterTimer.Stop()
                $wh.TargetOffset=[Math]::Max(0,[Math]::Min($wh.Sv.ScrollableHeight,$wh.Sv.VerticalOffset+(-$e2.Delta*0.6)))
                if(-not $wheelTimer.IsEnabled){$wheelTimer.Start()}
            })
            $pngPath=Join-Path (Split-Path $scriptPath -Parent) "images/icon.png"
            if(Test-Path $pngPath){
                try{
                    $bmp=New-Object System.Windows.Media.Imaging.BitmapImage
                    $bmp.BeginInit();$bmp.UriSource=[Uri]::new($pngPath);$bmp.CacheOption=[System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad;$bmp.EndInit()
                    $ic=$window.FindName("QuranIcon");if($ic){$ic.Source=$bmp}
                    $wm2=$window.FindName("WatermarkImg");if($wm2){$wm2.Source=$bmp}
                }catch{Log "PNG error: $_"}
            }
            $vb=$window.FindName("VerseBlock")
            if($vb){
                $vb.Inlines.Clear()
                $parts=$verseText -split '(\ufd3f\d+\ufd3e)'
                foreach($part in $parts){
                    if($part -match '^\ufd3f(\d+)\ufd3e$'){
                        $ntb=New-Object System.Windows.Controls.TextBlock
                        $ntb.Text=$part;$ntb.FontSize=15;$ntb.Foreground=[System.Windows.Media.Brushes]::LightSteelBlue
                        $ntb.FontFamily=$vb.FontFamily;$ntb.Margin=[System.Windows.Thickness]::new(6,0,6,-10)
                        $c=New-Object System.Windows.Documents.InlineUIContainer
                        $c.Child=$ntb;$c.BaselineAlignment=[System.Windows.BaselineAlignment]::Center;$vb.Inlines.Add($c)
                    }elseif($part.Length -gt 0){$run=New-Object System.Windows.Documents.Run;$run.Text=$part.Trim(' ');$vb.Inlines.Add($run)}
                }
            }
            $window.Dispatcher.InvokeAsync($slideIn,[System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
            $countdownTimer.Start();$masterTimer.Start()
            $wavName=if($isKhatma){"audio/notificationKhitma.wav"}else{"audio/notificationRandom.wav"}
            $wavPath=Join-Path (Split-Path $scriptPath -Parent) $wavName
            if(Test-Path $wavPath){
                try{
                    Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
public class MCINotif { [DllImport("winmm.dll",CharSet=CharSet.Auto)] public static extern int mciSendString(string cmd,string ret,int retLen,IntPtr hwnd); }
"@ -ErrorAction SilentlyContinue
                    [MCINotif]::mciSendString("close qurannotif",$null,0,[IntPtr]::Zero)|Out-Null
                    [MCINotif]::mciSendString("open `"$wavPath`" type mpegvideo alias qurannotif",$null,0,[IntPtr]::Zero)|Out-Null
                    [MCINotif]::mciSendString("play qurannotif",$null,0,[IntPtr]::Zero)|Out-Null
                    Log "Playing: $wavName"
                }catch{Log "WAV error: $_"}
            }
        })

        $doClose={
            $masterTimer.Stop();$countdownTimer.Stop();$wheelTimer.Stop()
            if($playState -ne $null -and $playState.IsPlaying){
                $playPollTimer.Stop(); $playState.IsPlaying=$false
                Start-Process powershell.exe -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -stopplay"
            }
            try{[MCINotif]::mciSendString("stop qurannotif",$null,0,[IntPtr]::Zero)|Out-Null;[MCINotif]::mciSendString("close qurannotif",$null,0,[IntPtr]::Zero)|Out-Null}catch{}
            $window.Dispatcher.InvokeAsync($slideOut,[System.Windows.Threading.DispatcherPriority]::Background)|Out-Null
        }

        $btnClose.Add_Click($doClose)
        $btnSurah.Add_Click({Start-Process "notepad.exe" -ArgumentList $surahFile})

        $playState = [PSCustomObject]@{ IsPlaying=$false }
        $playPidFile = Join-Path $env:TEMP "quran_play.pid"

        $playPollTimer = New-Object System.Windows.Threading.DispatcherTimer
        $playPollTimer.Interval = [TimeSpan]::FromMilliseconds(500)
        $playPollTimer.Add_Tick({
            if($playState.IsPlaying -and -not (Test-Path $playPidFile)){
                $playState.IsPlaying=$false
                $btnPlay.Content="🔊 تلاوة"
                $btnPlay.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x1a,0x4a,0x9e))
                $playPollTimer.Stop()
                Log "Play ended naturally, button reset"
            }
        })

        $btnPlay.Add_Click({
            if($playState.IsPlaying){
                Log "Stop clicked"
                $playPollTimer.Stop()
                $playState.IsPlaying=$false
                $btnPlay.Content="🔊 تلاوة"
                $btnPlay.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x1a,0x4a,0x9e))
                Start-Process powershell.exe -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -stopplay"
                return
            }
            $liveSettings = Get-Settings
            $liveAudioList = $audioList
            $liveHasAudio = $false
            if($liveSettings.reciter -ne "" -and $liveSettings.reciter -ne "none"){
                $liveReciters = Get-Reciters
                $liveReciter = $liveReciters | Where-Object{$_.url -eq $liveSettings.reciter} | Select-Object -First 1
                if($liveReciter -ne $null -and $liveReciter.status -eq 'full'){
                    $liveSafeName = $liveReciter.name -replace '[\\/:*?"<>|]','_'
                    $liveMp3Folder = Join-Path $audioBase $liveSafeName
                    if(Test-Path $liveMp3Folder){
                        $liveHasAudio = $true
                        if(-not $hasAudio -or $liveSettings.reciter -ne $liveSettings.reciter){
                            $liveFiles = @()
                            $liveBas = Join-Path $liveMp3Folder "001001.mp3"
                            if(Test-Path $liveBas){$liveFiles += $liveBas}
                            foreach($f in ($audioList -split ';')){
                                $fname = Split-Path $f -Leaf
                                $newPath = Join-Path $liveMp3Folder $fname
                                if(Test-Path $newPath){$liveFiles += $newPath}
                            }
                            if($liveFiles.Count -gt 0){$liveAudioList = $liveFiles -join ';'}
                        }
                    }
                }
            }
            if(-not $liveHasAudio){
                Log "Showing no-audio prompt"
                try{
                    $naReader=[System.Xml.XmlReader]::Create([System.IO.StringReader]::new($noAudioXmlStr))
                    $naWin=[System.Windows.Markup.XamlReader]::Load($naReader)
                    $naWin.FindName("BtnDismiss").Add_Click({$naWin.Close()})
                    $naWin.FindName("BtnGoSettings").Add_Click({
                        $naWin.Close()
                        Start-Process powershell.exe -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -settings"
                    })
                    $naWin.ShowDialog()|Out-Null
                }catch{Log "No-audio prompt error: $_`n$($_.ScriptStackTrace)"}
                return
            }
            $btnPlay.Content="⏳"
            $btnPlay.IsEnabled=$false
            Start-Process powershell.exe -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -play `"$liveAudioList`""
            $window.Dispatcher.InvokeAsync({
                $waited=0
                while(-not (Test-Path $playPidFile) -and $waited -lt 3000){ Start-Sleep -Milliseconds 100; $waited+=100 }
                $btnPlay.IsEnabled=$true
                if(Test-Path $playPidFile){
                    $playState.IsPlaying=$true
                    $btnPlay.Content="⏹ إيقاف"
                    $btnPlay.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x7a,0x1a,0x1a))
                    $playPollTimer.Start()
                    Log "Play started, button set to Stop"
                }else{
                    $btnPlay.Content="🔊 تلاوة"
                    $btnPlay.Background=[System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x1a,0x4a,0x9e))
                    Log "Play process did not start in time"
                }
            },[System.Windows.Threading.DispatcherPriority]::Background)|Out-Null
        })

        $btnSettings.Add_Click({
            Start-Process powershell.exe -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -settings"
        })

        if($btnRead){
            $btnRead.Add_Click({
                Log "Read clicked";$btnRead.IsEnabled=$false;$btnRead.Content="✅ حفظ"
                $mb=$window.FindName("MainBorder")
                if($mb){$mb.BorderBrush=[System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x00,0xCC,0x44))}
                $window.Dispatcher.InvokeAsync({
                    $nw=Join-Path (Split-Path $scriptPath -Parent) "audio/next.wav"
                    if(Test-Path $nw){
                        try{
                            [MCINotif]::mciSendString("close qurannotif",$null,0,[IntPtr]::Zero)|Out-Null
                            [MCINotif]::mciSendString("open `"$nw`" type mpegvideo alias qurannotif",$null,0,[IntPtr]::Zero)|Out-Null
                            [MCINotif]::mciSendString("play qurannotif",$null,0,[IntPtr]::Zero)|Out-Null
                        }catch{Log "audio/next.wav error: $_"}
                    }
                    Start-Process powershell.exe -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -advancekhatma" -Wait
                    Start-Sleep -Milliseconds 800;& $doClose
                },[System.Windows.Threading.DispatcherPriority]::Background)|Out-Null
            })
        }

        $autoCloseTimer=New-Object System.Windows.Threading.DispatcherTimer
        $autoCloseTimer.Interval=[TimeSpan]::FromSeconds($readSeconds)
        $autoCloseTimer.Add_Tick({
            Log "Auto-close";$autoCloseTimer.Stop();$masterTimer.Stop();$countdownTimer.Stop();$wheelTimer.Stop()
            try{[MCINotif]::mciSendString("stop qurannotif",$null,0,[IntPtr]::Zero)|Out-Null;[MCINotif]::mciSendString("close qurannotif",$null,0,[IntPtr]::Zero)|Out-Null}catch{}
            $window.Dispatcher.InvokeAsync($slideOut,[System.Windows.Threading.DispatcherPriority]::Background)|Out-Null
        })
        $autoCloseTimer.Start()

        $window.Add_MouseEnter({$mouseOver.Active=$true;$mouseOver.ScrollKilled=$true;$masterTimer.Stop();$autoCloseTimer.Stop();$countdownTimer.Stop()})
        $window.Add_MouseLeave({
            $mouseOver.Active=$false
            if($cd.SecondsLeft -gt 0){$autoCloseTimer.Stop();$autoCloseTimer.Interval=[TimeSpan]::FromSeconds($cd.SecondsLeft);$autoCloseTimer.Start();$countdownTimer.Start()}
        })

        $window.ShowDialog()|Out-Null;Log "ShowDialog returned"
    }catch{
        Log "POPUP FATAL ERROR: $_`nStack: $($_.ScriptStackTrace)"
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show("Error: $_`n`n$($_.ScriptStackTrace)","Debug Error")
    }
    exit
}

function ConvertTo-ArabicNumerals([int]$n) {
    $map=@{'0'=[char]0x0660;'1'=[char]0x0661;'2'=[char]0x0662;'3'=[char]0x0663;'4'=[char]0x0664;'5'=[char]0x0665;'6'=[char]0x0666;'7'=[char]0x0667;'8'=[char]0x0668;'9'=[char]0x0669}
    $s=[string]$n; foreach($k in $map.Keys){$s=$s.Replace($k,[string]$map[$k])}; return $s
}

function Get-CurrentMode($settings,$progress) {
    switch($settings.mode){
        'onlyKhatma'{return 'khatma'}
        'onlyRandom'{return 'random'}
        'equalBoth' {if($progress.PSObject.Properties['nextMode']){return $progress.nextMode}else{return 'khatma'}}
        'percentage'{
            $last=if($progress.PSObject.Properties['lastWasRandom']){[bool]$progress.lastWasRandom}else{$false}
            if($last){return 'khatma'}
            if((Get-Random -Min 0 -Max 100) -lt $settings.percentage){return 'random'}else{return 'khatma'}
        }
        default{return 'khatma'}
    }
}

# =============================================================================
# MODE: NOTIFY (-hidden)
# =============================================================================
if ($args -contains '-hidden') {
    Log "MODE: -hidden"
    try {
        if(-not(Test-Path $quranPath)){Log "ERROR: quran.json NOT FOUND";exit}
        $quranData=Get-Content $quranPath -Encoding UTF8|ConvertFrom-Json
        Log "quranData: $($quranData.Count) surahs"
        $settings=Get-Settings;$vc=[int]$settings.verseCount;$fullSurah=[bool]$settings.fullSurah
        Log "mode=$($settings.mode) verses=$vc fullSurah=$fullSurah reciter=$($settings.reciter)"

        $mp3Folder=$null;$hasAudio=$false
        if($settings.reciter -ne "" -and $settings.reciter -ne "none"){
            $rd2=Get-Reciters
            $ro2=$rd2|Where-Object{$_.url -eq $settings.reciter}|Select-Object -First 1
            if($ro2 -ne $null -and $ro2.status -eq 'full'){
                $sn2=$ro2.name -replace '[\\/:*?"<>|]','_'
                $cand=Join-Path $audioBase $sn2
                if(Test-Path $cand){$mp3Folder=$cand;$hasAudio=$true;Log "mp3Folder=$mp3Folder"}
                else{Log "Audio folder missing: $cand"}
            }else{Log "Reciter not downloaded: $($settings.reciter)"}
        }else{Log "No audio (none)"}

        $scriptPath=$PSCommandPath;$timestamp=Get-Date -Format 'yyyyMMdd_HHmmss'

        if(Test-Path $khatmaFile){
            $raw2=Get-Content $khatmaFile -Encoding UTF8|ConvertFrom-Json
            $progress=[PSCustomObject]@{
                surahIndex=[int]$raw2.surahIndex;verseIndex=[int]$raw2.verseIndex
                nextMode=if($raw2.PSObject.Properties['nextMode']){[string]$raw2.nextMode}else{'khatma'}
                lastWasRandom=if($raw2.PSObject.Properties['lastWasRandom']){[bool]$raw2.lastWasRandom}else{$false}
            }
        }else{$progress=[PSCustomObject]@{surahIndex=0;verseIndex=0;nextMode="khatma";lastWasRandom=$false}}

        $currentMode=Get-CurrentMode $settings $progress
        Log "currentMode=$currentMode"
        $progress.lastWasRandom=($currentMode -eq 'random')
        if($settings.mode -eq 'equalBoth'){$progress.nextMode=if($currentMode -eq 'khatma'){'random'}else{'khatma'}}
        $progress|ConvertTo-Json|Out-File -FilePath $khatmaFile -Encoding UTF8
        $isKhatma=($currentMode -eq 'khatma')

        function Build-Mp3List($verses,$folderPath){
            $files=@()
            if($folderPath){
                $bas=Join-Path $folderPath "001001.mp3";if(Test-Path $bas){$files+=$bas}
                foreach($v2 in $verses){$p2=Join-Path $folderPath "$($v2.surahNum.ToString('000'))$($v2.id.ToString('000')).mp3";if(Test-Path $p2){$files+=$p2}}
            }
            return $files
        }
        function Write-SurahFile($so,$ts){
            $f="$env:TEMP\quran_surah_$ts.txt"
            "Surah $($so.name)`n`n$(($so.verses|ForEach-Object{"$($_.text) [$($_.id)]"})-join"`n`n")"|Out-File -FilePath $f -Encoding UTF8
            return $f
        }

        if($isKhatma){
            $si2=[int]$progress.surahIndex;$vi2=[int]$progress.verseIndex
            if($fullSurah){
                $so=$quranData[$si2]
                $verses=$so.verses|ForEach-Object{[PSCustomObject]@{text=$_.text;id=$_.id;surahName=$so.name;surahNum=$si2+1}}
                $headerTitle="ختمة  ·  $($so.name)  (السورة كاملة)"
                $verseText=($verses|ForEach-Object{"$($_.text) ﴿$($_.id)﴾ "})-join" "
                $mp3Files=Build-Mp3List $verses $mp3Folder;$surahFile=Write-SurahFile $so $timestamp
            }else{
                $verses2=[System.Collections.Generic.List[object]]::new();$csi=$si2;$cvi=$vi2
                while($verses2.Count -lt $vc -and $csi -lt $quranData.Count){
                    $su=$quranData[$csi]
                    while($cvi -lt $su.verses.Count -and $verses2.Count -lt $vc){$v3=$su.verses[$cvi];$verses2.Add([PSCustomObject]@{text=$v3.text;id=$v3.id;surahName=$su.name;surahNum=$csi+1});$cvi++}
                    if($verses2.Count -lt $vc){$csi++;$cvi=0}
                }
                $verses=$verses2
                $headerTitle="ختمة  ·  $(($verses|Select-Object -ExpandProperty surahName -Unique)-join' / ')  (آية $($verses[-1].id)-$($verses[0].id))"
                $verseText=($verses|ForEach-Object{"$($_.text) ﴿$($_.id)﴾ "})-join" "
                $mp3Files=Build-Mp3List $verses $mp3Folder;$surahFile=Write-SurahFile $quranData[$si2] $timestamp
            }
        }else{
            if($fullSurah){
                $sui=Get-Random -Min 0 -Max $quranData.Count;$so=$quranData[$sui]
                $verses=$so.verses|ForEach-Object{[PSCustomObject]@{text=$_.text;id=$_.id;surahNum=$sui+1}}
                $headerTitle="عشوائي  ·  $($so.name)  (السورة كاملة)"
                $verseText=($verses|ForEach-Object{"$($_.text) ﴿$($_.id)﴾ "})-join" "
                $mp3Files=Build-Mp3List $verses $mp3Folder;$surahFile=Write-SurahFile $so $timestamp
            }else{
                $elig=[System.Collections.Generic.List[object]]::new()
                for($si3=0;$si3 -lt $quranData.Count;$si3++){$vc3=$quranData[$si3].verses.Count;for($vi3=0;$vi3 -le($vc3-$vc);$vi3++){$elig.Add([PSCustomObject]@{SurahIndex=$si3;VerseIndex=$vi3})}}
                $pick=$elig[(Get-Random -Min 0 -Max $elig.Count)];$sui=$pick.SurahIndex;$svi=$pick.VerseIndex
                $su2=$quranData[$sui]
                $verses=$su2.verses[$svi..($svi+$vc-1)]|ForEach-Object{[PSCustomObject]@{text=$_.text;id=$_.id;surahNum=$sui+1}}
                $headerTitle="عشوائي  ·  $($su2.name)  (آية $($verses[-1].id)-$($verses[0].id))"
                $verseText=($verses|ForEach-Object{"$($_.text) ﴿$($_.id)﴾ "})-join" "
                $mp3Files=Build-Mp3List $verses $mp3Folder;$surahFile=Write-SurahFile $quranData[$sui] $timestamp
            }
        }

        $charCount=($verseText -replace '\s','').Length
        $maxSecs=if($fullSurah){600}else{90}
        $baseSeconds=[Math]::Max(15,[Math]::Min($maxSecs,[int]($charCount/15)))
        $readSeconds=[Math]::Max(10,[int]($baseSeconds*$settings.timerMultiplier))
        Log "chars=$charCount base=${baseSeconds}s mult=$($settings.timerMultiplier) read=${readSeconds}s"

        $audioList=if($mp3Files.Count -gt 0){$mp3Files -join ';'}else{""}
        Log "hasAudio=$hasAudio  audioList length=$($audioList.Length)"

        $payload=@{scriptPath=$scriptPath;headerTitle=$headerTitle;verseText=$verseText;audioList=$audioList;surahFile=$surahFile;isKhatma=$isKhatma;readSeconds=$readSeconds;hasAudio=$hasAudio}|ConvertTo-Json -Compress
        $payloadFile="$env:TEMP\quran_payload_$timestamp.json"
        $payload|Out-File -FilePath $payloadFile -Encoding UTF8
        Log "Payload -> $payloadFile"
        Start-Process powershell.exe -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -popupfile `"$payloadFile`""
        Log "-hidden done"
    }catch{Log "HIDDEN ERROR: $_`nStack: $($_.ScriptStackTrace)"}
    exit
}

# =============================================================================
# LAUNCHER
# =============================================================================
Log "LAUNCHER mode"
$scriptPath=$MyInvocation.MyCommand.Path
Start-Process powershell.exe -WindowStyle Hidden -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -hidden"
Log "Launcher done"