############################################################################################################
# Window Resizer Ver 1.0.0                                                                                 #
############################################################################################################
# -------------------------------------------------------------------------------------------------------- #
# 変数
#   $ScriptPath               :   自分自身のパス
#   $AssemblyLocation         :   各種アセンブリの場所
#   $wh                       :   ウィンドウハンドラー
#   $WindowTitle              :   ウィンドウタイトル
#   $Width                    :   ウィンドウの幅
#   $Height                   :   ウィンドウの高さ
#   $rect                     :   ウィンドウの位置情報やサイズ情報が入る構造体
#   $xaml                     :   GUIを構成するxaml変数
#   $reader                   :   Xaml Reader
#   $window                   :   メインウィンドウオブジェクト
#   $selectedWindowTitle      :   コンボボックスで選択されたアイテムと一致するウィンドウタイトルが入る
#   $selectedProcess          :   選択されたアイテムのプロセスの詳細が入る。ウィンドウハンドラを取得する為に必要
# -------------------------------------------------------------------------------------------------------- #
# Info
#   Development Environment   : Powershell 7
#   Developer                 : saica
#   Used Library              : PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing, MahApps, materialDesign, ControlzEx, Micorosoft.Xaml.Behaviors
#   System Requirements       : Windows 10 以降 && Need [Powershell 7(.Net Core)]
# -------------------------------------------------------------------------------------------------------- #

# スクリプト本体があるディレクトリの取得
if($MyInvocation.MyCommand.CommandType -eq "ExternalScript"){
  $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
} else {
  $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
  if(!$ScriptPath){
    $ScriptPath = "."
  }
}

# アセンブリを読み込む
$AssemblyLocation = Join-Path -Path $ScriptPath -ChildPath assembly
foreach ($Assembly in (Get-ChildItem $AssemblyLocation -Filter *.dll)) {
  [System.Reflection.Assembly]::LoadFrom($Assembly.FullName) | out-null
}

Set-ExecutionPolicy RemoteSigned -Scope Process -Force

Add-Type @"
  using System;
  using System.Runtime.InteropServices;

  public class User32 {
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, [MarshalAs(UnmanagedType.Bool)] bool bRepaint);
  }

  public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
  }
"@

function Resize-Window {
  param(
    [IntPtr]$wh,
    [string]$WindowTitle,
    [int]$Width = -1,
    [int]$Height = -1
  )

  if ($wh -ne [IntPtr]::Zero) {
    $rect = New-Object RECT
    [User32]::GetWindowRect($wh, [ref]$rect)
    if($Width -eq -1){
      $Width = $rect.Right - $rect.Left
    }
    if($Height -eq -1){
      $Height = $rect.Bottom - $rect.Top
    }
    [User32]::MoveWindow($wh, $rect.Left, $rect.Top, $Width, $Height, $true)
  }
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

$xaml = @'
<Controls:MetroWindow
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:iconPacks="http://metro.mahapps.com/winfx/xaml/iconpacks"
  xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
  xmlns:metro="http://metro.mahapps.com/winfx/xaml/controls"
  GlowBrush="{DynamicResource PrimaryHueMidBrush}"
  BorderThickness="1"
  TitleCharacterCasing="Normal"
  Title="Window Resizer"
  Width="500"
  Height="240"
  ShowMaxRestoreButton="False"
  ShowMinButton="False"
  IsMaxRestoreButtonEnabled="False"
  IsMinButtonEnabled="False"
  ShowIconOnTitleBar="False"
  WindowStartupLocation="CenterScreen"
  ResizeMode="NoResize" >

  <Window.Resources>
    <ResourceDictionary>
      <ResourceDictionary.MergedDictionaries>
        <!-- Material Design -->
        <ResourceDictionary Source="pack://application:,,,/MaterialDesignThemes.Wpf;component/Themes/MaterialDesignTheme.Dark.xaml" />
        <ResourceDictionary Source="pack://application:,,,/MaterialDesignThemes.Wpf;component/Themes/MaterialDesignTheme.Defaults.xaml" />
        <ResourceDictionary Source="pack://application:,,,/MaterialDesignColors;component/Themes/Recommended/Primary/MaterialDesignColor.BlueGrey.xaml" />
        <ResourceDictionary Source="pack://application:,,,/MaterialDesignColors;component/Themes/Recommended/Accent/MaterialDesignColor.Blue.xaml" />

        <!-- MahApps.Metro resource dictionaries. Make sure that all file names are Case Sensitive! -->
        <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
        <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
        <!-- Accent and AppTheme setting -->
        <!-- “Light”, “Dark” -->
        <!--“Red”, “Green”, “Blue”, “Purple”, “Orange”, “Lime”, “Emerald”, “Teal”, “Cyan”, “Cobalt”, “Indigo”, “Violet”, “Pink”, “Magenta”, “Crimson”, “Amber”, “Yellow”, “Brown”, “Olive”, “Steel”, “Mauve”, “Taupe”, “Sienna” -->
        <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Cobalt.xaml" />

        <!-- Material Design: MahApps Compatibility -->
        <ResourceDictionary Source="pack://application:,,,/MaterialDesignThemes.MahApps;component/Themes/MaterialDesignTheme.MahApps.Fonts.xaml" />
        <ResourceDictionary Source="pack://application:,,,/MaterialDesignThemes.MahApps;component/Themes/MaterialDesignTheme.MahApps.Flyout.xaml" />
      </ResourceDictionary.MergedDictionaries>

      <!-- Primary -->
      <SolidColorBrush x:Key="PrimaryHueLightBrush" Color="#8493a2" />
      <SolidColorBrush x:Key="PrimaryHueLightForeGroundBrush" Color="#ffffff" />
      <SolidColorBrush x:Key="PrimaryHueMidBrush" Color="#576573" />
      <SolidColorBrush x:Key="PrimaryHueMidForeGroundBrush" Color="#ffffff" />
      <SolidColorBrush x:Key="PrimaryHueDarkBrush" Color="#2d3b48" />
      <SolidColorBrush x:Key="PrimaryHueDarkForeGroundBrush" Color="#ffffff" />
    </ResourceDictionary>
  </Window.Resources>

  <Grid>

    <Grid.RowDefinitions>
      <RowDefinition Height="*" />
      <RowDefinition Height="*" />
      <RowDefinition Height="*" />
      <RowDefinition Height="*" />
      <RowDefinition Height="*" />
      <RowDefinition Height="*" />
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*" />
      <ColumnDefinition Width="Auto" />
    </Grid.ColumnDefinitions>

    <Menu Grid.Row="0" Width="Auto" Height="Auto" HorizontalAlignment="Left" VerticalAlignment="Stretch">
      <MenuItem Header="UI Size" Width="Auto" Height="Auto" Name="UI_Size" >
        <MenuItem Header="FHD" Name="FHD" IsCheckable="True" IsChecked="True" />
        <MenuItem Header="4K" Name="FK" IsCheckable="True" />
      </MenuItem>
      <MenuItem Header="プロセスの再取得" Width="Auto" Height="Auto" Name="Get_p" />
    </Menu>

    <ComboBox Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2" Width="Auto" Height="Auto" Name="ComboBox1"  Style="{DynamicResource MahApps.Styles.ComboBox.Virtualized}" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="5,5,5,0" />
    <Viewbox Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="2" Width="Auto" Height="Auto" Name="WVB" Stretch="Uniform" HorizontalAlignment="Left">
      <Label Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="2" Width="Auto" Height="Auto" Name="Width" Content="Width(横幅)" Margin="5,0,5,0" />
    </Viewbox>

    <Slider Grid.Row="3" Grid.Column="0" Name="WSlider" HorizontalAlignment="Stretch" VerticalAlignment="Center"  Orientation="Horizontal" AutoToolTipPlacement="TopLeft" LargeChange="10" Maximum="3840" Minimum="10" SmallChange="10" Value="0" IsSnapToTickEnabled="True" TickFrequency="10" RenderTransformOrigin="1.0,1.0" Margin="5,5,5,0">
      <Slider.RenderTransform>
        <ScaleTransform ScaleX="1" ScaleY="1" />
      </Slider.RenderTransform>
    </Slider>

    <Controls:NumericUpDown Name="nUpDown1" Grid.Row="3" Grid.Column="2" Maximum="3850" Minimum="1"  Value="{Binding ElementName=WSlider, Path=Value}" Margin="0,0,5,0" />
    
    <Viewbox Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Width="Auto" Height="Auto" Name="HVB" Stretch="Uniform" HorizontalAlignment="Left">
      <Label Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Width="Auto" Height="Auto" Name="Height" Content="Height(高さ)" Margin="5,0,5,0" />
    </Viewbox>

    <Slider Grid.Row="5" Grid.Column="0" Name="HSlider" HorizontalAlignment="Stretch" VerticalAlignment="Center" Orientation="Horizontal" AutoToolTipPlacement="TopLeft" LargeChange="10" Maximum="2160" Minimum="10" SmallChange="10" Value="0" IsSnapToTickEnabled="True" TickFrequency="10" RenderTransformOrigin="1.0,1.0" Margin="5,5,5,5">
      <Slider.RenderTransform>
        <ScaleTransform ScaleX="1" ScaleY="1" />
      </Slider.RenderTransform>
    </Slider>

    <Controls:NumericUpDown Name="nUpDown2" Grid.Row="5" Grid.Column="2"  Maximum="2160" Minimum="1" Value="{Binding ElementName=HSlider, Path=Value}" Margin="0,0,5,5" />

  </Grid>
</Controls:MetroWindow>
'@
# Margin = 左,上,右,下

try{
  $reader = [XML.XMLReader]::Create([System.IO.StringReader]$xaml)
  $window = [Windows.Markup.XAMLReader]::Load($reader)

  # 各コントロールの追加
  # メンバ登録することにより $変数名.メンバ名でコントロールを操作できる
  $window | Add-Member NoteProperty -Name "UI_Size" -Value $window.FindName("UI_Size") -Force
  $window | Add-Member NoteProperty -Name "FHD" -Value $window.FindName("FHD") -Force
  $window | Add-Member NoteProperty -Name "FK" -Value $window.FindName("FK") -Force
  $window | Add-Member NoteProperty -Name "Get_p" -Value $window.FindName("Get_p") -Force
  $window | Add-Member NoteProperty -Name "ComboBox" -Value $window.FindName("ComboBox1") -Force
  $window | Add-Member NoteProperty -Name "LWidth" -Value $window.FindName("Width") -Force
  $window | Add-Member NoteProperty -Name "WSlider" -Value $window.FindName("WSlider") -Force
  $window | Add-Member NoteProperty -Name "nUpDown1" -Value $window.FindName("nUpDown1") -Force
  $window | Add-Member NoteProperty -Name "LHeight" -Value $window.FindName("Height") -Force
  $window | Add-Member NoteProperty -Name "HSlider" -Value $window.FindName("HSlider") -Force
  $window | Add-Member NoteProperty -Name "nUpDown2" -Value $window.FindName("nUpDown2") -Force
}catch{
    Write-Error "Error bulding Xaml Data ,`n$_"
    exit
}

# ウィンドウタイトルの取得
function Get_ProcessList {
  $window.ComboBox.ItemsSource = Get-Process | Where-Object {$_.MainWindowTitle -ne ""} | Select-Object -ExpandProperty MainWindowTitle
}


# アイテムの変更イベント
$window.ComboBox.add_SelectionChanged({
  $selectedWindowTitle = $window.ComboBox.SelectedItem
  if ($selectedWindowTitle -ne $null) {
    $selectedProcess = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -eq $selectedWindowTitle }
    if ($selectedProcess -ne $null) {
      # スライダーの値変更イベントを一時的に削除
      # 理由:ドロップダウンからアイテムを選択したときにスライダーの値変更イベントも呼ばれてしまうため
      $window.WSlider.remove_ValueChanged($WSlider_ValueChanged)
      $window.HSlider.remove_ValueChanged($HSlider_ValueChanged)

      $rect = New-Object RECT
      [User32]::GetWindowRect($selectedProcess.MainWindowHandle, [ref]$rect)
      $window.WSlider.Value = $rect.Right - $rect.Left
      $window.HSlider.Value = $rect.Bottom - $rect.Top

      # 削除したイベントを復元
      $window.WSlider.add_ValueChanged($WSlider_ValueChanged)
      $window.HSlider.add_ValueChanged($HSlider_ValueChanged)
    }
  }
})

# スライダーの値変更イベント
$WSlider_ValueChanged = {
  $selectedProcess = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -eq $window.ComboBox.Text }
  if ($selectedProcess -ne $null) {
    Resize-Window -wh $selectedProcess.MainWindowHandle -WindowTitle $selectedProcess.MainWindowTitle -Width $window.WSlider.Value
  }
}

$HSlider_ValueChanged = {
  $selectedProcess = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -eq $window.ComboBox.Text }
  if ($selectedProcess -ne $null) {
    Resize-Window -wh $selectedProcess.MainWindowHandle -WindowTitle $selectedProcess.MainWindowTitle -Height $window.HSlider.Value
  }
}

$window.add_Closed({
  $window = $null

})

# メニューアイテム [UI Size] -> [FHD] が選択された時
$window.FHD.add_Checked({
  $window.ResizeMode = "CanResize"
  $window.Width = 500
  $window.Height = 240
  $window.ResizeMode = "NoResize"
  $window.FK.IsChecked=$false
  $window.FHD.IsChecked=$true
  $window.ComboBox.FontSize = 12
  $window.UI_Size.FontSize = 14
  $window.FHD.FontSize = 14
  $window.FK.FontSize = 14
  $window.Get_p.FontSize = 14
  $window.nUpDown1.FontSize = 12
  $window.nUpDown2.FontSize = 12
  $window.WSlider.RenderTransform.ScaleY = 1
  $window.HSlider.RenderTransform.ScaleY = 1
})

# メニューアイテム [UI Size] -> [4K] が選択された時
$window.FK.add_Checked({
  $window.ResizeMode = "CanResize"
  $window.Width = 1000
  $window.Height = 480
  $window.ResizeMode = "NoResize"
  $window.FK.IsChecked=$true
  $window.FHD.IsChecked=$false
  $window.ComboBox.FontSize *= 2
  $window.UI_Size.FontSize *= 2
  $window.FHD.FontSize *= 2
  $window.FK.FontSize *= 2
  $window.Get_p.FontSize *= 2
  $window.nUpDown1.FontSize *= 2
  $window.nUpDown2.FontSize *= 2
  $window.WSlider.RenderTransform.ScaleY = 2
  $window.HSlider.RenderTransform.ScaleY = 2
})

$window.Get_p.add_Click({
  Get_ProcessList
})

Get_ProcessList

$window.ShowDialog() | Out-Null