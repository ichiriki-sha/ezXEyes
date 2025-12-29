<#
.SYNOPSIS
    「xeyes」風のマスコットアプリ。デスクトップ上に目を表示し、マウスカーソルに合わせて瞳が追従します。

.DESCRIPTION
    - デスクトップ上に左右の目を表示し、瞳がマウスカーソルを追従します
    - Canvas と WPF を使用して描画
    - ウィンドウの位置とサイズは ini ファイルに保存し、次回起動時に復元
    - タスクトレイに常駐し、右クリックメニューから終了可能
    - 同時起動防止のため Mutex を使用
    - ウィンドウは透明化可能で、必要に応じて最前面に表示

.NOTES
    作成者: ichiriki
    作成日: 2025/12/28
    PowerShell Version: 5.1 推奨
    注意点:
        - Canvas サイズ変更時は目のサイズを自動調整
        - ini ファイルは %APPDATA%\ezxEyes\ に保存

.EXAMPLE
    .\ezxEyes.ps1
    - デスクトップ上に目を表示し、タスクトレイに常駐します

#>

# -----------------------------
# アセンブリ読み込み
# -----------------------------
# WPF や Windows Forms、GDI+ を使用するために必要な .NET アセンブリをロード
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -----------------------------
# 単一起動チェック
# -----------------------------
# 同じスクリプトが複数起動されるのを防止するため、Mutex を使用
$scriptName = [System.IO.Path]::GetFileName($Myinvocation.MyCommand.Path)
$createdNew = $false
$mutexName = "Local\${scriptName}_single_instance_mutex"
$global:mutex = New-Object System.Threading.Mutex($true, $mutexName)

if (-not $global:mutex.WaitOne(0, $false)) {

    exit
}

# -----------------------------
# 設定ファイル・初期値
# -----------------------------
# アプリケーション用データ保存ディレクトリ
$script:ScriptName = [System.IO.Path]::GetFileName($MyInvocation.MyCommand.Path)
$script:ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($script:ScriptName)
$appDataDir = Join-Path -Path $env:APPDATA -ChildPath $script:ScriptBaseName

Write-Host "appDataDir:$appDataDir"

# ディレクトリがなければ作成
if (-not (Test-Path $appDataDir)) {
    New-Item -ItemType Directory -Path $appDataDir | Out-Null
}

# ini ファイルパス
$script:IniFile = Join-Path -Path $appDataDir -ChildPath "$($script:ScriptBaseName).ini"

# ウィンドウサイズ関連初期値
$script:CanvasWidth  = 220
$script:CanvasHeight = 150

$script:WindowPaddingWidth  = 16 + 8
$script:WindowPaddingHeight = 39 + 8

$script:DefaultLeft   = 100
$script:DefaultTop    = 100
$script:DefaultWidth  = $script:CanvasWidth  + $script:WindowPaddingWidth
$script:DefaultHeight = $script:CanvasHeight + $script:WindowPaddingHeight

# -----------------------------
# ウィンドウ設定の読み書き
# -----------------------------
function Read-WindowSettings {
    param($path)
    if (-Not (Test-Path $path)) {
        return @{Left=$script:DefaultLeft; Top=$script:DefaultTop; Width=$script:DefaultWidth; Height=$script:DefaultHeight}
    }

    $lines = Get-Content $path
    $settings = @{}
    foreach($line in $lines){
        if ($line -match '^(Left|Top|Width|Height)=(\d+)$'){
            $settings[$matches[1]] = [int]$matches[2]
        }
    }

    foreach($key in @('Left','Top','Width','Height')){
        if (-Not $settings.ContainsKey($key)){
            $settings[$key] = (Get-Variable "default$key").Value
        }
    }

    return $settings
}

function Save-WindowSettings {
    param($path,$left,$top,$width,$height)
    @(
        "Left=$left"
        "Top=$top"
        "Width=$width"
        "Height=$height"
    ) | Set-Content -Path $path -Encoding UTF8
}

# -----------------------------
# 目のレイアウト計算関数
# -----------------------------
function Calc-EyeLayout {
    param($canvasWidth, $canvasHeight)

    $centerX    = $canvasWidth / 2
    $centerY    = $canvasHeight / 2
    $eyeWidth   = [Math]::Max(20, $canvasWidth * 0.4)   # 最小サイズガード
    $eyeHeight  = [Math]::Max(12, $canvasHeight * 0.8)
    $pupilRatio = 0.35
    $spacing    = $eyeWidth * 0.15
    $leftX      = $centerX - ($eyeWidth / 2) - ($spacing / 2)
    $rightX     = $centerX + ($eyeWidth / 2) + ($spacing / 2)

    return @{
        LeftX = $leftX; RightX = $rightX; CenterY = $centerY; EyeWidth = $eyeWidth; EyeHeight = $eyeHeight; PupilRatio = $pupilRatio
    }
}

# -----------------------------
# トレイアイコン作成
# -----------------------------
function Create-TrayIcon {
    $bitmap                 = New-Object System.Drawing.Bitmap 16, 16
    $graphics               = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $graphics.Clear([System.Drawing.Color]::Transparent)

    # 目描画
    $eyeWidth         = 6
    $eyeHeight        = 8
    $pupilRatio       = 0.35
    $leftX            = 5
    $leftY            = 8
    $rightX           = 11
    $rightY           = 8

    $graphics.FillEllipse([System.Drawing.Brushes]::White, $leftX  - $eyeWidth / 2, $leftY  - $eyeHeight * 0.9, $eyeWidth, $eyeHeight)
    $graphics.DrawEllipse([System.Drawing.Pens]::Black   , $leftX  - $eyeWidth / 2, $leftY  - $eyeHeight * 0.9, $eyeWidth, $eyeHeight)
    $graphics.FillEllipse([System.Drawing.Brushes]::White, $rightX - $eyeWidth / 2, $rightY - $eyeHeight * 0.9, $eyeWidth, $eyeHeight)
    $graphics.DrawEllipse([System.Drawing.Pens]::Black   , $rightX - $eyeWidth / 2, $rightY - $eyeHeight * 0.9, $eyeWidth, $eyeHeight)

    $pupilOffsetX     =  2
    $pupilOffsetY     = -4
    $pupilW           = $eyeWidth  * $pupilRatio
    $pupilH           = $eyeHeight * $pupilRatio
    $graphics.FillEllipse([System.Drawing.Brushes]::Black, $leftX  - $pupilW / 2 + $pupilOffsetX, $leftY  - $pupilH / 2 + $pupilOffsetY, $pupilW, $pupilH)
    $graphics.FillEllipse([System.Drawing.Brushes]::Black, $rightX - $pupilW / 2 + $pupilOffsetX, $rightY - $pupilH / 2 + $pupilOffsetY, $pupilW, $pupilH)

    $hIcon            = $bitmap.GetHicon()
    $icon             = [System.Drawing.Icon]::FromHandle($hIcon)

    $trayIcon         = New-Object System.Windows.Forms.NotifyIcon
    $trayIcon.Icon    = $icon
    $trayIcon.Visible = $true
    $trayIcon.Text    = $script:ScriptBaseName

    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $exitItem = $contextMenu.Items.Add("終了(&X)")
    $exitItem.Add_Click({
        $trayIcon.Visible = $false
        $trayIcon.Dispose()
        $Window.Close()
    })
    $trayIcon.ContextMenuStrip = $contextMenu

    $trayIcon.Add_DoubleClick({
        if ($Window.WindowState -eq 'Minimized') { $Window.WindowState = 'Normal' }
        $Window.Topmost = $true
        $Window.Activate()
        $Window.Topmost = $false
    })

    # GDI リソース破棄
    $graphics.Dispose()
    $bitmap.Dispose()

    return $trayIcon
}

# -----------------------------
# XAML ウィンドウ作成
# -----------------------------
$script:WinSettings = Read-WindowSettings $script:IniFile

[xml]$xaml = @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    Left='$($script:WinSettings.Left)'
    Top='$($script:WinSettings.Top)'
    Width='$($script:WinSettings.Width)'
    Height='$($script:WinSettings.Height)'
    WindowStyle='None'
    ResizeMode='CanResizeWithGrip'
    AllowsTransparency='True'
    Background='Transparent'
    ShowInTaskbar='False'>

    <Grid>
        <Canvas x:Name='MainCanvas' Background='Transparent'/>
        <ResizeGrip HorizontalAlignment='Right' VerticalAlignment='Bottom' Width='16' Height='16' Margin='0'/>
    </Grid>
</Window>
"@

$reader        = New-Object System.Xml.XmlNodeReader $xaml
$script:Window = [Windows.Markup.XamlReader]::Load($reader)
$script:Canvas = $script:Window.FindName("MainCanvas")
$script:Eyes   = New-Object System.Collections.ArrayList

$script:Canvas.Add_MouseLeftButtonDown({ $script:Window.DragMove() })

# -----------------------------
# 目作成関数
# -----------------------------
function Create-Eye {
    param($centerX, $centerY, $eyeWidth, $eyeHeight, $pupilRatio)

    $eye = New-Object System.Windows.Shapes.Ellipse
    $eye.Width  = $eyeWidth
    $eye.Height = $eyeHeight
    $eye.Fill   = [System.Windows.Media.Brushes]::White
    $eye.Stroke = [System.Windows.Media.Brushes]::Black
    $eye.StrokeThickness = 2
    [System.Windows.Controls.Canvas]::SetLeft($eye , $centerX - $eyeWidth  / 2)
    [System.Windows.Controls.Canvas]::SetTop($eye , $centerY - $eyeHeight / 2)

    $pupil = New-Object System.Windows.Shapes.Ellipse
    $pupil.Width  = $eyeWidth * $pupilRatio
    $pupil.Height = $eyeHeight * $pupilRatio
    $pupil.Fill   = [System.Windows.Media.Brushes]::Black
    [System.Windows.Controls.Canvas]::SetLeft($pupil, $centerX - $pupil.Width  / 2)
    [System.Windows.Controls.Canvas]::SetTop($pupil, $centerY - $pupil.Height / 2)

    return [PSCustomObject]@{
        Eye     = $eye
        Pupil   = $pupil
        CenterX = $centerX
        CenterY = $centerY
        EyeW    = $eyeWidth
        EyeH    = $eyeHeight
    }
}

function Update-EyePositionAndSize {
    param($eyeData, $centerX, $centerY, $eyeWidth, $eyeHeight, $pupilRatio)
    $eyeData.Eye.Width  = $eyeWidth
    $eyeData.Eye.Height = $eyeHeight
    [System.Windows.Controls.Canvas]::SetLeft($eyeData.Eye, $centerX - $eyeWidth / 2)
    [System.Windows.Controls.Canvas]::SetTop($eyeData.Eye , $centerY - $eyeHeight / 2)

    $eyeData.Pupil.Width  = $eyeWidth * $pupilRatio
    $eyeData.Pupil.Height = $eyeHeight * $pupilRatio
    [System.Windows.Controls.Canvas]::SetLeft($eyeData.Pupil, $centerX - $eyeData.Pupil.Width / 2)
    [System.Windows.Controls.Canvas]::SetTop($eyeData.Pupil , $centerY - $eyeData.Pupil.Height / 2)

    $eyeData.CenterX = $centerX
    $eyeData.CenterY = $centerY
    $eyeData.EyeW    = $eyeWidth
    $eyeData.EyeH    = $eyeHeight
}

# -----------------------------
# 初期化
# -----------------------------
function Initialize-Eyes {
    $script:Canvas.Children.Clear()
    $script:Eyes.Clear()

    $cw = if ($script:Canvas.ActualWidth  -eq 0) { $script:Window.Width  } else { $script:Canvas.ActualWidth  }
    $ch = if ($script:Canvas.ActualHeight -eq 0) { $script:Window.Height } else { $script:Canvas.ActualHeight }

    $layout   = Calc-EyeLayout $cw $ch
    $leftEye  = Create-Eye $layout.LeftX  $layout.CenterY $layout.EyeWidth $layout.EyeHeight $layout.PupilRatio
    $rightEye = Create-Eye $layout.RightX $layout.CenterY $layout.EyeWidth $layout.EyeHeight $layout.PupilRatio

    $script:Canvas.Children.Add($leftEye.Eye)    | Out-Null
    $script:Canvas.Children.Add($leftEye.Pupil)  | Out-Null
    $script:Canvas.Children.Add($rightEye.Eye)   | Out-Null
    $script:Canvas.Children.Add($rightEye.Pupil) | Out-Null

    $script:Eyes.Add($leftEye)  | Out-Null
    $script:Eyes.Add($rightEye) | Out-Null
}

Initialize-Eyes

# -----------------------------
# リサイズ時更新
# -----------------------------
$script:Window.Add_SizeChanged({
    $cw = if ($script:Canvas.ActualWidth  -eq 0) { $script:Window.Width  } else { $script:Canvas.ActualWidth  }
    $ch = if ($script:Canvas.ActualHeight -eq 0) { $script:Window.Height } else { $script:Canvas.ActualHeight }

    $layout = Calc-EyeLayout $cw $ch
    Update-EyePositionAndSize $script:Eyes[0] $layout.LeftX  $layout.CenterY $layout.EyeWidth $layout.EyeHeight $layout.PupilRatio
    Update-EyePositionAndSize $script:Eyes[1] $layout.RightX $layout.CenterY $layout.EyeWidth $layout.EyeHeight $layout.PupilRatio
})

# -----------------------------
# タイマーで目追従
# -----------------------------
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(30)
$timer.Add_Tick({
    $cursor = [System.Windows.Forms.Cursor]::Position
    $windowRect = @{
        Left   = [int]$script:Window.Left
        Top    = [int]$script:Window.Top
        Right  = [int]($script:Window.Left + $script:Window.Width)
        Bottom = [int]($script:Window.Top  + $script:Window.Height)
    }

    $script:Window.Opacity = if (($cursor.X -ge $windowRect.Left -and $cursor.X -le $windowRect.Right) -and
                                 ($cursor.Y -ge $windowRect.Top  -and $cursor.Y -le $windowRect.Bottom)) { 1.0 } else { 0.6 }

    $formX = $cursor.X - $script:Window.Left
    $formY = $cursor.Y - $script:Window.Top

    foreach($eye in $script:Eyes) {
        $dx = $formX - $eye.CenterX
        $dy = $formY - $eye.CenterY
        if ($dx -ne 0 -or $dy -ne 0) {
            $angle = [Math]::Atan2($dy, $dx)
            $radiusX = ($eye.EyeW - $eye.Pupil.Width ) / 2
            $radiusY = ($eye.EyeH - $eye.Pupil.Height) / 2
            $px = $eye.CenterX + [Math]::Cos($angle) * $radiusX - $eye.Pupil.Width / 2
            $py = $eye.CenterY + [Math]::Sin($angle) * $radiusY - $eye.Pupil.Height / 2
            [System.Windows.Controls.Canvas]::SetLeft($eye.Pupil, $px)
            [System.Windows.Controls.Canvas]::SetTop($eye.Pupil , $py)
        }
    }
})
$timer.Start()

# -----------------------------
# ウィンドウ閉じたとき後片付け
# -----------------------------
$script:Window.Add_Closed({

    if ($timer) { $timer.Stop() }            # タイマー停止
    if ($script:TrayIcon -and $script:TrayIcon.Visible) {
        $script:TrayIcon.Visible = $false
        $script:TrayIcon.Dispose()
    }

    Save-WindowSettings $script:IniFile `
        ([Math]::Round($script:Window.Left)) `
        ([Math]::Round($script:Window.Top)) `
        ([Math]::Round($script:Window.Width)) `
        ([Math]::Round($script:Window.Height))

    #  Mutex 解放
    if ($global:mutex) {
        try {
            $global:mutex.ReleaseMutex()
        } catch {}
        $global:mutex.Dispose()
        $global:mutex = $null
    }
})

# -----------------------------
# トレイアイコン作成
# -----------------------------
$script:TrayIcon = Create-TrayIcon

# -----------------------------
# ウィンドウ表示
# -----------------------------
$script:Window.ShowDialog() | Out-Null
