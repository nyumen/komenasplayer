﻿# コメナスPlayer
# https://github.com/nyumen/komenasplayer

$current_path = Split-Path $MyInvocation.MyCommand.Path
Set-Location -Path $current_path

# include config
. ".\komenasplayer_config.ps1"

# Thanks
# http://kamifuji.dyndns.org/PS-Support/


# アセンブリの読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$SWP_NOSIZE = 0x0001     # ウインドウの現在のサイズを保持する
$SWP_NOMOVE = 0x0002     # ウインドウの現在位置を保持する
$SWP_NOZORDER = 0x0004   # ウインドウリスト内での現在位置を保持する
$SWP_SHOWWINDOW = 0x0040 #ウインドウを表示する

$HWND_TOPMOST = -1       # ウインドウをウインドウリストの一番上に配置する
$HWND_NOTOPMOST = -2     # すべての最前面ウィンドウの後ろに挿入

$signature=@' 
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(
        IntPtr hWnd, 
        int nCmdShow
    );

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow( IntPtr hWnd );

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(
        IntPtr hWnd,            // ウィンドウのハンドル
        IntPtr hWndInsertAfter, // 配置順序のハンドル
        int    X,               // 横方向の位置
        int    Y,               // 縦方向の位置
        int    cx,              // 幅
        int    cy,              // 高さ
        UInt32 uFlags           // ウィンドウ位置のオプション
    );
'@

$Win32 = Add-Type -memberDefinition $signature -name "Win32ApiFunctions" -namespace Win32ApiFunctions -passThru

# C#のソースコードを変数に保存
# http://blog.livedoor.jp/morituri/archives/53399411.html
if(!('Win32Api.RECT' -as [type])) {
    Add-Type –TypeDefinition `
      @"
using System;
using System.Runtime.InteropServices;
 
namespace Win32Api {

  public struct RECT {
    public int left;
    public int top;
    public int right;
    public int bottom;
  }
 
  public class Helper {

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
   
    public static RECT GetForegroundWindowRect(IntPtr hwnd) {
      RECT rect = new RECT();
      GetWindowRect(hwnd, out rect);
     
      return rect;
    }
  }
}
"@
}

#https://stuncloud.wordpress.com/2014/11/19/powershell_turnoff_ime_automatically/
# IMEの状態を取得、変更するクラスを定義
if(!('PowerShell.IME' -as [type])) {
    Add-Type –TypeDefinition `
        @'
using System;
using System.Runtime.InteropServices;
namespace PowerShell
{
    public class IME {
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
        
        [DllImport("imm32.dll")]
        private static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);
        public static int GetState(IntPtr hwnd) {
            IntPtr imeHwnd = ImmGetDefaultIMEWnd(hwnd);
            return SendMessage(imeHwnd, 0x0283, 0x0005, 0);
        }
        public static void SetState(IntPtr hwnd, bool state) {
            IntPtr imeHwnd = ImmGetDefaultIMEWnd(hwnd);
            SendMessage(imeHwnd, 0x0283, 0x0006, state?1:0);
        }
    }
}
'@
}

# スクリーンショット
# https://www.it-swarm-ja.tech/ja/powershell/windows-powershell%E3%81%A7%E7%94%BB%E9%9D%A2%E3%82%AD%E3%83%A3%E3%83%97%E3%83%81%E3%83%A3%E3%82%92%E5%AE%9F%E8%A1%8C%E3%81%99%E3%82%8B%E3%81%AB%E3%81%AF%E3%81%A9%E3%81%86%E3%81%99%E3%82%8C%E3%81%B0%E3%82%88%E3%81%84%E3%81%A7%E3%81%99%E3%81%8B%EF%BC%9F/969655941/
function screenshot([Drawing.Rectangle]$bounds, $path) {
    $myEncoder = [System.Drawing.Imaging.Encoder]::Quality
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1) 
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($myEncoder, $screenshot_quality)
    $myImageCodecInfo = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()|where {$_.MimeType -eq 'image/jpeg'}
    $bmp = New-Object Drawing.Bitmap $bounds.width, $bounds.height
    $graphics = [Drawing.Graphics]::FromImage($bmp)
    $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)
    $bmp.Save($path, $myImageCodecInfo, $encoderParam)
    $graphics.Dispose()
    $bmp.Dispose()
}

function getDefaultPos() {
    sleep -Milliseconds  100

    # $rectにPC TV Plusの4隅のスクリーン座標が格納される
    $ps_vnt = getProcess ( "Vnt" )
    $rect = [Win32Api.Helper]::GetForegroundWindowRect( $ps_vnt.MainWindowHandle )
    $VntX = $rect.Left + 20
    $VntY = $rect.Bottom - 180
}

function Send-Keys( $KeyStroke, $ProcessName ) {
    $ps = getProcess( $ProcessName )
    # IMEがオンだったらオフにする
    if ( $ps.Name -eq $ProcessName ) {
        if ([PowerShell.IME]::GetState( $ps.MainWindowHandle )) {
            [PowerShell.IME]::SetState( $ps.MainWindowHandle, $false )
            sleep -Milliseconds  100
        }
    }

#Write-Host $ps.Name
    [Void]$win32::SetForegroundWindow( $ps.MainWindowHandle )
    sleep -Milliseconds  100

    [System.Windows.Forms.SendKeys]::SendWait($KeyStroke)
}

function getProcess( $ProcessName ) {
    $ps = Get-Process | Where-Object {$_.Name -eq $ProcessName}
    if ( $ProcessName -eq "Vnt" ) {
        $script:ps_vnt = $ps
    } else {
        $script:ps_viewer = $ps
    }
    return $ps
}

$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height

$comment_viewer_file_name_pos = $comment_viewer_exe_path.LastIndexOf( "\" )
$comment_viewer_exe_name = $comment_viewer_exe_path.Substring( $comment_viewer_file_name_pos + 1 )
$comment_viewer_install_path = $comment_viewer_exe_path.Substring( 0, $comment_viewer_file_name_pos )
$comment_viewer_app = $comment_viewer_exe_name.Substring(0, $comment_viewer_exe_name.LastIndexOf( "." ))


# フォームの作成
$form = New-Object System.Windows.Forms.Form 
$form.Text = "コメナス"
$form.Size = New-Object System.Drawing.Size(230,190)
# 画面解像度に合わせて位置を調整する
$real_window_pos_top = $default_window_pos_top / ( 1080 + 36 ) * ( $screenHeight + 36 )
if ( $screenWidth -lt 1920 ) {
    $real_window_pos_left = 0
} else {
    $real_window_pos_left = $default_window_pos_left / 1920 * $screenWidth
}
$form.Top = $real_window_pos_top
$form.Left = $real_window_pos_left
$form.StartPosition = "Manual"
$form.BackColor = "#606060"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.FormBorderStyle = "FixedSingle"
$form.Opacity = 1

# ラベル SKIP A
$LabelSkipA = New-Object System.Windows.Forms.Label
$LabelSkipA.Location = "15,18"
$LabelSkipA.Size = New-Object System.Drawing.Size(55,30)
$LabelSkipA.Text = "SKIP A"
$LabelSkipA.BackColor = "black"
$LabelSkipA.Forecolor = "yellow"
$LabelSkipA.TextAlign = "MiddleCenter"

# ラベル SKIP B
$LabelSkipB = New-Object System.Windows.Forms.Label
$LabelSkipB.Location = "80,18"
$LabelSkipB.Size = New-Object System.Drawing.Size(55,30)
$LabelSkipB.Text = "SKIP B"
$LabelSkipB.BackColor = "black"
$LabelSkipB.Forecolor = "yellow"
$LabelSkipB.TextAlign = "MiddleCenter"

# ラベル 30秒スキップ
$LabelSkipPrev = New-Object System.Windows.Forms.Label
$LabelSkipPrev.Location = "145,18"
$LabelSkipPrev.Size = New-Object System.Drawing.Size(55,30)
$LabelSkipPrev.Text = "+30"
$LabelSkipPrev.BackColor = "black"
$LabelSkipPrev.Forecolor = "yellow"
$LabelSkipPrev.TextAlign = "MiddleCenter"

# ラベル 30秒バック
$LabelSkipBack = New-Object System.Windows.Forms.Label
$LabelSkipBack.Location = "145,63"
$LabelSkipBack.Size = New-Object System.Drawing.Size(55,30)
$LabelSkipBack.Text = "-30"
$LabelSkipBack.BackColor = "black"
$LabelSkipBack.Forecolor = "yellow"
$LabelSkipBack.TextAlign = "MiddleCenter"

# Yesボタンの設定
$YesButton = New-Object System.Windows.Forms.Button
$YesButton.Location = New-Object System.Drawing.Point(15,63)
$YesButton.Size = New-Object System.Drawing.Size(55,30)
$YesButton.Text = "OPEN"
$YesButton.DialogResult = "Yes"
$YesButton.Flatstyle = "Popup"
$YesButton.backcolor = "black"
$YesButton.forecolor = "yellow"

# ラベル
$LabelPause = New-Object System.Windows.Forms.Label
$LabelPause.Location = "80,63"
$LabelPause.Size = New-Object System.Drawing.Size(55,30)
$LabelPause.Text = "PAUSE" 
$LabelPause.BackColor = "black"
$LabelPause.Forecolor = "yellow"
$LabelPause.TextAlign = "MiddleCenter"

# ラベル はじめから
$LabelSkipZero = New-Object System.Windows.Forms.Label
$LabelSkipZero.Location = "15,108"
$LabelSkipZero.Size = New-Object System.Drawing.Size(55,30)
$LabelSkipZero.Text = "SKIP 0"
$LabelSkipZero.BackColor = "black"
$LabelSkipZero.Forecolor = "yellow"
$LabelSkipZero.TextAlign = "MiddleCenter"


# PC TV Plus 再起動
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(80,108)
$OKButton.Size = New-Object System.Drawing.Size(55,30)
$OKButton.Text = "PC TV"
$OKButton.DialogResult = "OK"
$OKButton.Flatstyle = "Popup"
$OKButton.Backcolor = "black"
$OKButton.forecolor = "yellow"

# キャンセルボタンの設定
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(145,108)
$CancelButton.Size = New-Object System.Drawing.Size(55,30)
$CancelButton.Text = "SS"
$CancelButton.DialogResult = "Retry"
$CancelButton.Flatstyle = "Popup"
$CancelButton.backcolor = "black"
$CancelButton.forecolor = "yellow"


$form.Controls.Add($LabelSkipA)
$form.Controls.Add($LabelSkipB)
$form.Controls.Add($LabelSkipPrev)
$form.Controls.Add($LabelSkipBack)
$form.Controls.Add($LabelRecovery)
$form.Controls.Add($LabelPause)
$form.Controls.Add($LabelSkipZero)

$file = $null

$FuncPCTVReStart = {
    sleep -Milliseconds 100
    $ps_vnt = getProcess( "Vnt" )
    if ( $ps_vnt -ne $null ) {
        $ret = $ps_vnt.CloseMainWindow()
    }
    sleep -Milliseconds 5000
    $vnt_file_name_pos = $pc_tv_plus_path.LastIndexOf( "\" )
    $vnt_exe_name = $pc_tv_plus_path.Substring( $vnt_file_name_pos + 1 )
    $vnt_working_directory = $pc_tv_plus_path.Substring( 0, $vnt_file_name_pos )
    $proc = Start-Process -FilePath $vnt_exe_name -WorkingDirectory $vnt_working_directory -PassThru
}

$FuncScreenShot = {
    # スクリーンショット
    if (  $comment_viewer_size_max -eq $True ) {
        $ps_viewer = getProcess( $comment_viewer_app )
        # ウィンドウ最大化
        if ( $ps_viewer.Name -eq $comment_viewer_app ) {
            $win32::ShowWindowAsync($ps_viewer.MainWindowHandle, 3) | Out-Null
        }
    }
    sleep -Milliseconds 500
    $screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
    $screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height
    $bounds = [Drawing.Rectangle]::FromLTRB(0, 0, $screenWidth, $screenHeight)
    $file_name = (Get-Date).ToString("yyyyMMddHHmmss") + "screenshot.jpg"
    $save_path = $screenshot_dir + "\" + $file_name
    screenshot $bounds $save_path
}


Function file_open(){
    
    #アセンブリのロード
    Add-Type -AssemblyName System.Windows.Forms

    #ダイアログインスタンス生成
    $dialog = New-Object Windows.Forms.OpenFileDialog
    
    $ret = $dialog.Filter = "実況ログファイル(*.xml) | *.xml"
    $ret = $dialog.InitialDirectory = $log_dir
  
    #ダイアログ表示
    $result = $dialog.ShowDialog()

    #「開くボタン」押下ならファイル名フルパスをリターン
    If($result -eq "OK"){
        return $dialog.FileName 
    } Else {
        return $null
    }

}

$FuncLogFileOpen = {
    #ファイル取得
    $script:file = file_open
    if ( $script:file -ne $none ) {
        $YesButton.Forecolor = "red"
    } else {
        $YesButton.Forecolor = "yellow"
    }
}

$FuncSkipA = {
    # 次のチャプターとAのコメントまで移動する
    . Send-Keys " a" $comment_viewer_app
    . Send-Keys "^({RIGHT})" Vnt
    sleep -Milliseconds $a_b_skip_wait
    . Send-Keys " " $comment_viewer_app
    [Void]$win32::SetForegroundWindow( $script:ps_vnt.MainWindowHandle )
}

$FuncSkipB = {
    # 次のチャプターとBのコメントまで移動する
    . Send-Keys " b" $comment_viewer_app
    . Send-Keys "^({RIGHT})" Vnt
    sleep -Milliseconds $a_b_skip_wait
    . Send-Keys " " $comment_viewer_app
    [Void]$win32::SetForegroundWindow( $script:ps_vnt.MainWindowHandle )
}

$FuncSkipPrev = {
    # 30秒飛ばし
    . Send-Keys " 3" $comment_viewer_app
    . Send-Keys "%({RIGHT})" Vnt
    # commenomiが進みすぎるときはこの値を増やす
    sleep -Milliseconds $prev_skip_wait
    . Send-Keys " " $comment_viewer_app
    [Void]$win32::SetForegroundWindow( $script:ps_vnt.MainWindowHandle )
}

$FuncSkipBack = {
    # 30秒戻し
    . Send-Keys " 2" $comment_viewer_app
    . Send-Keys "%({LEFT}{LEFT})" Vnt
    # commenomiが進みすぎるときはこの値を増やす
    sleep -Milliseconds $back_skip_wait
    . Send-Keys " " $comment_viewer_app
    [Void]$win32::SetForegroundWindow( $script:ps_vnt.MainWindowHandle )
}

$FuncPause = {
    # 一時停止
    . Send-Keys " " $comment_viewer_app
    . Send-Keys " " Vnt
#    sleep -Milliseconds 100
#    [Void]$win32::SetForegroundWindow( $script:ps_vnt.MainWindowHandle )
}

$FuncChangeVntSize = {
    param($Size)
    $ps_vnt = getProcess("Vnt")
    sleep -Milliseconds 100
    $screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
    $screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height
    # PC TV Plusウィンドウ位置調整
    if ( $ps_vnt -ne $null ) {
        switch ( $Size ) {
            "S" {
                $top, $width, $shift = $size_S_top, $size_S_width, $size_S_shift
            }
            "M" {
                $top, $width, $shift = $size_M_top, $size_M_width, $size_M_shift
            }
            "L" {
                $top, $width, $shift = $size_L_top, $size_L_width, $size_L_shift
            }
            "FULL" {
                sleep -Milliseconds 100
                $win32::SetWindowPos( $ps_vnt.MainWindowHandle, $HWND_NOTOPMOST, 0, 0, $screenWidth, $screenHeight, $SWP_SHOWWINDOW)
                return
            }
        }
        # 1920*1080の設定値から実際の解像度に合わせて再計算
        $real_width = [Math]::Round( $width * $screenWidth / 1920 )
        $seek_bar_height = 36
        $comment_line_count = 12
        $height = ($real_width / 16 * 9)

        $height_plus_per = $screenHeight / ($screenWidth / 16 * 9 )
        $real_top = [Math]::Round( ( ( $screenHeight - $seek_bar_height ) / $comment_line_count ) * ( $top / 87 ) * $height_plus_per )

        $left = [Math]::Round( ( ( $screenWidth - $real_width ) / 2) )
        if ( $left -lt ( $real_window_pos_left + 231 ) ) {
            $left = $real_window_pos_left + 231
        }
        $left = $left + $shift
        $win32::SetWindowPos( $ps_vnt.MainWindowHandle, $HWND_TOPMOST, $left, $real_top, $real_width, $height, $SWP_SHOWWINDOW)
    }
}

$FuncFileOpenClick = {
    if ($_.Button -eq "Left" ) {
        . $FuncLogFileOpen
    } else {
        $script:file = $null
        $LabelLog.Forecolor = "yellow"
    }
    $script:komenasne_option = $null
}

$FuncSkipZero = {
    . Send-Keys "%$("{LEFT}" * 30)" Vnt
    sleep -Milliseconds 1000
    . Send-Keys "0" $comment_viewer_app
    [Void]$win32::SetForegroundWindow( $script:ps_vnt.MainWindowHandle )
}


$FuncChangeOpenClose = {
    if ( $YesButton.Text -eq "OPEN" ) {
        $YesButton.Text = "CLOSE"
    } else {
        $YesButton.Text = "OPEN"
    }
}

$FuncOpenKomenasne = {
    # komenasneを開く
    $komenasne_file_name_pos = $komenasne_path.LastIndexOf( "\" )
    $komenasne_exe_name = $komenasne_path.Substring( $komenasne_file_name_pos + 1 )
    $komenasne_working_directory = $komenasne_path.Substring( 0, $komenasne_file_name_pos )
    if ( $komenasne_option -eq $null ) {
        $proc = Start-Process -FilePath $komenasne_exe_name -WorkingDirectory $komenasne_working_directory -PassThru
    } else {
    #Write-Host $komenasne_option
        $proc = Start-Process -FilePath $komenasne_exe_name -WorkingDirectory $komenasne_working_directory -ArgumentList $komenasne_option -PassThru
    }
}


$LabelSkipA.Add_Click($FuncSkipA)
$LabelSkipB.Add_Click($FuncSkipB)
$LabelSkipPrev.Add_Click($FuncSkipPrev)
$LabelSkipBack.Add_Click($FuncSkipBack)
$LabelPause.Add_Click($FuncPause)
$LabelSkipZero.Add_Click($FuncSkipZero)



# 最前面に表示：する
$form.Topmost = $True

# キーとボタンの関係
$form.AcceptButton = $OKButton
$form.CancelButton = $CancelButton

# ボタン等をフォームに追加
$form.Controls.Add($OKButton)
$form.Controls.Add($CancelButton)
$form.Controls.Add($YesButton)
$form.Controls.Add($AbortButton)
$form.Controls.Add($NoneButton)


# インスタンス化
$Context = New-Object System.Windows.Forms.ContextMenuStrip

# 項目を追加
[void]$Context.Items.Add("動画サイズ S")
[void]$Context.Items.Add("動画サイズ M")
[void]$Context.Items.Add("動画サイズ L")
[void]$Context.Items.Add("動画サイズ FULL")
[void]$Context.Items.Add("過去ログファイルを開く")
[void]$Context.Items.Add("チャンネルと日時を指定")
[void]$Context.Items.Add(" ") # 一行開ける
[void]$Context.Items.Add("OPEN CLOSE 切り替え")


$Click = {
    [String]$A = $_.ClickedItem
    $Context.Close()
    IF ( $A -eq "動画サイズ S") {
        . $FuncChangeVntSize "S"
    }elseif( $A -eq "動画サイズ M" ) {
        . $FuncChangeVntSize "M"
    }elseif( $A -eq "動画サイズ L" ) {
        . $FuncChangeVntSize "L"
    }elseif( $A -eq "動画サイズ FULL" ) {
        . $FuncChangeVntSize "FULL"
    }elseif( $A -eq "過去ログファイルを開く" ) {
        . $FuncLogFileOpen
    }elseif( $A -eq "チャンネルと日時を指定" ) {
        . $SubForm
    }elseif( $A -eq "OPEN CLOSE 切り替え" ) {
        . $FuncChangeOpenClose
    }
}
$Context.Add_ItemClicked($Click)

$FuncTimeTable = {
    param( $url )
    $response = Invoke-WebRequest $url -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36"
    if ( $response.StatusCode -eq 200 ) {
        $resp = [System.Text.Encoding]::UTF8.GetString( [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($response.Content) )
#$resp = '<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"><meta name="description" content="日テレ ルパン三世　「偽りのファンタジスタ」#2の番組詳細 （2015/10/09 01:29 ～ 2015/10/09 01:59）"><meta name="robots" content="noarchive"><title>番組詳細 日テレ ルパン三世　「偽りのファンタジスタ」#2 （2015/10/09 01:29 ～ 2015/10/09 01:59）</title><link rel="stylesheet" href="https://timetable.yanbe.net/css/time_table.css?20210303014027"><script src="https://timetable.yanbe.net/js/ex2.js?20210303014027"></script><script src="https://timetable.yanbe.net/js/p_set.js?20210303014027"></script></head><body><div id="banner"><a href="https://timetable.yanbe.net/?p=13">テレビ番組表の記録</a></div><!-- google_ad_section_start --><div id="content"><h2 class="title">番組詳細</h2><div class="title_menu">表示切替：<a href="https://timetable.yanbe.net/html/13/2015/10/08_1.html?13">番組表に戻る</a></div><div class="title_under_ads"><script src="https://timetable.yanbe.net/js/google_ads1.js"></script><script src="https://pagead2.googlesyndication.com/pagead/show_ads.js"></script></div><div class="clear_both"></div><div class="main_body"><h3 class="pdv_title">番組名</h3><p class="pdv_text">ルパン三世　「偽りのファンタジスタ」#2</p><h3 class="pdv_title">放送</h3><p class="pdv_text">日テレ<br>2015/10/09 01:29 ～ 2015/10/09 01:59　（30分）</p><h3 class="pdv_title">番組内容</h3><p class="pdv_text">1985年に放送終了した「ルパン三世PARTIII」以来30年ぶりにテレビシリーズが復活!新たなルパン三世が青いジャケットに身を包みイタリアとサンマリノを駆け巡る!</p><h3 class="pdv_title">その他</h3><p class="pdv_text">アニメ／特撮 - 国内アニメ</p><hr><p class="pdv_text"><a href="https://livesubject.yanbe.net/?p=0&bn=&sy=2015&sm=10&sd=09&sh=01&ey=2015&em=10&ed=09&eh=01&rs=2&sj=&sn=2&st=1" target="_blank" rel="noopener">放送時間帯の実況板スレを検索（別サイト）</a><br>全国ネット番組の場合はレス記録から<br>放送内容や延長情報などの情報が得れる場合があります</p></div></div><!-- google_ad_section_end --><div id="links"><form action="https://timetable.yanbe.net/" method="get" name="input"><input type="hidden" name="d" value="20151008"><div class="sidetitle">メニュー</div><div class="side"><div class="side_menu_area">地域<br><select name="p" class="p_select" onchange="goreload(this.form);return false"><option value="1">北海道</option><option value="2">青森</option><option value="3">岩手</option><option value="4">宮城</option><option value="5">秋田</option><option value="6">山形</option><option value="7">福島</option><option value="8">茨城</option><option value="9">栃木</option><option value="10">群馬</option><option value="11">埼玉</option><option value="12">千葉</option><option value="13" selected="selected">東京</option><option value="14">神奈川</option><option value="15">新潟</option><option value="16">富山</option><option value="17">石川</option><option value="18">福井</option><option value="19">山梨</option><option value="20">長野</option><option value="21">岐阜</option><option value="22">静岡</option><option value="23">愛知</option><option value="24">三重</option><option value="25">滋賀</option><option value="26">京都</option><option value="27">大阪</option><option value="28">兵庫</option><option value="29">奈良</option><option value="30">和歌山</option><option value="31">鳥取</option><option value="32">島根</option><option value="33">岡山</option><option value="34">広島</option><option value="35">山口</option><option value="36">徳島</option><option value="37">香川</option><option value="38">愛媛</option><option value="39">高知</option><option value="40">福岡</option><option value="41">佐賀</option><option value="42">長崎</option><option value="43">熊本</option><option value="44">大分</option><option value="45">宮崎</option><option value="46">鹿児島</option><option value="47">沖縄</option></select></div><div class="side_menu_year">年<br><a href="https://timetable.yanbe.net/?d=2021&amp;p=13">2021年</a><br><a href="https://timetable.yanbe.net/?d=2020&amp;p=13">2020年</a><br><a href="https://timetable.yanbe.net/?d=2019&amp;p=13">2019年</a><br><a href="https://timetable.yanbe.net/?d=2018&amp;p=13">2018年</a><br><a href="https://timetable.yanbe.net/?d=2017&amp;p=13">2017年</a><br><a href="https://timetable.yanbe.net/?d=2016&amp;p=13">2016年</a><br><a href="https://timetable.yanbe.net/?d=2015&amp;p=13">2015年</a><br><a href="https://timetable.yanbe.net/?d=2014&amp;p=13">2014年</a><br><a href="https://timetable.yanbe.net/?d=2013&amp;p=13">2013年</a><br><a href="https://timetable.yanbe.net/?d=2012&amp;p=13">2012年</a><br><a href="https://timetable.yanbe.net/?d=2011&amp;p=13">2011年</a><br><a href="https://timetable.yanbe.net/?d=2010&amp;p=13">2010年</a><br><a href="https://timetable.yanbe.net/?d=2009&amp;p=13">2009年</a><br><a href="https://timetable.yanbe.net/?d=2008&amp;p=13">2008年</a><br><a href="https://timetable.yanbe.net/?d=2007&amp;p=13">2007年</a><br></div></div></form><script src="https://timetable.yanbe.net/amazon_top.cgi"></script></div><script src="https://timetable.yanbe.net/js/google_analytics1.js"></script></body></html>'
        $resp -match "<title>(.+?)</title>"
        $page_title = $Matches[1]
        $ret = $page_title -match "番組詳細 (.+?)・(.+?) （(.+?) ～ (.+?)）"
        if ( $ret -eq $False ) {
            $ret = $page_title -match "番組詳細 (.+?) (.+?) （(.+?) ～ (.+?)）"
        }
        if ( $ret -eq $True ) {
            $channel = $Matches[1]
            $timetable_title = $Matches[2]
            $start_date_time = [DateTime]$Matches[3]
            $end_date_time = [DateTime]$Matches[4]
            $timetable_total_minutes = ($end_date_time - $start_date_time).TotalMinutes
            $timetable_start_date_time = $start_date_time.ToString("yyyy-MM-dd HH:mm")
            if ( $channel.StartsWith( "ＮＨＫ総合" ) ) {
                $timetable_channel_index = 0
            } elseif ( $channel.StartsWith( "ＮＨＫＥテレ" ) ) {
                $timetable_channel_index = 1
            } elseif ( $channel.StartsWith( "日テレ" ) ) {
                $timetable_channel_index = 2
            } elseif ( $channel.StartsWith( "テレビ朝日" ) ) {
                $timetable_channel_index = 3
            } elseif ( $channel.StartsWith( "ＴＢＳ" ) ) {
                $timetable_channel_index = 4
            } elseif ( $channel.StartsWith( "テレビ東京" ) ) {
                $timetable_channel_index = 5
            } elseif ( $channel.StartsWith( "フジテレビ" ) ) {
                $timetable_channel_index = 6
            } elseif ( $channel.StartsWith( "ＴＯＫＹＯ　ＭＸ" ) ) {
                $timetable_channel_index = 7
            } elseif ( $channel.StartsWith( "ＮＨＫ　ＢＳ１" ) ) {
                $timetable_channel_index = 8
            } elseif ( $channel.StartsWith( "ＮＨＫ　ＢＳプレミアム" ) ) {
                $timetable_channel_index = 9
            } elseif ( $channel.StartsWith( "ＢＳ日テレ" ) ) {
                $timetable_channel_index = 10
            } elseif ( $channel.StartsWith( "ＢＳ朝日" ) ) {
                $timetable_channel_index = 11
            } elseif ( $channel.StartsWith( "ＢＳ－ＴＢＳ" ) ) {
                $timetable_channel_index = 12
            } elseif ( $channel.StartsWith( "ＢＳテレ東" ) ) {
                $timetable_channel_index = 13
            } elseif ( $channel.StartsWith( "ＢＳフジ" ) ) {
                $timetable_channel_index = 14
            } elseif ( $channel.StartsWith( "ＷＯＷＯＷプライム" ) ) {
                $timetable_channel_index = 15
            } elseif ( $channel.StartsWith( "BS11" ) ) {
                $timetable_channel_index = 16
            } elseif ( $channel.StartsWith( "BS12" ) ) {
                $timetable_channel_index = 17
            } elseif ( $channel.StartsWith( "アニメシアターＸ" ) ) {
                $timetable_channel_index = 18
            }
        }
    }
}


$SubForm = {
    # フォームの作成
    $private:form = New-Object System.Windows.Forms.Form 
    $private:form.Text = "入力"
    $private:form.Size = New-Object System.Drawing.Size(515,340) 

    # OKボタンの設定
    $private:OKButton = New-Object System.Windows.Forms.Button
    $private:OKButton.Location = New-Object System.Drawing.Point(90,250)
    $private:OKButton.Size = New-Object System.Drawing.Size(75,30)
    $private:OKButton.Text = "OK"
    $private:OKButton.DialogResult = "OK"	
    # 列挙子名：None, OK, Cancel, Abort, Retry, Ignore, Yes, No

    # キャンセルボタンの設定
    $private:CancelButton = New-Object System.Windows.Forms.Button
    $private:CancelButton.Location = New-Object System.Drawing.Point(170,250)
    $private:CancelButton.Size = New-Object System.Drawing.Size(75,30)
    $private:CancelButton.Text = "Cancel"
    $private:CancelButton.DialogResult = "Cancel"
    # 列挙子名：None, OK, Cancel, Abort, Retry, Ignore, Yes, No

    # ラベルの設定
    $labelReset = New-Object System.Windows.Forms.Label
    $labelReset.Location = New-Object System.Drawing.Point(300,250) 
    $labelReset.Size = New-Object System.Drawing.Size(75,30) 
    $labelReset.Text = "Reset"
    $labelReset.BackColor = "lightgray"
    $labelReset.Forecolor = "black"
    $labelReset.TextAlign = "MiddleCenter"

    # ラベルの設定
    $labelUrl1 = New-Object System.Windows.Forms.Label
    $labelUrl1.Location = New-Object System.Drawing.Point(10,10) 
    $labelUrl1.Size = New-Object System.Drawing.Size(190,20) 
    $labelUrl1.Text = "（省略可能）番組詳細のURLを入力"
    
    # ラベルの設定
    $labelUrl = New-Object System.Windows.Forms.Label
    $labelUrl.Location = New-Object System.Drawing.Point(200,10) 
    $labelUrl.Size = New-Object System.Drawing.Size(200,20) 
    $labelUrl.Forecolor = "blue"
    $labelUrl.Text = "クリックして番組表を開く"

    # 入力ボックスの設定
    $textBoxTimeTable = New-Object System.Windows.Forms.TextBox 
    $textBoxTimeTable.Location = New-Object System.Drawing.Point(10,30) 
    $textBoxTimeTable.Size = New-Object System.Drawing.Size(470,50) 

    # ラベルの設定
    $labelGetUrl = New-Object System.Windows.Forms.Label
    $labelGetUrl.Location = New-Object System.Drawing.Point(10,55) 
    $labelGetUrl.Size = New-Object System.Drawing.Size(75,20) 
    $labelGetUrl.Text = "読み込み"
    $labelGetUrl.BackColor = "lightgray"
    $labelGetUrl.Forecolor = "black"
    $labelGetUrl.TextAlign = "MiddleCenter"

    # ラベルの設定
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,90) 
    $label1.Size = New-Object System.Drawing.Size(350,20) 
    $label1.Text = "チャンネルを選択"

    # コンボボックスを作成
    $comboBox = New-Object System.Windows.Forms.Combobox
    $comboBox.Location = New-Object System.Drawing.Point(10,110)
    $comboBox.size = New-Object System.Drawing.Size(100,30)
    $comboBox.DropDownStyle = "DropDown"
    $comboBox.FlatStyle = "standard"
    $comboBox.font = $Font
    $comboBox.BackColor = "#005050"
    $comboBox.ForeColor = "white"

    # ラベルの設定　
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,150) 
    $label2.Size = New-Object System.Drawing.Size(125,20) 
    $label2.Text = "YYYY-MM-DD HH:mm"

    # 入力ボックスの設定
    $textBoxDate = New-Object System.Windows.Forms.TextBox 
    $textBoxDate.Location = New-Object System.Drawing.Point(10,170) 
    $textBoxDate.Size = New-Object System.Drawing.Size(125,50) 

    # ラベルの設定
    $label3 = New-Object System.Windows.Forms.Label
    $label3.Location = New-Object System.Drawing.Point(150,150) 
    $label3.Size = New-Object System.Drawing.Size(250,20) 
    $label3.Text = "分数(半角数字)"

    # 入力ボックスの設定
    $textBoxMinutes = New-Object System.Windows.Forms.TextBox 
    $textBoxMinutes.Location = New-Object System.Drawing.Point(150,170) 
    $textBoxMinutes.Size = New-Object System.Drawing.Size(50,50) 

    # ラベルの設定
    $label4 = New-Object System.Windows.Forms.Label
    $label4.Location = New-Object System.Drawing.Point(10,200) 
    $label4.Size = New-Object System.Drawing.Size(250,20) 
    $label4.Text = "番組名(任意)"

    # 入力ボックスの設定
    $textBoxTitle = New-Object System.Windows.Forms.TextBox 
    $textBoxTitle.Location = New-Object System.Drawing.Point(10,220) 
    $textBoxTitle.Size = New-Object System.Drawing.Size(470,50) 

    # コンボボックスに項目を追加
    $ch_list = [ordered]@{
        "NHK総合"          = "jk1"
        "NHK Eテレ"        = "jk2"
        "日本テレビ"       = "jk4"
        "テレビ朝日"       = "jk5"
        "TBSテレビ"        = "jk6"
        "テレビ東京"       = "jk7"
        "フジテレビ"       = "jk8"
        "TOKYO MX"         = "jk9"
        "NHK BS1"          = "jk101"
        "NHK BSプレミアム" = "jk103"
        "BS日テレ"         = "jk141"
        "BS朝日"           = "jk151"
        "BS-TBS"           = "jk161"
        "BSテレ東"         = "jk171"
        "BSフジ"           = "jk181"
        "WOWOWプライム"    = "jk191"
        "BS11イレブン"     = "jk211"
        "BS12トゥエルビ"   = "jk222"
        "AT-X"             = "jk236"
    }
    foreach ( $ch in $ch_list.keys ) {
        [void] $comboBox.Items.Add( $ch )
    }

    # キーとボタンの関係
    $private:form.AcceptButton = $private:OKButton
    $private:form.CancelButton = $private:CancelButton
    
    # ボタン等をフォームに追加
    $private:form.Controls.Add($private:OKButton)
    $private:form.Controls.Add($private:CancelButton)
    $private:form.Controls.Add($labelReset)
    $private:form.Controls.Add($labelUrl1) 
    $private:form.Controls.Add($labelUrl) 
    $private:form.Controls.Add($textBoxTimeTable)
    $private:form.Controls.Add($labelGetUrl) 
    $private:form.Controls.Add($label1) 
    $private:form.Controls.Add($label2) 
    $private:form.Controls.Add($label3) 
    $private:form.Controls.Add($label4) 
    $private:form.Controls.Add($textBoxDate)
    $private:form.Controls.Add($textBoxMinutes)
    $private:form.Controls.Add($textBoxTitle)
    $form.Controls.Add($comboBox)

    $labelUrl.Add_Click(
        {
            start ‘https://timetable.yanbe.net/'
        }
    )

    $labelGetUrl.Add_Click(
        {
            . $FuncTimeTable $textBoxTimeTable.Text
            if ( $timetable_title -ne "" ) {
                $comboBox.SelectedIndex = $script:komenasChSelect =  $timetable_channel_index
                $textBoxDate.Text = $script:komenasDate = $timetable_start_date_time
                $textBoxMinutes.Text = $script:komenasMinutes = $timetable_total_minutes
                $textBoxTitle.Text = $script:komenasTitle = $timetable_title
            }
        }
    )

    $labelReset.Add_Click(
        {
            $textBoxTimeTable.Text = $script:timeTableUrl = ""
            $comboBox.SelectedIndex = $script:komenasChSelect = 0
            $textBoxDate.Text = $script:komenasDate = ""
            $textBoxMinutes.Text = $script:komenasMinutes = ""
            $textBoxTitle.Text = $script:komenasTitle = ""
            $textBoxTimeTable.Select()
        }
    )

    #フォームを常に手前に表示
    $private:form.Topmost = $True

    $textBoxTimeTable.Text = $script:timeTableUrl
    $comboBox.SelectedIndex = $script:komenasChSelect
    $textBoxDate.Text = $script:komenasDate
    $textBoxMinutes.Text = $script:komenasMinutes
    $textBoxTitle.Text = $script:komenasTitle

    #フォームをアクティブにし、テキストボックスにフォーカスを設定
    #$textBoxTimeTable.Text = "https://timetable.yanbe.net/pdv.cgi?d=20210218&p=13&v=1&c=101101024202102182345"
    $textBoxTimeTable.Select()
#    $form.Add_Shown({})
 
    # フォームを表示させ、その結果を受け取る
    $private:result = $private:form.ShowDialog()
    
    # 結果による処理分岐
    if ($private:result -eq "OK") {
        $ch_name = $comboBox.text
        $file = $null
        $script:timeTableUrl = $textBoxTimeTable.Text
        $script:komenasChSelect = $comboBox.SelectedIndex
        $script:komenasDate = $textBoxDate.Text
        $script:komenasMinutes = $textBoxMinutes.Text
        $script:komenasTitle = $textBoxTitle.Text
        if ( ( $textBoxDate.text -ne "" ) -and ( $textBoxMinutes.text -ne "" ) ) {
    	    $script:komenasne_option = """" + $ch_list.$ch_name + """ """ + $textBoxDate.text + """ " + $textBoxMinutes.Text + " """ + $textBoxTitle.Text + """"
            $YesButton.forecolor = "limegreen"
        } else {
            $script:komenasne_option = $null
        }
    }
}

$FuncHideForm = {
    if ( $form_auto_hide -eq $True ) {
        $form.Size = New-Object System.Drawing.Size(136, 55)
        $form.Opacity = $form_opacity
    }
}

$FuncShowForm = {
    $form.Size = New-Object System.Drawing.Size(230,190)
    $form.Opacity = 1
}



# クリックイベントの内容
$MouseClick = {
    if ( $_.Button -eq "Right" ) {
        $Context.Show()
    } elseif ( $_.Button -eq "Left" ) {
        . $FuncHideForm
    }
}
$form.Add_MouseDown($MouseClick)
$MouseHover = {
#    if ( $form.Size.Height -eq 55 ) {
        . $FuncShowForm
#    }
}
$form.Add_MouseHover($MouseHover)


# フォームにコンテキストメニューを追加
$form.ContextMenuStrip = $Context

$VerbosePreference = 'Continue'

# フォームを表示させ、その結果を受け取る
$result = $form.ShowDialog()


# 結果による処理分岐
while ($result -ne "Cancel") {
    if ($result -eq "OK") {
        . $FuncPCTVReStart
    }

    if ($result -eq "Yes") {

        if ($YesButton.Text -eq "OPEN") {
            $ps_vnt = getProcess("Vnt")
            $error_flg = $false
            if ( $file -ne $null ) {
                $file_name = """" + $file + """"
                $proc = Start-Process -FilePath $comment_viewer_exe_name -WorkingDirectory $comment_viewer_install_path -ArgumentList $file_name -PassThru
                $file = $null
                sleep -Milliseconds  4000
            } else {
                . $FuncOpenKomenasne
                Wait-Process -InputObject $proc
                if ( $proc.ExitCode -ne 0 ) {
                    # 一度だけリトライする
                    sleep -Milliseconds  3000
                    . $FuncOpenKomenasne
                    Wait-Process -InputObject $proc
                }
                if ( $proc.ExitCode -ne 0 ) {
#                        $f = New-Object System.Windows.Forms.form
#                        $f.TopMost = $true
#                        $f.Left = 0
#                        $ret = [System.Windows.Forms.MessageBox]::Show($f, "komenasneが見つかりません.", "Error")
                    $error_flg = $true
                }
            }
            if ( $error_flg -eq $false ) {
                if ( $ps_vnt -ne $null ) {
                    sleep -Milliseconds  4000
                    # PC TV Plusを倍速再生にする
                    if ( $enable_speed_up -eq $True ) {
                        . Send-Keys '+(^(G))' Vnt
                    }
                }
                $count = 0
                while ($count -lt 50) {
                    sleep -Milliseconds 100
                    $ps_viewer = getProcess( $comment_viewer_app )
                    if ( $ps_viewer -ne $null ) {
                        $YesButton.Text = "CLOSE"
                        break
                    }
                    $count++
                }
                if ( $ps_viewer -ne $null ) {
                    if (  $comment_viewer_size_max -eq $True ) {
                        # ウィンドウ最大化
                        $win32::ShowWindowAsync($ps_viewer.MainWindowHandle, 3) | Out-Null
                        $form.Size = New-Object System.Drawing.Size(136, 55)
                    } else {
                        . Send-Keys '{F5}' $comment_viewer_app
                    }
                }
                # フォームを半透明に
                . $FuncHideForm
            }
            $komenasne_option = $null
            $YesButton.Forecolor = "yellow"
        } else {
            # commenomiを閉じる
            sleep -Milliseconds 100
            $ps_viewer = getProcess( $comment_viewer_app )
            if ( $ps_viewer -ne $null  ) {
                $ret = $ps_viewer.CloseMainWindow()
            }
            $ps_vnt = getProcess("Vnt")
            if ( $ps_vnt -ne $null ) {
                sleep -Milliseconds 1000
                . Send-Keys "{BACKSPACE}" Vnt
            }
            $YesButton.Text = "OPEN"
            $form.Opacity = 1
        }
    }

    if ($result -eq "Retry") {
        # スクリーンショット
        . $FuncScreenShot
    }


    # フォームを表示させ、その結果を受け取るA
    $result = $form.ShowDialog()
}
