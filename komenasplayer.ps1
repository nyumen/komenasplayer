# コメントプレイヤーのAPP名
$comment_viewer_app = "commenomi"

# コメントプレイヤー最大化 ( $True or $False)
$comment_viewer_size_max = $True

# ウィンドウのデフォルト位置 1920x1080の場合 上から800,左から15（好みに合わせて変更）
$default_window_pos_top = 760
$default_window_pos_left = 20

# PC TV Plusを早見再生する ( $True or $False )
$enable_speed_up = $True

# 30秒送りでcommenomiが進みすぎるときはこの値を増やす（1000で1秒）
$prev_skip_wait = 1000

# 30秒戻しでcommenomiが進みすぎるときはこの値を増やす（1000で1秒）
$back_skip_wait = 1000

# フォームの透明度 ( 0～1 )
$form_opacity = 0.6

# 過去ログフォルダ
$log_dir = "..\komenasne\kakolog"

# スクリーンショットフォルダ
$screenshot_dir = ".\screenshot"

# Thanks
# http://kamifuji.dyndns.org/PS-Support/


# アセンブリの読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$signature=@' 
      [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
      public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@

$SendMouseEvent = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru

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

# ソースコードからセッションにクラスを追加
#Add-Type -TypeDefinition $src

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

# https://owlcamp.jp/powershell%E3%81%A7%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%BC%E3%83%B3%E3%82%AD%E3%83%A3%E3%83%97%E3%83%81%E3%83%A3%E3%82%92%E8%87%AA%E5%8B%95%E3%81%A7%E5%8F%96%E5%BE%97%E3%81%97%E3%81%A6%E3%83%95%E3%82%A1/
$dll_info = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
    Add-Type -MemberDefinition $dll_info -Name NativeMethods -Namespace Win32
 

# スクリーンショット
# https://www.it-swarm-ja.tech/ja/powershell/windows-powershell%E3%81%A7%E7%94%BB%E9%9D%A2%E3%82%AD%E3%83%A3%E3%83%97%E3%83%81%E3%83%A3%E3%82%92%E5%AE%9F%E8%A1%8C%E3%81%99%E3%82%8B%E3%81%AB%E3%81%AF%E3%81%A9%E3%81%86%E3%81%99%E3%82%8C%E3%81%B0%E3%82%88%E3%81%84%E3%81%A7%E3%81%99%E3%81%8B%EF%BC%9F/969655941/
function screenshot([Drawing.Rectangle]$bounds, $path) {
   $bmp = New-Object Drawing.Bitmap $bounds.width, $bounds.height
   $graphics = [Drawing.Graphics]::FromImage($bmp)

   $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)

   $bmp.Save($path)

   $graphics.Dispose()
   $bmp.Dispose()
}

function getDefaultPos() {
    sleep -Milliseconds  100

    # $rectにPC TV Plusの4隅のスクリーン座標が格納される
    $ps = getProcess ( "Vnt" )
    $rect = [Win32Api.Helper]::GetForegroundWindowRect( $ps.MainWindowHandle )
    $VntX = $rect.Left + 20
    $VntY = $rect.Bottom - 180
}

function Send-Keys($KeyStroke, $ProcessName) {
    $ps = getProcess($ProcessName)
    # IMEがオンだったらオフにする
    if ([PowerShell.IME]::GetState( $ps.MainWindowHandle )) {
        [PowerShell.IME]::SetState( $ps.MainWindowHandle, $false )
        sleep -Milliseconds  100
    }

    # PC TV Plusのウィンドウサイズを取得する
    . getDefaultPos 

    if ($ProcessName -eq "Vnt") {
        $x = $VntX
        $y = $VntY
    }
    if ($ProcessName -eq $comment_viewer_app ) {
        $x = $form.Left + 240
        $y = $form.Top - 1
    }

    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)


    # 左クリック
    $SendMouseEvent::mouse_event(0x00000002, 0, 0, 0, 0);
    $SendMouseEvent::mouse_event(0x00000004, 0, 0, 0, 0);

    sleep -Milliseconds  100

    [System.Windows.Forms.SendKeys]::SendWait($KeyStroke)
}

function getProcess($ProcessName) {
    return Get-Process | Where-Object {$_.Name -eq $ProcessName}
}

# C#
# https://maywork.net/computer/powershell-googlechrome-windows-size-reset/
Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class Win32Api {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    }
"@

$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height

# フォームの作成
$form = New-Object System.Windows.Forms.Form 
$form.Text = "こめなす"
#$form.Size = New-Object System.Drawing.Size(220,130)
$form.Size = New-Object System.Drawing.Size(220,180)
$form.Top = $default_window_pos_top
$form.Left = $default_window_pos_left
$form.StartPosition = "Manual"
$form.BackColor = "#606060"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.FormBorderStyle = "FixedSingle"
$form.Opacity = 1

# ラベル SKIP A
$LabelSkipA = New-Object System.Windows.Forms.Label
$LabelSkipA.Location = "15,15"
$LabelSkipA.Size = New-Object System.Drawing.Size(55,30)
$LabelSkipA.Text = "SKIP A"
$LabelSkipA.BackColor = "black"
$LabelSkipA.Forecolor = "yellow"
$LabelSkipA.TextAlign = "MiddleCenter"

# ラベル SKIP B
$LabelSkipB = New-Object System.Windows.Forms.Label
$LabelSkipB.Location = "80,15"
$LabelSkipB.Size = New-Object System.Drawing.Size(55,30)
$LabelSkipB.Text = "SKIP B"
$LabelSkipB.BackColor = "black"
$LabelSkipB.Forecolor = "yellow"
$LabelSkipB.TextAlign = "MiddleCenter"

# ラベル 30秒スキップ
$LabelSkipPrev = New-Object System.Windows.Forms.Label
$LabelSkipPrev.Location = "145,15"
$LabelSkipPrev.Size = New-Object System.Drawing.Size(55,30)
$LabelSkipPrev.Text = "+30"
$LabelSkipPrev.BackColor = "black"
$LabelSkipPrev.Forecolor = "yellow"
$LabelSkipPrev.TextAlign = "MiddleCenter"

# ラベル 30秒バック
$LabelSkipBack = New-Object System.Windows.Forms.Label
$LabelSkipBack.Location = "145,60"
$LabelSkipBack.Size = New-Object System.Drawing.Size(55,30)
$LabelSkipBack.Text = "-30"
$LabelSkipBack.BackColor = "black"
$LabelSkipBack.Forecolor = "yellow"
$LabelSkipBack.TextAlign = "MiddleCenter"


# Yesボタンの設定
$YesButton = New-Object System.Windows.Forms.Button
$YesButton.Location = New-Object System.Drawing.Point(15,60)
$YesButton.Size = New-Object System.Drawing.Size(55,30)
$YesButton.Text = "OPEN"
$YesButton.DialogResult = "Yes"
$YesButton.Flatstyle = "Popup"
$YesButton.backcolor = "black"
$YesButton.forecolor = "yellow"


# ラベル ログセット
$LabelLog = New-Object System.Windows.Forms.Label
$LabelLog.Location = "15,105"
$LabelLog.Size = New-Object System.Drawing.Size(55,30)
$LabelLog.Text = "SET LOG"
$LabelLog.BackColor = "black"
$LabelLog.Forecolor = "yellow"
$LabelLog.TextAlign = "MiddleCenter"


# PC TV Plus 再起動
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(80,105)
$OKButton.Size = New-Object System.Drawing.Size(55,30)
$OKButton.Text = "PC TV"
$OKButton.DialogResult = "OK"
$OKButton.Flatstyle = "Popup"
$OKButton.Backcolor = "black"
$OKButton.forecolor = "yellow"

# キャンセルボタンの設定
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(145,105)
$CancelButton.Size = New-Object System.Drawing.Size(55,30)
$CancelButton.Text = "SS"
$CancelButton.DialogResult = "Retry"
$CancelButton.Flatstyle = "Popup"
$CancelButton.backcolor = "black"
$CancelButton.forecolor = "yellow"


# ラベル
$LabelPause = New-Object System.Windows.Forms.Label
$LabelPause.Location = "80,60"
$LabelPause.Size = New-Object System.Drawing.Size(55,30)
$LabelPause.Text = "PAUSE" 
$LabelPause.BackColor = "black"
$LabelPause.Forecolor = "yellow"
$LabelPause.TextAlign = "MiddleCenter"

$form.Controls.Add($LabelSkipA)
$form.Controls.Add($LabelSkipB)
$form.Controls.Add($LabelSkipPrev)
$form.Controls.Add($LabelSkipBack)
$form.Controls.Add($LabelRecovery)
$form.Controls.Add($LabelPause)
$form.Controls.Add($LabelLog)

$file = $null

$FuncPCTVReStart = {
    sleep -Milliseconds 100
    $ps = getProcess( "Vnt" )
    if ( $ps -ne $null ) {
        $ret = $ps.CloseMainWindow()
    }
    sleep -Milliseconds 5000
    $obj = Start-Process -FilePath "Vnt.exe" -WorkingDirectory "C:\Program Files (x86)\Sony\PC TV Plus" -PassThru
}

$FuncScreenShot = {
    # スクリーンショット
    . Send-Keys "" $comment_viewer_app
    sleep -Milliseconds 100
    $bounds = [Drawing.Rectangle]::FromLTRB(0, 0, $screenWidth, $screenHeight)
    $file_name = (Get-Date).ToString("yyyyMMddHHmmss") + "screenshot.png"
    $save_path = $screenshot_dir + "\" + $file_name
    screenshot $bounds $save_path
}


Function file_open(){
    
    #アセンブリのロード
    Add-Type -AssemblyName System.Windows.Forms

    #ダイアログインスタンス生成
    $dialog = New-Object Windows.Forms.OpenFileDialog
    
    $dialog.Filter = "実況ログファイル(*.xml) | *.xml"
    $dialog.InitialDirectory = $log_dir
  
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
        $LabelLog.Forecolor = "red"
    } else {
        $LabelLog.Forecolor = "yellow"
    }
}

$FuncSkipA = {
    # 次のチャプターとAのコメントまで移動する
    . Send-Keys "^({RIGHT})" Vnt
    sleep -Milliseconds 4000
    . Send-Keys "a" $comment_viewer_app
    . Send-Keys "" Vnt
#    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($VntX, $VntY)
}

$FuncSkipB = {
    # 次のチャプターとBのコメントまで移動する
    . Send-Keys "^({RIGHT})" Vnt
    sleep -Milliseconds  4000
    . Send-Keys "b" $comment_viewer_app
    . Send-Keys "" Vnt
#    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point( $VntX, $VntY )
}

$FuncSkipPrev = {
    # 30秒飛ばし
    . Send-Keys "%({RIGHT})" Vnt
    . Send-Keys " 3" $comment_viewer_app
    # commenomiが進みすぎるときはこの値を増やす
    sleep -Milliseconds $prev_skip_wait
    . Send-Keys " " $comment_viewer_app
    sleep -Milliseconds 100
    . Send-Keys "" Vnt
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($VntX, $VntY)
}

$FuncSkipBack = {
    # 30秒戻し
    . Send-Keys "%({LEFT}{LEFT})" Vnt
    . Send-Keys " 2" $comment_viewer_app
    # commenomiが進みすぎるときはこの値を増やす
    sleep -Milliseconds $back_skip_wait
    . Send-Keys " " $comment_viewer_app
    sleep -Milliseconds 100
    . Send-Keys "" Vnt
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($VntX, $VntY)
}

$FuncPause = {
    # 一時停止
    . Send-Keys " " $comment_viewer_app
    . Send-Keys " " Vnt
}

$LabelSkipA.Add_Click($FuncSkipA)
$LabelSkipB.Add_Click($FuncSkipB)
$LabelSkipPrev.Add_Click($FuncSkipPrev)
$LabelSkipBack.Add_Click($FuncSkipBack)
$LabelPause.Add_Click($FuncPause)


$FuncFileOpenClick= {
    if ($_.Button -eq "Left" ) {
        . $FuncLogFileOpen
    } else {
        $script:file = $null
        $LabelLog.Forecolor = "yellow"
    }
}
$LabelLog.Add_MouseDown($FuncFileOpenClick)



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
            $ps = getProcess("Vnt")
            if ( $ps -ne $null ) {

                # PC TV Plusウィンドウ位置調整
                if ( $ps -ne $null ) {

#                    $top = 174 # PC TV Plusの縦位置
#                    $width = 1280 # PC TV Plusの幅

#                    $y = 260 # PC TV Plusの縦位置
#                    $width = 960 # PC TV Plusの幅
#                    $shift = 0 # 右にずらす

                    $top = 140 # PC TV Plusの縦位置
                    $width = 1440 # PC TV Plusの幅
                    $shift = 25 # 右にずらす

                    $left = (( 1920 - $width ) / 2) + $shift
                    $height = ($width / 16 * 9)
                    [Win32Api]::MoveWindow($ps.MainWindowHandle, $left, $top, $width, $height, $true) | Out-Null
                }

                if ( $file -ne $null ) {
                    $file_name = """" + $file + """"
                    $obj = Start-Process -FilePath ($comment_viewer_app + ".exe") -WorkingDirectory ("..\" + $comment_viewer_app + "\") -ArgumentList $file_name -PassThru
                    $file = $null
                    sleep -Milliseconds  2000
                } else {
                    # komenasneを開く
        #            $obj = Start-Process -FilePath "komenasne.exe" -WorkingDirectory ".\" -PassThru
                    $obj = Start-Process -FilePath "komenasne.exe" -WorkingDirectory "..\komenasne\" -PassThru
                }

                Wait-Process -InputObject $obj
                sleep -Milliseconds  4000
                # PC TV Plusを倍速再生にする
                if ( $enable_speed_up -eq $True ) {
                    . Send-Keys '+(^(G))' Vnt
                }
                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point( $VntX, $VntY )
                $count = 0
                while ($count -lt 50) {
                    sleep -Milliseconds 100
                    $ps = getProcess( $comment_viewer_app )
                    if ( $ps -ne $null ) {
                        $YesButton.Text = "CLOSE"
                        $count = 51
                    }
                    $count++
                }
                # ウィンドウ最大化
                if ( ( $comment_viewer_size_max -eq $True ) -and ( $ps -ne $null ) ) {
                    [Win32.NativeMethods]::ShowWindowAsync($ps.MainWindowHandle, 3) | Out-Null
                }
                $LabelLog.Forecolor = "yellow"
                $form.Opacity = $form_opacity
            }
        } else {
            #commenomiを閉じる
            sleep -Milliseconds 100
            $ps = getProcess( $comment_viewer_app )
            if ( $ps -ne $null ) {
                $ret = $ps.CloseMainWindow()
            }
            sleep -Milliseconds 1000
            . Send-Keys "{BACKSPACE}" Vnt
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
