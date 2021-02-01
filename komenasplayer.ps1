# コメントプレイヤーのAPP名
$comment_viewer_app = "commenomi"

# コメントプレイヤー最大化 ( $True or $False)
$comment_viewer_size_max = $True

# ウィンドウのデフォルト位置 1920x1080の場合 上から800,左から20（好みに合わせて変更）
$default_window_pos_top = 800
$default_window_pos_left = 20

#30秒バックのオプション ( $True or $False ) スクリーンショット機能と排他
$enable_back_button = $True 

#PC TV Plusを早見再生する ( $True or $False )
$enable_speed_up = $True

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

$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height

# フォームの作成
$form = New-Object System.Windows.Forms.Form 
$form.Text = "こめなす"
$form.Size = New-Object System.Drawing.Size(240,150)
$form.Top = $default_window_pos_top
$form.Left = $default_window_pos_left
$form.StartPosition = "Manual"
$form.BackColor = "#606060"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.FormBorderStyle = "FixedSingle"
$form.Opacity = 1

# OKボタンの設定
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(20,20)
$OKButton.Size = New-Object System.Drawing.Size(55,30)
$OKButton.Text = "Skip A"
$OKButton.DialogResult = "OK"
$OKButton.Flatstyle = "Popup"
$OKButton.Backcolor = "black"
$OKButton.forecolor = "yellow"

# OKボタンの設定
$OK2Button = New-Object System.Windows.Forms.Button
$OK2Button.Location = New-Object System.Drawing.Point(90,20)
$OK2Button.Size = New-Object System.Drawing.Size(55,30)
$OK2Button.Text = "Skip B"
$OK2Button.DialogResult = "Ignore"
$OK2Button.Flatstyle = "Popup"
$OK2Button.Backcolor = "black"
$OK2Button.forecolor = "yellow"

# Yesボタンの設定
$YesButton = New-Object System.Windows.Forms.Button
$YesButton.Location = New-Object System.Drawing.Point(20,70)
$YesButton.Size = New-Object System.Drawing.Size(55,30)
$YesButton.Text = "Open"
$YesButton.DialogResult = "Yes"
$YesButton.Flatstyle = "Popup"
$YesButton.backcolor = "black"
$YesButton.forecolor = "yellow"

# キャンセルボタンの設定
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(90,70)
$CancelButton.Size = New-Object System.Drawing.Size(55,30)
$CancelButton.Text = "ﾛﾛ"
$CancelButton.DialogResult = "Retry"
$CancelButton.Flatstyle = "Popup"
$CancelButton.backcolor = "black"
$CancelButton.forecolor = "yellow"

# Abortボタンの設定
$AbortButton = New-Object System.Windows.Forms.Button
$AbortButton.Location = New-Object System.Drawing.Point(160,20)
$AbortButton.Size = New-Object System.Drawing.Size(55,30)
$AbortButton.Text = "+30"
$AbortButton.DialogResult = "Abort"
$AbortButton.Flatstyle = "Popup"
$AbortButton.backcolor = "black"
$AbortButton.forecolor = "yellow"

# Noneボタンの設定
$NoneButton = New-Object System.Windows.Forms.Button
$NoneButton.Location = New-Object System.Drawing.Point(160,70)
$NoneButton.Size = New-Object System.Drawing.Size(55,30)
if ( $enable_back_button -eq $True ) {
    $NoneButton.Text = "-30"
} else {
    $NoneButton.Text = "SS"
}
$NoneButton.DialogResult = "No"
$NoneButton.Flatstyle = "Popup"
$NoneButton.backcolor = "black"
$NoneButton.forecolor = "yellow"

# 最前面に表示：する
$form.Topmost = $True

# キーとボタンの関係
$form.AcceptButton = $OKButton
$form.CancelButton = $CancelButton

# ボタン等をフォームに追加
$form.Controls.Add($OKButton)
$form.Controls.Add($CancelButton)
$form.Controls.Add($OK2Button)
$form.Controls.Add($YesButton)
$form.Controls.Add($AbortButton)
$form.Controls.Add($NoneButton)

$VerbosePreference = 'Continue'

# フォームを表示させ、その結果を受け取る
$result = $form.ShowDialog()


# 結果による処理分岐
while ($result -ne "Cancel") {
    if ($result -eq "OK") {
        # 次のチャプターとAのコメントまで移動する
        . Send-Keys "^({RIGHT})" Vnt
        sleep -Milliseconds 4000
        . Send-Keys "A" $comment_viewer_app
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($VntX, $VntY)
    }

    if ($result -eq "Ignore") {
        # 次のチャプターとBのコメントまで移動する
        . Send-Keys "^({RIGHT})" Vnt
        sleep -Milliseconds  4000
        . Send-Keys "B" $comment_viewer_app
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point( $VntX, $VntY )
    }

    if ($result -eq "Yes") {
        if ($YesButton.Text -eq "Open") {
            if ( getProcess("Vnt") -ne $null ) {
                # komenasneを開く
    #            $obj = Start-Process -FilePath "komenasne.exe" -WorkingDirectory ".\" -PassThru
                $obj = Start-Process -FilePath "komenasne.exe" -WorkingDirectory "..\komenasne\" -PassThru
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
                        $YesButton.Text = "Close"
                        $count = 51
                    }
                    $count++
                }
                # ウィンドウ最大化
                if ( $comment_viewer_size_max -eq $True ) {
                    [Win32.NativeMethods]::ShowWindowAsync($ps.MainWindowHandle, 3) | Out-Null
                }
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
            $YesButton.Text = "Open"
        }
    }

    if ($result -eq "Retry") {
        # 一時停止
        . Send-Keys " " $comment_viewer_app
        . Send-Keys "（ ）" Vnt
    }

    if ($result -eq "Abort") {
        # 30秒飛ばし
        . Send-Keys " 3" $comment_viewer_app
        . Send-Keys "%({RIGHT})" Vnt
        # commenomiが進みすぎるときはこの値を増やす
        sleep -Milliseconds 2000
        . Send-Keys " " $comment_viewer_app
        sleep -Milliseconds 100
        . Send-Keys "" Vnt
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($VntX, $VntY)
    }

    if ($result -eq "No") {
        if ( $enable_back_button -eq $True ) {
            # 30秒戻し
            . Send-Keys " 2" $comment_viewer_app
            . Send-Keys "%({LEFT}{LEFT})" Vnt
            # commenomiが進みすぎるときはこの値を増やす
            sleep -Milliseconds 1500
            . Send-Keys " " $comment_viewer_app
            sleep -Milliseconds 100
            . Send-Keys "" Vnt
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($VntX, $VntY)
        } else {
            # スクリーンショット
            . Send-Keys "" $comment_viewer_app
            sleep -Milliseconds 100
            $bounds = [Drawing.Rectangle]::FromLTRB(0, 0, $screenWidth, $screenHeight)
            $file_name = (Get-Date).ToString("yyyyMMddHHmmss") + "screenshot.png"
            screenshot $bounds $file_name
        }
    }


    # フォームを表示させ、その結果を受け取るA
    $result = $form.ShowDialog()
}
