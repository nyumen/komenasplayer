# コメントプレイヤーのAPP名
$comment_viewer_exe_path = "..\commenomi\commenomi.exe"

# komenasneのPATH
$komenasne_path = "..\komenasne\komenasne.exe"

# PC TV PlusのインストールPATH
$pc_tv_plus_path = "C:\Program Files (x86)\Sony\PC TV Plus\Vnt.exe"

# 過去ログフォルダ
$log_dir = "..\komenasne\kakolog"

# スクリーンショット保存フォルダ
$screenshot_dir = ".\screenshot"

# コメントプレイヤー最大化 ( $True or $False )
$comment_viewer_size_max = $True

# デフォルト位置 1920x1080の場合
$default_window_pos_top = 850
$default_window_pos_left = 15


# PC TV Plusの大きさ指定 1920x1080の場合 ( top:上からの位置 width:幅 shift:横方向補正 )
$size_S_top = 261
$size_S_width = 1024
$size_S_shift = 0
$size_M_top = 174
$size_M_width = 1280
$size_M_shift = 0
$size_L_top = 142
$size_L_width = 1408
$size_L_shift = 0

# PC TV Plusを早見再生する ( $True or $False )
$enable_speed_up = $True

# 30秒送りでcommenomiが進みすぎるときはこの値を増やす（1000で1秒）
$prev_skip_wait = 500

# 30秒戻しでcommenomiが進みすぎるときはこの値を増やす（1000で1秒）
$back_skip_wait = 2000

# AやBへのスキップでcommenomiが進みすぎるときはこの値を増やす（1000で1秒）
$a_b_skip_wait = 3000

# フォームの透明度 ( 0.05～1 )
#$form_opacity = 0.5
$form_opacity = 0.05

# フォームクリックでフォームを隠す ( $True or $False )
$form_auto_hide = $True

# スクリーンショットの画質 ( 90～100 )
$screenshot_quality = 92