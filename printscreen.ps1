#*&---------------------------------------------------------------------*
#*& Script printScreen.ps1
#*& 画面スクリーンショット取得
#*&---------------------------------------------------------------------*

param(
[parameter(Mandatory=$False,HelpMessage="SleepTime(Sec)")]
[int]$SleepTime,
[parameter(Mandatory=$False,HelpMessage="ファイル保管場所")]
[string]$SaveDir=$PWD.Path,
[parameter(Mandatory=$False,HelpMessage="isMsgbox")]
[ValidateSet("True","False")]
$isMsgBox=$true,
[parameter(Mandatory=$False,HelpMessage="isMouseCursor")]
[ValidateSet("True","False")]
$isMouseCursor=$true,
[parameter(Mandatory=$False,HelpMessage="isMouseCursor")]
[ValidateSet("JPEG","PNG","BMP","TIFF","GIF")]
[string]$Format="JPEG"
)

Process{

    #ショートカットファイルのショートカットから
    #powershell -file <<script.ps1>> [<Args>] の形で引数を渡すときboolean型で引数を渡す方法がなさそう
    #しかたなく引数を任意で受け取ってからbooleanに変換する。
    #(ショートカットファイルのショートカットから引数に  1 とか True　とか入れると
    #文字列で認識してそれをboolean型に突っ込もうとしてエラーになる)
    $isMsgBox = [System.Convert]::ToBoolean($isMsgBox)
    $isMouseCursor = [System.Convert]::ToBoolean($isMouseCursor)

    #.NetFrameworkのFormクラスを読み込み
    Add-Type -AssemblyName System.Windows.Forms
            
    #スリープタイム処理
    if ( $PSBoundParameters.ContainsKey("SleepTime") ){
        Start-Sleep $SleepTime
    }

    #画面スクリーンショット
        
    #DPI Scalingの取得
    $DPISetting = (Get-ItemProperty 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name AppliedDPI).AppliedDPI
    switch ($DPISetting){
        96 {$ActualDPI = 100}
        120 {$ActualDPI = 125}
        144 {$ActualDPI = 150}
        192 {$ActualDPI = 200}
    }

    [float]$DisplayScale=($ActualDPI /100)
 
    #プライマリのディスプレイサイズ取得
    $PrimarySize = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize
    [int]$Width = ($PrimarySize.Width * $DisplayScale)
    [int]$Height = ($PrimarySize.Height * $DisplayScale)

    #Bitmapクラスのインスタンス化
    $Bitmap = New-Object System.Drawing.Bitmap($Width,$Height)
    #Graphicsクラスのインスタンス化
    $Graphic = [System.Drawing.Graphics]::FromImage($Bitmap)

    #画像の始点作成
    $StartPoint = (New-Object System.Drawing.Point(0,0))

    #転送元の四角形の左上隅の点,転送先の四角形の左上隅の点,転送される領域のサイズ
    $Graphic.CopyFromScreen($StartPoint, $StartPoint, $Bitmap.Size)

    #マウスカーソルを画像に追加
    IF($isMouseCursor){
        #矢印のカーソル取得
        $ArrowCursor = [System.Windows.Forms.Cursors]::Arrow
        #現在のカーソル位置取得
        $cursorPosition=[System.Windows.Forms.Cursor]::Position
        #カーソル位置補正（スケーリング対応）
        [float]$PositionX=(($CursorPosition.x)*$DisplayScale)
        [float]$PositionY=(($CursorPosition.y)*$DisplayScale)
        #矩形生成
        $Rectangle = New-Object -TypeName System.Drawing.Rectangle($PositionX,$PositionY,0,0)
        #カーソル書込み
        $ArrowCursor.Draw($Graphic,$Rectangle)
    }
    #オブジェクト破棄
    $Graphic.Dispose()
    #EncoderParametersクラスのインスタンス生成
    $EncodeParams = New-Object Drawing.Imaging.EncoderParameters
    #EncoderParameterクラスのインスタンス生成
    $EncodeParams.Param[0] = New-Object Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]100)  
    #指定フォーマットの情報取得
    $Codec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.FormatDescription -eq $Format }
    #ファイル名用　年　月　日　時分秒ミリ秒
    $DateTime = Get-Date -Format "yyyyMMdd_HHmmssfff"
    #ファイル出力先の作成（フルパス）
    [string]$SaveFullPath = join-path $SaveDir "$DateTime.$Format"

    #ファイル出力
    $Bitmap.Save($SaveFullPath, $Codec, $EncodeParam)

    #メッセージボックス表示
    IF($isMsgBox){
        $MsgBox=[System.Windows.Forms.MessageBox]
        $buttons=[System.Windows.Forms.MessageBoxButtons]::OK
        $icon=[System.Windows.Forms.MessageBoxIcon]::information
        $defaultButton=[System.Windows.Forms.MessageBoxDefaultButton]::Button1
        $options=[System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly
        $msgbox::show("OK","Screenshot",$buttons,$icon,$defaultbutton,$options)
    }

}
