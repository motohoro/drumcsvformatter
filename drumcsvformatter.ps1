function Get-SubStringBytes([String]$Text, [int]$StartIndex = 0, [int]$Length = 0) {
$enc = [System.Text.Encoding]::Default
$bytes = $enc.GetBytes($Text)
return $enc.GetString($bytes, $StartIndex, $Length)
}
function Get-StringByteLen([String]$Text){
$enc = [System.Text.Encoding]::Default
$bytes = $enc.GetBytes($Text)
return $bytes.Length
}

$enc = [System.Text.Encoding]::Default
# http://www.atmarkit.co.jp/fwin2k/win2ktips/986psdialog/psdialog.html
# System.Windows.Formsアセンブリを有効化
[void][System.Reflection.Assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=Neutral, PublicKeyToken=b77a5c561934e089")

$todaystr = Get-Date -Format "yyyyMMdd"
$leader = "1,20160516,                    ,          ,                                        ,000,00,      ,00000,001218900,1502,1511,                              ,                ,                ,          ,                                        ,                                        ,0,                                        ,        ,        ,        ,          ,          ,                                        ,0 ,    ,"
$leadersize = Get-StringByteLen $leader

$unitblankstr = " ,              ,      ,                              ,0,0,0000000,0000000,0000000,000000,0,                                ,    ,0,0,0,          ,      ,                              ,0,                    ,"
$unitstr =      "D,1111111000007 ,000000,＊＊                          ,0,0,0000000,0000000,0000000,001000,0,                                ,g   ,0,0,1,          ,      ,                              ,0,                    ,"
$unitsize= Get-StringByteLen $unitstr


$cassettenum = 353
$cassettearray = New-Object object[] $cassettenum #num=352 ややこしいのでゼロは無視するため353個必要

# OpenFileDialogクラスをインスタンス化し、必要な情報を設定
$dialog = New-Object System.Windows.Forms.OpenFileDialog
#$dialog.Filter = "画像ファイル(*.PNG;*.JPG;*.GIF)|*.PNG;*.JPG;*.JPEG;*.GIF"
$dialog.Filter = "datファイル(*.dat)|*.dat"
$dialog.InitialDirectory = "C:\"
$dialog.Title = "ファイルを選択してください"
# ダイアログを表示
if($dialog.ShowDialog() -eq "OK"){
  # ［OK］ボタンがクリックされたら、選択されたファイル名（パス）を表示
  $dialog.FileName + " が選択されました。"
  $filepath = $dialog.FileName
  # http://capm-network.com/?tag=PowerShell%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E6%93%8D%E4%BD%9C
#    Get-Content $filepath | Foreach-Object {
#        echo $_
#    }

##http://win.just4fun.biz/PowerShell/%E3%83%86%E3%82%AD%E3%82%B9%E3%83%88%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%82%92%EF%BC%91%E8%A1%8C%E3%81%9A%E3%81%A4%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%82%80%E3%82%B5%E3%83%B3%E3%83%97%E3%83%AB%E3%82%B3%E3%83%BC%E3%83%89.html
    $enc = [Text.Encoding]::GetEncoding("Shift_JIS")
    $fh = New-Object System.IO.StreamReader($filepath, $enc)
    while (($l = $fh.ReadLine()) -ne $null) {
    #    Write-Host $i : $l
#        Write-Host $l
#        Write-Host $l.Substring($leadersize,$unitsize)
        $j=[int]0
        while($leadersize+$j*$unitsize -lt $l.Length){
            $oneunitstr =Get-SubStringBytes $l ($leadersize+$j*$unitsize) $unitsize
            $unitcsvarray = $oneunitstr -split ","
            if ($unitcsvarray[1] -like "1111111*"){
                $unitnum = [int]$unitcsvarray[1].SubString(7,5)
#                echo $unitnum
                $cassettearray[$unitnum]=$oneunitstr
            }
            $j++
        }
    #    $i++
    }

    $clipdata=""
    for($h =0;$h -lt 2;$h++){
        for ($k=0;$k -lt 8;$k++){
            $clipline= $leader #""
            for ($m=1;$m -le 2;$m++){
                for ($p=0;$p -lt 11;$p++){
#                    echo $h*176+$k*2+$m+$p*16= ($h*176+$k*2+$m+$p*16)
                    $linedata = ""
                    if ($cassettearray[$h*176+$k*2+$m+$p*16] -eq $null ){
#                        echo $unitstr
#                        $clipline += $unitstr
                        $linedata = $unitstr
                    }else{
#                        echo $cassettearray[$h*176+$k*2+$m+$p*16]
                        $linedata = $cassettearray[$h*176+$k*2+$m+$p*16]
                    }
                    $linedata_l = Get-SubStringBytes $linedata 0 17
                    $linedata_c = Get-SubStringBytes $linedata 17 37
                    $linedata_c = $linedata_c -replace "\s" , ""
                    $linedata_r = Get-SubStringBytes $linedata 54 ((Get-StringByteLen $linedata) -54)

                    $clipline += ($linedata_l + $linedata_c + "`t" + $linedata_r)

                }
            }
            $clipline+="`t"
            for($q=0;$q -lt 8;$q++){
                $clipline += $unitblankstr
            }

            # 文字列の操作 - Windows管理者のためのPowerShell http://powershell.wiki.fc2.com/wiki/%E6%96%87%E5%AD%97%E5%88%97%E3%81%AE%E6%93%8D%E4%BD%9C
            $clipline = $clipline -replace "\s+","`t"

            $clipline += "`n"
            $clipdata += $clipline
#            echo =============================
        }
    }
    ##PowerShellをはじめよう　～PowerShell入門～: PowerShellからクリップボードを扱う https://letspowershell.blogspot.jp/2016/03/powershell_12.html
    ##PowerShell の実行結果をクリップボードに入れたい - tech.guitarrapc.com http://tech.guitarrapc.com/entry/2013/07/19/200702
#    $OutputEncoding = [Text.Encoding]::Default #[console]::outputencoding
    $OutputEncoding = [console]::OutputEncoding;
#    $clipdata.Replace("," ,"`t") | clip
    $clipdata | clip
    #or
# アセンブリの読み込み
Add-Type -Assembly System.Windows.Forms
# 取得したアイテムをテキストとしてクリップボードへ送信
[Windows.Forms.Clipboard]::SetText($clipdata)

#    $excel = New-Object -ComObject Excel.Application
#    $excel.Visible = $False
#    $book = $excel.Workbooks.Open("C:\TEST\AAA.xlsx")
#    $sheet = $book.WorkSheets.item("Sheet1")
#    $sheet.Cells.Item(2,3) =  #Item(行, 列)。インデックスの番号は1から始まる。

}

#PowerShellをはじめよう　～PowerShell入門～: PowerShellでExcelを操作する　- シートの操作編 -
#https://letspowershell.blogspot.jp/2015/06/powershellexcel_11.html
#PowerShellでExcelの読み書き・ファイル作成
# http://blog.livedoor.jp/morituri/archives/54318641.html

#PowerShellファイル操作 CapmNetwork
#http://capm-network.com/?tag=PowerShell%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E6%93%8D%E4%BD%9C

#PowerShell: ◆Split演算子
#http://mtgpowershell.blogspot.jp/2010/11/split.html
#[PSv2]新しい演算子 -split と -join - PowerShell Scripting Weblog
#http://winscript.jp/powershell/198

#PowerShell の実行結果をクリップボードに入れたい - tech.guitarrapc.com
#http://tech.guitarrapc.com/entry/2013/07/19/200702

#[PowerShell] クリップボードに値をコピーする
#https://webcache.googleusercontent.com/search?q=cache:R_63Lf1jGxQJ:https://www.ipentec.com/document/document.aspx%3Fpage%3Dpowershell-copy-to-clipboard+&cd=1&hl=ja&ct=clnk&gl=jp

#PowerShellからクリップボードを扱う
# https://letspowershell.blogspot.jp/2016/03/powershell_12.html
#PowerShell ? 文字列を扱う | ITLAB51.COM
#http://www.itlab51.com/?p=5825

#powershell 入門 条件 IF | 技術的なこと、あれこれ
#http://www.sakutyuu.com/technology/?p=117

#PowerShellで日付書式にカスタム書式パターンを指定する - tech.guitarrapc.com
#http://tech.guitarrapc.com/entry/2013/02/09/030226
