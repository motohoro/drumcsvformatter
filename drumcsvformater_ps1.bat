@echo off
echo スクリプトを実行しています…
REM Windows7プリインストールのpowershellのversionは2 (wikipedia)
REM PowerShell2.0ではMTAモードがデフォルトだが、STAモードで実行しないと.NetDialogが表示されない http://funcs.org/907
powershell -v 2 -STA -NoProfile -ExecutionPolicy RemoteSigned -File .\drumcsvformatter.ps1
REM @powershell -NoProfile -ExecutionPolicy unrestricted -Command "Start-Process powershell.exe  -ArgumentList ""-file .\\drumcsvformatter.ps1"", "-Verb" ,"runas""
echo 完了しました！
pause > nul
exit

REM Powershellを楽に実行してもらうには - Qiita http://qiita.com/tomoko523/items/df8e384d32a377381ef9

