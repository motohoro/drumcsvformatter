@echo off
echo �X�N���v�g�����s���Ă��܂��c
REM Windows7�v���C���X�g�[����powershell��version��2 (wikipedia)
REM PowerShell2.0�ł�MTA���[�h���f�t�H���g�����ASTA���[�h�Ŏ��s���Ȃ���.NetDialog���\������Ȃ� http://funcs.org/907
powershell -v 2 -STA -NoProfile -ExecutionPolicy RemoteSigned -File .\drumcsvformatter.ps1
REM @powershell -NoProfile -ExecutionPolicy unrestricted -Command "Start-Process powershell.exe  -ArgumentList ""-file .\\drumcsvformatter.ps1"", "-Verb" ,"runas""
echo �������܂����I
pause > nul
exit

REM Powershell���y�Ɏ��s���Ă��炤�ɂ� - Qiita http://qiita.com/tomoko523/items/df8e384d32a377381ef9

