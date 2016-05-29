Option Explicit
'文字のバイト数をカウントする｜VBScript Tips
'http://www.kanaya440.com/contents/tips/vbs/001.html
Function lngCnLen(strVal)
    Dim i, strChr
    lngCnLen = 0
    If Trim(strVal) <> "" Then
        For i = 1 To Len(strVal)
            strChr = Mid(strVal, i, 1)
            '２バイト文字は＋２
            If (Asc(strChr) And &HFF00) <> 0 Then
                lngCnLen = lngCnLen + 2
            Else
                lngCnLen = lngCnLen + 1
            End If
        Next
    End If
End Function


'VBScript Tips (Tips0025) [文字列の一部分を取得する（ASCII）]
'http://www.whitire.com/vbs/tips0025.html
Function MidAscByte(ByVal strSjis, ByVal lngStartPos, ByVal lngGetByte, ByVal blnZenFlag)
' strSjis:      切り出す文字列
' lngStartPos:  開始位置
' lngGetByte:   取得バイト数（"" 時は lngStartPos 以降全て）
' blnZenFlag:   全角文字が区切り位置で分割するときの動作
'               True= スペースに変換, False= そのまま出力
    Dim lngByte             ' バイト数
    Dim lngLoop             ' ループカウンタ
    Dim strChkBuff          ' 確認用バッファ
    Dim strLastByte         ' 最終バイト

    On Error Resume Next

    MidAscByte = ""
    If lngGetByte = "" Then
        ' 最大文字数をセットしておく
        lngGetByte = Len(strSjis) * 2
    End If
    lngGetByte = CLng(lngGetByte)

    ' 開始位置
    lngByte = 0
    For lngLoop = 1 To Len(strSjis)
        strChkBuff = Mid(strSjis, lngLoop, 1)
        If (Asc(strChkBuff) And &HFF00) = 0 Then
            lngByte = lngByte + 1
        Else
            lngByte = lngByte + 2
            ' 全角の２バイト目が開始位置のとき
            If lngByte = lngStartPos Then
                If blnZenFlag = True Then
                    MidAscByte = " "
                Else
                    MidAscByte = Asc(strChkBuff) And &H00FF
                    If MidAscByte < 0 Then
                        MidAscByte = 256 + MidAscByte
                    End If
                    MidAscByte = ChrB(MidAscByte)
                End If
                lngLoop = lngLoop + 1
            End If
        End If
        If lngByte >= lngStartPos Then
            Exit For
        End If
    Next

    ' 取得
    lngByte = LenB(MidAscByte)
    If lngByte < lngGetByte Then
        For lngLoop = lngLoop To Len(strSjis)
            strChkBuff = Mid(strSjis, lngLoop, 1)
            MidAscByte = MidAscByte & strChkBuff
            If (Asc(strChkBuff) And &HFF00) = 0 Then
                lngByte = lngByte + 1
            Else
                lngByte = lngByte + 2
            End If
            If lngByte >= lngGetByte Then
                Exit For
            End If
        Next
    End If

    lngByte = LenAscByte(MidAscByte)
    If lngByte > lngGetByte Then
        ' 終端が全角１バイト目のとき。意味ないかも（笑）
        If blnZenFlag = True Then
            MidAscByte = Mid(MidAscByte, 1, Len(MidAscByte) - 1) & " "
        Else
            strLastByte = Fix((Asc(Right(MidAscByte, 1)) And &HFF00) / 256)
            If strLastByte < 0 Then
                strLastByte = 256 + strLastByte
            End If
            MidAscByte = Mid(MidAscByte, 1, Len(MidAscByte) - 1) & ChrB(strLastByte)
        End If
    End If
End Function

Function LeftAscByte(ByVal strSjis, ByVal lngGetByte, ByVal blnZenFlag)
    LeftAscByte = MidAscByte(strSjis, 1, lngGetByte, blnZenFlag)
End Function

Function RightAscByte(ByVal strSjis, ByVal lngGetByte, ByVal blnZenFlag)
    RightAscByte = StrReverse(strSjis)
    RightAscByte = MidAscByte(RightAscByte, 1, lngGetByte, blnZenFlag)
    RightAscByte = StrReverse(RightAscByte)
End Function


'==========================================================
'胡瓜メモ帳: VBScriptとVB.NETのドラッグ&ドロップされたファイルのパスを取得
'http://qri.seesaa.net/article/131805051.html
'ドロップされたファイルのパスを配列に取得
Dim myArray
Set myArray = WScript.Arguments
 
'配列の内容を出力
'For Each pass_str In myArray
'	MsgBox(pass_str)
'Next 
MsgBox(myArray(0))

Dim leader,leadersize
leader = "1,20160516,                    ,          ,                                        ,000,00,      ,00000,001218900,1502,1511,                              ,                ,                ,          ,                                        ,                                        ,0,                                        ,        ,        ,        ,          ,          ,                                        ,0 ,    ,"
leadersize = lngCnLen(leader)+1

Dim unitblankstr,unitstr,unitsize
unitblankstr=" ,              ,      ,                              ,0,0,0000000,0000000,0000000,000000,0,                                ,    ,0,0,0,          ,      ,                              ,0,                    ,"
unitstr="D,1111111000007 ,000000,＊＊                          ,0,0,0000000,0000000,0000000,001000,0,                                ,g   ,0,0,1,          ,      ,                              ,0,                    ,"
unitsize= lngCnLen(unitstr)

'Dim cassettenum
Const cassettenum = 353
Dim cassettearray(353)'cassettenum)


'VBScript Tips (Tips0072) [ファイル終端まで１行ずつ読み込む]
'http://www.whitire.com/vbs/tips0072.html

Dim objFSO      ' FileSystemObject
Dim objFile     ' ファイル読み込み用

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
    Set objFile = objFSO.OpenTextFile(myArray(0))
'    Set objFile = objFSO.OpenTextFile("D:\監査天秤\Chozai.dat")
    If Err.Number = 0 Then
        Do While objFile.AtEndOfStream <> True
            Dim line,oneunitstr
            line = objFile.ReadLine
'            WScript.Echo line 'objFile.ReadLine
            Dim j
            j=0
            While (leadersize+j*unitsize < lngCnLen(line))
                Dim unitcsvarray,unitnum
                oneunitstr =MidAscByte(line,leadersize+j*unitsize,unitsize,False) 'Get-SubStringBytes $l ($leadersize+$j*$unitsize) $unitsize
'If Msgbox(oneunitstr ,1)>1 Then
'WScript.Quit
'End If

                unitcsvarray = Split(oneunitstr , ",")
                If InStr(unitcsvarray(1) , "1111111") = 1 Then
                    unitnum = Mid(unitcsvarray(1),8,5)
'                    WScript.Echo unitnum
                    cassettearray(int(unitnum))=oneunitstr
                End If
                j=j+1
            WEnd
            
        Loop
        objFile.Close
    Else
        WScript.Echo "ファイルオープンエラー: " & Err.Description
    End If
Else
    WScript.Echo "エラー: " & Err.Description
End If

Set objFile = Nothing
Set objFSO = Nothing
Dim clipdata
clipdata=""
Dim TAB
TAB = Chr(9)
Dim CR
CR = Chr(13)

Dim h,k,m,p
For h=0 To 2-1
    For k=0 To 8-1
        Dim clipline
        clipline=leader
        For m=1 To 2
            For p=0 To 11-1
                Dim linedata
                linedata=""
                If cassettearray(h*176+k*2+m+p*16) = Null Then
                    linedata = unitstr
                Else
                    linedata = cassettearray(h*176+k*2+m+p*16)
                End If
'If h*176+k*2+m+p*16 = 21 Then
'If Msgbox("@"+linedata+"@" ,1)>1 Then
'WScript.Quit
'End If
'End If
                Dim linedata_l,linedata_c,linedata_r
                linedata_l = LeftAscByte(linedata,17,False)
                linedata_c = MidAscByte(linedata,18,37,False)
                Dim objRep
                Set objRep = New RegExp
                objRep.Pattern="\s+"
                objRep.Global = True
                linedata_c = objRep.replace(linedata_c,"##")
                linedata_r = RightAscByte(linedata,lngCnLen(linedata)-54,False)
                Dim clip_pre
                clip_pre=clipline
                clipline =clipline + linedata_l + linedata_c +TAB+ linedata_r
If (h*176+k*2+m+p*16 = 21) OR (h*176+k*2+m+p*16 = 3) Then
If Msgbox(clip_pre+"="+clipline+"?"+linedata+"@@"+linedata_l+"="+linedata_c+"=="+linedata_r+"?" ,1)>1 Then
'If Msgbox("?"+linedata+"@@"+linedata_l+"="+linedata_c+"=="+linedata_r+"?" ,1)>1 Then
WScript.Quit
End If
End If

'                clipline =clipline + (linedata_l + linedata_c + linedata_r)
            Next
        Next
        
'	    clipline = clipline + TAB
	    Dim q
	    For q=0 To 8-1 
	        clipline =clipline+ unitblankstr
	    Next
	        Dim objRep2
            Set objRep2 = New RegExp
            objRep2.Pattern="\s+"
            objRep2.Global = True 
            clipline = objRep2.Replace(clipline,TAB)
	    clipline = clipline + CR
	    clipdata =clipdata+ clipline
    Next
Next

WScript.Echo "Finished"'clipdata

'clipコマンドを利用してクリップボードに文字列をコピーするVBScript | 初心者備忘録
'http://www.ka-net.org/blog/?p=1563
'cmd = "cmd /c ""echo " & clipdata & "| clip"""
'CreateObject("WScript.Shell").Run cmd, 0

Dim objFSO2      ' FileSystemObject
Dim objFile2     ' ファイル書き込み用
Err.Number=0
Set objFSO2 = WScript.CreateObject("Scripting.FileSystemObject")
Dim appPath
appPath = objFSO2.GetParentFolderName(WScript.ScriptFullName)
If Err.Number = 0 Then
'VBScriptの実行時ユーザ/権限。アイコンをダブルクリックで実行と、他のア... - Yahoo!知恵袋
'http://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q1163358876
    Set objFile2 = objFSO2.OpenTextFile(appPath & "\test.txt", 2, True,-2)
    If Err.Number = 0 Then
        objFile2.Write(Join(cassettearray,"==="))
        objFile2.Write(CR)
        objFile2.Write(clipdata)
        objFile2.Close
    Else
        WScript.Echo "ファイルオープンエラー: " & Err.Description
    End If
Else
    WScript.Echo "エラー: " & Err.Description
End If

Set objFile2 = Nothing
Set objFSO2 = Nothing
