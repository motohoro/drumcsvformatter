Option Explicit
'�����̃o�C�g�����J�E���g����bVBScript Tips
'http://www.kanaya440.com/contents/tips/vbs/001.html
Function lngCnLen(strVal)
    Dim i, strChr
    lngCnLen = 0
    If Trim(strVal) <> "" Then
        For i = 1 To Len(strVal)
            strChr = Mid(strVal, i, 1)
            '�Q�o�C�g�����́{�Q
            If (Asc(strChr) And &HFF00) <> 0 Then
                lngCnLen = lngCnLen + 2
            Else
                lngCnLen = lngCnLen + 1
            End If
        Next
    End If
End Function


'VBScript Tips (Tips0025) [������̈ꕔ�����擾����iASCII�j]
'http://www.whitire.com/vbs/tips0025.html
Function MidAscByte(ByVal strSjis, ByVal lngStartPos, ByVal lngGetByte, ByVal blnZenFlag)
' strSjis:      �؂�o��������
' lngStartPos:  �J�n�ʒu
' lngGetByte:   �擾�o�C�g���i"" ���� lngStartPos �ȍ~�S�āj
' blnZenFlag:   �S�p��������؂�ʒu�ŕ�������Ƃ��̓���
'               True= �X�y�[�X�ɕϊ�, False= ���̂܂܏o��
    Dim lngByte             ' �o�C�g��
    Dim lngLoop             ' ���[�v�J�E���^
    Dim strChkBuff          ' �m�F�p�o�b�t�@
    Dim strLastByte         ' �ŏI�o�C�g

    On Error Resume Next

    MidAscByte = ""
    If lngGetByte = "" Then
        ' �ő啶�������Z�b�g���Ă���
        lngGetByte = Len(strSjis) * 2
    End If
    lngGetByte = CLng(lngGetByte)

    ' �J�n�ʒu
    lngByte = 0
    For lngLoop = 1 To Len(strSjis)
        strChkBuff = Mid(strSjis, lngLoop, 1)
        If (Asc(strChkBuff) And &HFF00) = 0 Then
            lngByte = lngByte + 1
        Else
            lngByte = lngByte + 2
            ' �S�p�̂Q�o�C�g�ڂ��J�n�ʒu�̂Ƃ�
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

    ' �擾
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
        ' �I�[���S�p�P�o�C�g�ڂ̂Ƃ��B�Ӗ��Ȃ������i�΁j
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
'�ӉZ������: VBScript��VB.NET�̃h���b�O&�h���b�v���ꂽ�t�@�C���̃p�X���擾
'http://qri.seesaa.net/article/131805051.html
'�h���b�v���ꂽ�t�@�C���̃p�X��z��Ɏ擾
Dim myArray
Set myArray = WScript.Arguments
 
'�z��̓��e���o��
'For Each pass_str In myArray
'	MsgBox(pass_str)
'Next 
MsgBox(myArray(0))

Dim leader,leadersize
leader = "1,20160516,                    ,          ,                                        ,000,00,      ,00000,001218900,1502,1511,                              ,                ,                ,          ,                                        ,                                        ,0,                                        ,        ,        ,        ,          ,          ,                                        ,0 ,    ,"
leadersize = lngCnLen(leader)+1

Dim unitblankstr,unitstr,unitsize
unitblankstr=" ,              ,      ,                              ,0,0,0000000,0000000,0000000,000000,0,                                ,    ,0,0,0,          ,      ,                              ,0,                    ,"
unitstr="D,1111111000007 ,000000,����                          ,0,0,0000000,0000000,0000000,001000,0,                                ,g   ,0,0,1,          ,      ,                              ,0,                    ,"
unitsize= lngCnLen(unitstr)

'Dim cassettenum
Const cassettenum = 353
Dim cassettearray(353)'cassettenum)


'VBScript Tips (Tips0072) [�t�@�C���I�[�܂łP�s���ǂݍ���]
'http://www.whitire.com/vbs/tips0072.html

Dim objFSO      ' FileSystemObject
Dim objFile     ' �t�@�C���ǂݍ��ݗp

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
    Set objFile = objFSO.OpenTextFile(myArray(0))
'    Set objFile = objFSO.OpenTextFile("D:\�č��V��\Chozai.dat")
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
        WScript.Echo "�t�@�C���I�[�v���G���[: " & Err.Description
    End If
Else
    WScript.Echo "�G���[: " & Err.Description
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

'clip�R�}���h�𗘗p���ăN���b�v�{�[�h�ɕ�������R�s�[����VBScript | ���S�Ҕ��Y�^
'http://www.ka-net.org/blog/?p=1563
'cmd = "cmd /c ""echo " & clipdata & "| clip"""
'CreateObject("WScript.Shell").Run cmd, 0

Dim objFSO2      ' FileSystemObject
Dim objFile2     ' �t�@�C���������ݗp
Err.Number=0
Set objFSO2 = WScript.CreateObject("Scripting.FileSystemObject")
Dim appPath
appPath = objFSO2.GetParentFolderName(WScript.ScriptFullName)
If Err.Number = 0 Then
'VBScript�̎��s�����[�U/�����B�A�C�R�����_�u���N���b�N�Ŏ��s�ƁA���̃A... - Yahoo!�m�b��
'http://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q1163358876
    Set objFile2 = objFSO2.OpenTextFile(appPath & "\test.txt", 2, True,-2)
    If Err.Number = 0 Then
        objFile2.Write(Join(cassettearray,"==="))
        objFile2.Write(CR)
        objFile2.Write(clipdata)
        objFile2.Close
    Else
        WScript.Echo "�t�@�C���I�[�v���G���[: " & Err.Description
    End If
Else
    WScript.Echo "�G���[: " & Err.Description
End If

Set objFile2 = Nothing
Set objFSO2 = Nothing
