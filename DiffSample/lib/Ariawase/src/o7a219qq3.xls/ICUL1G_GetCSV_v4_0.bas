Attribute VB_Name = "ICUL1G_GetCSV_v4_0"
Option Explicit

'  2013/10/17 H.Arikawa GetCVS_v4_0を自動化用にカスタマイズして入れ込み。
'  2013/10/23 H.Arikawa バグ修正&ディレクトリ構成1本化

Public OffsetFileName As String
Public TimingFileName As String
Public VoltageoffsetFileName As String
Public OptFileName As String
Public Clock_VoltageOffsetFileName As String
Public Power_Supply_VoltageoffsetFileName As String
Public Flg_CsvFileFailSafe As Boolean

Private Enum prCSVEOFEnum
    BothEOF = 0
    OneSEOF = 1
    LineRem = 2
End Enum

Public strPassForCSV As String

'Toolのディレクトリ分岐ケアの為、追加。　2013/10/17 H.Arikawa
'#Const EEE_AUTO_JOB_LOCATE = @EEE_AUTO_JOB_LOCATE@      '1:長崎200mm,2:長崎300mm,3:熊本
#Const EEE_AUTO_JOB_LOCATE = 2      '1:長崎200mm,2:長崎300mm,3:熊本             'CaptureModuleの変数置換を導入次第変更する。

Public intCntCsv As Integer             '共通変数の為、条件付きコンパイル引数の外へ移動。 2014/01/08 H.Arikawa
#If EEE_AUTO_JOB_LOCATE = 1 Or EEE_AUTO_JOB_LOCATE = 2 Then
'********** For CSV_CTRL 2013/4/24 M.Imamura   ****************************
Public Const CSV_VBS_PATH_SEVER = "G:\jobs\CSV_CTRL\CSV_CTRL.vbe"
Public Const CSV_VBS_PATH_LOCAL = "C:\CSV_CTRL.vbe"
'**************************************************************************
#Else
'********** For CSV_CTRL 2013/4/24 M.Imamura   ****************************
Public Const CSV_VBS_PATH_SEVER = "F:\Job\CIS\2PC\Debug_Fol\CSV_CTRL\CSV_CTRL.vbe"
Public Const CSV_VBS_PATH_LOCAL = "C:\CSV_CTRL.vbe"
'**************************************************************************
#End If

Public Sub GetCsvFileName()
    
    If Flg_Simulator = 1 Then
        strPassForCSV = ".\"    'シミュレータで動作させる時は、外部ファイルはJOBと同じ場所に置く
    Else
        strPassForCSV = ".\parameter\" & ComputerName & "\"
    End If
    
    '===== CSV-file FailSafe =====
    Flg_CsvFileFailSafe = True                                                      '外部ファイル管理フラグ
    '=============================

    '********** For CSV_CTRL 2013/4/24 M.Imamura   ****************************
    intCntCsv = 0
    ThisWorkbook.Worksheets("Read CSV").Buttons.Delete
    '**************************************************************************

    '===== 暫定策　最終的には長崎、熊本のディレクトリ構成ルールを1本化する。 =====
    OffsetFileName = sub_CSV_CTRL(strPassForCSV & "offset_" & Format(CStr(Sw_Node), "000") & ".csv")
'    TimingFileName = sub_CSV_CTRL(strPassForCSV & "timing_" & Format(CStr(Sw_Node), "000") & ".csv")
'    VoltageoffsetFileName = sub_CSV_CTRL(strPassForCSV & "pps_" & Format(CStr(Sw_Node), "000") & ".csv")
    OptFileName = sub_CSV_CTRL(strPassForCSV & "opt_" & Format(CStr(Sw_Node), "000") & ".csv")
    Power_Supply_VoltageoffsetFileName = sub_CSV_CTRL(strPassForCSV & "power_supply_" & Format(CStr(Sw_Node), "000") & ".csv")
    Clock_VoltageOffsetFileName = sub_CSV_CTRL(strPassForCSV & "clock_" & Format(CStr(Sw_Node), "000") & ".csv")
    
End Sub

Public Sub ReadOffsetFile()
    Dim strArg, temp5 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim FileNo As Integer

    '======== CSV File Exist Check.
    If Dir(OffsetFileName) = "" Then
        Flg_CsvFileFailSafe = False
        MsgBox "Error! [" & OffsetFileName & "] is Not Found!"
        Exit Sub
    End If

    '=======Start_ReadOffsetFile============
    Worksheets("Read CSV").Range("A1:AZ10000").Clear     'Clear Sheet
    
    FileNo = FreeFile

    Open OffsetFileName For Input As #FileNo              'CSV File OPEN

    On Error GoTo CloseFile                             'Error Check

    i = 0
    Do Until EOF(FileNo)                                     'Data Input to buffer
        Line Input #FileNo, temp5
        i = i + 1
        strArg = Split(temp5, ",")
        For j = 0 To UBound(strArg)
            Worksheets("Read CSV").Cells(i, j + 1) = strArg(j)                 'Data Input to sheet
        Next j
    Loop

    Close #FileNo                                       'CSV File Close

    On Error GoTo 0

Offset_csv_end:

Exit Sub

CloseFile:

    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    Close #FileNo
    MsgBox ("File Open Error! Please Check Offset File")
    GoTo Offset_csv_end
    
    '=======End_ReadOffsetFile============


End Sub

Public Sub ReadTimingFile()
    Dim strArg, temp5 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim FileNo As Integer


    '=======Start_ReadOffsetFile============
    Worksheets("Read CSV").Range("A1:AZ1000").Clear     'Clear Sheet
 
    FileNo = FreeFile

    Open TimingFileName For Input As #FileNo              'CSV File OPEN

    On Error GoTo CloseFile                             'Error Check

    i = 0
    Do Until EOF(FileNo)                                     'Data Input to buffer
        Line Input #FileNo, temp5
        i = i + 1
        strArg = Split(temp5, ",")
        For j = 0 To UBound(strArg)
            Worksheets("Read CSV").Cells(i, j + 1) = strArg(j)                 'Data Input to sheet
        Next j
    Loop

    Close #FileNo                                       'CSV File Close

    On Error GoTo 0

Offset_csv_end:

Exit Sub

CloseFile:

    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    Close #FileNo
    MsgBox ("File Open Error! Please Check Timing File")
    GoTo Offset_csv_end
    
    '=======End_ReadOffsetFile============



End Sub

Public Sub ReadVoltageoffsetFile()
    Dim strArg, temp5 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim FileNo As Integer


    '=======Start_ReadOffsetFile============
    Worksheets("Read CSV").Range("A1:AZ1000").Clear     'Clear Sheet
    
    FileNo = FreeFile

    Open VoltageoffsetFileName For Input As #FileNo              'CSV File OPEN

    On Error GoTo CloseFile                             'Error Check

    i = 0
    Do Until EOF(FileNo)                                     'Data Input to buffer
        Line Input #FileNo, temp5
        i = i + 1
        strArg = Split(temp5, ",")
        For j = 0 To UBound(strArg)
            Worksheets("Read CSV").Cells(i, j + 1) = strArg(j)                 'Data Input to sheet
        Next j
    Loop

    Close #FileNo                                       'CSV File Close

    On Error GoTo 0

Offset_csv_end:

Exit Sub

CloseFile:

    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    Close #FileNo
    MsgBox ("File Open Error! Please Check Voltage File")
    GoTo Offset_csv_end
    
    '=======End_ReadOffsetFile============


End Sub
Public Sub ReadOptFile()
    Dim strArg, temp5 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim FileNo As Integer

    
    '======== CSV File Exist Check.
    If Dir(OptFileName) = "" Then
        Flg_CsvFileFailSafe = False
        MsgBox "Error! [" & OptFileName & "] is Not Found!"
        Exit Sub
    End If

    '=======Start_ReadOffsetFile============
    Worksheets("Read CSV").Range("A1:AZ1000").Clear     'Clear Sheet
    
    FileNo = FreeFile

    Open OptFileName For Input As #FileNo              'CSV File OPEN

    On Error GoTo CloseFile                             'Error Check

    i = 0
    Do Until EOF(FileNo)                                     'Data Input to buffer
        Line Input #FileNo, temp5
        i = i + 1
        strArg = Split(temp5, ",")
        For j = 0 To UBound(strArg)
            Worksheets("Read CSV").Cells(i, j + 1) = strArg(j)                 'Data Input to sheet
        Next j
    Loop

    Close #FileNo                                       'CSV File Close

    On Error GoTo 0

Offset_csv_end:
    
Exit Sub

CloseFile:

    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    Close #FileNo
    MsgBox ("File Open Error! Please Check Opt File")
    GoTo Offset_csv_end
    
    '=======End_ReadOffsetFile============


End Sub

Private Sub CheckCsvFile(ByVal fileName As String, ByRef Flg_CsvFileFailSafe As Boolean)

    Dim i As Long
    Dim FileNo As Integer, BackUpFileNo As Integer
    
    Dim lCSVEOFChkFlg As Integer

    On Error GoTo ErrorDetected

'    FileNo = FreeFile
'    Open FileName For Input As #FileNo
    
    '====== Debug Job Skip Proc ==========.
    If UCase(Mid(NormalJobName, 3, 1)) = "Z" Then Exit Sub
    
    '======== CSV File Exist Check.
    If Dir(fileName) = "" Then
        Flg_CsvFileFailSafe = False
        MsgBox "Error! [" & fileName & "] is Not Found!"
        Exit Sub
    End If
    
    Dim BackUpFilePath As String, tmpBackUpFilePath As Variant
    tmpBackUpFilePath = Split(fileName, "\")
    
    BackUpFilePath = tmpBackUpFilePath(0)
    For i = 1 To (UBound(tmpBackUpFilePath) - 1)
        BackUpFilePath = BackUpFilePath & "\" & tmpBackUpFilePath(i)
    Next i
    BackUpFilePath = BackUpFilePath & "\BackUp\"

    Dim BackUpFileName As String, tmpFileName As String
    tmpFileName = Left(tmpBackUpFilePath(UBound(tmpBackUpFilePath)), Len(tmpBackUpFilePath(UBound(tmpBackUpFilePath))) - 4)
    
    '================= Exit Proc in case Back Up File is Not Found.
    If FindTargetFile(BackUpFilePath, tmpFileName, BackUpFileName) = False Then
        MsgBox "Error! " & tmpFileName & "'s Back Up File is Invalid."
        Exit Sub
    End If
    
    BackUpFileName = BackUpFilePath & BackUpFileName
    
    FileNo = FreeFile
    Open fileName For Input As #FileNo
    BackUpFileNo = FreeFile
    Open BackUpFileName For Input As #BackUpFileNo

    Dim tmpFile As String, tmpBackUpFile As String
    
    For i = 0 To 2 ^ 31 - 1
    
        Line Input #FileNo, tmpFile
        Line Input #BackUpFileNo, tmpBackUpFile
        
        lCSVEOFChkFlg = ConvertXorToInt(EOF(FileNo), EOF(BackUpFileNo))
        ' Route for lCSVEOFChkFlg Value.
        ' BothEOF(0) : Both EOF
        ' OneSEOF(1) : Only One File EOF
        ' LineRem(2) : Line Remained
        Select Case lCSVEOFChkFlg
            Case prCSVEOFEnum.BothEOF
                'Empty-file Reject Until Find "Location:".
                If i <= 2 Then
                    fileName = "empty-file [" & fileName & "]"
                    GoTo ErrorDetected
                End If
                Exit For
            Case prCSVEOFEnum.OneSEOF
                Flg_CsvFileFailSafe = False
                MsgBox "Error! [" & fileName & "]'s Line Number does'nt Match [" _
                                & BackUpFileName & "] "
                Exit For
            Case prCSVEOFEnum.LineRem
                If tmpFile <> tmpBackUpFile Then
                    Flg_CsvFileFailSafe = False
                    MsgBox "BackUp CSV File Miss Match Error! @" & fileName
                    Exit For
                End If
        End Select
        
    '===== CSV-file FailSafe =====
        Dim tmpFile_emptychkline As String, tmpFile_emptychkword As Variant
        If i = 2 Then
            tmpFile_emptychkline = tmpFile
            tmpFile_emptychkword = Split(tmpFile_emptychkline, ",")
            If tmpFile_emptychkword(1) <> "location:" And tmpFile_emptychkword(1) <> "UserDelay TAP" Then
                fileName = "empty-file [" & fileName & "]"
                GoTo ErrorDetected
            End If
        End If
    '=============================
        
        If i = 2 ^ 31 - 1 Then
            Flg_CsvFileFailSafe = False
            MsgBox "Error! CSV File is so Long! [2^31 Column]"
        End If
    Next i

    Close #FileNo
    Close #BackUpFileNo
    Exit Sub

ErrorDetected:
    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    Close #FileNo
    Close #BackUpFileNo
    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    MsgBox "BackUp CSV File Miss Match Error! @" & fileName
    Flg_CsvFileFailSafe = False
    
End Sub

Private Function FindTargetFile(ByVal Path As String, ByVal srcFileName As String, _
                        ByRef dstFileName As String) As Boolean
    Dim i As Long
    Dim buf As String
    Dim TmpSrcFileAddWild As String
    
    Dim NowRevNum As Integer
    Dim MaxRevNum As Integer
    
    FindTargetFile = False
    
On Error GoTo JumpEnd
    
    TmpSrcFileAddWild = Path & srcFileName & "_*"
    buf = Dir(TmpSrcFileAddWild, vbNormal)
    
    If buf = "" Then
        Flg_CsvFileFailSafe = False
        Exit Function
    End If
    
    MaxRevNum = 0
    Do While buf <> ""
        NowRevNum = CheckCSVBackUpFileAndGetRev(srcFileName & "_", buf)
        If MaxRevNum <= NowRevNum Then
            dstFileName = buf
            MaxRevNum = NowRevNum
        End If
        buf = Dir
    Loop
    
    '============ Set Retrun Value "True" in case Back up File is Found Safty.
    FindTargetFile = True
    Exit Function
    
JumpEnd:
    FindTargetFile = False
    Flg_CsvFileFailSafe = False
End Function

Private Function CheckCSVBackUpFileAndGetRev(ByVal InRefFName As String, ByVal InBackUpFName As String) As Integer

    Dim i As Integer

    Dim lRsFileName As String
    Dim lTmpAscCode As Integer
    Dim lRevNumStr As String

    CheckCSVBackUpFileAndGetRev = -1

    '============= lRsFileName Shuld be "XXX.CSV"
    lRsFileName = Right(InBackUpFName, Len(InBackUpFName) - Len(InRefFName))

    '============= lRsFileName Words Length Check. Shuld be 7Letters. "XXX.CSV"
    If Len(lRsFileName) <> 7 Then Exit Function

    '============= lRsFileName Exp Check. Shuld be ".CSV"
    If UCase(Right(lRsFileName, 4)) <> ".CSV" Then Exit Function
    
    '============= lRsFileName Revision Number Check. Shuld be 3 Letters "XXX"
    lRevNumStr = Left(lRsFileName, 3)
    
    '========= Ascii Code Number Check . 0-9 => 48-57 in Ascii Code.
    For i = 1 To 3
        lTmpAscCode = Asc(Mid(lRevNumStr, i, 1))
        If lTmpAscCode < 48 Or 57 < lTmpAscCode Then Exit Function
    Next i
    
    '========= Set Return Value Revison Number.
    CheckCSVBackUpFileAndGetRev = CInt(lRevNumStr)

End Function

Private Function ConvertXorToInt(ByVal InXLgc As Boolean, ByVal InYLgc As Boolean) As Integer
    ' X Y   Ret
    ' T T BothEOF : 0(F)
    ' T F OneSEOF : 1(T)
    ' F T OneSEOF : 1(T)
    ' F F LineRem : 2(F)
    ' Note. X Xor Y = (X Or Y) And (X NAnd Y)
    
    Dim i As Integer
    Dim lXorBool As Boolean

    lXorBool = InXLgc Xor InYLgc
    
    If lXorBool = True Then
        ConvertXorToInt = prCSVEOFEnum.OneSEOF
    Else
        If InXLgc = True Then
            ConvertXorToInt = prCSVEOFEnum.BothEOF
        Else
            ConvertXorToInt = prCSVEOFEnum.LineRem
        End If
    End If

End Function

Public Sub AllCSVCheckSub()
    Dim intMipiSetFor As Integer
    
    '===== CSV-file FailSafe =====
    '############### ParameterFile-Check Failsafe ###############
    '    Call StartTime
        Call CheckCsvFile(OffsetFileName, Flg_CsvFileFailSafe)
        Call CheckCsvFile(OptFileName, Flg_CsvFileFailSafe)
        Call CheckCsvFile(Power_Supply_VoltageoffsetFileName, Flg_CsvFileFailSafe)
        Call CheckCsvFile(Clock_VoltageOffsetFileName, Flg_CsvFileFailSafe)
        For intMipiSetFor = 0 To UBound(MipiSetFor1G)
            If MipiSetFor1G(intMipiSetFor).MipiKeyName <> "" Then
                Call CheckCsvFile(sub_CSV_CTRL(strPassForCSV & MipiSetFor1G(intMipiSetFor).MipiKeyName & "_" & Format(CStr(Sw_Node), "000") & ".csv"), Flg_CsvFileFailSafe)
            End If
        Next
    '    Call StopTimer
    '############################################################
    
End Sub

Public Sub ManualCSVCheckSub()
    
    '====== Debug Job Skip Check ==========.
    If UCase(Mid(NormalJobName, 3, 1)) = "Z" Then
        Flg_CsvFileFailSafe = False
        MsgBox "JobName's 3rd. Word is ""Z"". Can't Check CSV."
        Debug.Print "  [CSV Check NG! @" & Now() & "]"
        Exit Sub
    End If

    Call JobEnvInit
    Call GetCsvFileName
    Call ICUL1G_Parameter_Def
    Call AllCSVCheckSub

    If Flg_CsvFileFailSafe = True Then
        Debug.Print "  [CSV Check All OK! @" & Now() & "]"
    Else
        MsgBox "CSV Check NG!"
        Debug.Print "  [CSV Check NG! @" & Now() & "]"
    End If

End Sub

'********** For CSV_CTRL 2013/4/24 M.Imamura   ****************************
Public Function sub_CSV_CTRL(ByVal CsvFileName_buf As String, Optional FlgShowOnly As String = "No") As String
    Dim strVbsPathLine As String
    Dim objVbs As Object
    Dim objVbsReturn As Variant
        
    sub_CSV_CTRL = CsvFileName_buf
    
    If Flg_Simulator = 1 Then
        Exit Function
    End If
    
    '
    If FlgShowOnly = "Yes" And First_Exec = 0 Then
        MsgBox "Run this JOB To use Button ", vbCritical, "CSV_CTRL"
        Exit Function
    End If
    
    'サーバからvbsファイルをコピー
    On Error Resume Next
    Kill CSV_VBS_PATH_LOCAL
    On Error GoTo 0
    If sub_PalsFileCopy(CSV_VBS_PATH_SEVER, CSV_VBS_PATH_LOCAL) = False Then
        MsgBox "Copy Failed... " & CSV_VBS_PATH_SEVER, vbCritical, "CSV_CTRL"
        Exit Function
    End If
    
    'ドライブ情報が無い場合、絶対Pathに変換
    If InStr(1, CsvFileName_buf, ":\") = 0 And InStr(1, CsvFileName_buf, "\\") = 0 Then
        CsvFileName_buf = ThisWorkbook.Path & Right(CsvFileName_buf, Len(CsvFileName_buf) - 1)
    End If
    'コマンドライン引数を設定
    strVbsPathLine = CSV_VBS_PATH_LOCAL & " " & ThisWorkbook.Name & " " & CsvFileName_buf & " " & Flg_AutoMode & " " & FlgShowOnly & " " & ThisWorkbook.Path
    
On Error GoTo sub_CSV_CTRL_error
    'VBSの起動
    Set objVbs = CreateObject("WScript.Shell")
    objVbsReturn = objVbs.Run(strVbsPathLine, 0, True)  '0-非表示,True-スクリプトの終了を待つ
    Set objVbs = Nothing
    If FlgShowOnly = "Yes" Then
        Exit Function
    End If
    
    intCntCsv = intCntCsv + 1
    
    'CSVファイル名を更新する
    sub_CSV_CTRL = ThisWorkbook.Worksheets("Read CSV").Cells(1, 1).Value

    'CSV表示ボタンを作成、引数をセットする
    With Worksheets("Read CSV").Buttons.Add(200, 50 + (intCntCsv - 1) * 30, 150, 30)
        .Name = "ShowCSV" & CStr(intCntCsv)
        .OnAction = "'sub_CSV_CTRL " & """" & ThisWorkbook.Worksheets("Read CSV").Cells(1, 1).Value & """,""" & "Yes" & """'"
        .Characters.Text = "Show " & ThisWorkbook.Worksheets("Read CSV").Cells(2, 1).Value
    End With
    
    'オフセットがLIMITNGの場合フラグを立てる
    If ThisWorkbook.Worksheets("Read CSV").Cells(3, 1).Value = "LIMITNG" Then
        Flg_CsvFileFailSafe = False
    End If

    Exit Function
sub_CSV_CTRL_error:
    MsgBox "CSV_CTRL Script Error, NotFound vbsFile or Abort By User!!", vbCritical, "CSV_CTRL"
End Function

