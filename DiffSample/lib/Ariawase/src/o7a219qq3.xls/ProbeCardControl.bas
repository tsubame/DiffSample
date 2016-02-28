Attribute VB_Name = "ProbeCardControl"
'概要:
'　プローブカードコンタクト回数を管理する
'
'目的:
'　ZDownごとにカウントを行い、管理値を超えた場合、測定終了させる。
'
'作成者:
'   2014/02/06 Ver1.0 Y.Okabe
'   2014/04/28 Ver2.0 Y.Okabe CodeModify
'
'使用方法：
'下記3項目を対応すること。

'1.下記コードを　dc_setup() の先頭へ記入
'    '### Management of ProbeCardContact #############################
'            Call Init_ProbeCardInfoFILE
'    '################################################################

'2.下記コードを　EndOFTest_f()　のDコマンド送信後に記入
'    '### Management of ProbeCardContact #############################
'            Call ContactCountAndSave
'    '################################################################

'3.下記コードを　StopPMCMod()に記入
'    '### Management of ProbeCardContact #############################
'    If Flg_StopPMC_Contact = True Then
'        blnFlg_StopPMC = True
'    End If
'    '################################################################
    
'4.ウェーハ1スライスあたりの総コンタクト数を確認して
'Private Const Total_C As Integer = 1976 を変更すること
''

Option Explicit

Public Flg_StopPMC_Contact As Boolean
Private Const LastProcessInfo_FilePATH As String = "F:\Job\ProbeCardData\"
Private Const Total_C As Integer = 346  '1スライスあたりのコンタクト回数　Jobごとに設定

Private Flg_Nasa As Integer
Private ProbeCardContact As Long

Private Dname As String
Private C_Name As String
Private C_SirialNo As String
Private C_Spec As Long

Private CardDataArr() As String
Private EditRow As Integer
Private NowWaferNo As Integer
Private ProbeCardDataInfoFILE As String
Private ProbeCardBackupDataInfoFILE As String
Private CardTypeName As String
    
Private hProber As Integer
Private DataCNT As Integer

'ウェーハ先頭チップで動作。カードデータに不備または管理値オーバーの場合、強制終了する。

Public Function Init_ProbeCardInfoFILE()

Dim TestingNotContinue As Boolean
Dim ProgramName As String
    
If Flg_AutoMode = True And Flg_Tenken = 0 Then

    If WaferNo <> "" Then
       If NowWaferNo = 0 Then
           NowWaferNo = CInt(WaferNo)
           If CardDataInit = False Then
              TestingNotContinue = True
           End If
       Else
           If NowWaferNo = CInt(WaferNo) Then
                Exit Function
           Else
                NowWaferNo = CInt(WaferNo)
                If CardDataInit = False Then
                   TestingNotContinue = True
                End If
           End If
       End If
    End If
    
    ProgramName = ActiveWorkbook.Name
    If InStr(ProgramName, "AATJob") <> 0 Then    'For Nasa
            If Flg_Nasa = 0 Then
                If CardDataInit = False Then
                   TestingNotContinue = True
                End If
            End If
            Flg_Nasa = 1
    End If
    
    If TestingNotContinue = True Then
        Call MsgBox("TestProgram is Close!!", vbOKOnly + vbExclamation, "PROBECARD ARARM")
        ThisWorkbook.Saved = True
        Application.Quit
    End If
    
End If

End Function

'タイプ名取得してタイプ名から読み込みファイル中のカードデータを紐付てカードデータを格納する。
'読み込みファイルに異常があれば強制終了する。
'1スライスを見込んだコンタクト回数が管理値オーバーする場合、強制終了の有効フラグがたつ。

Private Function CardDataInit() As Boolean

    Dim FileNo As Integer
    Dim strText As String
    Dim fileData, fileData2 As Variant
        
    '############################# Youser Check Point #############################
    '測定するデバイスタイプとプローブカードタイプが同じ場合は""で出力される
    '測定するデバイスタイプとプローブカードタイプが異なる場合はタイプ番号が記述される
    CardTypeName = 219
    '##############################################################################
        
    If CardTypeName = "" Then
        Call ProberParameter_TypeNameCheck(CardTypeName)
    End If
    
    If Open_File = False Then
        CardDataInit = False
        Exit Function
    End If
    
    ReDim CardDataArr(DataCNT - 1)
    
    DataCNT = 0
    EditRow = 0
    
    FileNo = FreeFile
    Open ProbeCardDataInfoFILE For Input As #FileNo
    On Error GoTo CloseFile
    
    Do Until EOF(FileNo)
        Line Input #FileNo, strText
            fileData = Split(strText, ":")
            fileData2 = Split(fileData(1), ",")
            CardDataArr(DataCNT) = strText

            If DataCNT > 1 And EditRow = 0 Then
                If CardTypeName = Mid(fileData(0), 4, 3) Then
                    If fileData2(4) = "" Then
                        Dname = fileData(0)
                        C_Name = fileData2(0)
                        C_SirialNo = fileData2(1)
                        C_Spec = fileData2(2)
                        ProbeCardContact = fileData2(3)
                        EditRow = DataCNT
                    End If
                End If
            End If
            DataCNT = DataCNT + 1
    Loop
       
    Close #FileNo
    CardDataInit = True
 
    If EditRow = 0 Then
        CardDataInit = False
        Call MsgBox(" ProbeCardDataFile is Wrong! " & vbCrLf & " Please Check ", vbOKOnly + vbExclamation, "PROBECARD ARARM")
    Else
        If ProbeCardContact + Total_C > C_Spec Then
            CardDataInit = False
            Call MsgBox(" ProbeCard is ContactCount Over!! " & vbCrLf & " Don't Testing ", vbOKOnly + vbExclamation, "PROBECARD ARARM")
        End If
    End If
 
FILE_end:

Exit Function

CloseFile:

    Close #FileNo
    CardDataInit = False
    Call MsgBox(" ProbeCardDataFile is Wrong! " & vbCrLf & " Please Check ", vbOKOnly + vbExclamation, "PROBECARD ARARM")
    GoTo FILE_end
    
End Function

'外部CSVファイルを読み込み、データ数を確認。ファイルに異常があれば終了。

Private Function Open_File() As Boolean

    Dim FileNo As Integer
    Dim strText As String
    Dim fileData As Variant
    Dim SameType As Integer
    Dim FileDate As Date
    Dim BackFileDate As Date
    
    If Sw_Node = 0 Then
        Call JobEnvInit
    End If
    
    ProbeCardDataInfoFILE = LastProcessInfo_FilePATH & "SKMBPC" & Sw_Node & "\SKMBPC" & Sw_Node & ".txt"
    ProbeCardBackupDataInfoFILE = LastProcessInfo_FilePATH & "SKMBPC" & Sw_Node & "\Backup\SKMBPC" & Sw_Node & ".txt"
    
    DataCNT = 0
    SameType = 0
    
    If Dir(ProbeCardDataInfoFILE) = "" Or Dir(ProbeCardBackupDataInfoFILE) = "" Then
        Open_File = False
        Call MsgBox(" ProbeCardDataFile is Nothing!! " & vbCrLf & " Please Check ", vbOKOnly + vbExclamation, "PROBECARD ARARM")
        Exit Function
    End If
    
    FileDate = FileDateTime(ProbeCardDataInfoFILE)
    BackFileDate = FileDateTime(ProbeCardBackupDataInfoFILE)
    
    If FileDate = BackFileDate Then
    Else
        Call MsgBox(" BackUp ProbeCardDataFile Miss Match Error! " & vbCrLf & " Please Check ", vbOKOnly + vbExclamation, "PROBECARD ARARM")
        Open_File = False
        Exit Function
    End If

    FileNo = FreeFile
    Open ProbeCardDataInfoFILE For Input As #FileNo
    On Error GoTo CloseFile
    
    Do Until EOF(FileNo)
        Line Input #FileNo, strText
            fileData = Split(strText, ":")
            
            If DataCNT = 1 Then
                If Sw_Node = fileData(1) And Len(fileData(1)) < 4 And Mid(fileData(1), 1, 1) <> 0 Then
                Else
                GoTo CloseFile
                End If
            End If
            
            If DataCNT > 1 And CardTypeName = Mid(fileData(0), 4, 3) And Len(fileData(0)) < 7 Then
                SameType = SameType + 1
            End If
            DataCNT = DataCNT + 1
    Loop

    Close #FileNo
    
    If SameType = 1 Then
        Open_File = True
    Else
        Call MsgBox(" ProbeCardDataFile is Wrong! " & vbCrLf & " Please Check ", vbOKOnly + vbExclamation, "PROBECARD ARARM")
        Open_File = False
    End If
    
FILE_end:

Exit Function

CloseFile:

    Close #FileNo
    Open_File = False
    Call MsgBox(" ProbeCardDataFile is Wrong! " & vbCrLf & " Please Check ", vbOKOnly + vbExclamation, "PROBECARD ARARM")

    GoTo FILE_end

End Function

'コンタクト回数を外部ファイルへ上書き更新する。
Public Function ContactCountAndSave()

    On Error GoTo ErrorContactCountAndSave
    
    Dim FileNo As Integer
    Dim flag As Boolean
    Dim i As Long
    
    FileNo = FreeFile
    Open ProbeCardDataInfoFILE For Output As FileNo
    flag = True
    
    ProbeCardContact = ProbeCardContact + 1
    CardDataArr(EditRow) = Dname & ":" & C_Name & "," & C_SirialNo & "," & C_Spec & "," & ProbeCardContact & ","
    
    For i = 0 To UBound(CardDataArr)
        Print #FileNo, CStr(CardDataArr(i))
    Next i
    Close FileNo

    FileCopy ProbeCardDataInfoFILE, ProbeCardBackupDataInfoFILE
    
    Exit Function
    
ErrorContactCountAndSave:
    If flag = True Then Close FileNo
    Call MsgBox(" ProbeCardDataFile is Can't Saved!! " & vbCrLf & " Please Check ", vbOKOnly + vbExclamation, "PROBECARD ARARM")
    Flg_StopPMC_Contact = True
    If Flg_Nasa = 1 Then
        Call MsgBox("TestProgram is Close!!", vbOKOnly + vbExclamation, "PROBECARD ARARM")
        ThisWorkbook.Saved = True
        Application.Quit
    End If
 
End Function
Private Sub ProbIni()

'      GPIB Address
'************************************
'      prober   No.5

    Dim GpibAddress As Integer
    GpibAddress = 5
    Call ibdev(0, GpibAddress, 0, 13, 1, &H13, hProber)
End Sub

Private Sub ProberInput_Wait2s(cmd As String)

    '--- PROBER INIT ----
    If hProber = 0 Then
    Call ProbIni
    End If
    cmd = cmd + Chr(13) + Chr(10)
    Call Sleep(2000)
    Call ibwrt(hProber, cmd)
    Call Sleep(2000)

End Sub

Private Sub ProberParameter_TypeNameCheck(ProberParameter_Type As String)

    If Flg_Simulator = 1 Then Exit Sub

    '======= Check ProberParameter vs Program ===
    Dim buff As String * 250
    Dim Paramater As String
    Dim probcmd As String

    probcmd = "G"
    
    Call ProberInput_Wait2s(probcmd)
    buff = "0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"
    Paramater = ""
    
    Call ibrd(hProber, buff)
    Paramater = Paramater + buff

    If Not (Paramater = "0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000") Then
       ProberParameter_Type = Mid(Paramater, 4, 3)
    End If

    If ProberParameter_Type = "" Then
        MsgBox " DeviceTypeName is wrong!"
        Exit Sub
    Else
    End If

End Sub
