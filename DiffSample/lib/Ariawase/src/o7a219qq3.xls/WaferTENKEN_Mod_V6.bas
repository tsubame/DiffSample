Attribute VB_Name = "WaferTENKEN_Mod_V6"
Option Explicit

'Modul NAME     : WaferTENKEN_Mod
'AUTHOR         : N.Togo

Private hProber As Integer
Public TenkenWaferNo As Integer
Public TenkenX As Integer
Public TenkenY As Integer
Public TenkenTemp As Double

Private Sub ProbIni()

'      GPIB Address
'************************************
'      prober   No.5

    Dim GpibAddress As Integer
    GpibAddress = 5
    Call ibdev(0, GpibAddress, 0, 13, 1, &H13, hProber)
End Sub

Private Sub ProberInput(cmd As String)
    
    '--- PROBER INIT ----
    If hProber = 0 Then
        Call ProbIni
    End If
    cmd = cmd + Chr(13) + Chr(10)
    Call ibwrt(hProber, cmd)
    
End Sub

Public Sub TenkenSampleSet(ByRef Flg_TestEnd As Boolean)
    
    Dim xAddress As Integer
    Dim yAddress As Integer
    Dim result As Long
    Dim Flg_ProberAuto As Boolean
    
    '+++ Clear Val +++
    TenkenX = 0
    TenkenY = 0
    
    If TenkenWaferNo <> 0 Then
        result = MsgBox("Set Prober ManualMode?", vbYesNo, "Prober State Check!!")
        If result = vbYes Then
            Call AutoModeStop
            MsgBox "Are You Ready?  TENKEN START OK?"
            InputAddress.Show
            Exit Sub
        Else
            MsgBox "Is Prober AutoMode?"
        End If
    End If

'---- Prober AutoMode Check ----
OneMore:
    Flg_ProberAuto = False
    Call ProberStateCheck(Flg_ProberAuto)
    If Flg_ProberAuto = False Then
        result = MsgBox("Set Prober AutoMode!!", vbOKCancel, "Prober State Check!!")
    
        If result = vbOK Then
            GoTo OneMore
        Else
            result = MsgBox("May I Cancel?", vbYesNo, "Cancel Check!!")
            If result = vbYes Then
                Flg_TestEnd = True
                Exit Sub
            Else
                MsgBox "Set Prober AutoMode!!"
                GoTo OneMore
            End If
        End If
    End If
    
    '--- READ ADDRESS @Tenken_ref.dat ----
    Call ReadAddress(xAddress, yAddress)
    
    '--- MOVE ADDRESS ----
    Call XYMove(xAddress, yAddress)
   
StartCheck:
    result = MsgBox("Check Address!!      TENKEN  START  OK ?", vbYesNoCancel, "Check Address!!")
    If result = vbYes Then      'YES SELECT
        Call Zup
    ElseIf result = vbNo Then   'NO SELECT
        WaferTenkenForm.Show
        GoTo StartCheck
    Else
        result = MsgBox("May I Cancel?", vbYesNo, "Cancel Check!!")
        If result = vbYes Then
            Flg_TestEnd = True
            Exit Sub
        Else
            GoTo StartCheck
        End If
    End If

    '--- READ WAFER SLOT NO ----
    Call ReadWaferNo(TenkenWaferNo)
    
    '--- READ CHIP ADDRESS -----
    Call ReadXY(TenkenX, TenkenY)
    
    '-- READ STAGE TEMPERATURE --
    Call ReadTemp(TenkenTemp)
    
End Sub

Private Sub ProberStateCheck(ByRef Flg_ProberAuto As Boolean)

    Dim i As Long
    Dim answer As Integer
    
    Call MoveCardinalChip(False)
    For i = 0 To 20
        Sleep (50)
        Call ibrsp(hProber, answer)
        If (answer = 70) Or (answer = 74) Then
            Flg_ProberAuto = True
            Exit Sub
        End If
    Next i
    
    Flg_ProberAuto = False
    
End Sub

Private Sub ReadWaferNo(ByRef WaferNo As Integer)

    Dim buff As String * 10
    Dim answer As String
    Dim probcmd As String
    
    probcmd = "S"
    Call ProberInput(probcmd)
    
    buff = "0000000000"
    answer = ""
    Call ibrd(hProber, buff)
    answer = answer + buff
    
    If answer = "0000000000" Then
        Exit Sub
    End If
    
    WaferNo = CInt(Left(answer, 2))
    
End Sub

Private Sub ReadAddress(ByRef xAddress As Integer, ByRef yAddress As Integer)
    
    Dim fp As Integer
    Dim fileName As String
    Dim FileName2 As String
    Dim lngWafer As Long
    Dim strWafer As String
    Dim strXaddress As String
    Dim strYaddress As String
    
'    fileName = ThisWorkbook.Path & "\TENKEN\tenken_ref.dat"
    FileName2 = ThisWorkbook.Path & "\TENKEN\tenken_ref_" & Format(Sw_Node, "000") & ".dat"
    
    '### ファイル検索 ###
'    If Dir(fileName) <> "" Then
'        fp = FreeFile
'        Open fileName For Input As fp
'    ElseIf Dir(FileName2) <> "" Then
'        fp = FreeFile  '2012/11/16 175JobMakeDebug
'        Open FileName2 For Input As fp
'    Else
'        MsgBox "ファイルが見つかりません。"
'    End If

    If Dir(FileName2) <> "" Then
        fp = FreeFile  '2012/11/16 175JobMakeDebug
        Open FileName2 For Input As fp
    Else
        MsgBox "ファイルが見つかりません。"
    End If

    Line Input #fp, strWafer
    lngWafer = CLng(strWafer)
    Line Input #fp, strXaddress
    xAddress = CLng(strXaddress)
    Line Input #fp, strYaddress
    yAddress = CLng(strYaddress)

End Sub

Public Sub XYMove(ByVal TargetX As Integer, ByVal TargetY As Integer)

    Dim NowX As Integer
    Dim NowY As Integer
    Dim DeltaX As Integer
    Dim DeltaY As Integer
    Dim i As Long
    Dim Flg_OffWafer As Boolean
        
    Call MoveCardinalChip
    Call Zdown
    
    For i = 0 To 20
        Call ReadXY(NowX, NowY)
        
        DeltaX = TargetX - NowX
        DeltaY = TargetY - NowY
        
        If (DeltaX = 0) And (DeltaY = 0) Then Exit Sub
        
        '----- X-Move Y-Move   -99< \\ <99   -----------------
        If DeltaX < -99 Then DeltaX = -99:
        If 99 < DeltaX Then DeltaX = 99:
        If DeltaY < -99 Then DeltaY = -99:
        If 99 < DeltaY Then DeltaY = 99:
        Call Move(DeltaX, DeltaY, Flg_OffWafer)
        If Flg_OffWafer = True Then Exit For
    Next i
    
    MsgBox "Don't Select Target Sample!"
    
End Sub

Private Sub ReadXY(ByRef x As Integer, ByRef y As Integer)

    Dim buff As String * 8
    Dim answer As String
    Dim probcmd As String
    
    probcmd = "A"
    Call ProberInput(probcmd)
    
    buff = "00000000"
    answer = ""
    Call ibrd(hProber, buff)
    answer = answer + buff
    
    If answer = "00000000" Then
        Exit Sub
    End If
    
    x = Mid(answer, 1, 3)
    y = Mid(answer, 4, 3)
    
End Sub

Private Sub Move(x As Integer, y As Integer, ByRef Flg_OffWafer As Boolean)
    
    Dim xx As String
    Dim yy As String
    Dim probcmd As String
    
    '---- X SET ----
    If -99 <= x And x <= -10 Then
        xx = "-" & CStr(Abs(x))
    ElseIf -9 <= x And x <= -1 Then
        xx = "-0" & CStr(Abs(x))
    ElseIf 0 <= x And x <= 9 Then
        xx = "+0" & CStr(Abs(x))
    ElseIf 10 <= x And x <= 99 Then
        xx = "+" & CStr(Abs(x))
    End If
    
    '---- Y SET ----
    If -99 <= y And y <= -10 Then
        yy = "-" & CStr(Abs(y))
    ElseIf -9 <= y And y <= -1 Then
        yy = "-0" & CStr(Abs(y))
    ElseIf 0 <= y And y <= 9 Then
        yy = "+0" & CStr(Abs(y))
    ElseIf 10 <= y And y <= 99 Then
        yy = "+" & CStr(Abs(y))
    End If
    
    probcmd = "X" & xx & "Y" & yy
    Call ProberInput(probcmd)
    Flg_OffWafer = OffWaferCheck
        
End Sub

Private Function OffWaferCheck() As Boolean
    
    Dim i As Long
    Dim answer As Integer
    
    OffWaferCheck = False
    For i = 0 To 10
        Call ibrsp(hProber, answer)
        Sleep (50)
        '--- OffWafer or Error ---
        If (answer = 79) Or (answer = 75) Then
            OffWaferCheck = True
            Exit Function
        End If
    Next i

End Function

Private Sub Zup(Optional Flg_SRQCheck As Boolean = True)
    Dim probcmd As String
    probcmd = "Z"
    Call ProberInput(probcmd)
    Call Sleep(1000)
    If Flg_SRQCheck = True Then Call SRQCheck(67)
End Sub

Private Sub Zdown(Optional Flg_SRQCheck As Boolean = True)
    Dim probcmd As String
    probcmd = "D"
    Call ProberInput(probcmd)
    Call Sleep(1000)
    If Flg_SRQCheck = True Then Call SRQCheck(68)
End Sub

Private Sub MoveCardinalChip(Optional Flg_SRQCheck As Boolean = True)
    Dim probcmd As String
    probcmd = "B"
    Call ProberInput(probcmd)
    Call Sleep(1000)
    If Flg_SRQCheck = True Then Call SRQCheck(70, 74)
End Sub

Public Sub AutoModeStop()
    Dim probcmd As String
    probcmd = "e"
    Call ProberInput(probcmd)
End Sub

Private Sub SRQCheck(ByVal SrqNo1 As Integer, Optional ByVal SrqNo2 As Integer = -1)
    
    Dim i As Long
    Dim answer As Integer
    
    For i = 0 To 60000
        Sleep (20)
        Call ibrsp(hProber, answer)
        If (answer = SrqNo1) Or (answer = SrqNo2) Then
            Exit Sub
        End If
    Next i
    
    MsgBox "Prober doesn't respond."
    
End Sub

Public Sub ReadTemp(ByRef temp As Double)

    Dim buff As String * 7
    Dim answer As String
    Dim probcmd As String

    If Flg_Simulator = 1 Then Exit Sub
    
    probcmd = "f1"
    Call ProberInput(probcmd)
    
    buff = "0000000"
    answer = ""
    Call ibrd(hProber, buff)
    answer = answer + buff
    
    If answer = "00000000" Then
        Exit Sub
    End If
    
    temp = CDbl(Mid(answer, 1, 5))
    
End Sub


'***************: probe mark Check  for NagasakiTEC :*************************************
Public Sub c_Command()
    
    Dim probcmd As String
    Dim site As Long
    Dim PassFailInfo As String
    Dim i As Long
    Dim buf As String
    Dim BinNumber As Long
    
    probcmd = ""
    PassFailInfo = ""
    i = 1
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
        
            BinNumber = TheExec.sites.site(site).FirstBinNumber
            
            If BinNumber = -1 Then
                PassFailInfo = "0" & PassFailInfo
            Else
                PassFailInfo = "1" & PassFailInfo
            End If
        Else
            PassFailInfo = "0" & PassFailInfo
        End If
        
        If i = 4 Or site = nSite Then
            PassFailInfo = Format(PassFailInfo, "0000")
            
            Select Case PassFailInfo
                Case "0000": buf = "@"
                Case "0001": buf = "A"
                Case "0010": buf = "B"
                Case "0011": buf = "C"
                Case "0100": buf = "D"
                Case "0101": buf = "E"
                Case "0110": buf = "F"
                Case "0111": buf = "G"
                Case "1000": buf = "H"
                Case "1001": buf = "I"
                Case "1010": buf = "J"
                Case "1011": buf = "K"
                Case "1100": buf = "L"
                Case "1101": buf = "M"
                Case "1110": buf = "N"
                Case "1111": buf = "O"
            End Select
            
            probcmd = probcmd & buf
            PassFailInfo = ""
            i = 0
        End If
        
        i = i + 1
    Next site
    
    probcmd = "c" & probcmd
    
    Call ProberInput(probcmd)
    Call SRQCheck(89)
    
End Sub


