Attribute VB_Name = "JobRouteMod"
Option Explicit
'Ver1.1 2013/02/01 H.Arikawa Get_DcTestLastNo追加。
'                            GradeOnの処理を修正｡
'...
'Ver2.0 2013/09/25 H.Arikawa ParameterBankのクリア処理追加。通常測定をAcquirePlaneを複数Deleteしないまま白コブ・マージンに進んだ際のケア
'                            不要Enableword削除。(color,function ←自動化では使用しない為)
'Ver2.1 2013/10/28 H.Arikawa 条件設定省略のフラグ化


'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const EEE_AUTO_JOB_LOCATE = 2      '1:長崎200mm,2:長崎300mm,3:熊本

Public LastBin(nSite) As Double
Public SortBin(nSite) As Double

Dim FirstFail As Boolean

Public InGrade As Boolean
Public MarginOn As Boolean

Private hProber As Integer 'add chip kenma

Dim mChipExist(nSite) As Boolean

Public DcTestLastNumber As Integer
Public ImageTestLastNumber As Integer
Public Flg_FirstCompleteRun As Boolean

Public Flg_FailSite(nSite) As Boolean
Public Flg_FailSiteImage(nSite) As Boolean
Public G_Relay(nSite) As Long



Public Sub FlagSetup()

    Dim site As Long
    
    TheExec.Flow.EnableWord("dc") = False
    TheExec.Flow.EnableWord("current") = False
    TheExec.Flow.EnableWord("image") = False
    TheExec.Flow.EnableWord("ngCap1") = False
    TheExec.Flow.EnableWord("ngCap2") = False
    TheExec.Flow.EnableWord("ngCap3") = False
    TheExec.Flow.EnableWord("ngCap4") = False
    TheExec.Flow.EnableWord("ngCap5") = False
    MarginOn = False
    
    If TheExec.CurrentJob = NormalJobName Then
        TheExec.RunOptions.DoAll = True

        TheExec.Flow.EnableWord("dc") = True
        TheExec.Flow.EnableWord("current") = True
        TheExec.Flow.EnableWord("image") = True
        TheExec.Flow.EnableWord("grade") = True
       
        If Flg_shiroten = 1 And Chip_f >= 2 Then
            TheExec.Flow.EnableWord("shiroten") = True
        End If

        If Flg_margin = 1 And Chip_f >= 2 And Chip_f < 10 Then
            TheExec.Flow.EnableWord("margin") = True
            MarginOn = True
        End If

    Else
        TheExec.RunOptions.DoAll = True
        
        TheExec.Flow.EnableWord("dc") = True
        TheExec.Flow.EnableWord("current") = True
        TheExec.Flow.EnableWord("image") = True
        TheExec.Flow.EnableWord("grade") = True
        
    End If

    FirstFail = True
    InGrade = False

    For site = 0 To nSite
        DisableSiteCount(site) = 0
    Next site

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = False Then
            DisableSiteCount(site) = 1
        End If
    Next site

End Sub

Public Sub RouteSetup(ByRef testStatus As Long, ByVal site As Long)

    Dim DCTestNumber As Double
    
    If TheExec.CurrentJob = NormalJobName Then
    
        ' Report Status
        If testStatus <> logTestPass Then
            If FirstFail = True Then
                If InGrade = True Then
                    TheExec.Flow.EnableWord("dc") = False
                    TheExec.Flow.EnableWord("current") = False
                    TheExec.Flow.EnableWord("image") = False
                    TheExec.Flow.EnableWord("grade") = False
                    TheExec.Flow.EnableWord("ngCap1") = False
                    TheExec.Flow.EnableWord("ngCap2") = False
                    TheExec.Flow.EnableWord("ngCap3") = False
                    TheExec.Flow.EnableWord("ngCap4") = False
                    TheExec.Flow.EnableWord("ngCap5") = False
                Else
                    TheExec.Flow.EnableWord("dc") = False
                    TheExec.Flow.EnableWord("current") = False
                    TheExec.Flow.EnableWord("image") = False
                    TheExec.Flow.EnableWord("grade") = True
                    TheExec.Flow.EnableWord("ngCap1") = False
                    TheExec.Flow.EnableWord("ngCap2") = False
                    TheExec.Flow.EnableWord("ngCap3") = False
                    TheExec.Flow.EnableWord("ngCap4") = False
                    TheExec.Flow.EnableWord("ngCap5") = False
                End If
                TheExec.RunOptions.DoAll = True '    /*** 17/Mar/02 takayama append
                FirstFail = False
            
            Else
    '                TheExec.Flow.EnableWord("dc") = False
    '                TheExec.Flow.EnableWord("current") = False
    '                TheExec.Flow.EnableWord("image") = False
    '                TheExec.Flow.EnableWord("color") = False
    '                TheExec.Flow.EnableWord("function") = False
    '                TheExec.Flow.EnableWord("grade") = False
            End If
        Else
    '        theexec.Sites.site(0).TestResult = sitePass
        End If
    
    Else
        ' Report Status
        If testStatus <> logTestPass Then
    '        For site = 0 To nSite
                DCTestNumber = TheExec.sites.site(site).TestNumber
                If DCTestNumber < 1000 Then
                    TheExec.Flow.EnableWord("dc") = False
                    TheExec.Flow.EnableWord("current") = False
                    TheExec.Flow.EnableWord("image") = False
                    TheExec.Flow.EnableWord("shiroten") = False
                    TheExec.Flow.EnableWord("margin") = False
                    TheExec.Flow.EnableWord("grade") = True
                    TheExec.Flow.EnableWord("ngCap1") = False
                    TheExec.Flow.EnableWord("ngCap2") = False
                    TheExec.Flow.EnableWord("ngCap3") = False
                    TheExec.Flow.EnableWord("ngCap4") = False
                    TheExec.Flow.EnableWord("ngCap5") = False
                End If
                TheExec.sites.site(site).TestResult = siteFail
    '        Next site
        End If

    End If

End Sub
'2013/01/23 H.Arikawa EndCondition Beta Add
Private Function EndOFTest_f() As Long

    Dim site As Long

    '@@@ Captureシステム毎に中身を分ける必要有り @@@
    'Capture Unit rest
    Call CaptureResetSequence
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        
    Dim ArgArr() As String
    Dim strEndCondition As String
    
    If EeeAutoGetArgument(ArgArr, EEE_AUTO_ENDOFTEST_PARAM) Then
        strEndCondition = ArgArr(0)
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
            Call SetForceEnableTestCondition(strEndCondition)
    End If
        Call TheCondition.SetCondition(strEndCondition)
    Else
        MsgBox "The Number of EndOFTest_f's arguments is invalid!"
    End If
    
    '=== Except TestCondition End Process Start ===
    If nSite > 1 Then
        Call SET_RELAY_CONDITION("GND_All_Beta", "-")
    End If
    Call OptSet("DARK")
    '=== Except TestCondition End Process End ===
    
    Call UnInitializeEeeAutoModules 'EeeAutoModの終了処理
   
    For site = 0 To nSite
        DisableSiteCount(site) = 0
    Next site

    If TheHdw.Digital.Patgen.IsRunningAnySite = True Then
        TheHdw.Digital.Patgen.Ccall = True
        TheHdw.Digital.Patgen.HaltWait
        TheHdw.Digital.Patgen.Ccall = False
    End If

    If TheExec.CurrentJob = NormalJobName Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                TheExec.sites.site(site).BinNumber = LastBin(site)
                TheExec.sites.site(site).SortNumber = SortBin(site)
            End If
        Next site
    End If
    
    
    Call mf_CloseDefectFile
    
'{ ログレポートライターの後処理
    CloseDcLogReportWriter
'}
    TheHdw.WAIT 6 * mS

    If (Flg_AutoMode = True) Then Call c_Command
                    
     'add chip kenma
    Dim probcmd As String
    Dim probcmd2 As String
    
    'EEE_AUTO_JOB_LOCATE:1or2 : 長崎　研磨ユニット　  Dコマンド送信のみ
    'EEE_AUTO_JOB_LOCATE:3    : 熊本　研磨シート搬送　D,Qコマンド送信
    
#If EEE_AUTO_JOB_LOCATE = 1 Or EEE_AUTO_JOB_LOCATE = 2 Then
    If Flg_Tenken = 0 And Flg_AutoMode = True Then
        probcmd = "D"
        Call ProberInput(probcmd)
        Call SRQCheck(68)   '44
    End If
#Else
    If Flg_Tenken = 0 And Flg_AutoMode = True Then
        probcmd = "D"
        Call ProberInput(probcmd)
        Call SRQCheck(68)   '44
        
        '### Management of ProbeCardContact #############################
        'Kumamoto Only
        #If EEE_AUTO_JOB_LOCATE = 3 Then
            Call ContactCountAndSave
        #End If
        '################################################################
        
        probcmd2 = "Q"
        Call ProberInput(probcmd2)
        Call SRQCheck(75)   '4B
    End If
#End If

End Function

Public Sub GradeOn()

    Dim site As Long
    
    InGrade = True
    
    GradeLastBin = 1
    GradeSortBin = 1
    
    TheExec.RunOptions.DoAll = True
    If FirstFail = True Then
        For site = 0 To nSite
            TheExec.sites.site(site).Active = True
            
            '@@@ 基板毎に中身を分ける必要有り @@@
            Call GND_Connect(site)
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            
            If TheExec.sites.site(site).Active = True And DisableSiteCount(site) = 0 Then
                LastBin(site) = 1
                SortBin(site) = 1
                Ng_test(site) = 0
            End If
        Next site
    End If

    For site = 0 To nSite
        TheExec.sites.site(site).Active = True
        If TheExec.sites.site(site).BinNumber <> -1 Then
            LastBin(site) = TheExec.sites.site(site).FirstBinNumber
            SortBin(site) = TheExec.sites.site(site).FirstSortNumber
        Else
            If DisableSiteCount(site) = 0 Then
                LastBin(site) = GradeLastBin
                SortBin(site) = GradeSortBin
            End If
        End If
    Next site

End Sub

Private Function ShirotenCheck_f() As Long

    Dim flg As Integer
    Dim site As Long
    Dim AcitveSiteNum As Integer

    '++++++ TOPT STOP ++++++
    TheExec.RunOptions.AutoAcquire = False
    
    '===== PlaneClear =====
    'Fixed(読み取り専用)以外のプレーンをクリアする
    Call TheParameterBank.Clear
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            TheExec.sites.site(site).TestNumber = GetShirotenFirstTestNumber
        End If
    Next site
    
    AcitveSiteNum = 0
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then AcitveSiteNum = AcitveSiteNum + 1
    Next site
    
    flg = 0
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            'User Define Function
            If ActiveSiteCheck(site) = False Then flg = flg + 1
        End If
    Next site

    If flg = AcitveSiteNum Then
        TheExec.Flow.EnableWord("shiroten") = False
    End If

    For site = 0 To nSite
        If LastBin(site) > 1 And LastBin(site) <= 5 Then DisableSiteCount(site) = 0
    Next site

End Function

Private Function MarginCheck_f() As Long
    
    Dim i As Integer
    Dim flg As Integer
    Dim site As Long
    Dim AcitveSiteNum As Integer

    '++++++ TOPT STOP ++++++
    TheExec.RunOptions.AutoAcquire = False
    
    '===== PlaneClear =====
    'Fixed(読み取り専用)以外のプレーンをクリアする
    Call TheParameterBank.Clear
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            TheExec.sites.site(site).TestNumber = GetMarginFirstTestNumber
        End If
    Next site
    
    AcitveSiteNum = 0
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then AcitveSiteNum = AcitveSiteNum + 1
    Next site
    
    flg = 0
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            'User Define Function
            If ActiveSiteCheck(site) = False Then flg = flg + 1
        End If
    Next site

    If flg = AcitveSiteNum Then
        TheExec.Flow.EnableWord("margin") = False
    Else
        Marflg(Chip_f) = 1
    End If

    If Chip_f = 9 Then
        For i = 0 To 9
            Marflg(i) = 0
        Next
    End If

    For site = 0 To nSite
        If LastBin(site) > 1 And LastBin(site) <= 5 Then DisableSiteCount(site) = 0
    Next site

End Function

'From IMX145

Public Function SiteCheckForFW(argc As Long, argv() As String) As Long
    Call SiteCheck
End Function

Public Sub SiteCheck()

    Dim site As Long
    
    For site = 0 To nSite
        If DisableSiteCount(site) = 0 And TheExec.sites.site(site).Active = True Then
            Ng_test(site) = TheExec.sites.site(site).TestNumber
        End If
    Next site

    For site = 0 To nSite
        If DisableSiteCount(site) = 0 And TheExec.sites.site(site).Active = True Then
            '"grade"項目の直前項目まで流れたら、"1回は最後まで流れたよフラグ"をONにする
            'When test reaches to the test item just before "grade" items, the flag
            '"First Complete Run" is enabled.(MM)
            If Ng_test(site) = ImageTestLastNumber And Flg_FirstCompleteRun = False Then Flg_FirstCompleteRun = True
        End If
    Next site
    
    For site = 0 To nSite
        If CurrentJobName = NormalJobName Then  '2012/11/16 175JobMakeDebug
            If DisableSiteCount(site) = 1 Or mChipExist(site) = False Then
                If Flg_margin = 0 And Flg_shiroten = 0 Then
                    If Flg_FailSite(site) = False Then
                        If Ng_test(site) < DcTestLastNumber Then        'It is not ImageTest(DC Test)
                            '@@@ DUT情報毎に中身を替えないといけない。@@@
                            Call DisconnectAllDevicePins(site)                 'FailSite All OPEN   '2012/11/16 175JobMakeDebug
                            Call GND_DisConnect(site)                          '2012/11/16 175JobMakeDebug
                            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                            Flg_FailSite(site) = True
                        End If
                    End If
                    TheExec.sites.site(site).Active = False
                Else
                    If TheExec.sites.site(site).TestNumber > 5000 Then
                        If (LastBin(site) >= 24) Or (LastBin(site) = 14) Or (LastBin(site) = 12) Then
                            If Flg_FailSite(site) = False Then
                                '@@@ DUT情報毎に中身を替えないといけない。@@@
                                Call DisconnectAllDevicePins(site)   'FailSite All OPEN  '2012/11/16 175JobMakeDebug
                                Call GND_DisConnect(site)               '2012/11/16 175JobMakeDebug
                                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                            End If
                            TheExec.sites.site(site).Active = False
                            Flg_FailSite(site) = True
                        End If
                    Else
                        If Flg_FailSite(site) = False Then
                            If Ng_test(site) < DcTestLastNumber Then        'It is not ImageTest(DC Test)
                                '@@@ DUT情報毎に中身を替えないといけない。@@@
                                Call DisconnectAllDevicePins(site)   'FailSite All OPEN   '2012/11/16 175JobMakeDebug
                                Call GND_DisConnect(site)                '2012/11/16 175JobMakeDebug
                                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                                Flg_FailSite(site) = True
                            End If
                        End If
                        TheExec.sites.site(site).Active = False
                    End If
                End If
            End If
        
        ElseIf CurrentJobName = "TENKEN" Then
            If TheExec.sites.site(site).Active = True Then
                Call GND_Connect(site)
            ElseIf TheExec.sites.site(site).Active = False And G_Relay(site) = 0 Then
                Call DisconnectAllDevicePins(site)
                Call GND_DisConnect(site)
                G_Relay(site) = 1
            End If
        End If
    Next

End Sub

Public Sub mf_ChipExistenceCheck()
'内容:
'   マルチ測定におけるウェハ周辺チップ有無の確認
'
'パラメータ:
'
'注意事項:

    Dim site As Long

    For site = 0 To nSite
        If TheExec.CurrentJob = NormalJobName Then
            If TheExec.sites.site(site).Active = False Then
                mChipExist(site) = False
            Else
                mChipExist(site) = True
            End If
        End If
    Next site

End Sub

Private Sub ProberInput(cmd As String)
    
    '--- PROBER INIT ----
    If hProber = 0 Then
        Call ProbIni
    End If
    cmd = cmd + Chr(13) + Chr(10)
    Call ibwrt(hProber, cmd)
    Call Sleep(10)
    
End Sub

Private Sub ProbIni()

'      GPIB Address
'************************************
'      prober   No.5

    Dim GpibAddress As Integer
    GpibAddress = 5
    Call ibdev(0, GpibAddress, 0, 13, 1, &H13, hProber)
End Sub

Private Sub SRQCheck(ByVal SrqNo1 As Integer, Optional ByVal SrqNo2 As Integer = -1)
    
    Dim i As Long
    Dim answer As Integer
    
    For i = 0 To 120000
        Call ibrsp(hProber, answer)
        Sleep (10)
        If (answer = SrqNo1) Or (answer = SrqNo2) Then Exit Sub
    Next i
    
    MsgBox "Prober doesn't respond."
    
End Sub

Public Sub SepareteGnd(ByVal site As Long)

    If TheExec.sites.site(site).Active = False Then
    
        If Flg_vsubOpen(site) = False Then

            '強制ON
            TheExec.sites.site(site).Active = True
        
            'VDDSUBピン取得
            Dim strVddSubPinName As String
            strVddSubPinName = TheVarBank.Value(PIN_NAME_VDDSUB)
            
            '50mAはAPMUもDPSもPPMUも通る
            Call SetFVMI(strVddSubPinName, 0 * V, 50 * mA, site)
            Call DisconnectPins(strVddSubPinName, site)
            
            '強制ON
            Flg_vsubOpen(site) = True
            
            'フラグOFF
            TheExec.sites.site(site).Active = False
        
        End If
        
    End If

End Sub

'2013/02/01 H.Arikawa Debug
'2013/02/08 H.Arikawa 暫定
Public Function Get_DcTestLastNo() As Integer
'内容:
'　　Image ACQTBLから最後のDC項目のテストNo.を取得する
'
'パラメータ:
'
'注意事項:

    On Error GoTo ErrorDetected

    Dim i As Long
    Dim AutoON_LineNum As Integer
    Dim FlowT_LineNum As Integer
    Dim OFF_check_Flag As Boolean
    Dim MaxRow As Integer
    Dim Instance_Name() As String
    Dim TestI_Name As String
    Dim tmpInstance_Name As String

'--- Image ACQTBLシート ---
    '### Image ACQTBLシート読み込み ###
    Dim wkshtObj_IA As Object
    Set wkshtObj_IA = ThisWorkbook.Sheets("Image ACQTBL")
    '======= WorkSheet ErrorProcess ========
    If wkshtObj_IA Is Nothing Then
        MsgBox "Not Find Sheet : " & " Image ACQTBL"
        Exit Function
    End If
    
    '--- Image ACQTBLシートのグループ化を表示 ---
    wkshtObj_IA.Outline.ShowLevels RowLevels:=2, ColumnLevels:=2

    '### Image ACQTBLシートのAuto Acquire行を上から読み込み "ON"になる所を探す ###
    AutoON_LineNum = 5
    
    '### 検索開始位置を設定 ###
    i = AutoON_LineNum
    OFF_check_Flag = False
    
    MaxRow = wkshtObj_IA.Range("C5").End(xlDown).Row   'データが入っている最後の最終行を返す

    Do While wkshtObj_IA.Cells(i, 4) <> "ON"
        AutoON_LineNum = AutoON_LineNum + 1
        i = i + 1
        OFF_check_Flag = True
        If i = MaxRow Then
            OFF_check_Flag = False
            GoTo All_off:
        End If
    Loop
    
All_off:
    '### Auto Acquireが全てOFFの時のバカ除け ###
    If OFF_check_Flag = False Then
        AutoON_LineNum = 5
    End If

    '### "ON"になったInstance Nameを取得 ###
    tmpInstance_Name = wkshtObj_IA.Cells(AutoON_LineNum, 5)
    Instance_Name = Split(tmpInstance_Name, "_Con")

'--- Flow Tableシート ---
    '### Flow Tableシート読み込み ###
    Dim wkshtObj_FT As Object
    Set wkshtObj_FT = ThisWorkbook.Sheets("Flow Table")
    '======= WorkSheet ErrorProcess ========
    If wkshtObj_FT Is Nothing Then
        MsgBox "Not Find Sheet : " & " Flow Table"
        Exit Function
    End If

    '### Flow TableシートのParameter行を上から読み込み Test Instancesで検索した文字と同じになる所を探す ###
    FlowT_LineNum = 5
    
    '### 検索開始位置を設定 ###
    i = FlowT_LineNum
    
    Do While wkshtObj_FT.Cells(i, 8) <> Instance_Name(0)
        FlowT_LineNum = FlowT_LineNum + 1
        i = i + 1
    Loop

    '### 一致したTest Nameの1つ上のテストNo.を取得 ###
    Get_DcTestLastNo = CInt(wkshtObj_FT.Cells(FlowT_LineNum - 1, 10))
    
    Exit Function
    
ErrorDetected:
    MsgBox "Get_DcTestLastNo Process Fail!! Please Check Program!! "
    DisableAllTest
    First_Exec = 0

End Function

'通常項目の末尾項目のTNumを取得する関数。
'自動生成JOBでAutoAcquire=True時に、ムラのRef画像読み込み処理などでSetAllActiveを実行することがある。
'死んでいるサイトも復活するが、AutoAcquire=Trueであればキャプチャをしに行って、エラーになる。
'これを回避するために、ムラのRef画像読み込み処理などが完了するまでは、AutoAcquireをFalseにして
'流したい。ムラのRef画像読み込みなどの処理が、Flow上のどのテスト項目に埋まっているかは
'予めわかならいので、通常項目が一度は少なくとも1つのサイトで完全実行されることを確認する。
'本関数は、その「通常項目の末尾の項目の特定」を行うのが目的。
'尚、「通常項目」とは、Flow Tableの<enable>カラムの値が"image", "dc"のいずれかのもの。
'  To obtain the TNum value of the last normal (= <enable> is "image" or "dc") test instance
'on Flow Table.
Public Function Get_ImageTestLastNo() As Long

    Const FLOW_SHEET As String = "Flow Table"
    Const COLUMN_LABEL As String = "B"
    Const COLUMN_OFFSET_ENABLE As Long = 1
    Const COLUMN_OFFSET_TNUM As Long = 8
    Const LABEL_GRADE As String = "grade"
    Const ENABLE_IMAGE As String = "image"
    Const ENABLE_DC As String = "dc"

On Error GoTo ErrorDetected

    'To find "grade" in <Label> column on Flow Table sheet.
    Dim gradeRange As Range
    Set gradeRange = ThisWorkbook.Worksheets(FLOW_SHEET).Range(COLUMN_LABEL & ":" & COLUMN_LABEL).Find(LABEL_GRADE, lookat:=xlWhole)
    
    'To obtain the TNum of the previous line of the found row.
    If Not gradeRange Is Nothing Then
        If Trim(gradeRange.offset(0, COLUMN_OFFSET_ENABLE)) = "" And (LCase(Trim(gradeRange.offset(-1, COLUMN_OFFSET_ENABLE))) = ENABLE_IMAGE Or LCase(Trim(gradeRange.offset(-1, COLUMN_OFFSET_ENABLE))) = ENABLE_DC) Then
            Get_ImageTestLastNo = gradeRange.offset(-1, COLUMN_OFFSET_TNUM).Value
        Else
            'In case the previous line's enable is not "image" nor "dc", then disable test.
            Call MsgBox("Get_ImageTestLastNo: Failed in finding last test item with enable-word ""image"".")
            Call DisableAllTest
            First_Exec = 0
        End If
    Else
        Call MsgBox("Get_ImageTestLastNo: Failed in finding ""grade"" label on Flow Table sheet.")
        Call DisableAllTest
        First_Exec = 0
    End If
    Exit Function
    
ErrorDetected:
    Call MsgBox("Get_ImageTestLastNo: Possibly Flow Table sheet is absent.")
    Call DisableAllTest
    First_Exec = 0
End Function

Public Sub Get_ISPTNsheet()

    Dim i As Long
    Dim CellNum As Long

    '### ISPNシート読み込み ###
    Dim wkshtObj As Object
    Set wkshtObj = ThisWorkbook.Sheets("ISPTN_sheet")
    '======= WorkSheet ErrorProcess ========
    If wkshtObj Is Nothing Then
        MsgBox "Not Find Sheet : " & " ISPTN_sheet"
        Exit Sub
    End If

    '### ISPNシート表の行数読み込み ###
    i = 7
    CellNum = 0
    Do While wkshtObj.Cells(i, 4) <> ""
        i = i + 1
        CellNum = CellNum + 1
    Loop

    '### ISPNシート表から各数値読み込み ###
    ReDim IsptnValue(CellNum - 1) As String
    ReDim IsptnupValue(CellNum - 1) As String
    ReDim IsptndnValue(CellNum - 1) As String
    i = 7
    Do While wkshtObj.Cells(i, 4) <> ""
        IsptnValue(i - 7) = wkshtObj.Cells(i, 4)
        IsptnupValue(i - 7) = wkshtObj.Cells(i, 5)
        IsptndnValue(i - 7) = wkshtObj.Cells(i, 6)
        i = i + 1
    Loop

End Sub
