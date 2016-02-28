VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_PALS_OptAdj_Main 
   Caption         =   "PALS - Auto Opt Adjust"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   OleObjectBlob   =   "frm_PALS_OptAdj_Main.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frm_PALS_OptAdj_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

'##########################################################
'フォームの×ボタンを消す処理
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000

' ウィンドウに関する情報を返す
Private Declare Function GetWindowLong Lib "USER32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' ウィンドウの属性を変更
Private Declare Function SetWindowLong Lib "USER32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' Activeなウィンドウのハンドルを取得
Private Declare Function GetActiveWindow Lib "USER32.dll" () As Long
' メニューバーを再描画
Private Declare Function DrawMenuBar Lib "USER32.dll" (ByVal hWnd As Long) As Long

Private Sub cb_savecsv_Click()
    If MsgBox("If You Push [OK],PALS Save to CSV...", vbOKCancel + vbQuestion, OPTTOOLNAME) = vbCancel Then
        Exit Sub
    End If
    
    If OptCond.IllumMaker = NIKON Then
        Call sub_OutPutCsv(NIKON_WRKSHT_NAME, OptFileName)
    Else
        Call sub_OutPutCsv(IA_WRKSHT_NAME, OptFileName)
    End If
End Sub

Private Sub UserForm_Activate()
    Dim hWnd As Long
    Dim Wnd_STYLE As Long

    hWnd = GetActiveWindow()
    Wnd_STYLE = GetWindowLong(hWnd, GWL_STYLE)
    Wnd_STYLE = Wnd_STYLE And (Not WS_SYSMENU)
    SetWindowLong hWnd, GWL_STYLE, Wnd_STYLE
    DrawMenuBar hWnd
    Me.Caption = OPTTOOLNAME & " Ver:" & OPTTOOLVER

End Sub

Private Sub UserForm_Initialize()
    
    Call sub_PalsFileCheck(PALS_PARAMFOLDERNAME_OPT)
    
'    Set objLoadedJob = GetObject(, "excel.application")
    
    '>>>2011/9/05 M.IMAMURA Add. ForCIS
    If gblnForCis = True Then
        ob_Site0only.Value = False
        ob_AveAllSite.Value = True
        '>>>2011/8/31 M.IMAMURA Add. ForCIS
        op_AutoAdjust.enabled = False
        op_NotAdjust.Value = True
        Btn_ContinueOnFail.Value = True
        '<<<2011/8/31 M.IMAMURA Add. ForCIS
    End If
    '>>>2011/9/05 M.IMAMURA Add. ForCIS
    
    cbo_AdjNum.Clear
    cbo_AveNum.Clear

    Dim intSetNum As Integer
    For intSetNum = 1 To 20
        cbo_AdjNum.AddItem intSetNum
        cbo_AveNum.AddItem intSetNum
    Next

    cbo_AdjNum.Value = intOptTryNum
    cbo_AveNum.Value = intOptAveNum
    
    '>>>2011/4/22 M.IMAMURA ADD
    Dim intOptTestCondLoop As Integer
    Dim intOptCondLoop As Integer
    Dim blnExistOptcond As Boolean

    Set OptCond = Nothing
    Set OptCond = New csPALS_OptCond

    For intOptTestCondLoop = 0 To PALS.CommonInfo.TestCount
        If PALS.CommonInfo.TestInfo(intOptTestCondLoop).OptIdentifier <> "" Then
            blnExistOptcond = False
            For intOptCondLoop = 0 To OptCond.OptCondNum
                If OptCond.CondInfoI(intOptCondLoop).OptIdentifier = PALS.CommonInfo.TestInfo(intOptTestCondLoop).OptIdentifier Then
                    blnExistOptcond = True
                    Exit For
                End If
            Next intOptCondLoop
            If blnExistOptcond = False Then
                cmd_start.enabled = False
                Call sub_errPALS("Not Found OptIdentifier @Sheet[" & NIKON_WRKSHT_NAME & "]!! TestName[" & PALS.CommonInfo.TestInfo(intOptTestCondLoop).Parameter & "] - OptIdentifier[" & PALS.CommonInfo.TestInfo(intOptTestCondLoop).OptIdentifier & "] in TestInstancesSheet", "4-1-01-5-01")
                Exit For
            End If
        End If
    Next intOptTestCondLoop
    '<<<2011/4/22 M.IMAMURA ADD

    If g_blnUseCSV = False Then Me.cb_savecsv.Visible = False

End Sub

Public Sub cmd_start_Click()
    Dim ws As Worksheet
    Dim bExist As Boolean
    Dim blnOptCheck As Boolean
    Dim blnOptUpdate As Boolean
    Dim MeasureDatalogInfo As DatalogInfo

    Dim IdenRow As Long
    Dim ShtPoint As Variant
    Dim strOptIdenShTgt As String
    
    Dim OptPoint As Variant
    Dim OptRow As Long
    
    On Error GoTo errPALScmd_start_Click
    
    If sub_CheckResultFormat = False Then
        Exit Sub
    End If
    
    g_ErrorFlg_PALS = False
    blnOptCheck = False
    g_blnOptStop = False
    ReDim Preserve dblDataPrev(PALS.CommonInfo.TestCount)
    ReDim Preserve intWedgePrev(PALS.CommonInfo.TestCount)

    '######### UserStop
    If cmd_start.Caption = "Stop" Then
        If MsgBox("Pushed [Stop] Button" & vbCrLf & "Do You Want Stop?", vbYesNo, OPTTOOLNAME) = vbYes Then
            g_blnOptStop = True
            Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Canceled OptAdj...  ")
            cmd_start.enabled = False
        End If
        Exit Sub
    End If
    
    '######### Reset OptCondition
'    Call OptIni
    Set OptCond = Nothing
    Set OptCond = New csPALS_OptCond
    
    '######### StartCheck
    '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    If g_RunAutoFlg_PALS = False Then
        If MsgBox("Pushed [Start] Button, Ready?", vbOKCancel, OPTTOOLNAME) = vbCancel Then
            Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Canceled...  ")
            Exit Sub
        End If
    End If
    '<<<2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    
    '######### GetMaxLux
    If OptCond.IllumMaker = NIKON And g_blnOptDebOffline = False Then
        g_dblMaxLux = ReadOptLux(True)
    Else
        g_dblMaxLux = 20000
    End If

    '######### Disable[Exit]
    cmd_exit.enabled = False
    cmd_start.Caption = "Stop"
    
    '######### Adjusted to False
    Dim intOptCondLoop As Integer
    For intOptCondLoop = 0 To OptCond.OptCondNum
        g_blnOptCondAdjusted(intOptCondLoop) = False
        If OptCond.CondInfoI(intOptCondLoop).AxisLevel > 0 Then
            If sub_UpdateOpt(OptCond.CondInfoI(intOptCondLoop).OptIdentifier, OptCond.CondInfoI(intOptCondLoop).AxisLevel, "Lux", , True) = False Then
                Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Failed[sub_UpdateOpt]!", True)
                GoTo optexit
            End If

        ElseIf OptCond.CondInfoI(intOptCondLoop).WedgeFilter > 0 Then
            If sub_UpdateOpt(OptCond.CondInfoI(intOptCondLoop).OptIdentifier, OptCond.CondInfoI(intOptCondLoop).WedgeFilter, "Wedge", , True) = False Then
                Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Failed[sub_UpdateOpt]!", True)
                GoTo optexit
            End If
        End If
    Next
    
    '#########Set RunOption
    If g_blnOptDebOffline = False Then
        Call sub_exec_DoAll(True)
    End If
    
    '最大LOOP回数の取得
    g_MaxPalsCount = frm_PALS_OptAdj_Main.cbo_AveNum.Value

    Dim index As Long
    Dim sitez As Long
    '特性値データ配列数を項目数で再定義
    For index = 0 To PALS.CommonInfo.TestCount
        For sitez = 0 To nSite
            Call PALS.CommonInfo.TestInfo(index).site(sitez).sub_ChangeDataDivision(g_MaxPalsCount)
        Next sitez
    Next index
    
    '#################################################
    '#################################################
    '########################Main Lutin ForNormalAdjust
    'Loop All Test
    
    Dim lngNowLoopCnt As Long               '現在のループ回数
    Dim lngNowAveCnt As Long                '現在の平均回数
    Dim intFileNo As Integer                'ファイル番号
    Dim DatalogPosi As DatalogPosition      'データログの各項目データ位置を保存する構造体
    Dim flg_ShtAdj As Boolean               'SH調整かどうかのフラグ
    
    flg_ShtAdj = False
    strOptIdenShTgt = ""

    '>>>2011/6/13 M.IMAMURA ADD
    g_blnOptAdjusting = False
    Set OptCond = Nothing
    If g_blnOptDebOffline = True Then
        Set OptCond = New csPALS_OptCond
    End If
    '<<<2011/6/13 M.IMAMURA ADD

AdjStart:
    '########################
    '######### AdjustLoop
    '########################
    For lngNowLoopCnt = 1 To val(cbo_AdjNum.Value)
        'PALSのエラーフラグがTrueになっていた場合、測定終了
        If g_ErrorFlg_PALS Then
            Exit For
        End If
        
        '######### Set DataLog
        If g_blnOptDebOffline = False Then
            Call sub_set_datalog(False)
            If flg_ShtAdj Then
                Call sub_set_datalog(True, PALS_PARAMFOLDERNAME_OPT, "OptAdjDataSh")
            Else
                Call sub_set_datalog(True, PALS_PARAMFOLDERNAME_OPT, "OptAdjData")
            End If
        End If
        
        '>>>2011/6/13 M.IMAMURA Del.
        '######### Reset OptCondition
'        Set OptCond = Nothing
'        Set OptCond = New csPALS_OptCond
        '<<<2011/6/13 M.IMAMURA Del.
        
        '#########################
        '######### AverageLoop
        '#########################
       For lngNowAveCnt = 1 To val(cbo_AveNum.Value)
            'PALSのエラーフラグがTrueになっていた場合、測定終了
            If g_ErrorFlg_PALS Then
                Exit For
            End If

            Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Now TestRunning...  Try " & CStr(lngNowLoopCnt) & " / " & cbo_AdjNum.Value & " (Ave" & CStr(lngNowAveCnt) & "/" & CStr(cbo_AveNum.Value) & ")")
            
            '######### Run Test
            If g_blnOptDebOffline = False Then Call sub_exec_run
                        
            '>>>2011/6/13 M.IMAMURA ADD
            g_blnOptAdjusting = True
            '<<<2011/6/13 M.IMAMURA ADD
            
            '######### Read LogFile
            If lngNowAveCnt = 1 Then
                '######### OpenLogFile
                intFileNo = FreeFile                        'ファイル番号の取得
                If g_blnOptDebOffline = False Then
                    Open g_strOutputDataText For Input As #intFileNo
                Else
                    Open g_strOptDataTextDeb For Input As #intFileNo
                End If
            End If
            mSecSleep (500)
            '>>>2011/06/13 M.IMAMURA ContFailFlg Add.
            Call sub_ReadDatalog(lngNowAveCnt, intFileNo, DatalogPosi, True)
            '<<<2011/06/13 M.IMAMURA ContFailFlg Add.
        Next
        
        '######### Reset DataLog
        If g_blnOptDebOffline = False Then Call sub_set_datalog(False)
        '######### Close LogFile
        Close #intFileNo

        'Write OptParameter to Datalog
        Call sub_OutPutOptParam

        '#################################################
        '######### Write Log & Check Data & UpdateOpt
        '#################################################
        If g_blnOptStop = True Or lngNowLoopCnt = val(cbo_AdjNum.Value) Then
            blnOptUpdate = False
        Else
            blnOptUpdate = True
        End If
        
        If sub_CheckOptTarget(lngNowLoopCnt, blnOptUpdate, strOptIdenShTgt) = True Then
            'CheckOK
            blnOptCheck = True
            Exit For
        Else
            'CheckNG
            If g_blnOptDebOffline <> True And blnOptUpdate = True And g_ErrorFlg_PALS = False And OptCond.IllumMaker <> KESILLUM Then
                Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Running OptIni...  ")
            '>>>2011/06/13 M.IMAMURA OptReset Mod.
                '>>>2011/07/29 M.IMAMURA PALSLogReset Add.
                Set PALS = Nothing
                Set PALS = New csPALS
                Call ReadCategoryData
                For index = 0 To PALS.CommonInfo.TestCount
                    For sitez = 0 To nSite
                        Call PALS.CommonInfo.TestInfo(index).site(sitez).sub_ChangeDataDivision(g_MaxPalsCount)
                    Next sitez
                Next index
                '<<<2011/07/29 M.IMAMURA PALSLogReset Add.
                Set OptCond = Nothing
'                First_Exec = 0
'                Call sub_run_Optini
            '>>>2011/06/13 M.IMAMURA OptReset Mod.
            End If
        End If
        
        If g_blnOptStop = True Then Exit For
        If g_ErrorFlg_PALS = True Then Exit For
        If g_blnOptDebOffline = True Then GoTo exitsh
    
    Next
    '########################Main Lutin ForNormalAdjust
    '#################################################
    '#################################################
    
    If flg_ShtAdj = True Then GoTo ShAdjEnd
    
    '#################################################
    '#################################################
    '########################Main Lutin For SH & X****
    If blnOptCheck = True Then
        '##　調整済み項目をX倍へ反映
        For intOptCondLoop = 0 To OptCond.OptCondNum
            If g_blnOptCondAdjusted(intOptCondLoop) = True And OptCond.CondInfoI(intOptCondLoop).AxisLevel <> -1 Then
                '##　Xチェック 計算 書き込み
                Call sub_XOptCalculate(OptCond.CondInfoI(intOptCondLoop).OptIdentifier)
            End If
        Next
        
        If OptCond.IllumMaker = INTERACTION Or OptCond.IllumMaker = KESILLUM Or gblnForCis = True Then
            GoTo exitsh
        End If
    
        '##　PALSOptループ
        Dim intOptTestCondLoop As Integer
        Dim strOptIdenShLess As String
        Dim dblShLuxBack As Double
        Dim intShWedgeBack As Integer
        Dim intShNdBack As Integer
        Dim strShtCondBack As String
        
        For intOptTestCondLoop = 0 To PALS.CommonInfo.TestCount
        strOptIdenShTgt = PALS.CommonInfo.TestInfo(intOptTestCondLoop).OptIdentifier
        '##　SHチェック
        If InStr(1, strOptIdenShTgt, "SH") > 0 Then
        If g_blnOptCondAdjusted(OptCond.CondInfoNo(strOptIdenShTgt)) = True Then
            '##　SH無し文字列生成
            strOptIdenShLess = Left(strOptIdenShTgt, InStr(strOptIdenShTgt, "SH") - 1)
            If InStr(1, strOptIdenShTgt, "SH") + 1 < Len(strOptIdenShTgt) Then
                strOptIdenShLess = strOptIdenShLess & Right(strOptIdenShTgt, Len(strOptIdenShTgt) - InStr(strOptIdenShTgt, "SH") - 1)
            End If
            
            '##　PALSOptループ
            For intOptCondLoop = 0 To OptCond.OptCondNum
            '##　SHチェック 調整済チェック
            If OptCond.IllumMaker = NIKON And g_blnOptCondAdjusted(intOptCondLoop) = False And OptCond.CondInfoI(intOptCondLoop).OptIdentifier = strOptIdenShLess Then
            
                    '##　CSTチェック
                    If sub_SheetNameCheck(CONDSHTNAME) = False Then
                    '>>>2011/6/13 M.IMAMURA Error Changed
'                        Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "NotFoundSheet[ConditionSetTable]! ErrCode.4-2-02-8-12", True)
'                        GoTo optexit
                        Call sub_errPALS("NotFoundSheet[" & CONDSHTNAME & "]!!", "4-2-02-8-12")
                        blnOptCheck = False
                        g_blnOptStop = True
                        g_ErrorFlg_PALS = False
                        GoTo exitsh
                    '<<<2011/6/13 M.IMAMURA Error Changed
                    End If
                    
                    '##　CST検索
                    '>>>2011/5/17 M.IMAMURA OPTSET Search
                    Set OptPoint = Nothing
                    Set OptPoint = Worksheets(CONDSHTNAME).Range("C3:Z3").Find("OPTSET")
                    If OptPoint Is Nothing Then
                    '>>>2011/6/13 M.IMAMURA Error Changed
'                        Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "NotFoundParameter[OPTSET]@[ConditionSetTable]! ErrCode.4-2-02-8-13", True)
'                        GoTo optexit
                        Call sub_errPALS("NotFoundParameter[OPTSET]@[" & CONDSHTNAME & "]!", "4-2-02-8-13")
                        blnOptCheck = False
                        g_ErrorFlg_PALS = False
                        GoTo exitsh
                    '<<<2011/6/13 M.IMAMURA Error Changed
                    End If
                    '<<<2011/5/17 M.IMAMURA OPTSET Search
                    For IdenRow = 3 To 65536
                        If Worksheets(CONDSHTNAME).Cells(IdenRow, OptPoint.Column).Value = strOptIdenShTgt Then
                            Exit For
                        End If
                    Next
                    If IdenRow >= 65536 Then
                    '>>>2011/6/13 M.IMAMURA Error Changed
'                        Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "NotFoundParameter[" & strOptIdenShTgt & "]@[ConditionSetTable]! ErrCode.4-2-02-4-14", True)
'                        GoTo optexit
                        Call sub_errPALS("NotFoundParameter[" & strOptIdenShTgt & "]@[" & CONDSHTNAME & "]!", "4-2-02-4-14")
                        blnOptCheck = False
                        g_ErrorFlg_PALS = False
                        GoTo exitsh
                    '<<<2011/6/13 M.IMAMURA Error Changed
                    End If
                    Set ShtPoint = Nothing
                    Set ShtPoint = Worksheets(CONDSHTNAME).Range("C3:Z3").Find("SHUTTER")
                    If ShtPoint Is Nothing Then
                    '>>>2011/6/13 M.IMAMURA Error Changed
'                        Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "NotFoundParameter[SHUTTER]@[ConditionSetTable]! ErrCode.4-2-02-8-15", True)
'                        GoTo optexit
                        Call sub_errPALS("NotFoundParameter[SHUTTER]@[" & CONDSHTNAME & "]!", "4-2-02-8-15")
                        blnOptCheck = False
                        g_ErrorFlg_PALS = False
                        GoTo exitsh
                    '<<<2011/6/13 M.IMAMURA Error Changed
                    End If
                    
                    '##　CSTバックアップ
                    strShtCondBack = Worksheets(CONDSHTNAME).Cells(IdenRow, ShtPoint.Column).Value
                    
                    '##　CST変更&読み込み
                    If Worksheets(CONDSHTNAME).Cells(1, 2).Value = "ConditionSetTable_ShutterOnly" Then
                        Worksheets(CONDSHTNAME).Cells(IdenRow, ShtPoint.Column).Value = "SHUT_OFF"
                        Call Excel.Application.Run("Read_conditionsetSh")
                    Else
                        Worksheets(CONDSHTNAME).Cells(IdenRow, ShtPoint.Column).Value = "OFF"
                        Call Excel.Application.Run("Read_conditionset")
                    End If
                     
                    '##　SH光量バックアップ
                    dblShLuxBack = OptCond.CondInfo(strOptIdenShTgt).AxisLevel
                    intShWedgeBack = OptCond.CondInfo(strOptIdenShTgt).WedgeFilter
                    intShNdBack = OptCond.CondInfo(strOptIdenShTgt).NDFilter
                    
                    If sub_UpdateOpt(strOptIdenShTgt, OptCond.CondInfo(strOptIdenShTgt).AxisLevel / 4, "Lux", False) = False Then
                        Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Failed[sub_UpdateOpt]!", True)
                        GoTo optexit
                    End If
                    
                    '##　光量調整
                    blnOptCheck = False
                    flg_ShtAdj = True
                    GoTo AdjStart
ShAdjEnd:
                    '##　CST戻し
                    Worksheets(CONDSHTNAME).Cells(IdenRow, ShtPoint.Column).Value = strShtCondBack
                    strShtCondBack = ""
                    
                    '##　CST読み込み
                    If Worksheets(CONDSHTNAME).Cells(1, 2).Value = "ConditionSetTable_ShutterOnly" Then
                        Call Excel.Application.Run("Read_conditionsetSh")
                    Else
                        Call Excel.Application.Run("Read_conditionset")
                    End If
                    
                    '>>>2011/4/20 M.IMAMURA ADD
                    '##　光量差異をチェック
                    If Abs(OptCond.CondInfo(strOptIdenShTgt).AxisLevel - dblShLuxBack) < dblShLuxBack * 0.1 Then
                        Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "ShutterLuxCheck Error@" & strOptIdenShTgt & "! ErrCode.4-2-02-5-16", True)
                        GoTo optexit
                    End If
                    '<<<2011/4/20 M.IMAMURA ADD
                    
                    
                    '##　SH光量を戻し
                    If sub_UpdateOpt(strOptIdenShTgt, dblShLuxBack, "Lux", True) = False Then
                        Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Failed[sub_UpdateOpt]!", True)
                        GoTo optexit
                    End If
                    
                    '##　SH無しに書き込み
                    If blnOptCheck = True Then
                    If sub_UpdateOpt(strOptIdenShLess, OptCond.CondInfo(strOptIdenShTgt).AxisLevel, "Lux", True) = False Then
                        Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Failed[sub_UpdateOpt]!", True)
                        GoTo optexit
                    End If
                    End If
                    
                    '######### Reset OptCondition
                    Set OptCond = Nothing
                    Set OptCond = New csPALS_OptCond

                    '##　SH無しのXチェック 計算 書き込み
                    If blnOptCheck = True Then Call sub_XOptCalculate(strOptIdenShLess)
                    
                    '>>>2011/4/20 M.IMAMURA ADD
                    '>>>2011/06/13 M.IMAMURA OptReset Mod.
                    Set OptCond = Nothing
'                    First_Exec = 0
                    '<<<2011/06/13 M.IMAMURA OptReset Mod.
                    '<<<2011/4/20 M.IMAMURA ADD
                    
                    '##　動作完了チェック
                    If blnOptCheck = False Or g_blnOptStop = True Or g_ErrorFlg_PALS = True Then Exit For
            End If
            Next
            
            '##　動作完了チェック
            If blnOptCheck = False Or g_blnOptStop = True Or g_ErrorFlg_PALS = True Then Exit For
        
        End If
        End If
        Next
    
    End If
    '########################Main Lutin For SH & X****
    '#################################################
    '#################################################
exitsh:
    '>>> 2012/3/9 M.Imamura ConnectLoop Add.
    If blnOptCheck = True Then
        Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "OptAdjust Finished !", , True)
        If cb_ConnectLoop.Value = True Then
            mSecSleep (1000)
            Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Now Connecting To LoopTool....")
            mSecSleep (1000)
            Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Now Doing LoopTool....")
            If FLG_PALS_DISABLE.LoopAdj = False Then
                Call frm_PALS.sub_ShowForm(PALS_PARAMFOLDERNAME_LOOP)
                Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Loop Finished!!....")
            Else
                Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Disble LoopTool!", True)
            End If
        End If
    End If
    
    If blnOptCheck = False Then Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "OptAdjust Failed !", True)
    If g_blnOptStop = True Then Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "OptAdjust Stoped !", , True)
    If g_ErrorFlg_PALS = True Then Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "PALS Error !", True)

optexit:
    cmd_start.Caption = "Start"
    cmd_exit.enabled = True
    cmd_start.enabled = True
    If strShtCondBack <> "" Then
        Call sub_errPALS("Warning!! Shutter Condition Chenged [" & strShtCondBack & "]->[OFF]" & "@Sheet[" & CONDSHTNAME & "]", "4-1-02-7-02")
    End If
    '>>>2011/6/13 M.IMAMURA ADD
    g_blnOptAdjusting = False
    '<<<2011/6/13 M.IMAMURA ADD
Exit Sub

errPALScmd_start_Click:
    If strShtCondBack <> "" Then
        Call sub_errPALS("Warning!! Shutter Condition Chenged [" & strShtCondBack & "]->[OFF]" & "@Sheet[" & CONDSHTNAME & "]", "4-1-02-7-03")
    End If
    Call sub_errPALS("OptAdjust Tool Run error at 'cmd_Start_Click'", "4-1-02-0-04")
    '>>>2011/6/13 M.IMAMURA ADD
    g_blnOptAdjusting = False
    '<<<2011/6/13 M.IMAMURA ADD
    
End Sub


Private Sub cmd_exit_Click()
    Unload frm_PALS_OptAdj_Main
End Sub


Private Sub Btn_ContinueOnFail_Click()
    Btn_StopOnFail.Value = False
    Btn_ContinueOnFail.Value = True
End Sub

Private Sub Btn_StopOnFail_Click()
    Btn_StopOnFail.Value = True
    Btn_ContinueOnFail.Value = False
End Sub

Private Sub op_AutoAdjust_Click()
    txt_maxwait.enabled = True
    txt_maxtrial_num.enabled = True
End Sub

Private Sub op_NotAdjust_Click()
    txt_maxwait.enabled = False
    txt_maxtrial_num.enabled = False
End Sub

Private Sub txt_maxtrial_num_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txt_maxtrial_num_Change()
    txt_loop_num.Text = val(txt_loop_num.Text)
    If txt_loop_num.Text = "0" Then txt_loop_num.Text = "1"
End Sub

Private Sub txt_maxwait_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txt_maxwait_Change()
    txt_loop_num.Text = val(txt_loop_num.Text)
    If txt_loop_num.Text = "0" Then txt_loop_num.Text = "1"
End Sub

Private Sub txt_loop_num_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txt_loop_num_Change()
    txt_loop_num.Text = val(txt_loop_num.Text)
    If txt_loop_num.Text = "0" Then txt_loop_num.Text = "1"
'    If Val(txt_loop_num.Text) > 100 Then txt_loop_num.Text = "100"
End Sub


