VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_PALS_LoopAdj_Main 
   Caption         =   "PALS - Auto Loop Parameter Adjust"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   OleObjectBlob   =   "frm_PALS_LoopAdj_Main.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frm_PALS_LoopAdj_Main"
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

Private Sub UserForm_Activate()
    Dim hWnd As Long
    Dim Wnd_STYLE As Long

    hWnd = GetActiveWindow()
    Wnd_STYLE = GetWindowLong(hWnd, GWL_STYLE)
    Wnd_STYLE = Wnd_STYLE And (Not WS_SYSMENU)
    SetWindowLong hWnd, GWL_STYLE, Wnd_STYLE
    DrawMenuBar hWnd
    Me.Caption = LOOPTOOLNAME & " Ver:" & LOOPTOOLVER
End Sub


'********************************************************************************************
' 名前 : UserForm_Initialize
' 内容 : ユーザーフォーム出力時の初期化関数
' 引数 : なし
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
' 更新履歴： Rev1.1      2012/03/09　光量調整からの連携対応   M.Imamura
'********************************************************************************************
Private Sub UserForm_Initialize()
    g_blnLoopStop = False
    Call sub_PalsFileCheck(PALS_PARAMFOLDERNAME_LOOP)
    
    '>>>2011/8/31 M.IMAMURA Add. ForCIS
    If gblnForCis = True Then
        op_AutoAdjust.enabled = False
        op_NotAdjust.Value = True
        Btn_ContinueOnFail.Value = True
    End If
    '>>>2011/8/31 M.IMAMURA Add. ForCIS

    '>>>2012/3/9 M.IMAMURA Add. For Connection From OptAdj
    If FLG_PALS_RUN.OptAdj = True Then
        With frm_PALS_LoopAdj_Main
            .txt_loop_num.Value = frm_PALS_OptAdj_Main.txt_loop_num.Text
            
            If frm_PALS_OptAdj_Main.op_AutoAdjust.Value = False Then
                .op_NotAdjust.Value = True
                .op_AutoAdjust.Value = False
            Else
                .op_NotAdjust.Value = False
                .op_AutoAdjust.Value = True
            End If
            
            .Btn_ContinueOnFail.Value = frm_PALS_OptAdj_Main.Btn_ContinueOnFail.Value
            .txt_maxwait = frm_PALS_OptAdj_Main.txt_maxwait.Text
            .txt_maxtrial_num = frm_PALS_OptAdj_Main.txt_maxtrial_num.Text
            Call .cmd_start_Click
        End With
    End If
    '<<<2012/3/9 M.IMAMURA Add. For Connection From OptAdj

End Sub



'********************************************************************************************
' 名前 : cmd_Start_Click
' 内容 : LOOP調整開始ボタンクリックした際の動作
'        測定準備⇒測定⇒傾向判断⇒帳票出力の流れを制御している
' 引数 : なし
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Sub cmd_start_Click()
    
    If Not sub_CheckTestInstancesParams Then
        Exit Sub
    End If
    
    Dim Flg_AnalyzeDebug As Boolean

    If sub_CheckResultFormat = False Then
        Exit Sub
    End If
    
    If frm_PALS_LoopAdj_Main.Chk_DebugMode = True Then
        Flg_AnalyzeDebug = True
        frm_PALS_LoopAdj_Main.chk_IGXL_Check.Value = True
    End If
                
On Error GoTo errPALScmd_start_ClickLoop
        
    Dim intLoopTrialCnt As Integer      'ループ試行回数を示す変数
    Dim MeasureDatalogInfo As DatalogInfo
        
    '途中STOPを行った際の処理
    If cmd_start.Caption = "Stop" Then
        If MsgBox("Pushed [Stop] Button" & vbCrLf & "Do You Want Stop?", vbYesNo, LOOPTOOLNAME) = vbYes Then
            g_blnLoopStop = True
            Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Canceled Loop...  ")
            cmd_start.enabled = False
        End If
        Exit Sub
    End If
        
    '初期化
    intLoopTrialCnt = 0
    
    'TestConditionに設定してある各カテゴリのWaitが、フォームで指定した最大Wait以上になっていないかチェックする
    If Not sub_CheckTestConditionWaitData(val(frm_PALS_LoopAdj_Main.txt_maxwait)) Then
        Exit Sub
    End If
    
    'LOOP調整を開始するかの確認
    '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    If g_RunAutoFlg_PALS = False And frm_PALS_OptAdj_Main.cb_ConnectLoop.Value = False Then
        If MsgBox("Pushed [Start] Button, Ready?", vbOKCancel, LOOPTOOLNAME) = vbCancel Then
            Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Canceled...  ")
            Exit Sub
        End If
    End If
    '<<<2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    
    'フォームの状態の変更
    cmd_exit.enabled = False
    cmd_start.Caption = "Stop"

    'カテゴリ情報を格納しておく構造体の、配列数再定義、初期化
    ReDim ChangeParamsInfo(PALS.LoopParams.CategoryCount)
    Call sub_Init_ChangeLoopParamsInfo

'バラツキがあり再調整する場合、このフラグへ飛ぶ
LOOP_RETRY:

    'ループ試行回数を示す変数をインクリメント
    intLoopTrialCnt = intLoopTrialCnt + 1

    '最大LOOP回数の取得
    g_MaxPalsCount = frm_PALS_LoopAdj_Main.txt_loop_num.Text

g_AnalyzeIgnoreCnt = 5

    Dim index As Long
    Dim sitez As Long
    '特性値データ配列数を項目数で再定義
    For index = 0 To PALS.CommonInfo.TestCount
        For sitez = 0 To nSite
            Call PALS.CommonInfo.TestInfo(index).site(sitez).sub_ChangeDataDivision(g_MaxPalsCount)
        Next sitez
    Next index


    If Not Flg_AnalyzeDebug Then
        '######### Set DataLog
        'データログファイル名の設定
        Call sub_set_datalog(False)
        Call sub_set_datalog(True, PALS_PARAMFOLDERNAME_LOOP, "LoopAdjData")
        
        '######### Set RunOption
        'RunOptionをContinue On Failに変更
        Call sub_exec_DoAll(True)
    End If
    
    
    '#################################################
    '##############   Main Measure   #################
    '#################################################
    
    Dim lngNowLoopCnt As Long               '現在のループ回数
    Dim intFileNo As Integer                'ファイル番号
    Dim DatalogPosi As DatalogPosition      'データログの各項目データ位置を保存する構造体
    
    'ユーザーフォームで指定した回数分、繰り返す
    For lngNowLoopCnt = 1 To val(txt_loop_num.Text)
''        mSecSleep (100)

'>>>2011/05/16 K.SUMIYASHIKI ADD
        Call sub_InitActiveSiteInfo
'<<<2011/05/16 K.SUMIYASHIKI ADD

        If Flg_AnalyzeDebug Then
            g_strOutputDataText = frm_PALS_LoopAdj_Main.txt_AnalyzeDataPath.Text
            If Len(g_strOutputDataText) = 0 Then
                MsgBox ("Input analyze data path!!")
                Exit Sub
            End If
'            g_strOutputDataText = ""
        End If
        
        'PALSのエラーフラグがTrueになっていた場合、測定終了
        If g_ErrorFlg_PALS Then
            Exit For
        End If
        
        'フォームの進捗状況欄更新
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now TestRunning...  " & CStr(lngNowLoopCnt) & " / " & txt_loop_num.Text)
        
        If Not Flg_AnalyzeDebug Then
            '######### Run Test
            'IG-XLのRunを実行
            Call sub_exec_run
        End If
'        mSecSleep (500)
        mSecSleep (300)
    
        Dim lngDatalogFileValue As Long
        Dim WaitCnt As Long
        lngDatalogFileValue = 0
        
        For WaitCnt = 0 To 50
            If lngDatalogFileValue = FileLen(g_strOutputDataText) Then
                mSecSleep (100)
                If lngDatalogFileValue = FileLen(g_strOutputDataText) Then
                    Exit For
                End If
            Else
                lngDatalogFileValue = FileLen(g_strOutputDataText)
                mSecSleep (100)
            End If
        Next WaitCnt
    
        '######### 1回目の測定時のみ、データログを開く(読み取りモード指定)
        If lngNowLoopCnt = 1 Then
                        
            'ファイル番号の取得
            intFileNo = FreeFile
            
            '測定データログをInput(読み込み)モードで開く
            Open g_strOutputDataText For Input As #intFileNo
        End If
        
            
        '######### 測定データログの読込
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Reading...")
        '>>>2011/06/13 M.IMAMURA ContFailFlg Add.
        '>>>2011/08/04 K.SUMIYASHIKI UPDATE.
        'FAILした項目のデータを読み取りたい場合
        If frm_PALS_LoopAdj_Main.Btn_ContinueOnFail.Value = True Then
            Call sub_ReadDatalog(lngNowLoopCnt, intFileNo, DatalogPosi, True)
        'FAILした項目のデータは無視する場合
        Else
            Call sub_ReadDatalog(lngNowLoopCnt, intFileNo, DatalogPosi, False)
        End If
        '<<<2011/08/04 K.SUMIYASHIKI UPDATE.
        '<<<2011/06/13 M.IMAMURA ContFailFlg Add.
        
        '######### 途中STOP時の処理
        If g_blnLoopStop Then
            txt_loop_num.Text = CStr(lngNowLoopCnt)
            lngNowLoopCnt = lngNowLoopCnt
            Exit For
        End If

        '######### データ解析＆リトライ
        '30回目以降から傾向分析を開始し、その後1回おきにデータ解析を行う
        '最大ループ試行回数に達している場合は、データ解析は行わない
        If (lngNowLoopCnt >= FIRST_VARIATION_CHECK_CNT) And (lngNowLoopCnt Mod VARIATION_CHECK_STEP = 0) _
            And intLoopTrialCnt < val(frm_PALS_LoopAdj_Main.txt_maxtrial_num.Text) Then
            
            '傾向確認を行うモード(デフォルト)に指定されている場合、データ解析を行う
            If op_AutoAdjust.Value Then
            
                '######### Analyze LoopData
                '3σ/規格幅が規定値より大きい項目がないかチェック
                '1項目でも大きい項目があれば、Falseが返る
                If Not sub_CheckLoopData(lngNowLoopCnt) Then
                
                    '各カテゴリのパラメータを傾向に応じて変更、TestConditionシートの値も変更
                    If sub_UpdataLoopParams = False Then
                        Call sub_errPALS("Updata LoopParameter error at 'sub_UpdataLoopParams'", "2-1-01-0-01")
                        Exit For
                    End If
                    
                    '今回の傾向・取り込み回数・Wait情報を、ChangeParamsInfoに保存
                    '取り込み回数・Waitの変更１カテゴリでもあった場合は、Trueが返る
                    If sub_Update_ChangeLoopParamsInfo Then
                    
                        'ファイル(測定データログ)を閉じる
                        Close #intFileNo
                        
                        'TestConditionシート内のデータを測定データログの末尾に追加
                        Call sub_OutPutLoopParam(MeasureDatalogInfo)
                        
                        If Not Flg_AnalyzeDebug Then
                            'データログの設定をクリア
                            Call sub_set_datalog(False)
                        End If
                        
                        'csPALSクラスの解放
                        Set PALS = Nothing
                        
                        'csPALSクラスを再定義
                        Set PALS = New csPALS
                        
                        'TestConditionシートデータの再読込
                        Call ReadCategoryData
                        
                        'フラグの初期化
                        g_blnLoopStop = False
                        g_ErrorFlg_PALS = False
                        
                        '再測定実施
                        GoTo LOOP_RETRY
                    
                    Else
                        '傾向確認を行うモード(フォームのボタン)をFalseに変更
                        op_AutoAdjust.Value = False
                    
                    End If
                    
                End If
            End If
        End If
    Next lngNowLoopCnt
    
    'ファイル(測定データログ)を閉じる
    Close #intFileNo

    cmd_start.enabled = False
    
    '#################################################
    '#################################################

    'TestConditionシート内のデータを測定データログの末尾に追加
    Call sub_OutPutLoopParam(MeasureDatalogInfo)
    
    '######### Make LoopResultSheet
    'フォームの進捗状況欄更新
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Making LoopResultSheet...")
    
    'LOOP帳票出力
    '正しい測定回数を出力させる為、途中STOPを行った場合と、通常測定時の対応を分けている
    If g_blnLoopStop Then
        '途中STOP時
        Call sub_ShowLoopData(lngNowLoopCnt, MeasureDatalogInfo)
    Else
        '通常測定時
        Call sub_ShowLoopData(lngNowLoopCnt - 1, MeasureDatalogInfo)
    End If
    
    '######### Reset DataLog
    If Not Flg_AnalyzeDebug Then
        'データログの設定をクリア
        Call sub_set_datalog(False)
    End If
    
   'Average回数、Wait時間が最大に設定されていた場合、メッセージボックスを出し知らせる
    Call sub_LoopParamsCheck
    
    'フォームの進捗状況欄更新
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Finished...", , True)
    
    cmd_start.Caption = "Start"
    cmd_start.enabled = True
    cmd_exit.enabled = True
'    cmd_start.Enabled = True

Exit Sub

errPALScmd_start_ClickLoop:
    Call sub_errPALS("Loop Tool Run error at 'cmd_Start_Click'", "2-1-01-0-02")

    '既にファイルを開いていた場合、ファイルを閉じる
    If intFileNo <> 0 Then
        Close #intFileNo
    End If

'    cmd_start.Enabled = True
    
End Sub


'********************************************************************************************
' 名前 : cmd_readloopdata_Click
' 内容 : 指定データログのLOOP帳票を出力
' 引数 : なし
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub cmd_readloopdata_Click()
    
On Error GoTo errPALScmd_readloopdata_Click

    'LOOP帳票を出力させるデータログを選択
    Call sub_SetLoopData
    
    'データログを選択しなかった場合、関数を抜ける
    If g_strOutputDataText = "False" Then
        Exit Sub
    End If
    
    Dim lngNowLoopCnt As Long               '現在のループ回数
    Dim lngNowLoopEnd As Long               'データ数
    Dim intFileNo As Integer                'ファイル番号
    Dim DatalogPosi As DatalogPosition      'データログの各項目データ位置を保存する構造体
    Dim strbuf As String
    Dim MeasureDatalogInfo As DatalogInfo
    
    cmd_start.enabled = False
    
    '初期化
    lngNowLoopEnd = 0
    
    'ファイル番号の取得
    intFileNo = FreeFile
    
    '######### データログから測定回数を取得
    Open g_strOutputDataText For Input As #intFileNo
    Do Until EOF(intFileNo)
        Line Input #intFileNo, strbuf
        If strbuf = DATALOG_END Then
            lngNowLoopEnd = lngNowLoopEnd + 1
            
        ElseIf InStr(1, strbuf, "MEASURE DATE : ") <> 0 Then
            MeasureDatalogInfo.MeasureDate = sub_GetMeasureData(strbuf, "Date")
            
        ElseIf InStr(1, strbuf, "JOB NAME     : ") <> 0 Then
            MeasureDatalogInfo.JobName = sub_GetMeasureData(strbuf, "JobName")
            
        ElseIf InStr(1, strbuf, "SW_NODE      : ") <> 0 Then
            MeasureDatalogInfo.SwNode = sub_GetMeasureData(strbuf, "Node")
            
        End If
    Loop
    Close #intFileNo

    '測定データが無い場合、エラーを返し関数を抜ける
    If lngNowLoopEnd = 0 Then
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Not Data Found...", True)
        Exit Sub
    End If
    
    '最大の測定回数を取得
    g_MaxPalsCount = lngNowLoopEnd
    
    Dim index As Long                   'テスト項目を示すループカウンタ
    Dim sitez As Long                   'サイトを示すループカウンタ
    
    '特性値データ配列数を項目数で再定義
    For index = 0 To PALS.CommonInfo.TestCount
        For sitez = 0 To nSite
            Call PALS.CommonInfo.TestInfo(index).site(sitez).sub_ChangeDataDivision(g_MaxPalsCount)
        Next sitez
    Next index
    
    '######### データログを読み込む
    For lngNowLoopCnt = 1 To lngNowLoopEnd

'>>>2011/05/16 K.SUMIYASHIKI ADD
        Call sub_InitActiveSiteInfo
'<<<2011/05/16 K.SUMIYASHIKI ADD
        
        If lngNowLoopCnt = 1 Then
            intFileNo = FreeFile
            Open g_strOutputDataText For Input As #intFileNo
        End If
        '######### 測定データログの読込
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Reading...")
        '>>>2011/06/13 M.IMAMURA ContFailFlg Add.
        '>>>2011/08/04 K.SUMIYASHIKI UPDATE.
        'FAILした項目のデータを読み取りたい場合
        If frm_PALS_LoopAdj_Main.Btn_ContinueOnFail.Value = True Then
            Call sub_ReadDatalog(lngNowLoopCnt, intFileNo, DatalogPosi, True)
        'FAILした項目のデータは無視する場合
        Else
            Call sub_ReadDatalog(lngNowLoopCnt, intFileNo, DatalogPosi, False)
        End If
        '<<<2011/08/04 K.SUMIYASHIKI UPDATE.
        '<<<2011/06/13 M.IMAMURA ContFailFlg Add.
    Next lngNowLoopCnt
    
    Close #intFileNo

    '######### Make LoopResultSheet
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Making LoopResultSheet...")
    Call sub_ShowLoopData(g_MaxPalsCount, MeasureDatalogInfo)

    'フォームの状態の変更
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Finished...", , True)

    cmd_start.enabled = True

Exit Sub

errPALScmd_readloopdata_Click:
    Call sub_errPALS("Create Loop sheet error at 'cmd_readloopdata_Click'", "2-1-02-0-03")

    '既にファイルを開いていた場合、ファイルを閉じる
    If intFileNo <> 0 Then
        Close #intFileNo
    End If
    
    cmd_start.enabled = True

End Sub

Private Sub cmd_exit_Click()
    Unload frm_PALS_LoopAdj_Main
End Sub

Private Sub txt_loop_num_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txt_loop_num_Change()
    txt_loop_num.Text = val(txt_loop_num.Text)
    If txt_loop_num.Text = "0" Then txt_loop_num.Text = "1"
'    If Val(txt_loop_num.Text) > 100 Then txt_loop_num.Text = "100"
End Sub

Private Sub txt_lot_name_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    If KeyAscii >= 65 And KeyAscii <= 90 Then Exit Sub
    If KeyAscii >= 97 And KeyAscii <= 122 Then Exit Sub

    KeyAscii = 0

End Sub


'********************************************************************************************
' 名前: sub_ShowLoopData
' 内容: フォームのステータスを更新し、LOOP帳票を作成
' 引数: lngNowLoopCnt:測定データ数
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/20　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_ShowLoopData(ByVal lngNowLoopCnt As Long, ByRef MeasureDatalogInfo As DatalogInfo)
    
    'フォームの進捗状況欄更新
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Making...LoopResult")

    'LOOP帳票作成
    Call sub_MakeLoopResultSheet(lngNowLoopCnt, MeasureDatalogInfo)

End Sub

