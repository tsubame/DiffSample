VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' クラス名  ：  csPALS_Common.cls
' 概要      ：  PALSで共通に使用する情報をもつクラス
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/16　新規作成        K.Sumiyashiki
'==========================================================================

Private xTestInfo() As csPALS_TestInfo          'csPALS_TestInfoクラスの定義

Public TestnameInfoList As New Collection       'Tnameとインデックスを紐付ける為のコレクション

Public g_strTesterName As String                '測定ログのNode Name:以降のテスター名

'テストインスタンスのループ調整に関するパラメータの存在チェック変数
Private Type PALS_LoopParameterList
    Flg_CapCategory1   As Boolean
    Flg_CapCategory2   As Boolean
    Flg_LoopJudgeLimit As Boolean
    
    CapCategory1_X   As Integer
    CapCategory2_X   As Integer
    LoopJudgeLimit_X As Integer
End Type

'テストインスタンスの光量調整に関するパラメータの存在チェック変数
Private Type PALS_OptParameterList
    Flg_OptTarget     As Boolean
    Flg_OptIdentifier As Boolean
    Flg_OptJudgeLimit As Boolean

    OptTarget_X     As Integer
    OptIdentifier_X As Integer
    OptJudgeLimit_X As Integer
End Type
'テストインスタンスのWait調整に関するパラメータの存在チェック変数
Private Type PALS_WaitParameterList
    Flg_CapCategory1   As Boolean
    Flg_CapCategory2   As Boolean
    Flg_WaitAdjFlg     As Boolean
    
    CapCategory1_X   As Integer
    CapCategory2_X   As Integer
    WaitAdjFlg_X     As Integer
End Type

Private m_LoopParameter As PALS_LoopParameterList
Private m_OptParameter  As PALS_OptParameterList
Private m_WaitParameter  As PALS_WaitParameterList


'*********************************************************************************
' 関数名    ：  Class_Initialize
' 概要      ：  インスタンス生成時の初期化
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()

    'コレクションデータの初期化
    Call sub_InitCollection(TestnameInfoList)
    
    'テストインスタンスの2行目に記載されているPALS各パラメータ名の存在チェック
    Call sub_CheckPALSParameter

    'フローテーブルとテストインスタンスの情報取得
    Call sub_ReadTestInfo
    
End Sub


'*********************************************************************************
' 関数名    ：  Class_Terminate
' 概要      ：  インスタンス開放時の処理
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Terminate()
    Call sub_InitCollection(TestnameInfoList)
    Erase xTestInfo
End Sub


'*********************************************************************************
' プロパティ名  ：  TestInfoName
' 概要          ：  項目番号に対応するTestInfoNameオブジェクトを返すプロパティ
' 引数          ：  index :項目名を示すインデックス
' 備考          ：  読み取り専用
' 更新履歴      ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Public Property Get TestInfo(ByVal index As Integer) As csPALS_TestInfo
    Set TestInfo = xTestInfo(index)
End Property


'********************************************************************************************
' 名前: sub_CheckPALSParameter
' 内容: テストインスタンスの2行目に記載されているPALS各パラメータ名のチェック
'     　あった場合、各パラメータのフラグをTrueに変更
' 引数: なし
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_CheckPALSParameter()
    
    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALSsub_CheckPALSParameter
    
    '一致するパラメータがテストインスタンスに存在すれば、フラグをTrueに変更
    Const strRange As String = "A2:IV2"
    
    
   '*******LOOP TOOL******************************************************************************************************
    With m_LoopParameter
        '指定シート一致するパラメータがあった場合、各パラメータの存在フラグをTrueに変更
        .Flg_CapCategory1 = sub_FindPALSParameter(PALS_LOOPCATEGORY1, TEST_INSTANCES, strRange, .CapCategory1_X, True)
        .Flg_CapCategory2 = sub_FindPALSParameter(PALS_LOOPCATEGORY2, TEST_INSTANCES, strRange, .CapCategory2_X, True)
        .Flg_LoopJudgeLimit = sub_FindPALSParameter(PALS_LOOPJUDGELIMIT, TEST_INSTANCES, strRange, .LoopJudgeLimit_X, True)
        
        
        'ループツールに必要なパラメータが全て存在した場合、メインフォームのLOOPボタンを有効にする
        If sub_ModuleCheck("frm_PALS_LoopAdj_Main.frm") = True And (.Flg_CapCategory1 = True) And (.Flg_CapCategory2 = True) And (.Flg_LoopJudgeLimit = True) Then
            FLG_PALS_DISABLE.LoopAdj = False
        Else
            If sub_ModuleCheck("frm_PALS_WaitAdj_Main.frm") = False Then
                frm_PALS.cmd_LoopRun.ControlTipText = frm_PALS.cmd_LoopRun.ControlTipText & "NotFound [frm_PALS_WaitAdj_Main.frm]Form" & vbCrLf
            End If
            frm_PALS.cmd_LoopRun.ControlTipText = frm_PALS.cmd_LoopRun.ControlTipText & "Not Found Parameter at [" & TEST_INSTANCES & "]" & vbCrLf
            If .Flg_CapCategory1 = False Then
                frm_PALS.cmd_LoopRun.ControlTipText = frm_PALS.cmd_LoopRun.ControlTipText & "- CapCategory1 -" & vbCrLf
            End If
            If .Flg_CapCategory2 = False Then
                frm_PALS.cmd_LoopRun.ControlTipText = frm_PALS.cmd_LoopRun.ControlTipText & "- CapCategory2 -" & vbCrLf
            End If
            If .Flg_LoopJudgeLimit = False Then
                frm_PALS.cmd_LoopRun.ControlTipText = frm_PALS.cmd_LoopRun.ControlTipText & "- LoopJudgeLimit -" & vbCrLf
            End If
        End If
    End With
   '**********************************************************************************************************************

   '*******Wait TOOL******************************************************************************************************
    With m_WaitParameter
        '指定シート一致するパラメータがあった場合、各パラメータの存在フラグをTrueに変更
        .Flg_CapCategory1 = sub_FindPALSParameter(PALS_LOOPCATEGORY1, TEST_INSTANCES, strRange, .CapCategory1_X, True)
        .Flg_CapCategory2 = sub_FindPALSParameter(PALS_LOOPCATEGORY2, TEST_INSTANCES, strRange, .CapCategory2_X, True)
        .Flg_WaitAdjFlg = sub_FindPALSParameter(PALS_WAITADJFLG, TEST_INSTANCES, strRange, .WaitAdjFlg_X, True)
        
        
        'Waitツールに必要なパラメータが全て存在した場合、メインフォームのWaitボタンを有効にする
        If sub_ModuleCheck("frm_PALS_WaitAdj_Main.frm") = True And .Flg_CapCategory1 = True And .Flg_CapCategory2 = True And .Flg_WaitAdjFlg = True And sub_SheetNameCheck(WaitResultSheetName) = True And ((.Flg_CapCategory1 = True) And (.Flg_CapCategory2 = True)) Then
            FLG_PALS_DISABLE.WaitAdj = False
        Else
            If sub_ModuleCheck("frm_PALS_WaitAdj_Main.frm") = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "NotFound [frm_PALS_WaitAdj_Main.frm]Form" & vbCrLf
            End If
            frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "Not Found Parameter at [" & TEST_INSTANCES & "]" & vbCrLf
            If .Flg_CapCategory1 = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- CapCategory1 -" & vbCrLf
            End If
            If .Flg_CapCategory2 = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- CapCategory2 -" & vbCrLf
            End If
            If .Flg_WaitAdjFlg = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- WaitAdjFlg -" & vbCrLf
            End If
            If sub_SheetNameCheck(WaitResultSheetName) = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "NotFound [" & WaitResultSheetName & "]Sheet" & vbCrLf
            End If
            If .Flg_CapCategory1 = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- CapCategory1 -" & vbCrLf
            End If
            If .Flg_CapCategory2 = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- CapCategory2 -" & vbCrLf
            End If
        End If
    End With
   '**********************************************************************************************************************

   '*******OPT TOOL*******************************************************************************************************
    With m_OptParameter
        '指定シート一致するパラメータがあった場合、各パラメータの存在フラグをTrueに変更
        .Flg_OptTarget = sub_FindPALSParameter(PALS_OPTTARGET, TEST_INSTANCES, strRange, .OptTarget_X, True)
        .Flg_OptIdentifier = sub_FindPALSParameter(PALS_OPTIDENTIFIER, TEST_INSTANCES, strRange, .OptIdentifier_X, True)
        .Flg_OptJudgeLimit = sub_FindPALSParameter(PALS_OPTJUDGELIMIT, TEST_INSTANCES, strRange, .OptJudgeLimit_X, True)

        '光量調整ツールに必要なパラメータが全て存在した場合、メインフォームのOPTボタンを有効にする
        If sub_ModuleCheck("frm_PALS_OptAdj_Main.frm") = True And sub_SheetNameCheck(OptResultSheetName) = True And ((.Flg_OptTarget = True) And (.Flg_OptIdentifier = True) And (.Flg_OptJudgeLimit = True)) Then
            FLG_PALS_DISABLE.OptAdj = False
        Else
            If sub_SheetNameCheck(OptResultSheetName) = False Then
                frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "NotFound [" & OptResultSheetName & "]Sheet" & vbCrLf
            End If
            If sub_ModuleCheck("frm_PALS_OptAdj_Main.frm") = False Then
                frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "NotFound [frm_PALS_OptAdj_Main.frm]Form" & vbCrLf
            End If
            frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "Not Found Parameter at [" & TEST_INSTANCES & "]" & vbCrLf
            If .Flg_OptTarget = False Then
                frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptTarget -" & vbCrLf
            End If
            If .Flg_OptIdentifier = False Then
                frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptIdentifier -" & vbCrLf
            End If
            If .Flg_OptJudgeLimit = False Then
                frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptJudgeLimit -" & vbCrLf
            End If
        End If
    End With
   '**********************************************************************************************************************

   '*******VOLT TOOL*******************************************************************************************************
    '>>> 2011/5/6 M.Imamura
    If sub_SheetNameCheck(PinSheetnameChans) = True Then
        PinSheetname = PinSheetnameChans
    Else
        PinSheetname = PinSheetnameChannel
    End If
    '<<< 2011/5/6 M.Imamura
    If sub_SheetNameCheck(PinSheetname) = True And sub_SheetNameCheck(OutPutSheetname) = True And sub_SheetNameCheck(ReadSheetName) = True And sub_SheetNameCheck(ReadSheetNameInfo) = True And sub_ModuleCheck("PALS_VoltAdj_Mod.bas") = True And sub_ModuleCheck("frm_PALS_VoltAdj_Main.frm") = True Then
        FLG_PALS_DISABLE.VoltageAdj = False
    Else
        If sub_SheetNameCheck(PinSheetname) = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [" & PinSheetname & "]Sheet" & vbCrLf
        End If
        If sub_SheetNameCheck(OutPutSheetname) = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [" & OutPutSheetname & "]Sheet" & vbCrLf
        End If
        If sub_SheetNameCheck(ReadSheetName) = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [" & ReadSheetName & "]Sheet" & vbCrLf
        End If
        If sub_SheetNameCheck(ReadSheetNameInfo) = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [" & ReadSheetNameInfo & "]Sheet" & vbCrLf
        End If
        If sub_ModuleCheck("PALS_VoltAdj_Mod.bas") = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [PALS_VoltAdj_Mod]Bas" & vbCrLf
        End If
        If sub_ModuleCheck("frm_PALS_VoltAdj_Main.frm") = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [frm_PALS_VoltAdj_Main]Form" & vbCrLf
        End If
    End If
   '**********************************************************************************************************************

   '*******WAVE TOOL*******************************************************************************************************
    
    If sub_SheetNameCheck(WaveSetupSheetName) = True And sub_SheetNameCheck(WaveResultSheetName) = True And _
       sub_ModuleCheck("frm_PALS_WaveAdj_Main.frm") = True And sub_ModuleCheck("PALS_WaveAdj_mod_Common.bas") = True And sub_ModuleCheck("csPALS_WaveAdjust.cls") = True Then
        FLG_PALS_DISABLE.WaveAdj = False
    Else
        If sub_SheetNameCheck(WaveSetupSheetName) = False Then
            frm_PALS.cmd_WaveRun.ControlTipText = frm_PALS.cmd_WaveRun.ControlTipText & "NotFound [" & WaveSetupSheetName & "]Sheet" & vbCrLf
        End If
        If sub_SheetNameCheck(WaveResultSheetName) = False Then
            frm_PALS.cmd_WaveRun.ControlTipText = frm_PALS.cmd_WaveRun.ControlTipText & "NotFound [" & WaveResultSheetName & "]Sheet" & vbCrLf
        End If
        If sub_ModuleCheck("PALS_WaveAdj_mod_Common.bas") = False Then
            frm_PALS.cmd_WaveRun.ControlTipText = frm_PALS.cmd_WaveRun.ControlTipText & "NotFound [PALS_WaveAdj_mod_Common]Bas" & vbCrLf
        End If
        If sub_ModuleCheck("csPALS_WaveAdjust.bas") = False Then
            frm_PALS.cmd_WaveRun.ControlTipText = frm_PALS.cmd_WaveRun.ControlTipText & "NotFound [csPALS_WaveAdjust]Class" & vbCrLf
        End If
        If sub_ModuleCheck("frm_PALS_WaveAdj_Main.frm") = False Then
            frm_PALS.cmd_WaveRun.ControlTipText = frm_PALS.cmd_WaveRun.ControlTipText & "NotFound [frm_PALS_WaveAdj_Main]Form" & vbCrLf
        End If
    End If
   '**********************************************************************************************************************

   '*******BIAS TOOL*******************************************************************************************************
    
    If sub_ModuleCheck("frm_PALS_BiasAdj_Main.frm") = True And sub_ModuleCheck("PALS_BiasAdj_Mod.bas") = True Then
        FLG_PALS_DISABLE.BiasAdj = False
    Else
        If sub_ModuleCheck("frm_PALS_BiasAdj_Main.frm") = False Then
            frm_PALS.cmd_BiasRun.ControlTipText = frm_PALS.cmd_BiasRun.ControlTipText & "NotFound [frm_PALS_BiasAdj_Main]Form" & vbCrLf
        End If
        If sub_ModuleCheck("PALS_BiasAdj_Mod.bas") = False Then
            frm_PALS.cmd_BiasRun.ControlTipText = frm_PALS.cmd_BiasRun.ControlTipText & "NotFound [PALS_BiasAdj_Mod].Bas" & vbCrLf
        End If
    End If
   '**********************************************************************************************************************

   '*******TRACE TOOL*******************************************************************************************************
    
    If sub_ModuleCheck("frm_PALS_TraceAdj_Main.frm") = True And sub_ModuleCheck("PALS_TraceAdj_Mod.bas") = True And sub_ModuleCheck("PALS_TraceAcq_Mod.bas") = True Then
        FLG_PALS_DISABLE.TraceAdj = False
    Else
        If sub_ModuleCheck("frm_PALS_TraceAdj_Main.frm") = False Then
            frm_PALS.cmd_TraceRun.ControlTipText = frm_PALS.cmd_TraceRun.ControlTipText & "NotFound [frm_PALS_TraceAdj_Main]Form" & vbCrLf
        End If
        If sub_ModuleCheck("PALS_TraceAdj_Mod.bas") = False Then
            frm_PALS.cmd_TraceRun.ControlTipText = frm_PALS.cmd_TraceRun.ControlTipText & "NotFound [PALS_TraceAdj_Mod].Bas" & vbCrLf
        End If
        If sub_ModuleCheck("PALS_TraceAcq_Mod.bas") = False Then
            frm_PALS.cmd_TraceRun.ControlTipText = frm_PALS.cmd_TraceRun.ControlTipText & "NotFound [PALS_TraceAcq_Mod].Bas" & vbCrLf
        End If
    End If
   '**********************************************************************************************************************

Exit Sub

errPALSsub_CheckPALSParameter:
    Call sub_errPALS("Check PALS Parameter error at 'sub_CheckPALSParameter'", "0-5-01-0-14")

End Sub


'********************************************************************************************
' 名前: sub_FindPALSParameter
' 内容: strTargetで指定された文字列が、strTargetで指定されたセル範囲内に存在するかチェックする。
'　　　 存在した場合、パラメータ存在判断用のフラグをTrueにし、列位置を取得する。
' 引数: strTarget     :検索文字列
'       strWsName     :検索ワークシート名
'       strRange      :検索範囲
'       intPosi       :文字列が存在する列番号
'       blnPALS_Check :検索文字列がある列の1列目に"PALS"という文字列があるか確認するかのフラグ
'                      ⇒True:確認あり、False:確認なし
' 戻値: True          :一致文字列あり
'       False         :一致文字列なし
' 備考    ： 引数のintPosiはByrefで定義している為、内部で書き換わることに注意
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_FindPALSParameter(ByRef strTarget As String, ByVal strWsName As String, ByVal strRange As String, _
                                        Optional ByRef intPosi As Integer = 0, Optional blnPALS_Check As Boolean = False) As Boolean

Const PALSNAME As String = "PALS"

On Error GoTo errPALSsub_FindPALSParameter

    Dim obj As Object

    Set obj = Worksheets(strWsName).Range(strRange).Find(strTarget, lookat:=xlWhole, MatchCase:=True)
    
    If obj Is Nothing Then
        sub_FindPALSParameter = False
    Else
        intPosi = Worksheets(strWsName).Range(strRange).Find(strTarget, lookat:=xlWhole, MatchCase:=True).Column
    
        If blnPALS_Check Then
            If Worksheets(strWsName).Cells(1, intPosi).Value = PALSNAME Then
                sub_FindPALSParameter = True
            Else
                MsgBox "Error : " & strTarget & "not found." & vbCrLf & "         Please check PALS parameter" & vbCrLf _
                        & "         at 'Test Instances'" & "ErrCode.0-5-02-2-15", vbExclamation, PALS_ERRORTITLE
            End If
        Else
            sub_FindPALSParameter = True
        End If
    
    End If

Exit Function

errPALSsub_FindPALSParameter:
    Call sub_errPALS("Find PALS Parameter error at 'sub_FindPALSParameter'", "0-5-02-0-16")

End Function



'********************************************************************************************
' 名前: sub_ReadTestInfo
' 内容: Flow TableとTest InstancesのパラメータをcsLoopTestInfoクラスで定義されたLoopTestInfo配列に格納。
'       LoopTestInfoは項目数分の配列が動的に作成される。
' 引数: なし
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/10　新規作成   K.Sumiyashiki
'            Rev1.1      2011/06/06  処理追加   K.Sumiyashiki
'                                    ⇒Eee-JOBのテストインスタンス読み込み対応
'                                      規格に文字列が入っていた際の処理を追加
'********************************************************************************************
Public Sub sub_ReadTestInfo()

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

    On Error GoTo errPALSsub_ReadTestInfo

    Dim FT_y As Long        'フローテーブルのY方向セル位置
    Dim index As Long       '配列番号
    Dim intOptIdenCnt As Integer
    Dim intWaitAdjCnt As Integer
    
    intOptIdenCnt = 0
    intWaitAdjCnt = 0
    'Flow Tableシートの存在チェック
    If Not sub_SheetNameCheck(FLOW_TABLE) Then
        Call sub_errPALS("Flow Table not found at 'sub_ReadTestInfo'", "0-5-03-8-17")
        Exit Sub
    End If
    
    'フローテーブルの読み込み
    With Worksheets(FLOW_TABLE)
    
        '初期化
        FT_y = FT_START_Y
        index = 0
        
        'Opcodeが"set-device"になるまで繰り返す
        Do While (.Cells(FT_y, FT_OPCODE_X).Value <> FT_LASTROW_NAME)
        
            'TNumが入っている行のみを処理
            If Len(.Cells(FT_y, FT_TNUM_X).Value) Then
            
                '配列の再定義
                ReDim Preserve xTestInfo(index)
                
                'csLoopTestInfoクラスのインスタンス生成
                Set xTestInfo(index) = New csPALS_TestInfo
                
                '各パラメータを取得
                xTestInfo(index).tname = .Cells(FT_y, FT_TNAME_X).Value
                xTestInfo(index).Parameter = .Cells(FT_y, FT_PARAMETER_X).Value
                xTestInfo(index).Bin = .Cells(FT_y, FT_BIN_X).Value
                xTestInfo(index).OpCode = .Cells(FT_y, FT_OPCODE_X).Value
                
                'ラベルの取得。各ラベルの最初の項目にラベル情報を追加する
                If Len(.Cells(FT_y - 1, FT_LABEL_X).Value) > 0 Then
                    xTestInfo(index).Label = .Cells(FT_y - 1, FT_LABEL_X).Value
                ElseIf (Len(.Cells(FT_y - 1, FT_LABEL_X).Value) = 0) And (Len(.Cells(FT_y - 2, FT_LABEL_X).Value) > 0) _
                        And Len(.Cells(FT_y - 1, FT_TNUM_X).Value) = 0 Then
                    xTestInfo(index).Label = .Cells(FT_y - 2, FT_LABEL_X).Value
                End If
                
                'テスト名とインデックスを紐付けるコレクションにデータ追加(値:インデックス、キー:項目名)
                TestnameInfoList.Add Item:=index, key:=.Cells(FT_y, FT_TNAME_X).Value
                
                'インデックスのインクリメント
                index = index + 1
            End If
            'フローテーブル Y方向位置のインクリメント
            FT_y = FT_y + 1
        Loop
    End With


    Dim TI_y As Long                        'テストインスタンスのY方向セル位置
    Dim obj As Object

    index = 0

    'Test Instancesシートの存在チェック
    If Not sub_SheetNameCheck(TEST_INSTANCES) Then
        Call sub_errPALS("Test Instances not found at 'sub_ReadTestInfo'", "0-5-03-8-18")
    End If

    'テストインスタンスの読み込み
    With Worksheets(TEST_INSTANCES)
        
        'indexが項目数に達するまで繰り返し
        Do While (index <= UBound(xTestInfo))
            
            'テストインスタンスのB列に記載されているパラメータ名の検索
            'フローテーブルのParameterと一致するものがあれば処理を行う
            Set obj = .Range("B5:B10000").Find(xTestInfo(index).Parameter, lookat:=xlWhole, MatchCase:=True)
            
            '一致する項目が無い場合
            If obj Is Nothing Then
                MsgBox "TestInstances parameter name not equal FlowTable parameter name!" & vbCrLf & "ErrCode.0-5-03-2-19", vbExclamation, PALS_ERRORTITLE
'                MsgBox "フローテーブルとテストインスタンスの項目が一致しません。"
            Else
                TI_y = .Range("B5:B10000").Find(xTestInfo(index).Parameter, lookat:=xlWhole, MatchCase:=True).Row
                '各パラメータを取得
'>>>2011/06/06 K.SUMIYASHIKI UPDATE
'CIS対応
                If IsNumeric(.Cells(TI_y, TI_HIGHLIMIT_X).Value) = False Then
                    xTestInfo(index).HighLimit = 1     '上限規格
                    xTestInfo(index).LowLimit = 1       '下限規格
                    xTestInfo(index).Unit = ""               '単位
                    xTestInfo(index).arg2 = 3               'Arg2
    
                Else
                    xTestInfo(index).HighLimit = .Cells(TI_y, TI_HIGHLIMIT_X).Value     '上限規格
                    xTestInfo(index).LowLimit = .Cells(TI_y, TI_LOWLIMIT_X).Value       '下限規格
                    xTestInfo(index).Unit = .Cells(TI_y, TI_UNIT_X).Value               '単位
                    xTestInfo(index).arg2 = .Cells(TI_y, TI_ARG2_X).Value               'Arg2
                End If
'<<<2011/06/06 K.SUMIYASHIKI UPDATE
            
                If m_LoopParameter.Flg_CapCategory1 Then
                    xTestInfo(index).CapCategory1 = .Cells(TI_y, m_LoopParameter.CapCategory1_X).Value
                End If
            
                If m_LoopParameter.Flg_CapCategory2 Then
                    xTestInfo(index).CapCategory2 = .Cells(TI_y, m_LoopParameter.CapCategory2_X).Value
                End If
                                
                If m_LoopParameter.Flg_LoopJudgeLimit Then
                    xTestInfo(index).LoopJudgeLimit = val(.Cells(TI_y, m_LoopParameter.LoopJudgeLimit_X).Value)
                    If Len(.Cells(TI_y, m_LoopParameter.LoopJudgeLimit_X).Value) = 0 Then
                        xTestInfo(index).LoopJudgeLimit = 0.1
                    End If
                End If

                If m_WaitParameter.Flg_WaitAdjFlg Then
                    xTestInfo(index).WaitAdjFlg = .Cells(TI_y, m_WaitParameter.WaitAdjFlg_X).Value
                End If

                If m_OptParameter.Flg_OptIdentifier Then
                    xTestInfo(index).OptIdentifier = .Cells(TI_y, m_OptParameter.OptIdentifier_X).Value
                    
                    If xTestInfo(index).OptIdentifier <> "" Then intOptIdenCnt = intOptIdenCnt + 1
                    Dim intTestindex As Integer
                    'Reduplicate　Check
                    For intTestindex = 0 To index - 1
                        If xTestInfo(intTestindex).OptIdentifier = xTestInfo(index).OptIdentifier And xTestInfo(index).OptIdentifier <> "" Then
                            FLG_PALS_DISABLE.OptAdj = True
                            frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptIdentifier -" & "Duplicated at [" & TEST_INSTANCES & "]-" & xTestInfo(intTestindex).OptIdentifier & vbCrLf
                        End If
                    Next
                End If

                If m_OptParameter.Flg_OptTarget Then
                    xTestInfo(index).OptTarget = val(.Cells(TI_y, m_OptParameter.OptTarget_X).Value)
                    '>>>2011/4/22 M.IMAMURA ADD
                    If xTestInfo(index).OptIdentifier <> "" And xTestInfo(index).arg2 = 3 And (xTestInfo(index).OptTarget < xTestInfo(index).LowLimit Or xTestInfo(index).OptTarget > xTestInfo(index).HighLimit) Then
                        FLG_PALS_DISABLE.OptAdj = True
                        frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptTarget -" & "Out of TestLimt at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                    End If
                    If xTestInfo(index).OptIdentifier <> "" And xTestInfo(index).arg2 = 1 And (xTestInfo(index).OptTarget < xTestInfo(index).LowLimit) Then
                        FLG_PALS_DISABLE.OptAdj = True
                        frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptTarget -" & "Out of TestLimt at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                    End If
                    If xTestInfo(index).OptIdentifier <> "" And xTestInfo(index).arg2 = 2 And (xTestInfo(index).OptTarget > xTestInfo(index).HighLimit) Then
                        FLG_PALS_DISABLE.OptAdj = True
                        frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptTarget -" & "Out of TestLimt at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                    End If
                    '>>>2011/4/22 M.IMAMURA ADD
                End If

                If m_OptParameter.Flg_OptJudgeLimit Then
                    xTestInfo(index).OptJudgeLimit = val(.Cells(TI_y, m_OptParameter.OptJudgeLimit_X).Value)
                    If xTestInfo(index).OptJudgeLimit = 0 Then xTestInfo(index).OptJudgeLimit = 0.01
                End If
            

                If ((m_OptParameter.Flg_OptTarget = True) And (m_OptParameter.Flg_OptIdentifier = True) And (m_OptParameter.Flg_OptJudgeLimit = True)) Then
                    If .Cells(TI_y, m_OptParameter.OptIdentifier_X).Value <> "" Or val(.Cells(TI_y, m_OptParameter.OptJudgeLimit_X).Value) <> 0 And val(.Cells(TI_y, m_OptParameter.OptTarget_X).Value) <> 0 Then
                        '>>>2011/4/22 M.IMAMURA UPDATE
                        If .Cells(TI_y, m_OptParameter.OptIdentifier_X).Value = "" Then
                            FLG_PALS_DISABLE.OptAdj = True
                            frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptIdentifier -" & " Not Found Parameter at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                        End If
                        If .Cells(TI_y, m_OptParameter.OptJudgeLimit_X).Value <> "" And (val(.Cells(TI_y, m_OptParameter.OptJudgeLimit_X).Value) <= 0 Or val(.Cells(TI_y, m_OptParameter.OptJudgeLimit_X).Value) >= 1) Then
                            FLG_PALS_DISABLE.OptAdj = True
                            frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptJudgeLimit -" & " Illegal Parameter at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                        End If
                        If val(.Cells(TI_y, m_OptParameter.OptTarget_X).Value) <= 0 Then
                            FLG_PALS_DISABLE.OptAdj = True
                            frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptTarget -" & " Illegal Parameter at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                        End If
                        '<<<2011/4/22 M.IMAMURA UPDATE
                    End If

                End If

                If ((m_WaitParameter.Flg_CapCategory1 = True) And (m_WaitParameter.Flg_CapCategory2 = True) And (m_WaitParameter.Flg_WaitAdjFlg = True)) Then
                    If xTestInfo(index).CapCategory1 <> "" Or .Cells(TI_y, m_WaitParameter.WaitAdjFlg_X).Value <> "" Then
                        intWaitAdjCnt = intWaitAdjCnt + 1
                        If .Cells(TI_y, m_WaitParameter.CapCategory1_X).Value = "" Then
                            FLG_PALS_DISABLE.WaitAdj = True
                            frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- CapCategory1 -" & "Not Found Parameter at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                            intWaitAdjCnt = intWaitAdjCnt - 1
                        End If
                        If .Cells(TI_y, m_WaitParameter.WaitAdjFlg_X).Value <> "RUN" Then
'                            FLG_PALS_DISABLE.WaitAdj = True
'                            frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- WaitAdjFlg -" & "Not Found Parameter at [" & TEST_INSTANCES & "]-" & xTestInfo(Index).Parameter & vbCrLf
                            intWaitAdjCnt = intWaitAdjCnt - 1
                        End If

                    End If

                End If


            End If
            index = index + 1
        Loop
    End With

    If intOptIdenCnt = 0 Then
        FLG_PALS_DISABLE.OptAdj = True
        frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptIdentifier -" & "Not Found Parameter" & vbCrLf
    End If

    If intWaitAdjCnt <= 0 Then
        FLG_PALS_DISABLE.WaitAdj = True
        frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- WaitAdjFlg -" & "Not Found Parameter" & vbCrLf
    End If

Exit Sub

errPALSsub_ReadTestInfo:
    Call sub_errPALS("Read IG-XL sheet error at 'sub_ReadTestInfo'", "0-5-03-0-20")

End Sub


'********************************************************************************************
' 名前: TestCount
' 内容: 項目数を返す関数
' 引数: なし
' 戻値: 項目数
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/20　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Function TestCount() As Integer

On Error GoTo errPALSTestCount

    TestCount = UBound(xTestInfo)

Exit Function

errPALSTestCount:
    Call sub_errPALS("Get test count error at 'csPALS_Common -> TestCount'", "0-5-04-0-21")

End Function


