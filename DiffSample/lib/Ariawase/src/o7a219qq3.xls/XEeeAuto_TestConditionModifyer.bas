Attribute VB_Name = "XEeeAuto_TestConditionModifyer"
'概要:
'   TestConditionから呼ばれるマクロ集
'
'目的:
'   TestConditionシートを自動で省略するための機能
'
'作成者:
'   2012/03/24 Ver0.1 D.Maruyama    Draft
'   2012/04/19 Ver0.2 D.Maruyama    コードをきれいにした
'   2013/02/25 Ver0.3 H.Arikawa     光源回避用の定義・処理(実装はクラス)を追加

Option Explicit

'DCHANの条件チェックは時間がかかるので、DCHANに設定しない場合は見に行かないための定義
Public Const EEE_AUTO_TESTER_CHECKER_WITHOUT_DCHAN As String = "Without Dchan"
Public Const EEE_AUTO_TESTER_CHECKER_WITH_DCHAN As String = "Dchan"

'条件設定の種類
Public Enum eTestCnditionCheck
    TCC_TESTER_CHANNELS
    TCC_SETVOLTAGE
    TCC_ILLUMINATOR
    TCC_ILLUMINATOR_ESCAPSE
    TCC_ILLUMINATOR_MODZ1
    TCC_ILLUMINATOR_MODZ2
    TCC_APMU_UB
End Enum

'TestConditionシート周りの定義
Private Const CURRENT_SETTING As String = "C2"
Private Const ARG_MAX As Long = 10
Private Const TESTCONDITION_ITEM_START As String = "B5"
Private Const DEFAULT_ENABLES_START As String = "N5"
Private Const DEFAULT_LABEL As String = "Default"

'CEeeAuto_TestConditionItemを格納するコレクション
Private m_colTestItems As Collection

'テストに応じてチェックを行うクラス群
Private m_IllminatorChecker As IEeeAuto_TestConditionChecker
Private m_TesterChannelChecker As IEeeAuto_TestConditionChecker
Private m_ApmuUbChecker As IEeeAuto_TestConditionChecker
Private m_IllmEscapeChecker As IEeeAuto_TestConditionChecker
Private m_IllmOptModZSet1Checker As IEeeAuto_TestConditionChecker
Private m_IllmOptModZSet2Checker As IEeeAuto_TestConditionChecker


'最初の設定は必ず行わせるためのフラグ
Private m_IsFirstSetIlluminator As Boolean
Private m_IsFirstSetUB As Boolean
Private m_IsFirstSetVoltatge As Boolean
Private m_IsFirstSetIllmEscape As Boolean
Private m_IsFirstSetIllmOptModZSet1 As Boolean
Private m_IsFirstSetIllmOptModZSet2 As Boolean


'SetForceEnableTestConditionで設定されたコンディション名は覚えておき
'チェックは行わない
Private m_colForceEnableCondition As Collection

'このモジュールの状態定義
Private Enum TCCM_STATUS
    UNKNWON = 0
    INITIALIZED = 1
    CHECKED_BEFORE = 2
End Enum

'状態保持変数
Private m_State As TCCM_STATUS

'自動省略機能の初期化
Public Sub InitializeAutoConditionModify()
    
    Dim lColumn As Long
    Dim curColumn As Long
    Dim mySht As Worksheet
    
    m_State = UNKNWON
    
    On Error GoTo ErrorHandler:
    
    'AutoModifyModeであることの通知
    TheExec.Datalog.WriteComment "-----TestCondition Sheet Auto Modify Mode!! ---"
    
    'メンバの初期化
    Set m_colTestItems = Nothing
    Set m_colTestItems = New Collection
    Set m_colForceEnableCondition = Nothing
    Set m_colForceEnableCondition = New Collection
    
    'シートの取得
    Set mySht = ThisWorkbook.Worksheets(TheCondition.TestConditionSheet)
    
    'Defaultでない場合は動かない
    If (Not IsExecuteSetDefault(mySht)) Then
        Err.Raise 9999, "InitializeAutoConditionModify", "Execute list is not Default!"
        Exit Sub
    End If
        
    'TestConditionシートからの読み込み
    Call ReadTestCondition(mySht)
    
    'オブジェクトの構築
    Set m_IllminatorChecker = New CEeeAuto_IlluminatorChecker
    Set m_ApmuUbChecker = New CEeeAuto_ApmuUBChecker
    Set m_TesterChannelChecker = New CEeeAuto_TesterChannelChecker
    Set m_IllmEscapeChecker = New CEeeAuto_IllumEscapeChecker
    Set m_IllmOptModZSet1Checker = New CEeeAuto_IllumModeZSet1Checker
    Set m_IllmOptModZSet2Checker = New CEeeAuto_IllumModeZSet2Checker

    '最初のコマンドであるか示す変数を初期化
    m_IsFirstSetIlluminator = True
    m_IsFirstSetUB = True
    m_IsFirstSetVoltatge = True
    m_IsFirstSetIllmEscape = True
    m_IsFirstSetIllmOptModZSet1 = True
    m_IsFirstSetIllmOptModZSet2 = True
    
    '開放
    Set mySht = Nothing
    
    '状態遷移
    m_State = INITIALIZED

    Exit Sub
    
ErrorHandler:
    Set mySht = Nothing
    m_State = UNKNWON
    
End Sub

'自動省略機能の終了処理
Public Sub UninitializeAutoConditionModify()
    
    If m_State = INITIALIZED Then
        '結果を書き込む
        Call ModifyTestCondtitionSheet
    End If
    
    'オブジェクトを開放
    Set m_IllmOptModZSet2Checker = Nothing
    Set m_IllmOptModZSet1Checker = Nothing
    Set m_IllmEscapeChecker = Nothing
    Set m_TesterChannelChecker = Nothing
    Set m_ApmuUbChecker = Nothing
    Set m_IllminatorChecker = Nothing
    
    Set m_colForceEnableCondition = Nothing
    Set m_colTestItems = Nothing
    
    '念のため
    m_IsFirstSetIlluminator = True
    m_IsFirstSetUB = True
    m_IsFirstSetVoltatge = True
    m_IsFirstSetIllmEscape = True
    m_IsFirstSetIllmOptModZSet1 = True
    m_IsFirstSetIllmOptModZSet2 = True
    
    '状態遷移
    m_State = UNKNWON
    
End Sub

'条件設定前のコンディション取得
Public Sub CheckBeforeTestCondition(ByVal eMode As eTestCnditionCheck, ByRef pInfo As CSetFunctionInfo)
    
    '初期化済みでない場合はすぐに抜ける
    If m_State <> INITIALIZED Then
        Exit Sub
    End If
    
    TheHdw.StartStopwatch
    'ForceEnable化チェック
    If IsForceEnableTestCondition(pInfo.ConditionName) Then
        Exit Sub
    End If
    Dim sTime As Single
        
    'モードによって呼ぶアイテムを決める
    Dim strTemp As String
    Select Case eMode
        Case TCC_ILLUMINATOR
            If m_IsFirstSetIlluminator Then
                m_IsFirstSetIlluminator = False
                Exit Sub
            End If
            m_IllminatorChecker.CheckBeforeCondition
        Case TCC_APMU_UB
            If m_IsFirstSetUB Then
                m_IsFirstSetUB = False
                Exit Sub
            End If
            m_ApmuUbChecker.CheckBeforeCondition
        Case TCC_SETVOLTAGE
            If m_IsFirstSetVoltatge Then
                m_IsFirstSetVoltatge = False
                Exit Sub
            End If
            m_TesterChannelChecker.SetOperationMode (EEE_AUTO_TESTER_CHECKER_WITH_DCHAN)
            m_TesterChannelChecker.CheckBeforeCondition
        Case TCC_TESTER_CHANNELS
            If IsSetDigitalChannnel(pInfo) Then
               m_TesterChannelChecker.SetOperationMode (EEE_AUTO_TESTER_CHECKER_WITH_DCHAN)
            Else
                m_TesterChannelChecker.SetOperationMode (EEE_AUTO_TESTER_CHECKER_WITHOUT_DCHAN)
            End If
            m_TesterChannelChecker.CheckBeforeCondition
            
        Case TCC_ILLUMINATOR_ESCAPSE
            If m_IsFirstSetIllmEscape Then
                m_IsFirstSetIllmEscape = False
                Exit Sub
            End If
            m_IllmEscapeChecker.SetEndPosition GetFirstOptSetSameCategory(pInfo)
            m_IllmEscapeChecker.CheckBeforeCondition
            
        Case TCC_ILLUMINATOR_MODZ1
            If m_IsFirstSetIllmOptModZSet1 Then
                m_IsFirstSetIllmOptModZSet1 = False
                Exit Sub
            End If
            m_IllmOptModZSet1Checker.SetEndPosition pInfo.Arg(1)
            m_IllmOptModZSet1Checker.CheckBeforeCondition
        
        Case TCC_ILLUMINATOR_MODZ2
            If m_IsFirstSetIllmOptModZSet2 Then
                m_IsFirstSetIllmOptModZSet2 = False
                Exit Sub
            End If
            m_IllmOptModZSet2Checker.SetEndPosition pInfo.Arg(1)
            m_IllmOptModZSet2Checker.CheckBeforeCondition
            
    End Select
    
    sTime = TheHdw.ReadStopwatch
    Select Case eMode
        Case TCC_ILLUMINATOR
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumCheck Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_APMU_UB
            TheExec.Datalog.WriteComment pInfo.ConditionName & " APMU_UB Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_SETVOLTAGE
            TheExec.Datalog.WriteComment pInfo.ConditionName & " SetVoltage Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_TESTER_CHANNELS
            TheExec.Datalog.WriteComment pInfo.ConditionName & " Tester_Channnel Before " & pInfo.FunctionName & " " & CStr(sTime * 1000) & " " & pInfo.Arg(0)
        Case TCC_ILLUMINATOR_ESCAPSE
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumEscapse Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_ILLUMINATOR_MODZ1
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumModZSet1 Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_ILLUMINATOR_MODZ2
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumModZSet2 Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
    End Select
    
    '状態遷移
    m_State = CHECKED_BEFORE
    
    Exit Sub
    
ErrorHandler:
    m_State = UNKNWON

End Sub

'条件設定後のコンディション取得、意味のある条件設定か確認
Public Sub CheckAfterTestCondition(ByVal eMode As eTestCnditionCheck, ByRef pInfo As CSetFunctionInfo)
    
    '条件設定前の状態を取得していない場合はすぐに抜ける
    If m_State <> CHECKED_BEFORE Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler:
    
    TheHdw.StartStopwatch
    
     'Beforeが走らなければ実行できないので、ここでは行わない
'    'ForceEnable化チェック
'    If IsForceEnableTestCondition(pInfo.ConditionName) Then
'        Exit Sub
'    End If
    
    Dim sTime As Single
    Dim strTemp As String
    Dim IsValid As Boolean
    IsValid = True
    
    'モードによって呼ぶアイテムを決める
    Select Case eMode
        Case TCC_ILLUMINATOR
            IsValid = m_IllminatorChecker.CheckAfterCondition
        Case TCC_APMU_UB
            IsValid = m_ApmuUbChecker.CheckAfterCondition
        Case TCC_SETVOLTAGE
            IsValid = m_TesterChannelChecker.CheckAfterCondition
        Case TCC_TESTER_CHANNELS
            IsValid = m_TesterChannelChecker.CheckAfterCondition
    End Select
    
    '識別子の作成
    Dim strIdenfier As String
    strIdenfier = GetTestConditionIdenfier(pInfo)
    
    'Itemを探す
    Dim IsFound As Boolean
    Dim obj As CEeeAuto_TestConditionItem
    IsFound = False
    For Each obj In m_colTestItems
        If (obj.GetTestConditionIdenfier = strIdenfier) Then
            IsFound = True
            Exit For
        End If
    Next
    
    If IsFound Then
        Call obj.SetValideCodition(IsValid)
    End If
  
    sTime = TheHdw.ReadStopwatch
    Select Case eMode
        Case TCC_ILLUMINATOR
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumCheck After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_APMU_UB
            TheExec.Datalog.WriteComment pInfo.ConditionName & " APMU_UB After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_SETVOLTAGE
            TheExec.Datalog.WriteComment pInfo.ConditionName & " SetVoltage After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_TESTER_CHANNELS
            TheExec.Datalog.WriteComment pInfo.ConditionName & " Tester_Channnel After " & pInfo.FunctionName & " " & CStr(sTime * 1000) & " " & pInfo.Arg(0)
    End Select

    '状態遷移
    m_State = INITIALIZED
    
    Exit Sub
    
ErrorHandler:
    m_State = UNKNWON
  
End Sub

'内容:
'   条件設定マクロ側にこの条件設定が省略するか返す
'   対象関数は
'       OptEscape
'       OptModOrModZ1
'       OptModOrModZ2
'
'備考:
'　　OptEscape, OptModOrModZ1, OptModOrModZ2は呼ぶ瞬間に省略すべきかどうか決まる
'　　省略可能なのにメソッド実行すると、次の条件省略オブジェクトが正しく動作しない。
'　　たとえば退避不要なのにOptEscapeを実行すると、最終的な行き先が同じ場合でも
'　　OptModOrModZ2を省略不可と判断してしまう。具体例で示すと
' 　　現在地　DOWN, 最終目的地　DOWNだとすると
'　　  OptEscapeで退避位置(UP)へ移動
'　　  OptModOrModZ1で退避軸じゃない方向の移動
'　　  OptModOrModZ2で最終位置(DOWN)に移動
'　　OptModOrModZ2はUP→DOWNへ位置変更があるので、省略不可と判断する。
'　　なのでOptEscapeは省略可能と判断した場合、実行をSkipする必要がある。
Public Function IsValidTestCondition(ByVal eMode As eTestCnditionCheck, ByRef pInfo As CSetFunctionInfo) As Boolean

    '条件設定前の状態を取得していない場合はすぐに抜ける
    If m_State <> CHECKED_BEFORE Then
        IsValidTestCondition = True
        Exit Function
    End If
    
    On Error GoTo ErrorHandler:
    
    Dim sTime As Single
    Dim IsValid As Boolean
    IsValid = True
    
    'モードによって呼ぶアイテムを決める
    Select Case eMode
        Case TCC_ILLUMINATOR_ESCAPSE
            IsValid = m_IllmEscapeChecker.CheckAfterCondition
        Case TCC_ILLUMINATOR_MODZ1
            IsValid = m_IllmOptModZSet1Checker.CheckAfterCondition
        Case TCC_ILLUMINATOR_MODZ2
            IsValid = m_IllmOptModZSet2Checker.CheckAfterCondition
    End Select
   
    '識別子の作成
    Dim strIdenfier As String
    strIdenfier = GetTestConditionIdenfier(pInfo)
    
    'Itemを探す
    Dim IsFound As Boolean
    Dim obj As CEeeAuto_TestConditionItem
    IsFound = False
    For Each obj In m_colTestItems
        If (obj.GetTestConditionIdenfier = strIdenfier) Then
            IsFound = True
            Exit For
        End If
    Next
    
    If IsFound Then
        Call obj.SetValideCodition(IsValid)
    End If
  
    sTime = TheHdw.ReadStopwatch
    Select Case eMode
        Case TCC_ILLUMINATOR_ESCAPSE
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumEscapse After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_ILLUMINATOR_MODZ1
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumModZSet1 After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_ILLUMINATOR_MODZ2
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumModZSet2 After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
    End Select

    '状態遷移
    m_State = INITIALIZED
    
    IsValidTestCondition = IsValid

    Exit Function
    
ErrorHandler:
    m_State = UNKNWON

End Function

Private Sub ModifyTestCondtitionSheet()
    
On Error GoTo ErrorHandler:
    
    'シートの取得
    Dim mySht As Worksheet
    Set mySht = ThisWorkbook.Worksheets(TheCondition.TestConditionSheet)

    'デフォルトカラムの位置取得
    Dim lDefaultColumn As Long
    lDefaultColumn = mySht.Range(DEFAULT_ENABLES_START).Column
    
    '値をいったん取得する
    Dim myrange As Range
    Dim lRefRow As Long, lRefColumn As Long
    lRefRow = mySht.Range(TESTCONDITION_ITEM_START).Row
    lRefColumn = mySht.Range(TESTCONDITION_ITEM_START).Column
    Set myrange = mySht.Range(mySht.Cells(lRefRow, lDefaultColumn), mySht.Cells(lRefRow + m_colTestItems.Count - 1, lDefaultColumn))
 
    Dim aryTemp() As Variant
    aryTemp = myrange.Value
    
    'falseのところのみ上書き
    Dim obj As CEeeAuto_TestConditionItem
    Dim i As Long
    i = 1
    For Each obj In m_colTestItems
        If obj.IsEnable Then
            aryTemp(i, 1) = "enable"
        Else
            aryTemp(i, 1) = "disable"
        End If
        i = i + 1
    Next
   
    '値を戻す
    myrange.Value = aryTemp
    Erase aryTemp
    
    
    Set mySht = Nothing
    
    Exit Sub
ErrorHandler:
    Set mySht = Nothing


End Sub


Public Sub SetForceEnableTestCondition(ByVal strCondition As String)

    '条件設定を強制TRUE
    Dim obj As CEeeAuto_TestConditionItem
    For Each obj In m_colTestItems
        If (obj.ConditionName = strCondition) Then
            Call obj.SetValideCodition(True)
        End If
    Next

    'キーが重なるとエラーになるが、重なっているということは
    'すでにForceEnable化されているので改めて追加する必要がない
On Error Resume Next
    Call m_colForceEnableCondition.Add(strCondition, strCondition)
On Error GoTo 0

End Sub

Private Function IsExecuteSetDefault(ByRef mySht As Worksheet) As Boolean

    If mySht Is Nothing Then
        IsExecuteSetDefault = False
        Exit Function
    End If
        
    '現在の設定を取得
    Dim strCurSetting As String
    strCurSetting = mySht.Range(CURRENT_SETTING)
     
    'Default設定以外はエラー
    If strCurSetting <> DEFAULT_LABEL Then
        IsExecuteSetDefault = False
    End If
    
    IsExecuteSetDefault = True
 
End Function

Private Sub ReadTestCondition(ByRef mySht As Worksheet)

    Const ARG_START As Long = 3

    'エラーチェック
    If mySht Is Nothing Then
        Exit Sub
    End If
    
    'TestConditionの全情報を配列に格納
    Dim aryTestConditions As Variant
    aryTestConditions = mySht.Range(mySht.Range(TESTCONDITION_ITEM_START), _
                        mySht.Cells.SpecialCells(xlCellTypeLastCell))
    
    'デフォルトカラムの配列での位置取得
    Dim lDefaultColumn As Long
    lDefaultColumn = mySht.Range(DEFAULT_ENABLES_START).Column - mySht.Range(TESTCONDITION_ITEM_START).Column + 1
    
    '配列の絶対座標に対するオフセット行を取得
    Dim lOffsetRow As Long
    lOffsetRow = mySht.Range(TESTCONDITION_ITEM_START).Row
    
    
    'TestConditionの取得 TheConditionから取得しないのは行番号がわからないため
    Dim i As Long, j As Long
    Dim tempItem As CEeeAuto_TestConditionItem
    Dim strConditionName As String
    Dim strFuncName As String
    Dim lArgCount As Long
    Dim aryArg(9) As Variant
    Dim lRow As Long
    Dim IsEnable As Boolean
    For i = 1 To UBound(aryTestConditions, 1)
    
        '空白セルがきた場合は抜ける
        If (IsEmpty(aryTestConditions(i, 1))) Then
            Exit For
        End If
        
        'パラメータの読み込み
        strConditionName = aryTestConditions(i, 1)
        strFuncName = aryTestConditions(i, 2)
        If (aryTestConditions(i, lDefaultColumn) = "enable") Then
            IsEnable = True
        Else
            IsEnable = False
        End If
        j = ARG_START
        While ((Not IsEmpty(aryTestConditions(i, j))) And (aryTestConditions(i, j) <> "#EOP") Or j >= ARG_START + ARG_MAX)
            aryArg(j - ARG_START) = aryTestConditions(i, j)
            j = j + 1
        Wend
        lArgCount = j - ARG_START
        lRow = lOffsetRow + i - 1
        
        'オブジェクトを生成、コレクションに追加
        Set tempItem = New CEeeAuto_TestConditionItem
        Call tempItem.SetParams(strConditionName, strFuncName, lArgCount, aryArg, lRow, IsEnable)
        m_colTestItems.Add tempItem
        Set tempItem = Nothing
        
    Next i

End Sub

'識別子の作成　このモジュール専用
Private Function GetTestConditionIdenfier(ByRef pInfo As CSetFunctionInfo)

    Dim aryArg(ARG_MAX - 1) As Variant
    
    Dim i As Long
    With pInfo
        For i = 0 To .ArgParameterCount - 1
            aryArg(i) = .Arg(i)
        Next
        GetTestConditionIdenfier = GetTestConditionIdenfier_impl(.ConditionName, .FunctionName, .ArgParameterCount, aryArg)
    End With
    
End Function
 
'識別子の作成を共通でやらせたいため、別関数で外に出す
Public Function GetTestConditionIdenfier_impl(ByVal strCndName As String, ByVal strFunName As String, ByVal lCount As Long, ByRef aryArg() As Variant) As String

    Dim strIdenfier As String
    Dim i As Long
    
    strIdenfier = strCndName & "_" & strFunName
    
    For i = 0 To lCount - 1
        strIdenfier = strIdenfier & "_" & CStr(aryArg(i))
    Next i

    GetTestConditionIdenfier_impl = strIdenfier

End Function

'自動省略機能対象外のコンディションか確認する
Private Function IsForceEnableTestCondition(ByVal strConditionName As String) As Boolean

    If m_colForceEnableCondition.Count = 0 Then
        IsForceEnableTestCondition = False
    End If
        
    Dim obj As Variant
    
    For Each obj In m_colForceEnableCondition
        If obj = strConditionName Then
            IsForceEnableTestCondition = True
            Exit Function
        End If
    Next obj
    
    IsForceEnableTestCondition = False
    
    Exit Function
    
End Function

'デジタルピンは取り込みが遅いため、DCHANを含まない条件設定の場合はみにいきたくない
'このため関数名によって問い合わせを行い、DCHANを使うか確認し、使用しないなら見に行かない
Private Function IsSetDigitalChannnel(ByRef pInfo As CSetFunctionInfo) As Boolean
       
    Dim chanType As chtype

    Select Case pInfo.FunctionName
        Case "FW_DisconnectPins"
            chanType = GetChanType(pInfo.Arg(0))
            If (chanType = chAPMU) Or (chanType = chDPS) Then
                IsSetDigitalChannnel = False
                Exit Function
            End If
        Case "FW_SetFVMI"
            chanType = GetChanType(pInfo.Arg(0))
            If (chanType = chAPMU) Or (chanType = chDPS) Then
                IsSetDigitalChannnel = False
                Exit Function
            End If
    End Select
    
    IsSetDigitalChannnel = True

End Function

'チャンネルタイプの決定
Private Function GetChanType(ByVal PinList As String) As chtype

    Dim chanType As chtype
    Dim pinArr() As String
    Dim pinNum As Long
    Dim i As Long

    With TheExec.DataManager
        chanType = .chanType(PinList)

        If chanType = chUnk Then
            On Error GoTo INVALID_PINLIST
            Call .DecomposePinList(PinList, pinArr, pinNum)
            On Error GoTo -1

            chanType = .chanType(pinArr(0))
            For i = 0 To pinNum - 1
                If chanType <> .chanType(pinArr(i)) Then
                    chanType = chUnk
                    Exit For
                End If
            Next i
        End If
    End With

    GetChanType = chanType
    Exit Function

INVALID_PINLIST:
    GetChanType = chUnk

End Function

'本当はこんなことしたくないけど、ESCAPEで最終的にどこに移動するか引数で渡らないので
'自分で調べる
Private Function GetFirstOptSetSameCategory(ByRef pInfo As CSetFunctionInfo) As String
    
    GetFirstOptSetSameCategory = ""
    '行番号不要のため、おとなしくTheConditionの情報を使う
    Dim myCol As Collection
    Set myCol = TheCondition.GetCloneConditionInfo(pInfo.ConditionName)
        
    '指定したコンディショングループの中で最初に見つかった
    'FW_OptModOrModZ1の引数を利用する。複数回呼ばれてたらもれなく誤動作する
    Dim obj As CSetFunctionInfo
    
    For Each obj In myCol
        If obj.FunctionName = "FW_OptModOrModZ1" Or _
           obj.FunctionName = "FW_OptModOrModZ2" Or _
           obj.FunctionName = "FW_OptEscape" Or _
           obj.FunctionName = "FW_OptSet" Or _
           obj.FunctionName = "FW_OptSet_Test" Then
            GetFirstOptSetSameCategory = obj.Arg(1)
            Exit For
        End If
    Next obj

    Set myCol = Nothing
End Function

