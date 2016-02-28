VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   AcquireFrameWork（FWSetCondition）の情報を保持＆実行するクラス
'
'   Revision History:
'       Data        Description
'       2010/04/28  AcquireFrameWork（FWSetCondition）を実行する機能を実装した
'       2010/06/11  データ構造見直しのため、プログラムコードを変更した
'       2010/06/22  実行前後の情報を追加した（m_Statusを変更）
'                   実行情報リセット機能を実装した
'       2011/11/22  UserMacroの実行を動的/静的に切り替える機能を実装した
'
'作成者:
'   tomoyoshi.takase
'   0145184346
'

Option Explicit

Implements IAcquireAction

Private Const ERR_NUMBER = 9999         ' エラー番号を保持する
Private Const CLASS_NAME = "CSetAction" ' Class名称を保持する
Private Const DYNAMIC_WORD As String = "debug" ' 動的条件設定用のキーワード

Private m_FunctionInfo As CAcquireFunctionInfo ' ImageACQTBLシート1行のデータを保持する
'Private m_ActionLogger As CActionLogger        ' 実行LogデータをダンプするActionLoggerを保持する
Private m_InstanceName As String               ' TestInstance名称を保持する
Private m_ActionName As String                 ' AcquireFrameWorkが実行するUserMacro名称を保持する
Private m_Status As Long                       ' 実行結果を保持する（-1：実行前　0：正常終了　1：エラー発生）
Private m_ISetCondition As ISetCondition       ' 条件設定クラスのインターフェース

Public Property Get AsIAcquireAction() As IAcquireAction
    Set AsIAcquireAction = Me
End Property

Private Function IAcquireAction_Execute() As Integer
'内容:
'   SetConditionのUserMacroを実行する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'
    Call m_ISetCondition.Execute(m_FunctionInfo)
    m_Status = TL_SUCCESS

    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    m_Status = TL_ERROR
    Call TheError.Raise(ERR_NUMBER, "TheCondition.SetCondition", "Can not Execute Set Condition Macro: " & m_InstanceName)


End Function

Private Function IAcquireAction_ToptStart() As Integer
'内容:
'   TestOptimizer機能を使って、AcquireFrameWorkを実行する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    '#####  AcquireFrameWorkを実行する  #####
    With m_FunctionInfo
        If .macroName <> "" Then
            On Error GoTo IGXLErrHandler
            Call TheHdw.TOPT.Start(.GroupName, .macroName, .InstanceName)
        Else
            On Error GoTo ErrHandler
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".ToptStart", "Can Not Found Acquire Frame Work Name ! ( " _
                                & .GroupName & " / " & .InstanceName & " )")
        End If
    End With


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
IGXLErrHandler:
    IAcquireAction_ToptStart = TL_ERROR
    m_Status = IAcquireAction_ToptStart
    Exit Function
ErrHandler:
    IAcquireAction_ToptStart = TL_ERROR
    m_Status = IAcquireAction_ToptStart
    Exit Function


End Function

Private Function IAcquireAction_ClearStatus() As Integer
'内容:
'   AcquireFrameWorkインスタンスの実行履歴を削除する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWorkインスタンスの実行履歴を削除する  #####
    Call TheHdw.TOPT.ClearStatus(m_FunctionInfo.GroupName)


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    IAcquireAction_ClearStatus = TL_ERROR
    Exit Function


End Function

Private Property Get IAcquireAction_GetFrameWorkName() As String
    IAcquireAction_GetFrameWorkName = m_FunctionInfo.macroName
End Property

Private Property Get IAcquireAction_GetStatus() As Integer
    IAcquireAction_GetStatus = m_Status
End Property

Private Sub IAcquireAction_Reset()
    m_Status = -1
End Sub

Private Property Get IAcquireAction_AcquireFunctionInfo() As CAcquireFunctionInfo
    Set IAcquireAction_AcquireFunctionInfo = m_FunctionInfo
End Property

Private Sub IAcquireAction_Initialize(reader As IParameterReader)
    Set m_FunctionInfo = New CAcquireFunctionInfo
    Call m_FunctionInfo.Initialize(reader)
    m_InstanceName = reader.ReadAsString("Arg0@Parameters")
    
    'ここで動的or静的の条件設定インスタンスを作成
    Dim DebugWord As String
    DebugWord = LCase(reader.ReadAsString("Arg4@Parameters"))
    If StrComp(DebugWord, DYNAMIC_WORD, 0) = 0 Then

        '動的条件設定クラスを作成
        Dim DynamicSetCondition As CDynamicSetCondition
        Set DynamicSetCondition = New CDynamicSetCondition
        Set m_ISetCondition = DynamicSetCondition
    
    ElseIf StrComp(DebugWord, "", 0) = 0 Then
    
        '静的条件設定クラスを作成
        Dim StaticSetCondition As CStaticSetCondition
        Set StaticSetCondition = New CStaticSetCondition
        Set m_ISetCondition = StaticSetCondition
        
    Else
        'Arg4が"debug"、または""(空白)でない場合
        
        m_Status = TL_ERROR
        Call TheError.Raise(ERR_NUMBER, "CSetAction", "Please input ""debug"" or Null Character into Arg4." & vbCr & vbCr & "Input Char = " & DebugWord)

    End If
    
    m_Status = -1
End Sub

Private Sub Class_Initialize()
    Set m_FunctionInfo = Nothing
    Set m_ISetCondition = Nothing
    m_InstanceName = ""
    m_ActionName = ""
    m_Status = -1
End Sub

Private Sub Class_Terminate()
    Set m_FunctionInfo = Nothing
    Set m_ISetCondition = Nothing
    m_InstanceName = ""
    m_ActionName = ""
    m_Status = -1
End Sub


