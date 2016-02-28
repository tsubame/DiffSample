VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetFunctionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   セット関数情報クラス
'
'目的:
'   セット関数パラメータの型として使用
'
'作成者:
'   SLSI今手
'
'注意点:
'   各プロパティへの値の書き込みはWrite at onceです
'   パラメータの値にObjectは設定できません
'   TheErrorとして公開されているエラーマネージャObjectが必要です
'
Option Explicit

Private m_ConditionName As String              '測定条件名称
Private m_SetFunctionName As String            '条件設定機能名称
Private m_ConditionParameters As Collection    '条件設定機能のパラメータ値の保管用
Private m_IsEnableMacro As Boolean             'マクロの有効状態
Private m_MacroController As IMacroController  'マクロ実行機能提供Object
Private m_ConditionSheetName As String         '条件設定シートの名前
Private m_ExecuteSettingName As String         'マクログループ定義名
Private Const ERR_NUMBER As Long = 9999        'エラー時に渡すエラー番号

Public Sub RunMacro()
'内容:
'   Excelマクロの実行処理
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   MacroControllerにObjectが設定されていない場合はエラーとなります
'
    
    Call m_MacroController.RunMacro(Me)

End Sub

Public Property Get MacroController() As IMacroController
'内容:
'   Excelマクロ実行機能提供Object
'
'パラメータ:
'
'戻り値:
'   Excelマクロ実行機能提供Object
'
'注意事項:
'
    Set MacroController = m_MacroController
    
End Property

Public Property Set MacroController(ByRef pController As IMacroController)
'内容:
'   Excelマクロ実行機能提供Object
'
'パラメータ:
'   [pController]  In  IMacroController型: Excelマクロ実行機能提供Object
'
'戻り値:
'
'注意事項:
'
    
    Set m_MacroController = pController
    
End Property

Public Property Let ConditionSheetName(ByVal pWorksheetName As String)
'内容:
'   TestConditionワークシートの名前
'
'パラメータ:
'   [pWorksheetName]  In  string型:　ワークシート名
'
'戻り値:
'
'注意事項:
'   値の設定はWrite at onceです
'
    
    Static WriteFlag As Boolean
    If WriteFlag = False Then
        m_ConditionSheetName = pWorksheetName
        WriteFlag = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get ConditionSheetName() As String
'内容:
'   TestConditionワークシートの名前
'
'パラメータ:
'
'戻り値:
'   TestConditionワークシート名
'
'注意事項:
'
    ConditionSheetName = m_ConditionSheetName

End Property

Public Property Let IsEnableMacro(ByVal pSetValue As Boolean)
'内容:
'   マクロ有効無効設定値
'
'パラメータ:
'   [pSetValue]  In  Boolean型:   設定値(True:有効)
'
'戻り値:
'
'注意事項:
'   値の設定はWrite at onceです
'
    
    Static WriteFlag As Boolean
    If WriteFlag = False Then
        m_IsEnableMacro = pSetValue
        WriteFlag = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get IsEnableMacro() As Boolean
'内容:
'   マクロ有効無効設定値
'
'パラメータ:
'
'戻り値:
'   設定値 (True:有効)
'
'注意事項:
'
    
    IsEnableMacro = m_IsEnableMacro

End Property

Public Property Let ConditionName(ByVal pName As String)
'内容:
'   測定条件名称
'
'パラメータ:
'   [pName]  In  String型:   測定条件名文字列
'
'戻り値:
'
'注意事項:
'   値の設定はWrite at onceです
'
    Static WriteFlag As Boolean
    If WriteFlag = False Then
        m_ConditionName = pName
        WriteFlag = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get ConditionName() As String
'内容:
'   測定条件名称
'
'パラメータ:

'戻り値:
'   測定条件名文字列
'
'注意事項:
'
    ConditionName = m_ConditionName

End Property

Public Property Let FunctionName(ByVal pName As String)
'内容:
'   セット関数名称
'
'パラメータ:
'   [pName]  In  String型:   セット関数名文字列
'
'戻り値:
'
'注意事項:
'   値の設定はWrite at onceです
'
    Static WriteFlag As Boolean
    If WriteFlag = False Then
        m_SetFunctionName = pName
        WriteFlag = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get FunctionName() As String
'内容:
'   セット関数名称
'
'パラメータ:
'
'戻り値:
'   セット関数名文字列
'
'注意事項:
'
    FunctionName = m_SetFunctionName

End Property

Public Property Get ArgParameterCount() As Long
'内容:
'   セット関数の登録済みパラメータ数
'
'パラメータ:
'
'戻り値:
'   セット関数の登録済みパラメータ数
'
'注意事項:
'   Read onlyです
'
    ArgParameterCount = m_ConditionParameters.Count

End Property

Public Sub AddParameter(ByVal pParameter As Variant)
'内容:
'   セット関数のパラメータを登録する
'
'パラメータ:
'   [Parameter]  In  セット関数のパラメータ値
'
'戻り値:
'
'注意事項:
'   Objectは値として登録できません
'
    If IsObject(pParameter) Then
        'パラメータがObjectの時は登録させないでエラーとする
        Call RaiseError("The object cannot be registered as a parameter")
    Else
        Call m_ConditionParameters.Add(pParameter)
    End If

End Sub

Public Property Get Arg(ByVal pArgNumber As Long) As Variant
'内容:
'   登録されているセット関数のパラメータ値を取得
'
'パラメータ:
'   [pArgNumber]  In  取得するパラメータ番号値
'
'戻り値:
'   指定された番号のパラメータ値
'
'注意事項:
'   指定された番号のパラメータが存在しない場合、
'   取得値はEmptyとなります
'
    Arg = GetParameter(pArgNumber, m_ConditionParameters)

End Property
Public Property Let ExecuteSettingName(ByVal pExecuteSettingName As String)
'内容:
'   マクログループ定義名
'
'パラメータ:
'   [pExecuteSettingName]  In  String型:   セットマクログループ定義名文字列
'
'戻り値:
'
'注意事項:
'   値の設定はWrite at onceです
'

    Static WriteFlag As Boolean
    If WriteFlag = False Then
        m_ExecuteSettingName = pExecuteSettingName
        WriteFlag = True
    Else
        Call RaiseError("It has been already set.")
    End If

End Property

Public Property Get ExecuteSettingName() As String
'内容:
'   マクログループ定義名
'
'パラメータ:
'
'戻り値:
'   マクログループ定義名
'
'注意事項:
'
    ExecuteSettingName = m_ExecuteSettingName

End Property

Public Function Clone() As CSetFunctionInfo
'内容:
'   クローンObjectを提供する
'
'パラメータ:
'
'戻り値:
'   クローンObject
'
'注意事項:
'
    
    'クローン用Obj生成
    Dim cloneobj As CSetFunctionInfo
    Set cloneobj = New CSetFunctionInfo

    '情報の複製処理
    With cloneobj
        .ConditionName = m_ConditionName                  '条件名称の登録
        .FunctionName = m_SetFunctionName                 'セット関数名称の登録
        .IsEnableMacro = m_IsEnableMacro                  'マクロ有効設定値
        .ConditionSheetName = m_ConditionSheetName        '条件表ワークシート名
        .ExecuteSettingName = m_ExecuteSettingName        'マクログループ定義名
        Set .MacroController = m_MacroController          'マクロ実行機能提供Object
        Dim tmpRetParameter As Variant
        For Each tmpRetParameter In m_ConditionParameters 'セット関数のパラメータの登録
            .AddParameter (tmpRetParameter)
        Next tmpRetParameter
    End With

    '情報設定済みのObjectを返す
    Set Clone = cloneobj

End Function

'コンストラクタ
Private Sub Class_Initialize()
    Call ClearInfomation
End Sub

'測定条件情報初期化
Private Sub ClearInfomation()
    m_ConditionName = ""
    m_SetFunctionName = ""
    m_IsEnableMacro = False
    m_ConditionSheetName = ""
    m_ExecuteSettingName = ""
    Set m_MacroController = Nothing
    Set m_ConditionParameters = New Collection
End Sub

'パラメータ袋より指定番号のパラメータ値を取り出す関数（指定番号のパラメータが存在しない時はEmpty値を返す）
Private Function GetParameter(ByVal pArgNo As Long, ByVal pParemeterCollection As Collection) As Variant
    
    On Error GoTo DATA_GET_ERROR
    GetParameter = pParemeterCollection.Item(pArgNo + 1)
    On Error GoTo 0
    Exit Function

DATA_GET_ERROR:
    GetParameter = Empty
    Call RaiseError("Arg=" & pArgNo & " parameter is nothing")

End Function

'エラーレイズ用関数
Private Sub RaiseError(ByVal pErrMsg As String)
    If Not (TheError Is Nothing) Then
        'エラーマネージャへ
        Call TheError.Raise(ERR_NUMBER, TypeName(Me), pErrMsg)
    Else
        'VBのErrへ
        Call Err.Raise(ERR_NUMBER, TypeName(Me), pErrMsg)
    End If
End Sub


