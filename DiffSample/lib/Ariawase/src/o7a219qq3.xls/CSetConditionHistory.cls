VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetConditionHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   条件設定履歴情報用クラス
'
'目的:
'   設定履歴情報の型として使用
'
'作成者:
'   SLSI今手
'
'注意点:
'   CSetFunctionInfo が必要です。
'   各プロパティへの値の書き込みはWrite at onceです
'   TheErrorとして公開されているエラーマネージャObjectが必要です
'
Option Explicit

Private m_RecordDate As String                        '履歴の記録日
Private m_RecordTime As String                        '履歴の記録時間
Private m_ExecutionTime As Double                     'セット関数の実行時間保存用[Sec]
Private m_ExecTestInstancesName As String             'セット関数が実行されたテストインスタンス名称
Private m_SetFunctionParameterObj As CSetFunctionInfo 'セット関数のパラメータOBJ用
Private Const ERR_NUMBER As Long = 9999               'エラー時に渡すエラー番号

Public Property Let RecordTime(ByVal pTime As String)
'内容:
'   履歴記録時の時間
'
'パラメータ:
'   [pTime]  In  String型:   記録時の時間
'
'戻り値:
'
'注意事項:
'   値の設定はWrite at onceです
'
    Static WriteFlg As Boolean
    If WriteFlg = False Then
        m_RecordTime = pTime
        WriteFlg = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get RecordTime() As String
'内容:
'   履歴記録時の時間
'
'パラメータ:
'
'戻り値:
'   履歴記録時の時間
'
'注意事項:
'
    RecordTime = m_RecordTime

End Property

Public Property Let RecordDate(ByVal pRecDate As String)
'内容:
'   履歴記録時の日付
'
'パラメータ:
'   [RecDate]  In  String型:   記録時の日付
'
'戻り値:
'
'注意事項:
'   値の設定はWrite at onceです
'
    Static WriteFlg As Boolean
    If WriteFlg = False Then
        m_RecordDate = pRecDate
        WriteFlg = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get RecordDate() As String
'内容:
'   履歴記録時の日付
'
'パラメータ:
'
'戻り値:
'   履歴記録時の日付
'
'注意事項:
'
    RecordDate = m_RecordDate

End Property

Public Property Let ExecTime(ByVal pTime As Double)
'内容:
'   セット関数の実行時間
'
'パラメータ:
'   [pTime]  In  Double型:   実行時間
'
'戻り値:
'
'注意事項:
'   値の設定はWrite at onceです
'
    Static WriteFlg As Boolean
    If WriteFlg = False Then
        m_ExecutionTime = pTime
        WriteFlg = True
    Else
        Call RaiseError("It has already set it")
    End If
End Property

Public Property Get ExecTime() As Double
'内容:
'   セット関数の実行時間
'
'パラメータ:
'
'戻り値:
'   セット関数の実行時間
'
'注意事項:
'
    ExecTime = m_ExecutionTime

End Property

Public Property Let ExecTestInstanceName(ByVal pName As String)
'内容:
'   セット関数実行時のテストインスタンス名
'
'パラメータ:
'   [pName]  In  String型:   テストインスタンス名称
'
'戻り値:
'
'注意事項:
'   値の設定はWrite at onceです
'
    Static WriteFlg As Boolean
    If WriteFlg = False Then
        m_ExecTestInstancesName = pName
        WriteFlg = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get ExecTestInstanceName() As String
'内容:
'   セット関数実行時のテストインスタンス名
'
'パラメータ:
'
'戻り値:
'   セット関数実行時のテストインスタンス名
'
'注意事項:
'
    ExecTestInstanceName = m_ExecTestInstancesName

End Property

Public Property Set SetFunctionInfo(ByVal pParameterObj As CSetFunctionInfo)
'内容:
'   実行時のセット関数情報
'
'パラメータ:
'   [ParameterObj]  In  CSetFunctionInfo型:  セット関数情報Object
'
'戻り値:
'
'注意事項:
'   値の設定はWrite at onceです
'
    If m_SetFunctionParameterObj Is Nothing Then
        Set m_SetFunctionParameterObj = pParameterObj
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get SetFunctionInfo() As CSetFunctionInfo
'内容:
'   実行時のセット関数情報
'
'パラメータ:
'
'戻り値:
'   実行時のセット関数情報Object
'
'注意事項:
'
    Set SetFunctionInfo = m_SetFunctionParameterObj

End Property

'コンストラクタ
Private Sub Class_Initialize()
    Call ClearInfomation
End Sub

'測定条件情報初期化
Private Sub ClearInfomation()
    m_RecordDate = ""
    m_RecordTime = ""
    m_ExecutionTime = 0#
    m_ExecTestInstancesName = ""
    Set m_SetFunctionParameterObj = Nothing
End Sub

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

