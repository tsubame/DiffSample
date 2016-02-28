Attribute VB_Name = "XLibTheParameterBankUtility"
'概要:
'   TheParameterBankのユーティリティ
'
'   Revision History:
'       Data        Description
'       2011/02/10  ParameterBankのUtility機能を実装した
'
'作成者:
'   0145184304
'

Option Explicit

Public TheParameterBank As IParameterBank ' ParameterBankを宣言する

Private Const ERR_NUMBER = 9999                           ' Error番号を保持する
Private Const CLASS_NAME = "XLibTheParameterBankUtility" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty       ' Default値"Empty"を保持する

Public Sub CreateTheParameterBankIfNothing()
'内容:
'   TheParameterBankを初期化する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    On Error GoTo ErrHandler
    If TheParameterBank Is Nothing Then
        Set TheParameterBank = New CParameterBank
    End If
    Exit Sub
ErrHandler:
    Set TheParameterBank = Nothing
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

Public Sub InitializeTheParameterBank()
'内容:
'   TheParameterBankを初期化する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
End Sub

Public Sub DestroyTheParameterBank()
'内容:
'   TheParameterBankを破棄する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Set TheParameterBank = Nothing
End Sub

Public Function RunAtJobEnd() As Long
    If Not TheParameterBank Is Nothing Then
        Call TheParameterBank.Clear
    End If
End Function
