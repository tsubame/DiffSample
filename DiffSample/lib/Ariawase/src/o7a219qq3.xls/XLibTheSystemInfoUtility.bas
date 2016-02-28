Attribute VB_Name = "XLibTheSystemInfoUtility"
'概要:
'   TheSystemInfoのユーティリティ
'
'   Revision History:
'       Data        Description
'       2011/02/10  SystemInfoのUtility機能を実装した
'
'作成者:
'   0145184306
'

Option Explicit

Public TheSystemInfo As CSystemInfo ' SystemInfoを宣言する

Private Const ERR_NUMBER = 9999                           ' Error番号を保持する
Private Const CLASS_NAME = "XLibTheSystemInfoUtility"     ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty       ' Default値"Empty"を保持する

Public Sub CreateTheSystemInfoIfNothing()
'内容:
'   TheSystemInfoを初期化する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    On Error GoTo ErrHandler
    If TheSystemInfo Is Nothing Then
        Set TheSystemInfo = New CSystemInfo
    End If
    Exit Sub
ErrHandler:
    Set TheSystemInfo = Nothing
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

Public Sub InitializeTheSystemInfo()
End Sub

Public Sub DestroyTheSystemInfo()
'内容:
'   TheSystemInfoを破棄する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Set TheSystemInfo = Nothing
End Sub

Public Function RunAtJobEnd() As Long
End Function

