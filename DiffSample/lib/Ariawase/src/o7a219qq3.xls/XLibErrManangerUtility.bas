Attribute VB_Name = "XLibErrManangerUtility"
'概要:
'   TheErrorのユーティリティ
'
'目的:
'   TheError:CErrManagerの初期化/破棄のユーテリティを定義する
'
'作成者:
'   a_oshima

Option Explicit

Public TheError As CErrManager

Public Sub CreateTheErrorIfNothing()
'内容:
'   As Newの代替として初回に呼ばせるインスタンス生成処理
'   履歴のクリアも行う
'
'パラメータ:
'   なし
'
'注意事項:
'
    On Error GoTo ErrHandler
    If TheError Is Nothing Then
        Set TheError = New CErrManager
    End If
    Call TheError.ClearHistory
    Exit Sub
ErrHandler:
    Set TheError = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub DestroyTheError()
    Set TheError = Nothing
End Sub

Public Function RunAtJobEnd() As Long

End Function
