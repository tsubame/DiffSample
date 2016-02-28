VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   スナップショット機能提供I/F
'
'目的:
'   複数のスナップショット機能の操作統一
'
'作成者:
'   SLSI今手
'
'
'Code Checked
'Comment Checked
'

Option Explicit

'#Pass-Ok
Public Property Get FunctionName() As String
'内容:
'   スナップショット機能名
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   何スナップショットなのか確認用

End Property

'#Pass-Ok
Public Sub SaveSnapshot(Optional ByVal CommentMsg As String = "")
'内容:
'   スナップショット取得実行
'
'パラメータ:
'   [CommentMsg]　In  スナップショットログに出力するコメントメッセージ
'
'戻り値:
'
'注意事項:
'

End Sub
