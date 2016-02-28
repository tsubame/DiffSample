VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUserMacroAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   UserMacro情報を保持＆実行するクラス
'
'   Revision History:
'       Data        Description
'       2010/04/28  UserMacroを実行する機能を実装した
'       2010/05/12  プログラムコードを整理した
'       2010/05/19  UserMacro実行時に、戻り値を受け取るように変更した
'       2010/05/31  Error処理を調整した
'       2010/06/11  プログラムコードを整理した
'
'作成者:
'   0145184346
'

Option Explicit

Implements IUserMacroAction

Private Const CLASS_NAME As String = "CUserMacroAction" ' Class名称を保持する

Private mMacroName As String ' UserMacro名称を保持する

Public Property Get AsIUserMacroAction() As IUserMacroAction
    Set AsIUserMacroAction = Me
End Property

Private Function IUserMacroAction_Execute() As Integer
'内容:
'   UserMacroを実行する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    '#####  UserMacroを実行する  #####
    If mMacroName <> "" Then
        On Error GoTo VBAErrHandler
        IUserMacroAction_Execute = Application.Run(mMacroName)
    Else
        Call TheError.Raise(9999, CLASS_NAME & ".Execute", "Can Not Found User Macro Name !")
    End If


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
VBAErrHandler:
    IUserMacroAction_Execute = TL_ERROR
    Call TheError.Raise(9999, CLASS_NAME & ".Execute", "Can Not Execute User Macro ! ( " & mMacroName & " )")
End Function

Private Function IUserMacroAction_GetUserMacroName() As String
    IUserMacroAction_GetUserMacroName = mMacroName
End Function

Private Property Let IUserMacroAction_SetUserMacroName(ByVal strName As String)
    mMacroName = strName
End Property

Private Sub IUserMacroAction_Initialize()
    mMacroName = ""
End Sub

Private Sub Class_Initialize()
    mMacroName = ""
End Sub

Private Sub Class_Terminate()
    mMacroName = ""
End Sub
