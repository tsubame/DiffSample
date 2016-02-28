VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUserMacroAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   UserMacro����ێ������s����N���X
'
'   Revision History:
'       Data        Description
'       2010/04/28  UserMacro�����s����@�\����������
'       2010/05/12  �v���O�����R�[�h�𐮗�����
'       2010/05/19  UserMacro���s���ɁA�߂�l���󂯎��悤�ɕύX����
'       2010/05/31  Error�����𒲐�����
'       2010/06/11  �v���O�����R�[�h�𐮗�����
'
'�쐬��:
'   0145184346
'

Option Explicit

Implements IUserMacroAction

Private Const CLASS_NAME As String = "CUserMacroAction" ' Class���̂�ێ�����

Private mMacroName As String ' UserMacro���̂�ێ�����

Public Property Get AsIUserMacroAction() As IUserMacroAction
    Set AsIUserMacroAction = Me
End Property

Private Function IUserMacroAction_Execute() As Integer
'���e:
'   UserMacro�����s����
'
'�p�����[�^:
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    '#####  UserMacro�����s����  #####
    If mMacroName <> "" Then
        On Error GoTo VBAErrHandler
        IUserMacroAction_Execute = Application.Run(mMacroName)
    Else
        Call TheError.Raise(9999, CLASS_NAME & ".Execute", "Can Not Found User Macro Name !")
    End If


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
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