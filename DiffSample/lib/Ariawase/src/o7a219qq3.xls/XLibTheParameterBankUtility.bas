Attribute VB_Name = "XLibTheParameterBankUtility"
'�T�v:
'   TheParameterBank�̃��[�e�B���e�B
'
'   Revision History:
'       Data        Description
'       2011/02/10  ParameterBank��Utility�@�\����������
'
'�쐬��:
'   0145184304
'

Option Explicit

Public TheParameterBank As IParameterBank ' ParameterBank��錾����

Private Const ERR_NUMBER = 9999                           ' Error�ԍ���ێ�����
Private Const CLASS_NAME = "XLibTheParameterBankUtility" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty       ' Default�l"Empty"��ێ�����

Public Sub CreateTheParameterBankIfNothing()
'���e:
'   TheParameterBank������������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
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
'���e:
'   TheParameterBank������������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
End Sub

Public Sub DestroyTheParameterBank()
'���e:
'   TheParameterBank��j������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Set TheParameterBank = Nothing
End Sub

Public Function RunAtJobEnd() As Long
    If Not TheParameterBank Is Nothing Then
        Call TheParameterBank.Clear
    End If
End Function
