Attribute VB_Name = "XLibTheSystemInfoUtility"
'�T�v:
'   TheSystemInfo�̃��[�e�B���e�B
'
'   Revision History:
'       Data        Description
'       2011/02/10  SystemInfo��Utility�@�\����������
'
'�쐬��:
'   0145184306
'

Option Explicit

Public TheSystemInfo As CSystemInfo ' SystemInfo��錾����

Private Const ERR_NUMBER = 9999                           ' Error�ԍ���ێ�����
Private Const CLASS_NAME = "XLibTheSystemInfoUtility"     ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty       ' Default�l"Empty"��ێ�����

Public Sub CreateTheSystemInfoIfNothing()
'���e:
'   TheSystemInfo������������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
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
'���e:
'   TheSystemInfo��j������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Set TheSystemInfo = Nothing
End Sub

Public Function RunAtJobEnd() As Long
End Function

