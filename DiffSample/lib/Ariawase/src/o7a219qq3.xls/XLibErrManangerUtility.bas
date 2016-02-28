Attribute VB_Name = "XLibErrManangerUtility"
'�T�v:
'   TheError�̃��[�e�B���e�B
'
'�ړI:
'   TheError:CErrManager�̏�����/�j���̃��[�e���e�B���`����
'
'�쐬��:
'   a_oshima

Option Explicit

Public TheError As CErrManager

Public Sub CreateTheErrorIfNothing()
'���e:
'   As New�̑�ւƂ��ď���ɌĂ΂���C���X�^���X��������
'   �����̃N���A���s��
'
'�p�����[�^:
'   �Ȃ�
'
'���ӎ���:
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
