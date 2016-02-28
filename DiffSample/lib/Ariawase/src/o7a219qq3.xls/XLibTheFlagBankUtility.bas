Attribute VB_Name = "XLibTheFlagBankUtility"
'�T�v:
'   TheFlagBank��Utility���W���[��
'
'   Revision History:
'       Data        Description
'       2010/10/07  FlagBank��Utility�@�\����������
'       2010/10/28  �R�����g����ǉ����ύX����
'       2011/03/04�@CFlagBank�s��C���ɔ����ύX(by 0145206097)
'                   �_���v���[�h��Ԃ̔��f�������N���X�ֈړ�
'
'�쐬��:
'   0145184346
'

Option Explicit

'/** �p�u���b�N�t���O�o���N�I�u�W�F�N�g **/
Public TheFlagBank As CFlagBank
'/** ���O�t�@�C���� **/
Private mSaveFileName As String

Public Sub CreateTheFlagBankIfNothing()
'���e:
'   TheFlagBank�̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    On Error GoTo ErrHandler
    If TheFlagBank Is Nothing Then Set TheFlagBank = New CFlagBank
    Exit Sub
ErrHandler:
    Set TheFlagBank = Nothing
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

Public Sub SaveModeTheFlagBank(ByVal pDump As Boolean, Optional saveFileName As String)
'���e:
'   TheFlagBank�̃��O�擾���s�Ȃ�
'
'�p�����[�^:
'   [pDump]         In ���O�擾���[�h�w��
'   [SaveFileName]  In ���O�t�@�C����
'
'�߂�l:
'
'���ӎ���:
'
    If TheFlagBank Is Nothing Then Exit Sub
    mSaveFileName = saveFileName
    TheFlagBank.Dump pDump
End Sub

Public Sub DestroyTheFlagBank()
'���e:
'   TheFlagBank��j������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Set TheFlagBank = Nothing
    mSaveFileName = ""
End Sub

Public Function RunAtJobEnd() As Long
'���e:
'   �e�X�g���s�I�����ɁALogFile��ۑ�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    If TheFlagBank Is Nothing Then Exit Function
    TheFlagBank.Save mSaveFileName
End Function
