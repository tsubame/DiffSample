Attribute VB_Name = "XLibWkShtReaderUtility"
'�T�v:
'   ReaderManager�̃��[�e�B���e�B
'
'�ړI:
'   ReaderManager�̏�����/�j���̃��[�e���e�B���`����
'
'�쐬��:
'   a_oshima

Option Explicit

Private mReaderManager As CWorkSheetReaderManager

Public Sub CreateReaderManagerIfNothing()
'���e:
'   As New�̑�ւƂ��ď���ɌĂ΂���C���X�^���X��������
'
'�p�����[�^:
'   �Ȃ�
'
'���ӎ���:
'
    On Error GoTo ErrHandler
    If mReaderManager Is Nothing Then
        Set mReaderManager = New CWorkSheetReaderManager
        Call mReaderManager.GetReaderInstance(eSheetType.shtTypeDeviceConfigurations)
#If ITS <> 0 Then
        Call mReaderManager.GetReaderInstance(eSheetType.shtTypeImgTestScenario)
#End If
    End If
    Exit Sub
ErrHandler:
    Set mReaderManager = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function GetWkShtReaderManagerInstance() As CWorkSheetReaderManager
'���e:
'   ReaderManager�̃C���X�^���X��Ԃ�
'
'�p�����[�^:
'   �Ȃ�
'
'�߂�l:
'   ReaderManager�̃C���X�^���X
'
'��O:
'   �����������ɌĂ΂���VBA��O����
'  �i�p�t�H�[�}���X���P�̂���AsNew�̑�ւƂ��ėp�ӂ��Ă���ANothing�`�F�b�N�͍s��Ȃ��j
'
'���ӎ���:
'   �������������ɌĂсA�C���X�^���X����������Ă��邱��

    Set GetWkShtReaderManagerInstance = mReaderManager
End Function

Public Sub DestroyWkShtReaderManager()
    Set mReaderManager = Nothing
End Sub


