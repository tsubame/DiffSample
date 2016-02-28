VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �t�@�C���̏o�͊Ǘ��N���X
'�ړI:
'   �t�@�C���o�͂̃R�[�f�B���O��������B
'   �C���X�^���X���t�@�C���̑����Ƃ���B
'�g����:
'   �ENew �ŃC���X�^���X����
'   �ESetFileName �ŕۑ��t�@�C�����ݒ�B
'   �E�ȉ��̖��߂Ńe�L�X�g���t�@�C���ɏo��
'     �t�@�C�����o�͎��Ɏg�����߂͂S��
'   �@�@OpenFile   �F�t�@�C�����J��
'   �@�AFileNum    �F�o�͗p�BPrint #�C���X�^���X��.FileNum, �o�͕�����E�E�E �Ǝg��
'   �@�BCloseFile  �F�t�@�C�������
'
'     �CPrintOut   �F��L�R�� �@���A(Print)���B�܂Ƃ߂Ď��s����
'�L�q�T���v��:
'   Dim pFileOut As CFileOut
'   Set pFileOut = New CFileOut
'   Call pFileOut.SetFileName("tmp.txt")
'   Call pFileOut.PrintOut("hogehogehoge")
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��26��

Option Explicit

Public Enum eWriteMode
    OverWrite
    AddWrite
End Enum

Private m_WriteMode As eWriteMode
Private m_FileNum As Integer
Private m_FileName As String                '�t���p�X�t���t�@�C����
Private m_OrgFileName As String             '���[�U�[���ݒ肵���t�@�C����

Public Sub SetFileName(ByVal fileName As String, Optional WriteMode As eWriteMode = AddWrite)
'���e:
'   �t�@�C�����̐ݒ�
'   �t�@�C�����̂݁A�t���p�X���ǂ�����B
'   �t�@�C�����݂̂̏ꍇ�A�A�N�e�B�u���[�N�u�b�N�̃f�B���N�g���ɕۑ��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��26��
'�p�����[�^:
'   [FileName]  In      1):FileName�̎w��B
'�߂�l:
'
'���ӎ���:

    m_OrgFileName = fileName
    If (fileName Like "?:*") Or (fileName Like "\\*") Then
        m_FileName = fileName                                               '�t�@�C�����Ƀ��[�J���h���C�u��or�l�b�g���[�N�h���C�u�����܂�ł�����A���O���̂܂�
    Else
        m_FileName = ActiveWorkbook.Path & "\" & fileName                   '�t�@�C���������Ȃ�A�A�N�e�B�u���[�N�u�b�N�̃t���p�X��ǉ�
    End If
    m_WriteMode = WriteMode
'    '�t�@�C�����J���邩�����Ă݂�
'    Call OpenFile
'    Call CloseFile

End Sub

Public Property Get fileNum() As Integer
'���e:
'   ���̃C���X�^���X���J���Ă���t�@�C���ԍ���Ԃ��܂��B
'   �J���ĂȂ���΁A�[���ŃG���[�B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��26��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Integer
'
'���ӎ���:

    If m_FileNum > 0 Then
        fileNum = m_FileNum
    Else
        Call Err.Raise(9999, , "File is not opened. =>" & m_FileName)
    End If

End Property

Public Sub OpenFile()
'���e:
'   SetFileName �Őݒ肵�����Ńt�@�C�����J��
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��26��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
    
    On Error GoTo ErrHandler
    m_FileNum = FreeFile                                                      '�P�`�Q�T�T�̋󂢂Ă�t�@�C���ԍ����擾
    If m_WriteMode = OverWrite Then
        Open m_FileName For Output As m_FileNum                                   '�㏑���Ńt�@�C�����J��
    Else
        Open m_FileName For Append As m_FileNum                                   '�ǋL�Ńt�@�C�����J��
    End If
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    Call Err.Raise(9999, , "File Open is Failed! => " & m_FileName)

End Sub

Public Sub CloseFile()
'���e:
'   �J�����t�@�C�������
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��26��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:

    '�t�@�C�����J�������Ƃ����� or �����Ă��� => �������Ȃ��B
    If m_FileNum <> 0 Then
        Close #m_FileNum
        m_FileNum = 0                   '�������Ƃ̖����p
    End If
End Sub

Public Sub PrintOut(PrintMessage As String)
'���e:
'   ���̃C���X�^���X�ŊǗ�����t�@�C���� Open �� �������� �� Close ���Z�b�g�Ŏ��s
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��26��
'�p�����[�^:
'   [PrintMessage]  In  1):�t�@�C���ɏ������ޕ���
'�߂�l:
'
'���ӎ���:

    Call OpenFile
    Print #m_FileNum, PrintMessage
    Call CloseFile

End Sub

Public Property Get GetFileName() As String
'���e:
'   ���̃C���X�^���X�������Ă���t�@�C������Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��26��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   String
'
'���ӎ���:
    GetFileName = m_OrgFileName
End Property

Private Sub Class_Initialize()
'���e:
'   �R���X�g���N�^
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��26��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:

    m_FileNum = 0
    m_FileName = ""
End Sub

Private Sub Class_Terminate()
'���e:
'   �t�@�C���������ĂȂ����Colse���s�B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��26��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
    
    '�t�@�C�����J����Ă��Ȃ� or �����Ă���Ή������Ȃ��B
    If m_FileNum <> 0 Then
        Close #m_FileNum
        m_FileNum = 0
    End If
    m_FileName = ""
End Sub

