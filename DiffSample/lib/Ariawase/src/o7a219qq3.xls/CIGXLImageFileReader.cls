VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIGXLImageFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   stb�t�@�C���摜�ǂݍ��݃N���X
'
'�ړI:
'   stb�t�@�C���̉摜��ǂݍ��ރ��[�_�[
'
'�쐬��:
'   0145184306
'
Option Explicit

Implements IProcParamReader

Private Const ERR_NUMBER = 9999         ' �G���[�ԍ���ێ�����
Private Const CLASS_NAME = "CIGXLImageFileReader" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default�l"Empty"��ێ�����

Private m_FilePath As String
Private mParamName As String
Private mImgFileName() As String

Public Function AsIProcParamReader() As IProcParamReader
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IParameterReaderX20�^�I�u�W�F�N�g
'
'���l:
'
    Set AsIProcParamReader = Me
End Function

Public Sub Initialize(pFolder As String, pParamName As String)
'���e:
'   �C�j�V�����C�Y
'   ���[�_�[���������s���t�H���_���Z�b�g
'   �t�H���_�̑��݊m�F���s��
'   �p�����[�^�����Z�b�g
'   ���̓t�@�C�����o�͂̂��߁ACDatalogWriter�I�u�W�F�N�g�𐶐�
'
'����:
'[strFileName]  IN String�^:     �������ݐ�t�H���_
'[pParamName]  IN String�^:      �p�����[�^��
'
'���l:
'
    '�t�H���_�̑��݊m�F
    If Dir(pFolder, vbDirectory) = "" Then
        TheError.Raise ERR_NUMBER, CLASS_NAME & ".Initialize", "This Folder doesn't exist!"
    Else
        m_FilePath = pFolder
    End If
    mParamName = pParamName
End Sub

Private Function IProcParamReader_ReadAsBoolean(ByVal dataName As String) As Boolean
    TheError.Raise ERR_NUMBER, CLASS_NAME, "Unsupported Value Type!!"
End Function


Private Function IProcParamReader_ReadAsDouble(ByVal dataName As String) As Double
    TheError.Raise ERR_NUMBER, CLASS_NAME, "Unsupported Value Type!!"
End Function

Private Function IProcParamReader_ReadAsInteger(ByVal dataName As String) As Integer
    TheError.Raise ERR_NUMBER, CLASS_NAME, "Unsupported Value Type!!"
End Function

Private Function IProcParamReader_ReadAsLong(ByVal dataName As String) As Long
    TheError.Raise ERR_NUMBER, CLASS_NAME, "Unsupported Value Type!!"
End Function

Private Function IProcParamReader_ReadAsPlane(ByVal pParamName As String, ByRef pImgPlane As CImgPlane)
'���e:
'   �w��Plane��stb�t�@�C���摜��ǂݍ���
'
'����:
'[pParamName]  IN String�^:     �C���X�^���X��
'[pImgPlane]   IN/OUT CImgPlane�^:  �摜��荞�ݗpPlane
'
'���l:
'   �t�@�C������pParamName_site.stb�ɌŒ�
'

    Dim site As Long
    Dim ImgFileName As String
    ReDim mImgFileName(TheExec.sites.ExistingCount - 1)
    
    For site = 0 To TheExec.sites.ExistingCount - 1
        ImgFileName = makeReadFileName(site, mParamName)
        ImgFileName = m_FilePath & "\" & ImgFileName
        '�t�@�C���̑��݊m�F
        If Dir(ImgFileName, vbNormal) = "" Then GoTo ErrorHandle
        With pImgPlane
            Call .ReadFile(site, ImgFileName)
        End With
        mImgFileName(site) = ImgFileName
    Next site
    Exit Function
ErrorHandle:
    TheError.Raise ERR_NUMBER, CLASS_NAME & ".ReadAsPlane", ImgFileName & " File does'nt exist!!"
End Function

Private Function IProcParamReader_ReadAsString(ByVal dataName As String) As String
    TheError.Raise ERR_NUMBER, CLASS_NAME, "Unsupported Value Type!!"
End Function

Private Property Get IProcParamReader_MyName() As String
    IProcParamReader_MyName = CLASS_NAME
End Property

Private Sub Class_Initialize()
    m_FilePath = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    m_FilePath = INITIAL_EMPTY_VALUE
End Sub

Private Function makeReadFileName(ByVal pSite As Long, ByVal pParamName As String) As String
'���e:
'   ISTAR�`���̃t�@�C�������쐬����B
'
'����:
'[pSite]       IN Integer�^:    �T�C�g
'[pParamName] IN String�^:   �e�X�g�C���X�^���X��
'
'�߂�l:
'   �t�@�C����
'
'���l:
'   stb�t�@�C���̂�
'
    Const FILE_EXTENSION As String = ".stb"
    
    makeReadFileName = pParamName & "_" & pSite & FILE_EXTENSION
    
End Function

Public Function IProcParamReader_ReadAsBlockEnd(ByVal pParamName As String) As IProcParamReader
End Function

Public Function IProcParamReader_ReadAsBlockStart(ByVal pParamName As String) As IProcParamReader
End Function

Public Property Get ImgFileName(ByVal pIndex As Integer) As String
    ImgFileName = mImgFileName(pIndex)
End Property
