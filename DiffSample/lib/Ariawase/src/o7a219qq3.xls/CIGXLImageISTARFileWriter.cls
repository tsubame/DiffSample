VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIGXLImageISTARFileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   IDV�`���ŉ摜�o�͂��郉�C�^�[
'
'�ړI:
'   �w�肳�ꂽ�t�H���_�[��ISTAR�`���̃t�@�C���������ĉ摜���o�͂���
'
'�쐬��:
'   0145184306
'
Option Explicit

Implements IProcParamWriter

Private Const ERR_NUMBER = 9999         ' �G���[�ԍ���ێ�����
Private Const CLASS_NAME = "CIGXLImageISTARFileWriter" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default�l"Empty"��ێ�����

Private m_FilePath As String
Private mParamName As String
Private mImgFileName() As String

Public Function AsIProcParamWriter() As IProcParamWriter
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IProcParamWriter�^�I�u�W�F�N�g
'
'���l:
'
    Set AsIProcParamWriter = Me
End Function

Public Sub Initialize(ByVal pFolder As String, pParamName As String)
'���e:
'   �C�j�V�����C�Y
'   ���C�^�[���t�@�C���o�͂���t�H���_���Z�b�g
'   �t�H���_�̑��݊m�F���s��
'   �p�����[�^�����Z�b�g
'   �o�̓t�@�C�����o�͂̂��߁ACDatalogWriter�I�u�W�F�N�g�𐶐�
'
'����:
'[strFileName]  IN String�^:     �������ݐ�t�H���_
'[pParamName[ IN String�^�F      �p�����[�^��
'
'���l:
'
    If Dir(pFolder, vbDirectory) = "" Then
        TheError.Raise ERR_NUMBER, CLASS_NAME & ".Initialize", "This Folder doesn't exist!"
    Else
        m_FilePath = pFolder
    End If
    mParamName = pParamName
End Sub

Private Sub IProcParamWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
End Sub

Private Sub IProcParamWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
End Sub

Private Sub IProcParamWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
End Sub

Private Sub IProcParamWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
End Sub

Private Sub IProcParamWriter_WriteAsPlane(ByVal pParamName As String, ByRef pImgPlane As CImgPlane)
'���e:
'   �w��Plane�摜���o�͂���
'
'����:
'[pParamName]  IN String�^:     �e�X�g�C���X�^���X��
'[pParamPlane] IN CImgPlane�^:  �o��Plane
'
'���l:
'
'
    Dim site As Long
    Dim ImgFileName As String
    ReDim mImgFileName(TheExec.sites.ExistingCount - 1)

    If m_FilePath = INITIAL_EMPTY_VALUE Then
        TheError.Raise ERR_NUMBER, CLASS_NAME & ".WriteAsPlane", "Initialize Function is not done!!"
        Exit Sub
    Else
        For site = 0 To TheExec.sites.ExistingCount - 1
            ImgFileName = makeISTARImgFileName(site, mParamName)
            ImgFileName = m_FilePath & "\" & ImgFileName
            Call pImgPlane.WriteFile(site, ImgFileName)
            mImgFileName(site) = ImgFileName
            'File���d�����
            TheHdw.WAIT 1 * S
        Next site
    End If
End Sub

Private Sub IProcParamWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
End Sub

Private Property Get IProcParamWriter_MyName() As String
    IProcParamWriter_MyName = CLASS_NAME
End Property

Private Function makeISTARImgFileName(ByVal pSite As Long, ByVal pInstanceName As String) As String
'���e:
'   ISTAR�`���̃t�@�C�������쐬����
'
'����:
'[pSite]       IN/OUT Integer�^:    �T�C�g
'[pInstanceName] IN/OUT String�^:   �e�X�g�C���X�^���X��
'
'�߂�l:
'   ISTAR�`���t�@�C����
'
'���l:
'
'
    makeISTARImgFileName = "Test_" & pSite & ".stb"    'for debug

    Dim site As Long
    Dim strDeviceType As String
    Dim lngLotNo As Long
    Dim lngWaferNo As Long
    Dim lngChipNo As Long
    Dim lngXAddress As Long
    Dim lngYAddress As Long
    Dim lngDate As Variant

    Const REDUCE_VAL As Integer = 1   '�k���{���W��
    Const FILE_EXTENSION As String = ".stb"
    '�V�X�e�����擾
    With TheSystemInfo
        strDeviceType = .ProductName
        lngLotNo = .LotNo
        lngWaferNo = .WaferNo
        lngChipNo = .GetChipNoOfSite(pSite)
        lngXAddress = .GetXAddressOfSite(pSite)
        lngYAddress = .GetYAddressOfSite(pSite)
    End With
    '���t�擾
    lngDate = Now

    'ISTAR�p�t�@�C�������쐬
    makeISTARImgFileName = strDeviceType & "_LOTNAME" & Format(CStr(lngLotNo), "00") & "-" & _
                                Format(CStr(lngWaferNo), "00") & Format(CStr(lngChipNo), "0000") & "-" & _
                                    Format(CStr(lngXAddress), "00") & "-" & Format(CStr(lngYAddress), "00") & "-" & _
                                        pInstanceName & "-" & REDUCE_VAL & "-" & Format(CStr(lngDate), "yyyymmddHHMMSS") & FILE_EXTENSION

End Function

Private Sub Class_Initialize()
    m_FilePath = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    m_FilePath = INITIAL_EMPTY_VALUE
End Sub

Private Sub IProcParamWriter_WriteAsBlockEnd(ByVal pParamName As String)
End Sub

Private Sub IProcParamWriter_WriteAsBlockStart(ByVal pParamName As String)
End Sub

Public Sub IProcParamWriter_ReturnToRootBlock()
End Sub

Public Property Get ImgFileName(ByVal pIndex As Integer) As String
    ImgFileName = mImgFileName(pIndex)
End Property
