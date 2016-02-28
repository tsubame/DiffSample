VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDatalogWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ���o�̓t�@�C�����_���v�N���X
'
'�ړI:
'   �摜�̃t�@�C���Ǎ�/�t�@�C���o�͎���
'   �t�@�C������Datalog�ɏo�͂���
'
'�쐬��:
'   0145184934
'
Option Explicit

Implements IProcParamWriter

Private mMarker As String

Public Sub IProcParamWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
End Sub

Public Sub IProcParamWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
End Sub

Public Sub IProcParamWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
End Sub

Public Sub IProcParamWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
End Sub

Public Sub IProcParamWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
'���e:
'   �e�L�X�g����Datalog�Ƀ_���v����B
'
'[dataName]    IN String�^:     �f�[�^��
'[strData]     IN String�^:     �f�[�^
'
'���l:
'
    If dataName <> vbNullString Then
        TheExec.Datalog.WriteComment mMarker & dataName & ":" & strData
    Else
        TheExec.Datalog.WriteComment mMarker & strData
    End If
End Sub

Public Sub IProcParamWriter_WriteAsBlockEnd(ByVal pParamName As String)
'���e:
'   �}�[�J�[������������B
'
'[pParamName]    IN String�^:     ������
'
'���l:
'
    mMarker = ""
End Sub

Public Sub IProcParamWriter_WriteAsBlockStart(ByVal pParamName As String)
'���e:
'   �}�[�J�[��ݒ肷��B
'
'[pParamName]    IN String�^:     ������
'
'���l:
'
    mMarker = "#"
End Sub

Public Sub IProcParamWriter_ReturnToRootBlock()
End Sub

Public Property Get IProcParamWritermyName() As String
End Property

Public Sub IProcParamWriter_WriteAsPlane(ByVal pParamName As String, ByRef pImgPlane As CImgPlane)
End Sub

Private Property Get IProcParamWriter_MyName() As String
End Property

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