VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgKernelInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   �J�[�l���̏��
'
'�ړI:
'   �J�[�l���̏��̎��܂Ƃ�
'
'�쐬��:
'   0145184004
'
Option Explicit

Private m_Name As String
Private m_Type As IdpKernelType
Private m_Width As Long
Private m_Height As Long
Private m_ShiftR As Long
Private m_Value As String
Private m_Comment As String


Friend Sub Create(ByVal pName As String, ByVal pType As IdpKernelType, ByVal pWidth As Long, ByVal pHeight As Long, ByVal pShiftR As Long, ByVal pValue As String, Optional ByVal pComment As String = "")

    m_Name = pName
    m_Type = pType
    m_Width = pWidth
    m_Height = pHeight
    m_ShiftR = pShiftR
    m_Value = pValue
    m_Comment = pComment

End Sub

Public Property Get Name() As String
'���e:
'   �J�[�l���̖��O���擾
'
'���l:
'
    Name = m_Name
End Property

Public Property Get KernelType() As IdpKernelType
'���e:
'   �J�[�l���̃^�C�v���擾
'
'���l:
'
    KernelType = m_Type
End Property

Public Property Get width() As Long
'���e:
'   �J�[�l���̕����擾
'
'���l:
'
    width = m_Width
End Property

Public Property Get height() As Long
'���e:
'   �J�[�l���̍������擾
'
'���l:
'
    height = m_Height
End Property

Public Property Get ShiftR() As Long
'���e:
'   �J�[�l���̃V�t�g�ʂ��擾
'
'���l:
'
    ShiftR = m_ShiftR
End Property

Public Property Get Value() As String
'���e:
'   �J�[�l���̒l���擾
'
'���l:
'
    Value = m_Value
End Property

Public Property Get Comment() As String
'���e:
'   �J�[�l���̃R�����g���擾
'
'���l:
'
    Comment = m_Comment
End Property
