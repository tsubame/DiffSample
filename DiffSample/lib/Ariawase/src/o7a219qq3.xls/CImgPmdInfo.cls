VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPmdInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   PMD�̏���Z�߂�N���X�B
'
'�ړI:
'   PMD�̏���ۑ��B
'
'�쐬��:
'   0145184004
'
Option Explicit

Dim m_Name As String
Dim m_X As Long
Dim m_Y As Long
Dim m_Width As Long
Dim m_Height As Long


Friend Sub Create(ByVal pName As String, ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long)
'���e:
'   �f�[�^�쐬
'
'[pName]       IN   String�^:       PMD��
'[pX]          IN   Long�^:         ��_X�A�h���X
'[pY]          IN   Long�^:         ��_Y�A�h���X
'[pWidth]      IN   Long�^:         ��
'[pHeight]     IN   Long�^:         ����
'
'���l:
'
    m_Name = pName
    m_X = pX
    m_Y = pY
    m_Width = pWidth
    m_Height = pHeight
End Sub

Public Function Clone() As CImgPmdInfo
'���e:
'   �������g�Ɠ����f�[�^�̃I�u�W�F�N�g��Ԃ�
'
'���l:
'   �����I�Ȏg�p�ɂ������B
'
    Set Clone = New CImgPmdInfo
    Call Clone.Create(m_Name, m_X, m_Y, m_Width, m_Height)
End Function

Public Property Get Name() As String
'���e:
'   ���O�̎擾
'
'���l:
'
    Name = m_Name
End Property

Public Property Get XAdr() As Long
'���e:
'   X���_�̎擾
'
'���l:
'
    XAdr = m_X
End Property

Public Property Let XAdr(ByVal RHS As Long)
'���e:
'   X���_�̐ݒ�
'
'���l:
'
    m_X = RHS
End Property

Public Property Get YAdr() As Long
'���e:
'   Y���_�̎擾
'
'���l:
'
    YAdr = m_Y
End Property

Public Property Let YAdr(ByVal RHS As Long)
'���e:
'   Y���_�̐ݒ�
'
'���l:
'
    m_Y = RHS
End Property

Public Property Get width() As Long
'���e:
'   ���̎擾
'
'���l:
'
    width = m_Width
End Property

Public Property Let width(ByVal RHS As Long)
'���e:
'   ���̐ݒ�
'
'���l:
'
    m_Width = RHS
End Property

Public Property Get height() As Long
'���e:
'   �����̎擾
'
'���l:
'
    height = m_Height
End Property

Public Property Let height(ByVal RHS As Long)
'���e:
'   �����̐ݒ�
'
'���l:
'
    m_Height = RHS
End Property

Public Property Get Left() As Long
'���e:
'   PMD�̍��A�h���X
'
'���l:
'
    Left = m_X
End Property

Public Property Let Left(ByVal RHS As Long)
    m_Width = Right - RHS + 1
    m_X = RHS
End Property

Public Property Get Right() As Long
'���e:
'   PMD�̉E�A�h���X
'
'���l:
'
    Right = m_X + m_Width - 1
End Property

Public Property Let Right(ByVal RHS As Long)
    m_Width = RHS - Left + 1
End Property

Public Property Get Top() As Long
'���e:
'   PMD�̏�A�h���X
'
'���l:
'
    Top = m_Y
End Property

Public Property Let Top(ByVal RHS As Long)
    m_Height = Bottom - RHS + 1
    m_Y = RHS
End Property

Public Property Get Bottom() As Long
'���e:
'   PMD�̉��A�h���X
'
'���l:
'
    Bottom = m_Y + m_Height - 1
End Property

Public Property Let Bottom(ByVal RHS As Long)
    m_Height = RHS - Top + 1
End Property

Public Property Get Size() As Long
'���e:
'   PMD�̃T�C�Y(Read Only)
'
'���l:
'
    Size = m_Width * m_Height
End Property