VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImpLocalParamsContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ���[�J���p�����[�^�̏���ێ�����
'
'�ړI:
'   �p�����[�^�N���X�̃����o�[�ҏW�̏���ێ�����
'
'�쐬��:
'   0145184931
'
Option Explicit

Private mPath As String         '�p�����[�^�̃p�X��ێ�
Private mValue As String        '�p�����[�^�̒l��ێ�
Private mTypeName As String     '�p�����[�^�̃f�[�^�^����ێ�
Private mEditable As Boolean    '�p�����[�^���ҏW�\����ێ�
Private mLabel As String        '�p�����[�^�̃��x������ێ�

Public Sub Initialize(ByVal pPath As String, ByVal pValue As String, ByVal pTypeName As String, ByVal pEditable As Boolean, ByVal pLabel As String)
'���e:
'   �p�����[�^��������������
'
'[pPath]        IN String�^:    �p�����[�^�̃p�X
'[pValue]       IN String�^:    �p�����[�^�̒l
'[pTypeName]    IN String�^:    �p�����[�^�̃f�[�^�^����ێ�
'[pEditable]    IN Boolean�^:   �p�����[�^���ҏW�\����ێ�
'[pLabel]       IN String�^:    �p�����[�^�̃��x����
'
'���l:
'
    mPath = pPath
    mValue = pValue
    mTypeName = pTypeName
    mEditable = pEditable
    mLabel = pLabel
End Sub

Public Function Clone() As CImpLocalParamsContent
    Dim myClone As CImpLocalParamsContent
    Set myClone = New CImpLocalParamsContent
    Call myClone.Initialize(mPath, mValue, mTypeName, mEditable, mLabel)
    Set Clone = myClone
End Function

Public Property Get Path() As String
    Path = mPath
End Property

Public Property Let Path(ByVal pPath As String)
    mPath = pPath
End Property

Public Property Get Value() As String
    Value = mValue
End Property

Public Property Let Value(ByVal pValue As String)
    mValue = pValue
End Property

Public Property Get TypeName() As String
    TypeName = mTypeName
End Property

Public Property Let TypeName(ByVal pTypeName As String)
    mTypeName = pTypeName
End Property

Public Property Get Editable() As Boolean
    Editable = mEditable
End Property

Public Property Let Editable(ByVal pEditable As Boolean)
    mEditable = pEditable
End Property

Public Property Get Label() As String
    Label = mLabel
End Property

Public Property Let Label(ByVal pLabel As String)
    mLabel = pLabel
End Property