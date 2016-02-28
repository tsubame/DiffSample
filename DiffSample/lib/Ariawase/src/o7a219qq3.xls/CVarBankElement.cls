VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVarBankElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   CVarBank ��Collection.Add�����N���X
'   �l�₻�̑��̃p�����[�^��ێ�����
'   �ϐ��P�ɂ����̃C���X�^���X�P�����蓖�Ă�
'�ړI:
'   Public �ϐ��̎g�p���������
'   TOPT �̓v���V�[�W�����ŕʂ̃v���V�[�W���ɐ��䂪�ڂ邽�߁APublic�ϐ��ɈӐ}���Ȃ��ύX����������\�����o�Ă���B
'   ��{�A�ǂݎ���p�ɂ��Ă������Ƃŕs�p�ӂȏ������݂��������B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��22��

Option Explicit

Private m_Val As Variant         '�l�̊i�[�p
Private m_ReadOnly As Boolean    '�l�̌Œ��� �i�[�p
Private m_Name As String         '�L�[�i�[�p
Private m_Comment As String      '�R�����g�i�[�p

Public Property Let val(pVal As Variant)
    m_ReadOnly = True
    m_Val = pVal
End Property

Public Property Get val() As Variant
    val = m_Val
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = m_ReadOnly
End Property

Public Property Let ReadOnly(status As Boolean)
    m_ReadOnly = status
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let Comment(Comment As String)
    m_Comment = Comment
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Sub Initialize(Name As String, val As Variant, Comment As String)
    m_Name = Name
    m_Comment = Comment
    m_Val = val
End Sub

Private Sub Class_Initialize()
    m_Name = ""
    m_ReadOnly = True
    m_Val = ""
End Sub

Private Sub Class_Terminate()
'dummy
End Sub