VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamPixelLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �s�N�Z�����O�����ʃZ�b�g��ێ�����N���X
'
'�ړI:
'   �e�s�N�Z���̃s�N�Z�����O��ێ�����
'
'�쐬��:
'   0145184306
'
Option Explicit

Implements IProcParamReplaceable

Private Const ERR_NUMBER As Long = 9999                   ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CParamPixelLog"     ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty       ' Default�l"Empty"��ێ�����

Private Const PARAMETER_REPLACER_COLOR As String = "Color"
Private Const PARAMETER_REPLACER_VALUE As String = "Value"
Private Const PARAMETER_REPLACER_X As String = "x"
Private Const PARAMETER_REPLACER_Y As String = "y"

Private m_Color As CImgColor
Private m_Level As Double
Private m_XAddress As Long
Private m_YAddress As Long

Public Sub Initialize(ByVal pXAdd As Long, ByVal pYAdd As Long, ByVal pLevel As Double, ByRef pColor As CImgColor)
'���e:
'   ������
'
'����:
'[pXAdd]     IN     Long�^:  X�A�h���X
'[pYAdd]     IN     Long�^:  Y�A�h���X
'[pLevel]    IN     Double�^:  ���x��
'[pColor]    IN/OUT CImgColor�^:  �F���
'
'���l:
'
'
    Set m_Color = pColor
    m_Level = pLevel
    m_XAddress = pXAdd
    m_YAddress = pYAdd
End Sub

Public Property Get color() As CImgColor
'���e:
'   �ێ������F����Ԃ�
'
'�߂�l:
'   �ێ������F���
'
'���l:
'
    Set color = m_Color
End Property

Public Property Get Level() As Double
'���e:
'   �ێ��������x������Ԃ�
'
'�߂�l:
'   �ێ��������x�����
'
'���l:
'
    Level = m_Level
End Property

Public Property Get x() As Long
'���e:
'   �ێ�����X�A�h���X��Ԃ�
'
'�߂�l:
'   �ێ�����X�A�h���X
'
'���l:
'
    x = m_XAddress
End Property

Public Property Get y() As Long
'���e:
'   �ێ�����Y�A�h���X��Ԃ�
'
'�߂�l:
'   �ێ�����Y�A�h���X
'
'���l:
'
    y = m_YAddress
End Property

Private Sub Class_Initialize()
    Set m_Color = Nothing
    m_Level = 0
    m_XAddress = 0
    m_YAddress = 0
End Sub

Private Sub Class_Terminate()
    Set m_Color = Nothing
    m_Level = 0
    m_XAddress = 0
    m_YAddress = 0
End Sub

Private Sub IProcParamReplaceable_GetContents(ByVal pRoot As String, ByRef pRes As Collection)
'���e:
'   �p�����[�^�̐ݒ���e���擾����
'
'�p�����[�^:
'   [pRoot]     IN  String�^:  ���̃p�����[�^�N���X�̊K�w���
'   [pRes]      IN/OUT  Collection�^:  �p�����[�^�̐ݒ���e���i�[����R���N�V����
'
'���ӎ���:
'
'

    With pRes
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COLOR, m_Color.Name, "String", False, PARAMETER_REPLACER_COLOR)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_VALUE, CStr(m_Level), "Double", True, PARAMETER_REPLACER_VALUE)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_X, CStr(m_XAddress), "Long", True, PARAMETER_REPLACER_X)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_Y, CStr(m_YAddress), "Long", True, PARAMETER_REPLACER_Y)
    End With
End Sub

Private Sub IProcParamReplaceable_ReplaceContent(ByVal pRoot As String, ByRef pParamContents As CImpLocalParamsContent)
'���e:
'   �p�����[�^�̓��e������������
'
'�p�����[�^:
'   [pRoot]             IN  String�^:  ���̃p�����[�^�N���X�̊K�w���
'   [pParamContents]    IN  CImpLocalParamsContent�^:  �����������s���p�����[�^�̓��e
'
'���ӎ���:
'
'

    With pParamContents
        Dim subParamLabel As String
        subParamLabel = XLibImpExperimentUtility.GetSubParamLabel(.Path, pRoot)
        Select Case subParamLabel
            Case PARAMETER_REPLACER_VALUE
                m_Level = CDbl(.Value)
            Case PARAMETER_REPLACER_X
                m_XAddress = CLng(.Value)
            Case PARAMETER_REPLACER_Y
                m_YAddress = CLng(.Value)
        End Select
    End With
End Sub

Private Sub AddParamContents(ByRef pRes As Collection, ByVal pPath As String, ByVal pValue As String, ByVal pTypeName As String, ByVal pEditable As Boolean, ByVal pLabel As String)
'���e:
'   �p�����[�^�̏����R���N�V�����ɒǉ�����
'
'�p�����[�^:
'   [pRes]          IN/OUT  Collection�^:  �p�����[�^�̐ݒ���e��ǉ�����R���N�V����
'   [pPath]         IN  String�^:  �p�����[�^�̊K�w���
'   [pValue]        IN  String�^:  �p�����[�^�̒l
'   [pTypeName]     IN  String�^:  �p�����[�^�̃f�[�^�^
'   [pEditable]     IN  Boolean�^:  �p�����[�^���ҏW�\���ǂ���
'   [pLabel]        IN  String�^:  �p�����[�^��
'
'���ӎ���:
'
'

    Dim myContents As New CImpLocalParamsContent
    Call myContents.Initialize(pPath, pValue, pTypeName, pEditable, pLabel)
    Call pRes.Add(myContents)
End Sub