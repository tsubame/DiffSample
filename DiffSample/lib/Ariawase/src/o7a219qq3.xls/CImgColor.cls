VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �F�A�܂���idpColorType���Ƃ̊Ǘ��N���X
'�ړI:
'   ���ColorMap�ɂ�����AidpColorType��`�P���̏����Ǘ����܂��B
'   �Ή��F�̊Y���A�h���X��Collection�őS�Ď����܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N1��31��

Option Explicit

Implements IProcParamReplaceable

Private Const PARAMETER_REPLACER_NAME As String = "Name"
Private Const PARAMETER_REPLACER_IDPCOLOR As String = "IdpColor"
Private Const PARAMETER_REPLACER_PIXELS As String = "Pixels"
Private Const PARAMETER_REPLACER_MAPNAME As String = "MapName"

Private m_ColorName As String
Private m_IdpColorType As IdpColorType
Private m_Pixels As Collection
Private m_ParentName As String

Public Sub Init(p_ParentName As String, p_ColorName As String, p_IdpColorType As IdpColorType, p_Pixels As Collection)
'���e:
'   ���̃N���X�̏���ݒ肵�܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��1��
'�p�����[�^:
'   [p_ParentName]      In  1):���̐e�N���X�̖��O
'   [p_ColorName]       In  2):���̃N���X�̐F�̖��O
'   [p_IdpColorType]    In  3):���̃N���X��idpColorType
'   [p_Pixels]          In  4):���̃N���X�̐F���Y������S�A�h���X��Collection
'�߂�l:
'
'���ӎ���:
'
    m_ParentName = p_ParentName
    m_ColorName = p_ColorName
    m_IdpColorType = p_IdpColorType
    Set m_Pixels = p_Pixels
End Sub

Public Function Clone() As CImgColor
'���e:
'   ���̃N���X�̃N���[����Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��3��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   CImgColor
'
'���ӎ���:
'
'

    Dim p_Color As CImgColor
    Set p_Color = New CImgColor
    
    Dim p_Pixel As Variant
    Dim p_Pixels As Collection
    
    Set p_Pixels = New Collection
    
    For Each p_Pixel In m_Pixels
        Call p_Pixels.Add(p_Pixel.Clone)
    Next p_Pixel
    
    Call p_Color.Init(m_ParentName, m_ColorName, m_IdpColorType, p_Pixels)
    Set Clone = p_Color
    
End Function

Public Property Get Pixels() As Collection
    Set Pixels = m_Pixels
End Property

Public Property Get MapName() As String
    MapName = m_ParentName
End Property

Public Property Get Name() As String
    Name = m_ColorName
End Property

Public Property Get IdpColor() As IdpColorType
    IdpColor = m_IdpColorType
End Property

Private Sub Class_Terminate()
    Set m_Pixels = Nothing
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

    Dim bufSubParam As IProcParamReplaceable
    Dim itemCounter As Long
    With pRes
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_NAME, m_ColorName, "String", False, PARAMETER_REPLACER_NAME)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_IDPCOLOR, CStr(m_IdpColorType), "IdpColorType", False, PARAMETER_REPLACER_IDPCOLOR)
        If Not m_Pixels Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELS, vbNullString, "Collection", False, PARAMETER_REPLACER_PIXELS)
            itemCounter = 0
            For Each bufSubParam In m_Pixels
                If Not bufSubParam Is Nothing Then
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELS & "\Item(" & itemCounter & ")", vbNullString, "CImgPixel", False, "Item(" & itemCounter & ")")
                    Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_PIXELS & "\Item(" & itemCounter & ")", pRes)
                Else
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELS & "\Item(" & itemCounter & ")", "Nothing", "CImgPixel", False, "Item(" & itemCounter & ")")
                End If
                itemCounter = itemCounter + 1
            Next bufSubParam
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELS, "Nothing", "Collection", False, PARAMETER_REPLACER_PIXELS)
        End If
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_MAPNAME, m_ParentName, "String", False, PARAMETER_REPLACER_MAPNAME)
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