VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamSitePixelLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �T�C�g���̃s�N�Z�����O���Ǘ�����N���X
'
'�ړI:
'   �T�C�g���̃s�N�Z�����O���Ǘ�����
'
'�쐬��:
'   0145184306
'
Option Explicit

Implements IProcParamReplaceable

Private Const ERR_NUMBER As Long = 9999                   ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CParamSitePixelLog" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty       ' Default�l"Empty"��ێ�����

Private Const PARAMETER_REPLACER_SITE As String = "Site"
Private Const PARAMETER_REPLACER_PIXELLOGS As String = "PixelLogs"

Private m_PixelLogs As Collection
Private m_Site As Long

Public Sub Add(ByRef pPixelLog As CParamPixelLog)
'���e:
'   CParamPixelLog��o�^����
'
'����:
'[pPixelLog]     IN/OUT CParamPixelLog�^:  �e�A�h���X�̃s�N�Z�����O
'
'���l:
'
'
    Call m_PixelLogs.Add(pPixelLog)
End Sub

Public Sub Initialize(ByVal pSite As Long)
'���e:
'   ������
'
'����:
'[pSite]     IN Long�^:  �T�C�g�ԍ�
'
'���l:
'
'
    m_Site = pSite
End Sub

Public Property Get PixelLogs() As Collection
'���e:
'   �s�N�Z�����O��Z�߂ĕԂ�
'
'�߂�l:
'   �s�N�Z�����O�R���N�V����
'
'���l:
'
'
    Set PixelLogs = m_PixelLogs
End Property

Public Property Get Count() As Long
'���e:
'   �s�N�Z�����O�̐���Ԃ�
'
'�߂�l:
'   �s�N�Z�����O��
'
'���l:
'
'
    Count = m_PixelLogs.Count
End Property

Public Property Get site() As Long
'���e:
'   �T�C�g�ԍ���Ԃ�
'
'�߂�l:
'   �T�C�g�ԍ�
'
'���l:
'
'
    site = m_Site
End Property
Private Sub Class_Initialize()
    Set m_PixelLogs = New Collection
    m_Site = 0
End Sub

Private Sub Class_Terminate()
    Set m_PixelLogs = Nothing
    m_Site = 0
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
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SITE, CStr(m_Site), "Long", False, PARAMETER_REPLACER_SITE)
        If Not m_PixelLogs Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELLOGS, vbNullString, "Collection", False, PARAMETER_REPLACER_PIXELLOGS)
            itemCounter = 0
            For Each bufSubParam In m_PixelLogs
                If Not bufSubParam Is Nothing Then
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELLOGS & "\Item(" & itemCounter & ")", vbNullString, "CParamPixelLog", False, "Item(" & itemCounter & ")")
                    Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_PIXELLOGS & "\Item(" & itemCounter & ")", pRes)
                Else
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELLOGS & "\Item(" & itemCounter & ")", "Nothing", "CParamPixelLog", False, "Item(" & itemCounter & ")")
                End If
                itemCounter = itemCounter + 1
            Next bufSubParam
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELLOGS, "Nothing", "Collection", False, PARAMETER_REPLACER_PIXELLOGS)
        End If
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

    Dim targetSubParam As IProcParamReplaceable
    Dim targetIndex As Long
    With pParamContents
        Dim subParamLabel As String
        subParamLabel = XLibImpExperimentUtility.GetSubParamLabel(.Path, pRoot)
        Select Case subParamLabel
            Case PARAMETER_REPLACER_PIXELLOGS
                targetIndex = XLibImpExperimentUtility.GetSubParamIndex(.Path, pRoot)
                Set targetSubParam = m_PixelLogs.Item(targetIndex + 1)
                Call targetSubParam.ReplaceContent(pRoot & "\" & subParamLabel & "\Item(" & targetIndex & ")", pParamContents)
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