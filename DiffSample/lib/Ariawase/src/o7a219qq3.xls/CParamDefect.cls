VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamDefect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �s�N�Z�����O�̃Z�b�g�������N���X
'
'�ړI:
'   �s�N�Z�����O���Ǘ�����
'
'�쐬��:
'   0145184306
'
Option Explicit

Implements IProcParameter
Implements IProcParamReplaceable

Private Const ERR_NUMBER As Long = 9999                  ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CParamDefect"       ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default�l"Empty"��ێ�����

Private Const PARAMETER_REPLACER_NAME As String = "Name"
Private Const PARAMETER_REPLACER_DEVICEINFO As String = "DeviceInfo"
Private Const PARAMETER_REPLACER_SITES As String = "Sites"
Private Const PARAMETER_REPLACER_COMMENT As String = "Comment"

Private m_Comment As String
Private m_Name As String
Private m_DeviceInfo As CDeviceConfigInfo
Private m_Sites As Collection

Public Function AsIProcParameter() As IProcParameter
    Set AsIProcParameter = Me
End Function

Private Function IProcParameter_Clone() As IProcParameter
'���e:
'   ���g�̃N���[����Ԃ�
'
'�߂�l:
'   IProcParameter�^�N���X
'
'���l:
'
'
    
    '##### ���g�������Ă���DeviceInfo�̃N���[������� #####
    Dim objDevInfo As CDeviceConfigInfo
    If Not Me.DeviceConfig Is Nothing Then
        Set objDevInfo = Me.DeviceConfig.AsIProcParameter.Clone
    Else
        Set objDevInfo = Nothing
    End If
        
    '#####  ���g��Clone���쐬���Ė߂�  #####
    Dim objClone As CParamDefect
    Set objClone = New CParamDefect

    objClone.AsIProcParameter.Name = m_Name
    Set objClone.DeviceConfig = objDevInfo
    objClone.AsIProcParameter.Comment = m_Comment
    
    Dim lSite As Long
    For lSite = 1 To m_Sites.Count
        If m_Sites.Item(lSite) Is Nothing Then
            Call objClone.Add(Nothing, lSite - 1)
        Else
            Call objClone.Add(m_Sites.Item(lSite), m_Sites.Item(lSite).site)
        End If
    Next lSite
    
    Set IProcParameter_Clone = objClone
    
    Set objClone = Nothing
    Set objDevInfo = Nothing

End Function

Private Sub IProcParameter_Dump(ByRef pWriter As IProcParamWriter)
'���e:
'   ���g�̏������C�^�[�Ƀ_���v����
'
'����:
'[pWriter]     IN/OUT IProcParamWriter�^:  ���C�^�[
'
'���l:
'
'
    
    pWriter.WriteAsString "Parameter Type", m_Name
    If Not m_DeviceInfo Is Nothing Then
        m_DeviceInfo.AsIProcParameter.Dump pWriter
    End If
        
    Dim lSite As Long
    Dim vPixelLog As Variant
    
    For lSite = 0 To TheExec.sites.ExistingCount - 1
        If Me.site(lSite) Is Nothing Then
            pWriter.WriteAsString "Site = ", lSite
            pWriter.WriteAsString "PixelLogs", "Nothing"
        Else
            pWriter.WriteAsString "Site", Me.site(lSite).site
            For Each vPixelLog In Me.site(lSite).PixelLogs
                pWriter.WriteAsString m_Name, "x = " & vPixelLog.x
                pWriter.WriteAsString m_Name, "y = " & vPixelLog.y
                pWriter.WriteAsString m_Name, "Color = " & vPixelLog.color.Name
                pWriter.WriteAsString m_Name, "Value = " & vPixelLog.Level
            Next vPixelLog
        End If
    Next lSite
    
End Sub

Public Sub Add(ByRef pPixelSiteLog As CParamSitePixelLog, ByVal pSite As Long)
'���e:
'   ���g�̏������C�^�[�Ƀ_���v����
'
'����:
'[pPixelSiteLog]     IN/OUT CParamSitePixelLog�^:  �e�T�C�g�̃s�N�Z�����O
'
'���l:
'
'
    On Error GoTo err_handle
    Call m_Sites.Add(pPixelSiteLog, CStr(pSite))
    Exit Sub
err_handle:
    TheError.Raise ERR_NUMBER, CLASS_NAME, "This Site(" & pPixelSiteLog.site & ") has alreadey added."
End Sub

Public Function CountSites() As Long
'���e:
'   �o�^���ꂽ�T�C�g����Ԃ�
'
'���l:
'
'
    CountSites = m_Sites.Count
End Function
Public Function site(ByVal pSite As Long) As CParamSitePixelLog
'���e:
'   �w��T�C�g�̃s�N�Z�����O��Ԃ�
'
'����:
'[pSite]     IN Long�^:  �T�C�g�ԍ�
'
'�߂�l:
'   CParamSitePixelLog�^�N���X
'
'���l:
'
'
    On Error GoTo err_handle
    Set site = m_Sites.Item(CStr(pSite))
    Exit Function
err_handle:
    TheError.Raise ERR_NUMBER, CLASS_NAME, "There is not CParamSitePixelLog as Site(" & pSite & ")"
End Function

Public Property Get DeviceConfig() As CDeviceConfigInfo
'���e:
'   �ێ������f�o�C�X����Ԃ�
'
'�߂�l:
'   �ێ������f�o�C�X���
'
'���l:
'
    Set DeviceConfig = m_DeviceInfo
End Property

Public Property Set DeviceConfig(ByRef pDeviceConfig As CDeviceConfigInfo)
'���e:
'   �f�o�C�X����ێ�����
'
'����:
'[pDeviceInfo]   IN/OUT CDeviceConfigInfo�^:     �ێ�����f�o�C�X�����w�肷��
'
'���l:
'
    Set m_DeviceInfo = pDeviceConfig
End Property

Private Property Let IProcParameter_Name(ByVal pName As String)
'���e:
'   �p�����[�^����ێ�����
'
'[pName]   IN String�^:     �ێ�����p�����[�^�����w�肷��
'
'���l:
'

    m_Name = pName

End Property

Private Property Get IProcParameter_Name() As String
'���e:
'   �ێ������p�����[�^����Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^��
'
'���l:
'

    IProcParameter_Name = m_Name

End Property

Private Property Let IProcParameter_Comment(ByVal pComment As String)
'���e:
'   �R�����g��ێ�����
'
'[pComment]   IN String�^:     �ێ�����R�����g���w�肷��
'
'���l:
'

    m_Comment = pComment

End Property

Private Property Get IProcParameter_Comment() As String
'���e:
'   �ێ������R�����g��Ԃ�
'
'�߂�l:
'   �ێ������R�����g
'
'���l:
'

    IProcParameter_Comment = m_Comment

End Property

Private Sub Class_Initialize()
    m_Comment = INITIAL_EMPTY_VALUE
    m_Name = INITIAL_EMPTY_VALUE
    Set m_DeviceInfo = Nothing
    Set m_Sites = New Collection
End Sub

Private Sub Class_Terminate()
    m_Comment = INITIAL_EMPTY_VALUE
    m_Name = INITIAL_EMPTY_VALUE
    Set m_DeviceInfo = Nothing
    Set m_Sites = Nothing
End Sub

Public Sub IProcParameter_Read(ByRef pReader As IProcParamReader)
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
        '   2012-01-27 (0145184931): Name��ҏW�s�ɕύX
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_NAME, m_Name, "String", False, PARAMETER_REPLACER_NAME)
        Set bufSubParam = m_DeviceInfo
        If Not bufSubParam Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, vbNullString, "CDeviceConfigInfo", False, PARAMETER_REPLACER_DEVICEINFO)
            Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, pRes)
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, "Nothing", "CDeviceConfigInfo", False, PARAMETER_REPLACER_DEVICEINFO)
        End If
        If Not m_Sites Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SITES, vbNullString, "Collection", False, PARAMETER_REPLACER_SITES)
            itemCounter = 0
            For Each bufSubParam In m_Sites
                If Not bufSubParam Is Nothing Then
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SITES & "\Item(" & itemCounter & ")", vbNullString, "CParamSitePixelLog", False, "Item(" & itemCounter & ")")
                    Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_SITES & "\Item(" & itemCounter & ")", pRes)
                Else
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SITES & "\Item(" & itemCounter & ")", "Nothing", "CParamSitePixelLog", False, "Item(" & itemCounter & ")")
                End If
                itemCounter = itemCounter + 1
            Next bufSubParam
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SITES, "Nothing", "Collection", False, PARAMETER_REPLACER_SITES)
        End If
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COMMENT, m_Comment, "String", True, PARAMETER_REPLACER_COMMENT)
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
'   2012-01-27 (0145184931): Name��ҏW�s�ɕύX
'            Case PARAMETER_REPLACER_NAME
'                m_Name = .value
            Case PARAMETER_REPLACER_SITES
                targetIndex = XLibImpExperimentUtility.GetSubParamIndex(.Path, pRoot)
                Set targetSubParam = m_Sites.Item(targetIndex + 1)
                Call targetSubParam.ReplaceContent(pRoot & "\" & subParamLabel & "\Item(" & targetIndex & ")", pParamContents)
            Case PARAMETER_REPLACER_COMMENT
                m_Comment = .Value
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