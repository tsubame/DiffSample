VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �v���[���̊Ǘ�
'
'�ړI:
'   �v���[���̊Ǘ����s���B
'   �T�C�Y�A�J���[�}�b�v�������v���[���𑩂˂�(�r�b�g�Ⴂ�͓Z�߂�)
'
'�쐬��:
'   0145184004
'
Option Explicit

Implements IImgPlaneManager

Private Const POSTFIX_VAR_PMD = "_VAR"

Dim m_Name As String
Dim WithEvents m_Int16 As CImgPlanes
Attribute m_Int16.VB_VarHelpID = -1
Dim WithEvents m_Int32 As CImgPlanes
Attribute m_Int32.VB_VarHelpID = -1
Dim WithEvents m_Flt32 As CImgPlanes
Attribute m_Flt32.VB_VarHelpID = -1

Dim m_PMDs As Collection
Dim m_VarPMDName As String
Dim m_PlaneMap As CImgPlaneMap

'#FlagExpansion
Dim m_FlagPlanes As CImgFlagPlanes

'#FlagExpansion
Public Sub EnableInterceptor(ByVal pEnable As Boolean)
'���̃v���[���}�l�[�W��������Flag�C���X�^���X�؂�ւ��p
    If pEnable = True Then
        m_FlagPlanes.EnableInterceptor
    Else
        m_FlagPlanes.DisableInterceptor
    End If
End Sub

Public Function AsIImgPlaneManager() As IImgPlaneManager
    Set AsIImgPlaneManager = Me
End Function

Private Sub Class_Initialize()
    Set m_PMDs = New Collection
End Sub

'#FlagExpansion
Private Function IImgPlaneManager_ReleaseUnusedFlagPlane() As Long
    IImgPlaneManager_ReleaseUnusedFlagPlane = m_FlagPlanes.ReleaseUnusedFlagPlane
End Function

'#FlagExpansion
Private Sub IImgPlaneManager_ClearAllFlag()
    m_FlagPlanes.ClearAllFlag
End Sub

'#FlagExpansion
Private Function IImgPlaneManager_GetSharedFlagPlanes() As Collection
    Set IImgPlaneManager_GetSharedFlagPlanes = m_FlagPlanes.FlagPlanes
End Function

Private Sub IImgPlaneManager_GetPlanes(ByRef pPlanes As Collection)
    If Not m_Int16 Is Nothing Then Call m_Int16.GetPlanes(pPlanes)
    If Not m_Int32 Is Nothing Then Call m_Int32.GetPlanes(pPlanes)
    If Not m_Flt32 Is Nothing Then Call m_Flt32.GetPlanes(pPlanes)
End Sub

Private Sub IImgPlaneManager_GetNoUsePlanes(ByRef pPlanes As Collection)
    If Not m_Int16 Is Nothing Then Call m_Int16.GetNoUsePlanes(pPlanes)
    If Not m_Int32 Is Nothing Then Call m_Int32.GetNoUsePlanes(pPlanes)
    If Not m_Flt32 Is Nothing Then Call m_Flt32.GetNoUsePlanes(pPlanes)
End Sub

'#FlagExpansion
Private Sub IImgPlaneManager_CreateBasePMD(ByVal pName As String, ByVal pBasePmdName As String, ByVal pWidth As Long, ByVal pHeight As Long, ByRef pColorMap As CImgPlaneMap, ByVal pNumS16 As Long, ByVal pNumS32 As Long, ByVal pNumF32 As Long)
'���e:
'   �x�[�XPMD�쐬
'
'[pName]        IN  String�^:       �v���[����
'[pBasePmdName] IN  String�^:       �x�[�XPMD��
'[pWidth]       IN  Long�^:         ��
'[pHeight]      IN  Long�^:         ����
'[pColorMap]    IN  String�^:       �J���[�}�b�v
'[pNumS16]      IN  Long�^:         16Bit�����v���[���̖���
'[pNumS32]      IN  Long�^:         32Bit�����v���[���̖���
'[pNumF32]      IN  Long�^:         32Bit���������v���[���̖���
'
'���l:
'   CImgIDP�̂݌Ăяo����悤�ɂ������B
'   �G���[����
'   1,�w�肵���x�[�XPMD�̖��O�����ɑ��݂��鎞
'
    If TheIDP.isExistingPMD(pBasePmdName) = True Then
        '�G���[
        Call TheError.Raise(999, "CreateBasePMD", ErrMsgNameAlreadyExists(pBasePmdName))
    End If
    
    Set m_Int16 = New CImgPlanes
    Set m_Int32 = New CImgPlanes
    Set m_Flt32 = New CImgPlanes
        
    Call m_Int16.CreateBasePMD(pBasePmdName, pWidth, pHeight, pName, idpDepthS16, pNumS16, pColorMap.IdpColorsPmd)
    Call m_Int32.CreateBasePMD(pBasePmdName, pWidth, pHeight, pName, idpDepthS32, pNumS32, pColorMap.IdpColorsPmd)
    Call m_Flt32.CreateBasePMD(pBasePmdName, pWidth, pHeight, pName, idpDepthF32, pNumF32, pColorMap.IdpColorsPmd)
    
    m_Name = pName
    
    Set m_FlagPlanes = New CImgFlagPlanes
    m_FlagPlanes.Initialize pName
            
    Call AddPMD(pBasePmdName, 1, 1, pWidth, pHeight)
    
    'CreateBasePMD���ATheIDP������CImgPlaneMaps��CImgPlaneMap�̃R�s�[���APlaneManager�����B
    'GetFreePlane���A�����CImgPlane���R�s�[�����C���X�^���X�����B
    Set m_PlaneMap = pColorMap
    Set pColorMap = Nothing
    
    m_VarPMDName = pName & POSTFIX_VAR_PMD
    Call IImgPlaneManager_CreateSubPMD(m_VarPMDName, 1, 1, 1, 1)
        
End Sub

'#FlagExpansion
Private Function IImgPlaneManager_GetSharedFlagPlane(ByVal pFlagName As String) As IImgFlag

    On Error GoTo INSUFFICIENT_PLANE

    Set IImgPlaneManager_GetSharedFlagPlane = m_FlagPlanes.FlagPlane(pFlagName)
    Exit Function

INSUFFICIENT_PLANE:
    TheError.Raise 999, "GetSharedFlagPlane", "Flag plane is nothing."
    Set IImgPlaneManager_GetSharedFlagPlane = Nothing

End Function

Private Sub IImgPlaneManager_CreateSubPMD(ByVal pSubPmdName As String, ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long)
'���e:
'   �T�uPMD�쐬
'
'[pSubPmdName] IN   String�^:       �T�uPMD��
'[pX]          IN   Long�^:         ��_X�A�h���X
'[pY]          IN   Long�^:         ��_Y�A�h���X
'[pWidth]      IN   Long�^:         ��
'[pHeight]     IN   Long�^:         ����
'
'���l:
'   �G���[����
'   1,�w�肵���T�uPMD�̖��O�����ɑ��݂��鎞
'
'   �T���v���R�[�h
'    Call TheIDP.PlaneManager("vmcu").CreateSubPMD("ZONE2D", 80, 60, 640, 480)
'
    If TheIDP.isExistingPMD(pSubPmdName) = True Then
        '�G���[
        Call TheError.Raise(999, "CreateSubPMD", ErrMsgNameAlreadyExists(pSubPmdName))
    End If
    
    Call m_Int16.CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    Call m_Int32.CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    Call m_Flt32.CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    
    Call AddPMD(pSubPmdName, pX, pY, pWidth, pHeight)

End Sub

Private Function IImgPlaneManager_GetFreePlane(ByVal pDepth As IdpBitDepth, Optional ByVal pClear As Boolean = False, Optional ByVal pComment As String = "-") As CImgPlane
'���e:
'   �󂫃v���[���̎擾
'
'[pDepth]       IN  IdpBitDepth�^:  �r�b�g�w��
'[pClear]       IN  Boolean�^:      �N���A���邩���Ȃ���
'[pComment]     IN  String�^:      �摜�R�����g
'
'���l:
'   �G���[����
'   1,�p�ӂ����v���[�����S�Ďg���Ă��鎞
'
'   �T���v���R�[�h
'    Dim src As CImgPlane
'    Set src = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'
    Dim pPlane As CImgPlane
    Dim pColorMap As CImgPlaneMap
    Set pColorMap = m_PlaneMap
    
    Set pPlane = SelectPlanes(pDepth).GetFreePlane(pColorMap)
    Call pPlane.SetPMD(pPlane.BasePMD.Name)         'TOPT���쎞�̑΍�BPlane�m�ۂ��Ă���SetPMD
    pPlane.Comment = pComment
    If pClear = True Then
        Call pPlane.WritePixel(0)
    End If
    
    Set IImgPlaneManager_GetFreePlane = pPlane
    Set pPlane = Nothing
    Set pColorMap = Nothing
    
End Function

Private Property Get IImgPlaneManager_Name() As String
'���e:
'   ���O�̎擾
'
'���l:
'   �T���v���R�[�h
'    Debug.Print TheIDP.PlaneManager("vmcu").Name
'
    IImgPlaneManager_Name = m_Name
End Property

Private Property Get IImgPlaneManager_BasePMD() As CImgPmdInfo
'���e:
'   BasePMD�̏��擾
'
'���l:
'   �T���v���R�[�h
'    With TheIDP.PlaneManager("vmcu").BasePMD
'        �c�c�c
'    End With
'
    Set IImgPlaneManager_BasePMD = m_PMDs(1).Clone
End Property

Private Property Get IImgPlaneManager_PMD(ByVal pName As String) As CImgPmdInfo
'���e:
'   �w�肵�����O��PMD�̏����擾
'
'[pName]       IN   String�^:       �v���[����
'
'���l:
'

On Error GoTo NOT_FOUND
    Set IImgPlaneManager_PMD = m_PMDs(pName).Clone
    Exit Property

NOT_FOUND:
    Set IImgPlaneManager_PMD = Nothing
End Property

Private Property Get IImgPlaneManager_Count(ByVal pBitDepth As IdpBitDepth) As Long
'���e:
'   �w�肵���r�b�g�[���̃v���[���̖������擾
'
'[pBitDepth]    IN  IdpBitDepth�^:      �r�b�g�[���w��
'
'���l:
'   �T���v���R�[�h
'    Debug.Print TheIDP.PlaneManager("vmcu").Count(idpDepthS16)
'
    IImgPlaneManager_Count = SelectPlanes(pBitDepth).Count
End Property

Private Property Get IImgPlaneManager_UsedCount(ByVal pBitDepth As IdpBitDepth) As Long
'���e:
'   �w�肵���r�b�g�[���̃v���[���̖������擾
'
'[pBitDepth]    IN  IdpBitDepth�^:      �r�b�g�[���w��
'
'���l:
'   �T���v���R�[�h
'    Debug.Print TheIDP.PlaneManager("vmcu").Count(idpDepthS16)
'
    IImgPlaneManager_UsedCount = SelectPlanes(pBitDepth).UsedCount
End Property

Private Property Get IImgPlaneManager_VariablePMDName() As String
'���e:
'   ��PMD�̊�{���擾
'
'���l:
'
    IImgPlaneManager_VariablePMDName = m_VarPMDName
End Property

Private Property Get IImgPlaneManager_PlaneMap() As CImgPlaneMap
'���e:
'   �}�l�[�W���[�����J���[�}�b�v�̏��擾
'
'���l:
'   �T���v���R�[�h
'    With TheIDP.PlaneManager("vmcu").PlaneMap
'        �c�c�c
'    End With
'
    Set IImgPlaneManager_PlaneMap = m_PlaneMap
End Property

'Private Sub IImgPlaneManager_zSetFlagInstance(instance As IImgFlag)
''���̃v���[���}�l�[�W��������Flag�C���X�^���X�؂�ւ��p
'    Set m_FlgPlane = Nothing
'    Set m_FlgPlane = instance
'End Sub
'
Private Sub AddPMD(ByVal pName As String, ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long)
    
    'm_PMDs��PMD��ǉ�����
    Dim newPMD As CImgPmdInfo
    
    Set newPMD = New CImgPmdInfo
    Call newPMD.Create(pName, pX, pY, pWidth, pHeight)
    Call m_PMDs.Add(newPMD, pName)
    
End Sub

Private Function SelectPlanes(ByVal pBitDepth As IdpBitDepth) As CImgPlanes

    Select Case pBitDepth
    Case idpDepthS16
        Set SelectPlanes = m_Int16
    Case idpDepthS32
        Set SelectPlanes = m_Int32
    Case idpDepthF32
        Set SelectPlanes = m_Flt32
    Case Else
        TheError.Raise 999, "SelectPlanes", """" & pBitDepth & """ is Illegal Number."
    End Select
    
End Function

Private Function ErrMsgNameAlreadyExists(ByVal pName As String)
    ErrMsgNameAlreadyExists = """" & pName & """ already exists."
End Function

Private Sub m_Flt32_EVGETMANAGER(ByRef pMng As IImgPlaneManager)
    Set pMng = Me
End Sub

Private Sub m_Int16_EVGETMANAGER(ByRef pMng As IImgPlaneManager)
    Set pMng = Me
End Sub

Private Sub m_Int32_EVGETMANAGER(ByRef pMng As IImgPlaneManager)
    Set pMng = Me
End Sub

'#FlagExpansion
Private Sub Class_Terminate()
'term

Set m_Int16 = Nothing
Set m_Int32 = Nothing
Set m_Flt32 = Nothing
Set m_PMDs = Nothing
Set m_PlaneMap = Nothing
Set m_FlagPlanes = Nothing

End Sub