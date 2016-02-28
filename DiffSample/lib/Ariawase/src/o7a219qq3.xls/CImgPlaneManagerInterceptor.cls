VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneManagerInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   IImgPlaneManager割り込み用クラス
'目的:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月27日

Option Explicit

Implements IImgPlaneManager

Private m_PlaneManager As CImgPlaneManager
Private m_GetFreePlaneRetryTimes(2) As Long                         'GetFreePlane失敗回数のカウント用。配列はbitDepth
Private m_ExecutionCount(2) As Long                                    'GetFreePlane失敗回数のカウントリセット用。

Public Property Get TrueInstance() As IImgPlaneManager
    Set TrueInstance = m_PlaneManager.AsIImgPlaneManager
End Property

Public Property Set TrueInstance(Instance As IImgPlaneManager)
    Set m_PlaneManager = Instance
End Property

Public Function AsIImgPlaneManager() As IImgPlaneManager
    Set AsIImgPlaneManager = Me
End Function

'#FlagExpansion
Private Function IImgPlaneManager_ReleaseUnusedFlagPlane() As Long
    IImgPlaneManager_ReleaseUnusedFlagPlane = m_PlaneManager.AsIImgPlaneManager.ReleaseUnusedFlagPlane
End Function

'#FlagExpansion
Private Function IImgPlaneManager_GetSharedFlagPlanes() As Collection
    Set IImgPlaneManager_GetSharedFlagPlanes = m_PlaneManager.AsIImgPlaneManager.GetSharedFlagPlanes
End Function

'#FlagExpansion
Private Sub IImgPlaneManager_ClearAllFlag()
    m_PlaneManager.AsIImgPlaneManager.ClearAllFlag
End Sub

Private Sub IImgPlaneManager_GetPlanes(ByRef pPlanes As Collection)
    Call m_PlaneManager.AsIImgPlaneManager.GetPlanes(pPlanes)
End Sub

Private Sub IImgPlaneManager_GetNoUsePlanes(ByRef pPlanes As Collection)
    Call m_PlaneManager.AsIImgPlaneManager.GetNoUsePlanes(pPlanes)
End Sub

Private Sub IImgPlaneManager_CreateBasePMD(ByVal pName As String, ByVal pBasePmdName As String, ByVal pWidth As Long, ByVal pHeight As Long, ByRef pColorMap As CImgPlaneMap, ByVal pNumS16 As Long, ByVal pNumS32 As Long, ByVal pNumF32 As Long)
    Call m_PlaneManager.AsIImgPlaneManager.CreateBasePMD(pName, pBasePmdName, pWidth, pHeight, pColorMap, pNumS16, pNumS32, pNumF32)
End Sub

'#FlagExpansion
Private Function IImgPlaneManager_GetSharedFlagPlane(ByVal pFlagName As String) As IImgFlag
    Set IImgPlaneManager_GetSharedFlagPlane = m_PlaneManager.AsIImgPlaneManager.GetSharedFlagPlane(pFlagName)
End Function

Private Sub IImgPlaneManager_CreateSubPMD(ByVal pSubPmdName As String, ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long)
    Call m_PlaneManager.AsIImgPlaneManager.CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
End Sub

'######## ログ出力対象 ########
Private Function IImgPlaneManager_GetFreePlane(ByVal pDepth As IdpBitDepth, Optional ByVal pClear As Boolean = False, Optional ByVal pComment As String = "-") As CImgPlane
    
    '#########テスト１回ごとにカウントリセット######
    If m_ExecutionCount(pDepth) <> TheExec.ExecutionCount Then m_GetFreePlaneRetryTimes(pDepth) = 0
    m_ExecutionCount(pDepth) = TheExec.ExecutionCount
    
    On Error GoTo ErrorHandler
    Set IImgPlaneManager_GetFreePlane = m_PlaneManager.AsIImgPlaneManager.GetFreePlane(pDepth, pClear)
    'ログ出力
    Call TheIDP.PlaneList.NowState("CImgPlaneManager.GetFreePlane", IImgPlaneManager_GetFreePlane.Name, IImgPlaneManager_GetFreePlane.planeMap.Name, m_PlaneManager.AsIImgPlaneManager.Name, pDepth, , , "True", , , , pComment)
    Exit Function

ErrorHandler:
    m_GetFreePlaneRetryTimes(pDepth) = m_GetFreePlaneRetryTimes(pDepth) + 1
    Call TheIDP.PlaneList.NowState("CImgPlaneManager.GetFreePlane", , , m_PlaneManager.AsIImgPlaneManager.Name, pDepth, , , "False", CStr(m_GetFreePlaneRetryTimes(pDepth)), , , pComment)
    Call Err.Raise(999, "GetFreePlane", "Plane doesn't suffice. (Max:" & m_PlaneManager.AsIImgPlaneManager.Count(pDepth) & ")")

End Function

Private Property Get IImgPlaneManager_Name() As String
    IImgPlaneManager_Name = m_PlaneManager.AsIImgPlaneManager.Name
End Property

Private Property Get IImgPlaneManager_BasePMD() As CImgPmdInfo
    Set IImgPlaneManager_BasePMD = m_PlaneManager.AsIImgPlaneManager.BasePMD
End Property

Private Property Get IImgPlaneManager_PMD(ByVal pName As String) As CImgPmdInfo
    Set IImgPlaneManager_PMD = m_PlaneManager.AsIImgPlaneManager.PMD(pName)
End Property

Private Property Get IImgPlaneManager_Count(ByVal pBitDepth As IdpBitDepth) As Long
    IImgPlaneManager_Count = m_PlaneManager.AsIImgPlaneManager.Count(pBitDepth)
End Property

Private Property Get IImgPlaneManager_UsedCount(ByVal pBitDepth As IdpBitDepth) As Long
    IImgPlaneManager_UsedCount = m_PlaneManager.AsIImgPlaneManager.UsedCount(pBitDepth)
End Property

Private Property Get IImgPlaneManager_VariablePMDName() As String
    IImgPlaneManager_VariablePMDName = m_PlaneManager.AsIImgPlaneManager.VariablePMDName
End Property

Private Property Get IImgPlaneManager_PlaneMap() As CImgPlaneMap
    Set IImgPlaneManager_PlaneMap = m_PlaneManager.AsIImgPlaneManager.planeMap
End Property

'Private Sub IImgPlaneManager_zSetFlagInstance(instance As IImgFlag)
'    Call m_PlaneManager.AsIImgPlaneManager.zSetFlagInstance(instance)
'End Sub
'
'Private Sub IImgPlaneManager_EnableInterceptor(ByVal pEnable As Boolean)
'    Call m_PlaneManager.AsIImgPlaneManager.EnableInterceptor(pEnable)
'End Sub
'
Private Sub Class_Initialize()
    Set m_PlaneManager = New CImgPlaneManager
    Erase m_ExecutionCount
    Erase m_GetFreePlaneRetryTimes
End Sub

Private Sub Class_Terminate()
    Set m_PlaneManager = Nothing
    Erase m_ExecutionCount
    Erase m_GetFreePlaneRetryTimes
End Sub


