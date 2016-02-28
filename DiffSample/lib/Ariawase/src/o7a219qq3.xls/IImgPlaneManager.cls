VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IImgPlaneManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   IImgPlaneManagerのインターフェイス
'目的:
'   ログ出力のオンオフでCImgPlaneManagerとCImgPlaneManagerInterceptor を切り替える
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月27日

Option Explicit

Public Sub GetPlanes(ByRef pPlanes As Collection)
End Sub

Public Sub GetNoUsePlanes(ByRef pPlanes As Collection)
End Sub

Public Sub CreateBasePMD(ByVal pName As String, ByVal pBasePmdName As String, ByVal pWidth As Long, ByVal pHeight As Long, ByRef pColorMap As CImgPlaneMap, ByVal pNumS16 As Long, ByVal pNumS32 As Long, ByVal pNumF32 As Long)
End Sub

'#FlagExpansion
Public Function ReleaseUnusedFlagPlane() As Long
End Function

'#FlagExpansion
Public Function GetSharedFlagPlanes() As Collection
End Function

'#FlagExpansion
Public Function GetSharedFlagPlane(ByVal pFlagName As String) As IImgFlag
End Function

'#FlagExpansion
Public Sub ClearAllFlag()
End Sub

Public Sub CreateSubPMD(ByVal pSubPmdName As String, ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long)
End Sub

Public Function GetFreePlane(ByVal pDepth As IdpBitDepth, Optional ByVal pClear As Boolean = False, Optional ByVal pComment As String = "-") As CImgPlane
End Function

Public Property Get Name() As String
End Property

Public Property Get BasePMD() As CImgPmdInfo
End Property

Public Property Get PMD(ByVal pName As String) As CImgPmdInfo
End Property

Public Property Get Count(ByVal pBitDepth As IdpBitDepth) As Long
End Property

Public Property Get UsedCount(ByVal pBitDepth As IdpBitDepth) As Long
End Property

Public Property Get VariablePMDName() As String
End Property

Public Property Get planeMap() As CImgPlaneMap
End Property

'Public Sub zSetFlagInstance(instance As IImgFlag)
'End Sub
'
'Public Sub EnableInterceptor(ByVal pEnable As Boolean)
'End Sub
'
'
