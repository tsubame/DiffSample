VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgFlagInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   CImgFlag の割り込み用クラス
'目的:
'   ログ出力のオンオフでCImgFlagとCImgFlagInterceptor を切り替える
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月28日

Option Explicit

Implements IImgFlag

Private m_ImgFlag As CImgFlag

Public Property Get TrueInstance() As IImgFlag
    Set TrueInstance = m_ImgFlag.AsIImgFlag
End Property

Public Property Set TrueInstance(Instance As IImgFlag)
    Set m_ImgFlag = Instance
End Property

Public Function AsIImgFlag() As IImgFlag
    Set AsIImgFlag = Me
End Function

'#FlagExpansion
Private Property Get IImgFlag_Group() As String
    IImgFlag_Group = m_ImgFlag.AsIImgFlag.group
End Property

'#FlagExpansion
Private Property Get IImgFlag_Count() As Long
    IImgFlag_Count = m_ImgFlag.AsIImgFlag.Count
End Property

'######## ログ出力対象 ########
Private Property Let IImgFlag_Group(ByVal RHS As String)
    m_ImgFlag.AsIImgFlag.group = RHS
    'ログ出力
    Call TheIDP.PlaneList.NowState("CImgFlag.Group", , , RHS)
End Property

'######## ログ出力対象 ########
Private Property Set IImgFlag_FlgPlane(ByRef pPlane As CImgPlane)
    Set m_ImgFlag.AsIImgFlag.FlgPlane = pPlane
    'ログ出力
    Call TheIDP.PlaneList.NowState("CImgFlag.FlgPlane", pPlane.Name, pPlane.planeMap.Name, , , , , , , , , pPlane.Comment)
End Property

Private Property Get IImgFlag_FlgPlane() As CImgPlane
    Set IImgFlag_FlgPlane = m_ImgFlag.AsIImgFlag.FlgPlane
End Property

'######## ログ出力対象 ########
Private Sub IImgFlag_Clear()
    Call m_ImgFlag.AsIImgFlag.Clear
    'ログ出力
    Call TheIDP.PlaneList.NowState("CImgFlag.Clear")
End Sub

Private Property Get IImgFlag_FlagBit(ByVal pName As String) As Long
    IImgFlag_FlagBit = m_ImgFlag.AsIImgFlag.FlagBit(pName)
End Property

'######## ログ出力対象 ########
Private Function IImgFlag_SetFlagBit(ByVal pName As String) As Long
    IImgFlag_SetFlagBit = m_ImgFlag.AsIImgFlag.SetFlagBit(pName)
    'ログ出力
    Call TheIDP.PlaneList.NowState("CImgFlag.SetFlagBit", m_ImgFlag.AsIImgFlag.Name, m_ImgFlag.AsIImgFlag.FlgPlane.planeMap.Name, , , , , , , pName, m_ImgFlag.AsIImgFlag.FlagBit(pName), m_ImgFlag.AsIImgFlag.FlgPlane.Comment)
End Function

'######## ログ出力対象 ########
Private Sub IImgFlag_RemoveFlagBit(ByVal pName As String)
    'ログ出力
    Call TheIDP.PlaneList.NowState("CImgFlag.RemoveFlagBit", m_ImgFlag.AsIImgFlag.Name, m_ImgFlag.AsIImgFlag.FlgPlane.planeMap.Name, , , , , , , pName, m_ImgFlag.AsIImgFlag.FlagBit(pName), m_ImgFlag.AsIImgFlag.FlgPlane.Comment)
    
    Call m_ImgFlag.AsIImgFlag.RemoveFlagBit(pName)
End Sub

Private Sub IImgFlag_LNot(ByVal pDstName As String, ByVal pSrcName As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
    Call m_ImgFlag.AsIImgFlag.LNot(pDstName, pSrcName, pColor)
End Sub

Private Sub IImgFlag_LOr(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
    Call m_ImgFlag.AsIImgFlag.LOr(pDstName, pSrcName1, pSrcName2, pColor)
End Sub

Private Sub IImgFlag_LAnd(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
    Call m_ImgFlag.AsIImgFlag.LAnd(pDstName, pSrcName1, pSrcName2, pColor)
End Sub

Private Sub IImgFlag_LXor(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
    Call m_ImgFlag.AsIImgFlag.LXor(pDstName, pSrcName1, pSrcName2, pColor)
End Sub

Private Function IImgFlag_SetPMD(ByVal pPmd As Variant) As IImgFlag
    Set IImgFlag_SetPMD = m_ImgFlag.AsIImgFlag.SetPMD(pPmd)
End Function

Private Function IImgFlag_SetCustomPMD(ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long) As IImgFlag
    Set IImgFlag_SetCustomPMD = m_ImgFlag.AsIImgFlag.SetCustomPMD(pX, pY, pWidth, pHeight)
End Function

Private Property Get IImgFlag_Name() As String
    IImgFlag_Name = m_ImgFlag.AsIImgFlag.Name
End Property

Private Property Get IImgFlag_FlagBitList() As String
    IImgFlag_FlagBitList = m_ImgFlag.AsIImgFlag.FlagBitList
End Property

Private Sub Class_Initialize()
    Set m_ImgFlag = New CImgFlag
End Sub

Private Sub Class_Terminate()
    Set m_ImgFlag = Nothing
End Sub


