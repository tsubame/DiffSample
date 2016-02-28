VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneBankInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   IImgPlaneBank���荞�ݗp�N���X
'�ړI:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��27��

Option Explicit

Implements IImgPlaneBank

Private m_ImgPlaneBank As IImgPlaneBank

Public Property Get TrueInstance() As IImgPlaneBank
    Set TrueInstance = m_ImgPlaneBank
End Property

Public Property Set TrueInstance(Instance As IImgPlaneBank)
    Set m_ImgPlaneBank = Instance
End Property

Public Function AsIImgPlaneBank() As IImgPlaneBank
    Set AsIImgPlaneBank = Me
End Function

'######## ���O�o�͑Ώ� ########
Private Sub IImgPlaneBank_Add(ByVal pName As String, ByRef pPlane As CImgPlane, Optional ByVal pFixedPlane As Boolean = False, Optional ByVal pReadOnly As Boolean = False)
    Call m_ImgPlaneBank.Add(pName, pPlane, pFixedPlane, pReadOnly)
    '���O�o��
    Call TheIDP.PlaneList.NowState("CImgPlaneBank.Add", pPlane.Name, pPlane.planeMap.Name, pPlane.planeGroup, , pName, CStr(m_ImgPlaneBank.IsOverwriteMode), , , , , pPlane.Comment)
End Sub

Private Property Get IImgPlaneBank_IsOverwriteMode() As Boolean
    IImgPlaneBank_IsOverwriteMode = m_ImgPlaneBank.IsOverwriteMode
End Property

Private Property Let IImgPlaneBank_IsOverwriteMode(ByRef mode As Boolean)
    m_ImgPlaneBank.IsOverwriteMode = mode
    '���O�o��
    Call TheIDP.PlaneList.NowState("CImgPlaneBank.IsOverwriteMode", , , , , , CStr(mode))
End Property

Private Function IImgPlaneBank_Item(ByVal pName As String) As CImgPlane
    Set IImgPlaneBank_Item = m_ImgPlaneBank.Item(pName)
End Function

'######## ���O�o�͑Ώ� ########
Private Sub IImgPlaneBank_Delete(ByVal pName As String)
    '���O�o��
    On Error GoTo NOTHING_BANK
    Call TheIDP.PlaneList.NowState("CImgPlaneBank.Delete", m_ImgPlaneBank.Item(pName).Name, m_ImgPlaneBank.Item(pName).planeMap.Name, m_ImgPlaneBank.Item(pName).planeGroup, , pName, CStr(m_ImgPlaneBank.IsOverwriteMode), , , , , m_ImgPlaneBank.Item(pName).Comment)
    On Error GoTo 0
    Call m_ImgPlaneBank.Delete(pName)
    Exit Sub
    
NOTHING_BANK:
    Call TheIDP.PlaneList.NowState("CImgPlaneBank.Delete", "Nothing", , "Nothing", , pName, CStr(m_ImgPlaneBank.IsOverwriteMode))
    Call m_ImgPlaneBank.Delete(pName)
End Sub

Private Function IImgPlaneBank_isExisting(ByVal pName As String) As Boolean
    IImgPlaneBank_isExisting = m_ImgPlaneBank.isExisting(pName)
End Function

'######## ���O�o�͑Ώ� ########
Private Sub IImgPlaneBank_Clear()
    Call m_ImgPlaneBank.Clear
    '���O�o��
    Call TheIDP.PlaneList.NowState("CImgPlaneBank.Clear")
End Sub

Private Function IImgPlaneBank_List() As String
    IImgPlaneBank_List = m_ImgPlaneBank.List
End Function

Private Property Get IImgPlaneBank_Count() As Long
    IImgPlaneBank_Count = m_ImgPlaneBank.Count
End Property

Private Sub IImgPlaneBank_Display()
    Call m_ImgPlaneBank.Display
End Sub

Private Sub Class_Initialize()
    Set m_ImgPlaneBank = New CImgPlaneBank
End Sub

Private Sub Class_Terminate()
    Set m_ImgPlaneBank = Nothing
End Sub
