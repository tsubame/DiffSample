VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IImgPlaneBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   IImgPlaneBank�C���^�[�t�F�C�X
'�ړI:
'   ���O�o�͂̃I���I�t��CImgPlaneBank��CImgPlaneBankInterceptor ��؂�ւ���
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��27��

Option Explicit

Public Sub Add(ByVal pName As String, ByRef pPlane As CImgPlane, Optional ByVal pFixedPlane As Boolean = False, Optional ByVal pReadOnly As Boolean = False)
End Sub

Public Property Get IsOverwriteMode() As Boolean
End Property

Public Property Let IsOverwriteMode(ByRef mode As Boolean)
End Property

Public Function Item(ByVal pName As String) As CImgPlane
End Function

Public Sub Delete(ByVal pName As String)
End Sub

Public Function isExisting(ByVal pName As String) As Boolean
End Function

Public Sub Clear()
End Sub

Public Function List() As String
End Function

Public Property Get Count() As Long
End Property

Public Sub Display()
End Sub
