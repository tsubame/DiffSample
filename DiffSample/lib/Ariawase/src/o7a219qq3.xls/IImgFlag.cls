VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IImgFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   IImgFlag のインターフェイス
'目的:
'   ログ出力のオンオフでCImgFlagとCImgFlagInterceptor を切り替える
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月28日

Option Explicit

Public Property Let group(ByVal RHS As String)
End Property

'#FlagExpansion
Public Property Get group() As String
End Property

'#FlagExpansion
Public Property Get Count() As Long
End Property

Public Property Set FlgPlane(ByRef pPlane As CImgPlane)
End Property

Public Property Get FlgPlane() As CImgPlane
End Property

Public Sub Clear()
End Sub

Public Property Get FlagBit(ByVal pName As String) As Long
End Property

Public Function SetFlagBit(ByVal pName As String) As Long
End Function

Public Sub RemoveFlagBit(ByVal pName As String)
End Sub

Public Sub LNot(ByVal pDstName As String, ByVal pSrcName As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
End Sub

Public Sub LOr(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
End Sub

Public Sub LAnd(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
End Sub

Public Sub LXor(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
End Sub

Public Function SetPMD(ByVal pPmd As Variant) As IImgFlag
End Function

Public Function SetCustomPMD(ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long) As IImgFlag
End Function

Public Property Get Name() As String
End Property

Public Property Get FlagBitList() As String
End Property

