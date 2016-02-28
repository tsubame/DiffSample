VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_BitDepth As IdpBitDepth
Private m_Planes() As String
Private m_UsedCount As Long

Public Event EVGETMANAGER(ByRef pMng As IImgPlaneManager)
Public Event EVGETPLANES(ByVal pPlanes As Collection)

Friend Sub GetPlanes(ByVal pPlanes As Collection)
    RaiseEvent EVGETPLANES(pPlanes)
End Sub

Friend Sub GetNoUsePlanes(ByVal pPlanes As Collection)
    Dim i As Integer
    For i = 0 To Count - 1
        If m_Planes(i) <> "" Then
            Call pPlanes.Add(m_Planes(i))
        End If
    Next i
End Sub

Friend Sub CreateBasePMD(ByVal pBasePmdName As String, ByVal pWidth As Long, ByVal pHeight As Long, ByVal pPlaneName As String, ByVal pBitDepth As IdpBitDepth, ByVal PNum As Long, ByVal pColorMapDef As String)

    Dim i As Long
    
    'エラー処理
    If PNum < 0 Then
        Call TheError.Raise(999, "", "Invalid Parameter. (pNum:" & PNum & ")")
    End If
    
    m_BitDepth = pBitDepth
    ReDim m_Planes(PNum)
    
    If Count > 0 Then
        On Error GoTo ERROR_IGXL
        Call TheHdw.IDP.CreateBasePMD(pBasePmdName & Postfix, pBasePmdName & Postfix, 1, 1, pWidth, pHeight, pPlaneName & Postfix, PNum, pBitDepth, 1, 1, pColorMapDef)
        
        For i = 0 To Count - 1
            m_Planes(i) = pPlaneName & Postfix & Format$(i, "00")
        Next i
    End If
    
    Exit Sub
    
'エラー処理
ERROR_IGXL:
    Call ErrMsgIGXL("CreateBasePMD")

End Sub

Friend Sub CreateSubPMD(ByVal pPmdName As String, ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long)

    If Count > 0 Then
        On Error GoTo ERROR_IGXL
        Call TheHdw.IDP.CreateSubPMD(Manager.BasePMD.Name & Postfix, pPmdName & Postfix, pX, pY, pWidth, pHeight)
    End If

    Exit Sub
    
'エラー処理
ERROR_IGXL:
    Call ErrMsgIGXL("CreateSubPMD")
    
End Sub

Friend Function GetFreePlane(ByRef pColorMap As CImgPlaneMap) As CImgPlane
    
    Dim i As Long
    
    For i = 0 To Count - 1
        If m_Planes(i) <> "" Then
            Set GetFreePlane = New CImgPlane
            
            Call GetFreePlane.Create(m_Planes(i), m_BitDepth, Me, pColorMap)
            m_Planes(i) = ""
            
            If m_UsedCount < i + 1 Then
                m_UsedCount = i + 1
            End If
            
'2009/09/03 D.Maruyama Createの中でSetPMD(Base)を行うので不要
'            Call GetFreePlane.SetPMD(GetFreePlane.BasePMD.Name)
            Exit Function
        End If
    Next i
    
    'エラー
    Call TheError.Raise(999, "GetFreePlane", "Plane doesn't suffice. (Max:" & Count & ")")
    
End Function

Friend Sub ReleasePlane(ByRef pPlane As CImgPlane)

    Dim i As Long
    
    i = Right$(pPlane.Name, 2)
    m_Planes(i) = pPlane.Name
    
End Sub

Friend Function Postfix() As String

    Select Case m_BitDepth
    Case idpDepthS16
        Postfix = "_S"
        
    Case idpDepthS32
        Postfix = "_L"
    
    Case idpDepthF32
        Postfix = "_F"
    
    End Select
    
End Function

Friend Property Get Count() As Long
    Count = UBound(m_Planes)
End Property

Friend Property Get UsedCount() As Long
    UsedCount = m_UsedCount
End Property

Friend Property Get Manager() As IImgPlaneManager
    RaiseEvent EVGETMANAGER(Manager)
End Property

Private Sub ErrMsgIGXL(ByVal pErrSource As String)
    Call TheError.Raise(999, pErrSource, Err.Description)
End Sub

