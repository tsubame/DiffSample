VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPixInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Exist_Flag As Boolean
Public OverFlow_Flag As Boolean
Private m_PixInfo() As T_PIXINFO

Public Sub Class_Initialize()
    Exist_Flag = False
    OverFlow_Flag = False
    Dim dumyPix() As T_PIXINFO
    m_PixInfo = dumyPix
End Sub

Public Property Get ALLPixInfo() As T_PIXINFO()
    ALLPixInfo = m_PixInfo
End Property


Public Property Get PixInfo(i As Long) As T_PIXINFO
    PixInfo = m_PixInfo(i)
End Property

Public Function SetPixInfo(ByRef PixInfo() As T_PIXINFO)
    m_PixInfo = PixInfo
    Exist_Flag = True
    Call pixTrim
End Function

Public Function Count() As Double
    If Exist_Flag = False Then
        Count = 0
    Else
        Count = UBound(m_PixInfo) + 1
    End If
End Function

Public Function pixTrim() As Double
    If Exist_Flag = False Then Exit Function
    Dim MaxCount As Double
    Dim nowCount As Double
    MaxCount = Count - 1
    For nowCount = MaxCount To 0 Step -1
        If (m_PixInfo(nowCount).x <> 0 And m_PixInfo(nowCount).y <> 0) Then Exit For
    Next nowCount

    Dim dumyPix() As T_PIXINFO

    If nowCount = -1 Then
        
        Exist_Flag = False
        m_PixInfo = dumyPix
    Else
        If nowCount = MaxCount Then Exit Function
        ReDim Preserve m_PixInfo(nowCount) As T_PIXINFO
    End If
    
End Function
'2014/01/21 H.Arikawa
Public Function CpixMaxCount() As Double
    CpixMaxCount = 5000
'    CpixMaxCount = 1000000
End Function
