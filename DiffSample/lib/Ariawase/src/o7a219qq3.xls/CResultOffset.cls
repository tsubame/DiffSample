VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CResultOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_strTestName As String
Private m_arySlope(nSite) As Double
Private m_aryOffset(nSite) As Double
Private m_strUnit As String

Public Property Get Name() As String
    Name = m_strTestName
End Property

Public Property Let Name(ByVal strTestName As String)
    m_strTestName = strTestName
End Property

Public Property Get Slope(ByVal site As Long) As String
    Slope = m_arySlope(site)
End Property

Public Property Get ResultOffset(ByVal site As Long) As String
    ResultOffset = m_aryOffset(site)
End Property


Public Sub SetParam(ByRef arySlope() As Double, ByRef aryOffset() As Double, ByRef strUnit As String)

    Dim site As Long
    
    For site = 0 To nSite
        m_arySlope(site) = arySlope(site)
        m_aryOffset(site) = aryOffset(site)
    Next site
    m_strUnit = strUnit
End Sub

Public Function GetSlope(ByVal site As Long) As Double

    GetSlope = m_arySlope(site)
    
End Function

Public Function GetOffset(ByVal site As Long) As Double

    GetOffset = m_aryOffset(site)
    
End Function

Public Function GetUnit() As String

    GetUnit = m_strUnit
    
End Function

