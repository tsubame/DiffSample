VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPowerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'STD_CommonDCMod が必要

Option Explicit

Private m_ForceVBase As Double
Private m_ForceVOffset() As Double
Private m_ClampI As Double

Private Sub Class_Initialize()

    ReDim m_ForceVOffset(TheExec.sites.ExistingCount - 1)
    
End Sub

Public Property Let ClampI(ByVal argCurrent As Double)

    m_ClampI = argCurrent
    
End Property

Public Property Get ClampI() As Double

    ClampI = m_ClampI
    
End Property

Public Property Let ForceVBase(ByVal argVolt As Double)
    m_ForceVBase = argVolt
End Property

Public Property Let ForceVOffset(ByVal argSite As Long, ByVal argVolt As Double)

    m_ForceVOffset(argSite) = argVolt
    
End Property

Public Property Get ForceV() As Variant

    Dim tmpForceV() As Double
    Dim site As Long

    ReDim tmpForceV(UBound(m_ForceVOffset))
    For site = 0 To UBound(m_ForceVOffset)
        tmpForceV(site) = m_ForceVBase + m_ForceVOffset(site)
    Next site

    ForceV = tmpForceV

End Property

Public Sub Force(ByVal argPinName As String)

    Call SetFVMI(argPinName, ForceV, ClampI)
    
End Sub
'#No-Release
Public Sub ForceMulti(ByVal argPinName As String)

'注意事項:
'          設定値は全サイトがサイト0の設定値固定

    Call SetFVMIMulti(argPinName, ForceV, ClampI)
    
End Sub

Public Sub ForceWithOffset(ByVal argPinName As String, Optional ByVal offsetV As Double = 0)

    Dim i As Long
    Dim realForceV(nSite) As Double
    If IsArray(ForceV) Then
        For i = 0 To UBound(ForceV)
            realForceV(i) = ForceV(i) + offsetV
        Next i
    Else
        For i = 0 To UBound(realForceV)
            realForceV(i) = ForceV + offsetV
        Next i
    End If
        
    Call SetFVMI(argPinName, realForceV, ClampI)

End Sub
'
'Public Sub Force(ByVal argPinName As String)
'
'    Call SetFVMI(argPinName, ForceV, ClampI)
'
'End Sub
