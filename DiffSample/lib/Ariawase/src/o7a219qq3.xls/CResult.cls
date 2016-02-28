VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_Result() As Double

Public Property Let Name(ByVal RHS As String)
    m_Name = RHS
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let result(ByVal RHS As Variant)
    m_Result = RHS
End Property

Public Property Get result() As Variant
    result = m_Result
End Property
