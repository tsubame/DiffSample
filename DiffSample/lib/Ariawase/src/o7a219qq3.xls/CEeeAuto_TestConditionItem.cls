VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEeeAuto_TestConditionItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ARG_MAX As Long = 10

Private m_strConditionName As String
Private m_strFuncName As String
Private m_IsEnable As Boolean
Private m_IsChecked As Boolean
Private m_lRow As Long
Private m_lArgCount As Long
Private m_Arg(ARG_MAX - 1) As Variant

Private Sub Class_Initialize()
    m_strConditionName = ""
    m_strFuncName = ""
    m_IsEnable = False
    m_IsChecked = False
    
    m_lArgCount = 0
    m_lRow = 0
End Sub

Public Property Get ConditionName() As String
    ConditionName = m_strConditionName
End Property

Public Property Get funcName() As String
    funcName = m_strFuncName
End Property

Public Property Get IsChecked() As Boolean
    IsChecked = m_IsChecked
End Property

Public Property Get IsEnable() As Boolean
    IsEnable = m_IsEnable
End Property

Public Property Get Arg(ByVal pArgNumber As Long) As Variant
    Arg = m_Arg(pArgNumber)
End Property

Public Property Get Row() As Long
    Row = m_lRow
End Property

Public Sub SetParams(ByVal strConditionName As String, ByVal strFuncName As String, ByVal lArgCount As Long, ByRef aryArg() As Variant, ByVal lRow As Long, ByVal IsEnable As Boolean)

    m_strConditionName = strConditionName
    m_strFuncName = strFuncName
    m_lArgCount = lArgCount
    
    Dim i As Long
    For i = 0 To lArgCount - 1
        m_Arg(i) = aryArg(i)
    Next
    m_lRow = lRow
    m_IsEnable = IsEnable
    
End Sub


Public Function GetTestConditionIdenfier() As String
    
    'Disableになっているということは
    '・最初からDisable
    '・もしくはチェックの結果Disable、多重に読み込んでいる
    'ということなので、無視しなくてはならない
    If (Not m_IsEnable) Then
        GetTestConditionIdenfier = ""
        Exit Function
    End If

    If m_IsChecked = True Then
        GetTestConditionIdenfier = ""
        Exit Function
    End If
    
    GetTestConditionIdenfier = GetTestConditionIdenfier_impl(m_strConditionName, m_strFuncName, m_lArgCount, m_Arg)
    
End Function

Public Sub SetValideCodition(ByVal IsValid As Boolean)

    m_IsEnable = IsValid
    m_IsChecked = True
    
End Sub


