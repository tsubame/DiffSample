VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   名前用パラメータクラス
'
'目的:
'   名前としての文字列、及びその文字数を数値として保持する
'   Long型としての振る舞いをするためLong型パラメータオブジェクトを持つ
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameter

Dim mStrData As String
Dim mDataValue As CParamLong
'Dim mHighLimit As Long
'Dim mLowLimit As Long

Private Sub Class_Initialize()
    Set mDataValue = New CParamLong
    With mDataValue.AsIParameter
        .LowerLimit = 0                  'Corrected since Eee-JOB V3.0
        .UpperLimit = 50
    End With
End Sub

Private Sub Class_Terminate()
    Set mDataValue = Nothing
End Sub

Private Sub IParameter_Read(ByRef reader As IParameterReader)
    Me.AsIParameter.AsString = reader.ReadAsString(mDataValue.AsIParameter.Name)
End Sub

Private Sub IParameter_WriteOut(ByRef writer As IParameterWriter)
    writer.WriteAsString mDataValue.AsIParameter.Name, Me.AsIParameter.AsString
End Sub

Private Property Let IParameter_AsDouble(ByVal RHS As Double)
    TheError.Raise 9999, "CParamName.AsIParameter.AsDouble()", "[" & mDataValue.AsIParameter.Name & "] Can Not Be Treated As A Double Value !"
End Property

Private Property Get IParameter_AsDouble() As Double
    TheError.Raise 9999, "CParamName.AsIParameter.AsDouble()", "[" & mDataValue.AsIParameter.Name & "] Can Not Be Treated As A Double Value !"
End Property

Private Property Let IParameter_AsString(ByVal RHS As String)
    CheckAsString CStr(RHS)
    checkAsName CStr(RHS)
    mDataValue.AsIParameter.AsLong = Len(RHS)
    mStrData = CStr(RHS)
End Property

Private Property Get IParameter_AsString() As String
    IParameter_AsString = mStrData
End Property

Private Property Let IParameter_AsBoolean(ByVal RHS As Boolean)
    TheError.Raise 9999, "CParamName.AsIParameter.AsBoolean()", "[" & mDataValue.AsIParameter.Name & "] Can Not Be Treated As A Boolean Value !"
End Property

Private Property Get IParameter_AsBoolean() As Boolean
    TheError.Raise 9999, "CParamName.AsIParameter.AsBoolean()", "[" & mDataValue.AsIParameter.Name & "] Can Not Be Treated As A Boolean Value !"
End Property

Private Property Let IParameter_AsLong(ByVal RHS As Long)
    TheError.Raise 9999, "CParamName.AsIParameter.AsLong()", "[" & mDataValue.AsIParameter.Name & "] Can Not Be Treated As A Long Value !"
End Property

Private Property Get IParameter_AsLong() As Long
    TheError.Raise 9999, "CParamName.AsIParameter.AsLong()", "[" & mDataValue.AsIParameter.Name & "] Can Not Be Treated As A Long Value !"
End Property

Private Property Let IParameter_Name(ByVal RHS As String)
    mDataValue.AsIParameter.Name = RHS
End Property

Private Property Get IParameter_Name() As String
    IParameter_Name = mDataValue.AsIParameter.Name
End Property

Private Property Let IParameter_UpperLimit(ByVal RHS As Double)
    mDataValue.AsIParameter.UpperLimit = RHS
End Property

Private Property Get IParameter_UpperLimit() As Double
    IParameter_UpperLimit = mDataValue.AsIParameter.UpperLimit
End Property

Private Property Let IParameter_LowerLimit(ByVal RHS As Double)
    mDataValue.AsIParameter.LowerLimit = RHS
End Property

Private Property Get IParameter_LowerLimit() As Double
    IParameter_LowerLimit = mDataValue.AsIParameter.LowerLimit
End Property

Public Function AsIParameter() As IParameter
    Set AsIParameter = Me
End Function

Private Sub checkAsName(ByVal dataStr As String)
    Dim CHAR As String
    CHAR = Left$(dataStr, 1)
    If IsNumber(CHAR) Then
        TheError.Raise 9999, "CParamName.CheckAsName", mDataValue.AsIParameter.Name & " : [" & dataStr & "]  - The Number Of Head Of This String Are Invalid !"
    End If
End Sub
