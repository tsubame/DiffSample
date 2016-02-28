VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamStringWithUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   単位付文字列用パラメータクラス
'
'目的:
'   単位付文字列を単位、補助単位、及び補助単位を考慮しない数値として保持する
'   Double型としての振る舞いをするためDouble型パラメータオブジェクトを持つ
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameter
Implements IProcParameter
Implements ISiteArray
Implements IProcParamReplaceable

Private Const PARAMETER_REPLACER_MAINUNIT As String = "MainUnit"
Private Const PARAMETER_REPLACER_SUBUNIT As String = "SubUnit"
Private Const PARAMETER_REPLACER_SUBVALUE As String = "SubValue"
Private Const PARAMETER_REPLACER_VALUE As String = "Value"
Private Const PARAMETER_REPLACER_TYPE As String = "Type"

Dim mMainUnit As String
Dim mSubUnit As String
Dim mSubValue As Double
Dim mDataValue As CParamDouble

Private Sub Class_Initialize()
    Set mDataValue = New CParamDouble
    With mDataValue.AsIParameter
        .LowerLimit = -9999                  'Corrected since Eee-JOB V3.0
        .UpperLimit = 9999
    End With
End Sub

Private Sub Class_Terminate()
    Set mDataValue = Nothing
End Sub

Public Sub Initialize(ByVal MainUnit As String)
    mMainUnit = MainUnit
End Sub

Private Sub IParameter_Read(ByRef reader As IParameterReader)
    Me.AsIParameter.AsString = reader.ReadAsString(mDataValue.AsIParameter.Name)
End Sub

Private Sub IParameter_WriteOut(ByRef writer As IParameterWriter)
    writer.WriteAsString mDataValue.AsIParameter.Name, Me.AsIParameter.AsString
End Sub

Private Property Let IParameter_AsDouble(ByVal RHS As Double)
    mDataValue.AsIParameter.AsDouble = CDbl(RHS)
    mSubUnit = ""
    mSubValue = mDataValue.AsIParameter.AsDouble
End Property

Private Property Get IParameter_AsDouble() As Double
    IParameter_AsDouble = mDataValue.AsIParameter.AsDouble
End Property

Private Property Let IParameter_AsString(ByVal RHS As String)
    Dim SubUnit As String
    Dim SubValue As Double
    Dim MainUnit As String
    CheckAsString CStr(RHS)
    SplitUnitValue CStr(RHS), MainUnit, SubUnit, SubValue
    If (MainUnit <> mMainUnit) Then TheError.Raise 9999, "CParamStringWithUnit.AsIParameter.AsString()", mDataValue.AsIParameter.Name & " : [" & RHS & "]  - This Is Not Appropriate Main Unit !"
    
    'Corrected since Eee-JOB V3.0
    '  DO NOT Set mSubUnit and mSubValue if mDataValue setting is failed.
    mDataValue.AsIParameter.AsDouble = stringToValue(RHS, SubValue, SubUnit)
    mSubUnit = SubUnit
    mSubValue = SubValue
'    Dim dataValue As Double
'    dataValue = subValue * SubUnitToValue(mSubUnit)
'    mDataValue.AsIParameter.AsDouble = dataValue
End Property

Private Function stringToValue(ByVal strValue As String, ByVal SubValue As Double, ByVal SubUnit As String) As Double
    ''レンジ設定用変換テーブル（丸め誤差の回避）
    Select Case strValue
        Case "200nA":
            stringToValue = 0.0000002
        Case "2uA":
            stringToValue = 0.000002
        Case "5uA":
            stringToValue = 0.000005
        Case "10uA":
            stringToValue = 0.00001
        Case "20uA":
            stringToValue = 0.00002
        Case "40uA":
            stringToValue = 0.00004
        Case "50uA":
            stringToValue = 0.00005
        Case "200uA":
            stringToValue = 0.0002
        Case "500uA":
            stringToValue = 0.0005
        Case "1mA":
            stringToValue = 0.001
        Case "2mA":
            stringToValue = 0.002
        Case "5mA":
            stringToValue = 0.005
        Case "10mA":
            stringToValue = 0.01
        Case "20mA":
            stringToValue = 0.02
        Case "50mA":
            stringToValue = 0.05
        Case "100mA":
            stringToValue = 0.1
        Case "150mA":
            stringToValue = 0.15
        Case "200mA":
            stringToValue = 0.2
        Case "250mA":
            stringToValue = 0.25
        Case "300mA":
            stringToValue = 0.3
        Case "350mA":
            stringToValue = 0.35
        Case "400mA":
            stringToValue = 0.4
        Case "600mA":
            stringToValue = 0.6
        Case "800mA":
            stringToValue = 0.8
        Case Else
            stringToValue = SubValue * SubUnitToValue(SubUnit)
    End Select
End Function

Private Property Get IParameter_AsString() As String
    IParameter_AsString = CStr(mSubValue) & mSubUnit & mMainUnit
End Property

Private Property Let IParameter_AsBoolean(ByVal RHS As Boolean)
    TheError.Raise 9999, "CParamStringWithUnit.AsIParameter.AsBoolean()", "[" & mDataValue.AsIParameter.Name & "] Can Not Be Treated As A Boolean Value !"
End Property

Private Property Get IParameter_AsBoolean() As Boolean
    TheError.Raise 9999, "CParamStringWithUnit.AsIParameter.AsBoolean()", "[" & mDataValue.AsIParameter.Name & "] Can Not Be Treated As A Boolean Value !"
End Property

Private Property Let IParameter_AsLong(ByVal RHS As Long)
    TheError.Raise 9999, "CParamStringWithUnit.AsIParameter.AsLong()", "[" & mDataValue.AsIParameter.Name & "] Can Not Be Treated As A Long Value !"
End Property

Private Property Get IParameter_AsLong() As Long
    TheError.Raise 9999, "CParamStringWithUnit.AsIParameter.AsLong()", "[" & mDataValue.AsIParameter.Name & "] Can Not Be Treated As A Long Value !"
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



Public Function AsIProcParameter() As IProcParameter
    Set AsIProcParameter = Me
End Function

Private Sub IProcParamReplaceable_GetContents(ByVal pRoot As String, ByRef pRes As Collection)
'内容:
'   パラメータの設定内容を取得する
'
'パラメータ:
'   [pRoot]     IN  String型:  このパラメータクラスの階層情報
'   [pRes]      IN/OUT  Collection型:  パラメータの設定内容を格納するコレクション
'
'注意事項:
'
'

    Dim bufSubParam As IProcParamReplaceable
    With pRes
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_MAINUNIT, mMainUnit, "String", False, PARAMETER_REPLACER_MAINUNIT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SUBUNIT, mSubUnit, "String", False, PARAMETER_REPLACER_SUBUNIT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SUBVALUE, CStr(mSubValue), "Double", False, PARAMETER_REPLACER_SUBVALUE)
        Set bufSubParam = mDataValue
        If Not bufSubParam Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_VALUE, vbNullString, "CParamDouble", False, PARAMETER_REPLACER_VALUE)
            Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_VALUE, pRes)
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_VALUE, "Nothing", "CParamDouble", False, PARAMETER_REPLACER_VALUE)
        End If
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_TYPE, TypeName(Me), "String", False, PARAMETER_REPLACER_TYPE)
    End With
End Sub

Private Sub IProcParamReplaceable_ReplaceContent(ByVal pRoot As String, ByRef pParamContents As CImpLocalParamsContent)
'内容:
'   パラメータの内容を書き換える
'
'パラメータ:
'   [pRoot]             IN  String型:  このパラメータクラスの階層情報
'   [pParamContents]    IN  CImpLocalParamsContent型:  書き換えを行うパラメータの内容
'
'注意事項:
'
'

    Dim targetSubParam As IProcParamReplaceable
    With pParamContents
        Dim subParamLabel As String
        subParamLabel = XLibImpExperimentUtility.GetSubParamLabel(.Path, pRoot)
        Select Case subParamLabel
            Case PARAMETER_REPLACER_VALUE
                Set targetSubParam = mDataValue
                Call targetSubParam.ReplaceContent(pRoot & "\" & subParamLabel, pParamContents)
        End Select
    End With
End Sub

Private Sub AddParamContents(ByRef pRes As Collection, ByVal pPath As String, ByVal pValue As String, ByVal pTypeName As String, ByVal pEditable As Boolean, ByVal pLabel As String)
'内容:
'   パラメータの情報をコレクションに追加する
'
'パラメータ:
'   [pRes]          IN/OUT  Collection型:  パラメータの設定内容を追加するコレクション
'   [pPath]         IN  String型:  パラメータの階層情報
'   [pValue]        IN  String型:  パラメータの値
'   [pTypeName]     IN  String型:  パラメータのデータ型
'   [pEditable]     IN  Boolean型:  パラメータが編集可能かどうか
'   [pLabel]        IN  String型:  パラメータ名
'
'注意事項:
'
'

    Dim myContents As New CImpLocalParamsContent
    Call myContents.Initialize(pPath, pValue, pTypeName, pEditable, pLabel)
    Call pRes.Add(myContents)
End Sub

Private Property Let IProcParameter_Comment(ByVal pComment As String)
    mDataValue.AsIProcParameter.Comment = pComment
End Property
Private Property Get IProcParameter_Comment() As String
    IProcParameter_Comment = mDataValue.AsIProcParameter.Comment
End Property

Private Property Let IProcParameter_Name(ByVal pName As String)
    Me.AsIParameter.Name = pName
End Property
Private Property Get IProcParameter_Name() As String
    IProcParameter_Name = Me.AsIParameter.Name
End Property

Private Sub IProcParameter_Dump(ByRef pWriter As IProcParamWriter)
    With pWriter
        .WriteAsString "Name", mDataValue.AsIParameter.Name
        .WriteAsString "Type", TypeName(Me)
        .WriteAsDouble "Value", mDataValue.AsIParameter.AsDouble
        .WriteAsDouble "SubValue", mSubValue
        .WriteAsString "SubUnit", mSubUnit
        .WriteAsString "MainUnit", mMainUnit
        .WriteAsDouble "LowerLimit", mDataValue.AsIParameter.LowerLimit
        .WriteAsDouble "UpperLimit", mDataValue.AsIParameter.UpperLimit
        .WriteAsString "Comment", mDataValue.AsIProcParameter.Comment
    End With
End Sub

Private Function IProcParameter_Clone() As IProcParameter
    Dim ret As CParamStringWithUnit
    Set ret = New CParamStringWithUnit
    ret.Initialize mMainUnit
    With ret.AsIParameter
        .Name = mDataValue.AsIParameter.Name
        .AsString = Me.AsIParameter.AsString
        .LowerLimit = mDataValue.AsIParameter.LowerLimit
        .UpperLimit = mDataValue.AsIParameter.UpperLimit
    End With
    ret.AsIProcParameter.Comment = mDataValue.AsIProcParameter.Comment
    Set IProcParameter_Clone = ret
End Function



Function AsSiteArray() As ISiteArray
    Set AsSiteArray = Me
End Function

Private Function ISiteArray_AsDouble() As Double()
    Dim maxSiteIdx As Long
    maxSiteIdx = GetSiteCount - 1
    Dim ret() As Double
    ReDim ret(maxSiteIdx)
    Dim i As Long
    For i = 0 To maxSiteIdx
        ret(i) = IParameter_AsDouble
    Next i
    ISiteArray_AsDouble = ret
End Function
Private Function GetSiteCount() As Long
    GetSiteCount = TheExec.sites.ExistingCount
End Function

Private Function ISiteArray_AsRaw() As Variant
    ISiteArray_AsRaw = IParameter_AsString
End Function

Public Sub IProcParameter_Read(ByRef pReader As IProcParamReader)
End Sub
