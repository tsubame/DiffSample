VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMeasurePin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ピンクラス
'
'目的:
'   測定する各ピン名とその結果を格納するテストラベル等を持つ
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWritable

Dim mResult As Collection
Dim mPinName As CParamName
Dim mOperate As CParamString
Dim mTestLabel As CParamName
Dim mUnit As CParamString
Dim mLoLimit As CParamDouble
Dim mHiLimit As CParamDouble
Dim mJudge As CParamLong

Dim mBoardName As CParamString
Dim mBoardRange As CParamDouble
Dim mBoardForce As CParamDouble
Dim mValidResult As CParamLong

Dim mContainer As CContainer

Private Sub Class_Initialize()
    Set mContainer = New CContainer
    Set mPinName = CreateCParamName
    mPinName.AsIParameter.Name = TEST_PINS
    Set mOperate = CreateCParamString
    mOperate.AsIParameter.Name = OPERATE_RESULT
    Set mTestLabel = CreateCParamName
    mTestLabel.AsIParameter.Name = MEASURE_LABEL
    Set mUnit = CreateCParamString
    mUnit.AsIParameter.Name = USERMACRO_UNIT
    Set mLoLimit = CreateCParamDouble
    mLoLimit.AsIParameter.Name = USERMACRO_LOLIMIT
    Set mHiLimit = CreateCParamDouble
    mHiLimit.AsIParameter.Name = USERMACRO_HILIMIT
    Set mJudge = CreateCParamLong
    mJudge.AsIParameter.Name = USERMACRO_JUDGE
    Set mBoardName = CreateCParamString
    mBoardName.AsIParameter.Name = BOARD_NAME
    Set mBoardRange = CreateCParamDouble
    mBoardRange.AsIParameter.Name = BOARD_RANGE
    Set mBoardForce = CreateCParamDouble
    mBoardForce.AsIParameter.Name = BOARD_FORCE
    Set mValidResult = CreateCParamLong
    mValidResult.AsIParameter.Name = VALIDATE_RESULT
    Set mResult = mf_SiteLayOut
    Dim currParam As CParamDouble
    For Each currParam In mResult
        currParam.AsIParameter.Name = EXAMIN_RESULT
    Next currParam
End Sub

Public Sub Initialize(ByVal container As CContainer, ByVal reader As IParameterReader)
    Set mContainer = container
    mPinName.AsIParameter.Read reader
    mOperate.AsIParameter.Read reader
    mTestLabel.AsIParameter.Read reader
End Sub
'#V21-Release
Public Sub SetResult(ByRef resultFromDc As Variant, Optional pSite As Long = ALL_SITE)
    Dim resultFromOperate() As Double
    CalculateTempValue resultFromOperate, resultFromDc, mOperate.AsIParameter.AsString, mContainer, pSite
    Dim dataIndex As Long
    If pSite = ALL_SITE Then
     For dataIndex = 0 To UBound(resultFromOperate)
        mResult.Item(dataIndex + 1).AsIParameter.AsDouble = resultFromOperate(dataIndex)
     Next dataIndex
    Else
        dataIndex = pSite
        mResult.Item(dataIndex + 1).AsIParameter.AsDouble = resultFromOperate(dataIndex)
    End If
End Sub

Public Property Get Name() As String
    Name = mPinName.AsIParameter.AsString
End Property

Public Property Get TestLabel() As String
    TestLabel = mTestLabel.AsIParameter.AsString
End Property

Public Property Let TestLabel(ByVal LabelName As String)
    mTestLabel.AsIParameter.AsString = LabelName
End Property

Public Property Get result() As Double()
    Dim retResult() As Double
    Dim dataIndex As Long
    ReDim retResult(mResult.Count - 1)
    For dataIndex = 0 To UBound(retResult)
        retResult(dataIndex) = mResult.Item(dataIndex + 1).AsIParameter.AsDouble
    Next dataIndex
    result = retResult
End Property

Public Property Let BoardName(ByVal pinType As String)
    mBoardName.AsIParameter.AsString = pinType
End Property

Public Property Let BoardRange(ByRef rangeVal As Variant)
    mBoardRange.AsIParameter.AsDouble = validValue(rangeVal)
End Property

Public Property Let BoardForce(ByRef ForceVal As Variant)
    mBoardForce.AsIParameter.AsDouble = headerValue(ForceVal)
End Property

Public Property Get BoardName() As String
    BoardName = mBoardName.AsIParameter.AsString
End Property

Public Property Get BoardRange() As Variant
    BoardRange = mBoardRange.AsIParameter.AsDouble
End Property

Public Property Get UpperLimit() As Double
    UpperLimit = mHiLimit.AsIParameter.AsDouble
End Property

Public Property Get LowerLimit() As Double
    LowerLimit = mLoLimit.AsIParameter.AsDouble
End Property

Public Property Get JudgeNumber() As Long
    JudgeNumber = mJudge.AsIParameter.AsLong
End Property

Public Property Get Unit() As String
    Unit = mUnit.AsIParameter.AsString
End Property

Public Sub ValidateRange()
    mValidResult.AsIParameter.AsLong = ValidateMeasureRange(Me)
End Sub

Private Property Get IParameterWritable_ParameterList() As Collection
    Set IParameterWritable_ParameterList = New Collection
    With IParameterWritable_ParameterList
        .Add mPinName.AsIParameter.Name
        .Add mTestLabel.AsIParameter.Name
        .Add mOperate.AsIParameter.Name
    End With
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    mTestLabel.AsIParameter.WriteOut writer
    mPinName.AsIParameter.WriteOut writer
    mBoardName.AsIParameter.WriteOut writer
    mUnit.AsIParameter.WriteOut writer
    mBoardRange.AsIParameter.WriteOut writer
    mBoardForce.AsIParameter.WriteOut writer
    mHiLimit.AsIParameter.WriteOut writer
    mLoLimit.AsIParameter.WriteOut writer
    mJudge.AsIParameter.WriteOut writer
    mValidResult.AsIParameter.WriteOut writer
    mOperate.AsIParameter.WriteOut writer
    writer.WriteAsString EXAMIN_RESULT, ComposeParamList(mResult)
End Sub

Public Sub ReadSpecInfo(ByVal reader As IFileStream)
    If mTestLabel.AsIParameter.AsString = NOT_DEFINE Then
        mUnit.AsIParameter.AsString = NOT_DEFINE
    Else
        reader.SetLocation mTestLabel.AsIParameter.AsString
        mUnit.AsIParameter.Read reader
        mLoLimit.AsIParameter.Read reader
        mHiLimit.AsIParameter.Read reader
        mJudge.AsIParameter.Read reader
    End If
End Sub

Public Function AsIParameterWritable() As IParameterWritable
    Set AsIParameterWritable = Me
End Function

Private Function mf_SiteLayOut() As Collection
    Set mf_SiteLayOut = New Collection
    Dim SiteIndex As Long
    For SiteIndex = 0 To GetSiteCount
        mf_SiteLayOut.Add CreateCParamDouble
    Next SiteIndex
End Function

Private Function validValue(ByRef dataVal As Variant) As Double
    If Not IsArray(dataVal) Then
        validValue = dataVal
    Else
        Dim dataIndex As Long
        For dataIndex = 1 To UBound(dataVal)
            If dataVal(0) <> dataVal(dataIndex) Then
                validValue = INVALIDATION_VALUE
                Exit Function
            End If
        Next dataIndex
        validValue = dataVal(0)
    End If
End Function

Private Function headerValue(ByRef dataVal As Variant) As Double
    If Not IsArray(dataVal) Then
        headerValue = dataVal
    Else
        headerValue = dataVal(0)
    End If
End Function

Private Function ComposeParamList(ByVal paramResult As Collection) As String
    Dim currData As CParamDouble
    Dim dataIndex As Long
    For Each currData In paramResult
        If dataIndex = 0 Then
            ComposeParamList = currData.AsIParameter.AsString
        Else
            ComposeParamList = ComposeParamList & "," & currData.AsIParameter.AsString
        End If
        dataIndex = dataIndex + 1
    Next currData
End Function

Private Sub Class_Terminate()
    Set mContainer = Nothing
    Set mResult = Nothing
    Set mPinName = Nothing
    Set mOperate = Nothing
    Set mTestLabel = Nothing
    Set mUnit = Nothing
    Set mUnit = Nothing
    Set mLoLimit = Nothing
    Set mHiLimit = Nothing
    Set mJudge = Nothing
    Set mBoardName = Nothing
    Set mBoardRange = Nothing
    Set mBoardForce = Nothing
    Set mValidResult = Nothing
End Sub
