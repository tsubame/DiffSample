VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMeasureVBPMU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   BPMUピンのサイトシリアル電圧測定クラス
'
'目的:
'   各パラメータに従いBPMUピンの電圧測定を実行する
'   サイト間はシリアルループする
'
'作成者:
'   0145206097
'
Option Explicit

Implements IMeasure
Implements IAction
Implements IParameterWritable
Implements IMeasureLoopOption

Const ACTION_NAME = "Measure V"
Const WAIT_UNIT = "S"
Const AVG_LOLIMIT = 1

Dim mActionName As CParamString
Dim mPinList As CParamString
Dim mAvg As CParamLong
Dim mWait As CParamStringWithUnit
Dim mGroupID As CParamLong
Dim mLapTime As CParamDouble
Dim mMeasSite As CParamLong

Dim mMeasurePins As Collection
Dim mContainer As CContainer

Private Sub Class_Initialize()
    Set mContainer = New CContainer
    Set mMeasurePins = New Collection
    Set mActionName = CreateCParamString
    With mActionName.AsIParameter
        .Name = ACTION_LABEL
        .AsString = ACTION_NAME
    End With
    Set mGroupID = CreateCParamLong
    mGroupID.AsIParameter.Name = GROUP_ID
    Set mPinList = CreateCParamString
    mPinList.AsIParameter.Name = TEST_PINLIST
    Set mAvg = CreateCParamLong
    With mAvg.AsIParameter
        .Name = MEASURE_AVG
        .LowerLimit = AVG_LOLIMIT
    End With
    Set mWait = CreateCParamStringWithUnit
    With mWait
        .Initialize WAIT_UNIT
        .AsIParameter.Name = MEASURE_WAIT
    End With
    Set mLapTime = CreateCParamDouble
    mLapTime.AsIParameter.Name = EXAMIN_TIMESTAMP
    Set mMeasSite = CreateCParamLong
    With mMeasSite.AsIParameter
        .Name = SITE_INDEX
        .LowerLimit = -1
        .UpperLimit = GetSiteCount
    End With
End Sub

Private Sub IMeasure_Initialize(ByVal measSite As Long, ByVal MeasurePins As Collection, ByVal setContainer As CContainer, ByVal reader As IParameterReader)
    Set mMeasurePins = MeasurePins
    Set mContainer = setContainer
    mGroupID.AsIParameter.Read reader
    mPinList.AsIParameter.Read reader
    mAvg.AsIParameter.Read reader
    mWait.AsIParameter.Read reader
    mMeasSite.AsIParameter.AsLong = measSite
End Sub

Private Sub IMeasure_GetBoardInfo()
    Dim pinRange As Collection
    Dim pinForce As Collection
    With mContainer.DcTest
        Set pinRange = .GetMeasureRangeForBPMU(mPinList.AsIParameter.AsString)
        Set pinForce = .GetForceValueForBPMU(mPinList.AsIParameter.AsString)
    End With
    Dim currPin As CMeasurePin
    For Each currPin In mMeasurePins
        With currPin
            If Not pinRange Is Nothing Then
                .BoardRange = pinRange.Item(.Name)
            Else
                .BoardRange = INVALIDATION_VALUE
            End If
            If Not pinForce Is Nothing Then
                .BoardForce = pinForce.Item(.Name)
            Else
                .BoardForce = INVALIDATION_VALUE
            End If
            .BoardName = mContainer.DcTest.GetBoardNameForBPMU(.Name)
        End With
    Next currPin
End Sub
'#V21-Release
Private Sub IMeasure_WriteResult(ByVal resultKeeper As CResultManager)
    Dim currPin As CMeasurePin
    For Each currPin In mMeasurePins
        With currPin
            If .TestLabel <> NOT_DEFINE Then
                If mMeasSite.AsIParameter.AsLong = ALL_SITE Then
                resultKeeper.Add .TestLabel, .result
                Else
                resultKeeper.AddBySite .TestLabel, .result, mMeasSite.AsIParameter.AsLong
                End If
            End If
        End With
    Next currPin
End Sub

Private Property Get IMeasure_MeasurePins() As Collection
    Set IMeasure_MeasurePins = mMeasurePins
End Property
'#V21-Release
Private Sub IAction_Execute()
    With mContainer
        Do While .WaitTimer.GetLapTime < mWait.AsIParameter.AsDouble
        Loop
        .DcTest.MeasureVSerialForBPMU mPinList.AsIParameter.AsString, mAvg.AsIParameter.AsLong, mMeasSite.AsIParameter.AsLong
        mLapTime.AsIParameter.AsDouble = .WaitTimer.GetLapTime
        .DcTest.LoadResult mMeasurePins, mMeasSite.AsIParameter.AsLong
    End With
End Sub

Private Sub IMeasureLoopOption_Execute()
'内容:
'   LoopOption用Execute関数
'   Wait=0、Average=1で実行する
'
'備考:
'
    With mContainer
        Do While .WaitTimer.GetLapTime < 0
        Loop
        .DcTest.MeasureVSerialForBPMU mPinList.AsIParameter.AsString, 1, mMeasSite.AsIParameter.AsLong
        mLapTime.AsIParameter.AsDouble = .WaitTimer.GetLapTime
        .DcTest.LoadResult mMeasurePins, mMeasSite.AsIParameter.AsLong
    End With
End Sub
Private Property Get IAction_Name() As String
    IAction_Name = MEASURE_CLASS
End Property

Private Property Get IAction_GroupIndex() As Long
    IAction_GroupIndex = mGroupID.AsIParameter.AsLong
End Property

Private Property Get IParameterWritable_ParameterList() As Collection
    Set IParameterWritable_ParameterList = New Collection
    With IParameterWritable_ParameterList
        .Add mPinList.AsIParameter.Name
        .Add mWait.AsIParameter.Name
        .Add mAvg.AsIParameter.Name
    End With
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    mGroupID.AsIParameter.WriteOut writer
    mActionName.AsIParameter.WriteOut writer
    mPinList.AsIParameter.WriteOut writer
    mWait.AsIParameter.WriteOut writer
    mAvg.AsIParameter.WriteOut writer
    mMeasSite.AsIParameter.WriteOut writer
    mLapTime.AsIParameter.WriteOut writer
End Sub

Public Function AsIMeasure() As IMeasure
    Set AsIMeasure = Me
End Function

Public Function AsIAction() As IAction
    Set AsIAction = Me
End Function

Public Function AsIParameterWritable() As IParameterWritable
    Set AsIParameterWritable = Me
End Function

Private Sub Class_Terminate()
    Set mMeasurePins = Nothing
    Set mContainer = Nothing
    Set mActionName = Nothing
    Set mGroupID = Nothing
    Set mPinList = Nothing
    Set mAvg = Nothing
    Set mWait = Nothing
    Set mLapTime = Nothing
    Set mMeasSite = Nothing
End Sub
