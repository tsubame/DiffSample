VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetFVBPMU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   BPMUピンのFVMIモード設定クラス
'
'目的:
'   各パラメータに従いBPMUピンのFVMIモード設定を実行する
'
'作成者:
'   0145206097
'
Option Explicit

Implements ISetMode
Implements IAction
Implements IParameterWritable

Const ACTION_NAME = "Set FV Mode"
Const FORCE_UNIT = "V"
Const MEASURE_UNIT = "A"

Dim mGroupID As CParamLong
Dim mActionName As CParamString
Dim mPinList As CParamString
Dim mRange As CParamStringWithUnit
Dim mForce As CParamStringWithUnit
Dim mForceSite() As Double
Dim mOperate As CParamString
Dim mMeasSite As CParamLong

Dim mContainer As CContainer

Private Sub Class_Initialize()
    Set mContainer = New CContainer
    Set mActionName = CreateCParamString
    With mActionName.AsIParameter
        .Name = ACTION_LABEL
        .AsString = ACTION_NAME
    End With
    Set mGroupID = CreateCParamLong
    mGroupID.AsIParameter.Name = GROUP_ID
    Set mPinList = CreateCParamString
    mPinList.AsIParameter.Name = TEST_PINLIST
    Set mRange = CreateCParamStringWithUnit
    With mRange
        .Initialize MEASURE_UNIT
        .AsIParameter.Name = SET_RANGE
    End With
    Set mForce = CreateCParamStringWithUnit
    mForce.AsIParameter.Name = SET_FORCE
    Set mOperate = CreateCParamString
    mOperate.AsIParameter.Name = OPERATE_FORCE
    Set mMeasSite = CreateCParamLong
    With mMeasSite.AsIParameter
        .Name = SITE_INDEX
        .LowerLimit = -1
        .UpperLimit = GetSiteCount
    End With
    CreateSiteArray mForceSite
End Sub

Private Sub ISetMode_Initialize(ByVal measSite As Long, ByVal setContainer As CContainer, ByVal reader As IParameterReader)
    Set mContainer = setContainer
    mGroupID.AsIParameter.Read reader
    mPinList.AsIParameter.Read reader
    mOperate.AsIParameter.Read reader
    mRange.AsIParameter.Read reader
    Dim forceMode As CParamString
    Set forceMode = CreateCParamString
    With forceMode.AsIParameter
        .Name = SET_FORCE
        .Read reader
    End With
    With mForce
        If mOperate.AsIParameter.AsString <> "" Then
            .Initialize ""
            .AsIParameter.AsDouble = forceMode.AsIParameter.AsString
        Else
            .Initialize FORCE_UNIT
            Dim forceLimit() As Double
            forceLimit = mContainer.DcTest.GetForceVLimitForBPMU(mPinList.AsIParameter.AsString)
            With .AsIParameter
                .LowerLimit = forceLimit(0)
                .UpperLimit = forceLimit(1)
                .AsString = forceMode.AsIParameter.AsString
            End With
        End If
        Dim SiteIndex As Long
        For SiteIndex = 0 To UBound(mForceSite)
            mForceSite(SiteIndex) = mForce.AsIParameter.AsDouble
        Next SiteIndex
    End With
    mMeasSite.AsIParameter.AsLong = measSite
End Sub
'#V21-Release
Private Sub IAction_Execute()
    Dim convValue() As Double
    CalculateTempValue convValue, mForceSite, mOperate.AsIParameter.AsString, mContainer, mMeasSite.AsIParameter.AsLong
    With mContainer
        .DcTest.SetModeFVMIForBPMU mPinList.AsIParameter.AsString, convValue, mRange.AsIParameter.AsDouble, mMeasSite.AsIParameter.AsLong
        .WaitTimer.Start
    End With
End Sub

Private Property Get IAction_Name() As String
    IAction_Name = SETMODE_CLASS
End Property

Private Property Get IAction_GroupIndex() As Long
    IAction_GroupIndex = mGroupID.AsIParameter.AsLong
End Property

Private Property Get IParameterWritable_ParameterList() As Collection
    Set IParameterWritable_ParameterList = New Collection
    With IParameterWritable_ParameterList
        .Add mPinList.AsIParameter.Name
        .Add mForce.AsIParameter.Name
        .Add mRange.AsIParameter.Name
        .Add mOperate.AsIParameter.Name
    End With
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    mGroupID.AsIParameter.WriteOut writer
    mActionName.AsIParameter.WriteOut writer
    mPinList.AsIParameter.WriteOut writer
    mForce.AsIParameter.WriteOut writer
    mRange.AsIParameter.WriteOut writer
    mOperate.AsIParameter.WriteOut writer
    mMeasSite.AsIParameter.WriteOut writer
End Sub

Public Function AsISetMode() As ISetMode
    Set AsISetMode = Me
End Function

Public Function AsIAction() As IAction
    Set AsIAction = Me
End Function

Public Function AsIParameterWritable() As IParameterWritable
    Set AsIParameterWritable = Me
End Function

Private Sub Class_Terminate()
    Set mContainer = Nothing
    Set mActionName = Nothing
    Set mGroupID = Nothing
    Set mPinList = Nothing
    Set mRange = Nothing
    Set mForce = Nothing
    Set mOperate = Nothing
    Set mMeasSite = Nothing
End Sub
