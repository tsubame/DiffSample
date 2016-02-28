VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDisconnectBPMUConcurrent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   BPMUピンのDisconnectクラス
'
'目的:
'   ピンとBPMUリソースの切り離しを実行する(Site同時)
'
'作成者:
'    0020207264
'
Option Explicit

Implements IDisconnect
Implements IAction
Implements IParameterWritable

Const ACTION_NAME = "Disconnect"

Dim mActionName As CParamString
Dim mGroupID As CParamLong
Dim mPinList As CParamString
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
    Set mMeasSite = CreateCParamLong
    With mMeasSite.AsIParameter
        .Name = SITE_INDEX
        .AsLong = ALL_SITE
    End With
End Sub

Private Sub IDisconnect_Initialize(ByVal measSite As Long, ByVal setContainer As CContainer, ByVal reader As IParameterReader)
    Set mContainer = setContainer
    mGroupID.AsIParameter.Read reader
    mPinList.AsIParameter.Read reader
    
End Sub

Private Sub IAction_Execute()
    With mContainer
        .DcTest.DisconnectForBPMUConcurrent mPinList.AsIParameter.AsString
    End With
End Sub

Private Property Get IAction_Name() As String
    IAction_Name = DISCONNECT_CLASS
End Property

Private Property Get IAction_GroupIndex() As Long
    IAction_GroupIndex = mGroupID.AsIParameter.AsLong
End Property

Private Property Get IParameterWritable_ParameterList() As Collection
    Set IParameterWritable_ParameterList = New Collection
    With IParameterWritable_ParameterList
        .Add mPinList.AsIParameter.Name
    End With
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    mGroupID.AsIParameter.WriteOut writer
    mActionName.AsIParameter.WriteOut writer
    mPinList.AsIParameter.WriteOut writer
    mMeasSite.AsIParameter.WriteOut writer
End Sub

Public Function AsIDisconnect() As IDisconnect
    Set AsIDisconnect = Me
End Function

Public Function AsIAction() As IAction
    Set AsIAction = Me
End Function

Public Function AsIParameterWritable() As IParameterWritable
    Set AsIParameterWritable = Me
End Function

Private Sub Class_Terminate()
    Set mContainer = Nothing
    Set mGroupID = Nothing
    Set mActionName = Nothing
    Set mPinList = Nothing
    Set mMeasSite = Nothing
End Sub


