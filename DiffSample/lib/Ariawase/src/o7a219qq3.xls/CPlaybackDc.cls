VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlaybackDc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DC再生データ出力用クラス
'
'目的:
'   DC Playback Dataシート上に定義されたデータを出力する
'
'作成者:
'   0145206097
'
Option Explicit

Implements IDcTest

Dim mDcReplayData As Collection
Dim mPlaybackReader As IParameterReader

Private Sub Class_Initialize()
    Set mDcReplayData = New Collection
End Sub

Public Sub Initialize(ByVal reader As IParameterReader)
    Set mPlaybackReader = reader
    Set mDcReplayData = dcReplayData(mPlaybackReader)
End Sub

Private Sub Class_Terminate()
    Set mPlaybackReader = Nothing
    Set mDcReplayData = Nothing
End Sub

Private Function IDcTest_GetMeasureRange(ByVal PinList As String) As Collection
    Set IDcTest_GetMeasureRange = Nothing
End Function

Private Function IDcTest_GetMeasureRangeForBPMU(ByVal PinList As String) As Collection
    Set IDcTest_GetMeasureRangeForBPMU = Nothing
End Function

Private Function IDcTest_GetForceValue(ByVal PinList As String) As Collection
    Set IDcTest_GetForceValue = Nothing
End Function

Private Function IDcTest_GetForceValueForBPMU(ByVal PinList As String) As Collection
    Set IDcTest_GetForceValueForBPMU = Nothing
End Function

Private Function IDcTest_GetForceVLimit(ByVal PinList As String) As Double()
    IDcTest_GetForceVLimit = TheDC.Pins(PinList).GetForceVLimit
End Function

Private Function IDcTest_GetForceILimit(ByVal PinList As String) As Double()
    IDcTest_GetForceILimit = TheDC.Pins(PinList).GetForceILimit
End Function

Private Function IDcTest_GetForceVLimitForBPMU(ByVal PinList As String) As Double()
    Dim forceValue(1) As Double
    forceValue(0) = -24
    forceValue(1) = 24
    IDcTest_GetForceVLimitForBPMU = forceValue
End Function

Private Function IDcTest_GetForceILimitForBPMU(ByVal PinList As String) As Double()
    Dim forceValue(1) As Double
    forceValue(0) = -0.2
    forceValue(1) = 0.2
    IDcTest_GetForceILimitForBPMU = forceValue
End Function

Private Function IDcTest_GetBoardName(ByVal PinList As String) As String
    IDcTest_GetBoardName = TheDC.Pins(PinList).BoardName
End Function

Private Function IDcTest_GetBoardNameForBPMU(ByVal PinList As String) As String
    IDcTest_GetBoardNameForBPMU = "dcBPMU"
End Function

Private Sub IDcTest_SetModeFVMI(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
End Sub

Private Sub IDcTest_SetModeFVMIForBPMU(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
End Sub
'#V21-Release
Private Sub IDcTest_SetModeFVMIConcurrent(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal Connection As Boolean = True)
End Sub
'#V21-Release
Private Sub IDcTest_SetModeFVMIForBPMUConcurrent(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal Connection As Boolean = True)
End Sub
Private Sub IDcTest_SetModeFIMV(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
    If TheDC.Pins(PinList).BoardName = "dcDPS" Then Call OutputErrMsg(PinList & " is Invalid Pin List at PlaybackDc.SetModeFIMV")
End Sub

Private Sub IDcTest_SetModeFIMVForBPMU(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
End Sub
'#V21-Release
Private Sub IDcTest_SetModeFIMVConcurrent(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)
    If TheDC.Pins(PinList).BoardName = "dcDPS" Then Call OutputErrMsg(PinList & " is Invalid Pin List at PlaybackDc.SetModeFIMV")
End Sub
'#V21-Release
Private Sub IDcTest_SetModeFIMVForBPMUConcurrent(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)
    If TheDC.Pins(PinList).BoardName = "dcDPS" Then Call OutputErrMsg(PinList & " is Invalid Pin List at PlaybackDc.SetModeFIMV")
End Sub
Private Sub IDcTest_SetModeMV(ByVal PinList As String, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
    If TheDC.Pins(PinList).BoardName = "dcDPS" Then Call OutputErrMsg(PinList & " is Invalid Pin List at PlaybackDc.SetModeMV")
End Sub

Private Sub IDcTest_SetModeMVForBPMU(ByVal PinList As String, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
End Sub
'#V21-Release
Private Sub IDcTest_SetModeMVConcurrent(ByVal PinList As String, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)
    If TheDC.Pins(PinList).BoardName = "dcDPS" Then Call OutputErrMsg(PinList & " is Invalid Pin List at PlaybackDc.SetModeMV")
End Sub
'#V21-Release
Private Sub IDcTest_SetModeMVForBPMUConcurrent(ByVal PinList As String, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)
End Sub
Private Sub IDcTest_SetRangeMI(ByVal PinList As String, ByVal ClampI As Double, ByVal SiteIndex As Long)
End Sub

Private Sub IDcTest_SetRangeMIForBPMU(ByVal PinList As String, ByVal ClampI As Double, ByVal SiteIndex As Long)
End Sub
'#V21-Release
Private Sub IDcTest_SetRangeMIConcurrent(ByVal PinList As String, ByVal ClampI As Double)
End Sub
'#V21-Release
Private Sub IDcTest_SetRangeMIForBPMUConcurrent(ByVal PinList As String, ByVal ClampI As Double)
End Sub
Private Sub IDcTest_MeasureVConcurrent(ByVal PinList As String, ByVal avgNum As Long)
    If IsGangPins(PinList) Then Call OutputErrMsg(PinList & " Including Ganged Pins. (at PlaybackDc.MeasureVConcurrent)")
    If TheDC.Pins(PinList).BoardName = "dcDPS" Then Call OutputErrMsg(PinList & " is Invalid Pin List at PlaybackDc.MeasureVConcurrent")
End Sub

Private Sub IDcTest_MeasureIConcurrent(ByVal PinList As String, ByVal avgNum As Long)
    If IsGangPins(PinList) Then Call OutputErrMsg(PinList & " Including Ganged Pins. (at PlaybackDc.MeasureIConcurrent)")
End Sub

Private Sub IDcTest_MeasureVSerial(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    If TheDC.Pins(PinList).BoardName = "dcDPS" Then Call OutputErrMsg(PinList & " is Invalid Pin List at PlaybackDc.MeasureVSerial")
End Sub

Private Sub IDcTest_MeasureVSerialForBPMU(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Private Sub IDcTest_MeasureVSerialForGANG(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Private Sub IDcTest_MeasureISerial(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Private Sub IDcTest_MeasureISerialForBPMU(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Private Sub IDcTest_MeasureISerialForGANG(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
End Sub

Private Sub IDcTest_Connect(ByVal PinList As String, ByVal SiteIndex As Long)
End Sub

Private Sub IDcTest_ConnectForBPMU(ByVal PinList As String, ByVal SiteIndex As Long)
End Sub

Private Sub IDcTest_Disconnect(ByVal PinList As String, ByVal SiteIndex As Long)
End Sub

Private Sub IDcTest_DisconnectForBPMU(ByVal PinList As String, ByVal SiteIndex As Long)
End Sub
'#V21-Release
Private Sub IDcTest_DisconnectConCurrent(ByVal PinList As String)
End Sub
'#V21-Release
Private Sub IDcTest_DisconnectForBPMUConCurrent(ByVal PinList As String)
End Sub
'#V21-Release
Private Sub IDcTest_LoadResult(ByVal measurePinSet As Collection, Optional ByVal SiteIndex As Long = ALL_SITE)
    If mPlaybackReader.ReadAsBoolean(DATA_CHANGED) Then Set mDcReplayData = dcReplayData(mPlaybackReader)
    Dim currPin As CMeasurePin
    For Each currPin In measurePinSet
        With currPin
            If Not .TestLabel = NOT_DEFINE Then
                .SetResult mDcReplayData.Item(.TestLabel), SiteIndex
            End If
        End With
    Next currPin
End Sub

Public Function AsIDcTest() As IDcTest
    Set AsIDcTest = Me
End Function

Private Function dcReplayData(ByVal reader As IParameterReader) As Collection
    Set dcReplayData = New Collection
    On Error GoTo ErrHandler
    Dim mPlaybackData As Collection
    Set mPlaybackData = New Collection
    '### DC再生データシートリーダーの作成 #################
    Dim fileStream As IFileStream
    Set fileStream = reader
    fileStream.Rewind
    '### パラメータオブジェクトの作成 #####################
    Dim ParamLabel As CParamName
    Set ParamLabel = CreateCParamName
    ParamLabel.AsIParameter.Name = PB_LABEL
    Dim paramData As CParamString
    Set paramData = CreateCParamString
    paramData.AsIParameter.Name = PB_REF_DATA
    Dim paramDelta As CParamString
    Set paramDelta = CreateCParamString
    paramDelta.AsIParameter.Name = PB_DELTA_DATA
    '### データテーブル作成実行 ###########################
    Do While Not fileStream.IsEOR
        ParamLabel.AsIParameter.Read reader
        Dim MainUnit As String
        Dim SubUnit As String
        Dim SubValue As Double
        With paramData.AsIParameter
            .Read reader
            SplitUnitValue .AsString, MainUnit, SubUnit, SubValue
        End With
        Dim paramRef As CParamStringWithUnit
        Set paramRef = CreateCParamStringWithUnit
        With paramRef
            .Initialize MainUnit
            With .AsIParameter
                .Name = PB_REF_DATA
                .Read reader
            End With
        End With
        Dim deltaData As New Collection
        With paramDelta.AsIParameter
            .Read reader
            Set deltaData = DecomposeStringList(.AsString)
        End With
        Dim dataArray() As Double
        ReDim dataArray(deltaData.Count - 1)
        Dim dataIndex As Long
        For dataIndex = 0 To UBound(dataArray)
            If deltaData.Item(dataIndex + 1) = NOT_DEFINE Then
                dataArray(dataIndex) = paramRef.AsIParameter.AsDouble
            Else
                Dim paramPData As CParamStringWithUnit
                Set paramPData = CreateCParamStringWithUnit
                With paramPData
                    .Initialize MainUnit
                    With .AsIParameter
                        .AsString = deltaData.Item(dataIndex + 1)
                        dataArray(dataIndex) = paramRef.AsIParameter.AsDouble + .AsDouble
                    End With
                End With
            End If
        Next dataIndex
        dcReplayData.Add dataArray, ParamLabel.AsIParameter.AsString
        fileStream.MoveNext
    Loop
    Exit Function
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CPlaybackDc.DcReplayData", "Can Not Load The Playback Data !"
End Function

