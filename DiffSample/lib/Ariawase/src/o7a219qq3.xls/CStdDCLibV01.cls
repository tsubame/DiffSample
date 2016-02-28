VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStdDCLibV01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DC標準ライブラリ（@JOB標準化委員会）のラッパークラス
'
'目的:
'   DC標準ライブラリの各プロシージャを呼び出す
'
'注意:
'   HSD200, ICUL1Gボードは未対応
'
'作成者:
'   0145206097
'
Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1               '1CUL1Gボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>
#Const HSD200_USE = 1               'HSD200ボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

Implements IDcTest

Dim mDcResult As Collection

Private Sub Class_Initialize()
    Set mDcResult = New Collection
End Sub

Private Sub Class_Terminate()
    Set mDcResult = Nothing
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
    Call CheckBoardType(PinList)
    IDcTest_GetForceVLimit = TheDC.Pins(PinList).GetForceVLimit
End Function

Private Function IDcTest_GetForceILimit(ByVal PinList As String) As Double()
    Call CheckBoardType(PinList)
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
    Dim BdName As String
    BdName = TheDC.Pins(PinList).BoardName
    #If ICUL1G_USE <> 0 Then
        If BdName = "dcICUL1G" Then
            Call OutputErrMsg("ICUL1G PPMU is not supported by StdDCLib.")
        End If
    #End If
    IDcTest_GetBoardName = BdName
End Function

Private Function IDcTest_GetBoardNameForBPMU(ByVal PinList As String) As String
    IDcTest_GetBoardNameForBPMU = "dcBPMU"
End Function

Private Sub IDcTest_SetModeFVMI(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
    Call CheckBoardType(PinList)
    Call SetFVMI(PinList, ForceV, ClampI, SiteIndex, Connection)
End Sub

Private Sub IDcTest_SetModeFVMIForBPMU(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
    Call SetFVMI_BPMU(PinList, ForceV, ClampI, SiteIndex, , , Connection)
End Sub

Private Sub IDcTest_SetModeFIMV(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
    Call CheckBoardType(PinList)
    Call SetFIMV(PinList, ForceI, ClampV, SiteIndex, Connection)
End Sub

Private Sub IDcTest_SetModeFIMVForBPMU(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
    Call SetFIMV_BPMU(PinList, ForceI, ClampV, SiteIndex, , , Connection)
End Sub

Private Sub IDcTest_SetModeMV(ByVal PinList As String, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
    Call SetMV(PinList, ClampV, SiteIndex, Connection)
End Sub

Private Sub IDcTest_SetModeMVForBPMU(ByVal PinList As String, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
    Call SetMV_BPMU(PinList, ClampV, SiteIndex, , Connection)
End Sub

Private Sub IDcTest_SetRangeMI(ByVal PinList As String, ByVal ClampI As Double, ByVal SiteIndex As Long)
    Call CheckBoardType(PinList)
    Call ChangeMIRange(PinList, ClampI, SiteIndex)
End Sub

Private Sub IDcTest_SetRangeMIForBPMU(ByVal PinList As String, ByVal ClampI As Double, ByVal SiteIndex As Long)
    Call ChangeMIRange_BPMU(PinList, ClampI, SiteIndex)
End Sub

Private Sub IDcTest_MeasureVConcurrent(ByVal PinList As String, ByVal avgNum As Long)
    Dim pinIndex As Long
    Dim pinArray() As String
    Dim pinsNum As Long
    Dim result() As Double
    Call TheExec.DataManager.DecomposePinList(PinList, pinArray, pinsNum)
    CreateSiteArray result
'{
'厚木JOB自動検証用
    If Flg_JobAutoCheck = 1 Then
        'スナップショット取得
        Call CheckMeasureStatus
    End If
'}
    Call MeasureVMulti(PinList, avgNum)
    Set mDcResult = New Collection
    For pinIndex = 0 To pinsNum - 1
        Call ReadMVMultiResult(pinArray(pinIndex), result)
        mDcResult.Add result, pinArray(pinIndex)
    Next pinIndex
End Sub

Private Sub IDcTest_MeasureIConcurrent(ByVal PinList As String, ByVal avgNum As Long)
    Dim pinIndex As Long
    Dim pinArray() As String
    Dim pinsNum As Long
    Dim result() As Double
    Call TheExec.DataManager.DecomposePinList(PinList, pinArray, pinsNum)
    CreateSiteArray result
'{
'厚木JOB自動検証用
    If Flg_JobAutoCheck = 1 Then
        'スナップショット取得
        Call CheckMeasureStatus
    End If
'}
    Call MeasureIMulti(PinList, avgNum)
    Set mDcResult = New Collection
    For pinIndex = 0 To pinsNum - 1
        Call ReadMIMultiResult(pinArray(pinIndex), result)
        mDcResult.Add result, pinArray(pinIndex)
    Next pinIndex
End Sub

Private Sub IDcTest_MeasureVSerial(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    Dim pinIndex As Long
    Dim pinArray() As String
    Dim pinsNum As Long
    Dim result() As Double
    Call TheExec.DataManager.DecomposePinList(PinList, pinArray, pinsNum)
    CreateSiteArray result
    Set mDcResult = New Collection
'{
'厚木JOB自動検証用
    If Flg_JobAutoCheck = 1 Then
        'スナップショット取得
        Call CheckMeasureStatus
    End If
'}
    For pinIndex = 0 To pinsNum - 1
        Call MeasureV(pinArray(pinIndex), result, avgNum, SiteIndex)
        mDcResult.Add result, pinArray(pinIndex)
    Next pinIndex
End Sub

Private Sub IDcTest_MeasureVSerialForBPMU(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    Dim result() As Double
    CreateSiteArray result
    Set mDcResult = New Collection
'{
'厚木JOB自動検証用
    If Flg_JobAutoCheck = 1 Then
        'スナップショット取得
        Call CheckMeasureStatus
    End If
'}
    Call MeasureV_BPMU(PinList, result, avgNum, SiteIndex)
    mDcResult.Add result, PinList
End Sub

Private Sub IDcTest_MeasureVSerialForGANG(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    Dim result() As Double
    CreateSiteArray result
    Set mDcResult = New Collection
'{
'厚木JOB自動検証用
    If Flg_JobAutoCheck = 1 Then
        'スナップショット取得
        Call CheckMeasureStatus
    End If
'}
    Call MeasureV(PinList, result, avgNum, SiteIndex)
    mDcResult.Add result, PinList
End Sub

Private Sub IDcTest_MeasureISerial(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    Dim pinIndex As Long
    Dim pinArray() As String
    Dim pinsNum As Long
    Dim result() As Double
    Call TheExec.DataManager.DecomposePinList(PinList, pinArray, pinsNum)
    CreateSiteArray result
    Set mDcResult = New Collection
'{
'厚木JOB自動検証用
    If Flg_JobAutoCheck = 1 Then
        'スナップショット取得
        Call CheckMeasureStatus
    End If
'}
    For pinIndex = 0 To pinsNum - 1
        Call MeasureI(pinArray(pinIndex), result, avgNum, SiteIndex)
        mDcResult.Add result, pinArray(pinIndex)
    Next pinIndex
End Sub

Private Sub IDcTest_MeasureISerialForBPMU(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    Dim result() As Double
    CreateSiteArray result
    Set mDcResult = New Collection
'{
'厚木JOB自動検証用
    If Flg_JobAutoCheck = 1 Then
        'スナップショット取得
        Call CheckMeasureStatus
    End If
'}
    Call MeasureI_BPMU(PinList, result, avgNum, SiteIndex)
    mDcResult.Add result, PinList
End Sub

Private Sub IDcTest_MeasureISerialForGANG(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    Dim result() As Double
    CreateSiteArray result
    Set mDcResult = New Collection
'{
'厚木JOB自動検証用
    If Flg_JobAutoCheck = 1 Then
        'スナップショット取得
        Call CheckMeasureStatus
    End If
'}
    Call MeasureI(PinList, result, avgNum, SiteIndex)
    mDcResult.Add result, PinList
End Sub

Private Sub IDcTest_Connect(ByVal PinList As String, ByVal SiteIndex As Long)
    Call ConnectPins(PinList, SiteIndex)
End Sub

Private Sub IDcTest_ConnectForBPMU(ByVal PinList As String, ByVal SiteIndex As Long)
    Call ConnectPins_BPMU(PinList, SiteIndex)
End Sub

Private Sub IDcTest_Disconnect(ByVal PinList As String, ByVal SiteIndex As Long)
    Call DisconnectPins(PinList, SiteIndex)
End Sub

Private Sub IDcTest_DisconnectForBPMU(ByVal PinList As String, ByVal SiteIndex As Long)
    Call SetFVMI_BPMU(PinList, 0, 0.02, SiteIndex, , , False)
    Call DisconnectPins_BPMU(PinList, SiteIndex)
End Sub

'#V21-Release
Private Sub IDcTest_LoadResult(ByVal measurePinSet As Collection, Optional ByVal SiteIndex As Long = ALL_SITE)
    Dim currPin As CMeasurePin
    For Each currPin In measurePinSet
        currPin.SetResult mDcResult(currPin.Name), SiteIndex
    Next currPin
End Sub

Public Function AsIDcTest() As IDcTest
    Set AsIDcTest = Me
End Function

Private Sub CheckBoardType(ByVal pPinList As String)
    Dim BdName As String
    BdName = TheDC.Pins(pPinList).BoardName
    
    #If HSD200_USE <> 0 Then
        If BdName = "dcPPMU" Then
            Call OutputErrMsg("HSD200 PPMU is not supported by StdDCLib.")
        End If
    #End If
    #If ICUL1G_USE <> 0 Then
        If BdName = "dcICUL1G" Then
            Call OutputErrMsg("ICUL1G PPMU is not supported by StdDCLib.")
        End If
    #End If
End Sub


'//////////未実装クラス
Private Sub IDcTest_DisconnectConCurrent(ByVal PinList As String)
'#V21-Release
    Call OutputErrMsg("Can't use this Function as standard function")
End Sub
Private Sub IDcTest_SetModeFVMIConcurrent(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal Connection As Boolean = True)
'#V21-Release
Call OutputErrMsg("Can't use this function as standard function")
End Sub
Private Sub IDcTest_SetModeFIMVConcurrent(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)
'#V21-Release
    Call OutputErrMsg("Can't use this function as standard function")
End Sub
Private Sub IDcTest_SetModeMVConcurrent(ByVal PinList As String, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)
'#V21-Release
    Call OutputErrMsg("Can't use this function as standard function")
End Sub
Private Sub IDcTest_SetRangeMIConcurrent(ByVal PinList As String, ByVal ClampI As Double)
'#V21-Release
    Call OutputErrMsg("Can't use this function as standard function")
End Sub
Private Sub IDcTest_DisconnectForBPMUConCurrent(ByVal PinList As String)
'#V21-Release
    Call OutputErrMsg("Can't use this Function as standard function")
End Sub
Private Sub IDcTest_SetModeFVMIForBPMUConcurrent(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal Connection As Boolean = True)
'#V21-Release
Call OutputErrMsg("Can't use this function as standard function")
End Sub
Private Sub IDcTest_SetModeFIMVForBPMUConcurrent(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)
'#V21-Release
    Call OutputErrMsg("Can't use this function as standard function")
End Sub
Private Sub IDcTest_SetModeMVForBPMUConcurrent(ByVal PinList As String, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)
'#V21-Release
    Call OutputErrMsg("Can't use this function as standard function")
End Sub
Private Sub IDcTest_SetRangeMIForBPMUConcurrent(ByVal PinList As String, ByVal ClampI As Double)
'#V21-Release
    Call OutputErrMsg("Can't use this function as standard function")
End Sub

