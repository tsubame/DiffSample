VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVISConnectDcScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   DC�e�X�g�V�i���I��VIS�N���X��ڑ�
'
'�ړI:
'   �T�FDC�e�X�g�V�i���I��VIS�N���X��ڑ�
'   �U�FMeasure���s���O�ɃX�i�b�v�V���b�g���擾����@�\���
'
'�쐬��:
'   SLSI����
'
'
'Code Checked
'Comment Checked
'

Option Explicit

Implements IDcTest

Dim mDcResult As Collection
Dim mDcSrcObj As CVISVISrcSelector
Dim mSaveSnapshot As Boolean
Dim mSnapshotObjList As Collection

Private Const CS_NAME = "@CVISConnectDcScenario"

'#Pass-OK
Public Function AsIDcTest() As IDcTest
'���e:
'   IDcTest OBJ�擾
'
'�p�����[�^:
'
'�߂�l�F
'   IDcTest �I�u�W�F�N�g
'
'���ӎ���:
'
    Set AsIDcTest = Me

End Function

'#Pass-OK
Public Property Let CanUseSnapshot(ByVal UseSnapshot As Boolean)
'���e�F
'    Measure���s���O�ł̃X�i�b�v�V���b�g�擾�@�\�̐ݒ�
'
'�p�����[�^�F
'    [UseSnapshot]   In  �X�i�b�v�V���b�g�@�\�̗L��
'
'�߂�l�F
'
'���ӎ����F
'    �@�\�L���iUseSnapshot=True�j
'    [UseSnapshot]�̒l�ݒ莞�ɁASnapshotObjList�v���p�e�B��
'    �X�i�b�v�V���b�g�@�\���������I�u�W�F�N�gLIST�����ݒ��
'    �ꍇ�̓G���[�ƂȂ�܂��
'
    If (mSnapshotObjList Is Nothing) And (UseSnapshot = True) Then
        Call OutputErrMsg("SnapshotObj List is Nothing�@" & CS_NAME)
        Exit Property
    Else
        mSaveSnapshot = UseSnapshot
    End If

End Property

'#Pass-Ok
Public Property Get CanUseSnapshot() As Boolean
'���e�F
'    Measure���s���O�ł̃X�i�b�v�V���b�g�擾�@�\�̐ݒ�
'
'�p�����[�^�F
'
'�߂�l�F
'
'���ӎ����F
'    �@�\�L���iUseSnapshot=True�j
'
    CanUseSnapshot = mSaveSnapshot

End Property

'#Pass-Ok
Public Property Set SnapshotObjList(ByVal SnapshotObjList As Collection)
'���e�F
'    ISnapshot���������Ă���OBJ���ۑ�����Ă���R���N�V����
'
'�p�����[�^�F
'    [SnapshotObjList]   In  ISnapshot���������Ă���OBJ�������Ă���R���N�V����
'
'�߂�l�F
'
'���ӎ����F
'    ISnapshot���������Ă��Ȃ�OBJ�������Ă���R���N�V������
'    �ݒ肵���ꍇ�̓G���[�ƂȂ�܂��B
'
    Set mSnapshotObjList = SnapshotObjList

    '�R���N�V������Nothing�łȂ������m�F
    If mSnapshotObjList Is Nothing Then
        Call OutputErrMsg("SnapshotObjList is Nothing " & CS_NAME)
        Exit Property
    End If

    '�R���N�V�����ɋl�ߍ��܂�Ă���̂�ISnapshot���������Ă���OBJ���m�F
    Dim snapObj As ISnapshot
    Set snapObj = Nothing
    
    On Error GoTo SNAPSHOT_OBJ_ERROR
    For Each snapObj In mSnapshotObjList
'        Debug.Print "SnapshotFncObj= " & snapObj.FunctionName
        '�R���N�V�����̒���Nothing�������Ă��Ȃ������m�F
        If snapObj Is Nothing Then
            On Error GoTo 0
            Call OutputErrMsg("Object is Nothing in the SnapshotObjList " & CS_NAME)
            Exit Property
        End If
    Next snapObj
    
    Exit Property
    
SNAPSHOT_OBJ_ERROR:

    Call OutputErrMsg("The Object without the ISnapshot function is in the SnapshotObjList " & CS_NAME)

End Property

'#Pass-Ok
Public Property Set VISrcSelector(ByVal VISrcSelectObj As CVISVISrcSelector)
'���e�F
'    VIS�N���X��ݒ�ς݂�VISource�Z���N�^
'
'�p�����[�^�F
'    [VISrcSelectObj]   In  �ݒ�ς݂�VISource�Z���N�^
'
'�߂�l�F
'
'���ӎ����F
'   VISrcSelectObj�́A�g�p����d���{�[�h��ݒ�ς݂̂��̂�
'   �o�^���܂��B

    If VISrcSelectObj.BoardCount <> 0 Then
        Set mDcSrcObj = VISrcSelectObj
    Else
        Call OutputErrMsg("VIsource selector board list is nothing " & CS_NAME)
    End If

End Property

'#Pass-Ok
Private Sub Class_Initialize()
    
    '���茋�ʏ�����
    Set mDcResult = New Collection
    Set mDcSrcObj = Nothing
        
    '�X�i�b�v�V���b�g�@�\�̏�����
    Set mSnapshotObjList = Nothing
    mSaveSnapshot = False

End Sub

'#Pass-Ok
Private Sub Class_Terminate()
    
    Set mDcResult = Nothing
    Set mDcSrcObj = Nothing
    Set mSnapshotObjList = Nothing

End Sub

'######################## �ȉ� IDcTest�̎��� #############################

Private Function IDcTest_GetMeasureRange(ByVal PinList As String) As Collection
    Set IDcTest_GetMeasureRange = mDcSrcObj.Pins(PinList).GetMeasureRange
End Function

Private Function IDcTest_GetMeasureRangeForBPMU(ByVal PinList As String) As Collection
    Set IDcTest_GetMeasureRangeForBPMU = mDcSrcObj.Pins(PinList, dcBPMU).GetMeasureRange
End Function

Private Function IDcTest_GetForceValue(ByVal PinList As String) As Collection
    Set IDcTest_GetForceValue = mDcSrcObj.Pins(PinList).GetForceValue
End Function

Private Function IDcTest_GetForceValueForBPMU(ByVal PinList As String) As Collection
    Set IDcTest_GetForceValueForBPMU = mDcSrcObj.Pins(PinList, dcBPMU).GetForceValue
End Function

Private Function IDcTest_GetForceVLimit(ByVal PinList As String) As Double()
    IDcTest_GetForceVLimit = mDcSrcObj.Pins(PinList).GetForceVLimit
End Function

Private Function IDcTest_GetForceILimit(ByVal PinList As String) As Double()
    IDcTest_GetForceILimit = mDcSrcObj.Pins(PinList).GetForceILimit
End Function

Private Function IDcTest_GetForceVLimitForBPMU(ByVal PinList As String) As Double()
    IDcTest_GetForceVLimitForBPMU = mDcSrcObj.Pins(PinList, dcBPMU).GetForceVLimit
End Function

Private Function IDcTest_GetForceILimitForBPMU(ByVal PinList As String) As Double()
    IDcTest_GetForceILimitForBPMU = mDcSrcObj.Pins(PinList, dcBPMU).GetForceILimit
End Function

Private Function IDcTest_GetBoardName(ByVal PinList As String) As String
    IDcTest_GetBoardName = mDcSrcObj.Pins(PinList).BoardName
End Function

Private Function IDcTest_GetBoardNameForBPMU(ByVal PinList As String) As String
    IDcTest_GetBoardNameForBPMU = mDcSrcObj.Pins(PinList, dcBPMU).BoardName
End Function

'#Pass-OK
Private Sub IDcTest_SetModeFVMI(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)

    Call mDcSrcObj.Pins(PinList).SetFVMI(ForceV, ClampI, SiteIndex, Connection)

End Sub

'#Pass-Ok
Private Sub IDcTest_SetModeFVMIForBPMU(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)

    Call mDcSrcObj.Pins(PinList, dcBPMU).SetFVMI(ForceV, ClampI, SiteIndex, Connection)
    
End Sub
'#V21-Release
Private Sub IDcTest_SetModeFVMIConcurrent(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal Connection As Boolean = True)

    Call mDcSrcObj.Pins(PinList).SetFVMIMulti(ForceV, ClampI, Connection)

End Sub
'#V21-Release
Private Sub IDcTest_SetModeFVMIForBPMUConcurrent(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal Connection As Boolean = True)

    Call mDcSrcObj.Pins(PinList, dcBPMU).SetFVMIMulti(ForceV, ClampI, Connection)

End Sub
'#Pass-Ok
Private Sub IDcTest_SetModeFIMV(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)

    Call mDcSrcObj.Pins(PinList).SetFIMV(ForceI, ClampV, SiteIndex, Connection)

End Sub

'#Pass-OK
Private Sub IDcTest_SetModeFIMVForBPMU(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)

    Call mDcSrcObj.Pins(PinList, dcBPMU).SetFIMV(ForceI, ClampV, SiteIndex, Connection)

End Sub

'#V21-Release
Private Sub IDcTest_SetModeFIMVConcurrent(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)

    Call mDcSrcObj.Pins(PinList).SetFIMVMulti(ForceI, ClampV, Connection)

End Sub
'#V21-Release
Private Sub IDcTest_SetModeFIMVForBPMUConcurrent(ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)

    Call mDcSrcObj.Pins(PinList, dcBPMU).SetFIMVMulti(ForceI, ClampV, Connection)

End Sub
'#Pass-Ok
Private Sub IDcTest_SetModeMV(ByVal PinList As String, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
    
    Call mDcSrcObj.Pins(PinList).SetMV(ClampV, SiteIndex, Connection)

End Sub

'#Pass-Ok
Private Sub IDcTest_SetModeMVForBPMU(ByVal PinList As String, ByVal ClampV As Double, ByVal SiteIndex As Long, Optional ByVal Connection As Boolean = True)
    
    Call mDcSrcObj.Pins(PinList, dcBPMU).SetMV(ClampV, SiteIndex, Connection)
    
End Sub

'#V21-Release
Private Sub IDcTest_SetModeMVConcurrent(ByVal PinList As String, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)
    
    Call mDcSrcObj.Pins(PinList).SetMVMulti(ClampV, Connection)

End Sub
'#V21-Release
Private Sub IDcTest_SetModeMVForBPMUConcurrent(ByVal PinList As String, ByVal ClampV As Double, Optional ByVal Connection As Boolean = True)
    
    Call mDcSrcObj.Pins(PinList, dcBPMU).SetMVMulti(ClampV, Connection)

End Sub

'#Pass-Ok
Private Sub IDcTest_SetRangeMI(ByVal PinList As String, ByVal ClampI As Double, ByVal SiteIndex As Long)
    
    Call mDcSrcObj.Pins(PinList).ChangeMIRange(ClampI, SiteIndex)

End Sub

'#Pass-Ok
Private Sub IDcTest_SetRangeMIForBPMU(ByVal PinList As String, ByVal ClampI As Double, ByVal SiteIndex As Long)
    
    Call mDcSrcObj.Pins(PinList, dcBPMU).ChangeMIRange(ClampI, SiteIndex)

End Sub

'#V21-Release
Private Sub IDcTest_SetRangeMIConcurrent(ByVal PinList As String, ByVal ClampI As Double)
    
    Call mDcSrcObj.Pins(PinList).ChangeMIRangeMulti(ClampI)

End Sub
'#V21-Release
Private Sub IDcTest_SetRangeMIForBPMUConcurrent(ByVal PinList As String, ByVal ClampI As Double)
    
    Call mDcSrcObj.Pins(PinList, dcBPMU).ChangeMIRangeMulti(ClampI)

End Sub
'#Pass-Ok
Private Sub IDcTest_MeasureVConcurrent(ByVal PinList As String, ByVal avgNum As Long)
    
    Dim pinIndex As Long
    Dim pinArray() As String
    Dim pinsNum As Long
    Dim result() As Double
    
    Call TheExec.DataManager.DecomposePinList(PinList, pinArray, pinsNum)
    ReDim result(TheExec.sites.ExistingCount - 1)
       
    '�X�i�b�v�V���b�g�擾
    If mSaveSnapshot = True Then
        Dim snapFuncObj As ISnapshot
        For Each snapFuncObj In mSnapshotObjList
            Call snapFuncObj.SaveSnapshot
        Next snapFuncObj
    End If
    
    Call mDcSrcObj.Pins(PinList).MeasureVMulti(avgNum)
    Set mDcResult = New Collection
    
    For pinIndex = 0 To pinsNum - 1
        Call mDcSrcObj.Pins(pinArray(pinIndex)).ReadMVMultiResult(result)
        Call mDcResult.Add(result, pinArray(pinIndex))
    Next pinIndex

End Sub

'#Pass-Ok
Private Sub IDcTest_MeasureIConcurrent(ByVal PinList As String, ByVal avgNum As Long)
    
    Dim pinIndex As Long
    Dim pinArray() As String
    Dim pinsNum As Long
    Dim result() As Double
    
    Call TheExec.DataManager.DecomposePinList(PinList, pinArray, pinsNum)
    ReDim result(TheExec.sites.ExistingCount - 1)

    '�X�i�b�v�V���b�g�擾
    If mSaveSnapshot = True Then
        Dim snapFuncObj As ISnapshot
        For Each snapFuncObj In mSnapshotObjList
            Call snapFuncObj.SaveSnapshot
        Next snapFuncObj
    End If
    
    Call mDcSrcObj.Pins(PinList).MeasureIMulti(avgNum)
    Set mDcResult = New Collection
    
    For pinIndex = 0 To pinsNum - 1
        Call mDcSrcObj.Pins(pinArray(pinIndex)).ReadMIMultiResult(result)
        Call mDcResult.Add(result, pinArray(pinIndex))
    Next pinIndex

End Sub

'#Pass-Ok
Private Sub IDcTest_MeasureVSerial(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    
    Dim pinIndex As Long
    Dim pinArray() As String
    Dim pinsNum As Long
    Dim result() As Double
    
    Call TheExec.DataManager.DecomposePinList(PinList, pinArray, pinsNum)
    ReDim result(TheExec.sites.ExistingCount - 1)
    Set mDcResult = New Collection
    
    '�X�i�b�v�V���b�g�擾
    If mSaveSnapshot = True Then
        Dim snapFuncObj As ISnapshot
        For Each snapFuncObj In mSnapshotObjList
            Call snapFuncObj.SaveSnapshot
        Next snapFuncObj
    End If
    
    For pinIndex = 0 To pinsNum - 1
        Call mDcSrcObj.Pins(pinArray(pinIndex)).MeasureV(result, avgNum, SiteIndex)
        Call mDcResult.Add(result, pinArray(pinIndex))
    Next pinIndex

End Sub

'#Pass-Ok
Private Sub IDcTest_MeasureVSerialForBPMU(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    
    Dim result() As Double
    ReDim result(TheExec.sites.ExistingCount - 1)
    Set mDcResult = New Collection
    
    '�X�i�b�v�V���b�g�擾
    If mSaveSnapshot = True Then
        Dim snapFuncObj As ISnapshot
        For Each snapFuncObj In mSnapshotObjList
            Call snapFuncObj.SaveSnapshot
        Next snapFuncObj
    End If
    
    Call mDcSrcObj.Pins(PinList, dcBPMU).MeasureV(result, avgNum, SiteIndex)
    Call mDcResult.Add(result, PinList)

End Sub

Private Sub IDcTest_MeasureVSerialForGANG(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    
    Dim result() As Double
    ReDim result(TheExec.sites.ExistingCount - 1)
    Set mDcResult = New Collection
    
    '�X�i�b�v�V���b�g�擾
    If mSaveSnapshot = True Then
        Dim snapFuncObj As ISnapshot
        For Each snapFuncObj In mSnapshotObjList
            Call snapFuncObj.SaveSnapshot
        Next snapFuncObj
    End If
    
    Call mDcSrcObj.Pins(PinList).MeasureV(result, avgNum, SiteIndex)
    Call mDcResult.Add(result, PinList)

End Sub

'#Pass-Ok
Private Sub IDcTest_MeasureISerial(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    
    Dim pinIndex As Long
    Dim pinArray() As String
    Dim pinsNum As Long
    Dim result() As Double
    
    Call TheExec.DataManager.DecomposePinList(PinList, pinArray, pinsNum)
    ReDim result(TheExec.sites.ExistingCount - 1)
    Set mDcResult = New Collection
    
    '�X�i�b�v�V���b�g�擾
    If mSaveSnapshot = True Then
        Dim snapFuncObj As ISnapshot
        For Each snapFuncObj In mSnapshotObjList
            Call snapFuncObj.SaveSnapshot
        Next snapFuncObj
    End If
    
    For pinIndex = 0 To pinsNum - 1
        Call mDcSrcObj.Pins(pinArray(pinIndex)).MeasureI(result, avgNum, SiteIndex)
        Call mDcResult.Add(result, pinArray(pinIndex))
    Next pinIndex

End Sub

'#Pass-Ok
Private Sub IDcTest_MeasureISerialForBPMU(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    
    Dim result() As Double
    ReDim result(TheExec.sites.ExistingCount - 1)
    Set mDcResult = New Collection
    
    '�X�i�b�v�V���b�g�擾
    If mSaveSnapshot = True Then
        Dim snapFuncObj As ISnapshot
        For Each snapFuncObj In mSnapshotObjList
            Call snapFuncObj.SaveSnapshot
        Next snapFuncObj
    End If
    
    Call mDcSrcObj.Pins(PinList, dcBPMU).MeasureI(result, avgNum, SiteIndex)
    Call mDcResult.Add(result, PinList)

End Sub

Private Sub IDcTest_MeasureISerialForGANG(ByVal PinList As String, ByVal avgNum As Long, ByVal SiteIndex As Long)
    
    Dim result() As Double
    ReDim result(TheExec.sites.ExistingCount - 1)
    Set mDcResult = New Collection
    
    '�X�i�b�v�V���b�g�擾
    If mSaveSnapshot = True Then
        Dim snapFuncObj As ISnapshot
        For Each snapFuncObj In mSnapshotObjList
            Call snapFuncObj.SaveSnapshot
        Next snapFuncObj
    End If
    
    Call mDcSrcObj.Pins(PinList).MeasureI(result, avgNum, SiteIndex)
    Call mDcResult.Add(result, PinList)

End Sub

'#Pass-Ok
Private Sub IDcTest_Connect(ByVal PinList As String, ByVal SiteIndex As Long)
    
    Call mDcSrcObj.Pins(PinList).ConnectPins(SiteIndex)
    
End Sub

'#Pass-Ok
Private Sub IDcTest_ConnectForBPMU(ByVal PinList As String, ByVal SiteIndex As Long)
    
    Call mDcSrcObj.Pins(PinList, dcBPMU).ConnectPins(SiteIndex)

End Sub

'#Pass-Ok
Private Sub IDcTest_Disconnect(ByVal PinList As String, ByVal SiteIndex As Long)
    
    Call mDcSrcObj.Pins(PinList).DisconnectPins(SiteIndex)

End Sub

'#Pass-Ok
Private Sub IDcTest_DisconnectForBPMU(ByVal PinList As String, ByVal SiteIndex As Long)
    
    Call mDcSrcObj.Pins(PinList, dcBPMU).SetFVMI(0, 0.02, SiteIndex, False)
    Call mDcSrcObj.Pins(PinList, dcBPMU).DisconnectPins(SiteIndex)
        
End Sub
'#V21-Release
Private Sub IDcTest_DisconnectConCurrent(ByVal PinList As String)
    
    Call mDcSrcObj.Pins(PinList).DisconnectPinsMulti

End Sub
'#V21-Release
Private Sub IDcTest_DisconnectForBPMUConCurrent(ByVal PinList As String)
    
    Call mDcSrcObj.Pins(PinList, dcBPMU).SetFVMIMulti(0, 0.02, False)
    Call mDcSrcObj.Pins(PinList, dcBPMU).DisconnectPinsMulti

End Sub
'#Pass-Ok
'#V21-Release
Private Sub IDcTest_LoadResult(ByVal measurePinSet As Collection, Optional ByVal SiteIndex As Long = ALL_SITE)
    Dim currPin As CMeasurePin
    For Each currPin In measurePinSet
        currPin.SetResult mDcResult(currPin.Name), SiteIndex
    Next currPin
End Sub