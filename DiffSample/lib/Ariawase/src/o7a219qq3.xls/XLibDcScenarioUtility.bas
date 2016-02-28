Attribute VB_Name = "XLibDcScenarioUtility"
'�T�v:
'   DC�e�X�g�V�i���I�V�[�g�̂��߂̃��C�u�����Q
'
'�ړI:
'
'�쐬��:
'   SLSI��J
'   2013/10/16 H.Arikawa Ver:1.1 Eee-Job V2.14���ύX�_����ꍞ��

Option Explicit

#Const HSD200_USE = 1               'HSD200�{�[�h�̎g�p�@  0�F���g�p�A0�ȊO�F�g�p  <Tester��IP750EX�Ȃ�Default:1�ɂ��Ă���>

Public Enum CONTROL_STATUS
    INITIAL
    TEST_STEP
    TEST_CONTINUE
    TEST_REPEAT
    TEST_RETURN
    TEST_END
End Enum

Public Enum MEASURE_STATUS
    MEAS_INITIAL
    MEAS_STOP
    MEAS_RESTART
    MEAS_EXIT
End Enum

Public Enum status
    TEST_START
    FIRST_ACT
    RUNNING
    END_OF_ACT
    END_OF_TEST
End Enum

Private Const UPPER_VCLUMP_ICUL1G As Double = 6.5
Private Const LOWER_VCLUMP_ICUL1G As Double = -1.5
Private Const UPPER_VLIMIT_ICUL1G As Double = 6#
Private Const LOWER_VLIMIT_ICUL1G As Double = -1#

#If HSD200_USE = 0 Then
Private Const UPPER_VLIMIT_PPMU As Double = 7#
Private Const LOWER_VLIMIT_PPMU As Double = -2#
#Else
Private Const UPPER_VLIMIT_PPMU As Double = 6.5
Private Const LOWER_VLIMIT_PPMU As Double = -1.5
#End If
    
'#V21-Release
'############# �ȉ����[�e�B���e�B�Q ###############################################################
Public Sub CalculateTempValue(ByRef retValue() As Double, ByRef refValue As Variant, ByVal operateKey As String, ByVal container As CContainer, Optional ByVal site As Long = ALL_SITE)
'���e:
'   �w�肳�ꂽ�e���|�����l�Ƃ̌v�Z���ʂ�Ԃ�
'
'[retValue()]   OUT Double�^:       �v�Z��̃f�[�^�z��
'[refValue]     OUT Variant�^:      �v�Z�Ώۂ̃f�[�^�z��
'[operateKey]   IN String�^:        �v�Z���ƃe���|�����ϐ�����\��������
'[Container]    IN CContainer�^:    �e���|�������ʂ��i�[���ꂽ�R���e�i
'[Site]         In Long�^:          �T�C�g�w��@�\�p�@(Default:-1)
'���ӎ���:
'
'
    Dim dataIndex As Long
    Dim keyName As String
    Dim operator As String
    ReDim retValue(UBound(refValue))
    Dim TempValue() As Double
    operator = Left(operateKey, 1)
    keyName = Replace(operateKey, operator, "")
    On Error GoTo ErrHandler
    Select Case site
    Case ALL_SITE:
        Select Case operator
            Case "=":
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex)
                Next dataIndex
                container.AddTempResult keyName, retValue
            Case "+":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex) + TempValue(dataIndex)
                Next dataIndex
            Case "-":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex) - TempValue(dataIndex)
                Next dataIndex
            Case "*":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex) * TempValue(dataIndex)
                Next dataIndex
            Case "/":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex) / TempValue(dataIndex)
                Next dataIndex
            Case Else:
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex)
                Next dataIndex
        End Select
    Case Else:  'Site�w��
        dataIndex = site
        Select Case operator
            Case "=":
                retValue(dataIndex) = refValue(dataIndex)
                container.AddTempResult keyName, retValue
            Case "+":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                retValue(dataIndex) = refValue(dataIndex) + TempValue(dataIndex)
            Case "-":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                retValue(dataIndex) = refValue(dataIndex) - TempValue(dataIndex)
            Case "*":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                retValue(dataIndex) = refValue(dataIndex) * TempValue(dataIndex)
            Case "/":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                retValue(dataIndex) = refValue(dataIndex) / TempValue(dataIndex)
            Case Else:
                retValue(dataIndex) = refValue(dataIndex)
        End Select
    End Select
    Exit Sub
ErrHandler:
    For dataIndex = 0 To UBound(refValue)
        retValue(dataIndex) = refValue(dataIndex)
    Next dataIndex
    Err.Raise 9999, "CalculateTempValue", "Can Not Calculate The Operate [" & operateKey & "] !"
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

'############# �d���ݒ�N���X�̏����� #############################################################
Public Function CreateVISConnector() As IDcTest
'���e�F
'   DC�e�X�g�V�i���I��d���ݒ�Object��ڑ����邽�߂�
'   �ڑ��pObject�̐����Ə����ݒ�
'
'�p�����[�^�F
'
'�߂�l�F
'   IDcTest�����������d���N���X�ڑ��pObject
'
'���ӎ����F
'

    '�d���N���X�ڑ��p�I�u�W�F�N�g����
    Dim VisLibObj As CVISConnectDcScenario
    Set VisLibObj = New CVISConnectDcScenario

    '�d���ݒ�OBJ��ݒ�
    Set VisLibObj.VISrcSelector = TheDC
    
    '##### �X�i�b�v�V���b�g�擾�@�\ON���̐ݒ� #####
    If (IsSnapshotOn = True) And (TheExec.RunMode = runModeDebug) Then 'PMC���s���ɂ͓��삵�Ȃ��悤�ɔ���ǉ�
                                                                                                                    
        '�ݒ�ς݂̃X�i�b�v�V���b�gOBJ���R���N�V�����ɓo�^
        '�����@�\�̃X�i�b�v�V���b�g���擾���邱�Ƃ��l�������݌v�̂���
        '�R���N�V�����Ɏ擾�������X�i�b�v�V���b�g�@�\��ǉ�����`���Ƃ�
        Dim snapFncList As Collection
        Set snapFncList = New Collection
        
        'Snapshot�@�\��ǉ�
        Call snapFncList.Add(TheSnapshot)
        
        '�X�i�b�v�V���b�gOBJ���X�g�ɐݒ�
        Set VisLibObj.SnapshotObjList = snapFncList
                
        '�X�i�b�v�V���b�g�@�\���s�t���O��True�i���s����j�ɐݒ�
        VisLibObj.CanUseSnapshot = True
    
        '�X�i�b�v�V���b�g�擾���[�h�ɐݒ肳�ꂽ���Ƃ�\��
        Call MsgBox("Snapshot Save Mode !!", vbInformation, "XLibDcScenarioUtility")

    '##### �X�i�b�v�V���b�g�擾�@�\OFF���̐ݒ� #####
    Else
        '�X�i�b�v�V���b�g�@�\���s�t���O��False�i���s���Ȃ��j�ɐݒ�
        VisLibObj.CanUseSnapshot = False
    End If
    
    '���ׂĂ̐ݒ肪�I�������d���N���X�ڑ��p�I�u�W�F�N�g��Ԃ�
    Set CreateVISConnector = VisLibObj

End Function

'############# �e�X�g���n���W���[���Q ###########################################################
Public Function GetInstanceName() As String
'���e:
'   �J�����g�̃e�X�g�C���X�^���X����Ԃ�
'
'�p�����[�^:
'
'�߂�l�F
'   �C���X�^���X��
'
'���ӎ���:
'   �e�X�g�J�e�S���擾�p

    GetInstanceName = TheExec.DataManager.InstanceName

End Function

Public Function GetInstansNameAsUCase() As String
'���e:
'   �J�����g�̃e�X�g�C���X�^���X����啶���ŕԂ�
'
'�p�����[�^:
'
'�߂�l�F
'   �C���X�^���X���i�啶���j
'
'���ӎ���:
'   �e�X�g���x���擾�p

    GetInstansNameAsUCase = UCase(TheExec.DataManager.InstanceName)

End Function

Public Function GetSiteCount() As Long
'���e:
'   �T�C�g���̎擾
'
'�߂�l�F
'   �T�C�g������}�C�i�X1���������l
'
'���ӎ���:
'

    GetSiteCount = TheExec.sites.ExistingCount - 1
End Function

Public Function GetTesterNum() As Long
'���e:
'   Sw_Node�p�u���b�N�ϐ��̃��b�s���O
'
'�߂�l�F
'   �e�X�^�ԍ�
'
'���ӎ���:
'

    GetTesterNum = Sw_Node
End Function

Public Sub CreateSiteArray(ByRef retArray() As Double)
'���e:
'   �z����T�C�g�����m�ۂ���
'
'�߂�l�F
'   �T�C�g�����m�ۂ��ꂽ�z��ϐ�
'
'���ӎ���:
'

    ReDim retArray(GetSiteCount)
End Sub

Public Function IsGangPins(ByVal PinList As String) As Boolean
'���e:
'   �w�肳�ꂽ�s�����X�g�ɃM�����O�s�����܂܂�Ă��邩�m�F
'   IsGangPinList�̃��b�p�[
'
'�p�����[�^:
'    [PinList]       In   �m�F���s��PinList
'
'�߂�l:
'   �m�F���ʁi�M�����O�s�����܂܂�Ă���=True�j
'
'���ӎ���:
'
    If TheDC.Pins(PinList).BoardName = "dcICUL1G" Then
        IsGangPins = False
    Else
        IsGangPins = IsGangPinlist(PinList, GetChanType(PinList))
    End If
End Function

Public Function GetGangPinCount(ByVal PinList As String) As Long
'���e:
'   �M�����O�s������Ԃ�
'
'�p�����[�^:
'    [PinList]       In   �m�F���s��PinList
'
'�߂�l:
'   �M�����O�̃s����
'
'���ӎ���:
'
    Dim pinArr() As String
    TheExec.DataManager.DecomposePinList PinList, pinArr, GetGangPinCount
End Function

Public Function ValidateMeasureRange(ByVal mPin As CMeasurePin) As Long
'���e:
'   ���背���W���œK���ǂ����̔�����s��
'
'�p�����[�^:
'    [mPin]       In   ����Ώۂ�Pin�I�u�W�F�N�g
'
'�߂�l:
'   ���茋�ʁiConstant�^�j
'
'���ӎ���:
'
    With mPin
        '### �s���̃����W���肪�s�\�ȏꍇ ###############
        If .TestLabel = NOT_DEFINE Or .BoardRange = INVALIDATION_VALUE Then
            ValidateMeasureRange = DISABEL_TO_VALIDATION
            Exit Function
        End If
        '### ����p�����[�^��0�܂���3�̏ꍇ ###############
        If .JudgeNumber = 0 Or .JudgeNumber = 3 Then
            '### PPMU��MV���[�h�̂��߂̔��胍�W�b�N #######
            '�����W�� HSD100:-2�`7V, HSD200:-1.5V�`6.5V �݂̂Ȃ̂ōœK�l�̔���͂��Ȃ�
            If .BoardName = "dcPPMU" And GetUnit(.Unit) = "V" Then
                If Not isCorrectLimitForPPMU(.UpperLimit, .LowerLimit) Then
                    If .JudgeNumber = 3 Then
                        ValidateMeasureRange = VALIDATE_NG
                    ElseIf .JudgeNumber = 0 Then
                        ValidateMeasureRange = VALIDATE_NG_NO_JUDGE
                    End If
                Else
                    If .JudgeNumber = 3 Then
                        ValidateMeasureRange = VALIDATE_OK
                    ElseIf .JudgeNumber = 0 Then
                        ValidateMeasureRange = VALIDATE_OK_NO_JUDGE
                    End If
                End If
            '### ICLU1G��MV���[�h�̂��߂̔��胍�W�b�N #######
            '�����W�� -1V�`6V �݂̂Ȃ̂ōœK�l�̔���͂��Ȃ�
            ElseIf .BoardName = "dcICUL1G" And GetUnit(.Unit) = "V" Then
                If Not isCorrectLimitForICUL1G(.UpperLimit, .LowerLimit) Then
                    If .JudgeNumber = 3 Then
                        ValidateMeasureRange = VALIDATE_NG
                    ElseIf .JudgeNumber = 0 Then
                        ValidateMeasureRange = VALIDATE_NG_NO_JUDGE
                    End If
                Else
                    If .JudgeNumber = 3 Then
                        ValidateMeasureRange = VALIDATE_OK
                    ElseIf .JudgeNumber = 0 Then
                        ValidateMeasureRange = VALIDATE_OK_NO_JUDGE
                    End If
                End If
            '### ���̑��̔��胍�W�b�N #####################
            '�����W�̍œK�l�܂Ŋ܂߂�������s��
            Else
                Dim maxLimit As Double
                maxLimit = absMaxLimit(.UpperLimit, .LowerLimit)
                If maxLimit >= RoundDownDblData(.BoardRange, DIGIT_NUMBER) Then
                    If .JudgeNumber = 3 Then
                        ValidateMeasureRange = VALIDATE_NG
                    ElseIf .JudgeNumber = 0 Then
                        ValidateMeasureRange = VALIDATE_NG_NO_JUDGE
                    End If
                Else
                    '### �œK�����W�l���ǂ����̔��� #######
                    Dim isOptimalRange As Boolean
                    Select Case .BoardName
                        Case "dcAPMU":
                            isOptimalRange = isOptimalRangeAPMU(.Name, maxLimit, .BoardRange, .Unit)
                        Case "dcPPMU":
                            isOptimalRange = isOptimalRangePPMU(maxLimit, .BoardRange, .Unit)
                        Case "dcBPMU":
                            isOptimalRange = isOptimalRangeBPMU(maxLimit, .BoardRange, .Unit)
                        Case "dcDPS":
                            isOptimalRange = isOptimalRangeDPS(maxLimit, .BoardRange, .Unit)
                        Case "dcHDVIS":
                            isOptimalRange = isOptimalRangeHDVIS(.Name, maxLimit, .BoardRange, .Unit)
                        Case "dcICUL1G":
                            isOptimalRange = isOptimalRangeICUL1G(maxLimit, .BoardRange, .Unit)
                    End Select
                    If isOptimalRange Then
                        If .JudgeNumber = 3 Then
                            ValidateMeasureRange = VALIDATE_OK
                        ElseIf .JudgeNumber = 0 Then
                            ValidateMeasureRange = VALIDATE_OK_NO_JUDGE
                        End If
                    Else
                        If .JudgeNumber = 3 Then
                            ValidateMeasureRange = VALIDATE_WARNING
                        ElseIf .JudgeNumber = 0 Then
                            ValidateMeasureRange = VALIDATE_WARNING_NO_JUDGE
                        End If
                    End If
                End If
            End If
        '### ����p�����[�^��0��3�ȊO�̓m�[�`�F�b�N #######
        Else
            ValidateMeasureRange = NO_JUDGE
        End If

    End With
End Function

Private Function isCorrectLimitForPPMU(ByVal HiLimit As Double, ByVal LoLimit As Double) As Boolean
    Dim roundHLimit As Double
    Dim roundLLimit As Double
    roundHLimit = RoundDownDblData(HiLimit, DIGIT_NUMBER)
    roundLLimit = RoundDownDblData(LoLimit, DIGIT_NUMBER)
    If UPPER_VLIMIT_PPMU > roundHLimit And LOWER_VLIMIT_PPMU < roundLLimit Then
        isCorrectLimitForPPMU = True
    Else
        isCorrectLimitForPPMU = False
    End If
End Function

Private Function isCorrectLimitForICUL1G(ByVal HiLimit As Double, ByVal LoLimit As Double) As Boolean
    Dim roundHLimit As Double
    Dim roundLLimit As Double
    roundHLimit = RoundDownDblData(HiLimit, DIGIT_NUMBER)
    roundLLimit = RoundDownDblData(LoLimit, DIGIT_NUMBER)
    If UPPER_VLIMIT_ICUL1G > roundHLimit And LOWER_VLIMIT_ICUL1G < roundLLimit Then
        isCorrectLimitForICUL1G = True
    Else
        isCorrectLimitForICUL1G = False
    End If
End Function

Private Function absMaxLimit(ByVal HiLimit As Double, ByVal LoLimit As Double) As Double
    Dim roundHLimit As Double
    Dim roundLLimit As Double
    roundHLimit = RoundDownDblData(HiLimit, DIGIT_NUMBER)
    roundLLimit = RoundDownDblData(LoLimit, DIGIT_NUMBER)
    If Abs(roundHLimit) > Abs(roundLLimit) Then
        absMaxLimit = Abs(roundHLimit)
    Else
        absMaxLimit = Abs(roundLLimit)
    End If
End Function

Private Function isOptimalRangeAPMU(ByVal pName As String, ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            If maxLimit < 2 Then
                isOptimalRangeAPMU = CompareDblData(rangeVal, 2, DIGIT_NUMBER)
            ElseIf maxLimit < 5 Then
                isOptimalRangeAPMU = CompareDblData(rangeVal, 5, DIGIT_NUMBER)
            ElseIf maxLimit < 10 Then
                isOptimalRangeAPMU = CompareDblData(rangeVal, 10, DIGIT_NUMBER)
            ElseIf maxLimit < 35 Then
                isOptimalRangeAPMU = CompareDblData(rangeVal, 35, DIGIT_NUMBER)
            End If
        Case "A":
            If IsGangPins(pName) Then
                '�M�����O�s���͏��Warning��Ԃ��d�l�ɕύX '08/04/04
'                isOptimalRangeAPMU = CompareDblData(rangeVal, 0.05 * GetGangPinCount(pName), DIGIT_NUMBER)
                isOptimalRangeAPMU = False
            Else
                If maxLimit < 0.0000002 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.0000002, DIGIT_NUMBER)
                ElseIf maxLimit < 0.000002 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.000002, DIGIT_NUMBER)
                ElseIf maxLimit < 0.00001 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.00001, DIGIT_NUMBER)
                ElseIf maxLimit < 0.00004 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.00004, DIGIT_NUMBER)
                ElseIf maxLimit < 0.0002 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.0002, DIGIT_NUMBER)
                ElseIf maxLimit < 0.001 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.001, DIGIT_NUMBER)
                ElseIf maxLimit < 0.005 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.005, DIGIT_NUMBER)
                ElseIf maxLimit < 0.05 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.05, DIGIT_NUMBER)
                End If
            End If
    End Select
End Function

Private Function isOptimalRangePPMU(ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            'MV���[�h���͏��True��Ԃ��d�l�ɕύX '08/04/24
'            isOptimalRangePPMU = CompareDblData(rangeVal, 2, DIGIT_NUMBER)
            isOptimalRangePPMU = True
        Case "A":
            If maxLimit < 0.0000002 Then
                #If HSD200_USE = 0 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.0000002, DIGIT_NUMBER)
                #Else
                'HSD200�ɂ̓����W200nA�͂Ȃ�
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.000002, DIGIT_NUMBER)
                #End If
            ElseIf maxLimit < 0.000002 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.000002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.00002 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.00002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.0002 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.0002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.002 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.002, DIGIT_NUMBER)
            #If HSD200_USE <> 0 Then
            ElseIf maxLimit < 0.05 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.05, DIGIT_NUMBER)
            #End If
            End If
    End Select
End Function

Private Function isOptimalRangeICUL1G(ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            'MV���[�h���͏��True��Ԃ�
            isOptimalRangeICUL1G = True
        Case "A":
            If maxLimit < 0.00002 Then
                isOptimalRangeICUL1G = CompareDblData(rangeVal, 0.00002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.0002 Then
                isOptimalRangeICUL1G = CompareDblData(rangeVal, 0.0002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.002 Then
                isOptimalRangeICUL1G = CompareDblData(rangeVal, 0.002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.01 Then
                isOptimalRangeICUL1G = CompareDblData(rangeVal, 0.01, DIGIT_NUMBER)
            End If
    End Select
End Function

Private Function isOptimalRangeBPMU(ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            If maxLimit < 2 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 2, DIGIT_NUMBER)
            ElseIf maxLimit < 5 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 5, DIGIT_NUMBER)
            ElseIf maxLimit < 10 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 10, DIGIT_NUMBER)
            ElseIf maxLimit < 24 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 24, DIGIT_NUMBER)
            End If
        Case "A":
            If maxLimit < 0.000002 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.000002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.00002 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.00002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.0002 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.0002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.002 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.02 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.02, DIGIT_NUMBER)
            ElseIf maxLimit < 0.2 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.2, DIGIT_NUMBER)
            End If
    End Select
End Function

Private Function isOptimalRangeDPS(ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            Call OutputErrMsg("DPS is not support FI mode")
        Case "A":
            If maxLimit < 0.00005 Then
                isOptimalRangeDPS = CompareDblData(rangeVal, 0.00005, DIGIT_NUMBER)
            ElseIf maxLimit < 0.0005 Then
                isOptimalRangeDPS = CompareDblData(rangeVal, 0.0005, DIGIT_NUMBER)
            ElseIf maxLimit < 0.01 Then
                isOptimalRangeDPS = CompareDblData(rangeVal, 0.01, DIGIT_NUMBER)
            ElseIf maxLimit < 0.1 Then
                isOptimalRangeDPS = CompareDblData(rangeVal, 0.1, DIGIT_NUMBER)
            ElseIf maxLimit < 1 Then
                isOptimalRangeDPS = CompareDblData(rangeVal, 1, DIGIT_NUMBER)
            End If
    End Select
End Function

Private Function isOptimalRangeHDVIS(ByVal pName As String, ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            isOptimalRangeHDVIS = CompareDblData(rangeVal, 10, DIGIT_NUMBER)
        Case "A":
            Dim pinCount As Long
            pinCount = GetGangPinCount(pName)
            If maxLimit < RoundDownDblData(0.000005 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.000005 * pinCount, DIGIT_NUMBER)
            ElseIf maxLimit < RoundDownDblData(0.00005 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.00005 * pinCount, DIGIT_NUMBER)
            ElseIf maxLimit < RoundDownDblData(0.0005 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.0005 * pinCount, DIGIT_NUMBER)
            ElseIf maxLimit < RoundDownDblData(0.005 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.005 * pinCount, DIGIT_NUMBER)
            ElseIf maxLimit < RoundDownDblData(0.05 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.05 * pinCount, DIGIT_NUMBER)
            ElseIf maxLimit < RoundDownDblData(0.2 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.2 * pinCount, DIGIT_NUMBER)
            End If
    End Select
End Function

'############# �V�[�g�p�����[�^���͎x���p���[�e�B���e�B�֐��Q #####################################
Public Sub CreateActionParameterList(ByVal selCell As Range)
'���e:
'   DC�V�i���I���[�N�V�[�g�ɂ�������͎x���}�N���֐��@
'   ���[�U�[�̂��߂̃f�[�^���X�g�쐬
'
'�p�����[�^:
'   [selCell]      In   �ΏۃZ���I�u�W�F�N�g
'
'���ӎ���:
'
    If Not IsJobValid Then Exit Sub
    '### �f�[�^���x���ɉ������f�[�^���X�g�̍쐬 ###########
    Dim dataList As Collection
    Set dataList = Nothing
    On Error GoTo LIST_ERROR
    With selCell.parent
    '### �J�e�S���p�����[�^�̃��X�g�쐬 ###################
    If IsCategoryHeader(selCell) Then
        '### �����t���O���X�g�쐬 #########################
        If selCell.Column = .Range(EXAMIN_FLAG).Column Then
            Set dataList = examinFlagList
        '### �������[�h���X�g�쐬 #########################
        ElseIf selCell.Column = .Range(EXAMIN_MODE).Column Then
            If targetCell(selCell, EXAMIN_FLAG).Value Then
                Set dataList = examinModeList
            End If
        End If
    End If
    '### �s���O���[�v�p�����[�^�̃��X�g�쐬 ###############
    If IsGroupHeader(selCell) Then
        '### �s�����X�g�̎擾 #############################
        Dim PinList As String
        PinList = createPinList(selCell, False)
        '### �s�����X�g�������̏ꍇ�̓��X�g�쐬�͂��Ȃ� ###
        If PinList = NOT_DEFINE Then GoTo LIST_ERROR
        Dim boardType As String
        boardType = targetCell(selCell, TEST_PINTYPE).Value
        '### �A�N�V�������X�g�쐬 #########################
        If selCell.Column = .Range(TEST_ACTION).Column Then
            Set dataList = actionList
        '### �s���^�C�v�iI/O�s���̂݁j���X�g�쐬 ##########
        ElseIf selCell.Column = .Range(TEST_PINTYPE).Column Then
            Set dataList = ioPinTypeList(PinList)
        End If
        '### �ȉ��̃p�����[�^�͗L���ł���΃��X�g�쐬���� #
        Dim ParamList As Collection
        Set ParamList = actionParamList(selCell)
        '### ���胂�[�h���X�g�쐬 #########################
        If selCell.Column = .Range(SET_MODE).Column And _
               IsEnableParameter(ParamList, SET_MODE) Then
            Set dataList = modeList(PinList)
        '### ���背���W���X�g�쐬 #########################
        ElseIf selCell.Column = .Range(SET_RANGE).Column And _
               IsEnableParameter(ParamList, SET_RANGE) Then
            Dim measureMode As String
            measureMode = targetCell(selCell, SET_MODE).Value
            Set dataList = enableBoardRange(PinList, boardType, measureMode)
        '### �T�C�g���胂�[�h���X�g�쐬 ###################
        ElseIf selCell.Column = .Range(MEASURE_SITE).Column And _
               IsEnableParameter(ParamList, MEASURE_SITE) Then
            Set dataList = measureSiteList(PinList, boardType)
        End If
    '### �|�X�g�A�N�V�����p�p�����[�^�̃��X�g�쐬 #########
    ElseIf IsGroupFooter(selCell) Then
        If selCell.Column = .Range(TEST_ACTION).Column Then
            Set dataList = postActionList
        End If
    End If
    End With
LIST_ERROR:
    '### �f�[�^���X�g�����X�g�{�b�N�X�ɐݒ� ###############
    CreateListBox selCell, dataList
End Sub

Public Sub ValidateActionParameter(ByVal chCell As Range)
'���e:
'   DC�V�i���I���[�N�V�[�g�ɂ�������͎x���}�N���֐��A
'   �L���ȃp�����[�^���x���̏����`�F�b�N�y�уf�[�^�`�F�b�N
'
'�p�����[�^:
'   [chCell]      In   �ΏۃZ���I�u�W�F�N�g
'
'���ӎ���:
'
    If Not IsJobValid Then Exit Sub
    On Error GoTo DATA_ERROR
    If IsEmpty(chCell) Then enableCell chCell
    '### �J�e�S���p�����[�^�̃p�����[�^�`�F�b�N ###########
    If IsCategoryHeader(chCell) Then
        '���݂̓J�e�S���p�����[�^�̃`�F�b�N�͓��ɂȂ�
    End If
    '### �s���O���[�v�p�����[�^�̃`�F�b�N #################
    If IsGroupHeader(chCell) Then
        '### �s���O���[�v�̃`�F�b�N #######################
        Dim PinList As String
        PinList = createPinList(chCell, True)
        '### �s���O���[�v�擪�s���̃{�[�h���X�V ###########
        updateBoardName chCell
        '�s�����X�g�������̏ꍇ�̓`�F�b�N�͂��Ȃ�
        If PinList = NOT_DEFINE Then Exit Sub
        '### �A�N�V�����p�����[�^���X�g�̎擾 #############
        Dim ParamList As Collection
        Set ParamList = actionParamList(chCell)
        '### �A�N�V�����O���[�v�t�H�[�}�b�g�̐��` #########
        actionParamFormatter targetCell(chCell, TEST_ACTION), ParamList
        '�����ȃA�N�V�����̏ꍇ�͈ȉ��̃p�����[�^�`�F�b�N�͂��Ȃ�
        If ParamList.Count = 0 Then Exit Sub
        '### ���胂�[�h�p�����[�^�`�F�b�N #################
        Dim dataList As Collection
        If IsEnableParameter(ParamList, SET_MODE) Then
            Set dataList = modeList(PinList)
            verifyListParamData targetCell(chCell, SET_MODE), dataList
        End If
        '### �T�C�g���[�h�p�����[�^�`�F�b�N ###############
        If IsEnableParameter(ParamList, MEASURE_SITE) Then
            Dim boardType As String
            boardType = targetCell(chCell, TEST_PINTYPE).Value
            Set dataList = measureSiteList(PinList, boardType)
            verifyListParamData targetCell(chCell, MEASURE_SITE), dataList
        End If
        '### ���背���W�p�����[�^�`�F�b�N #################
        If IsEnableParameter(ParamList, SET_RANGE) Then
            verifyRangeData targetCell(chCell, SET_RANGE), PinList
        End If
        '### ����p�����[�^�`�F�b�N #######################
        If IsEnableParameter(ParamList, SET_FORCE) Then
            verifyForceValue targetCell(chCell, SET_FORCE), PinList
        End If
        '### �E�F�C�g�p�����[�^�`�F�b�N ###################
        If IsEnableParameter(ParamList, MEASURE_WAIT) Then
            verifyWaitValue targetCell(chCell, MEASURE_WAIT)
        End If
        '### ���ω񐔃p�����[�^�`�F�b�N ###################
        If IsEnableParameter(ParamList, MEASURE_AVG) Then
            verifyAverageValue targetCell(chCell, MEASURE_AVG)
        End If
    Else
        '### �s���O���[�v�̃`�F�b�N #######################
        If chCell.Column = chCell.parent.Range(TEST_PINS).Column Then
            createPinList chCell, True
        End If
    End If
    '### �|�X�g�A�N�V�����p�p�����[�^�`�F�b�N #############
    If IsGroupFooter(chCell) Then
        enableCell targetCell(chCell, TEST_ACTION)
        verifyListParamData targetCell(chCell, TEST_ACTION), postActionList
    End If
    Exit Sub
DATA_ERROR:
End Sub

Private Sub updateBoardName(ByVal chCell As Range)
    '### �O���[�v�擪�s���̃{�[�h���X�V ###################
    '�O������F�ΏۃZ�����s���O���[�v�̐擪�ł��邱��
    '08/05/12 OK
    '### �{�[�h���̎擾 ###################################
    Dim boardType As String
    boardType = queryBoardName(targetCell(chCell, TEST_PINS).Value)
    '### ���݂̐ݒ�l�擾 #################################
    Dim typeCell As Range
    Set typeCell = targetCell(chCell, TEST_PINTYPE)
    '### �{�[�h���̐ݒ�X�V ###############################
    If boardType <> typeCell.Value Then
        If boardType <> "PPMU" Or typeCell.Value <> "BPMU" Then
            Application.EnableEvents = False
            typeCell.Value = boardType
            Application.EnableEvents = True
        End If
    End If
End Sub

Private Sub verifyRangeData(ByVal chCell As Range, ByVal PinList As String)
    '### ���͑��背���W�����X�g�ɑ��݂��邩�ǂ����̔��� ###
    '�O������F�ΏۃZ�����s���O���[�v�̐擪�ł��邱��
    '�p�����[�^�������E�X�y�b�N�𖞂����Ȃ��ꍇ�A�p�����[�^�N���X�ŃG���[���N����
    '08/05/09 OK
    On Error GoTo DATA_ERROR
    enableCell chCell
    Dim boardType As String
    boardType = targetCell(chCell, TEST_PINTYPE).Value
    Dim measureMode As String
    measureMode = targetCell(chCell, SET_MODE).Value
    '### �����W���X�g�̎擾 ###############################
    Dim rangeList As Collection
    Set rangeList = enableBoardRange(PinList, boardType, measureMode)
    '### ���X�g��None���܂܂�Ă���ꍇ�͑��̓��͒l��NG ###
    If IsEnableParameter(rangeList, "None") And Not chCell.Value = "None" Then
        If boardType = "PPMU" Then
            GoTo DATA_ERROR
        ElseIf boardType = "ICUL1G" Then
            'Continue
        Else
            On Error GoTo 0
            Call Err.Raise(9999, "XLibDcScenarioUtility.verifyRangeData", "Internal error!")
            On Error GoTo DATA_ERROR
        End If
    End If
    '### ���͒l�����X�g�Ɋ܂܂�Ă��Ȃ���΃��[�j���O�����ݒ�
    '�s�����X�g��N.D�̎���IsEnableParameter��Flase�ƂȂ�
    If Not IsEnableParameter(rangeList, chCell.Value) Then
        '### ���͒l�̏����`�F�b�N #########################
        Dim paramRange As CParamStringWithUnit
        Set paramRange = CreateCParamStringWithUnit
        Select Case measureMode
            Case "MI":
                paramRange.Initialize "A"
            Case "MV":
                paramRange.Initialize "V"
            Case Else:
                GoTo DATA_ERROR
        End Select
        With paramRange.AsIParameter
            If boardType = "ICUL1G" And measureMode = "MV" Then
                If chCell.Value = Empty Then
                    chCell.Value = UPPER_VCLUMP_ICUL1G & "V"
                Else
                    'ICUL1G�Ɍ���㑤�N�����v�l������s��
                    '�i-1.5V<value<=6.5V�j
                    .UpperLimit = UPPER_VCLUMP_ICUL1G
                    .LowerLimit = LOWER_VCLUMP_ICUL1G
                    .AsString = chCell.Value
                    If CompareDblData(.AsDouble, LOWER_VCLUMP_ICUL1G, DIGIT_NUMBER) Then
                        Err.Raise 9999, "XLibDcScenarioUtility.verifyRangeData", "ICUL1G Upper Clamp must be > " & LOWER_VCLUMP_ICUL1G
                    End If
                End If
            Else
                .LowerLimit = 0
                .AsString = chCell.Value
            End If
        End With
        warningCell chCell
    End If
    Exit Sub
DATA_ERROR:
    errorCell chCell
End Sub

Private Sub verifyForceValue(ByVal chCell As Range, ByVal PinList As String)
    '### ���͈���p�����[�^�̏����E�X�y�b�N�̔��� #########
    '�O������F�ΏۃZ�����s���O���[�v�̐擪�ł��邱��
    '�s�����X�g��N.D�̎���compareForceLimit��Flase�ƂȂ�
    '�p�����[�^�������E�X�y�b�N�𖞂����Ȃ��ꍇ�A�p�����[�^�N���X�ŃG���[���N����
    '08/05/12 OK
    On Error GoTo DATA_ERROR
    enableCell chCell
    If IsEmpty(chCell) Then Exit Sub
    Dim boardType As String
    boardType = targetCell(chCell, TEST_PINTYPE).Value
    Dim measureMode As String
    measureMode = targetCell(chCell, SET_MODE).Value
    '### ���͒l�̏����`�F�b�N #############################
    Dim paramForce As CParamStringWithUnit
    Set paramForce = CreateCParamStringWithUnit
    '### �P�ʕt������łȂ��ꍇ ###########################
    If Not IsEmpty(targetCell(chCell, OPERATE_FORCE)) Then
        With paramForce
            .Initialize ""
            .AsIParameter.AsDouble = chCell.Value
        End With
    '### �P�ʕt������̏ꍇ ###############################
    Else
        Select Case measureMode
            Case "MI":
                paramForce.Initialize "V"
            Case "MV":
                paramForce.Initialize "A"
            Case Else:
                GoTo DATA_ERROR
        End Select
        paramForce.AsIParameter.AsString = chCell.Value
        '### ���͒l���X�y�b�N���łȂ���΃G���[�����ݒ� ###
        If Not compareForceLimit(PinList, boardType, measureMode, paramForce.AsIParameter.AsDouble) Then
            errorCell chCell
        End If
    End If
    Exit Sub
DATA_ERROR:
    errorCell chCell
End Sub

Private Sub verifyWaitValue(ByVal chCell As Range)
    '### �E�F�C�g���͒l�̏����`�F�b�N #####################
    '�O������F�ΏۃZ�����s���O���[�v�̐擪�ł��邱��
    '�p�����[�^�������E�X�y�b�N�𖞂����Ȃ��ꍇ�A�p�����[�^�N���X�ŃG���[���N����
    '08/05/12 OK
    On Error GoTo DATA_ERROR
    enableCell chCell
    Dim paramWait As CParamStringWithUnit
    Set paramWait = CreateCParamStringWithUnit
    With paramWait
        .Initialize "S"
        With .AsIParameter
            .LowerLimit = 0
            .AsString = chCell.Value
        End With
    End With
    Exit Sub
DATA_ERROR:
    errorCell chCell
End Sub

Private Sub verifyAverageValue(ByVal chCell As Range)
    '### �A�x���[�W���͒l�̏����`�F�b�N ###################
    '�O������F�ΏۃZ�����s���O���[�v�̐擪�ł��邱��
    '�p�����[�^�������E�X�y�b�N�𖞂����Ȃ��ꍇ�A�p�����[�^�N���X�ŃG���[���N����
    '08/05/12 OK
    On Error GoTo DATA_ERROR
    enableCell chCell
    Dim paramAvg As CParamLong
    Set paramAvg = CreateCParamLong
    With paramAvg.AsIParameter
        .LowerLimit = 1
        .AsLong = chCell.Value
    End With
    Exit Sub
DATA_ERROR:
    errorCell chCell
End Sub

Private Sub verifyListParamData(ByVal chCell As Range, ByVal ParamList As Collection)
    '### ���̑����X�g���͒l�̃`�F�b�N #####################
    '���X�g�ɑ��݂��Ȃ��p�����[�^��NG�Ƃ���
    '���X�g�����݂��Ȃ��ꍇ�̓`�F�b�N���Ȃ�
    '08/05/12 OK
    If ParamList Is Nothing Then Exit Sub
    If Not IsEnableParameter(ParamList, chCell.Value) Then
        errorCell chCell
    End If
End Sub

Private Sub verifyTestNameOrLabel(ByVal chCell As Range)
    '### �e�X�g�J�e�S���E�e�X�g���x���̒�`�̗L�������� ###
    '�����؁E���g�p
    On Error GoTo DATA_ERROR
    enableCell chCell
    '### �A�N�e�B�u�f�[�^�V�[�g�̎擾 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(JOB_LIST_TOOL, TEST_INSTANCES_TOOL)
    If DataSheet Is Nothing Then GoTo DATA_ERROR
    '### �C���X�^���X�V�[�g���[�_�[�̍쐬 #################
    Dim instanceReader As CInstanceSheetReader
    Set instanceReader = CreateCInstanceSheetReader
    With instanceReader
        .Initialize DataSheet.Name
        .AsIFileStream.SetLocation chCell.Value
    End With
    Exit Sub
DATA_ERROR:
    warningCell chCell
End Sub

Private Function actionList() As Collection
    '### �A�N�V�����p�����[�^���X�g #######################
    '����͑��݂���A�N�V�����̖₢���킹�悪�Ȃ��̂Œ��Ƀ��X�g�쐬
    Set actionList = New Collection
    With actionList
        .Add ""
        .Add "SET"
        .Add "MEASURE"
        .Add "OPEN"
    End With
End Function

Private Function postActionList() As Collection
    '### �|�X�g�A�N�V�����p�����[�^���X�g #################
    '����͑��݂���A�N�V�����̖₢���킹�悪�Ȃ��̂Œ��Ƀ��X�g�쐬
    Set postActionList = New Collection
    With postActionList
        .Add ""
        .Add "OPEN"
    End With
End Function

Private Function ioPinTypeList(ByVal PinList As String) As Collection
    '### �{�[�h�����X�g ###################################
    '�����PPMU/BPMU�̖₢���킹�悪�Ȃ��̂Ń��X�g�쐬
    Select Case queryBoardName(PinList)
        Case "PPMU":
            Set ioPinTypeList = New Collection
            With ioPinTypeList
                .Add "PPMU"
                .Add "BPMU"
            End With
        Case Else
            Set ioPinTypeList = Nothing
    End Select
End Function
'V21-Release
Private Function measureSiteList(ByVal PinList As String, ByVal boardType As String) As Collection
    '### ���[�W���[�T�C�g���[�h���X�g #####################
    '����̓��[�W���[�T�C�g���[�h�̖₢���킹�悪�Ȃ��̂Œ��Ƀ��X�g�쐬
    'Measure Action��BPMU�͂�����Concurrent�ɂ��Ă��A�N�V�����̃Z�b�g���ɋ����I��Serial�Ȃ̂Ŗ��Ȃ�

    Set measureSiteList = New Collection
    With measureSiteList
        If IsGangPins(PinList) Then
            .Add "Serial"
        Else
            .Add "Concurrent"
            .Add "Serial"
            'site�w��p
            Dim Num As Integer
            Num = GetSiteCount
            Dim i As Integer
            For i = 0 To Num
                .Add i
            Next i
        End If
    End With
    
End Function

Private Function modeList(ByVal PinList As String) As Collection
    '### ���[�W���[���[�h���X�g ###########################
    '����̓��[�W���[���[�h�̖₢���킹�悪�Ȃ��̂Œ��Ƀ��X�g�쐬
    Set modeList = New Collection
    With modeList
        Select Case queryBoardName(PinList)
            Case "DPS":
                .Add "MI"
            Case Else
                .Add "MV"
                .Add "MI"
        End Select
    End With
End Function

Private Function examinFlagList() As Collection
    '### �����t���O���X�g #################################
    Set examinFlagList = New Collection
    With examinFlagList
        .Add "FALSE"
        .Add "TRUE"
    End With
End Function

Private Function examinModeList() As Collection
    '### �������[�h���X�g #################################
    Set examinModeList = New Collection
    With examinModeList
        .Add ""
        .Add "BREAK"
        .Add "END"
    End With
End Function

Private Sub actionParamFormatter(ByVal chCell As Range, ByVal ParamList As Collection)
    '### �J�����g�A�N�V�����̃p�����[�^���W ###############
    '�O������F�ΏۃZ�����s���O���[�v�̐擪�ł��邱��

    '=== Add Eee-Job V2.14 ===
    Dim ErrorCells As Range
    Dim EnableCells As Range
    Dim DisableCells As Range
    '=== Add Eee-Job V2.14 ===

    Dim currCell As Range
    Set currCell = targetCell(chCell, TEST_ACTION)
    'EnableCell currCell                        'Add Eee-Job V2.14
    MakeUnionRange EnableCells, currCell        'Add Eee-Job V2.14

    '### �p�����[�^�����݂��Ȃ��A�N�V������NG #############
    If ParamList Is Nothing Or ParamList.Count = 0 Then
        'errorCell currCell                     'Add Eee-Job V2.14
        MakeUnionRange ErrorCells, currCell     'Add Eee-Job V2.14
    End If
    '### �V�[�g��f�[�^���x���̏��� #######################
    Dim shtDataList As New Collection
    With shtDataList
        .Add SET_MODE
        .Add SET_RANGE
        .Add SET_FORCE
        .Add MEASURE_WAIT
        .Add MEASURE_AVG
        .Add MEASURE_SITE
        .Add MEASURE_LABEL
        .Add OPERATE_FORCE
        .Add OPERATE_RESULT
    End With
    '### �s���O���[�v�Z���̎��W ###########################
    Dim pinCells As Collection
    Set pinCells = groupCellList(currCell, False)
    '### �p�����[�^�t�H�[�}�b�g�̐ݒ� #####################
    Dim currParam As Variant
    Dim currData As Collection
    Dim currPin As Range
    For Each currParam In shtDataList
        '### �t�H�[�}�b�g�ΏۃZ���̐ݒ� ###################
        If currParam = MEASURE_LABEL Or currParam = OPERATE_RESULT Then
            Set currData = pinCells
        Else
            Set currData = New Collection
            currData.Add pinCells.Item(1)
        End If
        '### �t�H�[�}�b�g�̏����ݒ� #######################
        For Each currPin In currData
            If IsEnableParameter(ParamList, currParam) Then
                'EnableCell targetCell(currPin, currParam)                      'Add Eee-Job V2.14
                MakeUnionRange EnableCells, targetCell(currPin, currParam)      'Add Eee-Job V2.14
            Else
                'disableCell targetCell(currPin, currParam)                     'Add Eee-Job V2.14
                MakeUnionRange DisableCells, targetCell(currPin, currParam)     'Add Eee-Job V2.14
            End If
        Next currPin
    Next currParam
    '=== Add Eee-Job V2.14 ===
    '�����ꊇ�ݒ�
    If Not EnableCells Is Nothing Then
        enableCell EnableCells
    End If
    If Not ErrorCells Is Nothing Then
        errorCell ErrorCells
    End If
    If Not DisableCells Is Nothing Then
        disableCell DisableCells
    End If
     '=== Add Eee-Job V2.14 ===
End Sub
'#V21-Release
Private Function actionParamList(ByVal currCell As Range) As Collection
    '### �A�N�V�����ɕK�v�ȃf�[�^���x���̃��X�g���쐬����
    '�O������F�ΏۃZ�����s���O���[�v�̐擪�ł��邱��
    '08/05/09 OK
    '10/08/20 BPMU��������+�G���[�Z���@�o�O�C��
    
    Set actionParamList = New Collection
    Dim currAction As Range
    Set currAction = targetCell(currCell, TEST_ACTION)
    Dim actions As New Collection
    If currAction.Value = "SET" Or IsEmpty(currAction) Then
        Dim setAct As New CSetFI
        actions.Add setAct
        actionParamList.Add SET_MODE
    End If
    If currAction.Value = "MEASURE" Or IsEmpty(currAction) Then
        Dim measAct As New CMeasureI
        Dim measPin As New CMeasurePin
        With actions
            .Add measAct
            .Add measPin
        End With
        actionParamList.Add SET_MODE
    End If
    If currAction.Value = "OPEN" Then
        Dim disconAct As New CDisconnect
        actions.Add disconAct
    End If
    actionParamList.Add MEASURE_SITE
    addActionParamList actionParamList, actions
    
    
End Function

Private Function IsEnableParameter(ByVal ParamList As Collection, ByVal dataLabel As String) As Boolean
    '### �p�����[�^���X�g�ɔC�ӂ̃f�[�^���x�������݂��邩�ǂ����𔻒肷��
    '08/05/09 OK
    If ParamList Is Nothing Then Exit Function
    Dim currParam As Variant
    For Each currParam In ParamList
        If currParam = dataLabel Then
            IsEnableParameter = True
            Exit Function
        End If
    Next currParam
End Function

Private Function addActionParamList(ByVal ParamList As Collection, ByVal actions As Collection)
    '### �o�^���ꂽ�A�N�V��������p�����[�^��ǂݍ��� #####
    '08/05/09 OK
    Dim currAct As IParameterWritable
    Dim currParam As Variant
    For Each currAct In actions
        For Each currParam In currAct.ParameterList
            ParamList.Add currParam
        Next currParam
    Next currAct
End Function

Private Function compareForceLimit(ByVal PinList As String, ByVal boardType As String, ByVal measureMode As String, ByVal inputVal As Double) As Boolean
    '### ���̓f�[�^�ƈ�����E�l�̔�r���s�� '08/04/30 OK ##
    'queryBoardSpec���疳���Ȓl���Ԃ��ꂽ�ꍇ��False
    Dim limitVal() As Double
    compareForceLimit = False
    Select Case boardType
        Case "BPMU":
            limitVal = queryBoardSpecForBPMU(PinList, measureMode)
        Case Else:
            limitVal = queryBoardSpec(PinList, measureMode)
    End Select
    If limitVal(0) = INVALIDATION_VALUE And limitVal(1) = INVALIDATION_VALUE Then GoTo IS_INVALID
    Dim roundHLimit As Double
    Dim roundLLimit As Double
    Dim roundFVal As Double
    roundHLimit = RoundDownDblData(limitVal(1), DIGIT_NUMBER)
    roundLLimit = RoundDownDblData(limitVal(0), DIGIT_NUMBER)
    roundFVal = RoundDownDblData(inputVal, DIGIT_NUMBER)
    If roundLLimit <= roundFVal And roundFVal <= roundHLimit Then
        compareForceLimit = True
    End If
    Exit Function
IS_INVALID:
End Function

Private Function enableBoardRange(ByVal PinList As String, ByVal boardType As String, ByVal measureMode As String) As Collection
    '### �s�����X�g����L���ȃ����W���X�g���擾���� #######
    '�@BPMU�ȊO�̓s�������烊�\�[�X���������肷��̂�boardType�̈����͖����ɂȂ�
    '�ABPMU�w��̓��\�[�X�̌ŗL�l��Ԃ��̂�PinList�̈����͖����ɂȂ�
    '�B�M�����O�s���̃��X�g�͍ő僌���W*�M�����O���̂�
    '08/05/09 OK
    Select Case boardType
        Case "BPMU":
            Set enableBoardRange = queryBoardRangeForBPMU(PinList, measureMode)
        Case Else
            Dim tempRange As Collection
            Set tempRange = queryBoardRange(PinList, measureMode)
            If tempRange Is Nothing Then
                Set enableBoardRange = Nothing
                Exit Function
            End If
            If IsGangPins(PinList) And measureMode = "MI" Then
                Set enableBoardRange = New Collection
                Dim maxRange As Variant
                'HDVIS��200mA���ߑł��F���t�F�[�Y�ŏC�����K�v
                Select Case boardType
                    Case "HDVIS":
                        maxRange = "200mA"
                    Case "APMU"
                        maxRange = tempRange.Item(1)
                End Select
                Dim MainUnit As String
                Dim SubUnit As String
                Dim SubValue As Double
                SplitUnitValue maxRange, MainUnit, SubUnit, SubValue
                Dim gangRange As Double
                gangRange = RoundDownDblData(SubValue * GetGangPinCount(PinList), DIGIT_NUMBER)
                enableBoardRange.Add gangRange & SubUnit & MainUnit
            Else
                Set enableBoardRange = tempRange
            End If
    End Select
End Function

Private Function queryBoardName(ByVal PinList As String) As String
    '### �{�[�h���̎擾 '08/04/28 OK ######################
    '�@�P���s��/�J���}��؂�̃s�����X�g��OK
    '�A�ꕔ�ɃM�����O/�}�[�W���܂ރs�����X�g�ł�OK
    '�BDPS�̃M�����O�s���܂��͂�����܂ރs�����X�g��NG
    '�C�قȂ�{�[�h�̃s�����X�g��NG
    '�D��`����Ă��Ȃ��s���܂��͂�����܂ރs�����X�g��NG
    Dim ip750board As New CVISVISrcSelector
    ip750board.Initialize
    On Error GoTo IS_INVALID
    queryBoardName = Replace(ip750board.Pins(PinList).BoardName, "dc", "")
    Exit Function
IS_INVALID:
    queryBoardName = NOT_DEFINE
End Function

Private Function queryBoardSpec(ByVal PinList As String, ByVal measureMode As String) As Double()
    '### �{�[�h������E�l�̎擾 '08/04/28 OK ##############
    '�@�P���s��/�J���}��؂�̃s�����X�g��OK
    '�A��L�̏ꍇ�ADPS��GetForceILimit��NG
    '�B�ꕔ�ɃM�����O/�}�[�W���܂ރs�����X�g��GetForceVLimit��OK
    '�C�ꕔ�ɃM�����O/�}�[�W���܂ރs�����X�g��GetForceILimit��NG
    '�DDPS�̃M�����O�s���܂��͂�����܂ރs�����X�g��NG
    '�E�قȂ�{�[�h�̃s�����X�g��NG
    '�F��`����Ă��Ȃ��s���܂��͂�����܂ރs�����X�g��NG
    Dim ip750board As New CVISVISrcSelector
    ip750board.Initialize
    Dim errLimit(1) As Double
    errLimit(0) = INVALIDATION_VALUE
    errLimit(1) = INVALIDATION_VALUE
    On Error GoTo IS_INVALID
    Select Case measureMode
        Case "MI":
            queryBoardSpec = ip750board.Pins(PinList).GetForceVLimit
        Case "MV":
            queryBoardSpec = ip750board.Pins(PinList).GetForceILimit
        Case Else
            queryBoardSpec = errLimit
    End Select
    Exit Function
IS_INVALID:
    queryBoardSpec = errLimit
End Function

Private Function queryBoardSpecForBPMU(ByVal PinList As String, ByVal measureMode As String) As Double()
    '### BPMU�{�[�h������E�l�̎擾 '08/04/28 OK ##########
    '�@�P���s��/�J���}��؂�̃s�����X�g��OK
    '�A�قȂ�{�[�h�s���܂��͂�����܂ރs�����X�g�ł�OK�ƂȂ��Ă��܂�
    '�B��`����Ă��Ȃ��s���܂��͂�����܂ރs�����X�g�ł�OK�ƂȂ��Ă��܂�
    '�CNOT_DEFINE�̃s����NG
    Dim ip750board As New CVISVISrcSelector
    ip750board.Initialize
    Dim errLimit(1) As Double
    errLimit(0) = INVALIDATION_VALUE
    errLimit(1) = INVALIDATION_VALUE
    If PinList = NOT_DEFINE Then GoTo IS_INVALID
    On Error GoTo IS_INVALID
    Select Case measureMode
        Case "MI":
            queryBoardSpecForBPMU = ip750board.Pins(PinList, dcBPMU).GetForceVLimit
        Case "MV":
            queryBoardSpecForBPMU = ip750board.Pins(PinList, dcBPMU).GetForceILimit
        Case Else
            queryBoardSpecForBPMU = errLimit
    End Select
    Exit Function
IS_INVALID:
    queryBoardSpecForBPMU = errLimit
End Function

Private Function queryBoardRange(ByVal PinList As String, ByVal measureMode As String) As Collection
    '### �{�[�h���背���W���X�g�̎擾 '08/04/28 OK ########
    '�@�P���s��/�J���}��؂�̃s�����X�g��OK
    '�A��L�̏ꍇ�APPMU/DPS��MeasVRangeList�̕Ԃ�l��"None"�̂�
    '�B�ꕔ�ɃM�����O/�}�[�W���܂ރs�����X�g�̓M�����O/�}�[�W�s�����D�悳���
    '�CDPS�̃M�����O�s���܂��͂�����܂ރs�����X�g��NG
    '�D�قȂ�{�[�h�̃s�����X�g��NG
    '�E��`����Ă��Ȃ��s���܂��͂�����܂ރs�����X�g��NG
    Dim ip750board As New CVISVISrcSelector
    ip750board.Initialize
    Dim rangeList As Collection
    On Error GoTo IS_INVALID
    Select Case measureMode
        Case "MI":
            Set queryBoardRange = DecomposeStringList(ip750board.Pins(PinList).MeasIRangeList)
        Case "MV":
            If ip750board.Pins(PinList).BoardName = "dcICUL1G" Then
                'ICUL1G�ł�Range�̃��X�g�͂Ȃ�
                Set queryBoardRange = Nothing
            Else
                Set queryBoardRange = DecomposeStringList(ip750board.Pins(PinList).MeasVRangeList)
            End If
        Case Else
            Set queryBoardRange = Nothing
    End Select
    Exit Function
IS_INVALID:
    Set queryBoardRange = Nothing
End Function

Private Function queryBoardRangeForBPMU(ByVal PinList As String, ByVal measureMode As String) As Collection
    '### BPMU�{�[�h���背���W���X�g�̎擾 '08/04/28 OK ####
    '�@�P���s��/�J���}��؂�̃s�����X�g��OK
    '�A�قȂ�{�[�h�s���܂��͂�����܂ރs�����X�g�ł�OK�ƂȂ��Ă��܂�
    '�B��`����Ă��Ȃ��s���܂��͂�����܂ރs�����X�g�ł�OK�ƂȂ��Ă��܂�
    '�CNOT_DEFINE�̃s����NG
    Dim ip750board As New CVISVISrcSelector
    ip750board.Initialize
    Dim rangeList As Collection
    If PinList = NOT_DEFINE Then GoTo IS_INVALID
    On Error GoTo IS_INVALID
    Select Case measureMode
        Case "MI":
            Set queryBoardRangeForBPMU = DecomposeStringList(ip750board.Pins(PinList, dcBPMU).MeasIRangeList)
        Case "MV":
            Set queryBoardRangeForBPMU = DecomposeStringList(ip750board.Pins(PinList, dcBPMU).MeasVRangeList)
        Case Else
            Set queryBoardRangeForBPMU = Nothing
    End Select
    Exit Function
IS_INVALID:
    Set queryBoardRangeForBPMU = Nothing
End Function

Private Function createPinList(ByVal currCell As Range, ByVal cellFrmt As Boolean) As String
    '### �s�����X�g�̍쐬 #################################
    '08/05/09 OK
    On Error GoTo IS_INVALID
    Dim pinCells As Collection
    Set pinCells = groupCellList(currCell, cellFrmt)
    If pinCells Is Nothing Then GoTo IS_INVALID
    '### �J���}��؂�̕�����ɓW�J #######################
    createPinList = ComposeStringList(pinCells)
    Exit Function
IS_INVALID:
    createPinList = NOT_DEFINE
End Function

Private Function groupCellList(ByVal currCell As Range, ByVal cellFrmt As Boolean) As Collection
    '### �O���[�v�Z���I�u�W�F�N�g�̎��W ###################
    '08/05/09 OK
    On Error GoTo IS_INVALID
    Dim dataCell As Range
    Set dataCell = targetCell(currCell, TEST_PINS)
    If IsEmpty(dataCell) Or Not isEnableArea(currCell) Then GoTo IS_INVALID
    '### �O���[�v�̓��o�� #################################
    Dim dataIndex As Long
    Dim cellIndex As Range
    Do While IsGroupHeader(dataCell.offset(dataIndex, 0)) = False
        dataIndex = dataIndex - 1
    Loop
    Set cellIndex = dataCell.offset(dataIndex, 0)
    '### �O���[�v�̃Z���I�u�W�F�N�g�����W #################
    dataIndex = 0
    Set groupCellList = New Collection
    Do While IsGroupFooter(cellIndex.offset(dataIndex, 0)) = False
        groupCellList.Add cellIndex.offset(dataIndex, 0)
        dataIndex = dataIndex + 1
    Loop
    '### �L���ȃs�����̂݃��X�g�Ɋi�[ #####################
    Dim enablePinList As Collection
    Set enablePinList = collectEnablePins(groupCellList, cellFrmt)
    '### �����ȃs�����������ꍇ�͖������X�g��Ԃ� #########
    If groupCellList.Count <> enablePinList.Count Then GoTo IS_INVALID
    Exit Function
IS_INVALID:
    Set groupCellList = Nothing
End Function

Private Function collectEnablePins(ByVal PinList As Collection, ByVal cellFrmt As Boolean) As Collection
    '### �s�����X�g����L���ȃs���������W���� #############
    '08/05/09 OK
    Set collectEnablePins = New Collection

    '=== Add Eee-Job V2.14 ===
    '�����ݒ�Z���p
    Dim EnableCells As Range
    Dim ErrorCells As Range
    '=== Add Eee-Job V2.14 ===

    Dim currPin As Range
    Dim topPinType As String
    '### �擪�s���̃{�[�h�����擾 #########################
    topPinType = queryBoardName(PinList.Item(1))
    Dim currType As String
    For Each currPin In PinList
        '### �擪�s���������̏ꍇ�͑S��NG #################
        If topPinType = NOT_DEFINE Then
            If cellFrmt Then
                'errorCell currPin                      '=== Add Eee-Job V2.14 ===
                MakeUnionRange ErrorCells, currPin      '=== Add Eee-Job V2.14 ===
            End If
        Else
        '=== Add Eee-Job V2.14 ===
            currType = queryBoardName(currPin)
            '### �����y�ѐ擪�s���ƈقȂ郊�\�[�X�̏ꍇ��NG
            If currType = NOT_DEFINE Or currType <> topPinType Then
                If cellFrmt Then
                    'errorCell currPin
                    MakeUnionRange ErrorCells, currPin
                End If
            '### �����s�����M�����O�s���̏ꍇ��NG #######
            ElseIf IsGangPins(currPin) And PinList.Count > 1 Then
                If cellFrmt Then
                    'errorCell currPin
                    MakeUnionRange ErrorCells, currPin
                End If
            '### �L���s���̂݃s���������W #################
            Else
                collectEnablePins.Add currPin.Value
                If cellFrmt Then
                    'enableCell currPin
                    MakeUnionRange EnableCells, currPin
                End If
            End If
        End If
    Next currPin

    '�ꊇ�����ݒ�
    If Not EnableCells Is Nothing Then
        enableCell EnableCells
    End If
    If Not ErrorCells Is Nothing Then
        errorCell ErrorCells
    End If
        '=== Add Eee-Job V2.14 ===
End Function
'=== Add Eee-Job V2.14 ===
'�����W�����p�̊֐�
Private Sub MakeUnionRange(ByRef pUnionRange As Range, ByRef pCurrentRange As Range)
    If pUnionRange Is Nothing Then
        Set pUnionRange = pCurrentRange
    Else
        Set pUnionRange = Union(pUnionRange, pCurrentRange)
    End If
End Sub

Private Sub enableCell(ByVal currCell As Range)
    '### �Z���̏����ݒ�F�L���ł��鎖��\������ ###########
    With currCell.Font
        If .Name <> "Tahoma" Or IsNull(.Name) Then .Name = "Tahoma"
        If .FontStyle <> "�W��" Or IsNull(.FontStyle) Then .FontStyle = "�W��"
        If .Size <> 10 Or IsNull(.Size) Then .Size = 10
        If .Strikethrough <> False Or IsNull(.Strikethrough) Then .Strikethrough = False
        If .Superscript <> False Or IsNull(.Superscript) Then .Superscript = False
        If .Subscript <> False Or IsNull(.Subscript) Then .Subscript = False
        If .OutlineFont <> False Or IsNull(.OutlineFont) Then .OutlineFont = False
        If .Shadow <> False Or IsNull(.Shadow) Then .Shadow = False
        If .Underline <> xlUnderlineStyleNone Or IsNull(.Underline) Then .Underline = xlUnderlineStyleNone
        If .ColorIndex <> xlAutomatic Or IsNull(.ColorIndex) Then .ColorIndex = xlAutomatic
    End With
    With currCell.Interior
        If .ColorIndex <> xlNone Or IsNull(.ColorIndex) Then .ColorIndex = xlNone
    End With
End Sub

Private Sub disableCell(ByVal currCell As Range)
    '### �Z���̏����ݒ�F�����ł��鎖��\������ ###########
    With currCell.Font
        If .Name <> "Tahoma" Or IsNull(.Name) Then .Name = "Tahoma"
        If .FontStyle <> "�W��" Or IsNull(.FontStyle) Then .FontStyle = "�W��"
        If .Size <> 10 Or IsNull(.Size) Then .Size = 10
        If .Strikethrough <> False Or IsNull(.Strikethrough) Then .Strikethrough = False
        If .Superscript <> False Or IsNull(.Superscript) Then .Superscript = False
        If .Subscript <> False Or IsNull(.Subscript) Then .Subscript = False
        If .OutlineFont <> False Or IsNull(.OutlineFont) Then .OutlineFont = False
        If .Shadow <> False Or IsNull(.Shadow) Then .Shadow = False
        If .Underline <> xlUnderlineStyleNone Or IsNull(.Underline) Then .Underline = xlUnderlineStyleNone
        If .ColorIndex <> 16 Or IsNull(.ColorIndex) Then .ColorIndex = 16
    End With
    With currCell.Interior
        If .ColorIndex <> 0 Or IsNull(.ColorIndex) Then .ColorIndex = 0
        If .Pattern <> xlLightUp Or IsNull(.Pattern) Then .Pattern = xlLightUp
        If .PatternColorIndex <> 48 Or IsNull(.PatternColorIndex) Then .PatternColorIndex = 48
    End With
End Sub

Private Sub errorCell(ByVal currCell As Range)
    '### �Z���̏����ݒ�F�f�[�^���̓~�X��\������ #########
    With currCell.Font
        If .Name <> "Tahoma" Or IsNull(.Name) Then .Name = "Tahoma"
        If .FontStyle <> "�W��" Or IsNull(.FontStyle) Then .FontStyle = "�W��"
        If .Size <> 10 Or IsNull(.Size) Then .Size = 10
        If .Strikethrough <> False Or IsNull(.Strikethrough) Then .Strikethrough = False
        If .Superscript <> False Or IsNull(.Superscript) Then .Superscript = False
        If .Subscript <> False Or IsNull(.Subscript) Then .Subscript = False
        If .OutlineFont <> False Or IsNull(.OutlineFont) Then .OutlineFont = False
        If .Shadow <> False Or IsNull(.Shadow) Then .Shadow = False
        If .Underline <> xlUnderlineStyleNone Or IsNull(.Underline) Then .Underline = xlUnderlineStyleNone
        If .ColorIndex <> xlAutomatic Or IsNull(.ColorIndex) Then .ColorIndex = xlAutomatic
    End With
    With currCell.Interior
        If .ColorIndex <> 38 Or IsNull(.ColorIndex) Then .ColorIndex = 38
    End With
End Sub

Private Sub warningCell(ByVal currCell As Range)
    '### �Z���̏����ݒ�F�f�[�^�������ł��鎖��\������ ###
    With currCell.Font
        If .Name <> "Tahoma" Or IsNull(.Name) Then .Name = "Tahoma"
        If .FontStyle <> "�W��" Or IsNull(.FontStyle) Then .FontStyle = "�W��"
        If .Size <> 10 Or IsNull(.Size) Then .Size = 10
        If .Strikethrough <> False Or IsNull(.Strikethrough) Then .Strikethrough = False
        If .Superscript <> False Or IsNull(.Superscript) Then .Superscript = False
        If .Subscript <> False Or IsNull(.Subscript) Then .Subscript = False
        If .OutlineFont <> False Or IsNull(.OutlineFont) Then .OutlineFont = False
        If .Shadow <> False Or IsNull(.Shadow) Then .Shadow = False
        If .Underline <> xlUnderlineStyleNone Or IsNull(.Underline) Then .Underline = xlUnderlineStyleNone
        If .ColorIndex <> xlAutomatic Or IsNull(.ColorIndex) Then .ColorIndex = xlAutomatic
    End With
    With currCell.Interior
        If .ColorIndex <> 36 Or IsNull(.ColorIndex) Then .ColorIndex = 36
    End With
End Sub
'=== Add Eee-Job V2.14 ===
Private Function IsCategoryHeader(ByVal currCell As Range) As Boolean
    '### �J�e�S���̃p�����[�^�L���s���� ###################
    '08/05/01 OK
    '�w��Z���̃J�e�S���J�����Ƀf�[�^���͂�����ƗL��
    If isEnableArea(currCell) Then
        Dim categoryCell As Range
        Set categoryCell = targetCell(currCell, TEST_CATEGORY)
        If Not IsEmpty(categoryCell) Then
            IsCategoryHeader = True
        End If
    End If
End Function

Private Function IsGroupHeader(ByVal currCell As Range) As Boolean
    '### �s���O���[�v�̃p�����[�^�L���s���� ###############
    '08/05/01 OK
    '�@�w��Z���̃��W���[�s���J�����ɓ��͂�����
    '�A���̃J�����̒��O�̃J�������󔒂ł��閔�̓p�����[�^�L���̈�̐擪�Z���ł���
    '��L2�_�𖞂����ƗL���ƂȂ�
    If isEnableArea(currCell) Then
        Dim pinNameCell As Range
        Set pinNameCell = targetCell(currCell, TEST_PINS)
        If Not IsEmpty(pinNameCell) Then
            Dim topCell As Range
            Set topCell = currCell.parent.Range(TEST_PINS)
            If IsEmpty(pinNameCell.offset(-1, 0)) Or _
               topCell.Row = pinNameCell.offset(-1, 0).Row Then
                IsGroupHeader = True
            End If
        End If
    End If
End Function

Private Function IsGroupFooter(ByVal currCell As Range) As Boolean
    '### �s���O���[�v�̃t�b�^�[�p�����[�^�L���s���� #######
    '08/05/01 OK
    '�@�w��Z���̃��W���[�s���J�������󔒂ł���
    '�A���̃J�����̒��O�̃J�����ɓ��͂����閔�̓p�����[�^�L���̈�̐擪�Z���łȂ�
    '��L2�_�𖞂����ƗL���ƂȂ�
    If isEnableArea(currCell) Then
        Dim pinNameCell As Range
        Set pinNameCell = targetCell(currCell, TEST_PINS)
        If IsEmpty(pinNameCell) Then
            Dim topCell As Range
            Set topCell = currCell.parent.Range(TEST_PINS)
            If Not IsEmpty(pinNameCell.offset(-1, 0)) Or _
               topCell.Row <> pinNameCell.offset(-1, 0).Row Then
                IsGroupFooter = True
            End If
        End If
    End If
End Function

Private Function isEnableArea(ByVal currCell As Range) As Boolean
    '### �w��Z���̃p�����[�^�L���̈攻�� #################
    '08/05/01 OK
    'END�L�[���[�h�����݂��Ȃ����͍ŏI�s�܂ŗL��
    Dim topCell As Range
    Dim endCell As Range
    With currCell.parent
        Set topCell = .Range(TEST_CATEGORY)
        Set endCell = .Columns(topCell.Column).Find("END")
    End With
    If currCell.Row > topCell.Row Then
        If endCell Is Nothing Then
            isEnableArea = True
        ElseIf endCell.Row > currCell.Row Then
            isEnableArea = True
        End If
    End If
End Function

Private Function targetCell(ByVal refCell As Range, ByVal targetLabel As String) As Range
    With refCell.parent
        Set targetCell = .Cells(refCell.Row, .Range(targetLabel).Column)
    End With
End Function

'############# �G�N�Z���V�[�g�}�N���֐��Q #########################################################
Public Sub ShowSpecInfo()
'���e:
'   DC�V�i���I�V�[�g�X�y�b�N�\���}�N���֐�
'
'�p�����[�^:
'
'���ӎ���:
'   �A�N�e�B�u��DC�e�X�g�V�i���I�V�[�g��Ńe�X�g���x���Z�����E�N���b�N�������̃��j���[����Ăяo�����
'
    On Error GoTo ErrHandler
    '### �A�N�e�B�u�f�[�^�V�[�g�̎擾 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(JOB_LIST_TOOL, TEST_INSTANCES_TOOL)
    If DataSheet Is Nothing Then
        Err.Raise 9999, "showSpecInfo", "Can Not Find The Active Instance Sheet !"
    End If
    '### �A�N�e�B�u�Z���̎擾 #############################
    Dim currCell As Range
    Set currCell = ActiveSheet.Application.ActiveCell
    '### �C���X�^���X�V�[�g���[�_�[�̍쐬 #################
    Dim instanceReader As CInstanceSheetReader
    Set instanceReader = CreateCInstanceSheetReader
    With instanceReader
        .Initialize DataSheet.Name
        .AsIFileStream.SetLocation currCell.Value
    End With
    '### �e�p�����[�^�I�u�W�F�N�g�쐬�Ɠǂݍ��� ###########
    Dim paramLoLimit As CParamDouble
    Set paramLoLimit = CreateCParamDouble
    With paramLoLimit.AsIParameter
        .Name = USERMACRO_LOLIMIT
        .Read instanceReader
    End With
    Dim paramHiLimit As CParamDouble
    Set paramHiLimit = CreateCParamDouble
    With paramHiLimit.AsIParameter
        .Name = USERMACRO_HILIMIT
        .Read instanceReader
    End With
    Dim paramJudge As CParamLong
    Set paramJudge = CreateCParamLong
    Dim judgeChar As String
    With paramJudge.AsIParameter
        .Name = USERMACRO_JUDGE
        .Read instanceReader
        If .AsDouble = 0 Then
            judgeChar = "NONE"
        ElseIf .AsDouble = 1 Then
            judgeChar = "D <"
        ElseIf .AsDouble = 2 Then
            judgeChar = "D >"
        ElseIf .AsDouble = 3 Then
            judgeChar = "< D >"
        Else
            judgeChar = "ERROR"
        End If
    End With
    Dim paramUnit As CParamString
    Set paramUnit = CreateCParamString
    Dim MainUnit As String
    Dim SubUnit As String
    Dim SubValue As Double
    With paramUnit.AsIParameter
        .Name = USERMACRO_UNIT
        .Read instanceReader
        SplitUnitValue "999" & .AsString, MainUnit, SubUnit, SubValue
    End With
    '### �X�y�b�N���̕\�� ###############################
    MsgBox "INSTANCE SHEET" & Chr(9) & " : [" & DataSheet.Name & "]" & Chr(13) & _
           "TEST LABEL" & Chr(9) & " : [" & currCell.Value & "]" & Chr(13) & _
           "LOWER LIMIT" & Chr(9) & " : [" & paramLoLimit.AsIParameter.AsDouble / SubUnitToValue(SubUnit) & paramUnit.AsIParameter.AsString & "]" & Chr(13) & _
           "UPPER LIMIT" & Chr(9) & " : [" & paramHiLimit.AsIParameter.AsDouble / SubUnitToValue(SubUnit) & paramUnit.AsIParameter.AsString & "]" & Chr(13) & _
           "JUDEGE" & Chr(9) & Chr(9) & " : LOWER [ " & judgeChar & " ] UPPER", vbOKOnly + vbInformation, "SPEC INFOMATION"
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub CheckExaminationMode()
'���e:
'   DC�V�i���I�V�[�g�����t���O�̃`�F�b�N�}�N���֐�
'
'�p�����[�^:
'
'���ӎ���:
'   ���[�N�u�b�N�������钼�O�ɌĂяo�����
'
    '### �A�N�e�B�u�f�[�^�V�[�g�̎擾 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    '### DC�V�i���I�V�[�g���[�_�[�̍쐬 ###################
    Dim scenarioReader As CDcScenarioSheetReader
    Set scenarioReader = CreateCDcScenarioSheetReader
    scenarioReader.Initialize DataSheet.Name
    '### �p�����[�^�I�u�W�F�N�g�̍쐬 #####################
    Dim paramExamin As CParamBoolean
    Set paramExamin = CreateCParamBoolean
    With paramExamin.AsIParameter
        .Name = EXAMIN_FLAG
    End With
    '### �����t���O��ON�ݒ�̃J�E���g #####################
    Dim trueMode As Long
    Do While Not scenarioReader.AsIActionStream.IsEndOfCategory
        paramExamin.AsIParameter.Read scenarioReader
        If paramExamin.AsIParameter.AsBoolean Then
            trueMode = trueMode + 1
        End If
        scenarioReader.AsIActionStream.MoveNextCategory
    Loop
    '### �����t���O��ON�̏ꍇ���ӂ𑣂� ###################
    Dim myAns As Integer
    If trueMode > 0 Then
        myAns = MsgBox("[DC Test Scenario]" & vbCrLf & trueMode & _
                       " Cells With 'TRUE' Found In [Examination - Flag] Field!" & vbCrLf & _
                       " Do You Want To Replace Them With 'FALSE' ?", _
                         vbYesNo + vbExclamation, "Examination Mode Alert")
        Select Case myAns:
            Case vbYes:
                clearExaminMode DataSheet
                clearResultData DataSheet
                reverseExaminFlag DataSheet, False
            Case vbNo:
        End Select
    End If
    '### �p�����[�^�I�u�W�F�N�g�̍쐬 #####################
    Dim paramMode As CParamBoolean
    Set paramMode = CreateCParamBoolean
    With paramMode.AsIParameter
        .Name = IS_VALIDATE
        .Read scenarioReader
    End With
    '### �����W�o���f�[�V�������[�h��ON�̏ꍇ���ӂ𑣂� ###
    If paramMode.AsIParameter.AsBoolean Then
        myAns = MsgBox("[DC Test Scenario]" & vbCrLf & _
                       " 'TRUE' Found In [Range Validation Check Box] Field!" & vbCrLf & _
                       " Do You Want To Replace It With 'FALSE' ?", _
                         vbYesNo + vbExclamation, "Examination Mode Alert")
        Select Case myAns:
            Case vbYes:
                clearValidationMode DataSheet
            Case vbNo:
        End Select
    End If
End Sub

Public Sub SetChangeStatus()
'���e:
'   DC�V�i���I�V�[�g�p�����[�^�ύX�v���p�e�B����}�N���֐�
'
'�p�����[�^:
'
'���ӎ���:
'
    '### �A�N�e�B�u�f�[�^�V�[�g�̎擾 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    '### �p�����[�^�v���p�e�B���� #########################
    Dim shtObject As Object
    Set shtObject = DataSheet
    shtObject.IsChanged = True
End Sub

Public Sub SwitchExamFlag()
'���e:
'   DC�V�i���I�V�[�g�����t���O�̐؂�ւ��}�N���֐�
'
'�p�����[�^:
'
'���ӎ���:
'
    On Error GoTo ErrHandler
    '### �A�N�e�B�u�f�[�^�V�[�g�̎擾 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    '### DC�V�i���I�V�[�g���[�_�[�̍쐬 ###################
    Dim scenarioReader As CDcScenarioSheetReader
    Set scenarioReader = CreateCDcScenarioSheetReader
    scenarioReader.Initialize DataSheet.Name
    '### �p�����[�^�I�u�W�F�N�g�̍쐬 #####################
    Dim paramExamin As CParamBoolean
    Set paramExamin = CreateCParamBoolean
    With paramExamin.AsIParameter
        .Name = EXAMIN_FLAG
        .Read scenarioReader
    End With
    '### �擪�̎����t���O�̓ǂݍ��� #######################
    Dim WriteFlag As Boolean
    If paramExamin.AsIParameter.AsBoolean Then
        WriteFlag = False
    Else
        WriteFlag = True
    End If
    '### �����t���O�̐؂�ւ����s #########################
    reverseExaminFlag DataSheet, WriteFlag
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub SetExamMode()
'���e:
'   DC�V�i���I�V�[�g�������[�h�N���A�}�N���֐�
'
'�p�����[�^:
'
'���ӎ���:
'   �}�N���o�^���ꂽ�{�^���̃N���b�N�ŌĂяo�����
'
    On Error GoTo ErrHandler
    '### �A�N�e�B�u�f�[�^�V�[�g�̎擾 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    '### �������[�h�ݒ�̃N���A���s #######################
    clearExaminMode DataSheet
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub ClearExamResult()
'���e:
'   DC�V�i���I�V�[�g�������ʃN���A�}�N���֐�
'
'�p�����[�^:
'
'���ӎ���:
'   �}�N���o�^���ꂽ�{�^���̃N���b�N�ŌĂяo�����
'
    On Error GoTo ErrHandler
    '### �A�N�e�B�u�f�[�^�V�[�g�̎擾 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    '### ���茋�ʃf�[�^�̃N���A���s #######################
    clearResultData DataSheet
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub MakePlaybackTable()
'���e:
'   DC�Đ����t�@�����X�f�[�^�쐬�}�N���֐�
'
'�p�����[�^:
'
'���ӎ���:
'   �}�N���o�^���ꂽ�{�^���̃N���b�N�ŌĂяo�����
'
    On Error GoTo ErrHandler
    '### DC�V�i���I�V�[�g���[�_�[�̍쐬 ###################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then
        Err.Raise 9999, "MakePlaybackTable", "Can Not Find The Active Dc Scenario Sheet !"
    End If
    Dim scenarioReader As CDcScenarioSheetReader
    Set scenarioReader = CreateCDcScenarioSheetReader
    scenarioReader.Initialize DataSheet.Name
    '### DC�Đ��f�[�^�V�[�g���C�^�[�̍쐬 #################
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_PLAYBACK_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    Dim playbackWriter As CDcPlaybackSheetWriter
    Set playbackWriter = CreateCDcPlaybackSheetWriter
    With playbackWriter
        .Initialize DataSheet.Name
        .ClearCells
    End With
    '### �C���X�^���X�V�[�g���[�_�[�̍쐬 #################
    Set DataSheet = GetUsableDataSht(JOB_LIST_TOOL, TEST_INSTANCES_TOOL)
    If DataSheet Is Nothing Then
        Err.Raise 9999, "MakePlaybackTable", "Can Not Find The Active Instance Sheet !"
    End If
    Dim instanceReader As CInstanceSheetReader
    Set instanceReader = CreateCInstanceSheetReader
    instanceReader.Initialize DataSheet.Name
    '### �p�����[�^�I�u�W�F�N�g�̍쐬 #####################
    Dim paramTName As CParamString
    Set paramTName = CreateCParamString
    Dim ParamLabel As CParamName
    Set ParamLabel = CreateCParamName
    Dim paramLLow As CParamDouble
    Set paramLLow = CreateCParamDouble
    paramLLow.AsIParameter.Name = USERMACRO_LOLIMIT
    Dim paramLHi As CParamDouble
    Set paramLHi = CreateCParamDouble
    paramLHi.AsIParameter.Name = USERMACRO_HILIMIT
    Dim paramUnit As CParamString
    Set paramUnit = CreateCParamString
    paramUnit.AsIParameter.Name = USERMACRO_UNIT
    Dim paramSLLow As CParamString
    Set paramSLLow = CreateCParamString
    paramSLLow.AsIParameter.Name = PB_LIMIT_LO
    Dim paramSLHi As CParamString
    Set paramSLHi = CreateCParamString
    paramSLHi.AsIParameter.Name = PB_LIMIT_HI
    '### �f�[�^�e�[�u���쐬���s ###########################
    Do While Not scenarioReader.AsIActionStream.IsEndOfCategory
        With paramTName.AsIParameter
            .Name = TEST_CATEGORY
            .Read scenarioReader
            .Name = PB_CATEGORY
            .WriteOut playbackWriter
        End With
        Do While Not scenarioReader.AsIActionStream.IsEndOfGroup
            Do While Not scenarioReader.AsIActionStream.IsEndOfData
                With ParamLabel.AsIParameter
                    .Name = MEASURE_LABEL
                    .Read scenarioReader
                End With
                If ParamLabel.AsIParameter.AsString <> NOT_DEFINE Then
                    With ParamLabel.AsIParameter
                        .Name = PB_LABEL
                        .WriteOut playbackWriter
                    End With
                    instanceReader.AsIFileStream.SetLocation ParamLabel.AsIParameter.AsString
                    Dim MainUnit As String
                    Dim SubUnit As String
                    Dim SubValue As Double
                    With paramUnit.AsIParameter
                        .Read instanceReader
                        SplitUnitValue "999" & .AsString, MainUnit, SubUnit, SubValue
                    End With
                    paramLHi.AsIParameter.Read instanceReader
                    paramLLow.AsIParameter.Read instanceReader
                    With paramSLHi.AsIParameter
                        .AsString = paramLHi.AsIParameter.AsDouble / SubUnitToValue(SubUnit) & paramUnit.AsIParameter.AsString
                        .WriteOut playbackWriter
                    End With
                    With paramSLLow.AsIParameter
                        .AsString = paramLLow.AsIParameter.AsDouble / SubUnitToValue(SubUnit) & paramUnit.AsIParameter.AsString
                        .WriteOut playbackWriter
                    End With
                    playbackWriter.AsIFileStream.MoveNext
                End If
                scenarioReader.AsIActionStream.MoveNextData
            Loop
            scenarioReader.AsIActionStream.MoveNextGroup
        Loop
        scenarioReader.AsIActionStream.MoveNextCategory
    Loop
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub RangeValidationCheckBox_Click()
'���e:
'   DC�V�i���I���O���|�[�g�`�F�b�N�{�b�N�X�̃C�x���g�}�N��
'
'�p�����[�^:
'
'���ӎ���:
'   �V�i���I�V�[�g��̃`�F�b�N�{�b�N�X�̃I��/�I�t�ŌĂяo�����
'
    On Error GoTo ErrHandler
    '### �A�N�e�B�u�f�[�^�V�[�g�̎擾 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    '### �V�[�g�}�l�[�W���̏����� #########################
    InitControlShtReader
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub HoldSheetInfo(ByVal chCell As Range, ByVal toolName As String)
'���e:
'   �f�[�^�V�[�g��̖��O�A�o�[�W�����̊Ǘ����s��
'
'�p�����[�^:
'    [changedCell]  In   �ύX���ꂽ�Z��
'    [toolName]     In   �ێ�����f�[�^�V�[�g��
'
'���ӎ���:
'
    If chCell.Address = TOOL_NAME_CELL Then
        Application.EnableEvents = False
        chCell.Value = toolName
        Application.EnableEvents = True
    ElseIf chCell.Address = VERSION_CELL Then
        Application.EnableEvents = False
        chCell.Value = CURR_VERSION
        Application.EnableEvents = True
    End If
End Sub

Private Sub reverseExaminFlag(ByVal ActiveSheet As Worksheet, ByVal examinFlag As Boolean)
    '### DC�V�i���I�V�[�g���C�^�[�̍쐬 ###################
    Dim scenarioWriter As CDcScenarioSheetWriter
    Set scenarioWriter = CreateCDcScenarioSheetWriter
    scenarioWriter.Initialize ActiveSheet.Name
    scenarioWriter.AsIActionStream.Rewind
    On Error GoTo SHEET_ERROR
    '### �p�����[�^�I�u�W�F�N�g�̍쐬 #####################
    Dim paramFlag As CParamBoolean
    Set paramFlag = CreateCParamBoolean
    With paramFlag.AsIParameter
        .Name = EXAMIN_FLAG
        .AsBoolean = examinFlag
    End With
    '### �����t���O�̐؂�ւ����s #########################
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    With scenarioWriter
        Do While Not .AsIActionStream.IsEndOfCategory
            paramFlag.AsIParameter.WriteOut scenarioWriter
            .AsIActionStream.MoveNextCategory
        Loop
        .AsIParameterWriter.WriteAsBoolean DATA_CHANGED, True
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    Exit Sub
SHEET_ERROR:
    scenarioWriter.AsIParameterWriter.WriteAsBoolean DATA_CHANGED, True
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

Private Sub clearExaminMode(ByVal ActiveSheet As Worksheet)
    '### DC�V�i���I�V�[�g���C�^�[�̍쐬 ###################
    Dim scenarioWriter As CDcScenarioSheetWriter
    Set scenarioWriter = CreateCDcScenarioSheetWriter
    scenarioWriter.Initialize ActiveSheet.Name
    scenarioWriter.AsIActionStream.Rewind
    On Error GoTo SHEET_ERROR
    '### �p�����[�^�I�u�W�F�N�g�̍쐬 #####################
    Dim paramMode As CParamString
    Set paramMode = CreateCParamString
    With paramMode.AsIParameter
        .Name = EXAMIN_MODE
        .AsString = ""
    End With
    '### �������[�h�ݒ�̃N���A���s #######################
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    With scenarioWriter
        Do While Not .AsIActionStream.IsEndOfCategory
            paramMode.AsIParameter.WriteOut scenarioWriter
            .AsIActionStream.MoveNextCategory
        Loop
        .AsIParameterWriter.WriteAsBoolean DATA_CHANGED, True
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    Exit Sub
SHEET_ERROR:
    scenarioWriter.AsIParameterWriter.WriteAsBoolean DATA_CHANGED, True
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

Private Sub clearResultData(ByVal ActiveSheet As Worksheet)
    '### DC�V�i���I�V�[�g���C�^�[�̍쐬 ###################
    Dim scenarioWriter As CDcScenarioSheetWriter
    Set scenarioWriter = CreateCDcScenarioSheetWriter
    scenarioWriter.Initialize ActiveSheet.Name
    scenarioWriter.AsIActionStream.Rewind
    '### �p�����[�^�I�u�W�F�N�g�̍쐬 #####################
    Dim paramTime As CParamString
    Set paramTime = CreateCParamString
    With paramTime.AsIParameter
        .Name = EXAMIN_EXECTIME
        .AsString = ""
    End With
    Dim paramUnit As CParamString
    Set paramUnit = CreateCParamString
    With paramUnit.AsIParameter
        .Name = EXAMIN_RESULTUNIT
        .AsString = ""
    End With
    Dim paramResult As New Collection
    Dim dataIndex As Long
    For dataIndex = 0 To GetSiteCount
        paramResult.Add ""
    Next dataIndex
    '### ���茋�ʃf�[�^�̃N���A���s #######################
    Application.ScreenUpdating = False
    With scenarioWriter
        Do While Not .AsIActionStream.IsEndOfCategory
            Do While Not .AsIActionStream.IsEndOfGroup
                paramTime.AsIParameter.WriteOut scenarioWriter
                Do While Not .AsIActionStream.IsEndOfData
                    .AsIParameterWriter.WriteAsString EXAMIN_RESULT, ComposeStringList(paramResult)
                    paramUnit.AsIParameter.WriteOut scenarioWriter
                    .AsIActionStream.MoveNextData
                Loop
                .AsIActionStream.MoveNextGroup
            Loop
            .AsIActionStream.MoveNextCategory
        Loop
    End With
    Application.ScreenUpdating = True
End Sub

Private Sub clearValidationMode(ByVal ActiveSheet As Worksheet)
    '### DC�V�i���I�V�[�g���C�^�[�̍쐬 ###################
    Dim scenarioWriter As CDcScenarioSheetWriter
    Set scenarioWriter = CreateCDcScenarioSheetWriter
    scenarioWriter.Initialize ActiveSheet.Name
    '### �p�����[�^�I�u�W�F�N�g�̍쐬 #####################
    Dim paramMode As CParamBoolean
    Set paramMode = CreateCParamBoolean
    '### �����W�o���f�[�V�������[�h�ݒ�̃N���A���s #######
    With paramMode.AsIParameter
        .Name = IS_VALIDATE
        .AsBoolean = False
        .WriteOut scenarioWriter
    End With
End Sub

Public Sub ValidateDCTestSenario()
'���e:
'   DC�V�i���I�V�[�g�t�H�[�}�b�g�𐮗�����}�N���֐�
'
'�p�����[�^:
'
'���ӎ���:
'   �@�Z���̃O���[�s���O����
'   �A�A�N�V�����O���[�v�t�H�[�}�b�g���`
'   �B�e�p�����[�^�̃`�F�b�N
'   �����s����
'
    On Error GoTo ErrHandler
    '### �A�N�e�B�u�f�[�^�V�[�g�̎擾 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    '### ���[�N�V�[�g���C�^�[�̏��� #######################
    Dim scenarioWriter As CDcScenarioSheetWriter
    Set scenarioWriter = CreateCDcScenarioSheetWriter
    Application.ScreenUpdating = False
    '### �t�H�[�}�b�g���`���s #############################
    With scenarioWriter
        .Initialize DataSheet.Name
        .SetGrouping
        .Validate
    End With
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub SetSheetBackground(ByVal activeSht As Object)
    '### �A�N�e�B�u�f�[�^�V�[�g�̎擾 #####################
    Dim DataSheet As Worksheet
    Dim toolName As String
    toolName = activeSht.Range("B1").Value
    On Error Resume Next
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, toolName)
    If Err.Number = 9999 Then
        activeSht.SetBackgroundPicture fileName:=""
        Exit Sub
    End If
    If DataSheet Is Nothing Then
        activeSht.SetBackgroundPicture fileName:= _
            GetJobRootPath & "\bin\DT_NotInJob.gif"
        Exit Sub
    End If
    If DataSheet.Name = activeSht.Name Then
        activeSht.SetBackgroundPicture fileName:=""
    Else
        activeSht.SetBackgroundPicture fileName:= _
            GetJobRootPath & "\bin\DT_NotInJob.gif"
    End If
End Sub

Public Function GetUsableDataSht(ByVal ctrlShName As String, ByVal toolName As String) As Worksheet
    '### �A�N�e�B�u�f�[�^�V�[�g�I�u�W�F�N�g�̍쐬 #########
    Dim ctrlSheet As CDataSheetManager
    Set ctrlSheet = CreateCDataSheetManager
    ctrlSheet.Initialize ctrlShName
    Set GetUsableDataSht = ctrlSheet.GetActiveDataSht(toolName)
End Function


