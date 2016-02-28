VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEeeAuto_TesterChannelChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �R���f�B�V�����`�F�b�J�@�d�����擾���W���[��
'
'�ړI:
'   ���̂Ƃ��̏����ݒ�𑕒u����擾����
'   �X�i�b�v�V���b�g�̂܂�p�N��
'�쐬��:
'   �ێR
'   �L�� �ҏW Ver1.1 2013/02/25

Option Explicit

'Tool�Ή���ɃR�����g�O���Ď��������ɂ���B�@2013/03/07 H.Arikawa
#Const HDVIS_USE = 0      'HVDIS���\�[�X�̎g�p�@ 0�F���g�p�A0�ȊO�F�g�p  <PALS��EeeAuto�����Ŏg�p>

'SnapShot�w�b�_���p�\����
Private Type type_SS_HEADER
    tResourceName() As String
    tIdLabel() As String
    tPinName() As String
    tSiteNumber() As Long
    tChannelNumber() As Long
End Type

'APMU���擾�p�\����
Private Type Type_APMU
    tGate As Long
    tRelay As Long
    tLowPassFilter As Long
    tExternalSense As Long
    tAlarm As Long
    tMode As ApmuMode
    tClampValue As Double
    tForceValue As Double
    tIRange As ApmuIRange
    tVRange As ApmuVRange
    tGangPinFlag As Long
'    tMeasureResult As Double
End Type

'APMU�擾���g�pTMP�\����
Private Type Type_TMP_APMU
    tGate() As Long
    tRelay() As Long
    tLowPassFilter() As Long
    tExternalSense() As Long
    tAlarm() As Long
    tMode() As ApmuMode
    tClampValue() As Double
    tForceValue() As Double
    tIRange() As ApmuIRange
    tVRange() As ApmuVRange
    tGangPinFlag() As Long
'    tMeasureResult() As Double
End Type

'DPS���擾�p�\����
Private Type type_DPS
    tCurrentLimit As Double
    tPrimaryVoltage As Double
    tAlternateVoltage As Double
    tCurrentRange As DpsIRange
    tOutputSource As Long
    tForceRelay As String
    tSenseRelay As String
'    tMeasureResult As Double
    tGangPinFlag As Long
    tMeasureSamples As Long
End Type

'PPMU���擾�p�\����
Private Type type_PPMU_INFO
    tForceVoltage As Double
    tForceCurrent As Double
    tCurrentRange As Long
    tHighLimit As Double
    tLowLimit As Double
    tConnect As Boolean
    tForceType As String
'    tMeasureResult As Double
    tMeasureSamples As Long
End Type

'BPMU���擾�p�\����
Private Type type_BPMU_INFO
    tClampCurrent As Double
    tClampVoltage As Double
    tForceCurrent As Double
    tForceVoltage As Double
    tVoltageRange As Long
    tCurrentRange As Long
    tHighLimit As Double
    tLowLimit As Double
    tVoltmeterMode As Boolean
    tBpmuGate As Boolean
    tConnectDut As Boolean
    tForcingMode As String
    tMeasureMode As String
'    tMeasureResult As Double
End Type

'Digtal-ch(PE)���擾�p
Private Type type_PE_INFO
    tVDriveLo As Double
    tVDriveHi As Double
    tVCompareLo As Double
    tVCompareHi As Double
    tVClampLo As Double
    tVClampHi As Double
    tVThreshold As Double
    tISource As Double
    tISink As Double
    tPeConnect As Boolean
    tHvConnect As Boolean
    tPpmuConnect As Boolean
    tBpmuConnect As Boolean
    tD0 As Double
    tD1 As Double
    tD2 As Double
    tD3 As Double
    tR0 As Double
    tR1 As Double
    tHvVph As Double
    tHvIph As Double
    tHvTpr As Double
End Type

'HDVIS���擾�p
Private Type type_HDVIS
    tGate As Long
    tRelay As Long
    tLowPassFilter As Long
    tAlarmOpenDgs As Long
    tAlarmOverLoad As Long
    tMargePinFlag As Long
    tMode As Long
    tVRange As Long
    tIRange As Long
    tSlewRate As Long
    tRelayMode As Long
    tClampValue As Double
    tForceValue As Double
    tMeasureResult As Double
    tExtMode As Long
    tExtSendRelay As Long
    tExtTriggerRelay As Long
    tSetupEnable As Boolean
End Type

'APMU���p�\����
Private Type type_APMU_INFO
    tSsHeader As type_SS_HEADER
    tApmuinf() As Type_APMU
End Type

'HDVIS���p�\����
Private Type type_HDVIS_INFO
    tSsHeader As type_SS_HEADER
    tHdvisInf() As type_HDVIS
End Type

'DPS���p�\����
Private Type type_DPS_INFO
    tSsHeader As type_SS_HEADER
    tDpsinf() As type_DPS
End Type

'I/O Pin(�f�W�^���s��) ���p�\����
Private Type type_IO_INFO
    tSsHeader As type_SS_HEADER
    tPeinf() As type_PE_INFO
    tPpmuinf() As type_PPMU_INFO
    tBpmuinf() As type_BPMU_INFO
End Type

'���̃��W���[���Ǝ��̍\����
Private Type type_ALL_Resouces
    tApmuInfo As type_APMU_INFO
    tDchanInfo As type_IO_INFO
    tDpsInfo As type_DPS_INFO
    tHdvisInfo As type_HDVIS_INFO
End Type
Implements IEeeAuto_TestConditionChecker

'���[�h�̒�`
Private m_strMode As String

Private m_TesterConditions(1) As type_ALL_Resouces

Private Sub Class_Initialize()
    
    m_strMode = ""

End Sub

Private Sub IEeeAuto_TestConditionChecker_CheckBeforeCondition()

'���e:
'   �d���Q�ƃe�[�u���̏�����
'
'���ӎ���:

    m_TesterConditions(0) = GetTesterResouceConditionInfo

End Sub

Private Function IEeeAuto_TestConditionChecker_CheckAfterCondition() As Boolean

    'Default��Valid
    IEeeAuto_TestConditionChecker_CheckAfterCondition = True
    
    m_TesterConditions(1) = GetTesterResouceConditionInfo
        
    If (IsEqualTwoCondition(m_TesterConditions(0), m_TesterConditions(1))) Then
        '�������ꏏ�������ꍇ InValid�Ƃ���
        IEeeAuto_TestConditionChecker_CheckAfterCondition = False
    End If

End Function

Private Sub IEeeAuto_TestConditionChecker_SetOperationMode(ByVal strMode As String)
    
    m_strMode = strMode
    
End Sub

Private Sub IEeeAuto_TestConditionChecker_SetEndPosition(ByVal strPos As String)

    '�ŏI�ʒu�͂��̃N���X�ɂ͂Ȃ�

End Sub
'=================================================================


Private Function GetTesterResouceConditionInfo() As type_ALL_Resouces
'���e:
'   �e�X�^���牞�����擾�������̂��̂�������
'
'���ӎ���:
'   �Ƃ��ɂȂ�

    Dim tPinNameArr() As String
    Dim tChansArr() As Long
    Dim resourceChk As Boolean
    Dim ResourceName As String

    Dim tApmuinf As type_APMU_INFO 'APMU���p�\����
    Dim tDpsinf As type_DPS_INFO 'DPS���p�\����
    Dim tPeinf As type_IO_INFO 'PE���p�\����
        
    '=====APMU======================================
    'APMU���\�[�X���g�p���Ă���Channel�𒲂ׂ�
    resourceChk = mf_ChkResourcePin_Site0(chAPMU, tChansArr, tPinNameArr)
    ResourceName = "[APMU]" 'IP750���\�[�X���ʗp���x��[APMU]
                                                                                                
    'APMU���\�[�X���g�p���Ă�����l���Ƃ�
    If resourceChk = True Then
        'APMU�����e���_�C��API���g�p���Ď擾
        Call mf_Make_SsHeader_Site0(ResourceName, "*", tPinNameArr, tChansArr, tApmuinf.tSsHeader)
        Call mf_GetApmuInfo(tApmuinf.tSsHeader.tChannelNumber, tApmuinf)
        GetTesterResouceConditionInfo.tApmuInfo = tApmuinf
    End If
    Erase tPinNameArr
    Erase tChansArr
    
    
    '=====DPS======================================
    'DPS���\�[�X���g�p���Ă���Channel�𒲂ׂ�
    resourceChk = mf_ChkResourcePin_Site0(chDPS, tChansArr, tPinNameArr)
    ResourceName = "[DPS]" 'IP750���\�[�X���ʗp���x��[DPS]
                                                                                                
    'DPS���\�[�X���g�p���Ă�����l���Ƃ�
    If resourceChk = True Then
        'DPS���̃w�b�_���̏����쐬
        Call mf_Make_SsHeader_Site0(ResourceName, "*", tPinNameArr, tChansArr, tDpsinf.tSsHeader)
        'DPS����TERADYNE-API����擾
        Call mf_GetDpsInfo(tDpsinf.tSsHeader.tChannelNumber, tDpsinf)
        GetTesterResouceConditionInfo.tDpsInfo = tDpsinf
    End If
    Erase tPinNameArr
    Erase tChansArr
    
    
    'DChan�����Ȃ��Ƃ��͖���
    If m_strMode <> EEE_AUTO_TESTER_CHECKER_WITHOUT_DCHAN Then
    
        '=====Dchan(IO)======================================
        'I/O(PE)���\�[�X���g�p���Ă���Channel�𒲂ׂ�
        resourceChk = mf_ChkResourcePin_Site0(chIO, tChansArr, tPinNameArr)
        ResourceName = "[PE]" 'IP750���\�[�X���ʗp���x��[PE]
    
        'I/O(PE)���\�[�X���g�p���Ă�����l���Ƃ�
        If resourceChk = True Then
            'PE���̃w�b�_���̏����쐬
            Call mf_Make_SsHeader_Site0(ResourceName, "*", tPinNameArr, tChansArr, tPeinf.tSsHeader)
            'PE����TERADYNE-API����擾
            Call mf_GetPeInfo(tPeinf.tSsHeader.tChannelNumber, tPeinf.tPeinf)
            'PPMU����TERADYNE-API����擾
            Call mf_GetPpmuInfo(tPeinf.tSsHeader.tChannelNumber, tPeinf.tPpmuinf)
            'BPMU����TERADYNE-API����擾
            Call mf_GetBpmuInfo(tPeinf.tSsHeader.tChannelNumber, tPeinf.tSsHeader.tPinName, tPeinf.tSsHeader.tSiteNumber, tPeinf.tBpmuinf)
            GetTesterResouceConditionInfo.tDchanInfo = tPeinf
        End If
        Erase tPinNameArr
        Erase tChansArr
        
    End If
   
#If HDVIS_USE <> 0 Then

    Const CH_HDVIS = 36
    ResourceName = "[HDVIS]" 'IP750���\�[�X���ʗp���x��[HDVIS]

    '=====HDVIS======================================
    'HDVIS���\�[�X���g�p���Ă���Channel�𒲂ׂ�
    resourceChk = mf_ChkResourcePin(CH_HDVIS, tChansArr, tPinNameArr)

    'HDVIS���\�[�X���g�p���Ă��Ȃ��Ƃ��͏I��
    If resourceChk = False Then
        'HDVIS���̃w�b�_���̏����쐬
        Call mf_Make_SsHeader(ResourceName, testIdLabel, tPinNameArr, tChansArr, tHdvisInf.tSsHeader)
    End If
    MsgBox "No Debug!! Don't Use HDVIS Condition Auto omission"
#End If
                                                                                                     
End Function

Private Function IsEqualTwoCondition(ByRef RHS As type_ALL_Resouces, ByRef lhs As type_ALL_Resouces) As Boolean
'���e:
'   type_ALL_Resouces�\���̂� operator =
'
'���ӎ���:
'   ��r���e�͂ЂƂ܂�����ŁB

    IsEqualTwoCondition = False
    
    Dim i As Long

    For i = 0 To UBound(RHS.tApmuInfo.tApmuinf)
        If RHS.tApmuInfo.tApmuinf(i).tClampValue <> lhs.tApmuInfo.tApmuinf(i).tClampValue Then
            Exit Function
        End If
        If RHS.tApmuInfo.tApmuinf(i).tForceValue <> lhs.tApmuInfo.tApmuinf(i).tForceValue Then
            Exit Function
        End If
        If RHS.tApmuInfo.tApmuinf(i).tIRange <> lhs.tApmuInfo.tApmuinf(i).tIRange Then
            Exit Function
        End If
        If RHS.tApmuInfo.tApmuinf(i).tGangPinFlag <> lhs.tApmuInfo.tApmuinf(i).tGangPinFlag Then
            Exit Function
        End If
        If RHS.tApmuInfo.tApmuinf(i).tRelay <> lhs.tApmuInfo.tApmuinf(i).tRelay Then
            Exit Function
        End If
    Next i

    For i = 0 To UBound(RHS.tDpsInfo.tDpsinf)
        If RHS.tDpsInfo.tDpsinf(i).tPrimaryVoltage <> lhs.tDpsInfo.tDpsinf(i).tPrimaryVoltage Then
            Exit Function
        End If
        If RHS.tDpsInfo.tDpsinf(i).tForceRelay <> lhs.tDpsInfo.tDpsinf(i).tForceRelay Then
            Exit Function
        End If
    Next i

    'DChan�����Ȃ��Ƃ��͖���
    If m_strMode <> EEE_AUTO_TESTER_CHECKER_WITHOUT_DCHAN Then
        For i = 0 To UBound(RHS.tDchanInfo.tBpmuinf)
            With RHS.tDchanInfo.tBpmuinf(i)
                If .tBpmuGate <> lhs.tDchanInfo.tBpmuinf(i).tBpmuGate Then
                    Exit Function
                End If
            End With
        Next
        
        For i = 0 To UBound(RHS.tDchanInfo.tPeinf)
            With RHS.tDchanInfo.tPeinf(i)
                If .tPpmuConnect <> lhs.tDchanInfo.tPeinf(i).tPpmuConnect Then
                    Exit Function
                End If
                If .tBpmuConnect <> lhs.tDchanInfo.tPeinf(i).tBpmuConnect Then
                    Exit Function
                End If
                If .tPeConnect <> lhs.tDchanInfo.tPeinf(i).tPeConnect Then
                    Exit Function
                End If
                If .tVDriveHi <> lhs.tDchanInfo.tPeinf(i).tVDriveHi Then
                    Exit Function
                End If
                If .tVDriveLo <> lhs.tDchanInfo.tPeinf(i).tVDriveLo Then
                    Exit Function
                End If
            End With
        Next
        
        For i = 0 To UBound(RHS.tDchanInfo.tPpmuinf)
            With RHS.tDchanInfo.tPpmuinf(i)
                If .tForceType <> lhs.tDchanInfo.tPpmuinf(i).tForceType Then
                    Exit Function
                End If
                If .tForceVoltage <> lhs.tDchanInfo.tPpmuinf(i).tForceVoltage Then
                    Exit Function
                End If
                If .tForceCurrent <> lhs.tDchanInfo.tPpmuinf(i).tForceCurrent Then
                    Exit Function
                End If
                If .tCurrentRange <> lhs.tDchanInfo.tPpmuinf(i).tCurrentRange Then
                    Exit Function
                End If
            End With
        Next
    End If

    IsEqualTwoCondition = True '�Ō�܂ł����炨�߂łƂ�
    
End Function


'�w�胊�\�[�X���g�p���Ă���`�����l����PinName�𒲂ׂ� Site0�̂ݔ�
Private Function mf_ChkResourcePin_Site0(ByVal ResourceName As chtype, _
ByRef rChansArr() As Long, _
ByRef rPinNameArr() As String) As Boolean

    Dim rPinCnt As Long
    Dim rChCnt As Long
    Dim rSiteCnt As Long
    Dim rAllPinsStr As String
    Dim funcName As String
    
    funcName = "@mf_ChkResourcePin"

    '�w�胊�\�[�X���g�p���Ă���PIN�����擾
    Call TheExec.DataManager.GetPinNames(rPinNameArr, ResourceName, rPinCnt)
                                                   
    '�w�肳�ꂽ��\�[�X���A��`����Ă��Ȃ��Ƃ���False��Ԃ��ďI��
    If rPinCnt = 0 Then
        mf_ChkResourcePin_Site0 = False
        Exit Function
    End If
                                                
    '�w�胊�\�[�X�Ƃ��Ē�`����Ă��邷�ׂĂ�PIN�̖��O���J���}��؂�ō쐬�@�@("P_PIN1,P_PIN2, .....")
    rAllPinsStr = mf_Make_PinNameStr(rPinNameArr)
                    
    '�w�胊�\�[�X�Ƃ��Ē�`����Ă��邷�ׂĂ�PIN�̃`�����l���ԍ����擾
    Call TheExec.DataManager.GetChanList(rAllPinsStr, 0, ResourceName, _
    rChansArr, rChCnt, rSiteCnt, "Resource Pin Check Error" & funcName)

    mf_ChkResourcePin_Site0 = True

End Function

'�X�i�b�v�V���b�g�̃w�b�_�����쐬
Private Sub mf_Make_SsHeader_Site0(ByVal ResourceName As String, ByVal tstIdLabel As String, _
ByRef pinNameArr() As String, ByRef chansArr() As Long, ByRef ssHeaderInf As type_SS_HEADER)
    
    Dim PinCnt As Long
    Dim siteNumCnt As Long
    Dim lopCnt As Long

    '�w�b�_�p�̍\���̂̔�������
    With ssHeaderInf
        ReDim .tResourceName(UBound(chansArr))
        ReDim .tIdLabel(UBound(chansArr))
        ReDim .tPinName(UBound(chansArr))
        ReDim .tSiteNumber(UBound(chansArr))
    End With
    
    '�`�����l���ԍ��z������炤
    ssHeaderInf.tChannelNumber = chansArr
    
    lopCnt = 0
    
    '�T�C�gNO�ƃs�����A���\�[�X���́A���x�����̍쐬
    For PinCnt = 0 To UBound(pinNameArr) '�Ώۃs�� LOOP
            
        With ssHeaderInf
            .tResourceName(lopCnt) = ResourceName
            .tIdLabel(lopCnt) = tstIdLabel
            .tPinName(lopCnt) = pinNameArr(PinCnt)
            .tSiteNumber(lopCnt) = 0 'Site0�Œ�
        End With
                   
        lopCnt = lopCnt + 1
        
    Next PinCnt

End Sub



'=================================================================
'������艺��XLibSnapshotIP750Mod ���̂܂�
'=================================================================

'�w�胊�\�[�X���g�p���Ă���`�����l����PinName�𒲂ׂ�
Private Function mf_ChkResourcePin(ByVal ResourceName As chtype, _
ByRef rChansArr() As Long, _
ByRef rPinNameArr() As String) As Boolean

    Dim rPinCnt As Long
    Dim rChCnt As Long
    Dim rSiteCnt As Long
    Dim rAllPinsStr As String
    Dim funcName As String
    
    funcName = "@mf_ChkResourcePin"

    '�w�胊�\�[�X���g�p���Ă���PIN�����擾
    Call TheExec.DataManager.GetPinNames(rPinNameArr, ResourceName, rPinCnt)
                                                   
    '�w�肳�ꂽ��\�[�X���A��`����Ă��Ȃ��Ƃ���False��Ԃ��ďI��
    If rPinCnt = 0 Then
        mf_ChkResourcePin = False
        Exit Function
    End If
                                                
    '�w�胊�\�[�X�Ƃ��Ē�`����Ă��邷�ׂĂ�PIN�̖��O���J���}��؂�ō쐬�@�@("P_PIN1,P_PIN2, .....")
    rAllPinsStr = mf_Make_PinNameStr(rPinNameArr)
                    
    '�w�胊�\�[�X�Ƃ��Ē�`����Ă��邷�ׂĂ�PIN�̃`�����l���ԍ����擾
    Call TheExec.DataManager.GetChanList(rAllPinsStr, -1, ResourceName, _
    rChansArr, rChCnt, rSiteCnt, "Resource Pin Check Error" & funcName)

    mf_ChkResourcePin = True

End Function

'��������z��Ɋi�[����Ă���v�f�̖��O���A�J���}��؂�`���ō쐬
Private Function mf_Make_PinNameStr(ByRef pinNameArr() As String) As String

    Dim tLoopCnt As Long
        
    '�z��Ɋi�[����Ă��邷�ׂĂ�PIN�̖��O���A�J���}��؂�`���ō쐬�@�@("P_PIN1,P_PIN2, .....")
    mf_Make_PinNameStr = pinNameArr(0)
    
    For tLoopCnt = 1 To UBound(pinNameArr)
        mf_Make_PinNameStr = mf_Make_PinNameStr & "," & pinNameArr(tLoopCnt)
    Next tLoopCnt

End Function

Private Sub mf_GetApmuInfo(ByRef apmuChans() As Long, ByRef typeApmuInf As type_APMU_INFO)
'���e:
'   �w��CH��APMU�����擾���܂�
'
'[apmuChans]      In  �����擾������APMU��CH
'[typeApmuInf]    Out  �擾���ʊi�[�pAPMU�\����
'
    Call mf_MakeApmuInfo(apmuChans, typeApmuInf)

End Sub

Private Sub mf_GetDpsInfo(ByRef dpsChans() As Long, ByRef typeDpsInf As type_DPS_INFO)
'���e:
'   �w��CH��DPS�����擾���܂�
'
'[dpsChans]       In  �����擾������DPS��CH
'[typeDpsInf]     Out  �擾���ʊi�[�pDPS�\����
'
    Call mf_MakeDpsInfo(dpsChans, typeDpsInf)

End Sub

Private Sub mf_GetPeInfo(ByRef peChans() As Long, ByRef typePeInf() As type_PE_INFO)
'���e:
'   �w��CH��PE�����擾���܂�
'
'[peChans]        In   �����擾������I/O(PE)��CH
'[typePeInf]      Out  �擾���ʊi�[�pPE�\����
'
    Call mf_MakePeInfo(peChans, typePeInf)

End Sub

Private Sub mf_GetPpmuInfo(ByRef ppmuChans() As Long, ByRef typePpmuInf() As type_PPMU_INFO)
'���e:
'   �w��CH��PPMU�����擾���܂�
'
'[ppmuChans]        In   �����擾������I/O(PPMU)��CH
'[typePpmuInf]      Out  �擾���ʊi�[�pPPMU�\����
'
    Call mf_MakePpmuInfo(ppmuChans, typePpmuInf)

End Sub

Private Sub mf_GetBpmuInfo(ByRef bpmuChans() As Long, _
ByRef bpmuPins() As String, _
ByRef siteNum() As Long, _
ByRef typeBpmuInf() As type_BPMU_INFO)
'���e:
'   �w��CH��PPMU�����擾���܂�
'
'[bpmuChans]        In   �����擾������I/O(BPMU)��CH
'[bpmuPins]         In   �����擾������I/O(BPMU)�̃s����
'[siteNum]          In   �����擾������I/O(BPMU)�̃T�C�g�ԍ�
'[typeBpmuInf]      Out  �擾���ʊi�[�pBPMU�\����
'
    Call mf_MakeBpmuInfo(bpmuChans, bpmuPins, siteNum, typeBpmuInf)

End Sub

Private Sub mf_GetHdvisInfo(ByRef hdvischans() As Long, ByRef typeHdvisInf As type_HDVIS_INFO)
'���e:
'   �w��CH��HDVIS�����擾���܂�
'
'[hdvisChans]      In  �����擾������HDVIS��CH
'[typeHdvisInf]    Out �擾���ʊi�[�pHDVIS�\����
'
    Call mf_MakeHdvisInfo(hdvischans, typeHdvisInf)

End Sub

'�w��CH��APMU����TERADYNE-API����擾
Private Sub mf_MakeApmuInfo(ByRef apmuChans() As Long, ByRef typeApmuInf As type_APMU_INFO)
    
    Dim tmpApmuMode() As ApmuMode
    Dim tmpForceValue() As Double
    Dim tmpClampValue() As Double
    Dim myRetVrange() As ApmuVRange
    Dim myRetIrange() As ApmuIRange
    Dim tchCnt As Long
    Dim read_apmu As Boolean
    
    Dim tmpApmuInf As Type_TMP_APMU
        

    '�w��CH��APMU���\�[�X�󋵂��擾
    
    With tmpApmuInf
        .tGate = TheHdw.APMU.Chans(apmuChans).Gate
        .tRelay = TheHdw.APMU.Chans(apmuChans).relay
        .tLowPassFilter = TheHdw.APMU.Chans(apmuChans).LowPassFilter
        .tExternalSense = TheHdw.APMU.Chans(apmuChans).ExternalSense
        .tAlarm = TheHdw.APMU.Chans(apmuChans).alarm
    End With

    '�w��CH��APMU���[�h�����擾�B
    With tmpApmuInf
        TheHdw.APMU.Chans(apmuChans).ReadRangesAndMode .tMode, .tVRange, .tIRange
    End With
    
    '�w��CH��APMU���擾�p�̍\���̂̔�������
    With tmpApmuInf
        ReDim .tForceValue(UBound(apmuChans))
        ReDim .tClampValue(UBound(apmuChans))
    End With
                
    '�w��CH�̃M�����O�ڑ��̊m�F�ƁA���[�^�̓ǂݎ��l�̎擾
    With tmpApmuInf
        .tGangPinFlag = TheHdw.APMU.Chans(apmuChans).GangedChannels     '�M�����O�ڑ���Ԃ̊m�F
'        Call TheHdw.APMU.chans(apmuChans).measure(1, .tMeasureResult)   '���[�^�ǂݎ��l���擾
    End With
                                
    '�w��CH��APMU���[�h�ʔ����Force�AClamp�̒l�̎擾
    'APMU�X�i�b�v�V���b�g�\���̂Ɍ��ʋl�ߍ���
    ReDim typeApmuInf.tApmuinf(UBound(apmuChans))
    
    For tchCnt = 0 To UBound(apmuChans) '�Ώ�CH LOOP
    'APMU���[�h�ɂ��킹�āA�����W�Ɛݒ�l���擾
        Select Case tmpApmuInf.tMode(tchCnt)
            Case apmuForceIMeasureV:
                read_apmu = TheHdw.APMU.Chans(apmuChans(tchCnt)).ReadForceCurrents(myRetIrange, tmpForceValue)
                read_apmu = TheHdw.APMU.Chans(apmuChans(tchCnt)).ReadClampVoltages(myRetVrange, tmpClampValue)
            Case apmuForceVMeasureI:
                read_apmu = TheHdw.APMU.Chans(apmuChans(tchCnt)).ReadForceVoltages(myRetVrange, tmpForceValue)
                read_apmu = TheHdw.APMU.Chans(apmuChans(tchCnt)).ReadClampCurrents(myRetIrange, tmpClampValue)
            Case apmuMeasureV:
                read_apmu = TheHdw.APMU.Chans(apmuChans(tchCnt)).ReadClampVoltages(myRetVrange, tmpClampValue)
                ReDim tmpForceValue(UBound(apmuChans))
            End Select
            
        'APMU���\���̂ɁACH LOOP�Ŏ擾���ʋl�ߍ���
        With typeApmuInf.tApmuinf(tchCnt)
            .tAlarm = tmpApmuInf.tAlarm(tchCnt)
            .tClampValue = tmpClampValue(0)
            .tExternalSense = tmpApmuInf.tExternalSense(tchCnt)
            .tForceValue = tmpForceValue(0)
            .tGangPinFlag = tmpApmuInf.tGangPinFlag(tchCnt)
            .tGate = tmpApmuInf.tGate(tchCnt)
            .tIRange = tmpApmuInf.tIRange(tchCnt)
            .tLowPassFilter = tmpApmuInf.tLowPassFilter(tchCnt)
'            .tMeasureResult = tmpApmuInf.tMeasureResult(tchCnt)
            .tMode = tmpApmuInf.tMode(tchCnt)
            .tRelay = tmpApmuInf.tRelay(tchCnt)
            .tVRange = tmpApmuInf.tVRange(tchCnt)
        End With
    
    Next tchCnt
    
End Sub

'�w��CH��DPS����TERADYNE-API����擾
Private Sub mf_MakeDpsInfo(ByRef dpsChans() As Long, ByRef typeDpsInf As type_DPS_INFO)
    
    Dim tchCnt As Long
    Dim tmpMesureVal() As Double
    Dim tmpCurrentLimit As Variant
    Dim tmpPrimaryVoltage As Variant
    Dim tmpAlternateVoltage As Variant
    Dim aveCnt As Long

    'DPS���̔�������
    ReDim typeDpsInf.tDpsinf(UBound(dpsChans))
    
    For tchCnt = 0 To UBound(dpsChans) Step 1
        
        With typeDpsInf.tDpsinf(tchCnt)
            '���\�[�X�ݒ��Ԏ擾
            tmpCurrentLimit = TheHdw.DPS.Chans(dpsChans(tchCnt)).CurrentLimit
            .tCurrentLimit = tmpCurrentLimit(0)
            
            tmpPrimaryVoltage = TheHdw.DPS.Chans(dpsChans(tchCnt)).forceValue(dpsPrimaryVoltage)
            .tPrimaryVoltage = tmpPrimaryVoltage(0)
                        
            tmpAlternateVoltage = TheHdw.DPS.Chans(dpsChans(tchCnt)).forceValue(dpsAlternateVoltage)
            .tAlternateVoltage = tmpAlternateVoltage(0)
            
            .tCurrentRange = TheHdw.DPS.Chans(dpsChans(tchCnt)).CurrentRange
            .tOutputSource = TheHdw.DPS.Chans(dpsChans(tchCnt)).OutputSource

            '�����[�ڑ���Ԏ擾
            If TheHdw.DPS.Chans(dpsChans(tchCnt)).ForceRelayClosed = True Then
                .tForceRelay = "Closed"
            Else
                .tForceRelay = "Open"
            End If

            If TheHdw.DPS.Chans(dpsChans(tchCnt)).SenseRelayClosed = True Then
                .tSenseRelay = "Closed"
            Else
                .tSenseRelay = "Open"
            End If
                
            '���[�^�[�̃A�x���[�W���擾
            .tMeasureSamples = TheHdw.DPS.Samples
            
'            '�d���v�̓d���l�擾
'            Call TheHdw.DPS.chans(dpsChans(tchCnt)).MeasureCurrents(.tCurrentRange, tmpMesureVal)
'
'            .tMeasureResult = 0
'
'            For aveCnt = 0 To UBound(tmpMesureVal)
'                .tMeasureResult = .tMeasureResult + tmpMesureVal(aveCnt)
'            Next aveCnt
'
'            .tMeasureResult = .tMeasureResult / .tMeasureSamples
        
        End With
                                                              
    Next tchCnt
    
End Sub

'�w��CH��PE����TERADYNE-API����擾
Private Sub mf_MakePeInfo(ByRef peChans() As Long, ByRef typePeInf() As type_PE_INFO)
    
    Dim tchCnt As Long
    
    ReDim typePeInf(UBound(peChans))
    
    For tchCnt = 0 To UBound(peChans) Step 1
        
        With typePeInf(tchCnt)
            .tVDriveLo = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVDriveLo)
            .tVDriveHi = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVDriveHi)
            .tVClampLo = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVCL)
            .tVClampHi = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVCH)
            .tVCompareLo = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVCompareLo)
            .tVCompareHi = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVCompareHi)
            .tVThreshold = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVT)
            .tISource = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chISource)
            .tISink = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chISink)
                                                                  
            'High Voltage Status
            If (peChans(tchCnt) Mod 32) = 0 Or ((peChans(tchCnt) - 4) Mod 32) = 0 Then
                Call TheHdw.PinLevels.chan(peChans(tchCnt)).ReadHighVoltageParams(.tHvVph, .tHvIph, .tHvTpr)
            End If
                        
'{
            'Rdge Set�̐ݒ�l��������ɂ���΂悢�̂�����Ȃ��̂ŕ���
'            Call TheHdw.Digital.Timing.chan(peChans(tchCnt)).readEdgeTimingRAM(0)
'            .tD0 = TheHdw.Digital.Timing.EdgeTime(chEdgeD0)
'            .tD1 = TheHdw.Digital.Timing.EdgeTime(chEdgeD1)
'            .tD2 = TheHdw.Digital.Timing.EdgeTime(chEdgeD2)
'            .tD3 = TheHdw.Digital.Timing.EdgeTime(chEdgeD3)
'            .tR0 = TheHdw.Digital.Timing.EdgeTime(chEdgeR0)
'            .tR1 = TheHdw.Digital.Timing.EdgeTime(chEdgeR1)
'}
                                                                  
            'IO�s��(HV)�̃����[�ڑ���Ԋm�F
             .tHvConnect = mf_ChkIoRelayStat(peChans(tchCnt), "HV")
            'IO�s��(PE)�̃����[�ڑ���Ԋm�F
             .tPeConnect = mf_ChkIoRelayStat(peChans(tchCnt), "PE")
            'IO�s��(PPMU)�̃����[�ڑ���Ԋm�F
             .tPpmuConnect = mf_ChkIoRelayStat(peChans(tchCnt), "PPMU")
            'IO�s��(BPMU)�̃����[�ڑ���Ԋm�F
             .tBpmuConnect = mf_ChkIoRelayStat(peChans(tchCnt), "BPMU")
        
        End With

    Next tchCnt
    
End Sub

'�w��CH��PPMU����TERADYNE-API����擾
Private Sub mf_MakePpmuInfo(ByRef ppmuChans() As Long, ByRef typePpmuInf() As type_PPMU_INFO)
    Dim tchCnt As Long
    Dim tmpMeasureVal() As Double
    Dim aveCnt As Long
    
    ReDim typePpmuInf(UBound(ppmuChans))
    
    For tchCnt = 0 To UBound(ppmuChans) Step 1
        
        With typePpmuInf(tchCnt)
            .tCurrentRange = TheHdw.PPMU.Chans(ppmuChans(tchCnt)).CurrentRange
            .tForceVoltage = TheHdw.PPMU.Chans(ppmuChans(tchCnt)).ForceVoltage(.tCurrentRange)
            .tForceCurrent = TheHdw.PPMU.Chans(ppmuChans(tchCnt)).ForceCurrent(.tCurrentRange)
            .tHighLimit = TheHdw.PPMU.Chans(ppmuChans(tchCnt)).TestLimitHigh
            .tLowLimit = TheHdw.PPMU.Chans(ppmuChans(tchCnt)).TestLimitLow
                                       
            If TheHdw.PPMU.Chans(ppmuChans(tchCnt)).IsForcingVoltage <> True Then
                .tForceType = "AMPS"
'                Call TheHdw.PPMU.chans(ppmuChans(tchCnt)).MeasureVoltages(tmpMeasureVal)
            Else
                .tForceType = "VOLTS"
'                Call TheHdw.PPMU.chans(ppmuChans(tchCnt)).MeasureCurrents(tmpMeasureVal)
            End If
                                                                                          
'            .tMeasureResult = 0
'            .tMeasureSamples = UBound(tmpMeasureVal) + 1
'
'            For aveCnt = 0 To UBound(tmpMeasureVal)
'                .tMeasureResult = .tMeasureResult + tmpMeasureVal(aveCnt)
'            Next aveCnt
'
'            .tMeasureResult = .tMeasureResult / .tMeasureSamples
                                                                                                                                                            
            'IO�s���̃����[�ڑ���Ԋm�F
             .tConnect = mf_ChkIoRelayStat(ppmuChans(tchCnt), "PPMU")
                            
        End With

    Next tchCnt
    
End Sub

'�w��CH��BPMU����TERADYNE-API����擾
Private Sub mf_MakeBpmuInfo(ByRef bpmuChans() As Long, _
ByRef bpmuPins() As String, _
ByRef siteNum() As Long, _
ByRef typeBpmuInf() As type_BPMU_INFO)

    Dim tchCnt As Long
    Dim tmpIrange() As Long
    Dim tmpVrange() As Long
    Dim tmpFvMode As Boolean
    Dim tmpMvMode As Boolean
    Dim tmpMeasureVal() As Double

    ReDim typeBpmuInf(UBound(bpmuChans))
        
    For tchCnt = 0 To UBound(bpmuChans) Step 1
        
        Call TheHdw.BPMU.Chans(bpmuChans(tchCnt)).ReadDriverRanges(tmpIrange, tmpVrange)
        
        With typeBpmuInf(tchCnt)
            .tCurrentRange = tmpIrange(0)
            .tVoltageRange = tmpVrange(0)
            .tClampCurrent = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).ClampCurrent(.tCurrentRange)
            .tClampVoltage = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).ClampVoltage(.tVoltageRange)
            .tForceCurrent = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).ForceCurrent(.tCurrentRange)
            .tForceVoltage = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).ForceVoltage(.tVoltageRange)
            .tHighLimit = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).TestLimitHigh
            .tLowLimit = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).TestLimitLow
            .tBpmuGate = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).GateOn
            .tConnectDut = mf_ChkIoRelayStat(bpmuChans(tchCnt), "BPMU")
        End With
                                                                                                                                    
        '�d������A�d������̃��[�h
        tmpFvMode = TheHdw.BPMU.Pins(bpmuPins(tchCnt)).BpmuIsForcingVoltage(siteNum(tchCnt))
        '�d������A�d������̃��[�h
        tmpMvMode = TheHdw.BPMU.Pins(bpmuPins(tchCnt)).BpmuIsMeasuringVoltage(siteNum(tchCnt))
        
        With typeBpmuInf(tchCnt)
            If tmpFvMode = True Then
                .tForcingMode = "FV"
            Else
                .tForcingMode = "FI"
            End If
                                                                          
            If tmpMvMode = True Then
                .tMeasureMode = "MV"
            Else
                .tMeasureMode = "MI"
            End If
        End With
        
        '���[�^�[���[�h
'        Call TheHdw.BPMU.chans(bpmuChans(tchCnt)).measure(1, tmpMeasureVal)
'        typeBpmuInf(tchCnt).tMeasureResult = tmpMeasureVal(0)
    
    Next tchCnt
    
End Sub

'�w��CH��HDVIS����TERADYNE-API����擾
Private Sub mf_MakeHdvisInfo(ByRef hdvischans() As Long, ByRef typeHdvisInf As type_HDVIS_INFO)
    
    Dim tmpGateStat() As Long
    Dim tmpRelayStat() As Long
    Dim tmpLowPassFilter() As Long
    Dim tmpAlarmOpnDgs() As Long
    Dim tmpAlarmOverLoad() As Long
    Dim tmpMergeFlg() As Long
    Dim tmpHdvisMode() As Long
    Dim tmpVrange() As Long
    Dim tmpIrange() As Long
    Dim tmpSlewRate() As Long
    Dim tmpRelayMode As Long
    Dim tmpForceValue() As Double
    Dim tmpClampValue() As Double
    Dim tmpMeasureValue() As Double
    Dim tmpExtMode() As Long
    Dim tmpExtSendRelay As Long
    Dim tmpExtTriggerRelay As Long
    Dim myForceValue() As Double
    Dim myClampValue() As Double
    Dim myRetIrange() As Long
    Dim myRetVrange() As Long
    Dim tchCnt As Long
    Dim readHdvis As Boolean
    Dim hdvisBoardNo As Long
    Dim setupEnable As Boolean
    
    'HDVIS���T�|�[�g���Ă��Ȃ�IG-XL�ŃR���p�C���G���[�ƂȂ�̂�
    '������邽��TheHdw.HDVIS��u������
'    Dim myHdvis As HdwDrivers.DriverHDVIS
    Dim myHdvis As Object
    Set myHdvis = TheHdw.HDVIS
                       
    'HDVIS���p�\���̏���
    ReDim typeHdvisInf.tHdvisInf(UBound(hdvischans))
            
    'CH���AHDVIS�p�����[�^��Ԏ擾
    With myHdvis.Chans(hdvischans)
        tmpGateStat = .Gate
        tmpRelayStat = .relay
        tmpLowPassFilter = .LowPassFilter
        tmpAlarmOpnDgs = .alarm(0)    'hdvisAlarmOpenDGS=0
        tmpAlarmOverLoad = .alarm(1)  'hdvisAlarmOverLoad=1
        tmpMergeFlg = .MergedChannels
        Call .ReadExternalModes(tmpExtMode)
        Call .ReadSlewRates(tmpSlewRate)
        Call .ReadRangesAndMode(tmpHdvisMode, tmpVrange, tmpIrange)
    End With
        
    '�����[���[�h�擾�i�����[���[�h�͂��ׂĂ�CH���ʁACH���̐ݒ�͂Ȃ��j
    tmpRelayMode = myHdvis.RelayMode
    
    'Measure�l�擾
    Call myHdvis.Chans(hdvischans).measure(1, tmpMeasureValue)
                
    '�w��CH��HDVIS���[�h�ʔ����Force�AClamp�̒l�̎擾
    For tchCnt = 0 To UBound(hdvischans)  '�擾�Ώ�CH LOOP
        
        'Force���[�h�ɉ����ă����W�Ɛݒ�l���擾
        Select Case tmpHdvisMode(tchCnt)
            Case 1 'hdvisForceIMeasureV: 'FI
                readHdvis = myHdvis.Chans(hdvischans(tchCnt)).ReadForceCurrents(myRetIrange, myForceValue)
                readHdvis = myHdvis.Chans(hdvischans(tchCnt)).ReadClampVoltages(myRetVrange, myClampValue)
            Case 0 'hdvisForceVMeasureI: 'FV
                readHdvis = myHdvis.Chans(hdvischans(tchCnt)).ReadForceVoltages(myRetVrange, myForceValue)
                readHdvis = myHdvis.Chans(hdvischans(tchCnt)).ReadClampCurrents(myRetIrange, myClampValue)
            Case 4 'hdvisMeasureV: 'MV HDVIS��MV���[�h����V-Clamp�̋@�\�͖���
                ReDim myForceValue(0)
                myForceValue(0) = 0#
                ReDim myClampValue(0)
                myClampValue(0) = 0#
        End Select
        
        '�{�[�h�P�ʂő��݂���ݒ�̏�Ԏ擾
        With myHdvis
            hdvisBoardNo = .SlotNumber(hdvischans(tchCnt)) 'CH�ԍ�����{�[�h�ԍ��擾
            setupEnable = .board(hdvisBoardNo).Setup.Enable
            tmpExtSendRelay = .board(hdvisBoardNo).ExternalSend.relay       '�ݒ�̓{�[�h��
            tmpExtTriggerRelay = .board(hdvisBoardNo).ExternalTrigger.relay '�ݒ�̓{�[�h��
        End With
        
        'HDVIS���A�\���̂֎擾���ʂ��l�ߍ���
        With typeHdvisInf.tHdvisInf(tchCnt)
            .tGate = tmpGateStat(tchCnt)
            .tRelay = tmpRelayStat(tchCnt)
            .tLowPassFilter = tmpLowPassFilter(tchCnt)
            .tAlarmOpenDgs = tmpAlarmOpnDgs(tchCnt)
            .tAlarmOverLoad = tmpAlarmOverLoad(tchCnt)
            .tMargePinFlag = tmpMergeFlg(tchCnt)
            .tMode = tmpHdvisMode(tchCnt)
            .tVRange = tmpVrange(tchCnt)
            .tIRange = tmpIrange(tchCnt)
            .tSlewRate = tmpSlewRate(tchCnt)
            .tRelayMode = tmpRelayMode
            .tForceValue = myForceValue(0)
            .tClampValue = myClampValue(0)
            .tMeasureResult = tmpMeasureValue(tchCnt)
            .tExtMode = tmpExtMode(tchCnt)
            .tExtSendRelay = tmpExtSendRelay
            .tExtTriggerRelay = tmpExtTriggerRelay
            .tSetupEnable = setupEnable
        End With
    
    Next tchCnt
    
    Set myHdvis = Nothing
    
End Sub

'�f�W�^��PIN�̃����[��Ԃ��m�F���A�w�胊�\�[�X�̐ڑ���Ԃ�Ԃ��܂��B
Private Function mf_ChkIoRelayStat(DigitalChNo As Long, ChkResourceName As String) As Boolean

    Select Case mf_GetIoRelayStat(DigitalChNo)
        
        Case rlyPE:
            If ChkResourceName = "PE" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
        
        Case rlyPPMU:
            If ChkResourceName = "PPMU" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
        
        Case rlyBPMU:
            If ChkResourceName = "BPMU" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
        
        Case rlyHV:
            If ChkResourceName = "HV" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
        
        Case rlyDisconnect:
            mf_ChkIoRelayStat = False
        
        Case rlyPPMU_PE:
            If ChkResourceName = "PPMU" Or ChkResourceName = "PE" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
                
        Case rlyPPMU_HV:
            If ChkResourceName = "PPMU" Or ChkResourceName = "HV" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
        
        Case rlyBPMU_PE:
            If ChkResourceName = "BPMU" Or ChkResourceName = "PE" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
                
        Case rlyBPMU_PPMU:
            If ChkResourceName = "BPMU" Or ChkResourceName = "PPMU" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
                
        Case rlyBPMU_HV:
            If ChkResourceName = "BPMU" Or ChkResourceName = "HV" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
                
    End Select

End Function
'�f�W�^��Pin�̃����[�ڑ���Ԃ��擾����
Private Function mf_GetIoRelayStat(chNumber As Long) As RlyType
    Dim rlyStat As RlyType

    On Error GoTo RLY_DISCON

    mf_GetIoRelayStat = TheHdw.Digital.relays.chan(chNumber).whichChanRelay
    
    Exit Function

RLY_DISCON:
    mf_GetIoRelayStat = rlyDisconnect

End Function



'�X�i�b�v�V���b�g�̃w�b�_�����쐬
Private Sub mf_Make_SsHeader(ByVal ResourceName As String, ByVal tstIdLabel As String, _
ByRef pinNameArr() As String, ByRef chansArr() As Long, ByRef ssHeaderInf As type_SS_HEADER)
    
    Dim PinCnt As Long
    Dim siteNumCnt As Long
    Dim lopCnt As Long

    '�w�b�_�p�̍\���̂̔�������
    With ssHeaderInf
        ReDim .tResourceName(UBound(chansArr))
        ReDim .tIdLabel(UBound(chansArr))
        ReDim .tPinName(UBound(chansArr))
        ReDim .tSiteNumber(UBound(chansArr))
    End With
    
    '�`�����l���ԍ��z������炤
    ssHeaderInf.tChannelNumber = chansArr
    
    lopCnt = 0
    
    '�T�C�gNO�ƃs�����A���\�[�X���́A���x�����̍쐬
    For PinCnt = 0 To UBound(pinNameArr) '�Ώۃs�� LOOP
        For siteNumCnt = 0 To TheExec.sites.ExistingCount - 1 '�}���`�T�C�g LOOP
            
            With ssHeaderInf
                .tResourceName(lopCnt) = ResourceName
                .tIdLabel(lopCnt) = tstIdLabel
                .tPinName(lopCnt) = pinNameArr(PinCnt)
                .tSiteNumber(lopCnt) = siteNumCnt
            End With
                       
            lopCnt = lopCnt + 1
        Next siteNumCnt
    Next PinCnt

End Sub
