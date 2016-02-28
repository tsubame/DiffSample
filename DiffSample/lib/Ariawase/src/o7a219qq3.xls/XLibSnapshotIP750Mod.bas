Attribute VB_Name = "XLibSnapshotIP750Mod"
'�T�v:
'   IP750�e�X�^�e���\�[�X�̐ݒ��Ԏ擾�ׂ̈̃��C�u����
'
'�ړI:
'   �T:IP750�e�X�^�̏�Ԏ擾
'   �U:IP750�e�X�^�̎擾��Ԃ�ۑ�
'
'�쐬��:
'   SLSI����
'
'2007-10-12 HDVIS���\�[�X�ǉ� for IG-XL 3.40.17�p
'2007-10-26 APMU�擾���̎󂯓n�����@�ύX
'           TheHdw.HDVIS�̎�舵�����@�ύX
'2007-11-14 BPMU�̃����W�擾��ReadDriverRanges�ɕύX
'           HDVIS��MV���[�h�ݒ莞�̃N�����v�d����0�ɐݒ�
'          �i�N�����v�@�\�͖������߁j
'2007-11-15 ��IG-XL�ł̃R���p�C���G���[����̂��߁AHDVIS��`���C��
'
'

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
    tMeasureResult As Double
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
    tMeasureResult() As Double
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
    tMeasureResult As Double
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
    tMeasureResult As Double
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
    tMeasureResult As Double
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

Public Sub GetTesterInfo(Optional ByVal idLabel As String, _
Optional ByVal ouputDataWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")
'���e:
'   ChannelMap�ɒ�`����Ă���IP750�e�X�^���\�[�X�̏����擾���܂�
'
'[idLabel]          In  �擾���ʂɕ\�����郉�x���i�w�肵�Ȃ��Ƃ��ɂ͎��s����TestInstances�������x���ɂȂ�܂��j
'[ouputDataWindow]  In  �擾���ʂ��f�[�^���OWindow�֏o�͂���ۂ�1��ݒ肷��
'[outPutLogName]    In  �擾���ʂ��O����Txt�t�@�C���֏o�͂���ۂ̓t�@�C�������w�肷��
'
'���ӎ���:
'   �擾�ł���̂͌���APMU/DPS/PE/PPMU/BPMU/HDVIS���ł�
    
    'TEST ID���x���Ɏw�肪�Ȃ���΁A�e�X�g�C���X�^���X�������x���Ƃ��Ďg�p����B
    idLabel = mf_Set_IdLabel(idLabel)

    Call CreateApmuInfo(idLabel, ouputDataWindow, outputLogName)
    Call CreateDpsInfo(idLabel, ouputDataWindow, outputLogName)
    Call CreatePeInfo(idLabel, ouputDataWindow, outputLogName)
    Call CreatePpmuInfo(idLabel, ouputDataWindow, outputLogName)
    Call CreateBpmuInfo(idLabel, ouputDataWindow, outputLogName)
    #If HDVIS_USE <> 0 Then
    Call CreateHdvisInfo(idLabel, ouputDataWindow, outputLogName)
    #End If

End Sub

Public Sub CreateApmuInfo(Optional ByVal testIdLabel As String = "*", _
Optional ByVal outputDataWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")
'���e:
'   ChannelMap�ɒ�`����Ă���APMU�s����APMU�����擾���܂�
'
'[testIdLabel]      In  �擾���ʂɕ\�����郉�x��
'[outputDataWindow] In  �擾���ʂ��f�[�^���OWindow�֏o�͂���ۂ�1��ݒ肷��
'[outPutLogName]    In  �擾���ʂ��O����Txt�t�@�C���֏o�͂���ۂ̓t�@�C�������w�肷��
'
'���ӎ���:
'   Mesure�l�́A1���荞�ݎ��̒l�ƂȂ�܂��B
        
    Dim tPinNameArr() As String
    Dim tChansArr() As Long
    Dim ResourceName As String
    Dim resourceChk As Boolean
    Dim tApmuinf As type_APMU_INFO 'APMU���p�\����
        
    
    ResourceName = "[APMU]" 'IP750���\�[�X���ʗp���x��[APMU]
            
    'APMU���\�[�X���g�p���Ă���Channel�𒲂ׂ�
    resourceChk = mf_ChkResourcePin(chAPMU, tChansArr, tPinNameArr)
                                                                                                
    'APMU���\�[�X���g�p���Ă��Ȃ��Ƃ��͏I��
    If resourceChk = False Then
        Call mf_OutputResourceNothingMsg("APMU", outputDataWindow, outputLogName)
        Exit Sub
    End If
                                                                                                
    'APMU���̃w�b�_���̏����쐬
    Call mf_Make_SsHeader(ResourceName, testIdLabel, tPinNameArr, tChansArr, tApmuinf.tSsHeader)
                                                                                           
    'APMU�����e���_�C��API���g�p���Ď擾
    Call mf_GetApmuInfo(tApmuinf.tSsHeader.tChannelNumber, tApmuinf)
    
    '�擾���ʂ��o�͂���֐���APMU��񂪓����Ă���\���̂�n��
    Call mf_DispApmuInfo(tApmuinf, outputDataWindow, outputLogName)

End Sub

Public Sub CreateDpsInfo(Optional ByVal testIdLabel As String = "*", _
Optional ByVal outputDataWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")
'���e:
'   ChannelMap�ɒ�`����Ă���DPS�s����DPS�����擾���܂�
'
'[testIdLabel]      In  �擾���ʂɕ\�����郉�x��
'[outputDataWindow] In  �擾���ʂ��f�[�^���OWindow�֏o�͂���ۂ�1��ݒ肷��
'[outPutLogName]    In  �擾���ʂ��O����Txt�t�@�C���֏o�͂���ۂ̓t�@�C�������w�肷��
'
    Dim tPinNameArr() As String
    Dim tChansArr() As Long
    Dim ResourceName As String
    Dim resourceChk As Boolean
    Dim tDpsinf As type_DPS_INFO 'DPS���p�\����
    
    ResourceName = "[DPS]" 'IP750���\�[�X���ʗp���x��[DPS]
            
    'DPS���\�[�X���g�p���Ă���Channel�𒲂ׂ�
    resourceChk = mf_ChkResourcePin(chDPS, tChansArr, tPinNameArr)
                                                                                                
    'DPS���\�[�X���g�p���Ă��Ȃ��Ƃ��͏I��
    If resourceChk = False Then
        Call mf_OutputResourceNothingMsg("DPS", outputDataWindow, outputLogName)
        Exit Sub
    End If
                                                                                            
    'DPS���̃w�b�_���̏����쐬
    Call mf_Make_SsHeader(ResourceName, testIdLabel, tPinNameArr, tChansArr, tDpsinf.tSsHeader)
                                                                                           
    'DPS����TERADYNE-API����擾
    Call mf_GetDpsInfo(tDpsinf.tSsHeader.tChannelNumber, tDpsinf)

    '�擾���ʂ��o�͂���֐���DPS��񂪓����Ă���\���̂�n��
    Call mf_DispDpsInfo(tDpsinf, outputDataWindow, outputLogName)

End Sub

Public Sub CreatePeInfo(Optional ByVal testIdLabel As String = "*", _
Optional ByVal outputDataWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")
'���e:
'   ChannelMap�ɒ�`����Ă���I/O�s����PE�����擾���܂�
'
'[testIdLabel]      In  �擾���ʂɕ\�����郉�x��
'[outputDataWindow] In  �擾���ʂ��f�[�^���OWindow�֏o�͂���ۂ�1��ݒ肷��
'[outPutLogName]    In  �擾���ʂ��O����Txt�t�@�C���֏o�͂���ۂ̓t�@�C�������w�肷��
'
    Dim tPinNameArr() As String
    Dim tChansArr() As Long
    Dim ResourceName As String
    Dim resourceChk As Boolean
    Dim tPeinf As type_IO_INFO 'PE���p�\����
    
    ResourceName = "[PE]" 'IP750���\�[�X���ʗp���x��[PE]
            
    'I/O(PE)���\�[�X���g�p���Ă���Channel�𒲂ׂ�
    resourceChk = mf_ChkResourcePin(chIO, tChansArr, tPinNameArr)
                                                                                                
    'I/O(PE)���\�[�X���g�p���Ă��Ȃ��Ƃ��͏I��
    If resourceChk = False Then
        Call mf_OutputResourceNothingMsg("I/O", outputDataWindow, outputLogName)
        Exit Sub
    End If
                                                                                            
    'PE���̃w�b�_���̏����쐬
    Call mf_Make_SsHeader(ResourceName, testIdLabel, tPinNameArr, tChansArr, tPeinf.tSsHeader)
                                                                                           
    'PE����TERADYNE-API����擾
    Call mf_GetPeInfo(tPeinf.tSsHeader.tChannelNumber, tPeinf.tPeinf)

    '�擾���ʂ��o�͂���֐���PE��񂪓����Ă���\���̂�n��
    Call mf_DispPeInfo(tPeinf, outputDataWindow, outputLogName)

End Sub

Public Sub CreateBpmuInfo(Optional ByVal testIdLabel As String = "*", _
Optional ByVal outputDataWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")
'���e:
'   ChannelMap�ɒ�`����Ă���I/O�s����BPMU�����擾���܂�
'
'[testIdLabel]      In  �擾���ʂɕ\�����郉�x��
'[outputDataWindow] In  �擾���ʂ��f�[�^���OWindow�֏o�͂���ۂ�1��ݒ肷��
'[outPutLogName]    In  �擾���ʂ��O����Txt�t�@�C���֏o�͂���ۂ̓t�@�C�������w�肷��
'
    Dim tPinNameArr() As String
    Dim tChansArr() As Long
    Dim ResourceName As String
    Dim resourceChk As Boolean
    Dim tIoinf As type_IO_INFO 'IO�s�����p�\����
    
    ResourceName = "[BPMU]" 'IP750���\�[�X���ʗp���x��[BPMU]
            
    'IO���\�[�X���g�p���Ă���Channel�𒲂ׂ�
    resourceChk = mf_ChkResourcePin(chIO, tChansArr, tPinNameArr)
                                                                                                
    'IO���\�[�X���g�p���Ă��Ȃ��Ƃ��͏I��
    If resourceChk = False Then
        Call mf_OutputResourceNothingMsg("I/O", outputDataWindow, outputLogName)
        Exit Sub
    End If
                                                                                            
    'BPMU���̃w�b�_���̏����쐬
    Call mf_Make_SsHeader(ResourceName, testIdLabel, tPinNameArr, tChansArr, tIoinf.tSsHeader)
                                                                                           
    'BPMU����TERADYNE-API����擾
    Call mf_GetBpmuInfo(tIoinf.tSsHeader.tChannelNumber, tIoinf.tSsHeader.tPinName, tIoinf.tSsHeader.tSiteNumber, tIoinf.tBpmuinf)

    '�擾���ʂ��o�͂���֐���PPMU��񂪓����Ă���\���̂�n��
    Call mf_DispBpmuInfo(tIoinf, outputDataWindow, outputLogName)

End Sub

Public Sub CreatePpmuInfo(Optional ByVal testIdLabel As String = "*", _
Optional ByVal outputDataWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")
'���e:
'   ChannelMap�ɒ�`����Ă���I/O�s����PPMU�����擾���܂�
'
'[testIdLabel]      In  �擾���ʂɕ\�����郉�x��
'[outputDataWindow] In  �擾���ʂ��f�[�^���OWindow�֏o�͂���ۂ�1��ݒ肷��
'[outPutLogName]    In  �擾���ʂ��O����Txt�t�@�C���֏o�͂���ۂ̓t�@�C�������w�肷��
'
    Dim tPinNameArr() As String
    Dim tChansArr() As Long
    Dim ResourceName As String
    Dim resourceChk As Boolean
    Dim tIoinf As type_IO_INFO 'IO�s�����p�\����
    
    ResourceName = "[PPMU]" 'IP750���\�[�X���ʗp���x��[PPMU]
            
    'IO���\�[�X���g�p���Ă���Channel�𒲂ׂ�
    resourceChk = mf_ChkResourcePin(chIO, tChansArr, tPinNameArr)
                                                                                                
    'IO���\�[�X���g�p���Ă��Ȃ��Ƃ��͏I��
    If resourceChk = False Then
        Call mf_OutputResourceNothingMsg("I/O", outputDataWindow, outputLogName)
        Exit Sub
    End If
                                                                                            
    'PPMU���̃w�b�_���̏����쐬
    Call mf_Make_SsHeader(ResourceName, testIdLabel, tPinNameArr, tChansArr, tIoinf.tSsHeader)
                                                                                           
    'PPMU����TERADYNE-API����擾
    Call mf_GetPpmuInfo(tIoinf.tSsHeader.tChannelNumber, tIoinf.tPpmuinf)

    '�擾���ʂ��o�͂���֐���PPMU��񂪓����Ă���\���̂�n��
    Call mf_DispPpmuInfo(tIoinf, outputDataWindow, outputLogName)

End Sub

Public Sub CreateHdvisInfo(Optional ByVal testIdLabel As String = "*", _
Optional ByVal outputDataWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")
'���e:
'   ChannelMap�ɒ�`����Ă���HDVIS�s����HDVIS�����擾���܂�
'
'[testIdLabel]      In  �擾���ʂɕ\�����郉�x��
'[outputDataWindow] In  �擾���ʂ��f�[�^���OWindow�֏o�͂���ۂ�1��ݒ肷��
'[outPutLogName]    In  �擾���ʂ��O����Txt�t�@�C���֏o�͂���ۂ̓t�@�C�������w�肷��
'
'���ӎ���:
'   Mesure�l�́A1���荞�ݎ��̒l�ƂȂ�܂��B
        
    Dim tPinNameArr() As String
    Dim tChansArr() As Long
    Dim ResourceName As String
    Dim resourceChk As Boolean
    Dim tHdvisInf As type_HDVIS_INFO 'HDVIS���p�\����
    
    Const CH_HDVIS = 36
        
    ResourceName = "[HDVIS]" 'IP750���\�[�X���ʗp���x��[HDVIS]
            
    'HDVIS���\�[�X���g�p���Ă���Channel�𒲂ׂ�
    resourceChk = mf_ChkResourcePin(CH_HDVIS, tChansArr, tPinNameArr)
                                                                                                
    'HDVIS���\�[�X���g�p���Ă��Ȃ��Ƃ��͏I��
    If resourceChk = False Then
        Call mf_OutputResourceNothingMsg("HDVIS", outputDataWindow, outputLogName)
        Exit Sub
    End If
                                                                                                
    'HDVIS���̃w�b�_���̏����쐬
    Call mf_Make_SsHeader(ResourceName, testIdLabel, tPinNameArr, tChansArr, tHdvisInf.tSsHeader)
                                                                                           
    'HDVIS�����e���_�C��API���g�p���Ď擾
    Call mf_GetHdvisInfo(tHdvisInf.tSsHeader.tChannelNumber, tHdvisInf)
    
    '�擾���ʂ��o�͂���֐���HDVIS��񂪓����Ă���\���̂�n��
    Call mf_DispHdvisInfo(tHdvisInf, outputDataWindow, outputLogName)

End Sub

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
        Call TheHdw.APMU.Chans(apmuChans).measure(1, .tMeasureResult)   '���[�^�ǂݎ��l���擾
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
            .tMeasureResult = tmpApmuInf.tMeasureResult(tchCnt)
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
            
            '�d���v�̓d���l�擾
            Call TheHdw.DPS.Chans(dpsChans(tchCnt)).MeasureCurrents(.tCurrentRange, tmpMesureVal)
            
            .tMeasureResult = 0

            For aveCnt = 0 To UBound(tmpMesureVal)
                .tMeasureResult = .tMeasureResult + tmpMesureVal(aveCnt)
            Next aveCnt
        
            .tMeasureResult = .tMeasureResult / .tMeasureSamples
        
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
                Call TheHdw.PPMU.Chans(ppmuChans(tchCnt)).MeasureVoltages(tmpMeasureVal)
            Else
                .tForceType = "VOLTS"
                Call TheHdw.PPMU.Chans(ppmuChans(tchCnt)).MeasureCurrents(tmpMeasureVal)
            End If
                                                                                          
            .tMeasureResult = 0
            .tMeasureSamples = UBound(tmpMeasureVal) + 1
            
            For aveCnt = 0 To UBound(tmpMeasureVal)
                .tMeasureResult = .tMeasureResult + tmpMeasureVal(aveCnt)
            Next aveCnt
                                                                              
            .tMeasureResult = .tMeasureResult / .tMeasureSamples
                                                                                                                                                            
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
        Call TheHdw.BPMU.Chans(bpmuChans(tchCnt)).measure(1, tmpMeasureVal)
        typeBpmuInf(tchCnt).tMeasureResult = tmpMeasureVal(0)
    
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

'�w��CH��PPMU����TERAPI-Logger����擾�i����j
'Private Sub mf_MakeTerapiPpmuInfo(ByRef ppmuChans() As Long, ByRef typePpmuInf() As type_PPMU_INFO)
'
'    Dim snapshotSupplier As TERAPISnapshotService.TERAPISnapshots
'    Set snapshotSupplier = New TERAPISnapshotService.TERAPISnapshots
'
'    Dim varSnapshot As Variant
'    Dim channelSnapshot As TERAPISnapshotService.SnapPPMUChannel
'
'    Dim chCnt As Long
'
'    ReDim typePpmuInf(UBound(ppmuChans))
'
'    chCnt = 0
'
'    For Each varSnapshot In snapshotSupplier.PPMUSnapshotUtility.GetSnapshotChannels(ppmuChans)
'
'        Set channelSnapshot = varSnapshot
'
'        With typePpmuInf(chCnt)
'            .tForceVoltage = channelSnapshot.ForceVoltage
'            .tForceCurrent = channelSnapshot.ForceCurrent
'            .tCurrentRange = channelSnapshot.LastCurrentRange
'            .tHighLimit = channelSnapshot.TestLimitHigh
'            .tLowLimit = channelSnapshot.TestLimitLow
'            .tConnect = channelSnapshot.IsConnected
'            .tForceType = channelSnapshot.IsForcing
'        End With
'
'        chCnt = chCnt + 1
'
'    Next varSnapshot
'
'    Set channelSnapshot = Nothing
'    Set snapshotSupplier = Nothing
'
'End Sub

'�擾APMU�����o��
Private Sub mf_DispApmuInfo(ByRef apmuInf As type_APMU_INFO, _
Optional ByVal outputWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")

    Dim dispMsg As String
    Dim infCnt As Long

    For infCnt = 0 To UBound(apmuInf.tSsHeader.tChannelNumber) Step 1
        'APMU���̏o�̓t�H�[�}�b�g�쐬
        dispMsg = mf_makeApmuInfFmt(infCnt, apmuInf)
                        
        '����OUTPUT Window�֏o��
        If outputWindow = 1 Then
            TheExec.Datalog.WriteComment dispMsg
        End If
        '�������O�t�@�C���֏o��
        If outputLogName <> "" Then
            Call mf_OutPutLog(outputLogName, dispMsg)
        End If
    Next infCnt

    If outputWindow = 1 Then
        TheExec.Datalog.WriteComment ""
    End If
    
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub

'�擾DPS�����o��
Private Sub mf_DispDpsInfo(ByRef dpsInf As type_DPS_INFO, _
Optional ByVal outputWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")

    Dim dispMsg As String
    Dim infCnt As Long

    For infCnt = 0 To UBound(dpsInf.tSsHeader.tChannelNumber) Step 1
        'DPS���̏o�̓t�H�[�}�b�g�쐬
        dispMsg = mf_makeDpsInfFmt(infCnt, dpsInf)
                        
        '����OUTPUT Window�֏o��
        If outputWindow = 1 Then
            TheExec.Datalog.WriteComment dispMsg
        End If
        '�������O�t�@�C���֏o��
        If outputLogName <> "" Then
            Call mf_OutPutLog(outputLogName, dispMsg)
        End If
    Next infCnt

    If outputWindow = 1 Then
        TheExec.Datalog.WriteComment ""
    End If
    
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub

'�擾PE�����o��
Private Sub mf_DispPeInfo(ByRef Peinf As type_IO_INFO, _
Optional ByVal outputWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")

    Dim dispMsg As String
    Dim infCnt As Long

    For infCnt = 0 To UBound(Peinf.tSsHeader.tChannelNumber) Step 1
        'PE���̏o�̓t�H�[�}�b�g�쐬
        dispMsg = mf_makePeInfFmt(infCnt, Peinf)
                        
        '����OUTPUT Window�֏o��
        If outputWindow = 1 Then
            TheExec.Datalog.WriteComment dispMsg
        End If
        '�������O�t�@�C���֏o��
        If outputLogName <> "" Then
            Call mf_OutPutLog(outputLogName, dispMsg)
        End If
    Next infCnt

    If outputWindow = 1 Then
        TheExec.Datalog.WriteComment ""
    End If
    
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub

'�擾PPMU�����o��
Private Sub mf_DispPpmuInfo(ByRef ioInf As type_IO_INFO, _
Optional ByVal outputWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")

    Dim dispMsg As String
    Dim infCnt As Long

    For infCnt = 0 To UBound(ioInf.tSsHeader.tPinName) Step 1
        'PPMU���̏o�̓t�H�[�}�b�g�쐬
        dispMsg = mf_makePpmuInfFmt(infCnt, ioInf)
                        
        '����OUTPUT Window�֏o��
        If outputWindow = 1 Then
            TheExec.Datalog.WriteComment dispMsg
        End If
        '�������O�t�@�C���֏o��
        If outputLogName <> "" Then
            Call mf_OutPutLog(outputLogName, dispMsg)
        End If
    Next infCnt

    If outputWindow = 1 Then
        TheExec.Datalog.WriteComment ""
    End If
    
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub

'�擾BPMU�����o��
Private Sub mf_DispBpmuInfo(ByRef ioInf As type_IO_INFO, _
Optional ByVal outputWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")

    Dim dispMsg As String
    Dim infCnt As Long

    For infCnt = 0 To UBound(ioInf.tSsHeader.tChannelNumber) Step 1
        'BPMU���̏o�̓t�H�[�}�b�g�쐬
        dispMsg = mf_makeBpmuInfFmt(infCnt, ioInf)
                        
        '����OUTPUT Window�֏o��
        If outputWindow = 1 Then
            TheExec.Datalog.WriteComment dispMsg
        End If
        '�������O�t�@�C���֏o��
        If outputLogName <> "" Then
            Call mf_OutPutLog(outputLogName, dispMsg)
        End If
    Next infCnt

    If outputWindow = 1 Then
        TheExec.Datalog.WriteComment ""
    End If
    
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub

'�擾HDVIS�����o��
Private Sub mf_DispHdvisInfo(ByRef hdvisInf As type_HDVIS_INFO, _
Optional ByVal outputWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")

    Dim dispMsg As String
    Dim infCnt As Long

    For infCnt = 0 To UBound(hdvisInf.tSsHeader.tChannelNumber) Step 1
        'HDVIS���̏o�̓t�H�[�}�b�g�쐬
        dispMsg = mf_makeHdvisInfFmt(infCnt, hdvisInf)
                        
        '����OUTPUT Window�֏o��
        If outputWindow = 1 Then
            TheExec.Datalog.WriteComment dispMsg
        End If
        '�������O�t�@�C���֏o��
        If outputLogName <> "" Then
            Call mf_OutPutLog(outputLogName, dispMsg)
        End If
    Next infCnt

    If outputWindow = 1 Then
        TheExec.Datalog.WriteComment ""
    End If
    
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub

'APMU���o�͗p��Format���`
Private Function mf_makeApmuInfFmt(ByVal infArrNo As Long, ByRef apmuInf As type_APMU_INFO) As String

    Dim makeMsg As String
    
    With apmuInf.tSsHeader
        makeMsg = (.tIdLabel(infArrNo) & "," _
            & .tResourceName(infArrNo) & "," _
            & "PIN=" & .tPinName(infArrNo) & "," _
            & "SITE=" & .tSiteNumber(infArrNo) & "," _
            & "CH_NUM=" & .tChannelNumber(infArrNo) & ",")
    End With
    
    With apmuInf.tApmuinf(infArrNo)
            makeMsg = (makeMsg _
            & "MODE=" & .tMode & "," _
            & "VRANGE=" & .tVRange & "," _
            & "IRANGE=" & .tIRange & "," _
            & "FORCE=" & .tForceValue & "," _
            & "CLAMP=" & .tClampValue & "," _
            & "GATE=" & .tGate & "," _
            & "RELAY=" & .tRelay & "," _
            & "LPF=" & .tLowPassFilter & "," _
            & "EXSENS=" & .tExternalSense & "," _
            & "ALARM=" & .tAlarm & "," _
            & "GANGED=" & .tGangPinFlag & "," _
            & "MEASURE_VAL=" & .tMeasureResult)
    End With

    mf_makeApmuInfFmt = makeMsg

End Function

'PPMU���o�͗p��Format���`
Private Function mf_makePpmuInfFmt(ByVal infArrNo As Long, ByRef ioInf As type_IO_INFO) As String

    Dim makeMsg As String
    
    With ioInf
        makeMsg = (.tSsHeader.tIdLabel(infArrNo) & "," _
        & .tSsHeader.tResourceName(infArrNo) & "," _
        & "PIN=" & .tSsHeader.tPinName(infArrNo) & "," _
        & "SITE=" & .tSsHeader.tSiteNumber(infArrNo) & "," _
        & "CH_NUM=" & .tSsHeader.tChannelNumber(infArrNo) & "," _
        & "FORCE_TYPE=" & .tPpmuinf(infArrNo).tForceType & "," _
        & "FORCE_VOLTAGE=" & .tPpmuinf(infArrNo).tForceVoltage & "," _
        & "FORCE_CURRENT=" & .tPpmuinf(infArrNo).tForceCurrent & "," _
        & "CURRENT_RANGE=" & .tPpmuinf(infArrNo).tCurrentRange & "," _
        & "HIGH_LIMIT=" & .tPpmuinf(infArrNo).tHighLimit & "," _
        & "LOW_LIMIT=" & .tPpmuinf(infArrNo).tLowLimit & "," _
        & "CONNECT=" & .tPpmuinf(infArrNo).tConnect & "," _
        & "MEASURE_SAMPLES=" & .tPpmuinf(infArrNo).tMeasureSamples & "," _
        & "MEASURE_VAL=" & .tPpmuinf(infArrNo).tMeasureResult)

    End With
    
    mf_makePpmuInfFmt = makeMsg

End Function

'BPMU���o�͗p��Format���`
Private Function mf_makeBpmuInfFmt(ByVal infArrNo As Long, ByRef ioInf As type_IO_INFO) As String

    Dim makeMsg As String
    
    With ioInf
        makeMsg = (.tSsHeader.tIdLabel(infArrNo) & "," _
        & .tSsHeader.tResourceName(infArrNo) & "," _
        & "PIN=" & .tSsHeader.tPinName(infArrNo) & "," _
        & "SITE=" & .tSsHeader.tSiteNumber(infArrNo) & "," _
        & "CH_NUM=" & .tSsHeader.tChannelNumber(infArrNo) & "," _
        & "FORCE_MODE=" & .tBpmuinf(infArrNo).tForcingMode & "," _
        & "MEASURE_MODE=" & .tBpmuinf(infArrNo).tMeasureMode & "," _
        & "FORCE-V=" & .tBpmuinf(infArrNo).tForceVoltage & "," _
        & "FORCE-I=" & .tBpmuinf(infArrNo).tForceCurrent & "," _
        & "CLAMP-V=" & .tBpmuinf(infArrNo).tClampVoltage & "," _
        & "CLAMP-I=" & .tBpmuinf(infArrNo).tClampCurrent & "," _
        & "I-RANGE=" & .tBpmuinf(infArrNo).tCurrentRange & "," _
        & "V-RANGE=" & .tBpmuinf(infArrNo).tVoltageRange & "," _
        & "HIGH_LIMIT=" & .tBpmuinf(infArrNo).tHighLimit & "," _
        & "LOW_LIMIT=" & .tBpmuinf(infArrNo).tLowLimit & "," _
        & "BPMU_GATE=" & .tBpmuinf(infArrNo).tBpmuGate & "," _
        & "CONNECT_DUT=" & .tBpmuinf(infArrNo).tConnectDut)
    End With
    
    mf_makeBpmuInfFmt = makeMsg

End Function

'DPS���o�͗p��Format���`
Private Function mf_makeDpsInfFmt(ByVal infArrNo As Long, ByRef dpsInf As type_DPS_INFO) As String

    Dim makeMsg As String
    
    With dpsInf
        makeMsg = (.tSsHeader.tIdLabel(infArrNo) & "," _
        & .tSsHeader.tResourceName(infArrNo) & "," _
        & "PIN=" & .tSsHeader.tPinName(infArrNo) & "," _
        & "SITE=" & .tSsHeader.tSiteNumber(infArrNo) & "," _
        & "CH_NUM=" & .tSsHeader.tChannelNumber(infArrNo) & "," _
        & "AMP_RNG=" & .tDpsinf(infArrNo).tCurrentRange & "," _
        & "OUT_SRC=" & .tDpsinf(infArrNo).tOutputSource & "," _
        & "PRI-V=" & .tDpsinf(infArrNo).tPrimaryVoltage & "," _
        & "ALT-V=" & .tDpsinf(infArrNo).tAlternateVoltage & "," _
        & "CURRENT_LIM=" & .tDpsinf(infArrNo).tCurrentLimit & "," _
        & "FORCE_RLY=" & .tDpsinf(infArrNo).tForceRelay & "," _
        & "SENS_RLY=" & .tDpsinf(infArrNo).tSenseRelay & "," _
        & "MEASURE_SAMPLES=" & .tDpsinf(infArrNo).tMeasureSamples & "," _
        & "MEASURE_VAL=" & .tDpsinf(infArrNo).tMeasureResult)
    End With
    
    mf_makeDpsInfFmt = makeMsg

End Function

'PE���o�͗p��Format���`
Private Function mf_makePeInfFmt(ByVal infArrNo As Long, ByRef Peinf As type_IO_INFO) As String

    Dim makeMsg As String
        
    With Peinf
        makeMsg = (.tSsHeader.tIdLabel(infArrNo) & "," _
        & .tSsHeader.tResourceName(infArrNo) & "," _
        & "PIN=" & .tSsHeader.tPinName(infArrNo) & "," _
        & "SITE=" & .tSsHeader.tSiteNumber(infArrNo) & "," _
        & "CH_NUM=" & .tSsHeader.tChannelNumber(infArrNo) & "," _
        & "VCH=" & .tPeinf(infArrNo).tVClampHi & "," _
        & "VCL=" & .tPeinf(infArrNo).tVClampLo & "," _
        & "VIH=" & .tPeinf(infArrNo).tVDriveHi & "," _
        & "VIL=" & .tPeinf(infArrNo).tVDriveLo & "," _
        & "VOH=" & .tPeinf(infArrNo).tVCompareHi & "," _
        & "VOL=" & .tPeinf(infArrNo).tVCompareLo & "," _
        & "IOH=" & .tPeinf(infArrNo).tISink & "," _
        & "IOL=" & .tPeinf(infArrNo).tISource & "," _
        & "VT=" & .tPeinf(infArrNo).tVThreshold & "," _
        & "HV_VPH=" & .tPeinf(infArrNo).tHvVph & "," _
        & "HV_IPH=" & .tPeinf(infArrNo).tHvIph & "," _
        & "HV_TPR=" & .tPeinf(infArrNo).tHvTpr & "," _
        & "HV_CONNECT=" & .tPeinf(infArrNo).tHvConnect & "," _
        & "PE_CONNECT=" & .tPeinf(infArrNo).tPeConnect & "," _
        & "PPMU_CONNECT=" & .tPeinf(infArrNo).tPpmuConnect & "," _
        & "BPMU_CONNECT=" & .tPeinf(infArrNo).tBpmuConnect)

    End With
    
    mf_makePeInfFmt = makeMsg

End Function

'HDVIS���o�͗p��Format���`
Private Function mf_makeHdvisInfFmt(ByVal infArrNo As Long, ByRef hdvisInf As type_HDVIS_INFO) As String

    Dim makeMsg As String
    
    With hdvisInf.tSsHeader
        makeMsg = (.tIdLabel(infArrNo) & "," _
            & .tResourceName(infArrNo) & "," _
            & "PIN=" & .tPinName(infArrNo) & "," _
            & "SITE=" & .tSiteNumber(infArrNo) & "," _
            & "CH_NUM=" & .tChannelNumber(infArrNo) & ",")
    End With

    With hdvisInf.tHdvisInf(infArrNo)
        makeMsg = (makeMsg _
            & "MODE=" & .tMode & "," _
            & "VRANGE=" & .tVRange & "," _
            & "IRANGE=" & .tIRange & "," _
            & "FORCE=" & .tForceValue & "," _
            & "CLAMP=" & .tClampValue & "," _
            & "GATE=" & .tGate & "," _
            & "RLY=" & .tRelay & "," _
            & "LPF=" & .tLowPassFilter & "," _
            & "ALM_OPNDGS=" & .tAlarmOpenDgs & "," _
            & "ALM_OVRLOAD=" & .tAlarmOverLoad & "," _
            & "MARGED=" & .tMargePinFlag & "," _
            & "RLYMODE=" & .tRelayMode & "," _
            & "SLEWRATE=" & .tSlewRate & "," _
            & "EXTMODE=" & .tExtMode & "," _
            & "SETUP_ENA=" & .tSetupEnable & "," _
            & "EXSEND_RLY=" & .tExtSendRelay & "," _
            & "EXTRIG_RLY=" & .tExtTriggerRelay & "," _
            & "MEASURE_VAL=" & .tMeasureResult)
    End With

    mf_makeHdvisInfFmt = makeMsg

End Function

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

'�f�W�^��Pin�̃����[�ڑ���Ԃ��擾����
Private Function mf_GetIoRelayStat(chNumber As Long) As RlyType
    Dim rlyStat As RlyType

    On Error GoTo RLY_DISCON

    mf_GetIoRelayStat = TheHdw.Digital.relays.chan(chNumber).whichChanRelay
    
    Exit Function

RLY_DISCON:
    mf_GetIoRelayStat = rlyDisconnect

End Function

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

'�f�W�^��Pin�̃����[�ڑ���Ԃ��擾��������ɕϊ��i�����߂��p TestCode�j
Private Function mf_DispIoRelayStat(chNumber As Long) As String

    Dim rlyStat As RlyType

    On Error GoTo RLY_DISCON

    rlyStat = TheHdw.Digital.relays.chan(chNumber).whichChanRelay

    Select Case rlyStat
        Case rlyPE:
            mf_DispIoRelayStat = "PE"
        Case rlyPPMU:
            mf_DispIoRelayStat = "PPMU"
        Case rlyBPMU:
            mf_DispIoRelayStat = "BPMU"
        Case rlyHV:
            mf_DispIoRelayStat = "HV"
        Case rlyDisconnect:
            mf_DispIoRelayStat = "Disconnect"
        Case rlyPPMU_PE:
            mf_DispIoRelayStat = "PPMU & PE"
        Case rlyPPMU_HV:
            mf_DispIoRelayStat = "PPMU & HV"
        Case rlyBPMU_PE:
            mf_DispIoRelayStat = "BPMU & PE"
        Case rlyBPMU_PPMU:
            mf_DispIoRelayStat = "BPMU & PPMU"
        Case Else
            mf_DispIoRelayStat = "Unknown"
    End Select
    
    Exit Function

RLY_DISCON:
    mf_DispIoRelayStat = "Disconnect"
    
End Function

'�X�i�b�v�V���b�g�p�̃��O���t�@�C���ɏo�͂���B
Private Sub mf_OutPutLog(ByVal LogFileName As String, outPutMessage As String)
    Dim fp As Integer
    On Error GoTo OUT_PUT_LOG_ERR
    
    fp = FreeFile
    Open LogFileName For Append As fp
    Print #fp, outPutMessage
    Close fp
    
    Exit Sub

OUT_PUT_LOG_ERR:
    Call MsgBox(LogFileName & " MsgOutPut Error", vbFalse Or vbCritical, "@mf_OutPutLog")
    Stop

End Sub

'��������z��Ɋi�[����Ă���v�f�̖��O���A�J���}��؂�`���ō쐬
Private Function mf_Make_PinNameStr(ByRef pinNameArr() As String) As String

    Dim tLoopCnt As Long
        
    '�z��Ɋi�[����Ă��邷�ׂĂ�PIN�̖��O���A�J���}��؂�`���ō쐬�@�@("P_PIN1,P_PIN2, .....")
    mf_Make_PinNameStr = pinNameArr(0)
    
    For tLoopCnt = 1 To UBound(pinNameArr)
        mf_Make_PinNameStr = mf_Make_PinNameStr & "," & pinNameArr(tLoopCnt)
    Next tLoopCnt

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

'TEST ID���x���Ɏw�肪�Ȃ���΁A�e�X�g�C���X�^���X�������x���Ƃ��Ďg�p����B
Private Function mf_Set_IdLabel(ByVal idLabel As String) As String
    If idLabel = "" Then
        mf_Set_IdLabel = TheExec.DataManager.InstanceName
    Else
        mf_Set_IdLabel = idLabel
    End If
    
End Function

'�w�胊�\�[�X��ChannelMAP�ɂ܂�������`����Ă��Ȃ��Ƃ��̃��b�Z�[�W�\���p
Private Sub mf_OutputResourceNothingMsg(ByVal ResourceName As String, _
Optional ByVal outputWindow As Integer = 1, _
Optional ByVal outputLogName As String = "")

    Dim dispMsg As String
    Dim nothingHeader As String
    
    nothingHeader = "@@@"

    dispMsg = nothingHeader & "," & "[SnapShot]" & "," & ResourceName & ".Type_doesn't_exist_in_the_ChannelMap"
                        
    '����OUTPUT Window�֏o��
    If outputWindow = 1 Then
        TheExec.Datalog.WriteComment dispMsg
        TheExec.Datalog.WriteComment ""
    End If
    '�������O�t�@�C���֏o��
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, dispMsg)
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub


