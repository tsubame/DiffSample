Attribute VB_Name = "XLibJob"
'�T�v:
'   ���ʎg�p����Object�̐����Ə�����
'
'       Revision History:
'           Date        Description
'           2013/6/11   TheDcTest�̌^��CDcScenario=>IDcScenario�ɕύX(0145184306)
'
'�ړI:
'   ���ʎg�p����Object�̐����Ə������̏������܂Ƃ߂�
'   �������菇�A�����ŊeObject�̐����A�����������s����B
'
'�쐬��:
'   SLSI����
'

Option Explicit

'Tool�Ή���ɃR�����g�O���Ď��������ɂ���B�@2013/03/07 H.Arikawa
#Const CUB_UB_USE = 0    'CUB UB�̐ݒ�          0�F���g�p�A0�ȊO�F�g�p

'Public Object�Q
Public TheUB As New CUtyBitController 'UB�ݒ�Object
Public TheDC As New CVISVISrcSelector '�d���ݒ�Object
Public TheSnapshot As New CSnapIP750  '�X�i�b�v�V���b�g�@�\Object

Public TheDcTest As IDcScenario
Public TheOffsetResult As COffsetManager

Dim mDataManagerReader As CDataSheetManager
Dim mJobListReader As CDataSheetManager
Dim mDcScenarioReader As CDcScenarioSheetReader
Dim mDcScenarioWriter As CDcScenarioSheetLogWriter
Dim mDcReplayDataReader As CDcPlaybackSheetReader
Dim mInstanceReader As CInstanceSheetReader
Dim mOffsetReader As COffsetSheetReader

Dim mDcLogReportWriter As CDcLogReportWriter

#If CUB_UB_USE <> 0 Then
Public CUBUtilBit As New CUBUtilityBits.UtilityBits 'CUB UB�ݒ�pObject
#End If

Public Sub InitJob()
'���e:
'   JOB�̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'

    Call InitTheDC       'DC�{�[�h�Z���N�^�̏�����
    Call InitTheUB       'UB�R���g���[���̏�����
    Call InitTheSnapshot '�X�i�b�v�V���b�g�@�\�̏�����

End Sub

Public Sub InitTheSnapshot()
'���e:
'   �X�i�b�v�V���b�g�@�\�̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    With TheSnapshot
        .Initialize                      '�X�i�b�v�V���b�g�@�\�̏�����
        .LogFileName = GetSnapFilename   '�O��TXT�o�̓t�@�C����
        .OutputPlace = snapTXT_FILE      '�擾���ʏo�͐�
        .OutputSaveStatus = True         '�X�i�b�v�V���b�g�@�\�̓���󋵂��f�[�^���O�ɏo��
        .SerialNumber = 1                '���O�ɏo�͂���V���A���ԍ��̏����l
    End With

End Sub

#If CUB_UB_USE <> 0 Then
Public Sub InitCub()
'���e:
'   CUB UB�̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'   ���p���ɂ͏����t�R���p�C�������� CUB_UB_USE=1�̋L�ڂ��K�v�ł��B
'   CUB UB�g�p���ɂ͖{��������Ƃ��e�X�^�[�C�j�V�������Ɏ��s����K�v������܂��B
'
    
    With CUBUtilBit
        .SetTheHdw TheHdw
        .SetTheExec TheExec
        .Clear
    End With
    
    TheHdw.DIB.LeavePowerOn = True

End Sub
#End If

Public Sub InitTheDC()
'���e:
'   �d���ݒ�{�[�h�Z���N�^�̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    
    Call TheDC.Initialize

End Sub

Public Sub InitTheUB()
'���e:
'   'UB�R���g���[���̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'   CUB��UB���p���ɂ͏����t�R���p�C�������� CUB_UB_USE=1�̋L�ڂ��K�v�ł��B
'

    '######### ������ #########################################
    Call TheUB.Initialize

    '######### �ݒ� #########################################
    'APMU UB
    With TheUB.AsAPMU
        Set .UBSetSht = Worksheets("APMU UB") '�����\�̃V�[�g�w��
        Call .LoadCondition                       '�����\��Load
    End With

    #If CUB_UB_USE <> 0 Then
    'CUB UB
    With TheUB.AsCUB
        Set .UBSetSht = Worksheets("CUB UB") '�����\�̃V�[�g�w��
        Call .LoadCondition                      '�����\��Load
    End With
    #End If

End Sub

Public Sub InitControlShtReader()
'���e:
'   JOB���X�g���̃R���g���[���V�[�g���[�_�[��������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'   �A�N�e�B�u�ȃf�[�^�V�[�g�w�肪�ύX���ꂽ�ꍇ��
'   �e���_�C���f�[�^�c�[����̃f�[�^�ύX���ꂽ�ꍇ��
'   �o���f�[�V�����̍ۂɍs��
'
    Set mDataManagerReader = Nothing
    Set mJobListReader = Nothing
End Sub

Public Sub InitTestScenario()
'���e:
'   �e�I�u�W�F�N�g�̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
'
    On Error GoTo ErrHandler
    '### ���茋�ʃ}�l�[�W�������� #########################
    XLibResultManagerUtility.InitResult
    '### �e�I�u�W�F�N�g�̂̏����� #########################
    If mDataManagerReader Is Nothing Or mJobListReader Is Nothing Then
        InitActiveDataSheet
        InitTheDcScenario
        InitTheOffsetResult
    Else
        ReInitTheDcScenario
    End If
    '### ���茋�ʂ̃N���A #################################
    TheDcTest.ClearContainer
    TheDcTest.ResultManager = TheResult
        
    '### DcLoopOption�ݒ�
    Call XLibDcScenarioLoopOption.ApplyDcScenarioLoopOptionMode

    Exit Sub
ErrHandler:
    InitControlShtReader
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    DisableAllTest
End Sub

Private Sub InitActiveDataSheet()
    '### JOB���X�g�}�l�[�W���̏��� ########################
    Set mJobListReader = CreateCDataSheetManager
    mJobListReader.Initialize JOB_LIST_TOOL
    '### �A�N�e�B�u�e�X�g�C���X�^���X�V�[�g�̎擾 #########
    Set mInstanceReader = CreateCInstanceSheetReader
    mInstanceReader.Initialize mJobListReader.GetActiveDataSht(TEST_INSTANCES_TOOL).Name
    '### �V�[�g�}�l�[�W���̏��� ###########################
    Set mDataManagerReader = CreateCDataSheetManager
    mDataManagerReader.Initialize SHEET_MANAGER_TOOL
    '### �A�N�e�B�uDC�V�i���I�V�[�g�̎擾 #################
    Dim activeSht As Worksheet
    Set activeSht = mDataManagerReader.GetActiveDataSht(DC_SCENARIO_TOOL)
    '### �A�N�e�B�uDC�V�i���I���[�_�[�̏��� ###############
    If activeSht Is Nothing Then
        Set mDcScenarioReader = Nothing
        Set mDcScenarioWriter = Nothing
    Else
        Set mDcScenarioReader = CreateCDcScenarioSheetReader
        mDcScenarioReader.Initialize activeSht.Name
        If mDcScenarioReader.AsIParameterReader.ReadAsBoolean(IS_VALIDATE) Then
            Set mDcLogReportWriter = New CDcLogReportWriter
            If XLibDcScenarioLoopOption.DcLoop = False Then mDcLogReportWriter.Initialize
        Else
            Set mDcLogReportWriter = Nothing
        End If
        Set mDcScenarioWriter = CreateCDcScenarioSheetLogWriter
        mDcScenarioWriter.Initialize activeSht.Name, GetSiteCount
    End If
    '### �A�N�e�B�uDC�Đ��f�[�^���[�_�[�̏��� #############
    Set activeSht = mDataManagerReader.GetActiveDataSht(DC_PLAYBACK_TOOL)
    If activeSht Is Nothing Then
        Set mDcReplayDataReader = Nothing
    Else
        Set mDcReplayDataReader = CreateCDcPlaybackSheetReader
        mDcReplayDataReader.Initialize activeSht.Name, GetSiteCount
    End If
    '### �I�t�Z�b�g���[�_�[�̏��� #########################
    Set activeSht = mDataManagerReader.GetActiveDataSht(OFFSET_TOOL)
    If activeSht Is Nothing Then
        Set mOffsetReader = Nothing
    Else
        Set mOffsetReader = CreateCOffsetSheetReader
        mOffsetReader.Initialize activeSht.Name, GetTesterNum, GetSiteCount
    End If
End Sub

Private Sub InitTheDcScenario()
    '### DC�e�X�g�V�i���I���s�G���W���̏����� #############
    If mDcScenarioReader Is Nothing Then
        Set TheDcTest = Nothing
    Else
        Dim dcPerformer As IDcTest
        If Not mDcReplayDataReader Is Nothing Then
            Dim replayDc As CPlaybackDc
            Set replayDc = CreateCPlaybackDc
            replayDc.Initialize mDcReplayDataReader
            Set dcPerformer = replayDc
        Else
            Set dcPerformer = CreateVISConnector '�d���ݒ��VIS�N���X���g�p����
'            Set dcPerformer = CreateCStdDCLibV01
        End If
        Set TheDcTest = CreateCDCScenario
        TheDcTest.Initialize dcPerformer, mDcScenarioReader, mInstanceReader, mDcScenarioWriter, mDcLogReportWriter
    End If
End Sub

Private Sub InitTheOffsetResult()
    '### �I�t�Z�b�g�}�l�[�W���̏����� #####################
    If mOffsetReader Is Nothing Then
        Set TheOffsetResult = Nothing
    Else
        Set TheOffsetResult = CreateCOffsetManager
        TheOffsetResult.Initialize mOffsetReader
    End If
End Sub

Private Sub ReInitTheDcScenario()
    '### DC�e�X�g�V�i���I���s�G���W���̏����� #############
    If mDcScenarioReader Is Nothing Then
        Set TheDcTest = Nothing
        Exit Sub
    Else
        With mDcScenarioReader
            If .AsIParameterReader.ReadAsBoolean(DATA_CHANGED) Then
                TheDcTest.Load
                mDcScenarioWriter.AsIActionStream.Rewind
            End If
            If Not mDcLogReportWriter Is Nothing Then
                If .AsIParameterReader.ReadAsBoolean(IS_VALIDATE) Then
                    If XLibDcScenarioLoopOption.DcLoop = False Then mDcLogReportWriter.Initialize
                Else
                    Set mDcLogReportWriter = Nothing
                End If
            End If
        End With
    End If
End Sub

Public Sub CloseDcLogReportWriter()
    If Not mDcLogReportWriter Is Nothing Then
        If XLibDcScenarioLoopOption.DcLoop = False Then mDcLogReportWriter.AsIFileStream.IsEOR
    End If
End Sub
