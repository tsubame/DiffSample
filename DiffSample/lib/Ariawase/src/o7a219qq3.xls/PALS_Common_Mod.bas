Attribute VB_Name = "PALS_Common_Mod"
Option Explicit

'==========================================================================
' ���W���[�����F  PALS_Common_mod.bas
' �T�v        �F  PALS�S�̂ŋ��ʂɎg�p����֐��Q
' ���l        �F  �Ȃ�
' �X�V����    �F  Rev1.0      2010/09/30�@�V�K�쐬        K.Sumiyashiki
'==========================================================================

'###########debug!!!!!!!!!!!!
'Public Const nSite As Long = 3
'Public Const Sw_Node As Long = 65
'Public Const g_MaxPalsCount As Long = 100
'###########debug!!!!!!!!!!!!


Public Declare Sub mSecSleep Lib "kernel32" Alias "Sleep" (ByVal lngmSec As Long)
Public Declare Function sub_PalsCopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public PALS As csPALS
Public blnRunPals As Boolean

Public Const PALSNAME As String = "PALS     ParameterAuto-adjustLinkSystem"
Public Const PALSVER As String = "1.50beta"

Public g_ErrorFlg_PALS As Boolean       '�p���X�ł̃G���[�������t���O(�G���[��������True�ɕύX)
Public Const PALS_ERRORTITLE As String = "- PALS Error -"

Public g_RunAutoFlg_PALS As Boolean       '�����A�g�@�\���쒆�t���O�@True:�����N�����AFalse�F�ʏ�N��

Public blnPALS_ANI  As Boolean
Public blnPALS_STOP As Boolean

Public Const PALS_OPTTARGET      As String = "OptTarget"
Public Const PALS_OPTIDENTIFIER  As String = "OptIdentifier"
Public Const PALS_OPTJUDGELIMIT  As String = "OptJudgeLimit"

Public Const PALS_LOOPCATEGORY1  As String = "CapCategory1"
Public Const PALS_LOOPCATEGORY2  As String = "CapCategory2"
Public Const PALS_LOOPJUDGELIMIT As String = "LoopJudgeLimit"

Public Const PALS_WAITADJFLG     As String = "WaitAdjFlg"

Public Const PALS_CHECKROW   As Integer = 1
Public Const PALS_CONTENTROW As Integer = 2
Public Const EXCEL_MAXCOLUMN As Integer = 256

Public PALS_ParamFolder As String
Public Const PALS_PARAMFOLDERNAME As String = "PALS_Params"
Public Const PALS_PARAMFOLDERNAME_VOLT As String = "PALS_Volt"
Public Const PALS_PARAMFOLDERNAME_WAVE As String = "PALS_Wave"
Public Const PALS_PARAMFOLDERNAME_WAIT As String = "PALS_Wait"
Public Const PALS_PARAMFOLDERNAME_OPT As String = "PALS_Opt"
Public Const PALS_PARAMFOLDERNAME_LOOP As String = "PALS_Loop"
Public Const PALS_PARAMFOLDERNAME_BIAS As String = "PALS_Bias"
Public Const PALS_PARAMFOLDERNAME_TRACE As String = "PALS_Trace"

Type PALS_TOOL_LIST
    PalsAdj     As Boolean
    VoltageAdj  As Boolean
    WaveAdj     As Boolean
    WaitAdj     As Boolean
    OptAdj      As Boolean
    LoopAdj     As Boolean
    BiasAdj     As Boolean
    TraceAdj     As Boolean
End Type


Public FLG_PALS_DISABLE As PALS_TOOL_LIST
Public FLG_PALS_RUN     As PALS_TOOL_LIST

'Public objLoadedJob     As Object

'****************************************
'****************  �萔  ****************
'****************************************
'�V�[�g���̒�`
Public Const FLOW_TABLE     As String = "Flow Table"
Public Const TEST_INSTANCES As String = "Test Instances"
Public Const TESTCONDITION  As String = "TestCondition"

'>>> 2011/5/6 M.Imamura
Public PinSheetname As String                               'Chans or ChannelMap SheetName
Public Const PinSheetnameChans = "Chans"                    'Chans SheetName
Public Const PinSheetnameChannel = "Channel Map"            'ChannelMap SheetName
'<<< 2011/5/6 M.Imamura

Public Const ReadSheetName = "Power-Supply Voltage"         'Read  SheetName
Public Const ReadSheetNameInfo = "Power-Supply Pin Info"    'Read  SheetName
Public Const OutPutSheetname = "Voltage Backup"             'Write SheetName

Public Const WaveSetupSheetName   As String = "WaveAdjustSetup"
Public Const WaveResultSheetName  As String = "WaveAdjustResult"

Public Const OptResultSheetName  As String = "OptAdjustResult"

Public Const WaitResultSheetName  As String = "WaitAdjustResult"

Public Const CONDSHTNAME As String = "ConditionSetTable"
Public Const ACQTBLSHTNAME As String = "Image ACQTBL"

Public Const intOscAdd            As Integer = 11

'FlowTable�ǂݎ��p�̒萔
'�ȍ~��FT��FlowTable�̗�
Public Const FT_LABEL_X      As Integer = 2
Public Const FT_START_Y      As Integer = 7
Public Const FT_TNAME_X      As Integer = 9
Public Const FT_OPCODE_X     As Integer = 7
Public Const FT_PARAMETER_X  As Integer = 8
Public Const FT_BIN_X        As Integer = 12
Public Const FT_TNUM_X       As Integer = 10
Public Const FT_LASTROW_NAME As String = "set-device"
Public Const FT_SURGE_NAME   As String = "D_SURGE"

'TestInstances�ǂݎ��p�̒萔
'�ȍ~��TI��TestInstances�̗�
Public Const TI_START_Y      As Integer = 6
Public Const TI_TESTNAME_X   As Integer = 2
Public Const TI_LOWLIMIT_X   As Integer = 14
Public Const TI_HIGHLIMIT_X  As Integer = 15
Public Const TI_UNIT_X       As Integer = 17
Public Const TI_CATEGORY1_X  As Integer = 19
Public Const TI_CATEGORY2_X  As Integer = 20
Public Const TI_JUDGELIMIT_X As Integer = 21
Public Const TI_ARG2_X       As Integer = 16

'TestCondition�ǂݎ��p�̒萔
'�ȍ~��TC��TestCondition�̗�
Public Const TC_START_Y         As Integer = 5
Public Const TC_CONDINAME_X     As Integer = 2
Public Const TC_PROCEDURENAME_X As Integer = 3
Public Const TC_ARG1_X          As Integer = 4
Public Const TC_SWNODE_X        As Integer = 1

'�f�[�^���O�̍ŏI�s������������
Public Const DATALOG_END As String = "========================================================================="

'�f�[�^���O�̍��ڈꗗ������������
Public Const DATALOG_INDEX As String = " Number  Site Result   Test Name       Pin       Channel Low            Measured       High           Force          Loc"
Public Const DATALOG_INDEX2 As String = " Number  Site Result   Test Name       Pin        Channel Low            Measured       High           Force          Loc"

'�f�[�^���O�̊e���ڈʒu����������ۂɎd�l���镶����
Public Const SITE_POSI      As String = "Site"
Public Const RESULT_POSI    As String = "Result"
Public Const TESTNAME_POSI  As String = "Test Name"
Public Const PIN_POSI       As String = "Pin"
Public Const MEASURED_POSI  As String = "Measured"
Public Const HIGH_POSI      As String = "High"
Public Const CHAN_POSI      As String = "Channel"

'>>>2011/10/3 M.IMAMURA �R���s���[�^�l�[���̌���INDEX
Public Const TESTERNAME_INDEX      As String = "      Node Name:"
'<<<2011/10/3 M.IMAMURA �R���s���[�^�l�[���̌���INDEX

'�f�[�^���O�̊e���ڈʒu�̕ۑ����s���\����
Public Type DatalogPosition
    SiteStart     As Integer    'site���X�^�[�g�ʒu
    SiteCount     As Integer    'site���̍ő啶����
    TestNameStart As Integer    '�e�X�g���̃X�^�[�g�ʒu
    TestNameCount As Integer    '�e�X�g���̍ő啶����
    MeasuredStart As Integer    '�����l�̃X�^�[�g�ʒu
    MeasuredCount As Integer    '�����l�̍ő啶����
    PinNameStart  As Integer    '[Pin]�̃X�^�[�g�ʒu
    PinNameCount  As Integer    '[Pin]�̍ő啶����
'>>>2011/05/12 K.SUMIYASHIKI ADD
    ResultStart  As Integer     '[Result(PASS/FAIL���)]�̃X�^�[�g�ʒu
    ResultCount  As Integer     '[Result(PASS/FAIL���)]�̍ő啶����
'<<<2011/05/12 K.SUMIYASHIKI ADD
End Type

'�P�ʊ��Z�p�W��
'Private Const TERA As Double = 1000000000000#      '�e��
'Private Const GIGA As Long = 1000000000            '�M�K
Private Const MEGA   As Long = 1000000              '���K
Private Const KIRO   As Long = 1000                 '�L��
Private Const MILLI  As Double = 0.001              '�~��
Private Const MAICRO As Double = 0.000001           '�}�C�N��
Private Const NANO   As Double = 0.000000001        '�i�m
Private Const PIKO   As Double = 0.000000000001     '�s�R
Private Const FEMTO  As Double = 0.000000000000001  '�t�F���g

'�f�[�^���O��(�t���p�X)
Public g_strOutputDataText As String


Private Const GC_NUMBER   As String = "Number"
Private Const GC_SITE     As String = "Site"
Private Const GC_RESULT   As String = "Result"
Private Const GC_TESTNAME As String = "Test Name"
Private Const GC_PIN      As String = "Pin"
Private Const GC_CHANNEL  As String = "Channel"
Private Const GC_LOW      As String = "Low"
Private Const GC_MEASURED As String = "Measured"
Private Const GC_HIGH     As String = "High"
Private Const GC_FORCE    As String = "Force"
Private Const GC_LOC      As String = "Loc"

Public Const SET_WAIT     As String = "xxSetWait"
Public Const SET_AVERAGE  As String = "xxSetAverage"
Public Const ACQUIRE_MODE As String = "xxAcquireMode"

Private colDataIndex As New Collection        '�f�[�^���O�����l�̃C���f�b�N�X�����p��������i�[����R���N�V����

'>>>2011/05/12 K.SUMIYASHIKI ADD
Public Type ActiveCheck
    Enable As Boolean                '�e�T�C�g�̏��(Active���ǂ���)���i�[���Ă���ϐ�(Active��True)�̒�`
End Type

Public Type ActiveSiteInformation
    site(nSite) As ActiveCheck       '�e�T�C�g�̏�Ԃ��i�[����\����(�T�C�g�����̔z��Œ�`)�̒�`
End Type

Public g_ActiveSiteInfo As ActiveSiteInformation    '�e�T�C�g�̏�Ԃ��i�[����\����
'<<<2011/05/12 K.SUMIYASHIKI ADD

Public CategoryData As csPALS_LoopMain    'csPALS_LoopMain�N���X�̒�`
'>>>2011/08/29 M.IMAMURA ADD
Public Const gblnForCis As Boolean = True
Public Flg_StopPMC_PALS As Boolean
Public Enum Enm_ErrFileBank
    Enm_ErrFileBank_LOCAL
    Enm_ErrFileBank_SERVER
End Enum
Public Const FILEBANK_LOCALPATH     As String = "C:\ERROR_LOG_"
'<<<2011/08/29 M.IMAMURA ADD

Public Sub Pause(interval As Single)
    Dim T1 As Single
    T1 = timer
    Do
        DoEvents
    Loop While timer - T1 < interval
End Sub

Public Sub RunPALS(Optional PalsRunNormal As Boolean = True)
    
On Error GoTo errPALSRunPALS
    
    PALS_ParamFolder = ThisWorkbook.Path & "\" & PALS_PARAMFOLDERNAME
    blnPALS_ANI = PalsRunNormal

    Call sub_PalsFileCheck

    If Sw_Node = 0 Then
        Call sub_errPALS("Sw_Node=0!! Please Check Your Condition!!", "0-2-01-5-02")
        Exit Sub
    End If

    '�p���X�̃G���[�t���O������
    '�p���X���ŃG���[�����������ꍇ�ATrue�ɕύX
    g_ErrorFlg_PALS = False
    
    'TOOL CHECK
    FLG_PALS_DISABLE.BiasAdj = True
    FLG_PALS_DISABLE.WaitAdj = True
    FLG_PALS_DISABLE.LoopAdj = True
    FLG_PALS_DISABLE.OptAdj = True
    FLG_PALS_DISABLE.TraceAdj = True
    FLG_PALS_DISABLE.VoltageAdj = True
    FLG_PALS_DISABLE.WaveAdj = True
    
    Set PALS = Nothing
    Set PALS = New csPALS
    
    If g_ErrorFlg_PALS Then
        End
    End If
    
    frm_PALS.Show
    
    g_ErrorFlg_PALS = False

Exit Sub

errPALSRunPALS:
    Call sub_errPALS("PALS initialize Failed at 'RunPALS'", "0-2-01-0-03")

End Sub

'Start Tester Run
Public Sub sub_exec_run()
    TheExec.RunTestProgram
End Sub

'Set Measure Condition [Do All]
Public Sub sub_exec_DoAll(blnDoAll As Boolean)
    TheExec.RunOptions.DoAll = blnDoAll
End Sub

'Do Optini
Public Sub sub_run_Optini()
    Call OptIni
End Sub



'********************************************************************************************
' ���O: sub_set_datalog
' ���e: IG-XL�̃f�[�^���O�ݒ���s��
' ����: blnSetLog      :True�ˊ֐����Ńf�[�^���O���̐ݒ���s��
'       blnSetLog      :False�ˊ֐����Ńf�[�^���O���̏��������s��
'       strFileHeader  :�ݒ�t�@�C�����̐擪�ɕt�^���镶����(��:LoopAdjData)
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_set_datalog(ByVal blnSetLog As Boolean, Optional strFileHeader1 As String = vbNullString, Optional strFileHeader2 As String = vbNullString)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_set_datalog

    With TheExec.Datalog
        If blnSetLog = True Then
            '�f�[�^���O���̐ݒ�(��:LoopAdjData_q7a163xa2_tool_debug_#65_20100927_170600.txt)
            g_strOutputDataText = PALS_ParamFolder & "\" & strFileHeader1 & "\" & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node) & "\" & strFileHeader2 & "_" & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & _
                                    "_#" & CStr(Sw_Node) & "_" & Format(Date, "yyyymmdd") & "_" & Format(TIME, "hhmmss") & ".txt"
            'Set Output txt Log
            .Setup.DatalogSetup.TextOutput = True
            'Set Output File
            .Setup.DatalogSetup.TextOutputFile = g_strOutputDataText
        Else
            'Set Output txt Log
            .Setup.DatalogSetup.TextOutput = False
            'Set Output File
            .Setup.DatalogSetup.TextOutputFile = vbNullString
            'Set EndLot
            .Setup.LotSetup.EndLot = True
            'Data Apply
            .ApplySetup
        End If
    End With

Exit Sub

errPALSsub_set_datalog:
    Call sub_errPALS("Set datalog name error at 'sub_set_datalog'", "0-2-02-0-04")

End Sub


'********************************************************************************************
' ���O: sub_ReadDatalog
' ���e: ��񕪂̑���f�[�^�̓ǂݎ����s��
' ����: lngNowLoopCnt :�����
'       intFileNo     :�I�[�v���t�@�C���̃t�@�C��No
'       DatalogPosi   :�f�[�^���O�̊e�����l�L���ʒu���i�[����\����
'       blnContFail   :Continue On Fail��Stop On Fail���𔻒f����t���O
'                      True ��FAIL���ڃf�[�^���ǂݎ��
'                      False��FAIL���ڃf�[�^�͏��O����
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/20�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_ReadDatalog(ByVal lngNowLoopCnt As Long, ByVal intFileNo As Integer, _
                            ByRef DatalogPosi As DatalogPosition, ByVal blnContFail As Boolean)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_ReadDatalog

    Dim strbuf As String           '�e�L�X�g�t�@�C������ǂݍ��񂾕�������i�[
    Dim blnFlgRead As Boolean      '�Ǎ�����p�t���O
    Dim blnFlgGetPos As Boolean    '�|�W�V��������p�t���O
    
    '�t���O������
    blnFlgRead = False
    blnFlgGetPos = False

    'DATALOG_END�Őݒ肳�ꂽ�s������܂ŌJ��Ԃ�
    Do Until blnFlgRead
        '�t�@�C������P�s�ǂݍ���
        Line Input #intFileNo, strbuf
        
        '���ڂ̑��莞�̂݁A�f�[�^���O�̊e���ڈʒu���m�F
        If (lngNowLoopCnt = 1) And (blnFlgGetPos = False) Then
            
            Call sub_InputDataIndex
            
            '�f�[�^���O���ڈꗗ�̍s�̏ꍇ�AstrBuf�Ɉ�s���̕�������i�[
'            Do While (strBuf <> DATALOG_INDEX And strBuf <> DATALOG_INDEX2)
            Do While Not sub_CheckDatalogIndex(strbuf)
                '>>>2011/10/3 M.IMAMURA �R���s���[�^�l�[���̎擾
                If InStr(1, strbuf, TESTERNAME_INDEX) > 0 Then
                    PALS.CommonInfo.g_strTesterName = Trim$(Mid(strbuf, InStr(1, strbuf, ":") + 1))
                End If
                '<<<2011/10/3 M.IMAMURA �R���s���[�^�l�[���̎擾
                Line Input #intFileNo, strbuf
            Loop
            
'            '��荞�ݍ��ڂ̈ʒu����
'            Call sub_GetDataPosition(strbuf, DatalogPosi)
            
            '�|�W�V��������p�t���O��True�ɕύX
            blnFlgGetPos = True
            
            '���s��strBuf�Ɋi�[
            Line Input #intFileNo, strbuf

            '��荞�ݍ��ڂ̈ʒu����
            Call sub_GetDataPosition(strbuf, DatalogPosi)
        End If
        
        '�f�[�^���O�̒l���擾
        If (Mid(strbuf, DatalogPosi.PinNameStart, 5) = "Empty") And (InStr(1, strbuf, "NGTEST") = 0) And _
                (InStr(1, strbuf, "WATCHS") = 0) And Len(strbuf) > 0 Then
            '�����l�擾�֐�
            Call sub_GetDatalogData(lngNowLoopCnt, strbuf, DatalogPosi, blnContFail)
        End If
        
        'DATALOG_END�Őݒ肳�ꂽ�s�ɓ��B������A�Ǎ�����p�t���O��True�ɕύX
        If strbuf = DATALOG_END Then
            blnFlgRead = True
        End If
    Loop

Exit Sub

errPALSsub_ReadDatalog:
    Call sub_errPALS("Read datalog data error at 'sub_ReadDatalog'", "0-2-03-0-05")

End Sub


'********************************************************************************************
' ���O: sub_InputDataIndex
' ���e: �f�[�^���O�����l�̃C���f�b�N�X����������ۂɎg�p���镶������R���N�V�����ɒǉ��B
'       �R���N�V�����́A�f�[�^���O��������l��ǂݎ��ۂɁA�C���f�b�N�X�����o����ׂɎg�p����B
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_InputDataIndex()

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

    '�f�[�^���O�����l�̃C���f�b�N�X�����p��������R���N�V�����ɒǉ�
    With colDataIndex
        .Add GC_NUMBER
        .Add GC_SITE
        .Add GC_RESULT
        .Add GC_TESTNAME
        .Add GC_PIN
        .Add GC_CHANNEL
        .Add GC_LOW
        .Add GC_MEASURED
        .Add GC_HIGH
        .Add GC_FORCE
        .Add GC_LOC
    End With

End Sub


'********************************************************************************************
' ���O: sub_CheckDatalogIndex
' ���e: �����œn���ꂽ������ɁA�R���N�V�������̑S�����񂪊܂܂�Ă��邩�`�F�b�N����B
'       �f�[�^���O�����l�̃C���f�b�N�X�s�����o����ׂɎg�p����B
' ����: strBuf  :�f�[�^���O��1�s
' �ߒl: True    :�R���N�V�������̕����񂪑S�Ċ܂܂�Ă���ꍇ
'     : False   :�R���N�V�������̕����񂪈�ł������Ă���ꍇ
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_CheckDatalogIndex(ByRef strbuf As String) As Boolean

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_CheckDatalogIndex

    '������
    sub_CheckDatalogIndex = False
        
    Dim varIndexName As Variant     'ForEach�Ŏg�p����ϐ�
    
    '�f�[�^���O�̃C���f�b�N�X�Ɋ܂܂�Ă��镶���񕪌J��Ԃ�
    For Each varIndexName In colDataIndex
        
        '�f�[�^���O�C���f�b�N�X�Ɋ܂܂�Ă��镶���񂪈�ł������Ă����False��Ԃ��A�֐��𔲂���
        If InStr(1, strbuf, varIndexName) = 0 Then
            Exit Function
        End If
    Next varIndexName

    '�S�Ċ܂܂�Ă���ꍇ�ATrue��Ԃ�
    sub_CheckDatalogIndex = True

Exit Function

errPALSsub_CheckDatalogIndex:
    Call sub_errPALS("Check DatalogIndex error at 'sub_CheckDatalogIndex'", "0-2-04-0-06")

End Function



'********************************************************************************************
' ���O: sub_GetDataPosition
' ���e: ����f�[�^�ʒu�̎擾���s��
' ����: strBuf      :��s���̃f�[�^���O(�f�[�^�̃C���f�b�N�X���܂܂�Ă���f�[�^)
'       DatalogPosi :�f�[�^���O�̊e�����l�L���ʒu���i�[����\����
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/20�@�V�K�쐬   K.Sumiyashiki
'            Rev1.1      2011/05/12�@�����ύX   K.Sumiyashiki
'                                    ��IG-XL�̃o�[�W������ResultFormat�̃f�[�^�ʒu���ω�����_�ɑΉ�
'********************************************************************************************
Private Sub sub_GetDataPosition(ByRef strbuf As String, ByRef DatalogPosi As DatalogPosition)
    
    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALSsub_GetDataPosition
    
    '�e���ڂ̈ʒu����
    
    Dim index(10) As Integer        '�e�f�[�^�̊J�n�ʒu���i�[����z��
    Dim IndexCnt As Integer         '�z��ԍ����C���N�������g����ׂ̕ϐ�
    '������
    IndexCnt = 0

    Dim blnCheckStart As Boolean    '�e�f�[�^���J�n���Ă��邩���f����t���O
    '������
    blnCheckStart = False

    Dim i As Long   'strbuf�̕�������������ۂɁA���������C���N�������g���Ă����ׂ�LOOP�J�E���^
    
    'strbuf�Ɋi�[����Ă��镶�����1�������A���ԂɌ���
    For i = 1 To Len(strbuf) - 1
        '�e�C���f�b�N�X�̊J�n�ʒu�𔻒f
        If blnCheckStart = False Then
            '�������������ʒu���󔒂łȂ��ӏ����A�e�f�[�^���n�܂����ʒu�Ɣ��f
            If Mid(strbuf, i, 1) <> " " Then
                blnCheckStart = True
                index(IndexCnt) = i
                IndexCnt = IndexCnt + 1
            End If

        '�A��2�������󔒂̏ꍇ�A�e�f�[�^���I������Ɣ��f
        ElseIf blnCheckStart = True Then
            If Mid(strbuf, i, 1) = " " And Mid(strbuf, i + 1, 1) = " " Then
                blnCheckStart = False
            End If
        End If
    Next i

'Index�z��̒��g
'Index(0) -> Number
'Index(1) -> Site
'Index(2) -> Result
'Index(3) -> Test Name
'Index(4) -> Pin
'Index(5) -> Channel
'Index(6) -> Low
'Index(7) -> Measured
'Index(8) -> High
'Index(9) -> Force
'Index(10)-> Loc

    With DatalogPosi
        .SiteStart = index(1)                               'site���X�^�[�g�ʒu
        .SiteCount = index(2) - .SiteStart                  'site���̍ő啶����
        .TestNameStart = index(3)                           '�e�X�g���̃X�^�[�g�ʒu
        .TestNameCount = index(4) - .TestNameStart          '�e�X�g���̍ő啶����
        .MeasuredStart = index(7)                           '�����l�̃X�^�[�g�ʒu
        .MeasuredCount = index(8) - .MeasuredStart          '�����l�̍ő啶����
        .PinNameStart = index(4)                            'Pin�̃X�^�[�g�ʒu
        .PinNameCount = index(5) - .PinNameStart            'Pin�̍ő啶����
'>>>2011/05/12 K.SUMIYASHIKI ADD
        .ResultStart = index(2)                             'Result(PASS/FAIL���)�̃X�^�[�g�ʒu
        .ResultCount = .TestNameStart - .ResultStart        'Result(PASS/FAIL���)�̍ő啶����
'<<<2011/05/12 K.SUMIYASHIKI ADD
    End With

Exit Sub

errPALSsub_GetDataPosition:
    Call sub_errPALS("Get Data Position error at 'sub_GetDataPosition'", "0-2-05-0-07")

End Sub


'********************************************************************************************
' ���O: sub_GetDatalogData
' ���e: �f�[�^���O��������l�̒��o���s���A�P�ʕϊ���A�ϐ��ɓ���
' ����: lngNowLoopCnt  :�����
'       strBuf         :��s���̃f�[�^���O(�����l���܂܂�Ă���f�[�^)
'       DatalogPosi    :�f�[�^���O�̊e�����l�L���ʒu���i�[����\����
'       blnContFail   :Continue On Fail��Stop On Fail���𔻒f����t���O
'                      True ��FAIL���ڃf�[�^���ǂݎ��
'                      False��FAIL���ڃf�[�^�͏��O����
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/20�@�V�K�쐬   K.Sumiyashiki
'            Rev1.1      2011/05/12�@�����ǉ�   K.Sumiyashiki
'                                    ��PASS/FAIL�̏��擾�����ǉ�
'********************************************************************************************
Private Sub sub_GetDatalogData(ByVal lngNowLoopCnt As Long, ByRef strbuf As String, ByRef DatalogPosi As DatalogPosition, ByVal blnContFail As Boolean)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_GetDatalogData

    Dim strTestName As String           'TestName(���ږ�)
    Dim strSiteNo As String             '�T�C�g�ԍ�
    Dim dblMeasured As Double           '�����l
    Dim intIndex As Integer             'LoopTestInfo�̃C���f�b�N�X
'>>>2011/05/12 K.SUMIYASHIKI ADD
    Dim strPassFail As String           'PASS/FAIL�̏��
'<<<2011/05/12 K.SUMIYASHIKI ADD

    With DatalogPosi
        '�e�X�g���̓ǂݎ��
        strTestName = RTrim$(Mid$(strbuf, .TestNameStart, .TestNameCount))
        If strTestName = FT_SURGE_NAME Then
            Exit Sub
        End If

        '�T�C�g���̓ǂݎ��
        strSiteNo = RTrim$(Mid$(strbuf, .SiteStart, .SiteCount))
                
        '�����l�̓ǂݎ��
        'sub_ConvertUnit�ŒP�ʊ��Z���s���Ă���(ex:"510 m" -> 0.51)
        dblMeasured = sub_ConvertUnit(RTrim$(Mid$(strbuf, .MeasuredStart, .MeasuredCount)))
        
        'TestnameInfoList�R���N�V�������g�p���A�Y���e�X�g���ڂ̃C���f�b�N�X���擾
        intIndex = PALS.CommonInfo.TestnameInfoList(strTestName)
        
'>>>2011/05/12 K.SUMIYASHIKI UPDATE
'>>>2011/06/16 M.IMAMURA blnContFail Add.
        'Stop On Fail�̏ꍇ�A���̃T�C�g�̗L���t���O��False�ɕύX���A�ȍ~�̃f�[�^��ǂݎ��Ȃ�
        If blnContFail = False Then
            If g_ActiveSiteInfo.site(CInt(strSiteNo)).Enable = False Then
                PALS.CommonInfo.TestInfo(intIndex).site(val(strSiteNo)).Enable(lngNowLoopCnt) = False
                Exit Sub
            End If
        End If
'<<<2011/06/16 M.IMAMURA blnContFail Add.
        
        '�w�荀�ڂ�PASS���f�t���O����UTrue�ɏ�����
        PALS.CommonInfo.TestInfo(intIndex).site(val(strSiteNo)).Enable(lngNowLoopCnt) = True

        '�����l��������
        PALS.CommonInfo.TestInfo(intIndex).site(val(strSiteNo)).Data(lngNowLoopCnt) = val(dblMeasured)

        strPassFail = RTrim$(Mid$(strbuf, .ResultStart, .ResultCount))
        '�w�荀�ڂ�FAIL���Ă����ꍇ�A�L�����f�t���O��False�֕ύX
        If strPassFail = "FAIL" Then
            g_ActiveSiteInfo.site(CInt(strSiteNo)).Enable = False
'''            PALS.CommonInfo.TestInfo(intIndex).Site(Val(strSiteNo)).Enable(lngNowLoopCnt) = False
        Else
'''            PALS.CommonInfo.TestInfo(intIndex).Site(Val(strSiteNo)).Enable(lngNowLoopCnt) = True
        End If
'<<<2011/05/12 K.SUMIYASHIKI UPDATE
    End With

Exit Sub

errPALSsub_GetDatalogData:
    Call sub_errPALS("Get datalog data error at 'sub_GetDatalogData'", "0-2-06-0-08")

End Sub


'********************************************************************************************
' ���O: sub_ConvertUnit
' ���e: �����l�̒P�ʕϊ����s��
'       "0.51 m" => "0.00051"
' ����: strBuf  :�����l�f�[�^�@ex)"0.51 m"
' �ߒl: �P�ʕϊ���̓����l
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/20�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_ConvertUnit(ByRef strbuf As String) As Double

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_ConvertUnit

    Dim SplitData As Variant    '�������ꂽ��������i�[

    'strBuf���󔒂ŕ���
    'ex)"0.51 m" => SplitData(0)="0.51",SplitData(1)="m"
    SplitData = Split(strbuf, " ")
    
    '�P�ʊ��Z
    '(UBound�ɂ��邱�ƂŁA�󔒂�2�ȏ゠�����ꍇ�ł��Ή���)
    Select Case SplitData(UBound(SplitData))
        Case "M"
            sub_ConvertUnit = val(SplitData(0)) * MEGA
        
        Case "MV"
            sub_ConvertUnit = val(SplitData(0)) * MEGA
            
        Case "K"
            sub_ConvertUnit = val(SplitData(0)) * KIRO
        
        Case "KV"
            sub_ConvertUnit = val(SplitData(0)) * KIRO
        
        Case "V"
            sub_ConvertUnit = val(SplitData(0))
        
        Case "mV", "mW"
            sub_ConvertUnit = val(SplitData(0)) * MILLI
        
        Case "m"
            sub_ConvertUnit = val(SplitData(0)) * MILLI
                
        Case "u"
            sub_ConvertUnit = val(SplitData(0)) * MAICRO
        
        Case "uV", "uW"
            sub_ConvertUnit = val(SplitData(0)) * MAICRO
        
        Case "n"
            sub_ConvertUnit = val(SplitData(0)) * NANO
        
        Case "nV", "nW"
            sub_ConvertUnit = val(SplitData(0)) * NANO
        
        Case "p"
            sub_ConvertUnit = val(SplitData(0)) * PIKO
        
        Case "pV", "pW"
            sub_ConvertUnit = val(SplitData(0)) * PIKO
        
        Case "f"
            sub_ConvertUnit = val(SplitData(0)) * FEMTO
        
        Case "fV"
            sub_ConvertUnit = val(SplitData(0)) * FEMTO
        
        Case Else
            sub_ConvertUnit = val(val(SplitData(0)))
        
    End Select

Exit Function

errPALSsub_ConvertUnit:
    Call sub_errPALS("Convert Unit error at 'sub_ConvertUnit'", "0-2-07-0-09")

End Function


Public Function all_mod_export()
    Call make_moddir
    
    Call export_module("frm_PALS.frm")
    Call export_module("frm_PALS_BiasAdj_Main.frm")
    Call export_module("frm_PALS_LoopAdj_Main.frm")
    Call export_module("frm_PALS_OptAdj_Main.frm")
    Call export_module("frm_PALS_TraceAdj_Main.frm")
    Call export_module("frm_PALS_VoltAdj_Main.frm")
    Call export_module("frm_PALS_WaitAdj_Main.frm")
    Call export_module("frm_PALS_WaveAdj_Confirm.frm")
    Call export_module("frm_PALS_WaveAdj_Doing.frm")
    Call export_module("frm_PALS_WaveAdj_Main.frm")
    Call export_module("frm_PALS_WaveAdj_Warning.frm")
    
    Call export_module("Conditionset_Mod_ShutOnly.bas")
    Call export_module("PALS_BiasAdj_Mod.bas")
    Call export_module("PALS_Common_Mod.bas")
    Call export_module("PALS_LoopAdj_Mod.bas")
    Call export_module("PALS_OptAdj_Mod.bas")
    Call export_module("PALS_Sub_Mod.bas")
    Call export_module("PALS_TraceAcq_Mod.bas")
    Call export_module("PALS_TraceAdj_Mod.bas")
    Call export_module("PALS_VoltAdj_Mod.bas")
    Call export_module("PALS_WaitAdj_Mod.bas")

    Call export_module("PALS_WaveAdj_mod_Common.bas")
    Call export_module("PALS_WaveAdj_mod_GetWave.bas")
    Call export_module("PALS_WaveAdj_mod_H.bas")
    Call export_module("PALS_WaveAdj_mod_HShared.bas")
    Call export_module("PALS_WaveAdj_mod_LH.bas")
    Call export_module("PALS_WaveAdj_mod_RG.bas")
    Call export_module("PALS_WaveAdj_mod_Shutter.bas")
    Call export_module("PALS_WaveAdj_mod_TVCfunctions.bas")
    Call export_module("PALS_WaveAdj_mod_VVT.bas")

    Call export_module("PALS_IlluminatorMod.bas")

    Call export_module("csPALS.cls")
    Call export_module("csPALS_Common.cls")
    Call export_module("csPALS_LoopCategoryParams.cls")
    Call export_module("csPALS_LoopMain.cls")
    Call export_module("csPALS_OptCond.cls")
    Call export_module("csPALS_OptCondParams.cls")
    Call export_module("csPALS_TestInfo.cls")
    Call export_module("csPALS_TestInfoParams.cls")

    Call export_module("csPALS_WaveACPSet.cls")
    Call export_module("csPALS_WaveACSet.cls")
    Call export_module("csPALS_WaveAdjust.cls")
    Call export_module("csPALS_WaveDCPSet.cls")
    Call export_module("csPALS_WaveDcSet.cls")
    Call export_module("csPALS_WaveDevicePin.cls")
    Call export_module("csPALS_WaveOscPSet.cls")
    Call export_module("csPALS_WaveOscSet.cls")
    Call export_module("csPALS_WaveResource.cls")


End Function


Public Sub make_moddir()
    Dim modfir As String

    On Error Resume Next

    modfir = ActiveWorkbook.Path & "\bas"
    MkDir modfir

    modfir = ActiveWorkbook.Path & "\frm"
    MkDir modfir

    modfir = ActiveWorkbook.Path & "\cls"
    MkDir modfir

    On Error GoTo 0
End Sub
Public Sub export_module(mymodule As String)

    Dim waveI As Long
    Dim mytype As Integer
    Dim modfir As String
    
    With Workbooks(ActiveWorkbook.Name).VBProject
        'check all project
        For waveI = 1 To .VBComponents.Count
            'delete pin point!!
            'check type & name
            If Right(mymodule, 3) = "bas" Then
                mytype = 1
                modfir = ActiveWorkbook.Path & "\bas"
            End If
            If Right(mymodule, 3) = "cls" Then
                mytype = 2
                modfir = ActiveWorkbook.Path & "\cls"
            End If
            If Right(mymodule, 3) = "frm" Then
                mytype = 3
                modfir = ActiveWorkbook.Path & "\frm"
            End If
            If .VBComponents(waveI).Name = Left(mymodule, Len(mymodule) - 4) And .VBComponents(waveI).Type = mytype Then
                'delete compornents
                If mytype = 1 Then .VBComponents(waveI).Export modfir & "\" & .VBComponents(waveI).Name & ".bas"
                If mytype = 2 Then .VBComponents(waveI).Export modfir & "\" & .VBComponents(waveI).Name & ".cls"
                If mytype = 3 Then .VBComponents(waveI).Export modfir & "\" & .VBComponents(waveI).Name & ".frm"
                Exit For
            End If
        Next waveI
    End With
    
End Sub

'********************************************************************************************
' ���O: sub_ModuleCheck
' ���e: �����œn���ꂽ���W���[�������݂��邩�`�F�b�N���s��
' ����: mymodule :�g���q�t�����W���[����
' �ߒl: True  : ��v����
'       False : ��v�Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/10/29�@�V�K�쐬   M.Imamura
'********************************************************************************************
Public Function sub_ModuleCheck(mymodule As String) As Boolean

    Dim vbc As Object
    Dim lngLoopCnt As Long
    Dim mytype As Integer
    
    sub_ModuleCheck = False
    
    If Right(mymodule, 3) = "bas" Then mytype = 1
    If Right(mymodule, 3) = "cls" Then mytype = 2
    If Right(mymodule, 3) = "frm" Then mytype = 3

    With Workbooks(ActiveWorkbook.Name).VBProject
        For lngLoopCnt = 1 To .VBComponents.Count
            'check type & name
            If .VBComponents(lngLoopCnt).Name = Left(mymodule, Len(mymodule) - 4) And .VBComponents(lngLoopCnt).Type = mytype Then
                sub_ModuleCheck = True
                Exit Function
            End If
        Next lngLoopCnt
    End With
    
    
End Function

'********************************************************************************************
' ���O: sub_SheetNameCheck
' ���e: �����œn���ꂽ���O�̃V�[�g�����݂��邩�`�F�b�N���s��
' ����: strSheetName :�����V�[�g��
' �ߒl: True  : ��v�V�[�g����
'       False : ��v�V�[�g�Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Function sub_SheetNameCheck(ByVal strSheetName As String) As Boolean

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_SheetNameCheck
    
    sub_SheetNameCheck = False

    Dim objWorkSheet As Worksheet
    
    For Each objWorkSheet In Worksheets
        If strSheetName = objWorkSheet.Name Then
            sub_SheetNameCheck = True
            Exit For
        End If
    Next

Exit Function

errPALSsub_SheetNameCheck:
    Call sub_errPALS("SheetName check error at 'sub_SheetNameCheck'", "0-2-08-0-10")

End Function


'********************************************************************************************
' ���O: sub_InitCollection
' ���e: �R���N�V�����f�[�^�̏�����
' ����: col:�R���N�V����
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_InitCollection(ByRef col As Collection)

    If g_ErrorFlg_PALS Or col.Count = 0 Then
        Exit Sub
    End If

On Error GoTo errPALSsub_InitCollection
    
    Dim i As Long   '���[�v�J�E���^

    '�R���N�V�����f�[�^�̍폜
    For i = col.Count To 1 Step -1
        col.Remove (i)
    Next i

Exit Sub

errPALSsub_InitCollection:
    Call sub_errPALS("Collection remove error at 'sub_InitCollection'", "0-2-09-0-11")

End Sub

'********************************************************************************************
' ���O: sub_errPALS
' ���e: �G���[�\���y�уG���[���O�쐬
' ����: strPalsErrMsg:�G���[�ڍ׏��
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/10/19�@�V�K�쐬   M.Imamura
' �X�V�����F Rev1.1      2011/05/06�@�G���[�R�[�h��ǉ�   M.Imamura
' �X�V�����F Rev1.2      2012/03/14�@���[�J��PC�ւ̕ۑ���ǉ�   M.Imamura
'********************************************************************************************

Public Sub sub_errPALS(ByVal strPalsErrMsg As String, Optional strPalsErrorCode As String = "", Optional enumFileBank As Enm_ErrFileBank = Enm_ErrFileBank_SERVER)
    Dim strPalsErrDescription As String
    
    'PALS�G���[�t���O����
    g_ErrorFlg_PALS = True
    
    '���b�Z�[�W�{�b�N�X�\��
    If Len(strPalsErrorCode) > 0 Then strPalsErrorCode = " " & strPalsErrorCode
    If Err.Number = 0 Then
        'PALS�Z���t�`�F�b�N���̃R�����g
        strPalsErrDescription = "PALS Check Error"
    Else
        'VBA�G���[�̏ڍ�
        strPalsErrDescription = Err.Description
    End If
    
    '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    If g_RunAutoFlg_PALS = False Then
        MsgBox "Error" & strPalsErrorCode & " : " & strPalsErrMsg & vbCrLf & "Description : " & strPalsErrDescription, vbExclamation, PALS_ERRORTITLE
    End If
    '<<<2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    
    '�G���[���O�f���o��
    Dim intFileNo As Integer
    Dim strOutputDataText As String
    intFileNo = FreeFile

    If enumFileBank = Enm_ErrFileBank_SERVER Then
        If PALS_ParamFolder = "" Then PALS_ParamFolder = ThisWorkbook.Path
        strOutputDataText = PALS_ParamFolder & "\PALS_ERROR_LOG_" & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & _
                                        "_#" & CStr(Sw_Node) & ".txt"
    ElseIf enumFileBank = Enm_ErrFileBank_LOCAL Then '�I�[�g�����~�p
        Flg_StopPMC_PALS = True 'StopPMC�Ŏ~�߂�
        strOutputDataText = FILEBANK_LOCALPATH & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node) & ".txt"
    End If
    
    Open strOutputDataText For Append As #intFileNo

    Print #intFileNo, "--------------------------------------------"
    Print #intFileNo, "Date         : " & Format(Date, "yyyy/mm/dd") & " " & Format(TIME, "hh:mm:ss")
    Print #intFileNo, "Message      : " & strPalsErrMsg
    Print #intFileNo, "Description  : " & strPalsErrDescription
    Print #intFileNo, "ErrorCode    : " & strPalsErrorCode
    
    Print #intFileNo, ""
    
    Close #intFileNo

End Sub
Public Sub sub_PalsFileCheck(Optional ByVal strPalsCheckDir As String = "")
    Dim strPalsCheckFile As String
    Dim intFileCount As Integer
    Dim intFileNo As Integer

On Error Resume Next
    strPalsCheckFile = PALS_ParamFolder
    If strPalsCheckDir <> "" Then strPalsCheckFile = strPalsCheckFile & "\" & strPalsCheckDir
    
'    With Application.FileSearch
'        .LookIn = strPalsCheckFile
'        .filename = "*.*"
'        .Execute
'        intFileCount = .FoundFiles.count
'
'        If .FoundFiles.count = 0 Then
            MkDir strPalsCheckFile
'        End If
'
'    End With
        
    'Write RunningLogData
    intFileNo = FreeFile
    If strPalsCheckDir <> "" Then
        Open strPalsCheckFile & "\" & strPalsCheckDir & ".log" For Append As #intFileNo
        Print #intFileNo, Format(Date, "yyyy/mm/dd") & " " & Format(TIME, "hh:mm:ss") & " " & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node)
        Close #intFileNo
        strPalsCheckFile = strPalsCheckFile & "\" & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node)
        MkDir strPalsCheckFile
    Else
        Open strPalsCheckFile & "\" & PALS_PARAMFOLDERNAME & ".log" For Append As #intFileNo
        Print #intFileNo, Format(Date, "yyyy/mm/dd") & " " & Format(TIME, "hh:mm:ss") & " " & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node)
        Close #intFileNo
    End If
    

On Error GoTo 0

End Sub


'********************************************************************************************
' ���O: sub_TestingStatusOutPals
' ���e: �t�H�[���̃t�H���g�F��ύX
' ����: objPalsForm:�t�H�[���I�u�W�F�N�g
'       strPalsMsg:�\�����镶����
'       RedColor:True�Ńt�H���g��(�f�t�H���g:False)
'       BlueColor:True�Ńt�H���g��(�f�t�H���g:False)
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_TestingStatusOutPals(objPalsForm As Object, strPalsMsg As String, Optional RedColor As Boolean = False, Optional BlueColor As Boolean = False)

    On Error GoTo errPALSsub_TestingStatusOutPals
    
    If RedColor = True Then
        objPalsForm.lblProcess.ForeColor = vbRed
    ElseIf BlueColor = True Then
        objPalsForm.lblProcess.ForeColor = vbBlue
    Else
        objPalsForm.lblProcess.ForeColor = vbBlack
    End If
    
    objPalsForm.lblProcess.Caption = strPalsMsg
    DoEvents

    Exit Sub

errPALSsub_TestingStatusOutPals:
    Call sub_errPALS("Status change error at 'sub_TestingStatusOutPals'", "0-2-10-0-12")

End Sub


'********************************************************************************************
' ���O: sub_InitActiveSiteInfo
' ���e: �e�T�C�g��PASS/FAIL�����������B�S��True(Active���)�ɏ������B
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2011/05/16�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_InitActiveSiteInfo()

    Dim sitez As Long

    With g_ActiveSiteInfo
        '�S�T�C�g���J��Ԃ�
        For sitez = LBound(.site) To UBound(.site)
            '�e�T�C�g��Active��Ԃŏ�����
            .site(sitez).Enable = True
        Next sitez
    End With

End Sub

'********************************************************************************************
' ���O: ReadCategoryData
' ���e: csPALS_LoopMain�̃C���X�^���X�𐶐�
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l: �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
' �X�V�����F Rev1.1      2011/07/29�@PALS_Sub����ړ�   M.Imamura
'********************************************************************************************
Public Sub ReadCategoryData()

'>>>2011/06/02 K.SUMIYASHIKI ADD
    g_ErrorFlg_PALS = False
'<<<2011/06/02 K.SUMIYASHIKI ADD

    Set CategoryData = Nothing
    Set CategoryData = New csPALS_LoopMain

End Sub
'********************************************************************************************
' ���O: ResetPals
' ���e: PALS�p�����[�^���ēǂݍ���
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l: �Ȃ�
' �X�V�����F Rev1.0      2011/12/06�@�V�K�쐬   M.Imamura
'********************************************************************************************
Public Sub ResetPals(Optional ByVal strResetMode As String = "ALL")

    
    Select Case strResetMode
        Case "PALS"
            Set PALS = Nothing
            Set PALS = New csPALS
        Case "ALL"
            Set OptCond = Nothing
            Set OptCond = New csPALS_OptCond
            
            Set CategoryData = Nothing
            Set CategoryData = New csPALS_LoopMain
        
            Call Get_Power_Condition
        
            Call Excel.Application.Run("TimingInit_PALS")
        
        Case "OPT"
            Set OptCond = Nothing
            Set OptCond = New csPALS_OptCond

        Case "TESTCOND"
            Set CategoryData = Nothing
            Set CategoryData = New csPALS_LoopMain

        Case "VOLT"
            Call Get_Power_Condition
        
        Case "TIME"
            Call Excel.Application.Run("TimingInit_PALS")

    End Select
End Sub

Public Function sub_CheckResultFormat() As Boolean
  sub_CheckResultFormat = True
  If TheExec.Datalog.Setup.DatalogSetup.SelectSetupFile = False Then
    sub_CheckResultFormat = False
    Call sub_errPALS("ResultFormat is NotChecked!! Please Select ResultFormat", "0-2-11-2-12")
  End If
End Function
Public Sub sub_OutPutCsv(ByVal InputWorkSheetName As String, ByVal OutPutCSVFName As String, Optional ByVal bln_ShowMsg As Boolean = True)

' Excel Application/Book/Sheet�I�u�W�F�N�g��`
    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlWS As Excel.Worksheet

    Dim intMaxRow As Integer
    Dim intMaxColumn As Integer
    
    Dim intWriteRow As Integer
    Dim intWriteColumn As Integer
    
    '2012/7/9 FullPath Get. M.Imamura
    Dim strOutPutCSVFName_Full As String
    
    If Left(OutPutCSVFName, 2) = "./" Or Left(OutPutCSVFName, 2) = ".\" Then
        strOutPutCSVFName_Full = ThisWorkbook.Path & Mid(OutPutCSVFName, 2, Len(OutPutCSVFName) - 1)
    Else
        strOutPutCSVFName_Full = OutPutCSVFName
    End If
    
    intMaxRow = Worksheets(InputWorkSheetName).UsedRange.Rows.Count
    intMaxColumn = Worksheets(InputWorkSheetName).UsedRange.Columns.Count
    
    On Error GoTo errPALSsub_OutPutCsv

    'BackUp CSV FIle
    '2012/7/9 FunctionNameChanged. M.Imamura FileCopy -> sub_PalsFileCopy
    If sub_PalsFileCopy(strOutPutCSVFName_Full, strOutPutCSVFName_Full & "_" & Format(Date, "yyyymmdd") & "_" & Format(TIME, "hhmmss")) = False Then
        GoTo errPALSsub_OutPutCsv
    End If

    Set xlApp = CreateObject("Excel.Application")

    xlApp.DisplayAlerts = False
    xlApp.Visible = False
    xlApp.ScreenUpdating = False

    Set xlWB = xlApp.Workbooks.Open(strOutPutCSVFName_Full)
    Set xlWS = xlWB.Worksheets(1)

    xlWB.Worksheets(1).Cells.Select
    xlApp.Selection.ClearContents

    For intWriteRow = 1 To intMaxRow
        For intWriteColumn = 1 To intMaxColumn
            xlWS.Cells(intWriteRow, intWriteColumn).Value = Worksheets(InputWorkSheetName).Cells(intWriteRow, intWriteColumn).Value
        Next intWriteColumn
    Next intWriteRow
    xlWS.Cells(1, 1).Font.color = vbRed
    xlWS.Cells(3, 2).Value = "location:"

    xlWB.SaveAs strOutPutCSVFName_Full, xlCSV

    xlWB.Close
    xlApp.Quit

    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    
    If bln_ShowMsg = True Then
        MsgBox "PALS saved [" & InputWorkSheetName & "]Sheet to " & vbCrLf & "  CSV[" & strOutPutCSVFName_Full & "]", vbOKOnly, PALSNAME
    End If
Exit Sub
    
errPALSsub_OutPutCsv:
    Call sub_errPALS("CsvOutPutError " & strOutPutCSVFName_Full & " at 'sub_OutPutCsv'", "0-5-05-6-37")

    If Not (xlWB Is Nothing) Then xlWB.Close
    Set xlWS = Nothing
    Set xlWB = Nothing
    If Not (xlApp Is Nothing) Then xlApp.Quit
    Set xlApp = Nothing
    
End Sub

'2012/7/9 FunctionNameChanged. M.Imamura FileCopy -> sub_PalsFileCopy
Public Function sub_PalsFileCopy(tgtFile As String, newFile As String) As Boolean
    Dim RetNum As Long
    RetNum = sub_PalsCopyFile(tgtFile, newFile, True) 'Arg3 True means overwrite, False means preserve.
    If RetNum Then
        sub_PalsFileCopy = True
    Else
        sub_PalsFileCopy = False
    End If
End Function
