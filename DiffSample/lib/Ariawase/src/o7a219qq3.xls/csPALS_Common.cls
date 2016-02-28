VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' �N���X��  �F  csPALS_Common.cls
' �T�v      �F  PALS�ŋ��ʂɎg�p����������N���X
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/16�@�V�K�쐬        K.Sumiyashiki
'==========================================================================

Private xTestInfo() As csPALS_TestInfo          'csPALS_TestInfo�N���X�̒�`

Public TestnameInfoList As New Collection       'Tname�ƃC���f�b�N�X��R�t����ׂ̃R���N�V����

Public g_strTesterName As String                '���胍�O��Node Name:�ȍ~�̃e�X�^�[��

'�e�X�g�C���X�^���X�̃��[�v�����Ɋւ���p�����[�^�̑��݃`�F�b�N�ϐ�
Private Type PALS_LoopParameterList
    Flg_CapCategory1   As Boolean
    Flg_CapCategory2   As Boolean
    Flg_LoopJudgeLimit As Boolean
    
    CapCategory1_X   As Integer
    CapCategory2_X   As Integer
    LoopJudgeLimit_X As Integer
End Type

'�e�X�g�C���X�^���X�̌��ʒ����Ɋւ���p�����[�^�̑��݃`�F�b�N�ϐ�
Private Type PALS_OptParameterList
    Flg_OptTarget     As Boolean
    Flg_OptIdentifier As Boolean
    Flg_OptJudgeLimit As Boolean

    OptTarget_X     As Integer
    OptIdentifier_X As Integer
    OptJudgeLimit_X As Integer
End Type
'�e�X�g�C���X�^���X��Wait�����Ɋւ���p�����[�^�̑��݃`�F�b�N�ϐ�
Private Type PALS_WaitParameterList
    Flg_CapCategory1   As Boolean
    Flg_CapCategory2   As Boolean
    Flg_WaitAdjFlg     As Boolean
    
    CapCategory1_X   As Integer
    CapCategory2_X   As Integer
    WaitAdjFlg_X     As Integer
End Type

Private m_LoopParameter As PALS_LoopParameterList
Private m_OptParameter  As PALS_OptParameterList
Private m_WaitParameter  As PALS_WaitParameterList


'*********************************************************************************
' �֐���    �F  Class_Initialize
' �T�v      �F  �C���X�^���X�������̏�����
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()

    '�R���N�V�����f�[�^�̏�����
    Call sub_InitCollection(TestnameInfoList)
    
    '�e�X�g�C���X�^���X��2�s�ڂɋL�ڂ���Ă���PALS�e�p�����[�^���̑��݃`�F�b�N
    Call sub_CheckPALSParameter

    '�t���[�e�[�u���ƃe�X�g�C���X�^���X�̏��擾
    Call sub_ReadTestInfo
    
End Sub


'*********************************************************************************
' �֐���    �F  Class_Terminate
' �T�v      �F  �C���X�^���X�J�����̏���
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Terminate()
    Call sub_InitCollection(TestnameInfoList)
    Erase xTestInfo
End Sub


'*********************************************************************************
' �v���p�e�B��  �F  TestInfoName
' �T�v          �F  ���ڔԍ��ɑΉ�����TestInfoName�I�u�W�F�N�g��Ԃ��v���p�e�B
' ����          �F  index :���ږ��������C���f�b�N�X
' ���l          �F  �ǂݎ���p
' �X�V����      �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Public Property Get TestInfo(ByVal index As Integer) As csPALS_TestInfo
    Set TestInfo = xTestInfo(index)
End Property


'********************************************************************************************
' ���O: sub_CheckPALSParameter
' ���e: �e�X�g�C���X�^���X��2�s�ڂɋL�ڂ���Ă���PALS�e�p�����[�^���̃`�F�b�N
'     �@�������ꍇ�A�e�p�����[�^�̃t���O��True�ɕύX
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_CheckPALSParameter()
    
    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALSsub_CheckPALSParameter
    
    '��v����p�����[�^���e�X�g�C���X�^���X�ɑ��݂���΁A�t���O��True�ɕύX
    Const strRange As String = "A2:IV2"
    
    
   '*******LOOP TOOL******************************************************************************************************
    With m_LoopParameter
        '�w��V�[�g��v����p�����[�^���������ꍇ�A�e�p�����[�^�̑��݃t���O��True�ɕύX
        .Flg_CapCategory1 = sub_FindPALSParameter(PALS_LOOPCATEGORY1, TEST_INSTANCES, strRange, .CapCategory1_X, True)
        .Flg_CapCategory2 = sub_FindPALSParameter(PALS_LOOPCATEGORY2, TEST_INSTANCES, strRange, .CapCategory2_X, True)
        .Flg_LoopJudgeLimit = sub_FindPALSParameter(PALS_LOOPJUDGELIMIT, TEST_INSTANCES, strRange, .LoopJudgeLimit_X, True)
        
        
        '���[�v�c�[���ɕK�v�ȃp�����[�^���S�đ��݂����ꍇ�A���C���t�H�[����LOOP�{�^����L���ɂ���
        If sub_ModuleCheck("frm_PALS_LoopAdj_Main.frm") = True And (.Flg_CapCategory1 = True) And (.Flg_CapCategory2 = True) And (.Flg_LoopJudgeLimit = True) Then
            FLG_PALS_DISABLE.LoopAdj = False
        Else
            If sub_ModuleCheck("frm_PALS_WaitAdj_Main.frm") = False Then
                frm_PALS.cmd_LoopRun.ControlTipText = frm_PALS.cmd_LoopRun.ControlTipText & "NotFound [frm_PALS_WaitAdj_Main.frm]Form" & vbCrLf
            End If
            frm_PALS.cmd_LoopRun.ControlTipText = frm_PALS.cmd_LoopRun.ControlTipText & "Not Found Parameter at [" & TEST_INSTANCES & "]" & vbCrLf
            If .Flg_CapCategory1 = False Then
                frm_PALS.cmd_LoopRun.ControlTipText = frm_PALS.cmd_LoopRun.ControlTipText & "- CapCategory1 -" & vbCrLf
            End If
            If .Flg_CapCategory2 = False Then
                frm_PALS.cmd_LoopRun.ControlTipText = frm_PALS.cmd_LoopRun.ControlTipText & "- CapCategory2 -" & vbCrLf
            End If
            If .Flg_LoopJudgeLimit = False Then
                frm_PALS.cmd_LoopRun.ControlTipText = frm_PALS.cmd_LoopRun.ControlTipText & "- LoopJudgeLimit -" & vbCrLf
            End If
        End If
    End With
   '**********************************************************************************************************************

   '*******Wait TOOL******************************************************************************************************
    With m_WaitParameter
        '�w��V�[�g��v����p�����[�^���������ꍇ�A�e�p�����[�^�̑��݃t���O��True�ɕύX
        .Flg_CapCategory1 = sub_FindPALSParameter(PALS_LOOPCATEGORY1, TEST_INSTANCES, strRange, .CapCategory1_X, True)
        .Flg_CapCategory2 = sub_FindPALSParameter(PALS_LOOPCATEGORY2, TEST_INSTANCES, strRange, .CapCategory2_X, True)
        .Flg_WaitAdjFlg = sub_FindPALSParameter(PALS_WAITADJFLG, TEST_INSTANCES, strRange, .WaitAdjFlg_X, True)
        
        
        'Wait�c�[���ɕK�v�ȃp�����[�^���S�đ��݂����ꍇ�A���C���t�H�[����Wait�{�^����L���ɂ���
        If sub_ModuleCheck("frm_PALS_WaitAdj_Main.frm") = True And .Flg_CapCategory1 = True And .Flg_CapCategory2 = True And .Flg_WaitAdjFlg = True And sub_SheetNameCheck(WaitResultSheetName) = True And ((.Flg_CapCategory1 = True) And (.Flg_CapCategory2 = True)) Then
            FLG_PALS_DISABLE.WaitAdj = False
        Else
            If sub_ModuleCheck("frm_PALS_WaitAdj_Main.frm") = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "NotFound [frm_PALS_WaitAdj_Main.frm]Form" & vbCrLf
            End If
            frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "Not Found Parameter at [" & TEST_INSTANCES & "]" & vbCrLf
            If .Flg_CapCategory1 = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- CapCategory1 -" & vbCrLf
            End If
            If .Flg_CapCategory2 = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- CapCategory2 -" & vbCrLf
            End If
            If .Flg_WaitAdjFlg = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- WaitAdjFlg -" & vbCrLf
            End If
            If sub_SheetNameCheck(WaitResultSheetName) = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "NotFound [" & WaitResultSheetName & "]Sheet" & vbCrLf
            End If
            If .Flg_CapCategory1 = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- CapCategory1 -" & vbCrLf
            End If
            If .Flg_CapCategory2 = False Then
                frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- CapCategory2 -" & vbCrLf
            End If
        End If
    End With
   '**********************************************************************************************************************

   '*******OPT TOOL*******************************************************************************************************
    With m_OptParameter
        '�w��V�[�g��v����p�����[�^���������ꍇ�A�e�p�����[�^�̑��݃t���O��True�ɕύX
        .Flg_OptTarget = sub_FindPALSParameter(PALS_OPTTARGET, TEST_INSTANCES, strRange, .OptTarget_X, True)
        .Flg_OptIdentifier = sub_FindPALSParameter(PALS_OPTIDENTIFIER, TEST_INSTANCES, strRange, .OptIdentifier_X, True)
        .Flg_OptJudgeLimit = sub_FindPALSParameter(PALS_OPTJUDGELIMIT, TEST_INSTANCES, strRange, .OptJudgeLimit_X, True)

        '���ʒ����c�[���ɕK�v�ȃp�����[�^���S�đ��݂����ꍇ�A���C���t�H�[����OPT�{�^����L���ɂ���
        If sub_ModuleCheck("frm_PALS_OptAdj_Main.frm") = True And sub_SheetNameCheck(OptResultSheetName) = True And ((.Flg_OptTarget = True) And (.Flg_OptIdentifier = True) And (.Flg_OptJudgeLimit = True)) Then
            FLG_PALS_DISABLE.OptAdj = False
        Else
            If sub_SheetNameCheck(OptResultSheetName) = False Then
                frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "NotFound [" & OptResultSheetName & "]Sheet" & vbCrLf
            End If
            If sub_ModuleCheck("frm_PALS_OptAdj_Main.frm") = False Then
                frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "NotFound [frm_PALS_OptAdj_Main.frm]Form" & vbCrLf
            End If
            frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "Not Found Parameter at [" & TEST_INSTANCES & "]" & vbCrLf
            If .Flg_OptTarget = False Then
                frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptTarget -" & vbCrLf
            End If
            If .Flg_OptIdentifier = False Then
                frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptIdentifier -" & vbCrLf
            End If
            If .Flg_OptJudgeLimit = False Then
                frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptJudgeLimit -" & vbCrLf
            End If
        End If
    End With
   '**********************************************************************************************************************

   '*******VOLT TOOL*******************************************************************************************************
    '>>> 2011/5/6 M.Imamura
    If sub_SheetNameCheck(PinSheetnameChans) = True Then
        PinSheetname = PinSheetnameChans
    Else
        PinSheetname = PinSheetnameChannel
    End If
    '<<< 2011/5/6 M.Imamura
    If sub_SheetNameCheck(PinSheetname) = True And sub_SheetNameCheck(OutPutSheetname) = True And sub_SheetNameCheck(ReadSheetName) = True And sub_SheetNameCheck(ReadSheetNameInfo) = True And sub_ModuleCheck("PALS_VoltAdj_Mod.bas") = True And sub_ModuleCheck("frm_PALS_VoltAdj_Main.frm") = True Then
        FLG_PALS_DISABLE.VoltageAdj = False
    Else
        If sub_SheetNameCheck(PinSheetname) = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [" & PinSheetname & "]Sheet" & vbCrLf
        End If
        If sub_SheetNameCheck(OutPutSheetname) = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [" & OutPutSheetname & "]Sheet" & vbCrLf
        End If
        If sub_SheetNameCheck(ReadSheetName) = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [" & ReadSheetName & "]Sheet" & vbCrLf
        End If
        If sub_SheetNameCheck(ReadSheetNameInfo) = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [" & ReadSheetNameInfo & "]Sheet" & vbCrLf
        End If
        If sub_ModuleCheck("PALS_VoltAdj_Mod.bas") = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [PALS_VoltAdj_Mod]Bas" & vbCrLf
        End If
        If sub_ModuleCheck("frm_PALS_VoltAdj_Main.frm") = False Then
            frm_PALS.cmd_VoltageRun.ControlTipText = frm_PALS.cmd_VoltageRun.ControlTipText & "NotFound [frm_PALS_VoltAdj_Main]Form" & vbCrLf
        End If
    End If
   '**********************************************************************************************************************

   '*******WAVE TOOL*******************************************************************************************************
    
    If sub_SheetNameCheck(WaveSetupSheetName) = True And sub_SheetNameCheck(WaveResultSheetName) = True And _
       sub_ModuleCheck("frm_PALS_WaveAdj_Main.frm") = True And sub_ModuleCheck("PALS_WaveAdj_mod_Common.bas") = True And sub_ModuleCheck("csPALS_WaveAdjust.cls") = True Then
        FLG_PALS_DISABLE.WaveAdj = False
    Else
        If sub_SheetNameCheck(WaveSetupSheetName) = False Then
            frm_PALS.cmd_WaveRun.ControlTipText = frm_PALS.cmd_WaveRun.ControlTipText & "NotFound [" & WaveSetupSheetName & "]Sheet" & vbCrLf
        End If
        If sub_SheetNameCheck(WaveResultSheetName) = False Then
            frm_PALS.cmd_WaveRun.ControlTipText = frm_PALS.cmd_WaveRun.ControlTipText & "NotFound [" & WaveResultSheetName & "]Sheet" & vbCrLf
        End If
        If sub_ModuleCheck("PALS_WaveAdj_mod_Common.bas") = False Then
            frm_PALS.cmd_WaveRun.ControlTipText = frm_PALS.cmd_WaveRun.ControlTipText & "NotFound [PALS_WaveAdj_mod_Common]Bas" & vbCrLf
        End If
        If sub_ModuleCheck("csPALS_WaveAdjust.bas") = False Then
            frm_PALS.cmd_WaveRun.ControlTipText = frm_PALS.cmd_WaveRun.ControlTipText & "NotFound [csPALS_WaveAdjust]Class" & vbCrLf
        End If
        If sub_ModuleCheck("frm_PALS_WaveAdj_Main.frm") = False Then
            frm_PALS.cmd_WaveRun.ControlTipText = frm_PALS.cmd_WaveRun.ControlTipText & "NotFound [frm_PALS_WaveAdj_Main]Form" & vbCrLf
        End If
    End If
   '**********************************************************************************************************************

   '*******BIAS TOOL*******************************************************************************************************
    
    If sub_ModuleCheck("frm_PALS_BiasAdj_Main.frm") = True And sub_ModuleCheck("PALS_BiasAdj_Mod.bas") = True Then
        FLG_PALS_DISABLE.BiasAdj = False
    Else
        If sub_ModuleCheck("frm_PALS_BiasAdj_Main.frm") = False Then
            frm_PALS.cmd_BiasRun.ControlTipText = frm_PALS.cmd_BiasRun.ControlTipText & "NotFound [frm_PALS_BiasAdj_Main]Form" & vbCrLf
        End If
        If sub_ModuleCheck("PALS_BiasAdj_Mod.bas") = False Then
            frm_PALS.cmd_BiasRun.ControlTipText = frm_PALS.cmd_BiasRun.ControlTipText & "NotFound [PALS_BiasAdj_Mod].Bas" & vbCrLf
        End If
    End If
   '**********************************************************************************************************************

   '*******TRACE TOOL*******************************************************************************************************
    
    If sub_ModuleCheck("frm_PALS_TraceAdj_Main.frm") = True And sub_ModuleCheck("PALS_TraceAdj_Mod.bas") = True And sub_ModuleCheck("PALS_TraceAcq_Mod.bas") = True Then
        FLG_PALS_DISABLE.TraceAdj = False
    Else
        If sub_ModuleCheck("frm_PALS_TraceAdj_Main.frm") = False Then
            frm_PALS.cmd_TraceRun.ControlTipText = frm_PALS.cmd_TraceRun.ControlTipText & "NotFound [frm_PALS_TraceAdj_Main]Form" & vbCrLf
        End If
        If sub_ModuleCheck("PALS_TraceAdj_Mod.bas") = False Then
            frm_PALS.cmd_TraceRun.ControlTipText = frm_PALS.cmd_TraceRun.ControlTipText & "NotFound [PALS_TraceAdj_Mod].Bas" & vbCrLf
        End If
        If sub_ModuleCheck("PALS_TraceAcq_Mod.bas") = False Then
            frm_PALS.cmd_TraceRun.ControlTipText = frm_PALS.cmd_TraceRun.ControlTipText & "NotFound [PALS_TraceAcq_Mod].Bas" & vbCrLf
        End If
    End If
   '**********************************************************************************************************************

Exit Sub

errPALSsub_CheckPALSParameter:
    Call sub_errPALS("Check PALS Parameter error at 'sub_CheckPALSParameter'", "0-5-01-0-14")

End Sub


'********************************************************************************************
' ���O: sub_FindPALSParameter
' ���e: strTarget�Ŏw�肳�ꂽ�����񂪁AstrTarget�Ŏw�肳�ꂽ�Z���͈͓��ɑ��݂��邩�`�F�b�N����B
'�@�@�@ ���݂����ꍇ�A�p�����[�^���ݔ��f�p�̃t���O��True�ɂ��A��ʒu���擾����B
' ����: strTarget     :����������
'       strWsName     :�������[�N�V�[�g��
'       strRange      :�����͈�
'       intPosi       :�����񂪑��݂����ԍ�
'       blnPALS_Check :���������񂪂�����1��ڂ�"PALS"�Ƃ��������񂪂��邩�m�F���邩�̃t���O
'                      ��True:�m�F����AFalse:�m�F�Ȃ�
' �ߒl: True          :��v�����񂠂�
'       False         :��v������Ȃ�
' ���l    �F ������intPosi��Byref�Œ�`���Ă���ׁA�����ŏ�������邱�Ƃɒ���
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_FindPALSParameter(ByRef strTarget As String, ByVal strWsName As String, ByVal strRange As String, _
                                        Optional ByRef intPosi As Integer = 0, Optional blnPALS_Check As Boolean = False) As Boolean

Const PALSNAME As String = "PALS"

On Error GoTo errPALSsub_FindPALSParameter

    Dim obj As Object

    Set obj = Worksheets(strWsName).Range(strRange).Find(strTarget, lookat:=xlWhole, MatchCase:=True)
    
    If obj Is Nothing Then
        sub_FindPALSParameter = False
    Else
        intPosi = Worksheets(strWsName).Range(strRange).Find(strTarget, lookat:=xlWhole, MatchCase:=True).Column
    
        If blnPALS_Check Then
            If Worksheets(strWsName).Cells(1, intPosi).Value = PALSNAME Then
                sub_FindPALSParameter = True
            Else
                MsgBox "Error : " & strTarget & "not found." & vbCrLf & "         Please check PALS parameter" & vbCrLf _
                        & "         at 'Test Instances'" & "ErrCode.0-5-02-2-15", vbExclamation, PALS_ERRORTITLE
            End If
        Else
            sub_FindPALSParameter = True
        End If
    
    End If

Exit Function

errPALSsub_FindPALSParameter:
    Call sub_errPALS("Find PALS Parameter error at 'sub_FindPALSParameter'", "0-5-02-0-16")

End Function



'********************************************************************************************
' ���O: sub_ReadTestInfo
' ���e: Flow Table��Test Instances�̃p�����[�^��csLoopTestInfo�N���X�Œ�`���ꂽLoopTestInfo�z��Ɋi�[�B
'       LoopTestInfo�͍��ڐ����̔z�񂪓��I�ɍ쐬�����B
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/10�@�V�K�쐬   K.Sumiyashiki
'            Rev1.1      2011/06/06  �����ǉ�   K.Sumiyashiki
'                                    ��Eee-JOB�̃e�X�g�C���X�^���X�ǂݍ��ݑΉ�
'                                      �K�i�ɕ����񂪓����Ă����ۂ̏�����ǉ�
'********************************************************************************************
Public Sub sub_ReadTestInfo()

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

    On Error GoTo errPALSsub_ReadTestInfo

    Dim FT_y As Long        '�t���[�e�[�u����Y�����Z���ʒu
    Dim index As Long       '�z��ԍ�
    Dim intOptIdenCnt As Integer
    Dim intWaitAdjCnt As Integer
    
    intOptIdenCnt = 0
    intWaitAdjCnt = 0
    'Flow Table�V�[�g�̑��݃`�F�b�N
    If Not sub_SheetNameCheck(FLOW_TABLE) Then
        Call sub_errPALS("Flow Table not found at 'sub_ReadTestInfo'", "0-5-03-8-17")
        Exit Sub
    End If
    
    '�t���[�e�[�u���̓ǂݍ���
    With Worksheets(FLOW_TABLE)
    
        '������
        FT_y = FT_START_Y
        index = 0
        
        'Opcode��"set-device"�ɂȂ�܂ŌJ��Ԃ�
        Do While (.Cells(FT_y, FT_OPCODE_X).Value <> FT_LASTROW_NAME)
        
            'TNum�������Ă���s�݂̂�����
            If Len(.Cells(FT_y, FT_TNUM_X).Value) Then
            
                '�z��̍Ē�`
                ReDim Preserve xTestInfo(index)
                
                'csLoopTestInfo�N���X�̃C���X�^���X����
                Set xTestInfo(index) = New csPALS_TestInfo
                
                '�e�p�����[�^���擾
                xTestInfo(index).tname = .Cells(FT_y, FT_TNAME_X).Value
                xTestInfo(index).Parameter = .Cells(FT_y, FT_PARAMETER_X).Value
                xTestInfo(index).Bin = .Cells(FT_y, FT_BIN_X).Value
                xTestInfo(index).OpCode = .Cells(FT_y, FT_OPCODE_X).Value
                
                '���x���̎擾�B�e���x���̍ŏ��̍��ڂɃ��x������ǉ�����
                If Len(.Cells(FT_y - 1, FT_LABEL_X).Value) > 0 Then
                    xTestInfo(index).Label = .Cells(FT_y - 1, FT_LABEL_X).Value
                ElseIf (Len(.Cells(FT_y - 1, FT_LABEL_X).Value) = 0) And (Len(.Cells(FT_y - 2, FT_LABEL_X).Value) > 0) _
                        And Len(.Cells(FT_y - 1, FT_TNUM_X).Value) = 0 Then
                    xTestInfo(index).Label = .Cells(FT_y - 2, FT_LABEL_X).Value
                End If
                
                '�e�X�g���ƃC���f�b�N�X��R�t����R���N�V�����Ƀf�[�^�ǉ�(�l:�C���f�b�N�X�A�L�[:���ږ�)
                TestnameInfoList.Add Item:=index, key:=.Cells(FT_y, FT_TNAME_X).Value
                
                '�C���f�b�N�X�̃C���N�������g
                index = index + 1
            End If
            '�t���[�e�[�u�� Y�����ʒu�̃C���N�������g
            FT_y = FT_y + 1
        Loop
    End With


    Dim TI_y As Long                        '�e�X�g�C���X�^���X��Y�����Z���ʒu
    Dim obj As Object

    index = 0

    'Test Instances�V�[�g�̑��݃`�F�b�N
    If Not sub_SheetNameCheck(TEST_INSTANCES) Then
        Call sub_errPALS("Test Instances not found at 'sub_ReadTestInfo'", "0-5-03-8-18")
    End If

    '�e�X�g�C���X�^���X�̓ǂݍ���
    With Worksheets(TEST_INSTANCES)
        
        'index�����ڐ��ɒB����܂ŌJ��Ԃ�
        Do While (index <= UBound(xTestInfo))
            
            '�e�X�g�C���X�^���X��B��ɋL�ڂ���Ă���p�����[�^���̌���
            '�t���[�e�[�u����Parameter�ƈ�v������̂�����Ώ������s��
            Set obj = .Range("B5:B10000").Find(xTestInfo(index).Parameter, lookat:=xlWhole, MatchCase:=True)
            
            '��v���鍀�ڂ������ꍇ
            If obj Is Nothing Then
                MsgBox "TestInstances parameter name not equal FlowTable parameter name!" & vbCrLf & "ErrCode.0-5-03-2-19", vbExclamation, PALS_ERRORTITLE
'                MsgBox "�t���[�e�[�u���ƃe�X�g�C���X�^���X�̍��ڂ���v���܂���B"
            Else
                TI_y = .Range("B5:B10000").Find(xTestInfo(index).Parameter, lookat:=xlWhole, MatchCase:=True).Row
                '�e�p�����[�^���擾
'>>>2011/06/06 K.SUMIYASHIKI UPDATE
'CIS�Ή�
                If IsNumeric(.Cells(TI_y, TI_HIGHLIMIT_X).Value) = False Then
                    xTestInfo(index).HighLimit = 1     '����K�i
                    xTestInfo(index).LowLimit = 1       '�����K�i
                    xTestInfo(index).Unit = ""               '�P��
                    xTestInfo(index).arg2 = 3               'Arg2
    
                Else
                    xTestInfo(index).HighLimit = .Cells(TI_y, TI_HIGHLIMIT_X).Value     '����K�i
                    xTestInfo(index).LowLimit = .Cells(TI_y, TI_LOWLIMIT_X).Value       '�����K�i
                    xTestInfo(index).Unit = .Cells(TI_y, TI_UNIT_X).Value               '�P��
                    xTestInfo(index).arg2 = .Cells(TI_y, TI_ARG2_X).Value               'Arg2
                End If
'<<<2011/06/06 K.SUMIYASHIKI UPDATE
            
                If m_LoopParameter.Flg_CapCategory1 Then
                    xTestInfo(index).CapCategory1 = .Cells(TI_y, m_LoopParameter.CapCategory1_X).Value
                End If
            
                If m_LoopParameter.Flg_CapCategory2 Then
                    xTestInfo(index).CapCategory2 = .Cells(TI_y, m_LoopParameter.CapCategory2_X).Value
                End If
                                
                If m_LoopParameter.Flg_LoopJudgeLimit Then
                    xTestInfo(index).LoopJudgeLimit = val(.Cells(TI_y, m_LoopParameter.LoopJudgeLimit_X).Value)
                    If Len(.Cells(TI_y, m_LoopParameter.LoopJudgeLimit_X).Value) = 0 Then
                        xTestInfo(index).LoopJudgeLimit = 0.1
                    End If
                End If

                If m_WaitParameter.Flg_WaitAdjFlg Then
                    xTestInfo(index).WaitAdjFlg = .Cells(TI_y, m_WaitParameter.WaitAdjFlg_X).Value
                End If

                If m_OptParameter.Flg_OptIdentifier Then
                    xTestInfo(index).OptIdentifier = .Cells(TI_y, m_OptParameter.OptIdentifier_X).Value
                    
                    If xTestInfo(index).OptIdentifier <> "" Then intOptIdenCnt = intOptIdenCnt + 1
                    Dim intTestindex As Integer
                    'Reduplicate�@Check
                    For intTestindex = 0 To index - 1
                        If xTestInfo(intTestindex).OptIdentifier = xTestInfo(index).OptIdentifier And xTestInfo(index).OptIdentifier <> "" Then
                            FLG_PALS_DISABLE.OptAdj = True
                            frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptIdentifier -" & "Duplicated at [" & TEST_INSTANCES & "]-" & xTestInfo(intTestindex).OptIdentifier & vbCrLf
                        End If
                    Next
                End If

                If m_OptParameter.Flg_OptTarget Then
                    xTestInfo(index).OptTarget = val(.Cells(TI_y, m_OptParameter.OptTarget_X).Value)
                    '>>>2011/4/22 M.IMAMURA ADD
                    If xTestInfo(index).OptIdentifier <> "" And xTestInfo(index).arg2 = 3 And (xTestInfo(index).OptTarget < xTestInfo(index).LowLimit Or xTestInfo(index).OptTarget > xTestInfo(index).HighLimit) Then
                        FLG_PALS_DISABLE.OptAdj = True
                        frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptTarget -" & "Out of TestLimt at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                    End If
                    If xTestInfo(index).OptIdentifier <> "" And xTestInfo(index).arg2 = 1 And (xTestInfo(index).OptTarget < xTestInfo(index).LowLimit) Then
                        FLG_PALS_DISABLE.OptAdj = True
                        frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptTarget -" & "Out of TestLimt at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                    End If
                    If xTestInfo(index).OptIdentifier <> "" And xTestInfo(index).arg2 = 2 And (xTestInfo(index).OptTarget > xTestInfo(index).HighLimit) Then
                        FLG_PALS_DISABLE.OptAdj = True
                        frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptTarget -" & "Out of TestLimt at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                    End If
                    '>>>2011/4/22 M.IMAMURA ADD
                End If

                If m_OptParameter.Flg_OptJudgeLimit Then
                    xTestInfo(index).OptJudgeLimit = val(.Cells(TI_y, m_OptParameter.OptJudgeLimit_X).Value)
                    If xTestInfo(index).OptJudgeLimit = 0 Then xTestInfo(index).OptJudgeLimit = 0.01
                End If
            

                If ((m_OptParameter.Flg_OptTarget = True) And (m_OptParameter.Flg_OptIdentifier = True) And (m_OptParameter.Flg_OptJudgeLimit = True)) Then
                    If .Cells(TI_y, m_OptParameter.OptIdentifier_X).Value <> "" Or val(.Cells(TI_y, m_OptParameter.OptJudgeLimit_X).Value) <> 0 And val(.Cells(TI_y, m_OptParameter.OptTarget_X).Value) <> 0 Then
                        '>>>2011/4/22 M.IMAMURA UPDATE
                        If .Cells(TI_y, m_OptParameter.OptIdentifier_X).Value = "" Then
                            FLG_PALS_DISABLE.OptAdj = True
                            frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptIdentifier -" & " Not Found Parameter at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                        End If
                        If .Cells(TI_y, m_OptParameter.OptJudgeLimit_X).Value <> "" And (val(.Cells(TI_y, m_OptParameter.OptJudgeLimit_X).Value) <= 0 Or val(.Cells(TI_y, m_OptParameter.OptJudgeLimit_X).Value) >= 1) Then
                            FLG_PALS_DISABLE.OptAdj = True
                            frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptJudgeLimit -" & " Illegal Parameter at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                        End If
                        If val(.Cells(TI_y, m_OptParameter.OptTarget_X).Value) <= 0 Then
                            FLG_PALS_DISABLE.OptAdj = True
                            frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptTarget -" & " Illegal Parameter at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                        End If
                        '<<<2011/4/22 M.IMAMURA UPDATE
                    End If

                End If

                If ((m_WaitParameter.Flg_CapCategory1 = True) And (m_WaitParameter.Flg_CapCategory2 = True) And (m_WaitParameter.Flg_WaitAdjFlg = True)) Then
                    If xTestInfo(index).CapCategory1 <> "" Or .Cells(TI_y, m_WaitParameter.WaitAdjFlg_X).Value <> "" Then
                        intWaitAdjCnt = intWaitAdjCnt + 1
                        If .Cells(TI_y, m_WaitParameter.CapCategory1_X).Value = "" Then
                            FLG_PALS_DISABLE.WaitAdj = True
                            frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- CapCategory1 -" & "Not Found Parameter at [" & TEST_INSTANCES & "]-" & xTestInfo(index).Parameter & vbCrLf
                            intWaitAdjCnt = intWaitAdjCnt - 1
                        End If
                        If .Cells(TI_y, m_WaitParameter.WaitAdjFlg_X).Value <> "RUN" Then
'                            FLG_PALS_DISABLE.WaitAdj = True
'                            frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- WaitAdjFlg -" & "Not Found Parameter at [" & TEST_INSTANCES & "]-" & xTestInfo(Index).Parameter & vbCrLf
                            intWaitAdjCnt = intWaitAdjCnt - 1
                        End If

                    End If

                End If


            End If
            index = index + 1
        Loop
    End With

    If intOptIdenCnt = 0 Then
        FLG_PALS_DISABLE.OptAdj = True
        frm_PALS.cmd_OptRun.ControlTipText = frm_PALS.cmd_OptRun.ControlTipText & "- OptIdentifier -" & "Not Found Parameter" & vbCrLf
    End If

    If intWaitAdjCnt <= 0 Then
        FLG_PALS_DISABLE.WaitAdj = True
        frm_PALS.cmd_WaitRun.ControlTipText = frm_PALS.cmd_WaitRun.ControlTipText & "- WaitAdjFlg -" & "Not Found Parameter" & vbCrLf
    End If

Exit Sub

errPALSsub_ReadTestInfo:
    Call sub_errPALS("Read IG-XL sheet error at 'sub_ReadTestInfo'", "0-5-03-0-20")

End Sub


'********************************************************************************************
' ���O: TestCount
' ���e: ���ڐ���Ԃ��֐�
' ����: �Ȃ�
' �ߒl: ���ڐ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/20�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Function TestCount() As Integer

On Error GoTo errPALSTestCount

    TestCount = UBound(xTestInfo)

Exit Function

errPALSTestCount:
    Call sub_errPALS("Get test count error at 'csPALS_Common -> TestCount'", "0-5-04-0-21")

End Function

