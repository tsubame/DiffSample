Attribute VB_Name = "XLibToptFrameWork_T"
'�T�v:
'   ToptScenario�p�J�X�^���e���v���[�g
'   �iDCTestScenario�p�J�X�^���e���v���[�g�𗬗p���Ă���j
'
'�ړI:
'   Eee-JOB ToptScenario�e�X�g�C���X�^���X�t���[�����[�N�p�ɐ�p�e���v���[�g���쐬
'   IG-XL Ver.3.40.17��Empty_T���x�[�X�Ƀp�����[�^��ǉ��폜���J�X�^�}�C�Y���s����
'   �C���X�^���X�G�f�B�^�͐�p��ToptScenario_IE���g�p����
'
'   ���i�K�ł̓e���v���[�g�̏������e�͈ȉ��̒ʂ�
'
'   �@PreBody�F TheTOPT�I�u�W�F�N�g��.SetScenario���\�b�h�Ăяo��
'   �ABody�F    StartOfBody�i�����j�̎��s
'               TheTOPT�I�u�W�F�N�g��.Execute���\�b�h�Ăяo��
'               EndOfBody�i�����j�̎��s
'   �BPostBody: �����Ȃ�
'
'   Revision History:
'       Data        Description
'       2010/04/28  �J�X�^���e���v���[�g�ł�ToptScenario����������
'       2010/05/12  Parameter��Check�@�\���폜����
'       2010/05/20  ActionLogger�̃Z�[�u�@�\���폜�����iInterPoseFunction�Ŏ��s�j
'       2010/05/28  ValidationError���̕s����C������
'       2010/06/11  �v���O�����R�[�h�𐮗�����
'
'�쐬��:
'   0145184346
'   �@�@2013/08/21  H.Arikawa�@Func:SetupParameters�@TestInstances�V�[�g�ւ�Comment�o�͂𖳌���
'   �@�@2013/10/16  H.Arikawa�@Eee-Job V2.14�̓��e�̓��ꍞ��

Option Explicit

' (c) Teradyne, Inc, 1997, 1998, 1999
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
' Revision History:
' Date        Description
' 11/01/07    Ganesh Fix for the defect tersw00097677 - Can't validate Acquire_T template.
' 12/27/06    Chandra Shekhar S Since TI test programs are failing in 3.40.17, fix for tersw00090947 is reverted
' 07/28/06    Chandra Shekhar S Fix for tersw0090947- State Init/Start setting is
'                               changed to StateOff forcedly on  PostBody of Template
' 04/03/06     Ganesh Pandiyan K Fix tersw00072341 - Added an Edit box to enter comments in the Template GUI.
' 09/27/99    Release 3.30 Development
'
'
Dim Arg_DcCategory As String
Dim Arg_DcSelector As String
Dim Arg_AcCategory As String
Dim Arg_AcSelector As String
Dim Arg_Timing As String
Dim Arg_Edgeset As String
Dim Arg_Levels As String

'------------------------------------------------------
'Eee-JOB ToptScenario�e���v���[�g�p�ϐ��ǉ�
Dim Arg_AcquireInstanceF As String
Dim Arg_UserMacroF As String
Dim Arg_StartOfBodyF As String
Dim Arg_EndOfBodyF As String
Dim Arg_StartOfBodyFInput As String
Dim Arg_EndOfBodyFInput As String

Private Const ARGNUM_ACQUIREINSTANCE = 0
Private Const ARGNUM_USERMACRO = 1
Private Const ARGNUM_STARTOFBODYF = 3
Private Const ARGNUM_ENDOFBODYF = 4
Private Const ARGNUM_STARTOFBODYFINPUT = 5
Private Const ARGNUM_ENDOFBODYFINPUT = 6
Private Const COMMENTCOLUMN_FOR_EEEJOB = 80
Private Const ARGNUM_MAXARG = ARGNUM_ENDOFBODYFINPUT

Private Const EEEJOB_FORM_CAPTION = "Topt Frame Work -- Instance Editor"
Private Const EEEJOB_TEMPLATE_NAME = "XLibTopFrameWork_T"
Private Const EEEJOB_DEFAULT_COMMENT = "Topt Frame Work For Eee-JOB"
Private Const EEEJOB_INIT_SCENARIO As String = "XLibToptTemplate.SetScenario"
Private Const EEEJOB_EXEC_SCENARIO As String = "XLibToptTemplate.Execute"
'------------------------------------------------------

Function TestTemplate() As Integer
'���e:
'   Eee-JOB ToptScenario�e���v���[�g�p���C���֐�
'
'�߂�l�F
'   �e���v���[�g���s����
'   �����F0
'   ���s�F1
'
'���ӎ���:
'   �e���_�C���W���e���v���[�g���C���֐���ύX
'   PreBody�ABody�APostBody�̊e���s���ʂɊ֏���݂���
'
    TestTemplate = TL_SUCCESS
    If PreBody() <> TL_SUCCESS Then GoTo ErrHandler
    If Body() <> TL_SUCCESS Then GoTo ErrHandler
    If PostBody() <> TL_SUCCESS Then GoTo ErrHandler
    Exit Function
ErrHandler:
    TestTemplate = TL_ERROR
End Function

Function PreBody() As Integer
    If TheExec.Flow.IsRunning = False Then Exit Function
    On Error GoTo ErrHandler
    Call GetTemplateParameters
    '------------------------------------------------------
    'Eee-JOB ToptScenario�e���v���[�g�p�����ύX
    'JOB��Topt�e�X�g�V�i���I�G���W��������
    '�������藈�Ȃ������i�K�ł̓e���_�C��API�ŃG���W���������̃��b�p�[�֐����Ăяo��
    '�������Ŏ��s�����ꍇ�G���[��Ԃ�
    PreBody = TheExec.Flow.CallFuncWithArgs(EEEJOB_INIT_SCENARIO, TL_C_EMPTYSTR)
    '------------------------------------------------------
    Exit Function
ErrHandler:
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    Call TheExec.ErrorReport
    PreBody = TL_ERROR
End Function

Function Body() As Integer
    If TheExec.Flow.IsRunning = False Then Exit Function
    On Error GoTo ErrHandler
    If Arg_StartOfBodyF <> TL_C_EMPTYSTR Then _
        Call TheExec.Flow.CallFuncWithArgs(Arg_StartOfBodyF, Arg_StartOfBodyFInput)
    '------------------------------------------------------
    'Eee-JOB ToptScenario�e���v���[�g�p�����ύX
    'JOB��Topt�e�X�g�V�i���I�G���W�����s
    '�������藈�Ȃ������i�K�ł̓e���_�C��API�ŃG���W�����s�̃��b�p�[�֐����Ăяo��
    '�G���W�����s�Ŏ��s�����ꍇ�G���[��Ԃ�
    Body = TheExec.Flow.CallFuncWithArgs(EEEJOB_EXEC_SCENARIO, TL_C_EMPTYSTR)
    '------------------------------------------------------
    If Arg_EndOfBodyF <> TL_C_EMPTYSTR Then _
        Call TheExec.Flow.CallFuncWithArgs(Arg_EndOfBodyF, Arg_EndOfBodyFInput)
    Exit Function
ErrHandler:
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    Call TheExec.ErrorReport
    Body = TL_ERROR
End Function

Function PostBody() As Integer
    If TheExec.Flow.IsRunning = False Then Exit Function
    PostBody = TL_SUCCESS
End Function

Sub GetTemplateParameters()
    Dim ArgStr() As String
    Call tl_tm_GetInstanceValues(ARGNUM_MAXARG, ArgStr)
    Arg_DcCategory = ArgStr(TL_C_DCCATCOLNUM)
    Arg_DcSelector = ArgStr(TL_C_DCSELCOLNUM)
    Arg_AcCategory = ArgStr(TL_C_ACCATCOLNUM)
    Arg_AcSelector = ArgStr(TL_C_ACSELCOLNUM)
    Arg_Timing = ArgStr(TL_C_TIMESETCOLNUM)
    Arg_Edgeset = ArgStr(TL_C_EDGESETCOLNUM)
    Arg_Levels = ArgStr(TL_C_LEVELSCOLNUM)

    '------------------------------------------------------
    'Eee-JOB ToptScenario�e���v���[�g�p�����ύX
    '�C���X�^���X�V�[�g����̃p�����[�^�擾
    Arg_AcquireInstanceF = ArgStr(ARGNUM_ACQUIREINSTANCE)
    Arg_UserMacroF = ArgStr(ARGNUM_USERMACRO)
    Arg_StartOfBodyF = ArgStr(ARGNUM_STARTOFBODYF)
    Arg_EndOfBodyF = ArgStr(ARGNUM_ENDOFBODYF)
    Arg_StartOfBodyFInput = ArgStr(ARGNUM_STARTOFBODYFINPUT)
    Arg_EndOfBodyFInput = ArgStr(ARGNUM_ENDOFBODYFINPUT)
    '------------------------------------------------------
End Sub

Function DatalogType() As Integer
    DatalogType = logFunctional
End Function

' End of Execution Section

Public Function RunIE(Optional FocusArg As Integer) As Boolean
    tl_tm_FocusArg = FocusArg
    Call tl_fs_ResetIECtrl(tl_tm_InstanceEditor)

    '------------------------------------------------------
    'Eee-JOB ToptScenario�C���X�^���X�G�f�B�^�p�����ύX
    '��p�C���X�^���X�G�f�B�^�̂��߂̃v���p�e�B�ݒ�
    With tl_tm_InstanceEditor
        .Name = EEEJOB_TEMPLATE_NAME
        .InterposePage = True
        .UserReqPage1 = True
        .Caption = EEEJOB_FORM_CAPTION
        .HelpValue = 0
    End With
    '------------------------------------------------------

    '------------------------------------------------------
    'Eee-JOB ToptScenario�C���X�^���X�G�f�B�^�p�����ύX
    '��p�C���X�^���X�G�f�B�^�̋N��
    Call ToptScenario_IE.Show
    '------------------------------------------------------

    'the return value will be true if the 'Apply' button was not enabled and if the workbook was valid when the form initialized
    RunIE = (Not (tl_tm_FormCtrl.ButtonEnabled)) And tl_tm_BookIsValid
End Function

Sub AssignTemplateValues()
    Dim ArgStr() As String
    Call tl_tm_GetInstanceValues(ARGNUM_MAXARG, ArgStr)
    For Each tl_tm_ParThisPar In AllPars
        With tl_tm_ParThisPar
            'Fix for the defect tersw00072341 - Added a edit box to enter the comments
            If (.Argnum <= UBound(ArgStr)) Then
                .ParameterValue = ArgStr(.Argnum)
            End If
        End With
    Next

    '------------------------------------------------------
    'Eee-JOB ToptScenario�C���X�^���X�G�f�B�^�p�����ύX
    'IG-XL Ver3.40.10����ł̓R�����g�̃p�����[�^���l������Ă��Ȃ����߁A
    'Eee-JOB ToptScenario�e���v���[�g��p�֐��ɌĂяo����ύX
    'Call Eee_ManageDefault(AllPars, ARGNUM_MAXARG)             '=== Add Eee-Job V2.14 ===
    '------------------------------------------------------

End Sub

Sub ApplyDefaults()
    Call SetupParameters
    For Each tl_tm_ParThisPar In AllPars
        With tl_tm_ParThisPar
            Call tl_tm_PutDefaultIfNeeded(.Argnum, .defaultvalue)
        End With
    Next
    Call tl_tm_CleanUp
End Sub

Function GetArgNames() As String
    Dim CallSetup As Boolean
    CallSetup = False
    If AllPars.Count = 0 Then
        Call SetupParameters    'acquire the Argument information, if needed
        CallSetup = True
    End If

    '------------------------------------------------------
    'Eee-JOB ToptScenario�C���X�^���X�G�f�B�^�p�����ύX
    'IG-XL Ver3.40.10����ł̓R�����g�̃p�����[�^���l������Ă��Ȃ����߁A
    'Eee-JOB ToptScenario�e���v���[�g��p�֐��ɌĂяo����ύX
    GetArgNames = Eee_ListArgNames(ARGNUM_MAXARG)
    '------------------------------------------------------

    If CallSetup = True Then Call tl_tm_CleanUp
End Function

Sub SetupParameters()
    Call tl_tm_SetupCatSelValidation
    Call tl_tm_SetupTimLevValidation
    Call tl_tm_SetupOverlayValidation

    '------------------------------------------------------
    'Eee-JOB ToptScenario�C���X�^���X�G�f�B�^�p�����ύX
    '�C���X�^���X�G�f�B�^��ɂ�����p�����[�^���A
    '�C���X�^���X�V�[�g��̃p�����[�^�J�����ʒu���̐ݒ�
    With tl_tm_ParUserReq1
        .AllParsAdd
        .Argnum = ARGNUM_ACQUIREINSTANCE
        .ParameterName = "Acquire Instance Name"
        .tl_tm_ParSetParam
        .TestFunctionName = False ' Parameter Check Simular TestInstanceName
''        .TestFunctionName = True
        .TestNotBlank = True ' Validation Error Check
        .SetEnabler tl_tm_ParUserReq2, TL_C_BLANK
        .TestingEnabled = True
    End With
    With tl_tm_ParUserReq2
        .AllParsAdd
        .Argnum = ARGNUM_USERMACRO
        .ParameterName = "User Macro Name"
        .tl_tm_ParSetParam
        .TestFunctionName = False ' Parameter Check Simular TestInstanceName
''        .TestFunctionName = True
        .TestNotBlank = True ' Validation Error Check
        .SetEnabler tl_tm_ParUserReq1, TL_C_BLANK
        .TestingEnabled = True
    End With
    With tl_tm_ParStartOfBodyF
        .AllParsAdd
        .Argnum = ARGNUM_STARTOFBODYF
        .ParameterName = TL_C_StartOfBodyFStr
        .tl_tm_ParSetParam
        .TestFunctionName = True
    End With
    With tl_tm_ParStartOfBodyFInput
        .AllParsAdd
        .Argnum = ARGNUM_STARTOFBODYFINPUT
        .ParameterName = TL_C_StartOfBodyFStr & TL_C_IpfInputStr
        .tl_tm_ParSetParam
     End With
    With tl_tm_ParEndOfBodyF
        .AllParsAdd
        .Argnum = ARGNUM_ENDOFBODYF
        .ParameterName = TL_C_EndOfBodyFStr
        .tl_tm_ParSetParam
        .TestFunctionName = True
    End With
    With tl_tm_ParEndOfBodyFInput
        .AllParsAdd
        .Argnum = ARGNUM_ENDOFBODYFINPUT
        .ParameterName = TL_C_EndOfBodyFStr & TL_C_IpfInputStr
        .tl_tm_ParSetParam
    End With
    '------------------------------------------------------

    '------------------------------------------------------
    '2013 ��� IS�e�X�g�Z�p4�ہ@�������Č��@Validation���ԒZ�k
    'Eee-JOB ToptScenario�C���X�^���X�G�f�B�^�p�����ύX
    'IG-XL Ver3.40.10����ł�tl_tm_ParFuncCommentsTextBox���T�|�[�g���Ă��Ȃ����߁A
    '����Ƀ��[�U�[�J�X�^���pTemplateArg�I�u�W�F�N�g�ϐ�[tl_tm_ParUserOpt10]��q��
    With tl_tm_ParUserOpt10
        .AllParsAdd
        .Argnum = COMMENTCOLUMN_FOR_EEEJOB
        .tl_tm_ParSetParam
        .TestIsLegalChoice = False
        .defaultvalue = EEEJOB_DEFAULT_COMMENT
    End With
'''    '------------------------------------------------------

End Sub

Function ValidateParameters(Optional VDCint As Integer) As Integer
    'This function is used, at validation time, to determine whether the data
    '   to be executed is proper, valid, and copacetic.  It can be called by
    '   an Instance Editor, or by the Job Validation routines.
    Dim TestResult As Integer
    Dim temp As String
    '   This has modes to run in.  If a mode of '0' is specified for
    '   input, it is assumed that the mode is TL_C_VALDATAMODEJOBVAL.
    '   The modes that .ValidateParameters can operate in are:
    '   TL_C_VALDATAMODEJOBVAL  -   Job Validation mode; report errors to sheet.
    '   TL_C_VALDATAMODENORMAL  -   Instance Editor mode; Fix the current parameter being evaluated.
    '   TL_C_VALDATAMODENOSTOP  -   Instance Editor mode; Do not stop to fix any parameters.
    '   It can return with different modes, such as:
    '   TL_C_VALDATAMODENOFIX   -   Instance Editor mode; Error found, that specific one was not fixed.
    '   TL_C_VALDATAMODEFIXNONE -   Instance Editor mode; Error(s) found, none were fixed.

    'Success is first assumed; if a problem is noted, ValidateParameters will be
    '   set to failure by this routine.
    ValidateParameters = TL_SUCCESS

    If VDCint = 0 Then VDCint = TL_C_VALDATAMODEJOBVAL
    If (VDCint <> TL_C_VALDATAMODENORMAL) And (VDCint <> TL_C_VALDATAMODENOSTOP) _
        And (VDCint <> TL_C_VALDATAMODEJOBVAL) Then
        'denote an error
        temp = TheExec.DataManager.InstanceName
        Call TheExec.ErrorLogMessage("ValidateParameters: Improper mode, instance: " & temp)
        Call TheExec.ErrorReport
        ValidateParameters = TL_ERROR
        Exit Function
    End If

    If VDCint = TL_C_VALDATAMODEJOBVAL Then
        With JobData
            'Get list of pins and pin-groups from datatools.
            Call tl_fs_TemplateJobDataPinlistStrings(JobData, VDCint)

            'Get lists of Categories, Selectors, Timesets, Edgesets, and Levels
            Call tl_fs_TemplateCatSelStrings(.AvailDcCat, .AvailDcSel, _
                .AvailAcCat, .AvailAcSel, .AvailTimeSetAll, .AvailTimeSetExtended, _
                .AvailEdgeSet, .AvailLevels)
            'Get list of Overlay
            Call tl_fs_TemplateOverlayString(.AvailOverlay)
        End With

        'Define the Parameter types and tests to be performed
        Call SetupParameters

        'Now, acquire the values of the parameters for this Template Instance
        '   from the DataManager and assign them to the TemplateArg structures.
        Call AssignTemplateValues
    End If

    ValidateParameters = TL_SUCCESS

    ' Choose tests to perform
    Call tl_tm_ChooseTests(AllPars, VDCint)

    '#####  �p�����[�^�`�F�b�N�̎��{�L���𐧌䂷��  #####
    tl_tm_ParUserReq1.TestingEnabled = True
    tl_tm_ParUserReq2.TestingEnabled = True
    If tl_tm_ParUserReq1.GetParStr = TL_C_EMPTYSTR And tl_tm_ParUserReq2.GetParStr <> TL_C_EMPTYSTR Then
        tl_tm_ParUserReq1.TestingEnabled = False
    End If
    If tl_tm_ParUserReq1.GetParStr <> TL_C_EMPTYSTR And tl_tm_ParUserReq2.GetParStr = TL_C_EMPTYSTR Then
        tl_tm_ParUserReq2.TestingEnabled = False
    End If

    ' Now run the tests on each Argument
    Call tl_tm_RunTests(AllPars, VDCint, TestResult)
    If TestResult <> TL_SUCCESS Then ValidateParameters = TL_ERROR
    If (TestResult <> TL_SUCCESS) And (VDCint = TL_C_VALDATAMODENORMAL) Then Exit Function

'    Warning: Be aware that interpose functions are not validated

    If VDCint = TL_C_VALDATAMODEJOBVAL Then Call tl_tm_CleanUp
    
'     ValidateParameters = TL_SUCCESS

End Function

Function ValidateDriverParameters() As Integer
    Dim RetVal As Long
    ValidateDriverParameters = TL_SUCCESS
    Call SetupParameters
    'Now, acquire the values of the parameters for this Template Instance
    '   from the DataManager and assign them to the TemplateArg structures.
    Call AssignTemplateValues
    Call tl_tm_CleanUp
End Function

Private Function Eee_ListArgNames(MaxArg As Long) As String
'���e:
'   �e���v���[�g�œo�^���ꂽ�S�Ẵp�����[�^�����J���}��؂�̃��X�g�����s��
'
'�p�����[�^:
'    [MaxArg]    In   �p�����[�^�����͂����J�����̍ő�ʒu
'
'�߂�l:
'   �J���}��؂�̃p�����[�^�����X�g
'
'���ӎ���:
'   IG-XL Ver3.40.10����[ValSupport�v���W���[���́utl_tm_ListArgNames�v�֐��ł�
'   �R�����g�̃p�����[�^���l������Ă��Ȃ����߁AVer3.40.17����C���|�[�g���֐�����ύX
'   Eee-JOB ToptScenario�e���v���[�g��p�֐��Ƃ���
'
    Dim NameArr() As String
    Dim intX As Integer
    ReDim NameArr(MaxArg - TL_C_DCCATCOLNUM)
    For Each tl_tm_ParThisPar In AllPars
        If ((tl_tm_ParThisPar.Argnum - TL_C_DCCATCOLNUM) <= (MaxArg - TL_C_DCCATCOLNUM)) Then
                NameArr(tl_tm_ParThisPar.Argnum - TL_C_DCCATCOLNUM) = tl_tm_ParThisPar.ParameterName
        End If
    Next
    Eee_ListArgNames = TL_C_EMPTYSTR
    For intX = 0 To MaxArg
        Eee_ListArgNames = Eee_ListArgNames & NameArr(intX - TL_C_DCCATCOLNUM) & TL_C_DELIMITERSTD
    Next intX
    Eee_ListArgNames = Left$(Eee_ListArgNames, Len(Eee_ListArgNames) - 1)
End Function

Private Sub Eee_ManageDefault(AllArgs As Collection, MaxArgCnt As Long)
'���e:
'   �e���v���[�g�œo�^���ꂽ�S�Ẵp�����[�^�����C���X�^���X�V�[�g��̃J�������x���ɐݒ肷��
'
'�p�����[�^:
'    [AllArgs]    In   �p�����[�^�I�u�W�F�N�g�̃R���N�V����
'    [MaxArgCnt]  In   �p�����[�^�����͂����J�����̍ő�ʒu
'
'�߂�l:
'
'���ӎ���:
'   IG-XL Ver3.40.10����[ValSupport�v���W���[���́utl_tm_ManageDefault�v�֐��ł�
'   �R�����g�̃p�����[�^���l������Ă��Ȃ����߁AVer3.40.17����C���|�[�g���֐�����ύX
'   Eee-JOB ToptScenario�e���v���[�g��p�֐��Ƃ���
'
    Dim DefSet As Boolean
    Dim AllArgArr() As String
    Dim MemNum As Long
    ReDim AllArgArr(MaxArgCnt)

    DefSet = False
    For Each tl_tm_ParThisPar In AllArgs
        With tl_tm_ParThisPar
            If (.ParameterValue = TL_C_EMPTYSTR) And (.defaultvalue <> TL_C_EMPTYSTR) Then
                .ParameterValue = .defaultvalue
                Call tl_fs_PutData(.Argnum, .defaultvalue)
                DefSet = True
            End If
            If .Argnum >= 0 Then
                If (.Argnum < (MaxArgCnt - TL_C_DCCATCOLNUM)) Then
                    AllArgArr(.Argnum) = .ParameterValue
                End If
            End If
        End With
    Next
    If DefSet = True Then
        Dim temp As String
        temp = TheExec.DataManager.InstanceName
        MemNum = TheExec.DataManager.memberindex
        DefSet = TheExec.DataManager.ReloadInstance(temp, AllArgArr, MemNum)
    End If
End Sub
