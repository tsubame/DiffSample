VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_PALS_LoopAdj_Main 
   Caption         =   "PALS - Auto Loop Parameter Adjust"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   OleObjectBlob   =   "frm_PALS_LoopAdj_Main.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frm_PALS_LoopAdj_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

'##########################################################
'�t�H�[���́~�{�^������������
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000

' �E�B���h�E�Ɋւ������Ԃ�
Private Declare Function GetWindowLong Lib "USER32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' �E�B���h�E�̑�����ύX
Private Declare Function SetWindowLong Lib "USER32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' Active�ȃE�B���h�E�̃n���h�����擾
Private Declare Function GetActiveWindow Lib "USER32.dll" () As Long
' ���j���[�o�[���ĕ`��
Private Declare Function DrawMenuBar Lib "USER32.dll" (ByVal hWnd As Long) As Long

Private Sub Btn_ContinueOnFail_Click()
    Btn_StopOnFail.Value = False
    Btn_ContinueOnFail.Value = True
End Sub

Private Sub Btn_StopOnFail_Click()
    Btn_StopOnFail.Value = True
    Btn_ContinueOnFail.Value = False
End Sub

Private Sub op_AutoAdjust_Click()
    txt_maxwait.enabled = True
    txt_maxtrial_num.enabled = True
End Sub

Private Sub op_NotAdjust_Click()
    txt_maxwait.enabled = False
    txt_maxtrial_num.enabled = False
End Sub

Private Sub txt_maxtrial_num_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txt_maxtrial_num_Change()
    txt_loop_num.Text = val(txt_loop_num.Text)
    If txt_loop_num.Text = "0" Then txt_loop_num.Text = "1"
End Sub

Private Sub txt_maxwait_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txt_maxwait_Change()
    txt_loop_num.Text = val(txt_loop_num.Text)
    If txt_loop_num.Text = "0" Then txt_loop_num.Text = "1"
End Sub

Private Sub UserForm_Activate()
    Dim hWnd As Long
    Dim Wnd_STYLE As Long

    hWnd = GetActiveWindow()
    Wnd_STYLE = GetWindowLong(hWnd, GWL_STYLE)
    Wnd_STYLE = Wnd_STYLE And (Not WS_SYSMENU)
    SetWindowLong hWnd, GWL_STYLE, Wnd_STYLE
    DrawMenuBar hWnd
    Me.Caption = LOOPTOOLNAME & " Ver:" & LOOPTOOLVER
End Sub


'********************************************************************************************
' ���O : UserForm_Initialize
' ���e : ���[�U�[�t�H�[���o�͎��̏������֐�
' ���� : �Ȃ�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
' �X�V�����F Rev1.1      2012/03/09�@���ʒ�������̘A�g�Ή�   M.Imamura
'********************************************************************************************
Private Sub UserForm_Initialize()
    g_blnLoopStop = False
    Call sub_PalsFileCheck(PALS_PARAMFOLDERNAME_LOOP)
    
    '>>>2011/8/31 M.IMAMURA Add. ForCIS
    If gblnForCis = True Then
        op_AutoAdjust.enabled = False
        op_NotAdjust.Value = True
        Btn_ContinueOnFail.Value = True
    End If
    '>>>2011/8/31 M.IMAMURA Add. ForCIS

    '>>>2012/3/9 M.IMAMURA Add. For Connection From OptAdj
    If FLG_PALS_RUN.OptAdj = True Then
        With frm_PALS_LoopAdj_Main
            .txt_loop_num.Value = frm_PALS_OptAdj_Main.txt_loop_num.Text
            
            If frm_PALS_OptAdj_Main.op_AutoAdjust.Value = False Then
                .op_NotAdjust.Value = True
                .op_AutoAdjust.Value = False
            Else
                .op_NotAdjust.Value = False
                .op_AutoAdjust.Value = True
            End If
            
            .Btn_ContinueOnFail.Value = frm_PALS_OptAdj_Main.Btn_ContinueOnFail.Value
            .txt_maxwait = frm_PALS_OptAdj_Main.txt_maxwait.Text
            .txt_maxtrial_num = frm_PALS_OptAdj_Main.txt_maxtrial_num.Text
            Call .cmd_start_Click
        End With
    End If
    '<<<2012/3/9 M.IMAMURA Add. For Connection From OptAdj

End Sub



'********************************************************************************************
' ���O : cmd_Start_Click
' ���e : LOOP�����J�n�{�^���N���b�N�����ۂ̓���
'        ���菀���ˑ���ˌX�����f�˒��[�o�̗͂���𐧌䂵�Ă���
' ���� : �Ȃ�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Sub cmd_start_Click()
    
    If Not sub_CheckTestInstancesParams Then
        Exit Sub
    End If
    
    Dim Flg_AnalyzeDebug As Boolean

    If sub_CheckResultFormat = False Then
        Exit Sub
    End If
    
    If frm_PALS_LoopAdj_Main.Chk_DebugMode = True Then
        Flg_AnalyzeDebug = True
        frm_PALS_LoopAdj_Main.chk_IGXL_Check.Value = True
    End If
                
On Error GoTo errPALScmd_start_ClickLoop
        
    Dim intLoopTrialCnt As Integer      '���[�v���s�񐔂������ϐ�
    Dim MeasureDatalogInfo As DatalogInfo
        
    '�r��STOP���s�����ۂ̏���
    If cmd_start.Caption = "Stop" Then
        If MsgBox("Pushed [Stop] Button" & vbCrLf & "Do You Want Stop?", vbYesNo, LOOPTOOLNAME) = vbYes Then
            g_blnLoopStop = True
            Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Canceled Loop...  ")
            cmd_start.enabled = False
        End If
        Exit Sub
    End If
        
    '������
    intLoopTrialCnt = 0
    
    'TestCondition�ɐݒ肵�Ă���e�J�e�S����Wait���A�t�H�[���Ŏw�肵���ő�Wait�ȏ�ɂȂ��Ă��Ȃ����`�F�b�N����
    If Not sub_CheckTestConditionWaitData(val(frm_PALS_LoopAdj_Main.txt_maxwait)) Then
        Exit Sub
    End If
    
    'LOOP�������J�n���邩�̊m�F
    '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    If g_RunAutoFlg_PALS = False And frm_PALS_OptAdj_Main.cb_ConnectLoop.Value = False Then
        If MsgBox("Pushed [Start] Button, Ready?", vbOKCancel, LOOPTOOLNAME) = vbCancel Then
            Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Canceled...  ")
            Exit Sub
        End If
    End If
    '<<<2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    
    '�t�H�[���̏�Ԃ̕ύX
    cmd_exit.enabled = False
    cmd_start.Caption = "Stop"

    '�J�e�S�������i�[���Ă����\���̂́A�z�񐔍Ē�`�A������
    ReDim ChangeParamsInfo(PALS.LoopParams.CategoryCount)
    Call sub_Init_ChangeLoopParamsInfo

'�o���c�L������Ē�������ꍇ�A���̃t���O�֔��
LOOP_RETRY:

    '���[�v���s�񐔂������ϐ����C���N�������g
    intLoopTrialCnt = intLoopTrialCnt + 1

    '�ő�LOOP�񐔂̎擾
    g_MaxPalsCount = frm_PALS_LoopAdj_Main.txt_loop_num.Text

g_AnalyzeIgnoreCnt = 5

    Dim index As Long
    Dim sitez As Long
    '�����l�f�[�^�z�񐔂����ڐ��ōĒ�`
    For index = 0 To PALS.CommonInfo.TestCount
        For sitez = 0 To nSite
            Call PALS.CommonInfo.TestInfo(index).site(sitez).sub_ChangeDataDivision(g_MaxPalsCount)
        Next sitez
    Next index


    If Not Flg_AnalyzeDebug Then
        '######### Set DataLog
        '�f�[�^���O�t�@�C�����̐ݒ�
        Call sub_set_datalog(False)
        Call sub_set_datalog(True, PALS_PARAMFOLDERNAME_LOOP, "LoopAdjData")
        
        '######### Set RunOption
        'RunOption��Continue On Fail�ɕύX
        Call sub_exec_DoAll(True)
    End If
    
    
    '#################################################
    '##############   Main Measure   #################
    '#################################################
    
    Dim lngNowLoopCnt As Long               '���݂̃��[�v��
    Dim intFileNo As Integer                '�t�@�C���ԍ�
    Dim DatalogPosi As DatalogPosition      '�f�[�^���O�̊e���ڃf�[�^�ʒu��ۑ�����\����
    
    '���[�U�[�t�H�[���Ŏw�肵���񐔕��A�J��Ԃ�
    For lngNowLoopCnt = 1 To val(txt_loop_num.Text)
''        mSecSleep (100)

'>>>2011/05/16 K.SUMIYASHIKI ADD
        Call sub_InitActiveSiteInfo
'<<<2011/05/16 K.SUMIYASHIKI ADD

        If Flg_AnalyzeDebug Then
            g_strOutputDataText = frm_PALS_LoopAdj_Main.txt_AnalyzeDataPath.Text
            If Len(g_strOutputDataText) = 0 Then
                MsgBox ("Input analyze data path!!")
                Exit Sub
            End If
'            g_strOutputDataText = ""
        End If
        
        'PALS�̃G���[�t���O��True�ɂȂ��Ă����ꍇ�A����I��
        If g_ErrorFlg_PALS Then
            Exit For
        End If
        
        '�t�H�[���̐i���󋵗��X�V
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now TestRunning...  " & CStr(lngNowLoopCnt) & " / " & txt_loop_num.Text)
        
        If Not Flg_AnalyzeDebug Then
            '######### Run Test
            'IG-XL��Run�����s
            Call sub_exec_run
        End If
'        mSecSleep (500)
        mSecSleep (300)
    
        Dim lngDatalogFileValue As Long
        Dim WaitCnt As Long
        lngDatalogFileValue = 0
        
        For WaitCnt = 0 To 50
            If lngDatalogFileValue = FileLen(g_strOutputDataText) Then
                mSecSleep (100)
                If lngDatalogFileValue = FileLen(g_strOutputDataText) Then
                    Exit For
                End If
            Else
                lngDatalogFileValue = FileLen(g_strOutputDataText)
                mSecSleep (100)
            End If
        Next WaitCnt
    
        '######### 1��ڂ̑��莞�̂݁A�f�[�^���O���J��(�ǂݎ�胂�[�h�w��)
        If lngNowLoopCnt = 1 Then
                        
            '�t�@�C���ԍ��̎擾
            intFileNo = FreeFile
            
            '����f�[�^���O��Input(�ǂݍ���)���[�h�ŊJ��
            Open g_strOutputDataText For Input As #intFileNo
        End If
        
            
        '######### ����f�[�^���O�̓Ǎ�
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Reading...")
        '>>>2011/06/13 M.IMAMURA ContFailFlg Add.
        '>>>2011/08/04 K.SUMIYASHIKI UPDATE.
        'FAIL�������ڂ̃f�[�^��ǂݎ�肽���ꍇ
        If frm_PALS_LoopAdj_Main.Btn_ContinueOnFail.Value = True Then
            Call sub_ReadDatalog(lngNowLoopCnt, intFileNo, DatalogPosi, True)
        'FAIL�������ڂ̃f�[�^�͖�������ꍇ
        Else
            Call sub_ReadDatalog(lngNowLoopCnt, intFileNo, DatalogPosi, False)
        End If
        '<<<2011/08/04 K.SUMIYASHIKI UPDATE.
        '<<<2011/06/13 M.IMAMURA ContFailFlg Add.
        
        '######### �r��STOP���̏���
        If g_blnLoopStop Then
            txt_loop_num.Text = CStr(lngNowLoopCnt)
            lngNowLoopCnt = lngNowLoopCnt
            Exit For
        End If

        '######### �f�[�^��́����g���C
        '30��ڈȍ~����X�����͂��J�n���A���̌�1�񂨂��Ƀf�[�^��͂��s��
        '�ő僋�[�v���s�񐔂ɒB���Ă���ꍇ�́A�f�[�^��͍͂s��Ȃ�
        If (lngNowLoopCnt >= FIRST_VARIATION_CHECK_CNT) And (lngNowLoopCnt Mod VARIATION_CHECK_STEP = 0) _
            And intLoopTrialCnt < val(frm_PALS_LoopAdj_Main.txt_maxtrial_num.Text) Then
            
            '�X���m�F���s�����[�h(�f�t�H���g)�Ɏw�肳��Ă���ꍇ�A�f�[�^��͂��s��
            If op_AutoAdjust.Value Then
            
                '######### Analyze LoopData
                '3��/�K�i�����K��l���傫�����ڂ��Ȃ����`�F�b�N
                '1���ڂł��傫�����ڂ�����΁AFalse���Ԃ�
                If Not sub_CheckLoopData(lngNowLoopCnt) Then
                
                    '�e�J�e�S���̃p�����[�^���X���ɉ����ĕύX�ATestCondition�V�[�g�̒l���ύX
                    If sub_UpdataLoopParams = False Then
                        Call sub_errPALS("Updata LoopParameter error at 'sub_UpdataLoopParams'", "2-1-01-0-01")
                        Exit For
                    End If
                    
                    '����̌X���E��荞�݉񐔁EWait�����AChangeParamsInfo�ɕۑ�
                    '��荞�݉񐔁EWait�̕ύX�P�J�e�S���ł��������ꍇ�́ATrue���Ԃ�
                    If sub_Update_ChangeLoopParamsInfo Then
                    
                        '�t�@�C��(����f�[�^���O)�����
                        Close #intFileNo
                        
                        'TestCondition�V�[�g���̃f�[�^�𑪒�f�[�^���O�̖����ɒǉ�
                        Call sub_OutPutLoopParam(MeasureDatalogInfo)
                        
                        If Not Flg_AnalyzeDebug Then
                            '�f�[�^���O�̐ݒ���N���A
                            Call sub_set_datalog(False)
                        End If
                        
                        'csPALS�N���X�̉��
                        Set PALS = Nothing
                        
                        'csPALS�N���X���Ē�`
                        Set PALS = New csPALS
                        
                        'TestCondition�V�[�g�f�[�^�̍ēǍ�
                        Call ReadCategoryData
                        
                        '�t���O�̏�����
                        g_blnLoopStop = False
                        g_ErrorFlg_PALS = False
                        
                        '�đ�����{
                        GoTo LOOP_RETRY
                    
                    Else
                        '�X���m�F���s�����[�h(�t�H�[���̃{�^��)��False�ɕύX
                        op_AutoAdjust.Value = False
                    
                    End If
                    
                End If
            End If
        End If
    Next lngNowLoopCnt
    
    '�t�@�C��(����f�[�^���O)�����
    Close #intFileNo

    cmd_start.enabled = False
    
    '#################################################
    '#################################################

    'TestCondition�V�[�g���̃f�[�^�𑪒�f�[�^���O�̖����ɒǉ�
    Call sub_OutPutLoopParam(MeasureDatalogInfo)
    
    '######### Make LoopResultSheet
    '�t�H�[���̐i���󋵗��X�V
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Making LoopResultSheet...")
    
    'LOOP���[�o��
    '����������񐔂��o�͂�����ׁA�r��STOP���s�����ꍇ�ƁA�ʏ푪�莞�̑Ή��𕪂��Ă���
    If g_blnLoopStop Then
        '�r��STOP��
        Call sub_ShowLoopData(lngNowLoopCnt, MeasureDatalogInfo)
    Else
        '�ʏ푪�莞
        Call sub_ShowLoopData(lngNowLoopCnt - 1, MeasureDatalogInfo)
    End If
    
    '######### Reset DataLog
    If Not Flg_AnalyzeDebug Then
        '�f�[�^���O�̐ݒ���N���A
        Call sub_set_datalog(False)
    End If
    
   'Average�񐔁AWait���Ԃ��ő�ɐݒ肳��Ă����ꍇ�A���b�Z�[�W�{�b�N�X���o���m�点��
    Call sub_LoopParamsCheck
    
    '�t�H�[���̐i���󋵗��X�V
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Finished...", , True)
    
    cmd_start.Caption = "Start"
    cmd_start.enabled = True
    cmd_exit.enabled = True
'    cmd_start.Enabled = True

Exit Sub

errPALScmd_start_ClickLoop:
    Call sub_errPALS("Loop Tool Run error at 'cmd_Start_Click'", "2-1-01-0-02")

    '���Ƀt�@�C�����J���Ă����ꍇ�A�t�@�C�������
    If intFileNo <> 0 Then
        Close #intFileNo
    End If

'    cmd_start.Enabled = True
    
End Sub


'********************************************************************************************
' ���O : cmd_readloopdata_Click
' ���e : �w��f�[�^���O��LOOP���[���o��
' ���� : �Ȃ�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub cmd_readloopdata_Click()
    
On Error GoTo errPALScmd_readloopdata_Click

    'LOOP���[���o�͂�����f�[�^���O��I��
    Call sub_SetLoopData
    
    '�f�[�^���O��I�����Ȃ������ꍇ�A�֐��𔲂���
    If g_strOutputDataText = "False" Then
        Exit Sub
    End If
    
    Dim lngNowLoopCnt As Long               '���݂̃��[�v��
    Dim lngNowLoopEnd As Long               '�f�[�^��
    Dim intFileNo As Integer                '�t�@�C���ԍ�
    Dim DatalogPosi As DatalogPosition      '�f�[�^���O�̊e���ڃf�[�^�ʒu��ۑ�����\����
    Dim strbuf As String
    Dim MeasureDatalogInfo As DatalogInfo
    
    cmd_start.enabled = False
    
    '������
    lngNowLoopEnd = 0
    
    '�t�@�C���ԍ��̎擾
    intFileNo = FreeFile
    
    '######### �f�[�^���O���瑪��񐔂��擾
    Open g_strOutputDataText For Input As #intFileNo
    Do Until EOF(intFileNo)
        Line Input #intFileNo, strbuf
        If strbuf = DATALOG_END Then
            lngNowLoopEnd = lngNowLoopEnd + 1
            
        ElseIf InStr(1, strbuf, "MEASURE DATE : ") <> 0 Then
            MeasureDatalogInfo.MeasureDate = sub_GetMeasureData(strbuf, "Date")
            
        ElseIf InStr(1, strbuf, "JOB NAME     : ") <> 0 Then
            MeasureDatalogInfo.JobName = sub_GetMeasureData(strbuf, "JobName")
            
        ElseIf InStr(1, strbuf, "SW_NODE      : ") <> 0 Then
            MeasureDatalogInfo.SwNode = sub_GetMeasureData(strbuf, "Node")
            
        End If
    Loop
    Close #intFileNo

    '����f�[�^�������ꍇ�A�G���[��Ԃ��֐��𔲂���
    If lngNowLoopEnd = 0 Then
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Not Data Found...", True)
        Exit Sub
    End If
    
    '�ő�̑���񐔂��擾
    g_MaxPalsCount = lngNowLoopEnd
    
    Dim index As Long                   '�e�X�g���ڂ��������[�v�J�E���^
    Dim sitez As Long                   '�T�C�g���������[�v�J�E���^
    
    '�����l�f�[�^�z�񐔂����ڐ��ōĒ�`
    For index = 0 To PALS.CommonInfo.TestCount
        For sitez = 0 To nSite
            Call PALS.CommonInfo.TestInfo(index).site(sitez).sub_ChangeDataDivision(g_MaxPalsCount)
        Next sitez
    Next index
    
    '######### �f�[�^���O��ǂݍ���
    For lngNowLoopCnt = 1 To lngNowLoopEnd

'>>>2011/05/16 K.SUMIYASHIKI ADD
        Call sub_InitActiveSiteInfo
'<<<2011/05/16 K.SUMIYASHIKI ADD
        
        If lngNowLoopCnt = 1 Then
            intFileNo = FreeFile
            Open g_strOutputDataText For Input As #intFileNo
        End If
        '######### ����f�[�^���O�̓Ǎ�
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Reading...")
        '>>>2011/06/13 M.IMAMURA ContFailFlg Add.
        '>>>2011/08/04 K.SUMIYASHIKI UPDATE.
        'FAIL�������ڂ̃f�[�^��ǂݎ�肽���ꍇ
        If frm_PALS_LoopAdj_Main.Btn_ContinueOnFail.Value = True Then
            Call sub_ReadDatalog(lngNowLoopCnt, intFileNo, DatalogPosi, True)
        'FAIL�������ڂ̃f�[�^�͖�������ꍇ
        Else
            Call sub_ReadDatalog(lngNowLoopCnt, intFileNo, DatalogPosi, False)
        End If
        '<<<2011/08/04 K.SUMIYASHIKI UPDATE.
        '<<<2011/06/13 M.IMAMURA ContFailFlg Add.
    Next lngNowLoopCnt
    
    Close #intFileNo

    '######### Make LoopResultSheet
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Making LoopResultSheet...")
    Call sub_ShowLoopData(g_MaxPalsCount, MeasureDatalogInfo)

    '�t�H�[���̏�Ԃ̕ύX
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Finished...", , True)

    cmd_start.enabled = True

Exit Sub

errPALScmd_readloopdata_Click:
    Call sub_errPALS("Create Loop sheet error at 'cmd_readloopdata_Click'", "2-1-02-0-03")

    '���Ƀt�@�C�����J���Ă����ꍇ�A�t�@�C�������
    If intFileNo <> 0 Then
        Close #intFileNo
    End If
    
    cmd_start.enabled = True

End Sub

Private Sub cmd_exit_Click()
    Unload frm_PALS_LoopAdj_Main
End Sub

Private Sub txt_loop_num_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txt_loop_num_Change()
    txt_loop_num.Text = val(txt_loop_num.Text)
    If txt_loop_num.Text = "0" Then txt_loop_num.Text = "1"
'    If Val(txt_loop_num.Text) > 100 Then txt_loop_num.Text = "100"
End Sub

Private Sub txt_lot_name_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    If KeyAscii >= 65 And KeyAscii <= 90 Then Exit Sub
    If KeyAscii >= 97 And KeyAscii <= 122 Then Exit Sub

    KeyAscii = 0

End Sub


'********************************************************************************************
' ���O: sub_ShowLoopData
' ���e: �t�H�[���̃X�e�[�^�X���X�V���ALOOP���[���쐬
' ����: lngNowLoopCnt:����f�[�^��
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/20�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_ShowLoopData(ByVal lngNowLoopCnt As Long, ByRef MeasureDatalogInfo As DatalogInfo)
    
    '�t�H�[���̐i���󋵗��X�V
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Making...LoopResult")

    'LOOP���[�쐬
    Call sub_MakeLoopResultSheet(lngNowLoopCnt, MeasureDatalogInfo)

End Sub

