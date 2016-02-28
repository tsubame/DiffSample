Attribute VB_Name = "PALS_LoopAdj_Mod"
Option Explicit

'==========================================================================
' ���W���[�����F  PALS_LoopAdj_mod.bas
' �T�v        �F  LOOP�����Ŏg�p����֐��Q
' ���l        �F  �Ȃ�
' �X�V����    �F  Rev1.0      2010/09/30�@�V�K�쐬        K.Sumiyashiki
'==========================================================================

Public Const LOOPTOOLNAME As String = "Auto Loop Parameter Adjust"
Public Const LOOPTOOLVER As String = "1.41"

Public g_blnLoopStop As Boolean

Enum enum_DataTrendType
    em_trend_None       '�X������
    em_trend_Shift      '�V�t�g
    em_trend_Slope      '�㏸�E���~
    em_trend_Sudden     '��ђl
    em_trend_Uneven     '�o���c�L
End Enum


Public Const CLM_NO     As Integer = 1                '����No���L�������
Public Const CLM_TEST   As Integer = 2                '���ږ����L�������
Public Const CLM_UNIT   As Integer = 3                '�P�ʂ��L�������
Public Const CLM_CNT    As Integer = 4                '���[�v�񐔂��L�������
Public Const CLM_MIN    As Integer = 5                '�ŏ��l���L�������
Public Const CLM_AVG    As Integer = 6                '���ϒl���L�������
Public Const CLM_MAX    As Integer = 7                '�ő�l���L�������
Public Const CLM_SIGMA  As Integer = 8                '�Ђ��L�������
Public Const CLM_3SIGMA As Integer = 9                '3�Ђ��L�������
Public Const CLM_1PAR10 As Integer = 10               '�K�i��/10���L�������
Public Const CLM_LOW    As Integer = 11               '�����K�i���L�������
Public Const CLM_HIGH   As Integer = 12               '����K�i���L�������
Public Const CLM_3SIGMAPARSPEC As Integer = 13        '3��/�K�i�����L�������
'>>>2010/12/13 K.SUMIYASHIKI ADD
Public Const CLM_JUDGELIMIT As Integer = 14           '���[�v�o���c�L�̔��f����x�����L�������(Test Instances��LoopJudgeLimit���L�������)
'<<<2010/12/13 K.SUMIYASHIKI ADD

Public Const ROW_NAME   As Integer = 1                '����Lot�����L������s
Public Const ROW_WAFER  As Integer = 2                '�E�F�[�nNo���L������s
Public Const ROW_MACHINEJOB As Integer = 3            '���葕�u�AJOB�����L������s
Public Const ROW_LOOPCOUNT As Integer = 4             '���[�v��(�Ssite���v)���L������s
Public Const ROW_DATE   As Integer = 5                '��������L������s
Public Const ROW_LABEL  As Integer = 6                '���[�v���ʂ̃��x�����L������s
Public Const ROW_DATASTART As Integer = ROW_LABEL + 1 '���[�v���ʂ̃f�[�^���L������擪�s

Public Const MODE_AUTO As String = "AUTO"


Public Type ChangeParamsInformation
    MinWait As Double                        '�ݒ�\�ȍŏ��E�F�C�g
    MaxWait As Double                        '�ݒ�\�ȍő�E�F�C�g
    WaitTrialCnt As Integer                  'Wait�ύX��
    AveTrialCnt As Integer                   'Average�ύX��
    Pre_Average As Integer                   '�O��̎�荞�݉�
    Pre_Wait As Double                       '�O��̎�荞�ݑO�E�F�C�g
    Pre_VariationTrend As enum_DataTrendType '�O���莞�̃o���c�L�X��
    Flg_WaitFinish As Boolean                'Wait���������t���O
    Flg_AverageFinish As Boolean             '��荞�݉񐔒��������t���O
End Type


'�P�ʊ��Z�p�W��
'Private Const TERA   As Double = 1000000000000#         '�e��
'Private Const GIGA   As Long = 1000000000               '�M�K
Private Const MEGA    As Long = 1000000                  '���K
Private Const KIRO    As Long = 1000                     '�L��
Private Const MILLI   As Double = 0.001                  '�~��
Private Const MAICRO  As Double = 0.000001               '�}�C�N��
Private Const NANO    As Double = 0.000000001            '�i�m
Private Const PIKO    As Double = 0.000000000001         '�s�R
Private Const FEMTO   As Double = 0.000000000000001      '�t�F���g
Private Const percent As Double = 0.01

Private Const LABEL_GRADE As String = "grade"

Public g_MaxPalsCount As Long     '�t�H�[���ɓ��͂����ő呪���(�f�t�H���g:100��)

Public Const FIRST_VARIATION_CHECK_CNT As Integer = 25        '�ŏ��ɌX�����͂��s����(�f�t�H���g:30��)
Public Const VARIATION_CHECK_STEP As Integer = 1            '�X�����͂��s���X�e�b�v(�f�t�H���g:30��ȍ~�A������)

Public ChangeParamsInfo() As ChangeParamsInformation        '�e�J�e�S���̃p�����[�^���ڂ�ۑ�����ϐ�

'����f�[�^�̏���ۑ�����\����
Public Type DatalogInfo
    MeasureDate As String
    JobName     As String
    SwNode      As String
End Type

Public g_AnalyzeIgnoreCnt As Integer

Public Sub sub_LoopFrmShow()
    frm_PALS_LoopAdj_Main.Show
End Sub

'********************************************************************************************
' ���O : sub_SetLoopData
' ���e : �_�C�A���O����LOOP���[������f�[�^���O��I�����A�t�@�C���p�X���擾
' ���� : �Ȃ�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_SetLoopData()

On Error GoTo errPALSsub_SetLoopData

    g_strOutputDataText = ""

    g_strOutputDataText = Application.GetOpenFilename( _
        title:="!!!!!!!!!!!!!!!!!!!!   Select Target LoopData   !!!!!!!!!!!!!!!!!!!!", _
        fileFilter:="IP750 LoopDataFile (*.txt), *.txt ")

Exit Sub

errPALSsub_SetLoopData:
    Call sub_errPALS("Set datalog name error at 'sub_SetLoopData'", "2-2-01-0-04")

End Sub


'********************************************************************************************
' ���O : sub_CheckLoopData
' ���e : �_�C�A���O����LOOP���[������f�[�^���O��I�����A�t�@�C���p�X���擾
' ���� : lngNowLoopCnt:����̑����
' �ߒl�F True  :�o���c�L�Ȃ�
'        False :�o���c�L����
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Function sub_CheckLoopData(ByVal lngNowLoopCnt As Long) As Boolean
    
    '�Ԃ�l��True�ŏ�����
    '�o���c�L�ɖ�肪�Ȃ����True���Ԃ�
    sub_CheckLoopData = True
    
    If g_ErrorFlg_PALS Then
        Exit Function
    End If
    
On Error GoTo errPALSsub_CheckLoopData
    
    Dim TestNo As Long          '���[�v�J�E���^(�e�X�g���ڂ�����)
    Dim sitez As Long           '���[�v�J�E���^(Site�ԍ�������)
    
    '���[�U�[�t�H�[���̃X�e�[�^�X��ύX
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Checking...")

    '�S���ڂ��J��Ԃ�
    For TestNo = 0 To PALS.CommonInfo.TestCount
        
        'DC�ȊO�̃f�[�^�X�����f���s��
        If PALS.CommonInfo.TestInfo(TestNo).CapCategory1 <> "DC" _
            Or Len(PALS.CommonInfo.TestInfo(TestNo).CapCategory1) > 0 Then                       'DC��萔��
            
            '�SSite�J��Ԃ�
            For sitez = 0 To nSite
                '3��/�K�i����1���ڂł��K��l�𒴂��Ă����ꍇ�A�Ԃ�l��False�ɕύX
                If Not sub_JudgeLoopData(TestNo, sitez, lngNowLoopCnt) Then
                    sub_CheckLoopData = False
                End If
            Next sitez
        End If
    Next TestNo

    If sub_CheckLoopData = False And g_AnalyzeIgnoreCnt > 0 Then
        g_AnalyzeIgnoreCnt = g_AnalyzeIgnoreCnt - 1
        sub_CheckLoopData = True
    End If

Exit Function

errPALSsub_CheckLoopData:
    Call sub_errPALS("Check LoopData error at 'sub_CheckLoopData'", "2-2-02-0-05")

End Function


'********************************************************************************************
' ���O: sub_JudgeLoopData
' ���e: ����"lngTestNo"��"sitez"�œn���ꂽ���ځE�T�C�g�̓����l�E�K�i������A3��/�K�i�����v�Z���A
'       ���̌��ʂ����e�͈�(��{��0.1)�𒴂����ꍇ�A�o���c�L�X�����f���s���B
'       �o���c�L���������ꍇ�A�e�J�e�S���̃o���c�L���(�ő�o���c�L���ړ�)���X�V����B
' ����: lngTestNo      : ���ڂ������ԍ�
'       sitez          : �T�C�g�ԍ�
'       lngNowLoopCnt  : ����ς݉�
' �ߒl: True  : �o���c�L���Ȃ�
'       False : �o���c�L����
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_JudgeLoopData(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As Boolean

    '�Ԃ�l��True�ŏ�����
    '�o���c�L�ɖ�肪�Ȃ����True���Ԃ�
    sub_JudgeLoopData = True

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_JudgeLoopData

    Dim dblStandardWidth As Double          '�K�i��
    Dim enumJudge As enum_DataTrendType     '�X�����ނ������񋓑�
    
    '������
    enumJudge = em_trend_None
    
    With PALS.CommonInfo.TestInfo(lngTestNo)
    
        Select Case .arg2
            '�K�i����
            Case 0
                sub_JudgeLoopData = True
                Exit Function
            
            '�����K�i�̂�
            Case 1
                dblStandardWidth = Abs(.LowLimit)
            
            '����K�i�̂�
            Case 2
                dblStandardWidth = Abs(.HighLimit)
            
            '�㉺���K�i����
            Case 3
                dblStandardWidth = .HighLimit - .LowLimit
    
            Case Else
                Call sub_errPALS("Get standard width error at 'sub_JudgeLoopData'", "2-2-03-2-06")
        End Select
        
        '�K�i����0�ȊO�̎��̂ݎ��s�i0����h�~�j
        If dblStandardWidth <> 0 Then
        
            '3��/�K�i�����K��l�ȏ�̏ꍇ�A�X���m�F���s��
            If ((.site(sitez).Sigma * 3# / dblStandardWidth)) >= .LoopJudgeLimit And .LoopJudgeLimit <> 0 Then
    
                '�o���c�L������ꍇ�AFalse��Ԃ�l�ɐݒ�
                sub_JudgeLoopData = False
    
                '�X�����m�F����ׂɁA�w��񐔂͌X�����f�֍s���Ȃ�
                If g_AnalyzeIgnoreCnt > 0 Then
                    Exit Function
                End If
    
                '�X���𔻒f���A�X������(�񋓑�:enum_DataTrendType�Œ�`)��Ԃ�l�Ƃ��ĕԂ�
                enumJudge = sub_AnalyzeLoopData(lngTestNo, sitez, lngNowLoopCnt)
    
                '������ŉ��P���s���ׂɕK�v�Ȋe�J�e�S���[�̏����X�V
                If Not sub_UpdateVariationLoopData(lngTestNo, sitez, dblStandardWidth, enumJudge) Then
                    '�r���ŃG���[�ɂȂ����ꍇ�A�G���[��Ԃ�
                    Call sub_errPALS("Update variation Loopdata error at 'sub_UpdateVariationLoopData'", "2-2-03-0-07")
                End If
            End If
        End If
    
    End With
        
Exit Function

errPALSsub_JudgeLoopData:
    Call sub_errPALS("Judge LoopData error at 'sub_JudgeLoopData'", "2-2-03-0-08")

End Function
        

'********************************************************************************************
' ���O: sub_AnalyzeLoopData
' ���e: �o���c�L�̌X�����m�F����֐�
' ����: lngTestNo         : ���ڂ������ԍ�
'       sitez             : �T�C�g�ԍ�
'       lngNowLoopCnt     : ����ς݉�
' �ߒl: �o���c�L�X���������񋓑�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
' �@�@�@�@�@ Rev2.0      2011/06/20�@�����ύX   K.Sumiyashiki
'                                    ��F������g�p���Ă̔��f�A���S���Y���֕ύX
'                                      (�֐��S�̂̃t���[��ύX)
'********************************************************************************************
Private Function sub_AnalyzeLoopData(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As enum_DataTrendType

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_AnalyzeLoopData

    '������
    sub_AnalyzeLoopData = em_trend_None
    
    With PALS.CommonInfo.TestInfo(lngTestNo).site(sitez)
        Dim dbl_F_Data As Double
        '�ʏ�̓����l�΂���Ɣ����l�̃��[�v�΂�����AF�l���Z�o
        dbl_F_Data = ((.Sigma ^ 2) * 2) / (.Differential_Sigma(lngNowLoopCnt) ^ 2)
    End With
        
    'F������A�����_�����L��Ɣ��f
    If sub_Get_F_Value(lngNowLoopCnt, 2, "bottom") < dbl_F_Data And dbl_F_Data < sub_Get_F_Value(lngNowLoopCnt, 2, "top") Then

'>>>2010/12/13 K.SUMIYASHIKI ADD
        '��ђl�Ɣ��f�����悤�ȑ傫�ȃo���c�L�𔻒f
        sub_AnalyzeLoopData = sub_JudgeBaratuki(lngTestNo, sitez, lngNowLoopCnt)
'<<<2010/12/13 K.SUMIYASHIKI ADD


'>>>2010/12/13 K.SUMIYASHIKI UPDATE
        '�傫�ȃo���c�L�Ŗ����ꍇ�̂ݏ���
        If sub_AnalyzeLoopData = em_trend_None Then
            '��ђl�̔��f
            '->���f�B�A��������A���̕���+2�Јȏ�̂��̂�����Δ�ђl�Ɣ��f
            sub_AnalyzeLoopData = sub_JudgeTobiti(lngTestNo, sitez, lngNowLoopCnt)
        End If
'<<<2010/12/13 K.SUMIYASHIKI UPDATE

    'F������A�����_���������Ɣ��f
    Else
        '��ђl�Ŗ����ꍇ�̂ݏ���
        If sub_AnalyzeLoopData = em_trend_None Then
            '�V�t�gor�㏸or���~�̔��f
            sub_AnalyzeLoopData = sub_JudgeShift(lngTestNo, sitez, lngNowLoopCnt)
        End If

    End If
    
    '��ђl�E�V�t�g�E�㏸�E���~�Ŗ����ꍇ�A�o���c�L�Ɣ��f
    If sub_AnalyzeLoopData = em_trend_None Then
        '�o���c�L�������l��Ԃ�l�ɐݒ�
        sub_AnalyzeLoopData = em_trend_Uneven
        Debug.Print ("Baratuki")
        Debug.Print ("TestName : " & PALS.CommonInfo.TestInfo(lngTestNo).tname)
        Debug.Print ("Site     : " & sitez) & vbCrLf
    End If
    
Exit Function

errPALSsub_AnalyzeLoopData:
    Call sub_errPALS("Analyze Loopdata error at 'sub_AnalyzeLoopData'", "2-2-04-0-09")
    
End Function


'********************************************************************************************
' ���O: sub_UpdateVariationLoopData
' ���e: �o���c�L���������ꍇ�A�o���c�L���������J�e�S���̃o���c�L�f�[�^���m�F���A
'       �����ڂ̃o���c�L�f�[�^�Ɣ�r����B���̍ہA�ȑO�̃f�[�^���o���c�L���傫���A���A�f�[�^�̌X����������΁A
'       ����̍��ڂ̃o���c�L�f�[�^�ōX�V���s���B
'       �X�V���e:�X���E�o���c�L�l�E���ږ��E�T�C�g���B
'       �f�[�^�̃o���c�L�́A�o���c�L�˔�ђl�ˏ㏸�E���~�˃V�t�g�̏��ɑΉ����s���B
' ����: lngTestNo         : ���ڂ������ԍ�
'       sitez             : �T�C�g�ԍ�
'       dblStandardWidth  : �K�i��
'       enumJudge         : �o���c�L�X��
' �ߒl: True  : �o���c�L���Ȃ�
'       False : �o���c�L����
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_UpdateVariationLoopData(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal dblStandardWidth As Double, _
                                        ByVal enumJudge As enum_DataTrendType) As Boolean
                
    If g_ErrorFlg_PALS Then
        sub_UpdateVariationLoopData = True
        Exit Function
    End If
                
On Error GoTo errPALSsub_UpdateVariationLoopData
                
    Dim colTargetCategory As New Collection     '�J�e�S���[���i�[����R���N�V����
    
    With PALS.CommonInfo.TestInfo(lngTestNo)
        '�R���N�V������CapCategory1�̒l(ex:OF,ML)��ǉ�
        colTargetCategory.Add Item:=PALS.LoopParams.CategoryInfoList(.CapCategory1)
        If Len(.CapCategory2) Then
            'CapCategory2�ɒl(ex:OF,ML)���L�q����Ă���΁A�R���N�V�����ɒǉ�
            colTargetCategory.Add PALS.LoopParams.CategoryInfoList(.CapCategory2)
        End If
    End With
        
    Dim valTargetCategory As Variant                '�R���N�V�����Ɋi�[����Ă���J�e�S���[��������
    Dim enumCategoryTrend As enum_DataTrendType     '�I�����ꂽ���ڂ̃f�[�^�X�����i�[
    
    '�f�[�^���R���N�V�������J��Ԃ�
    For Each valTargetCategory In colTargetCategory
    
        '����̎w��J�e�S���[�̃f�[�^�X��(�ň��l)��enumCategoryTrend�Ɉꎞ�i�[
        enumCategoryTrend = PALS.LoopParams.LoopCategory(valTargetCategory).VariationTrend
                
        With PALS.LoopParams.LoopCategory(valTargetCategory)
            '����̎w��J�e�S���[�̃f�[�^�X��(�ň��l)���A����̃f�[�^�X���������ꍇ�́A�e�f�[�^���㏑��
            '�f�[�^�X���������ꍇ�́A3��/�K�i�����r���A����̕��������ꍇ�́A�e�f�[�^���㏑��
            '�f�[�^�X���̔�r����->�΂���˔�ђl�ˏ㏸�E���~�˃V�t�g
            If enumCategoryTrend = enumJudge Then
                '�ȑO��3��/�K�i���̃f�[�^�Ɣ�r���A���񂪈�����Ίe�f�[�^���㏑��
                If .VariationLevel < PALS.CommonInfo.TestInfo(lngTestNo).site(sitez).Sigma * 3# / dblStandardWidth Then
                    .VariationLevel = PALS.CommonInfo.TestInfo(lngTestNo).site(sitez).Sigma * 3# / dblStandardWidth
                    .TargetTestName = PALS.CommonInfo.TestInfo(lngTestNo).tname
                    .VariationSite = sitez
                End If
            ElseIf enumCategoryTrend < enumJudge Then
                '����̃f�[�^�X���̕��������ꍇ�́A�e�f�[�^���㏑��
                .VariationLevel = PALS.CommonInfo.TestInfo(lngTestNo).site(sitez).Sigma * 3# / dblStandardWidth
                .TargetTestName = PALS.CommonInfo.TestInfo(lngTestNo).tname
                .VariationSite = sitez
                .VariationTrend = enumJudge
            End If
        End With
    Next valTargetCategory

    '�r���ŃG���[��������΁ATrue��Ԃ�
    sub_UpdateVariationLoopData = True

Exit Function

errPALSsub_UpdateVariationLoopData:
    Call sub_errPALS("Update variation Loopdata error at 'sub_UpdateVariationLoopData'", "2-2-05-0-10")

End Function


'********************************************************************************************
' ���O: sub_JudgeTobiti
' ���e: �o���c�L�X������ђl���ǂ����m�F����֐�
'       ����f�[�^��3�^�b�v�̃��f�B�A���t�B���^���|���A���f�B�A��������̒l�ƌ��̒l�̌��Z���s���B
'       ���̍ۂ̍�����2�Јȏ゠��΁A��ђl�Ƃ���B
' ����: lngTestNo         : ���ڂ������ԍ�
'       sitez             : �T�C�g�ԍ�
'       lngNowLoopCnt     : ����ς݉�
' �ߒl: �o���c�L�X���������񋓑�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_JudgeTobiti(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As enum_DataTrendType

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_JudgeTobiti

'��ђl�̔��f
'->���f�B�A��������A���̕���+2�Јȏ�̂��̂�����Δ�ђl�Ɣ��f

    '���f�B�A���^�b�v���̐ݒ�(�Ƃ肠�����̓��[�J���̐ÓI�ϐ��Œ�`)
    Const Median_Tap As Integer = 3
    
    '���f�B�A���^�b�v�������� or 1 �̏ꍇ�G���[���b�Z�[�W��\��
    If Median_Tap Mod 2 = 0 Or Median_Tap = 1 Then
        Call sub_errPALS("Median tap number is even number or 1." & vbCrLf & "         Please check median tap number !" & vbCrLf & "         at 'sub_JudgeTobiti'", "2-2-06-5-11")
        Exit Function
    End If

    '���f�B�A�����̏��O�͈͐ݒ�
    Dim intRemoveArea As Integer
    intRemoveArea = Int(Median_Tap / 2)

    '���f�B�A����̃f�[�^�i�[�z��
    '���O�͈͕��A�z�񐔂��폜
    Dim dblConvertData() As Double
    ReDim dblConvertData(lngNowLoopCnt - intRemoveArea)

    Dim data_cnt As Long            '���[�v�J�E���^(�f�[�^�C���f�b�N�X������)
    Dim tap_cnt As Long             '���[�v�J�E���^(���f�B�A�����̃^�b�v�ԍ�������)
    Dim dblTmpData() As Double      '�^�b�v�����̃f�[�^���ꎞ�i�[����z��
    
    '���O�͈͂��������ӏ����J��Ԃ�
    For data_cnt = 1 + intRemoveArea To lngNowLoopCnt - intRemoveArea
        
        '���f�B�A���^�b�v���ɉ����čĒ�`�y�я�����
        ReDim dblTmpData(Median_Tap - 1)
        
        '�^�b�v�����J��Ԃ�
        For tap_cnt = 0 To Median_Tap - 1
            'data_cnt�̑O��^�b�v�����̃f�[�^��dblTmpData�Ɉꎞ�i�[
            dblTmpData(tap_cnt) = PALS.CommonInfo.TestInfo(lngTestNo).site(sitez).Data(data_cnt - intRemoveArea + tap_cnt)
        Next tap_cnt

        'dblTmpData���~���Ńo�u���\�[�g
        Call sub_BubbleSort(dblTmpData)

        '�����l��dblConvertData�ɑ��
        dblConvertData(data_cnt) = dblTmpData(UBound(dblTmpData) - intRemoveArea)

    Next data_cnt

    '���f�B�A�������ɂ��A����J�n����̏��O�G���A��⊮
    For data_cnt = 1 To intRemoveArea
        dblConvertData(data_cnt) = dblConvertData(intRemoveArea + 1)
    Next data_cnt

    With PALS.CommonInfo.TestInfo(lngTestNo).site(sitez)
        '���݂̑���񐔕��A�f�[�^���J��Ԃ�
        For data_cnt = 1 To UBound(dblConvertData)
            '���f�B�A��������̃f�[�^���猳�̕��ϒl�����Z���A���̒l��2�Јȏ�̏ꍇ��ђl�Ɣ��f
            If (Abs((.Data(data_cnt) - dblConvertData(data_cnt))) - (.Sigma * 2)) > 0 Then
                Debug.Print ("Tobiti error")
                Debug.Print ("TestName : " & PALS.CommonInfo.TestInfo(lngTestNo).tname)
                Debug.Print ("Site     : " & sitez) & vbCrLf
'                Debug.Print ("Tobiti error")
                sub_JudgeTobiti = em_trend_Sudden
                
            End If
        Next data_cnt
    End With

Exit Function

errPALSsub_JudgeTobiti:
    Call sub_errPALS("Check LoopData error at 'sub_JudgeTobiti'", "2-2-06-0-12")

End Function


'********************************************************************************************
' ���O: sub_BubbleSort
' ���e: �o�u���\�[�g
' ����: dblVal         : ���ёւ����s���z��
'       blnSortAsc     : ����or�~�����w�肷��I�v�V����
'                        (�f�t�H���g��False:����)
' �ߒl: �Ȃ�
' ���l�F�\�[�g��̌��ʗၫ��
'       False �� dblVal(1):10, dblVal(2):5, dblVal(3):1
'       True  �� dblVal(1):1 , dblVal(2):5, dblVal(3):10
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_BubbleSort(ByRef dblVal() As Double, Optional ByVal blnSortAsc As Boolean = False)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_BubbleSort

    Dim i As Long       '���[�v�J�E���^
    Dim j As Long       '���[�v�J�E���^

    For i = LBound(dblVal) To UBound(dblVal) - 1
        For j = LBound(dblVal) To LBound(dblVal) + UBound(dblVal) - i - 1
            If dblVal(IIf(blnSortAsc, j, j + 1)) > dblVal(IIf(blnSortAsc, j + 1, j)) Then
                Call sub_Swap(dblVal(j), dblVal(j + 1))
            End If
        Next j
    Next i

Exit Sub

errPALSsub_BubbleSort:
    Call sub_errPALS("BubbleSort error at 'sub_BubbleSort'", "2-2-07-0-13")

End Sub


'********************************************************************************************
' ���O: sub_Swap
' ���e: �����œn���ꂽ2�̒l�����ւ���֐�
' ����: dblVal1 : ����ւ���ϐ�1
'       dblVal2 : ����ւ���ϐ�2
' �ߒl: �Ȃ�
' ���l�F�Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_Swap(ByRef dblVal1 As Double, ByRef dblVal2 As Double)

    Dim dblBuf As Double    '�ꎞ�i�[�ϐ�
    
    dblBuf = dblVal1
    dblVal1 = dblVal2
    dblVal2 = dblBuf

End Sub


'********************************************************************************************
' ���O: sub_JudgeShift
' ���e: �o���c�L�X�����㏸or���~or�V�t�g���ǂ����m�F����֐�
'       ����f�[�^�̊J�n�t�߂̃f�[�^�ƁA���݂̑���񐔕t�߂̃f�[�^�̕��ϒl���r���A
'   �@�@���̍�����1�Јȏ゠�����ꍇ�A�㏸or���~or�V�t�g�Ɣ��f����B
'       ���ϒl�����ۂ̃f�[�^���́A���݂̑���񐔂�1/10�Ƃ��Ă���B
' ����: lngTestNo         : ���ڂ������ԍ�
'       sitez             : �T�C�g�ԍ�
'       lngNowLoopCnt     : ����ς݉�
' �ߒl: �o���c�L�X���������񋓑�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_JudgeShift(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As enum_DataTrendType

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_JudgeShift

    Dim dblBeforeDataSum As Double    '����J�n�t�߂̃f�[�^���v(�f�[�^����intCalcWidth�Őݒ�)
    Dim dblAfterDataSum As Double     '���݂̑���񐔕t�߂̃f�[�^���v(�f�[�^����intCalcWidth�Őݒ�)
    Dim dblBeforeDataAve As Double    '����J�n�t�߂̃f�[�^����(�f�[�^����intCalcWidth�Őݒ�)
    Dim dblAfterDataAve As Double     '���݂̑���񐔕t�߂̃f�[�^����(�f�[�^����intCalcWidth�Őݒ�)

    Dim data_cnt As Long                '���[�v�J�E���^(�f�[�^�C���f�b�N�X������)
    Dim intCalcWidth As Integer         '���f���s���ۂɎg�p����f�[�^��

    '�f�[�^�����A���݂̑���񐔂�1/10�ɐݒ�
    intCalcWidth = Int(lngNowLoopCnt / 10)

    With PALS.CommonInfo.TestInfo(lngTestNo).site(sitez)

        '����J�n�t�߂̃f�[�^�̍��v���擾(�f�[�^����intCalcWidth�Őݒ�)
        For data_cnt = 1 To intCalcWidth
            dblBeforeDataSum = dblBeforeDataSum + .Data(data_cnt)
        Next data_cnt
        '���ϒl���v�Z
        dblBeforeDataAve = (dblBeforeDataSum / intCalcWidth)
    
        '���݂̑���񐔕t�߂̃f�[�^�̍��v���擾(�f�[�^����intCalcWidth�Őݒ�)
        For data_cnt = lngNowLoopCnt - intCalcWidth + 1 To lngNowLoopCnt
            dblAfterDataSum = dblAfterDataSum + .Data(data_cnt)
        Next data_cnt
        '���ϒl���v�Z
        dblAfterDataAve = (dblAfterDataSum / intCalcWidth)
    
        '����J�n���̃f�[�^���ςƌ��݂̑���񐔕t�߂̃f�[�^���ς��r
        '������1�Јȏ゠�����ꍇ�A�V�t�gor�㏸or���~�Ɣ��f
        '1�Јȉ��̏ꍇ�A�֐��𔲂���
        If (Abs(dblBeforeDataAve - dblAfterDataAve) < .Sigma * 2) Then
            sub_JudgeShift = em_trend_None
            Exit Function
        End If
    End With

    '�V�t�gor�㏸or���~�̏ꍇ�A�f�[�^���ǂ̃p�^�[�������f����
    sub_JudgeShift = sub_CheckShiftType(lngTestNo, sitez, lngNowLoopCnt)

Exit Function

errPALSsub_JudgeShift:
    Call sub_errPALS("Data Judge error at 'sub_JudgeShift'", "2-2-08-0-14")

End Function


'********************************************************************************************
' ���O: sub_CheckShiftType
' ���e: �o���c�L�X�����㏸or���~���V�t�g���ǂ����؂蕪�����s���֐�
'       ����f�[�^�ɑ΂��A���т̃f�[�^�Ƃ̍�������鏈�����s���A���̃f�[�^���ɁA
'       1�Јȏ�̍��������݂����ꍇ�A�V�t�g�Ɣ��f����B
'       1�Јȏ�̍������Ȃ��ꍇ�A���X�ɕϓ������Ƃ��A�㏸or���~�Ɣ��f����B
' ����: lngTestNo         : ���ڂ������ԍ�
'       sitez             : �T�C�g�ԍ�
'       lngNowLoopCnt     : ����ς݉�
' �ߒl: �o���c�L�X���������񋓑�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_CheckShiftType(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As enum_DataTrendType

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_CheckShiftType

    '���̃f�[�^�Ƃ̍����l���i�[����z��
    Dim dblConvertData() As Double
    '�Ō�̃f�[�^�͍��������߂��Ȃ��ׁA���݂̑���񐔂���-1�����z����Ē�`
    ReDim dblConvertData(lngNowLoopCnt - 1)

    Dim data_cnt As Long        '���[�v�J�E���^(�f�[�^�C���f�b�N�X������)
    
    With PALS.CommonInfo.TestInfo(lngTestNo).site(sitez)
    
        '�f�[�^���}�C�i�X1�����J��Ԃ�
        For data_cnt = 2 To lngNowLoopCnt - 1
            
            '���т̃f�[�^�Ƃ̍������擾
            dblConvertData(data_cnt) = .Data(data_cnt + 1) - .Data(data_cnt - 1)
            
            '�f�[�^�̍������Јȏ�̏ꍇ�V�t�g�Ƃ���
            If Abs(dblConvertData(data_cnt)) - (.Sigma) > 0 Then
                '�V�t�g�������l��Ԃ��I��
                sub_CheckShiftType = em_trend_Shift
                Call sub_errPALS("This data trend is Shift." & vbCrLf & "Please check data!", "2-2-09-7-15")
                Debug.Print ("Shift error")
                Debug.Print ("TestName : " & PALS.CommonInfo.TestInfo(lngTestNo).tname)
                Debug.Print ("Site     : " & sitez) & vbCrLf
                Exit Function
            End If
        Next data_cnt
    End With

    '�V�t�g�ł͂Ȃ��ꍇ�A�㏸or���~�������l��Ԃ�
    sub_CheckShiftType = em_trend_Slope
        Debug.Print ("Rise or Fall!")
        Debug.Print ("TestName : " & PALS.CommonInfo.TestInfo(lngTestNo).tname)
        Debug.Print ("Site     : " & sitez) & vbCrLf
    Call sub_errPALS("This data trend is Rise or Fall." & vbCrLf & "Please check data!")

Exit Function

errPALSsub_CheckShiftType:
    Call sub_errPALS("Data Judge error at 'sub_CheckShiftType'", "2-2-09-7-16")

End Function


'********************************************************************************************
' ���O : sub_UpdataLoopParams
' ���e : �X�����f�̌��ʂ���A�J�e�S������Wait�AAverage�̕ύX���s��
' ���� : �Ȃ�
' �ߒl�F True  :�G���[�Ȃ�
'        False :�G���[����
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
'�o���c�L�̌X���m�F���ʂ��p�����[�^�֔��f
Public Function sub_UpdataLoopParams() As Boolean

    '�Ԃ�l�̏������B�G���[���Ȃ����True���Ԃ�
    sub_UpdataLoopParams = True

    If g_ErrorFlg_PALS Then
        sub_UpdataLoopParams = True
        Exit Function
    End If
    
On Error GoTo errPALSsub_UpdataLoopParams

    Dim cnt As Long             '���[�v�J�E���^�B�J�e�S���C���f�b�N�X�������B
    Dim dblAdjValue As Double   '�p�����[�^��ύX����ۂɁA�ꎞ�I�ɒl��ۑ�����ϐ�
    Dim dblStep As Double       'Wait��ύX����ۂ̃X�e�b�v���B�ŏ�Wait�ƍő�Wait����v�Z�B

    With PALS.LoopParams
        
        '�J�e�S���̐������J��Ԃ�
        For cnt = 1 To .CategoryCount
    
            '������
            dblAdjValue = -1
            dblStep = -1
            dblAdjValue = 0
    
            Select Case .LoopCategory(cnt).VariationTrend
                
                '**********����̌X������ђl�̏ꍇ**********
                Case enum_DataTrendType.em_trend_Sudden
                    
                    'Wait�̃X�e�b�v�������߂�
''                    dblStep = Format((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 6, "#.000")
                    dblStep = Int((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 6)
                    
                    'Wait�̏����l���ݒ�ő�Wait�̏ꍇ�A���������Ƃ���
                    If dblStep = 0 Then
                        ChangeParamsInfo(cnt).Flg_WaitFinish = True
                    End If
                    
                    '�ȑO��Wait�̕ύX���s���Ă��Ȃ��ꍇ
                    If ChangeParamsInfo(cnt).WaitTrialCnt = 0 Then
''                        dblAdjValue = .LoopCategory(cnt).Wait + Format((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 2, "#.000")
                        dblAdjValue = .LoopCategory(cnt).WAIT + Int((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 2)
    
                        'TestCondition�V�[�g���̃p�����[�^��ύX
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                        
                        'Wait�ύX���s�����񐔂̃C���N�������g
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1

                    '�ȑO��Wait�̕ύX��1��s���Ă���ꍇ
                    ElseIf ChangeParamsInfo(cnt).WaitTrialCnt = 1 Then
                        dblAdjValue = .LoopCategory(cnt).WAIT + (dblStep * 2)
                                            
                        'TestCondition�V�[�g���̃p�����[�^��ύX
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                    
                        'Wait�ύX���s�����񐔂̃C���N�������g
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                    
                    '�ȑO��Wait�̕ύX��2��s���Ă���ꍇ
                    ElseIf ChangeParamsInfo(cnt).WaitTrialCnt = 2 Then
                        dblAdjValue = .LoopCategory(cnt).WAIT + dblStep
                                            
                        'TestCondition�V�[�g���̃p�����[�^��ύX
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                    
                        'Wait�ύX���s�����񐔂̃C���N�������g
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                    
                    '�ȑO��Wait�̕ύX��3��ȏ�s���Ă���ꍇ
                    ElseIf ChangeParamsInfo(cnt).WaitTrialCnt >= 3 Then
                        
                        If Not ChangeParamsInfo(cnt).Flg_WaitFinish Then
                        
''                            dblAdjValue = .LoopCategory(cnt).Wait + Format(dblAdjValue + (dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2))), "#.000")
''                            dblAdjValue = .LoopCategory(cnt).Wait + Int(dblAdjValue + (dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2))))
                            dblAdjValue = .LoopCategory(cnt).WAIT + Int(dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2)))
                            
                            'TestCondition�V�[�g���̃p�����[�^��ύX
                            If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                                sub_UpdataLoopParams = False
                                Exit Function
                            End If
                        
                            'Wait�ύX���s�����񐔂̃C���N�������g
                            ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                                            
                            If dblAdjValue > (ChangeParamsInfo(cnt).MaxWait * 0.99) Then
                                If dblAdjValue > ChangeParamsInfo(cnt).MaxWait Then
                                    dblAdjValue = ChangeParamsInfo(cnt).MaxWait
                                End If
                                ChangeParamsInfo(cnt).Flg_WaitFinish = True
                            End If
                                            
                        End If
                    Else
                    
                    
                    End If
    

                '**********����̌X�����o���c�L�̏ꍇ**********
                Case enum_DataTrendType.em_trend_Uneven
                    If Not ChangeParamsInfo(cnt).Flg_AverageFinish Then
                        
                        Dim intItemNum As Integer       '�J�e�S���C���f�b�N�X���ꎞ�ۑ�����ׂ̕ϐ�
                                                
                        '��荞�݉񐔂̏����l��511�̏ꍇ�A���������Ƃ���
                        If PALS.LoopParams.LoopCategory(cnt).Average = 511 Then
                            
                            ChangeParamsInfo(cnt).Flg_AverageFinish = True
                        
                        Else
                            
                            '�ύX���s���J�e�S���̃C���f�b�N�X���擾
                            intItemNum = PALS.CommonInfo.TestnameInfoList(.LoopCategory(cnt).TargetTestName)
                            
                            '���̎�荞�݉񐔂��v�Z
                            dblAdjValue = Int(.LoopCategory(cnt).Average * (.LoopCategory(cnt).VariationLevel _
                                            / PALS.CommonInfo.TestInfo(intItemNum).LoopJudgeLimit) ^ 2)
                            
                            '��荞�݉񐔂̔{���w��(TestCondition�V�[�g��Mode)���������ꍇ�A��肱�݉񐔂̒������s��
                            With .LoopCategory(cnt)
                                '�v�Z��̎�肱�݉񐔂�512�ȏ�̏ꍇ
                                If dblAdjValue > 511 Then
                                    'Mode��Auto�ɐݒ肳��Ă���΁A511�ɕύX
                                    If .mode = MODE_AUTO Then
                                        dblAdjValue = 511
                                    '�{���w�肪����΁A511�ɍł��߂����{���ɕύX
                                    Else
                                        dblAdjValue = 511 - (511 Mod val(.mode))
                                    End If
                                    
                                    '�A�x���[�W���������������t���O�𗧂Ă�
                                    ChangeParamsInfo(cnt).Flg_AverageFinish = True
                                
                                Else
                                    'Mode��Auto�̏ꍇ�͂��̂܂�
                                    '�{���w�肪����΁A����l�ȏ�ōł����������{���ɕύX
                                    If .mode <> MODE_AUTO Then
                                        dblAdjValue = dblAdjValue + (val(.mode) - (dblAdjValue Mod val(.mode)))
                                    End If
                                
                                    '�v�Z��̎�肱�݉񐔂�512�ȏ�̏ꍇ
                                    If dblAdjValue > 511 Then
                                        '511�ɍł��߂����{���ɕύX
                                        dblAdjValue = 511 - (511 Mod val(.mode))
                                        '�A�x���[�W���������������t���O�𗧂Ă�
                                        ChangeParamsInfo(cnt).Flg_WaitFinish = True
                                    End If
                                End If
                            End With
                            
                            'TestCondition�V�[�g���̃p�����[�^��ύX
                            If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Average", dblAdjValue) Then
                                sub_UpdataLoopParams = False
                                Exit Function
                            End If
                            
                            'Average�񐔕ύX���s�����񐔂̃C���N�������g
                            ChangeParamsInfo(cnt).AveTrialCnt = ChangeParamsInfo(cnt).AveTrialCnt + 1
                        End If
                    End If
    
                '����̌X�����㏸or���~�̏ꍇ
                Case enum_DataTrendType.em_trend_Slope
    
                '����̌X�����V�t�g�̏ꍇ
                Case enum_DataTrendType.em_trend_Shift
    
                Case Else
                    
            End Select
    
    
    
    
    
            '**********����̌X������ђl�ȊO�ŁA�ȑO�ɔ�ђl�̑Ή�(Wait�ύX)���s���Ă���ꍇ**********
            If ChangeParamsInfo(cnt).WaitTrialCnt > 0 _
                And .LoopCategory(cnt).VariationTrend <> enum_DataTrendType.em_trend_Sudden Then
                
                '����Wait�̒������������Ă���ꍇ�́A�������s��Ȃ�
                If Not ChangeParamsInfo(cnt).Flg_WaitFinish Then
                
                    'Wait�̃X�e�b�v�������߂�
''                    dblStep = Format((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 6, "#.000")
                    dblStep = Int((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 6)
    
                    '�ȑO��Wait�̕ύX��1��s���Ă���ꍇ
                    If ChangeParamsInfo(cnt).WaitTrialCnt = 1 Then
                        dblAdjValue = .LoopCategory(cnt).WAIT - (dblStep * 2)
                                            
                        'TestCondition�V�[�g���̃p�����[�^��ύX
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                    
                        'Wait�ύX���s�����񐔂̃C���N�������g
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                    
                    '�ȑO��Wait�̕ύX��2��s���Ă���ꍇ
                    ElseIf ChangeParamsInfo(cnt).WaitTrialCnt = 2 Then
                        dblAdjValue = .LoopCategory(cnt).WAIT - dblStep
                                            
                        'TestCondition�V�[�g���̃p�����[�^��ύX
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                    
                        'Wait�ύX���s�����񐔂̃C���N�������g
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                    
                    '�ȑO��Wait�̕ύX��3��ȏ�s���Ă���ꍇ
                    ElseIf ChangeParamsInfo(cnt).WaitTrialCnt >= 3 Then
                    
                    
''                        dblAdjValue = .LoopCategory(cnt).Wait - Format(dblAdjValue - (dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2))), "#.000")
''                        dblAdjValue = .LoopCategory(cnt).Wait - Int(dblAdjValue - (dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2))))
                        dblAdjValue = .LoopCategory(cnt).WAIT - Int(dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2)))
                                            
                        If dblAdjValue < ChangeParamsInfo(cnt).MinWait Then
                            dblAdjValue = ChangeParamsInfo(cnt).MinWait
                        End If
                                            
                        'TestCondition�V�[�g���̃p�����[�^��ύX
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                    
                        'Wait�ύX���s�����񐔂̃C���N�������g
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                
                        '�E�F�C�g���ő�Wait��99%�ȏ�ɂȂ����ꍇ�AWait�������I��
                        '(�؎̂Č덷������ׁA99%�ȏ�Ƃ��Ă���)
                        If dblAdjValue > (ChangeParamsInfo(cnt).MaxWait * 0.99) Then
                            If dblAdjValue > ChangeParamsInfo(cnt).MaxWait Then
                                dblAdjValue = ChangeParamsInfo(cnt).MaxWait
                            End If
                            ChangeParamsInfo(cnt).Flg_WaitFinish = True
                        End If
                    
                    Else
                
                         
                    End If
                End If
            End If
        Next cnt
    End With

Exit Function

errPALSsub_UpdataLoopParams:
    Call sub_errPALS("Updata LoopParameter error at 'sub_UpdataLoopParams'", "2-2-10-0-17")

End Function


'********************************************************************************************
' ���O: sub_Init_ChangeLoopParamsInfo
' ���e: �e�J�e�S���̃p�����[�^�ύX�󋵂�X�����i�[����\���̂̃f�[�^������
'       �ŏ��E�ő�E�F�C�g�́A�����ɐݒ���s���Ă���
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/20�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_Init_ChangeLoopParamsInfo()
    
    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALSsub_Init_ChangeLoopParamsInfo

    Dim i As Long       '���[�v�J�E���^
    
    '�J�e�S���̐������J��Ԃ�
    For i = 1 To UBound(ChangeParamsInfo)
        With ChangeParamsInfo(i)
            .MinWait = PALS.LoopParams.LoopCategory(i).WAIT                 '�ݒ�\�ȍŏ��E�F�C�g
            .MaxWait = val(frm_PALS_LoopAdj_Main.txt_maxwait)               '�ݒ�\�ȍő�E�F�C�g
            .WaitTrialCnt = 0                                               'Wait�ύX��
            .AveTrialCnt = 0                                                'Average�ύX��
            .Pre_Average = -1                                               '�O��̎�荞�݉�
            .Pre_Wait = -1                                                  '�O��̎�荞�ݑO�E�F�C�g
            .Pre_VariationTrend = enum_DataTrendType.em_trend_None          '�O���莞�̃o���c�L�X��
            .Flg_WaitFinish = False                                         'Wait���������t���O
            .Flg_AverageFinish = False                                      '��荞�݉񐔒��������t���O
        End With
    Next i
    
Exit Sub

errPALSsub_Init_ChangeLoopParamsInfo:
    Call sub_errPALS("Init ChangeLoopParamsInfo error at 'sub_Init_ChangeLoopParamsInfo'", "2-2-11-0-18")
    
End Sub


'********************************************************************************************
' ���O: sub_Update_ChangeLoopParamsInfo
' ���e: �o���c�L������A�đ�������{����ۂɁA����̊e�J�e�S����
'       Average�AWait�A�o���c�L�X����ۑ����鏈��
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/20�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Function sub_Update_ChangeLoopParamsInfo() As Boolean

    '�Ԃ�l�̏�����
    'Wait��Average�̕ύX����ł�����΁ATrue�ɕύX�����
    sub_Update_ChangeLoopParamsInfo = False

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_Update_ChangeLoopParamsInfo

    Dim i As Long       '���[�v�J�E���^
    
    '�J�e�S���̐������J��Ԃ�
    For i = 1 To UBound(ChangeParamsInfo)
        With PALS.LoopParams.LoopCategory(i)
            
            'Average�̕ύX������΁A�f�[�^�̃A�b�v�f�[�g���s���A�Ԃ�l��True�ɕύX
            If ChangeParamsInfo(i).Pre_Average <> .Average Then
                ChangeParamsInfo(i).Pre_Average = .Average
                sub_Update_ChangeLoopParamsInfo = True
            End If
            
            'Wait�̕ύX������΁A�f�[�^�̃A�b�v�f�[�g���s���A�Ԃ�l��True�ɕύX
            If ChangeParamsInfo(i).Pre_Wait <> .WAIT Then
                ChangeParamsInfo(i).Pre_Wait = .WAIT
                sub_Update_ChangeLoopParamsInfo = True
            End If
            
            '�o���c�L�X���̃A�b�v�f�[�g
            ChangeParamsInfo(i).Pre_VariationTrend = .VariationTrend
        End With
    Next i
    
Exit Function

errPALSsub_Update_ChangeLoopParamsInfo:
    Call sub_errPALS("Update ChangeLoopParameterInfo error at 'sub_Update_ChangeLoopParamsInfo'", "2-2-12-0-19")
    
End Function


'********************************************************************************************
' ���O : sub_OutPutLoopParam
' ���e : TestCondition�V�[�g�̃p�����[�^���A����f�[�^���O�̖����Ƀe�L�X�g�Œǉ�
'        ���L�̂悤�ȃf�[�^���ǉ������
'        ########### Parameter ###########
'        Category  Wait      Average
'        ML        0.1       10
'        OF        0.2       20
'        LL        0.4       40
'        SMR       0.7       70
'        DK        1         100
'        #################################
' ���� : �Ȃ�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   M.Imamura
'********************************************************************************************
Public Sub sub_OutPutLoopParam(ByRef MeasureDatalogInfo As DatalogInfo)

'    If g_ErrorFlg_PALS Then
'        Exit Sub
'    End If

On Error GoTo errPALSsub_OutPutLoopParam

    Dim intFileNo As Integer                '�t�@�C���ԍ�
    Dim intCategoryNum As Long              '�J�e�S�������񂷃��[�v�J�E���^
    
    intFileNo = FreeFile                    '�t�@�C���ԍ��̎擾
    
    With MeasureDatalogInfo
        .MeasureDate = Year(Date) & "/" & Month(Date) & "/" & Day(Date)
        .JobName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 4)
        .SwNode = Sw_Node
    End With
    
    'TestCondition�V�[�g�̃p�����[�^���A�f�[�^���O�ɒǋL
    'Append(�ǋL)���[�h�ő���f�[�^���O���J���A�e�p�����[�^��ǋL
    Open g_strOutputDataText For Append As #intFileNo

        With MeasureDatalogInfo
            Print #intFileNo, ""
            Print #intFileNo, "MEASURE DATE : " & .MeasureDate
            Print #intFileNo, "JOB NAME     : " & .JobName
            Print #intFileNo, "SW_NODE      : " & .SwNode
        End With
        
        Print #intFileNo, "########### Parameter ###########"
        Print #intFileNo, "Category" & Space(10 - Len("Category")) & "Wait" & Space(10 - Len("Wait")) & "Average"

        '�J�e�S�����J��Ԃ�
        For intCategoryNum = 1 To PALS.LoopParams.CategoryCount
            With PALS.LoopParams.LoopCategory(intCategoryNum)
                Print #intFileNo, .category & Space(10 - Len(.category)) & .WAIT & Space(10 - Len(CStr(.WAIT))) & .Average
            End With
        Next
        
        Print #intFileNo, "#################################"
    
    '�f�[�^���O�����
    Close #intFileNo

Exit Sub

errPALSsub_OutPutLoopParam:
    Call sub_errPALS("OutPut LoopParameter error at 'sub_OutPutLoopParam'", "2-2-13-0-20")

End Sub


'#######################################################################################################################
'########           �@���[�쐬�y�f�[�^���O�A���ʁz        �@     #######################################################
'#######################################################################################################################

'********************************************************************************************
' ���O: sub_MakeLoopResultSheet
' ���e: LOOP���[�쐬�֐�
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   H.Ishibashi
'********************************************************************************************
Public Sub sub_MakeLoopResultSheet(ByVal lngMaxCnt As Long, ByRef MeasureDatalogInfo As DatalogInfo)

    ' Excel ApplicationBookSheet�I�u�W�F�N�g��`
    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlWS As Excel.Worksheet

On Error GoTo errPALSsub_MakeLoopResultSheet

    Const DATASTEP As Integer = 200                 '�f�[�^���O���L�����鉡��MAX��(��256�����̂���)
    
    Dim DataRowStep As Integer     '�f�[�^���O���L������c�̔�΂���
    Dim Data_Num As Integer     '�f�[�^��(���[�v��)
    Dim lngsite As Long
    Dim lngRowJump As Long
    Dim i As Long
    Dim j As Long
    Dim retsu As Long
    Dim gyou As Long

    Dim intSheetCheck As Integer
    Dim SheetName As Worksheet

    '�n�b�`���O�p�ϐ��@----------
    Const ROW_HAIFUN As Integer = 3
    Const ROW_0 As Integer = 4
    Const ROW_01 As Integer = 5
    Dim lngCount_haifun As Long
    Dim lngCount0 As Long
    Dim lngCount01 As Long
    Dim lngColorAqua As Long
    Dim lngColorYellow As Long
    Dim lngColorOrange As Long
    
    '�F�̃C���f�b�N�X����
    '>>>2011/06/14 M.IMAMURA RGB->VBA.RGB Mod.
    '>>>2011/10/03 M.IMAMURA ColorCode Mod.
    lngColorAqua = VBA.RGB(200, 255, 255)
    lngColorYellow = VBA.RGB(250, 250, 200)
    lngColorOrange = VBA.RGB(255, 200, 150)
'    lngColorAqua = VBA.RGB(204, 255, 255)
'    lngColorYellow = VBA.RGB(250, 250, 204)
'    lngColorOrange = VBA.RGB(255, 204, 153)
    '----------------------------
    '<<<2011/10/03 M.IMAMURA ColorCode Mod.
    '<<<2011/06/14 M.IMAMURA RGB->VBA.RGB Mod.

    
    Set xlApp = CreateObject("Excel.Application")       ' Excel Application Object �����B�{�v���O�����̐eExcel�ł͂Ȃ��A
                                                        ' �V�KExcel���N�����鎖�ɒ���
    xlApp.DisplayAlerts = False                         ' �x�����b�Z�[�W��False�ݒ�B�㏑�����܂����ȂǕ����Ă��Ȃ�
                                                        ' �����ŏ㏑���ۑ�����ꍇ���ɕ֗�
    xlApp.Visible = False                               ' Book��\���ɂ���BFalse�Ŕ�\��(���œ���)�B�����������݂Ȃǂ���ꍇ�A
                                                        ' ��\���ɂ��鎖��User�̌둀���h�����Ƃ��o����B�����Sample�Ȃ̂ŕ\���B
    Set xlWB = xlApp.Workbooks.Add                      ' Excel�ɐV�KBook��ǉ��B.Open(FileName)���\�b�h�Ŋ�����ExcelBook���J�����Ƃ��\�B
    xlApp.ScreenUpdating = False

    Data_Num = lngMaxCnt
    
    '>>>2011/10/3 M.IMAMURA Add. Darts�ƃf�[�^�̊Ԋu�����킹��
    DataRowStep = Data_Num + 3 + 3
    '<<<2011/10/3 M.IMAMURA Add. Darts�ƃf�[�^�̊Ԋu�����킹��

    For lngsite = nSite To 0 Step -1 '------------------------------------------------------------------ Site_Loop

        '====================================================================================
        '========================= �f�[�^���O�V�[�g�������ݏ����J�n =========================
        '====================================================================================
        '�` �V�[�g�ǉ� (�V�[�g��:Data Log_Site0�AData Log_Site1�A���)�@�`

'>>>2011/06/02 K.SUMIYASHIKI UPDATE
'�V�[�g�ǉ��������֐���
'''        intSheetCheck = 0
'''        '�����V�[�g�������݂���΃t���O�𗧂Ă�
'''        For Each SheetName In Worksheets
'''            If SheetName.Name = "Data Log_Site" & lngsite Then
'''                intSheetCheck = 1
'''            End If
'''        Next
'''
'''        '���ɃV�[�g������ꍇ�A�Ԃ�U�����V�[�g����t���邽�߃t���O���C���N�������g
'''        For Each SheetName In Worksheets
'''            If SheetName.Name = "Data Log_Site" & lngsite & "(" & intSheetCheck & ")" Then
'''                intSheetCheck = intSheetCheck + 1
'''            End If
'''        Next
'''        '�V�[�g���ύX (�W���FData Log_Site0�A�����V�[�g���ݎ��FData Log_Site0(1)�AData Log_Site0(2)�A�A�A��)
'''        If intSheetCheck = 0 Then
'''           Sheets.Add.Name = "Data Log_Site" & lngsite
'''        Else
'''           Sheets.Add.Name = "Data Log_Site" & lngsite & "(" & intSheetCheck & ")"
'''        End If
        If lngsite = nSite - 1 Then
            xlWB.Worksheets("Sheet1").Delete
            xlWB.Worksheets("Sheet2").Delete
            xlWB.Worksheets("Sheet3").Delete
        End If
        Set xlWS = xlWB.Worksheets.Add       ' �V�K�u�b�N��Sheet1��xlWS�I�u�W�F�N�g���Z�b�g�B
        xlWS.Name = "Data Log_Site" & CStr(lngsite)

'        Call sub_AddSheet("Data Log_Site", lngsite)
'<<<2011/06/02 K.SUMIYASHIKI UPDATE
        
        'IG-XL�������Ă��Ȃ�PC�Ŏ��s����ۂ́A�o���f�[�V�������s��Ȃ�
        If Not frm_PALS_LoopAdj_Main.chk_IGXL_Check Then
            Call sub_Validate
        End If
        
        '�` �f�[�^��200���ږ��ɍs��ύX���ď����o���@�`
        lngRowJump = 0
        retsu = 0
        For i = 0 To PALS.CommonInfo.TestCount - 1  'koumoku_num=���ڐ�
        
            If PALS.CommonInfo.TestInfo(i).Label = LABEL_GRADE Then
                Exit For
            End If
        
        
            If i > 0 Then
                If i Mod DATASTEP = 0 Then
                    lngRowJump = lngRowJump + 1 '200���ڂɒB������s��ς���J�E���g�A�b�v
                    retsu = 0                   '���0�ɖ߂�
                End If
            End If
            retsu = retsu + 1
            gyou = 1 + (DataRowStep * lngRowJump) 'DataRowStep 103

            '>>>2011/10/3 M.IMAMURA Add. Darts�ƃf�[�^�̊Ԋu�����킹��
            If retsu = 1 And gyou > 1 Then
                xlWS.Cells(gyou - 1, retsu).Value = "----------" '���ږ��o��
            End If
            '<<<2011/10/3 M.IMAMURA Add. Darts�ƃf�[�^�̊Ԋu�����킹��
            
            xlWS.Cells(gyou, retsu).Value = PALS.CommonInfo.TestInfo(i).tname  '���ږ��o��


'>>>2011/04/20 K.SUMIYASHIKI UPDATE
            With PALS.CommonInfo.TestInfo(i)
                '�f�[�^���O�V�[�g�ɒP�ʂ����
                xlWS.Cells(gyou + 1, retsu).Value = "[" & .Unit & "]"
                
                '�����K�i�����
                If .arg2 = 0 Or .arg2 = 2 Then
                    xlWS.Cells(gyou + 2, retsu).Value = "No_Limit"
                Else
                    xlWS.Cells(gyou + 2, retsu).Value = sub_ReverseConvertUnit(.LowLimit, i)
                End If
                
                '����K�i�����
                If .arg2 = 0 Or .arg2 = 1 Then
                    xlWS.Cells(gyou + 3, retsu).Value = "No_Limit"
                Else
                    xlWS.Cells(gyou + 3, retsu).Value = sub_ReverseConvertUnit(.HighLimit, i)
                End If
                gyou = gyou + 4
            End With
'<<<2011/04/20 K.SUMIYASHIKI UPDATE


            For j = 1 To lngMaxCnt '[����񐔕����[�v]
'>>>2010/12/13 K.SUMIYASHIKI UPDATE
'>>>2011/05/13 K.SUMIYASHIKI UPDATE
'old 101213               Cells(gyou + 1 + j, retsu).value = PALS.CommonInfo.TestInfo(i).Site(lngSite).Data(j)
'old 110513               Cells(gyou + j, retsu).value = PALS.CommonInfo.TestInfo(i).Site(lngSite).Data(j)
                '>>>2011/10/3 M.IMAMURA Add.
                If j = 1 Then xlWS.Cells(gyou, retsu).Value = "0"
                '<<<2011/10/3 M.IMAMURA Add.
                If PALS.CommonInfo.TestInfo(i).site(lngsite).Enable(j) = True Then
                    xlWS.Cells(gyou + j, retsu).Value = sub_ReverseConvertUnit(PALS.CommonInfo.TestInfo(i).site(lngsite).Data(j), i)
                End If
'<<<2011/05/13 K.SUMIYASHIKI UPDATE
'<<<2010/12/13 K.SUMIYASHIKI UPDATE
            Next j
        Next i
        
        '>>>2011/10/3 M.IMAMURA Add.
        xlWS.Cells(gyou + lngMaxCnt * 2, 1).Value = "LoopTimes"
        xlWS.Cells(gyou + lngMaxCnt * 2, 2).Value = lngMaxCnt
        '<<<2011/10/3 M.IMAMURA Add.
        
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Making LoopResultSheet..." & Int((nSite - lngsite + 1) / (nSite + 1) * 50) & "%")
        
    Next lngsite
        
        
        '====================================================================================
        '========================= �f�[�^���O�V�[�g�������ݏ����I�� =========================
        '====================================================================================

    For lngsite = nSite To 0 Step -1 '------------------------------------------------------------------ Site_Loop
        '>>>2011/10/3 M.Imamura Add.
        '�F�t�Z���̌��J�E���g�ϐ���������
        lngCount0 = 0
        lngCount_haifun = 0
        lngCount01 = 0
        '<<<2011/10/3 M.Imamura Add.

        '//////////////////////////////////////////////////////////////////////////////
        '////////////////////// ���[�v���ʃV�[�g�������ݏ����J�n //////////////////////
        '//////////////////////////////////////////////////////////////////////////////
        '�` �V�[�g�ǉ� (��{��TestResult_Site0�ATestResult_Site1�A����)�@�`
'>>>2011/06/02 K.SUMIYASHIKI UPDATE
'�V�[�g�ǉ��������֐���
'''        intSheetCheck = 0
'''        '''�����V�[�g�������݂���΃t���O�𗧂Ă�
'''        For Each SheetName In Worksheets
'''            If SheetName.Name = "TestResult_Site" & lngsite Then  '�����V�[�g�������݂���΃t���O�𗧂Ă�
'''                intSheetCheck = 1
'''            End If
'''        Next
'''        '''���łɃV�[�g������ꍇ�� (TestResult_Site0(1)�ATestResult_Site0(2)�A���A�Ԃ�U��)
'''        For Each SheetName In Worksheets
'''            If SheetName.Name = "TestResult_Site" & lngsite & "(" & intSheetCheck & ")" Then
'''                intSheetCheck = intSheetCheck + 1
'''            End If
'''        Next
'''        '''�V�[�g���ύX
'''        If intSheetCheck = 0 Then
'''           Sheets.Add.Name = "TestResult_Site" & lngsite
'''        Else
'''           Sheets.Add.Name = "TestResult_Site" & lngsite & "(" & intSheetCheck & ")"
'''        End If

        Set xlWS = xlWB.Worksheets.Add       ' �V�K�u�b�N��Sheet1��xlWS�I�u�W�F�N�g���Z�b�g�B
        xlWS.Name = "TestResult_Site" & CStr(lngsite)
'        Call sub_AddSheet("TestResult_Site", lngsite)
'<<<2011/06/02 K.SUMIYASHIKI UPDATE

        'IG-XL�������Ă��Ȃ�PC�Ŏ��s����ۂ́A�o���f�[�V�������s��Ȃ�
        If Not frm_PALS_LoopAdj_Main.chk_IGXL_Check Then
            Call sub_Validate
        End If
        
        '�` ���x���������o���@�`
        xlWS.Cells(ROW_LABEL, CLM_NO).Value = "No"
        xlWS.Cells(ROW_LABEL, CLM_TEST).Value = "Test"
        xlWS.Cells(ROW_LABEL, CLM_UNIT).Value = "Unit"
        xlWS.Cells(ROW_LABEL, CLM_CNT).Value = "Cnt"
        xlWS.Cells(ROW_LABEL, CLM_MIN).Value = "MIN"
        xlWS.Cells(ROW_LABEL, CLM_AVG).Value = "AVG"
        xlWS.Cells(ROW_LABEL, CLM_MAX).Value = "MAX"
        xlWS.Cells(ROW_LABEL, CLM_SIGMA).Value = "sigma"
        xlWS.Cells(ROW_LABEL, CLM_3SIGMA).Value = "3sigma"
        xlWS.Cells(ROW_LABEL, CLM_1PAR10).Value = "'l/10"
        xlWS.Cells(ROW_LABEL, CLM_LOW).Value = "Low"
        xlWS.Cells(ROW_LABEL, CLM_HIGH).Value = "High"
        xlWS.Cells(ROW_LABEL, CLM_3SIGMAPARSPEC).Value = "3sigma/spec width"
'>>>2010/12/13 K.SUMIYASHIKI ADD
        xlWS.Cells(ROW_LABEL, CLM_JUDGELIMIT).Value = "LoopJudgeLimit"
'<<<2010/12/13 K.SUMIYASHIKI ADD

        '�` �e���ږ��̃f�[�^�������o���@�`
        For i = 0 To PALS.CommonInfo.TestCount - 1

            If PALS.CommonInfo.TestInfo(i).Label = LABEL_GRADE Then
                Exit For
            End If


            With PALS.CommonInfo.TestInfo(i)
                xlWS.Cells(ROW_DATASTART + i, CLM_NO).Value = i + 1
                xlWS.Cells(ROW_DATASTART + i, CLM_TEST).Value = .tname
                xlWS.Cells(ROW_DATASTART + i, CLM_UNIT).Value = "[" & .Unit & "]"
                
'>>>2010/12/13 K.SUMIYASHIKI ADD
                If (.LoopJudgeLimit <> 0.1 And .LoopJudgeLimit <> 0) Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_JUDGELIMIT).Value = .LoopJudgeLimit
                End If
'<<<2010/12/13 K.SUMIYASHIKI ADD
                
                With .site(lngsite)
'>>>2011/05/13 K.SUMIYASHIKI ADD
'old101213                    Cells(ROW_DATASTART + i, CLM_CNT).value = lngMaxCnt
                    xlWS.Cells(ROW_DATASTART + i, CLM_CNT).Value = .ActiveValueCnt
'<<<2011/05/13 K.SUMIYASHIKI ADD
                    xlWS.Cells(ROW_DATASTART + i, CLM_MIN).Value = sub_ReverseConvertUnit(.Min, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_AVG).Value = sub_ReverseConvertUnit(.ave, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_MAX).Value = sub_ReverseConvertUnit(.max, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_SIGMA).Value = sub_ReverseConvertUnit(.Sigma, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMA).Value = xlWS.Cells(ROW_DATASTART + i, CLM_SIGMA).Value * 3
                End With
            


                '����K�i�Ȃ��A�����K�i����̏ꍇ�@���@�K�i��/10�������K�i�~1/10�A3��/�K�i����3��/�����K�i
                If .arg2 = 1 Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value = sub_ReverseConvertUnit(.LowLimit, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).Value = Abs(xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value / 10)
                    
                    
                    If sub_ReverseConvertUnit(.LowLimit, i) = 0 Then
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = 0
                    Else
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = Abs(xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMA).Value / xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value)
                    End If
                    xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value = "-"
    
    
                '�����K�i�Ȃ��A����K�i����̏ꍇ�@���@�K�i��/10������K�i�~1/10�A3��/�K�i����3��/����K�i
                ElseIf .arg2 = 2 Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value = sub_ReverseConvertUnit(.HighLimit, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).Value = Abs(xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value / 10)
                    
                    If sub_ReverseConvertUnit(.HighLimit, i) = 0 Then
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = 0
                    Else
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = Abs(xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMA).Value / xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value)
                    End If
                    xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value = "-"
        
    
                '�㉺�K�i�Ȃ��̏ꍇ�@���@�K�i��/10��"-"�A3��/�K�i����"-"
                ElseIf .arg2 = 0 Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).Value = "-"
                    xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).HorizontalAlignment = xlCenter
                    xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = "-"
                    xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).HorizontalAlignment = xlCenter
                    xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value = "-"
                    xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value = "-"
    
    
                '�㉺�K�i����̏ꍇ�@���@�K�i��/10��(���-����)/10�A3��/�K�i����3��/(���-����)
                Else
                    xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value = sub_ReverseConvertUnit(.LowLimit, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value = sub_ReverseConvertUnit(.HighLimit, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).Value = (xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value - xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value) / 10
                    
                    If sub_ReverseConvertUnit((.HighLimit - .LowLimit), i) = 0 Then
'>>>2011/12/12 M.IMAMURA MOD
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = "-"
                        xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).Value = "-"
'                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).value = 0
'<<<2011/12/12 M.IMAMURA MOD
                    Else
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = Abs(xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMA).Value / _
                                                                            (xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value - xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value))
                    End If
    

                End If
    
'-----------------------------------�@3�Ё��K�i��=0�A0.1�̎��̃n�b�`���O�ݒ�@------------------------------------
                '3�Ё��K�i��=0 �Ȃ�@No,Test�𔖉��F�Ńn�b�`���O����
                If xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = 0 Then
                    lngCount0 = lngCount0 + 1
                    xlWS.Range(xlWS.Cells(ROW_DATASTART + i, CLM_NO), xlWS.Cells(ROW_DATASTART + i, CLM_TEST)).Interior.color = lngColorYellow
                    
                '3�Ё��K�i��="-" �Ȃ�@No,Test�𐅐F�Ńn�b�`���O����
                ElseIf xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = "-" Then
                    lngCount_haifun = lngCount_haifun + 1
                    xlWS.Range(xlWS.Cells(ROW_DATASTART + i, CLM_NO), xlWS.Cells(ROW_DATASTART + i, CLM_TEST)).Interior.color = lngColorAqua
                    '�������A��=0�Ȃ�@No,Test�𔖉��F�Ńn�b�`���O����
                    If xlWS.Cells(ROW_DATASTART + i, CLM_SIGMA).Value = 0 Then
                        lngCount_haifun = lngCount_haifun - 1
                        lngCount0 = lngCount0 + 1
                        xlWS.Range(xlWS.Cells(ROW_DATASTART + i, CLM_NO), xlWS.Cells(ROW_DATASTART + i, CLM_TEST)).Interior.color = lngColorYellow
                    End If
                '3�Ё��K�i��>=0.1 �Ȃ�@���̍s�𔖃I�����W�F�Ńn�b�`���O����
                ElseIf val(xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value) >= 0.1 Then
                    lngCount01 = lngCount01 + 1
                    xlWS.Range(xlWS.Cells(ROW_DATASTART + i, CLM_NO), xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC)).Interior.color = lngColorOrange
                End If
                '----------------------------------------------------------------------------------------------------------------------
    
    
                If xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value = "-" Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_LOW).HorizontalAlignment = xlCenter
                End If
                If xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value = "-" Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).HorizontalAlignment = xlCenter
                End If

            End With

        Next


        '�` ������o�́@�`
        Dim strLotName As String
        strLotName = Mid$(g_strOutputDataText, InStrRev(g_strOutputDataText, "\") + 1)
        strLotName = Left$(strLotName, Len(strLotName) - 4)
'>>>2011/12/07 M.IMAMURA Mod
        If Left$(strLotName, 12) = "LoopAdjData_" Then
            strLotName = Mid$(strLotName, 13)
        End If
        
        xlWS.Cells(ROW_NAME, CLM_NO).Value = "TestResult_Site" & CStr(lngsite) & "[" & strLotName & "] (no exclusion)"
'        Cells(ROW_NAME, CLM_NO).value = "TestResult[" & strLotName & "] (no exclusion)"
'<<<2011/12/07 M.IMAMURA Mod
        xlWS.Cells(ROW_WAFER, CLM_NO).Value = "Wafer : 1"

'>>>2011/12/07 M.IMAMURA Mod
        xlWS.Cells(ROW_MACHINEJOB, CLM_NO).Value = "Device : " & PALS.CommonInfo.g_strTesterName & "  Program : " & MeasureDatalogInfo.JobName & "  (Re-calculation : " & Left$(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 4) & " )"
'        xlWS.Cells(ROW_MACHINEJOB, CLM_NO).value = "Device : " & PALS.CommonInfo.g_strTesterName & "  Program : " & MeasureDatalogInfo.JobName & "  (Re-calculation : " & MeasureDatalogInfo.JobName & " )"
'<<<2011/12/07 M.IMAMURA Mod
'        xlWS.Cells(ROW_MACHINEJOB, CLM_NO).value = "Device : SKCCDS" & MeasureDatalogInfo.SwNode & "  Program : " & MeasureDatalogInfo.JobName & "  (Re-calculation : " & MeasureDatalogInfo.JobName & " )"
        xlWS.Cells(ROW_LOOPCOUNT, CLM_NO).Value = "Measurement count : " & lngMaxCnt
               
        xlWS.Cells(ROW_DATE, CLM_NO).Value = "Measurement date : " & MeasureDatalogInfo.MeasureDate

        '�`�@3�Ё��K�i��=0�A"-"�A0.1�ȏ�̌������ꂼ��o�́@�`
        xlWS.Cells(ROW_HAIFUN, CLM_1PAR10).Value = lngCount_haifun
        xlWS.Cells(ROW_0, CLM_1PAR10).Value = lngCount0
        xlWS.Cells(ROW_01, CLM_1PAR10).Value = lngCount01

        '�`�@�F�Â��@�`
        xlWS.Cells(ROW_HAIFUN, CLM_3SIGMA).Interior.color = lngColorAqua
        xlWS.Cells(ROW_0, CLM_3SIGMA).Interior.color = lngColorYellow
        xlWS.Cells(ROW_01, CLM_3SIGMA).Interior.color = lngColorOrange

        '�` �|���A�����ݒ�@�`
        Call sub_SetLoopFormat(xlApp, xlWB, xlWS)

        '�`�@����͈͐ݒ�@�`
        Call sub_PrintSetting(xlApp, xlWB, xlWS)
        
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Making LoopResultSheet..." & 50 + Int((nSite - lngsite + 1) / (nSite + 1) * 50) & "%")

        '//////////////////////////////////////////////////////////////////////////////
        '////////////////////// ���[�v���ʃV�[�g�������ݏ����I�� //////////////////////
        '//////////////////////////////////////////////////////////////////////////////
        
    Next lngsite '------------------------------------------------------------------------------------- Site_Loop

    xlApp.ScreenUpdating = True
    
    Dim xlFileName As String
    xlFileName = Left$(g_strOutputDataText, Len(g_strOutputDataText) - 4) & ".xls"
    xlWB.SaveAs xlFileName              ' �V�K�u�b�N��ʖ��ۑ�
'    xlApp.Visible = True
    xlWB.Close              ' �V�K�u�b�N����
    xlApp.Quit              ' Excel�𗎂Ƃ�
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    ThisWorkbook.Activate

Exit Sub
    Dim xlFileName2 As String
    xlFileName2 = Left$(g_strOutputDataText, Len(g_strOutputDataText) - 4) & ".xls"
    xlWB.SaveAs xlFileName2              ' �V�K�u�b�N��ʖ��ۑ�
'    xlApp.Visible = True
    xlWB.Close              ' �V�K�u�b�N����
    xlApp.Quit              ' Excel�𗎂Ƃ�
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing

errPALSsub_MakeLoopResultSheet:
    Call sub_errPALS("Make Loop sheet error at 'sub_MakeLoopResultSheet'", "2-2-14-0-21")

End Sub


'#######################################################################################################################
'########           �@�|���A�����y���ʃV�[�g�z        �@     ###########################################################
'#######################################################################################################################
Private Sub sub_SetLoopFormat(xlApp As Excel.Application, xlWB As Excel.Workbook, xlWS As Excel.Worksheet)

On Error GoTo errPALSsub_SetLoopFormat

    '�t�H���g�T�C�Y��{9�A��ԏゾ��16
    xlWS.Cells.Select
    xlApp.Selection.Font.Size = 9
    xlWS.Range(xlWS.Cells(ROW_NAME, CLM_NO), xlWS.Cells(ROW_NAME, CLM_NO)).Select
    xlApp.Selection.Font.Size = 16

    '�t�H���g��MS �S�V�b�N
    xlWS.Cells.Font.Name = "�l�r �S�V�b�N"

    '����������MIN�`3�ЁA3��/�K�i���̗�������_��4�ʂ܂ŕ\��
    xlWS.Range(xlWS.Cells(ROW_DATASTART, CLM_MIN), xlWS.Cells(xlWS.Rows.Count, CLM_3SIGMA)).NumberFormatLocal = "0.00000;-0.00000;0;@"
    xlWS.Range(xlWS.Cells(ROW_DATASTART, CLM_3SIGMAPARSPEC), xlWS.Cells(xlWS.Rows.Count, CLM_3SIGMAPARSPEC)).NumberFormatLocal = "0.00000;-0.00000;0;@"

    '�|���A�����œK��
    xlWS.Cells(ROW_LABEL, CLM_NO).Select
    xlWS.Range(xlApp.Selection, xlApp.Selection.End(xlToRight)).Select
    xlWS.Range(xlApp.Selection, xlApp.Selection.End(xlDown)).Select
    With xlApp.Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Columns.AutoFit
    End With

    Dim i As Long
    '�f�[�^�ŏ��A���ρA�ő�Z�������I�[�g����0.5�L����(�����#####�ɂȂ�\������̂���)
    For i = CLM_MIN To CLM_MAX
        xlWS.Cells(ROW_LABEL, i).ColumnWidth = xlWS.Cells(ROW_LABEL, i).ColumnWidth + 0.5
    Next

    '�Z��������12�ɓ���
    xlWS.Cells.RowHeight = 12

    '�擪�s�̂݃Z��������24�ɂ���
    xlWS.Cells(ROW_NAME, CLM_NO).RowHeight = 24

    '���x���s�𒆉������ɐݒ�
    xlWS.Range(xlWS.Cells(ROW_LABEL, CLM_NO), xlWS.Cells(ROW_LABEL, CLM_3SIGMAPARSPEC)).HorizontalAlignment = xlCenter

Exit Sub

errPALSsub_SetLoopFormat:
    Call sub_errPALS("Set Line error at 'sub_SetLoopFormat'", "2-2-15-0-22")

End Sub
'#######################################################################################################################
'########           �@�v�����g�ݒ�y���ʃV�[�g�z        �@     #########################################################
'#######################################################################################################################
Private Sub sub_PrintSetting(xlApp As Excel.Application, xlWB As Excel.Workbook, xlWS As Excel.Worksheet)

On Error GoTo errPALSsub_PrintSetting

    Const NEWPAGE = 61

    Dim i As Long
    Dim lngRowlast As Long '���[�v���ʃV�[�g�̍ŏI�s�̍s�ԍ�

    '�f�[�^�̕���������͈͂Ɏw��
    xlWS.Range(xlWS.Cells(ROW_NAME, CLM_NO), xlWS.Cells(ROW_NAME, CLM_3SIGMAPARSPEC)).Select
    xlWS.Range(xlApp.Selection, xlApp.Selection.End(xlDown)).Select
    lngRowlast = xlApp.Selection.Rows.Count + xlApp.Selection.Row - 1 '�ŏI�s�̍s�ԍ��擾
''    ActiveSheet.PageSetup.PrintArea = ActiveCell.CurrentRegion.Address
'���C����
''    ActiveSheet.PageSetup.PrintArea = Range(Cells(ROW_NAME, CLM_NO), Cells(lngRowlast, CLM_3SIGMAPARSPEC)).Address
    
    '���s�ݒ�
'''''    For i = Page1 + 1 To lngRowlast
    For i = 1 To lngRowlast
        If i Mod NEWPAGE = 0 Then xlWS.Rows(i).PageBreak = xlPageBreakManual
    Next

    '���x���̍s��S�y�[�W�Ɉ�������悤�ݒ�
    xlWS.Range(xlWS.Cells(ROW_LABEL, CLM_NO), xlWS.Cells(ROW_LABEL, CLM_NO)).Select
''    ActiveSheet.PageSetup.PrintTitleRows = "$6:$6"
''
''    '�w�b�_�[�A�t�b�^�[�ݒ�
''    With ActiveSheet.PageSetup
''        .CenterHeader = ActiveSheet.name        '�����w�b�_�[�F�V�[�g��
''        .CenterFooter = "&P / &N" & " �y�[�W"   '�����t�b�^�[�F�y�[�W�ԍ�
''        .RightFooter = "���[�v�c�[�����["       '�E���t�b�^�[�F�c�[���̖��O�Ȃ�
''    End With
''
''    '��1�y�[�W���Ŏ��߂�
''    ActiveSheet.PageSetup.FitToPagesWide = 1

    'A1�Z����I�����ďI��
    xlWS.Range("A1").Select

Exit Sub

errPALSsub_PrintSetting:
    Call sub_errPALS("Set PrintArea error at 'sub_PrintSetting'", "2-2-16-0-23")

End Sub


'********************************************************************************************
' ���O: sub_ReverseConvertUnit
' ���e: LOOP���[�ɏo�͂���l���A�e�X�g�C���X�^���X����擾�����P�ʂɂ���ĒP�ʕϊ�����֐�
' ����: dblValue   : �ϊ��O�̓����l
'       lngTestCnt : ���ڔԍ��������l
' �ߒl: �P�ʕϊ���̒l
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_ReverseConvertUnit(ByVal dblValue As Double, ByVal lngTestCnt As Long) As Double

On Error GoTo errPALSsub_ReverseConvertUnit

    '�P�ʊ��Z
    Select Case PALS.CommonInfo.TestInfo(lngTestCnt).Unit
        Case ""
            sub_ReverseConvertUnit = dblValue
        
        Case "MA"
            sub_ReverseConvertUnit = dblValue / MEGA

        Case "MV"
            sub_ReverseConvertUnit = dblValue / MEGA

        Case "KV"
            sub_ReverseConvertUnit = dblValue / KIRO

        Case "KA"
            sub_ReverseConvertUnit = dblValue / KIRO
        
        Case "V"
            sub_ReverseConvertUnit = dblValue
        
        Case "v"
            sub_ReverseConvertUnit = dblValue
        
        Case "A"
            sub_ReverseConvertUnit = dblValue
        
        Case "a"
            sub_ReverseConvertUnit = dblValue
        
        Case "mV"
            sub_ReverseConvertUnit = dblValue / MILLI
        
        Case "mv"
            sub_ReverseConvertUnit = dblValue / MILLI
                        
        Case "mA"
            sub_ReverseConvertUnit = dblValue / MILLI
                
        Case "uV"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
        Case "uv"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
        Case "uA"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
        Case "nV"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "nv"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "nA"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "pV"
            sub_ReverseConvertUnit = dblValue / PIKO
        
        Case "pv"
            sub_ReverseConvertUnit = dblValue / PIKO
        
        Case "pA"
            sub_ReverseConvertUnit = dblValue / PIKO
        
        Case "fV"
            sub_ReverseConvertUnit = dblValue / FEMTO
        
        Case "fv"
            sub_ReverseConvertUnit = dblValue / FEMTO
        
        Case "fA"
            sub_ReverseConvertUnit = dblValue / FEMTO
        
        Case "ms"
            sub_ReverseConvertUnit = dblValue / MILLI
        
        Case "us"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
'>>>2010/12/13 K.SUMIYASHIKI ADD
        Case "ns"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "ps"
            sub_ReverseConvertUnit = dblValue / PIKO
'<<<2010/12/13 K.SUMIYASHIKI ADD
        Case "S"
            sub_ReverseConvertUnit = dblValue
                        
        Case "mS"
            sub_ReverseConvertUnit = dblValue / MILLI
        
        Case "uS"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
        Case "nS"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "pS"
            sub_ReverseConvertUnit = dblValue / PIKO
        
        Case "%"
            sub_ReverseConvertUnit = dblValue / percent
                                                
        Case "Kr"
            sub_ReverseConvertUnit = dblValue
'>>>2013/12/03 T.Morimoto ADD
        Case "W"
            sub_ReverseConvertUnit = dblValue
                        
        Case "mW"
            sub_ReverseConvertUnit = dblValue / MILLI
        
        Case "uW"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
        Case "nW"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "pW"
            sub_ReverseConvertUnit = dblValue / PIKO
'<<<2013/12/04

        Case Else

'>>>2010/12/13 K.SUMIYASHIKI MESSEGE CHANGE
            Call MsgBox("Error! Not Entry Unit" & "->" & PALS.CommonInfo.TestInfo(lngTestCnt).Unit & vbCrLf & "ErrCode.2-2-17-4-24", vbExclamation)
'<<<2010/12/13 K.SUMIYASHIKI MESSEGE CHANGE
        
    End Select

Exit Function

errPALSsub_ReverseConvertUnit:
    Call sub_errPALS("Convert Unit error at 'sub_ReverseConvertUnit'", "2-2-17-0-25")

End Function


'********************************************************************************************
' ���O: sub_CheckTestConditionWaitData
' ���e: TestCondition�ɐݒ肵�Ă���e�J�e�S����Wait���A�t�H�[���Ŏw�肵���ő�Wait�ȏ�ɂȂ��Ă����ꍇ�G���[��Ԃ�
' ����: dblMaxWait : �t�H�[���Ŏw�肵���ő�Wait
' �ߒl: True  :�ُ�l�Ȃ�
'       False :�ُ�l����
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Function sub_CheckTestConditionWaitData(ByVal dblMaxWait As Double) As Boolean

    '������
    '���Ȃ����False���Ԃ�
    sub_CheckTestConditionWaitData = True

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_CheckTestConditionWaitData

    With PALS.LoopParams
        Dim cnt As Long
        '�S�J�e�S�����J��Ԃ�
        For cnt = 1 To .CategoryCount
            '�t�H�[���Ŏw�肵���l�ȏ��Wait���ݒ肳��Ă����ꍇ�A�G���[��Ԃ�
            If .LoopCategory(cnt).WAIT > (dblMaxWait) Then
                sub_CheckTestConditionWaitData = True
                '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
                If g_RunAutoFlg_PALS = False Then
                    Call MsgBox("Error" & vbCrLf & "TestCondition no Wait ga saidaiti wo koeteimasu!" & vbCrLf & .LoopCategory(cnt).category & vbCrLf & "ErCode.2-2-18-5-26", vbExclamation)
                End If
                '<<<2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
            End If
        Next cnt
    End With

Exit Function

errPALSsub_CheckTestConditionWaitData:
    Call sub_errPALS("Check TestCondition Wait Data error at 'sub_CheckTestConditionWaitData'", "2-2-18-0-27")

End Function


'********************************************************************************************
' ���O: sub_Validate
' ���e: IG-XL�̃o���f�[�V�������s���B
'       IG-XL�ŃV�[�g��ǉ�����ۂɁA�o���f�[�V�������s��Ȃ��ƃG���[����������׎g�p���Ă���B
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_Validate()
On Error GoTo errPALSsub_Validate
    
    TheExec.Validate

Exit Sub

errPALSsub_Validate:
    Call sub_errPALS("IG-XL Validate error at 'sub_Validate'", "2-2-19-0-28")

End Sub


'********************************************************************************************
' ���O: sub_GetMeasureData
' ���e: LOOP�c�[���ɂ���č쐬���ꂽ�f�[�^���O�̉����ɂ���f�[�^���A�����œn���ꂽ1�s���̃f�[�^���O����
'       �ǂݎ��ׂ̊֐��
' ����: strBuf     :�f�[�^�������Ă���1�s���̃f�[�^���O
'     : strGetType :��������f�[�^�̎��
' �ߒl: 1�s���̃f�[�^���O���甲���o�����e��ނ̃f�[�^
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/09/30�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Function sub_GetMeasureData(ByVal strbuf As String, ByVal strGetType As String) As String

On Error GoTo errPALSsub_GetMeasureData

    Select Case strGetType

        Case "Date"
            sub_GetMeasureData = Mid$(strbuf, 16)
        
        Case "JobName"
            sub_GetMeasureData = Mid$(strbuf, 16)

        Case "Node"
            sub_GetMeasureData = Mid$(strbuf, 16)

    End Select

Exit Function

errPALSsub_GetMeasureData:
    Call sub_errPALS("Get Measure Data Error at 'sub_GetMeasureData'", "2-2-20-0-29")

End Function

'>>>2010/12/13 K.SUMIYASHIKI ADD FUNCTION
'********************************************************************************************
' ���O: sub_JudgeBaratuki
' ���e: �X������ђl�Ɣ��f�����悤�ȑ傫�ȃo���c�L�����f����֐�
'       ���̑���f�[�^�Ƃ̍���(��Βl)��ώZ���A�����̕��ϒl��1�Јȏ゠��΁A
'       �傫�ȃo���c�L������Ɣ��f����B
' ����: lngTestNo         : ���ڂ������ԍ�
'       sitez             : �T�C�g�ԍ�
'       lngNowLoopCnt     : ����ς݉�
' �ߒl: �o���c�L�X���������񋓑�
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2010/12/13�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_JudgeBaratuki(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As enum_DataTrendType

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_JudgeBaratuki

'��ђl�Ɣ��f�����\���̂���A�傫�ȃo���c�L���Ȃ����̔��f
'->������̓����l�Ƃ̍����̕��ς����߁A�l��1�Јȏ�̏ꍇ�A�傫�ȃo���c�L������Ɣ��f����

    Dim dblSumDelta As Double   '���̃f�[�^�Ƃ̍�����ώZ����ׂ̕ϐ�
    Dim data_cnt As Long        '���[�v�J�E���^(�f�[�^�C���f�b�N�X������)
    
    dblSumDelta = 0
    
    With PALS.CommonInfo.TestInfo(lngTestNo).site(sitez)
    
        '�f�[�^���}�C�i�X1�����J��Ԃ�
        For data_cnt = 1 To lngNowLoopCnt - 1
            '���̃f�[�^�Ƃ̍����̐�Βl���擾
            dblSumDelta = dblSumDelta + Abs(.Data(data_cnt + 1) - .Data(data_cnt))
        Next data_cnt
    
        '���̃f�[�^�Ƃ̍������ς�0.9�Јȏ�̏ꍇ�o���c�L�Ƃ���
        If (dblSumDelta / (lngNowLoopCnt - 1)) > (.Sigma * 0.9) Then
            Debug.Print ("Big Baratuki!")
            Debug.Print ("TestName : " & PALS.CommonInfo.TestInfo(lngTestNo).tname)
            Debug.Print ("Site     : " & sitez) & vbCrLf            '�o���c�L�������l��Ԃ��I��
            sub_JudgeBaratuki = em_trend_Uneven
        End If
    End With
Exit Function

errPALSsub_JudgeBaratuki:
    Call sub_errPALS("Check LoopData error at 'sub_JudgeBaratuki'", "2-2-21-0-30")

End Function



Public Function sub_CheckTestInstancesParams() As Boolean

    sub_CheckTestInstancesParams = True

    Dim CategoryCnt As Long
    Dim TestItemCnt As Long
    Dim strTmpCategoryName1 As String
    Dim strTmpCategoryName2 As String

    Dim blnCategory1_OK As Boolean
    Dim blnCategory2_OK As Boolean

    With PALS
        For TestItemCnt = 0 To .CommonInfo.TestCount
            strTmpCategoryName1 = .CommonInfo.TestInfo(TestItemCnt).CapCategory1
            strTmpCategoryName2 = .CommonInfo.TestInfo(TestItemCnt).CapCategory2
            
            If strTmpCategoryName1 = "" Or strTmpCategoryName1 = "DC" Then
                blnCategory1_OK = True
            Else
                blnCategory1_OK = False
            End If
            
            If strTmpCategoryName2 = "" Or strTmpCategoryName1 = "DC" Then
                blnCategory2_OK = True
            Else
                blnCategory2_OK = False
            End If
            
            If blnCategory1_OK = False Or blnCategory2_OK = False Then
                
                For CategoryCnt = 1 To .LoopParams.CategoryCount
                    If blnCategory1_OK = False And strTmpCategoryName1 = .LoopParams.LoopCategory(CategoryCnt).category Then
                        blnCategory1_OK = True
                    ElseIf blnCategory2_OK = False And strTmpCategoryName2 = .LoopParams.LoopCategory(CategoryCnt).category Then
                        blnCategory2_OK = True
                    End If
                Next CategoryCnt
            
                If blnCategory1_OK = False And blnCategory2_OK = False Then
                    '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
                    If g_RunAutoFlg_PALS = False Then
                        Call MsgBox(.CommonInfo.TestInfo(TestItemCnt).tname & " CapCategory1 and 2 is wrong!!" & vbCrLf & "Please Chack CapCategory." & vbCrLf & "ErrCode.2-2-22-4-31", vbExclamation)
                    End If
                    sub_CheckTestInstancesParams = False
                    Exit Function
                ElseIf blnCategory1_OK = False Then
                    If g_RunAutoFlg_PALS = False Then
                        Call MsgBox(.CommonInfo.TestInfo(TestItemCnt).tname & " CapCategory1 is wrong!!" & vbCrLf & "Please Chack CapCategory." & vbCrLf & "ErrCode.2-2-22-4-32", vbExclamation)
                    End If
                    sub_CheckTestInstancesParams = False
                    Exit Function
                ElseIf blnCategory2_OK = False Then
                    If g_RunAutoFlg_PALS = False Then
                        Call MsgBox(.CommonInfo.TestInfo(TestItemCnt).tname & " CapCategory2 is wrong!!" & vbCrLf & "Please Chack CapCategory." & vbCrLf & "ErrCode.2-2-22-4-33", vbExclamation)
                    End If
                    sub_CheckTestInstancesParams = False
                    Exit Function
                End If
            End If
        Next TestItemCnt
    End With
    
End Function


Public Sub sub_LoopParamsCheck()

    Dim cnt As Long
    
    With PALS.LoopParams
        For cnt = 1 To .CategoryCount
            If .LoopCategory(cnt).Average = 511 Then
                '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
                If g_RunAutoFlg_PALS = False Then
                    Call MsgBox(.LoopCategory(cnt).category & "Average count is 511..." & vbCrLf & "Please check average!" & vbCrLf & "ErrCode.2-2-23-5-34", vbExclamation)
                End If
'            ElseIf .LoopCategory(cnt).Wait > ChangeParamsInfo(cnt).MaxWait * 0.99 Then
            ElseIf .LoopCategory(cnt).WAIT > val(frm_PALS_LoopAdj_Main.txt_maxwait) * 0.99 Then
                If g_RunAutoFlg_PALS = False Then
                    Call MsgBox(.LoopCategory(cnt).category & "Wait is max..." & vbCrLf & "Please check average!" & vbCrLf & "ErrCode.2-2-23-5-35", vbExclamation)
                End If
            End If
        Next cnt
    End With

End Sub


'>>>2011/06/20 K.SUMIYASHIKI ADD FUNCTION
'********************************************************************************************
' ���O: sub_AddSheet
' ���e: �w�肵�����[�N�V�[�g���̃V�[�g��ǉ�����B�����V�[�g������΁A�������C���N�������g���ǉ�����
' ����: strSheetName   : ���[�N�V�[�g��
'       sitez          : �T�C�g
' �ߒl: �Ȃ�
' ���l�F�Ȃ�
' �X�V�����F Rev1.0      2011/06/20�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_AddSheet(ByVal strSheetName As String, ByVal sitez As Integer)

    Sheets.Add.Name = "TempAddSheet"
    
    Dim intSheetCheck As Long
    intSheetCheck = 0 '�}�ԏ����l
    On Error Resume Next
    Do
        Err.Clear
        If intSheetCheck = 0 Then
            ActiveSheet.Name = strSheetName & sitez
        Else
            ActiveSheet.Name = strSheetName & sitez & "(" & intSheetCheck & ")"
        End If
        intSheetCheck = intSheetCheck + 1
    Loop Until Err.Number = 0
    On Error GoTo 0

End Sub


'>>>2011/06/20 K.SUMIYASHIKI ADD FUNCTION
'********************************************************************************************
' ���O: sub_Get_F_Value
' ���e: �w�肵������񐔁A�L�Ӑ����A���/�����ł�F�l�f�[�^���e�[�u������擾����
' ����: MeasureCnt      : �����
'       SigmaNum        : �L�Ӑ���
'       TopOrBottom     : ���or����("top"or"bottom"�Ŏw��)
' �ߒl: F�l�f�[�^
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2011/06/20�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_Get_F_Value(ByVal MeasureCnt As Integer, ByVal SigmaNum As String, ByVal TopOrBottom As String) As Double

    Dim F_Table() As Double     'F�l�e�[�u������擾�����f�[�^�z����i�[
    
    'F�l�e�[�u�����z��f�[�^���擾
    F_Table = sub_Get_F_Table(SigmaNum, TopOrBottom)
    
    '����񐔂�100��ȏ�̏ꍇ�́A100�񎞂�F�l�ŋߎ�
    If MeasureCnt > 100 Then
        sub_Get_F_Value = F_Table(100)
    Else
        '�w�葪��񐔎���F�l��Ԃ�
        sub_Get_F_Value = F_Table(MeasureCnt)
    End If


End Function


'>>>2011/06/20 K.SUMIYASHIKI ADD FUNCTION
'********************************************************************************************
' ���O: sub_Get_F_Table
' ���e: �w�肵������񐔁A�L�Ӑ����A���/�����ł�F�l�f�[�^���e�[�u������擾����
' ����: SigmaNum        : �L�Ӑ���
'       TopOrBottom     : ���or����("top"or"bottom"�Ŏw��)
' �ߒl: F�l�f�[�^�z��
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2011/06/20�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_Get_F_Table(ByVal SigmaNum As String, ByVal TopOrBottom As String) As Double()

    Dim DataTable(100) As Double        'F�l�f�[�^���i�[����z��

    '�L�Ӑ���3�Ђ̏ꍇ
    If SigmaNum = 3 Then
    
        '�����f�[�^
        If TopOrBottom = "bottom" Then
            DataTable(3) = 0.001503
            DataTable(4) = 0.008852
            DataTable(5) = 0.021938
            DataTable(6) = 0.038185
            DataTable(7) = 0.055678
            DataTable(8) = 0.07335
            DataTable(9) = 0.090656
            DataTable(10) = 0.107334
            DataTable(11) = 0.123277
            DataTable(12) = 0.138451
            DataTable(13) = 0.152869
            DataTable(14) = 0.166561
            DataTable(15) = 0.179568
            DataTable(16) = 0.191932
            DataTable(17) = 0.203697
            DataTable(18) = 0.214907
            DataTable(19) = 0.225599
            DataTable(20) = 0.235812
            DataTable(21) = 0.245578
            DataTable(22) = 0.25493
            DataTable(23) = 0.263895
            DataTable(24) = 0.2725
            DataTable(25) = 0.280768
            DataTable(26) = 0.288721
            DataTable(27) = 0.296379
            DataTable(28) = 0.30376
            DataTable(29) = 0.310882
            DataTable(30) = 0.317758
            DataTable(31) = 0.324405
            DataTable(32) = 0.330833
            DataTable(33) = 0.337057
            DataTable(34) = 0.343085
            DataTable(35) = 0.34893
            DataTable(36) = 0.354601
            DataTable(37) = 0.360105
            DataTable(38) = 0.365453
            DataTable(39) = 0.37065
            DataTable(40) = 0.375705
            DataTable(41) = 0.380624
            DataTable(42) = 0.385413
            DataTable(43) = 0.390079
            DataTable(44) = 0.394626
            DataTable(45) = 0.399059
            DataTable(46) = 0.403385
            DataTable(47) = 0.407606
            DataTable(48) = 0.411728
            DataTable(49) = 0.415754
            DataTable(50) = 0.419688
            DataTable(51) = 0.423534
            DataTable(52) = 0.427295
            DataTable(53) = 0.430975
            DataTable(54) = 0.434576
            DataTable(55) = 0.438101
            DataTable(56) = 0.441553
            DataTable(57) = 0.444935
            DataTable(58) = 0.448249
            DataTable(59) = 0.451498
            DataTable(60) = 0.454682
            DataTable(61) = 0.457806
            DataTable(62) = 0.460871
            DataTable(63) = 0.463878
            DataTable(64) = 0.466829
            DataTable(65) = 0.469727
            DataTable(66) = 0.472573
            DataTable(67) = 0.475369
            DataTable(68) = 0.478115
            DataTable(69) = 0.480814
            DataTable(70) = 0.483467
            DataTable(71) = 0.486076
            DataTable(72) = 0.488641
            DataTable(73) = 0.491163
            DataTable(74) = 0.493645
            DataTable(75) = 0.496087
            DataTable(76) = 0.498491
            DataTable(77) = 0.500856
            DataTable(78) = 0.503185
            DataTable(79) = 0.505478
            DataTable(80) = 0.507737
            DataTable(81) = 0.509961
            DataTable(82) = 0.512153
            DataTable(83) = 0.514313
            DataTable(84) = 0.516441
            DataTable(85) = 0.518538
            DataTable(86) = 0.520606
            DataTable(87) = 0.522645
            DataTable(88) = 0.524655
            DataTable(89) = 0.526638
            DataTable(90) = 0.528593
            DataTable(91) = 0.530522
            DataTable(92) = 0.532426
            DataTable(93) = 0.534304
            DataTable(94) = 0.536157
            DataTable(95) = 0.537987
            DataTable(96) = 0.539793
            DataTable(97) = 0.541575
            DataTable(98) = 0.543336
            DataTable(99) = 0.545074
            DataTable(100) = 0.546791


        '����f�[�^
        ElseIf TopOrBottom = "top" Then
            DataTable(3) = 222221.722183
            DataTable(4) = 665.833264
            DataTable(5) = 104.378009
            DataTable(6) = 42.000687
            DataTable(7) = 24.318542
            DataTable(8) = 16.822538
            DataTable(9) = 12.868018
            DataTable(10) = 10.479585
            DataTable(11) = 8.899609
            DataTable(12) = 7.784354
            DataTable(13) = 6.958128
            DataTable(14) = 6.322781
            DataTable(15) = 5.819582
            DataTable(16) = 5.411412
            DataTable(17) = 5.073739
            DataTable(18) = 4.789744
            DataTable(19) = 4.547531
            DataTable(20) = 4.338458
            DataTable(21) = 4.156104
            DataTable(22) = 3.995602
            DataTable(23) = 3.853195
            DataTable(24) = 3.725944
            DataTable(25) = 3.611509
            DataTable(26) = 3.508013
            DataTable(27) = 3.413927
            DataTable(28) = 3.327995
            DataTable(29) = 3.249176
            DataTable(30) = 3.1766
            DataTable(31) = 3.109535
            DataTable(32) = 3.047358
            DataTable(33) = 2.989539
            DataTable(34) = 2.93562
            DataTable(35) = 2.885207
            DataTable(36) = 2.837959
            DataTable(37) = 2.793575
            DataTable(38) = 2.751795
            DataTable(39) = 2.712388
            DataTable(40) = 2.675149
            DataTable(41) = 2.639898
            DataTable(42) = 2.606473
            DataTable(43) = 2.574731
            DataTable(44) = 2.544542
            DataTable(45) = 2.51579
            DataTable(46) = 2.488372
            DataTable(47) = 2.462192
            DataTable(48) = 2.437164
            DataTable(49) = 2.413212
            DataTable(50) = 2.390263
            DataTable(51) = 2.368254
            DataTable(52) = 2.347124
            DataTable(53) = 2.326821
            DataTable(54) = 2.307293
            DataTable(55) = 2.288496
            DataTable(56) = 2.270387
            DataTable(57) = 2.252927
            DataTable(58) = 2.23608
            DataTable(59) = 2.219812
            DataTable(60) = 2.204094
            DataTable(61) = 2.188895
            DataTable(62) = 2.17419
            DataTable(63) = 2.159953
            DataTable(64) = 2.146162
            DataTable(65) = 2.132794
            DataTable(66) = 2.119829
            DataTable(67) = 2.107249
            DataTable(68) = 2.095035
            DataTable(69) = 2.083171
            DataTable(70) = 2.071642
            DataTable(71) = 2.060431
            DataTable(72) = 2.049527
            DataTable(73) = 2.038915
            DataTable(74) = 2.028583
            DataTable(75) = 2.01852
            DataTable(76) = 2.008715
            DataTable(77) = 1.999157
            DataTable(78) = 1.989837
            DataTable(79) = 1.980746
            DataTable(80) = 1.971873
            DataTable(81) = 1.963212
            DataTable(82) = 1.954755
            DataTable(83) = 1.946493
            DataTable(84) = 1.93842
            DataTable(85) = 1.930528
            DataTable(86) = 1.922813
            DataTable(87) = 1.915266
            DataTable(88) = 1.907883
            DataTable(89) = 1.900658
            DataTable(90) = 1.893585
            DataTable(91) = 1.88666
            DataTable(92) = 1.879877
            DataTable(93) = 1.873232
            DataTable(94) = 1.866721
            DataTable(95) = 1.860339
            DataTable(96) = 1.854083
            DataTable(97) = 1.847947
            DataTable(98) = 1.84193
            DataTable(99) = 1.836026
            DataTable(100) = 1.830233
        
        Else
            MsgBox ("Program Argument Error!!")
                
        End If

    '�L�Ӑ���2�Ђ̏ꍇ
    ElseIf SigmaNum = 2 Then
        
        '�����f�[�^
        If TopOrBottom = "bottom" Then
            DataTable(3) = 0.02597
            DataTable(4) = 0.062328
            DataTable(5) = 0.100208
            DataTable(6) = 0.135357
            DataTable(7) = 0.167013
            DataTable(8) = 0.195366
            DataTable(9) = 0.220821
            DataTable(10) = 0.243786
            DataTable(11) = 0.264623
            DataTable(12) = 0.283634
            DataTable(13) = 0.30107
            DataTable(14) = 0.317141
            DataTable(15) = 0.332017
            DataTable(16) = 0.345844
            DataTable(17) = 0.358742
            DataTable(18) = 0.370814
            DataTable(19) = 0.382148
            DataTable(20) = 0.392818
            DataTable(21) = 0.402889
            DataTable(22) = 0.412416
            DataTable(23) = 0.421449
            DataTable(24) = 0.430031
            DataTable(25) = 0.438199
            DataTable(26) = 0.445986
            DataTable(27) = 0.453423
            DataTable(28) = 0.460536
            DataTable(29) = 0.467348
            DataTable(30) = 0.473882
            DataTable(31) = 0.480155
            DataTable(32) = 0.486186
            DataTable(33) = 0.491991
            DataTable(34) = 0.497583
            DataTable(35) = 0.502976
            DataTable(36) = 0.508182
            DataTable(37) = 0.513211
            DataTable(38) = 0.518074
            DataTable(39) = 0.522781
            DataTable(40) = 0.527339
            DataTable(41) = 0.531757
            DataTable(42) = 0.536041
            DataTable(43) = 0.540199
            DataTable(44) = 0.544237
            DataTable(45) = 0.548161
            DataTable(46) = 0.551977
            DataTable(47) = 0.555688
            DataTable(48) = 0.559301
            DataTable(49) = 0.562819
            DataTable(50) = 0.566247
            DataTable(51) = 0.569589
            DataTable(52) = 0.572848
            DataTable(53) = 0.576028
            DataTable(54) = 0.579132
            DataTable(55) = 0.582163
            DataTable(56) = 0.585124
            DataTable(57) = 0.588018
            DataTable(58) = 0.590847
            DataTable(59) = 0.593614
            DataTable(60) = 0.596321
            DataTable(61) = 0.598971
            DataTable(62) = 0.601564
            DataTable(63) = 0.604105
            DataTable(64) = 0.606593
            DataTable(65) = 0.609032
            DataTable(66) = 0.611422
            DataTable(67) = 0.613765
            DataTable(68) = 0.616064
            DataTable(69) = 0.618319
            DataTable(70) = 0.620531
            DataTable(71) = 0.622703
            DataTable(72) = 0.624835
            DataTable(73) = 0.626929
            DataTable(74) = 0.628985
            DataTable(75) = 0.631006
            DataTable(76) = 0.632991
            DataTable(77) = 0.634942
            DataTable(78) = 0.636861
            DataTable(79) = 0.638747
            DataTable(80) = 0.640602
            DataTable(81) = 0.642427
            DataTable(82) = 0.644222
            DataTable(83) = 0.645989
            DataTable(84) = 0.647728
            DataTable(85) = 0.649439
            DataTable(86) = 0.651125
            DataTable(87) = 0.652784
            DataTable(88) = 0.654419
            DataTable(89) = 0.656029
            DataTable(90) = 0.657615
            DataTable(91) = 0.659178
            DataTable(92) = 0.660719
            DataTable(93) = 0.662237
            DataTable(94) = 0.663734
            DataTable(95) = 0.66521
            DataTable(96) = 0.666665
            DataTable(97) = 0.668101
            DataTable(98) = 0.669517
            DataTable(99) = 0.670913
            DataTable(100) = 0.672291


        '����f�[�^
        ElseIf TopOrBottom = "top" Then
            DataTable(3) = 799.5
            DataTable(4) = 39.165495
            DataTable(5) = 15.100979
            DataTable(6) = 9.364471
            DataTable(7) = 6.977702
            DataTable(8) = 5.69547
            DataTable(9) = 4.899341
            DataTable(10) = 4.357233
            DataTable(11) = 3.963865
            DataTable(12) = 3.664914
            DataTable(13) = 3.429613
            DataTable(14) = 3.239263
            DataTable(15) = 3.081854
            DataTable(16) = 2.949321
            DataTable(17) = 2.836047
            DataTable(18) = 2.737998
            DataTable(19) = 2.652204
            DataTable(20) = 2.576425
            DataTable(21) = 2.508943
            DataTable(22) = 2.448414
            DataTable(23) = 2.393775
            DataTable(24) = 2.344171
            DataTable(25) = 2.298907
            DataTable(26) = 2.257412
            DataTable(27) = 2.219213
            DataTable(28) = 2.183913
            DataTable(29) = 2.15118
            DataTable(30) = 2.120728
            DataTable(31) = 2.092317
            DataTable(32) = 2.065736
            DataTable(33) = 2.040804
            DataTable(34) = 2.017366
            DataTable(35) = 1.995283
            DataTable(36) = 1.974435
            DataTable(37) = 1.954715
            DataTable(38) = 1.936029
            DataTable(39) = 1.918292
            DataTable(40) = 1.901431
            DataTable(41) = 1.885377
            DataTable(42) = 1.870071
            DataTable(43) = 1.855459
            DataTable(44) = 1.841492
            DataTable(45) = 1.828124
            DataTable(46) = 1.815317
            DataTable(47) = 1.803033
            DataTable(48) = 1.791239
            DataTable(49) = 1.779903
            DataTable(50) = 1.769
            DataTable(51) = 1.758501
            DataTable(52) = 1.748384
            DataTable(53) = 1.738628
            DataTable(54) = 1.729211
            DataTable(55) = 1.720115
            DataTable(56) = 1.711323
            DataTable(57) = 1.702819
            DataTable(58) = 1.694588
            DataTable(59) = 1.686616
            DataTable(60) = 1.678891
            DataTable(61) = 1.671399
            DataTable(62) = 1.664131
            DataTable(63) = 1.657075
            DataTable(64) = 1.650222
            DataTable(65) = 1.643562
            DataTable(66) = 1.637087
            DataTable(67) = 1.630789
            DataTable(68) = 1.62466
            DataTable(69) = 1.618692
            DataTable(70) = 1.61288
            DataTable(71) = 1.607216
            DataTable(72) = 1.601695
            DataTable(73) = 1.59631
            DataTable(74) = 1.591057
            DataTable(75) = 1.58593
            DataTable(76) = 1.580925
            DataTable(77) = 1.576037
            DataTable(78) = 1.571261
            DataTable(79) = 1.566594
            DataTable(80) = 1.562031
            DataTable(81) = 1.557569
            DataTable(82) = 1.553204
            DataTable(83) = 1.548933
            DataTable(84) = 1.544753
            DataTable(85) = 1.54066
            DataTable(86) = 1.536652
            DataTable(87) = 1.532725
            DataTable(88) = 1.528878
            DataTable(89) = 1.525107
            DataTable(90) = 1.521411
            DataTable(91) = 1.517786
            DataTable(92) = 1.514231
            DataTable(93) = 1.510743
            DataTable(94) = 1.50732
            DataTable(95) = 1.503961
            DataTable(96) = 1.500664
            DataTable(97) = 1.497426
            DataTable(98) = 1.494246
            DataTable(99) = 1.491123
            DataTable(100) = 1.488054
        Else
            
            MsgBox ("Program Argument Error!!")
        End If
    
    Else
        MsgBox ("Program Argument Error!!")
    
    End If

    '�z��f�[�^�������ŕԂ�
    sub_Get_F_Table = DataTable

End Function

'********************************************************************************************
' ���O: sub_RunLoopAuto
' ���e: �w�肵������񐔁E�m�[�h���ŁA���[�v����������Ŏ��{����
' ����: lngLoopCnt    : ���[�v��
'       intSwNode     : �e�X�^�m�[�h
' �ߒl: �I���t���O
' ���l�F �Ȃ�
' �X�V�����F Rev1.0      2011/08/10�@�V�K�쐬   K.Sumiyashiki
' �X�V�����F Rev2.0      2012/03/08�@���ʂƂ̘A�g�@�\�ǉ�   M.Imamura
'********************************************************************************************

Public Function sub_RunLoopAuto(ByVal lngLoopCnt As Long, ByVal intSwNode As Integer, Optional ByVal blnRunMode As Boolean = False, Optional blnDataMode As Boolean = True, Optional intMaxWait As Integer = 500, Optional intMaxTrialCount As Integer = 1) As Long

    sub_RunLoopAuto = 1
    Sw_Node = intSwNode

On Error GoTo errPALSsub_RunLoopAuto
    
    ThisWorkbook.Activate

    PALS_ParamFolder = ThisWorkbook.Path & "\" & PALS_PARAMFOLDERNAME
    Call sub_PalsFileCheck

    Set PALS = Nothing
    Set PALS = New csPALS

    '�A�g���t���O 2012/6/19
    g_RunAutoFlg_PALS = True

    'TestCondition�V�[�g�f�[�^�̍ēǍ�
    Call ReadCategoryData

    With frm_PALS_LoopAdj_Main
        .Show vbModeless
        .txt_loop_num.Value = lngLoopCnt
        
        If blnRunMode = False Then
            .op_NotAdjust.Value = True
            .op_AutoAdjust.Value = False
        Else
            .op_NotAdjust.Value = False
            .op_AutoAdjust.Value = True
        End If
        
        .Btn_ContinueOnFail.Value = blnDataMode
        .txt_maxwait = intMaxWait
        .txt_maxtrial_num = intMaxTrialCount
        Call .cmd_start_Click
    End With

    Unload frm_PALS_LoopAdj_Main

    If g_ErrorFlg_PALS = True Then
        GoTo errPALSsub_RunLoopAuto
    End If
    
    g_RunAutoFlg_PALS = False
    sub_RunLoopAuto = 0
    Exit Function

errPALSsub_RunLoopAuto:
    g_RunAutoFlg_PALS = False
    g_ErrorFlg_PALS = False

End Function





