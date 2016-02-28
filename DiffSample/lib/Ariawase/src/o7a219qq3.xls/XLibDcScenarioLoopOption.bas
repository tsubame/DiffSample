Attribute VB_Name = "XLibDcScenarioLoopOption"
'�T�v:
'   DC�V�i���ILoopOption�p���C�u����
'
'�ړI:
'
'�쐬��:
'   0145184306

Option Explicit

Public Enum SAMPLING_PARAM
    PARAM_TIME
    PARAM_COUNTS
End Enum ' SAMPLING_PARAM

Public Type LOOP_CONDITION
    LOOP_CATEGORIES As Collection
    Param As SAMPLING_PARAM
    PARAM_LIMIT As Double
    FOLDER As String
End Type 'LOOP_CONDITION


Private mDcLoop As Boolean  'LoopOption�؂�ւ��p�t���O
Private mNormalScenario As CDCScenario  '�ʏ�V�i���I�ۊǗp

Public Const SAMPLE_MODE_COUNTS As String = "SamplingCounts"
Public Const SAMPLE_MODE_TIME As String = "SamplingTime"
Public Const DCLOOP_FILE_HEADER As String = "DcLoopFileHeader"

Public Sub SetDcLoop(ByVal pLoop As Boolean)
    mDcLoop = pLoop
End Sub

Public Sub ApplyDcScenarioLoopOptionMode()
    If Not TheDcTest Is Nothing Then
        If mDcLoop = True Then
            '�ʏ�V�i���I�ۊ�
            Set mNormalScenario = TheDcTest
            
            Dim scenario As Collection
            Set scenario = mNormalScenario.Categories
            
            'Form�ݒ�
            Dim loopOptionForm As CDcScenarioLoopOptionForm
            Set loopOptionForm = New CDcScenarioLoopOptionForm
            Call loopOptionForm.Initialize(scenario)
            'Loop���s��
            If loopOptionForm.Show = True Then
                'Loop�p�V�i���I����
                Dim loopScenario As CDcScenarioLoopOption
                Set loopScenario = New CDcScenarioLoopOption
                Set loopScenario.Categories = scenario
                Call loopScenario.SetLoopOption(loopOptionForm.LoopCondition)
                '�V�i���I�����ւ�
                Set TheDcTest = loopScenario
                
            '�ʏ���s��
            Else
                Set scenario = Nothing
                Set mNormalScenario = Nothing
                mDcLoop = False
            End If
            Set loopOptionForm = Nothing
        End If
    End If
End Sub

Public Sub RunAtJobEnd()
    'Loop���s���`�F�b�N
    '�t���O�͎��s���ɏ�����������\��������̂�
    '�ʏ�V�i���I��ۊǂ��Ă��邩�Ń`�F�b�N����
    If Not mNormalScenario Is Nothing Then
        Set TheDcTest = Nothing
        Set TheDcTest = mNormalScenario
        Set mNormalScenario = Nothing
    End If
    mDcLoop = False
End Sub

Public Property Get DcLoop() As Boolean
    If mDcLoop = True Or Not mNormalScenario Is Nothing Then
        DcLoop = True
    Else
        DcLoop = False
    End If
End Property
