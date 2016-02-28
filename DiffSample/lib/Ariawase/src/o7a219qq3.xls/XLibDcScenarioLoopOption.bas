Attribute VB_Name = "XLibDcScenarioLoopOption"
'概要:
'   DCシナリオLoopOption用ライブラリ
'
'目的:
'
'作成者:
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


Private mDcLoop As Boolean  'LoopOption切り替え用フラグ
Private mNormalScenario As CDCScenario  '通常シナリオ保管用

Public Const SAMPLE_MODE_COUNTS As String = "SamplingCounts"
Public Const SAMPLE_MODE_TIME As String = "SamplingTime"
Public Const DCLOOP_FILE_HEADER As String = "DcLoopFileHeader"

Public Sub SetDcLoop(ByVal pLoop As Boolean)
    mDcLoop = pLoop
End Sub

Public Sub ApplyDcScenarioLoopOptionMode()
    If Not TheDcTest Is Nothing Then
        If mDcLoop = True Then
            '通常シナリオ保管
            Set mNormalScenario = TheDcTest
            
            Dim scenario As Collection
            Set scenario = mNormalScenario.Categories
            
            'Form設定
            Dim loopOptionForm As CDcScenarioLoopOptionForm
            Set loopOptionForm = New CDcScenarioLoopOptionForm
            Call loopOptionForm.Initialize(scenario)
            'Loop実行時
            If loopOptionForm.Show = True Then
                'Loop用シナリオ準備
                Dim loopScenario As CDcScenarioLoopOption
                Set loopScenario = New CDcScenarioLoopOption
                Set loopScenario.Categories = scenario
                Call loopScenario.SetLoopOption(loopOptionForm.LoopCondition)
                'シナリオ差し替え
                Set TheDcTest = loopScenario
                
            '通常実行時
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
    'Loop実行かチェック
    'フラグは実行中に書き換えられる可能性があるので
    '通常シナリオを保管しているかでチェックする
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
