VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDCScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DCシナリオの実行エンジンクラス
'       Revision History:
'           Date        Description
'           2013/6/11   IDcScnarioを実装、自身の関数の呼び出しをAsIDcScenarioを介するように修正(0145184306)
'
'目的:
'   DCアクションオブジェクトの作成と実行制御を行う
'
'作成者:
'   0145206097
'
Option Explicit

Implements IDcScenario

Dim mScenarioReader As IParameterReader
Dim mInstanceReader As IParameterReader
Dim mScenarioWriter As IParameterWriter
Dim mResultKeeper As CResultManager
Dim mDcContainer As CContainer
Dim mDcScenario As Collection
Dim mCurrScenario As CDcCategory
Dim mLogWriter As IParameterWriter
Dim mTextFileWriter As IParameterWriter
Dim mTempLogWriter As IParameterWriter
Dim mPinInfoReportWriter As IParameterWriter

Dim mExecWatch As CStopWatch
Dim mMeasureCounter As Long
'仮置変数(Pins/Chans変更用)
Dim mAPMUapichoice As String
Dim mBPMUapichoice As String
Dim mPPMUapichoice As String
Dim mHDVISapichoice As String
Dim mDPSapichoice As String
Dim mICUL1Gapichoice As String

'テスト用ライター
'Dim mActionParamWriter As CActionLogWriter

Private Sub Class_Initialize()
    Set mDcContainer = New CContainer
    Set mDcScenario = New Collection
    Set mCurrScenario = New CDcCategory
End Sub

Private Sub Class_Terminate()
    Set mScenarioReader = Nothing
    Set mInstanceReader = Nothing
    Set mScenarioWriter = Nothing
    Set mResultKeeper = Nothing
    Set mDcContainer = Nothing
    Set mDcScenario = Nothing
    Set mCurrScenario = Nothing
    Set mLogWriter = Nothing
    Set mTextFileWriter = Nothing
    Set mTempLogWriter = Nothing
    Set mExecWatch = Nothing
    Set mPinInfoReportWriter = Nothing
'    Set mActionParamWriter = Nothing
End Sub

Public Function AsIDcScenario() As IDcScenario
    Set AsIDcScenario = Me
End Function

Public Property Get Categories() As Collection
    Set Categories = mDcScenario
End Property
Private Sub IDcScenario_Initialize(ByVal dcPerformer As IDcTest, ByVal scenarioReader As IParameterReader, ByVal instanceReader As IParameterReader, ByVal scenarioWriter As IParameterWriter, ByVal pinInfoWriter As IParameterWriter)
    On Error GoTo ErrHandler
    Set mScenarioReader = scenarioReader
    Set mInstanceReader = instanceReader
    Set mScenarioWriter = scenarioWriter
    Set mPinInfoReportWriter = pinInfoWriter
    mDcContainer.DcTest = dcPerformer
    Call Me.AsIDcScenario.Apichoice
    Call Me.AsIDcScenario.Load
    
    '### テスト用アクションパラメータライターの準備 #######
'    Set mActionParamWriter = New CActionLogWriter
'    mActionParamWriter.Initialize "DC Scenario Exec Report", GetSiteCount
'    Set mActionParamWriter = Nothing
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CDcScenario.Initialize", "Can Not Initialize DC Test Scenario !"
End Sub

Private Sub IDcScenario_Load()
    On Error GoTo ErrHandler
    constructScenario mScenarioReader
    Exit Sub
ErrHandler:
'    OutputErrMsg "Can Not Construct DC Test Scenario !"
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CDcScenario.Load", "Can Not Construct DC Test Scenario !"
End Sub
'#V21-Release
Private Sub IDcScenario_Apichoice()
    
    Dim APMUsetapichoice As String
    Dim BPMUsetapichoice As String
    Dim PPMUsetapichoice As String
    Dim HDVISsetapichoice As String
    Dim DPSsetapichoice As String
    Dim ICUL1Gsetapichoice As String
    Dim APMUconapichoice As String
    Dim BPMUconapichoice As String
    Dim PPMUconapichoice As String
    Dim HDVISconapichoice As String
    Dim DPSconapichoice As String
    Dim ICUL1Gconapichoice As String
    
    'Pins/Chans指定  検証用仮ベタ 本来はCVSファイルべた書き
    APMUsetapichoice = "Chans"   'APMUsetapichoice = Worksheets("CVS").Range("B100")
    BPMUsetapichoice = "Pins"   'BPMUsetapichoice = Worksheets("CVS").Range("B101")
    PPMUsetapichoice = "Pins"   'PPMUsetapichoice = Worksheets("CVS").Range("B102")
    HDVISsetapichoice = "Pins" 'HDVISsetapichoice = Worksheets("CVS").Range("B103")
    DPSsetapichoice = "Pins"   'DPSsetapichoice = Worksheets("CVS").Range("B104")
    ICUL1Gsetapichoice = "Pins"   'DPSsetapichoice = Worksheets("CVS").Range("B???")
    APMUconapichoice = "Pins"   'APMUconapichoice = Worksheets("CVS").Range("B105")
    BPMUconapichoice = "Pins"   'BPMUconapichoice = Worksheets("CVS").Range("B106")
    PPMUconapichoice = "Pins"   'PPMUconapichoice = Worksheets("CVS").Range("B107")
    HDVISconapichoice = "Pins" 'HDVISconapichoice = Worksheets("CVS").Range("B108")
    DPSconapichoice = "Pins"   'DPSconapichoice = Worksheets("CVS").Range("B109")
    ICUL1Gconapichoice = "Pins"   'DPSconapichoice = Worksheets("CVS").Range("B???")
    'error check予定　　On Error Goto xxx
    '                   xxx:
    '                    "Error Occured !! CVS WorkSheets is strange"
    'Pins/Chans設定仕組み用
    '  On Error GoTo Apichoicehandler
    '  APMUsetapichoice = Worksheets("CVS").Range("B100")
    '  BPMUsetapichoice = Worksheets("CVS").Range("B101")
    '  PPMUsetapichoice = Worksheets("CVS").Range("B102")
    '  HDVISsetapichoice = Worksheets("CVS").Range("B103")
    '  DPSsetapichoice = Worksheets("CVS").Range("B104")
    '  ICUL1Gsetapichoice = Worksheets("CVS").Range("B???")
    '  APMUconapichoice = Worksheets("CVS").Range("B105")
    '  BPMUconapichoice = Worksheets("CVS").Range("B106")
    '  PPMUconapichoice = Worksheets("CVS").Range("B107")
    '  HDVISconapichoice = Worksheets("CVS").Range("B108")
    '   DPSconapichoice = Worksheets("CVS").Range("B109")
    '  ICUL1Gconapichoice = Worksheets("CVS").Range("B???")
    'Apichoicehandler: Err.Raise 9999, "Error Occured !! CVS WorkSheets is strange"
    
    
    
    'Status設定　0:Set-Pins/Connect-Pins , 1:Set-Pins/Connect-Chans , 2:Set-Chans/Connect-Chans , 3:Set-Chans/Connect-Pins
    Select Case APMUsetapichoice
        Case "Pins":
        Select Case APMUconapichoice
               Case "Pins":
                    mAPMUapichoice = 0
               Case "Chans":
                    mAPMUapichoice = 1
               Case Else:
                    'error表示予定　必要性を後で判断 MsgBox "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case "Chans":
        Select Case APMUconapichoice
               Case "Pins":
                    mAPMUapichoice = 3
               Case "Chans":
                    mAPMUapichoice = 2
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
    End Select

    Select Case BPMUsetapichoice
        Case "Pins":
        Select Case BPMUconapichoice
               Case "Pins":
                    mBPMUapichoice = 0
               Case "Chans":
                    mBPMUapichoice = 1
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case "Chans":
        Select Case BPMUconapichoice
               Case "Pins":
                    mBPMUapichoice = 3
               Case "Chans":
                    mBPMUapichoice = 2
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
    End Select

    Select Case PPMUsetapichoice
        Case "Pins":
        Select Case PPMUconapichoice
               Case "Pins":
                    mPPMUapichoice = 0
               Case "Chans":
                    mPPMUapichoice = 1
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case "Chans":
        Select Case PPMUconapichoice
               Case "Pins":
                    mPPMUapichoice = 3
               Case "Chans":
                    mPPMUapichoice = 2
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
    End Select

    Select Case HDVISsetapichoice
        Case "Pins":
        Select Case HDVISconapichoice
               Case "Pins":
                    mHDVISapichoice = 0
               Case "Chans":
                    mHDVISapichoice = 1
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case "Chans":
        Select Case HDVISconapichoice
               Case "Pins":
                    mHDVISapichoice = 3
               Case "Chans":
                    mHDVISapichoice = 2
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
    End Select

    Select Case DPSsetapichoice
        Case "Pins":
        Select Case DPSconapichoice
               Case "Pins":
                    mDPSapichoice = 0
               Case "Chans":
                    mDPSapichoice = 1
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case "Chans":
        Select Case DPSconapichoice
               Case "Pins":
                    mDPSapichoice = 3
               Case "Chans":
                    mDPSapichoice = 2
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
    End Select

    Select Case ICUL1Gsetapichoice
        Case "Pins":
        Select Case ICUL1Gconapichoice
               Case "Pins":
                    mICUL1Gapichoice = 0
               Case "Chans":
                    mICUL1Gapichoice = 1
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case "Chans":
        Select Case ICUL1Gconapichoice
               Case "Pins":
                    mICUL1Gapichoice = 3
               Case "Chans":
                    mICUL1Gapichoice = 2
               Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
        End Select
        Case Else:
                    'error表示予定　必要性を後で判断 "Error Occured !! CVS WorkSheets's Cell Changed"
    End Select

End Sub
Private Sub IDcScenario_ClearContainer()
    mDcContainer.Initialize
End Sub
'#V21-Release
Private Sub IDcScenario_GetTempResult(ByVal pName As String, ByRef pResult() As Double)
    On Error GoTo ErrHandler
        mDcContainer.TempResults.GetResult pName, pResult
    Exit Sub
    
ErrHandler:
    Call OutputErrMsg("CDCScenario.GetTempResult() [" & pName & "] Doesn't Exist.")
    
End Sub
'Private Function IsContain(ByVal pName As String) As Boolean消去


Private Property Let IDcScenario_ResultManager(ByVal resultKeeper As CResultManager)
    Set mResultKeeper = resultKeeper
'    mActionParamWriter.AsIActionStream.Rewind
End Property

'############# DCシナリオ実行エンジンメソッド #####################################################
Private Sub IDcScenario_SetScenario(ByVal currTest As String)
'内容:
'   実行するテストカテゴリの決定と実験モード時の各種設定を行う
'
'パラメータ:
'[currTest]   In  テストインスタンス名
'
'注意事項:
'
    On Error GoTo ErrHandler
    '### カレントテストの決定 #############################
    Set mCurrScenario = mDcScenario.Item(currTest)
    '### 実験モード設定 ###################################
    setExaminationMode
    '### 条件設定前 一時停止 ##############################
    pauseExecution TEST_START
    Exit Sub
ErrHandler:
    Err.Raise 9999, "CDcScenario.SetScenario", "Can Not Find The Test Name [" & currTest & "] !"
End Sub

Private Sub IDcScenario_Execute()
'内容:
'   指定されたテストカテゴリのアクションを実行
'
'パラメータ:
'
'注意事項:
'   実行前に必ずSetScenarioで実行するカテゴリを決定する必要がある
'
    On Error GoTo ErrHandler
    '### シナリオ実行開始 #################################
    Dim currGroup As Collection
    Dim actionIndex As Long
    Dim lastAction As IAction
    Dim measures As Collection
    Dim subMeasure As Collection
    Dim currMeasure As IMeasure
    Do
        Set measures = New Collection
        For Each currGroup In mCurrScenario.Groups
            Do
                Set subMeasure = New Collection
                '### グループアクション開始前 一時停止 ####
                pauseExecution FIRST_ACT, currGroup.Item(1)
                '### タイムスタンプ #######################
                If Not mExecWatch Is Nothing Then mExecWatch.Start
                '### グループアクション実行開始 ###########
                For actionIndex = 1 To currGroup.Count
                    Select Case currGroup.Item(actionIndex).Name
                        Case MEASURE_CLASS:
                            mMeasureCounter = 0
                            Dim tempMeasure As IMeasure
                            Do While loopComplete = False
                                '@@@ アクション実行 @@@@@@@
                                currGroup.Item(actionIndex).Execute
                                mMeasureCounter = mMeasureCounter + 1
                                '### メジャーアクション登録
                                Set tempMeasure = currGroup.Item(actionIndex)
                                '### ピンパラメータのダンプ
                                measureActionLogger tempMeasure, mMeasureCounter, mLogWriter
                                '### ピンパラメータのダンプ
                                measurePinInfoReport tempMeasure, mPinInfoReportWriter
                                '### 測定後 一時停止 ######
                                pauseExecution , currGroup.Item(actionIndex)
                            Loop

                            '### MEASUREアクションの登録 ##
                            subMeasure.Add tempMeasure
                        Case Else
                            '@@@ アクション実行 @@@@@@@@@@@
                            currGroup.Item(actionIndex).Execute
                            '### アクションピンパラメータのダンプ
                            actionLogger currGroup.Item(actionIndex), mLogWriter
                            '### アクション実行後 一時停止
                            pauseExecution , currGroup.Item(actionIndex)
                    End Select
                    Set lastAction = currGroup.Item(actionIndex)
                Next actionIndex
                '### タイムスタンプ #######################
                If Not mExecWatch Is Nothing Then timeStamp lastAction, mExecWatch.GetLapTime, mLogWriter
                '### グループアクション終了後 一時停止 ####
                pauseExecution END_OF_ACT, lastAction
            Loop While isRetuneMode
            '### MEASUREアクションの登録 ##################
            For Each currMeasure In subMeasure
                measures.Add currMeasure
            Next currMeasure
        Next currGroup
        '### シナリオ終了後ブレイクポイント ###############
        pauseExecution END_OF_TEST
    Loop While isRetuneMode
    '### 測定結果の格納 ###################################
    For Each currMeasure In measures
        currMeasure.WriteResult mResultKeeper
    Next currMeasure
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CDcScenario.Execute", "Can Not Execute The Test Name [" & mCurrScenario.Name & "] !"
End Sub

'############# 実験用実行制御プロシージャ群 #######################################################
Private Sub setExaminationMode()
    '### 実験フォームの初期設定 ###########################
    With DcTestExaminForm
        .ControlStatus = INITIAL
        .loopCount = 1
    End With
    '### ライター・実験フォームの設定 #####################
    With mCurrScenario
        If .IsExamination Then
            '### カテゴリパラメータのダンプ ###############
            Set mLogWriter = mScenarioWriter
            categoryLogger mCurrScenario, mLogWriter
            Set mExecWatch = New CStopWatch
            Select Case .ExamMode
                Case "BREAK":
                    DcTestExaminForm.ControlStatus = TEST_STEP
                Case "END":
                    DisableAllTest
            End Select
        Else
            Set mLogWriter = Nothing
            Set mExecWatch = Nothing
        End If
'        Set mLogWriter = mActionParamWriter
'        categoryLogger mCurrScenario, mLogWriter
    End With
End Sub

Private Sub pauseExecution(Optional ByVal testStatus As status = RUNNING, Optional ByVal action As IAction = Nothing)
    Dim textWriter As CDcTextFileLogWriter
    With DcTestExaminForm
        If .ControlStatus = TEST_REPEAT Then
                '### アクション繰り返し時のフォーム設定 ###
                With DcTestContinuousForm
                    DoEvents
                    .MeasureCounter.Caption = mMeasureCounter
                    If .MeasureStatus = MEAS_STOP Then
                        .ExitButton.enabled = True
                        .RestartButton.enabled = True
                        .StopButton.enabled = False
                        Do While .MeasureStatus = MEAS_STOP
                            DoEvents
                        Loop
                        Select Case .MeasureStatus
                            Case MEAS_EXIT:
                                DcTestExaminForm.loopCount = mMeasureCounter
                                .MeasureStatus = MEAS_INITIAL
                                .Hide
                            Case MEAS_RESTART:
                                .MeasureStatus = MEAS_INITIAL
                                .ExitButton.enabled = False
                                .RestartButton.enabled = False
                                .StopButton.enabled = True
                        End Select
                    End If
                End With
            '### アクション繰り返し終了 ###################
            If loopComplete = True Then
                DcTestContinuousForm.Hide
                TheResult.IsOverwriteMode = False
                .ControlStatus = TEST_STEP
                If .OutPutFileCheck Then
                    Set textWriter = Nothing
                    Set mLogWriter = mTempLogWriter
                    .OutPutFileCheck = False
                End If
            End If
        End If
        If .ControlStatus = TEST_STEP Then
            '### 次のステップの問い合わせ #################
            Select Case queryNextStep(testStatus, action)
                Case TEST_STEP:
                    mMeasureCounter = 1
                    .loopCount = mMeasureCounter
                    .OutPutFileCheck = False
                Case TEST_END:
                    mMeasureCounter = 1
                    .loopCount = mMeasureCounter
                    .OutPutFileCheck = False
                    Call DisableAllTest
                Case TEST_RETURN:
                    .OutPutFileCheck = False
                Case TEST_REPEAT:
                    mMeasureCounter = 0
                    '### アクション繰り返し時のフォーム設定
                    With DcTestContinuousForm
                        .ExitButton.enabled = False
                        .RestartButton.enabled = False
                        .StopButton.enabled = True
                        .Show vbModeless
                    End With
                    '### テキストファイルライター設定 #####
                    If .OutPutFileCheck Then
                        Set mTempLogWriter = mLogWriter
                        Set textWriter = CreateCDcTextFileLogWriter
                        categoryLogger mCurrScenario, textWriter
                        Set mLogWriter = textWriter
                    End If
                Case TEST_CONTINUE:
                    mMeasureCounter = 1
                    .loopCount = mMeasureCounter
                    .OutPutFileCheck = False
                    .ControlStatus = INITIAL
            End Select
        End If
    End With
End Sub

Private Function loopComplete() As Boolean
    '### アクション繰り返し終了判定 #######################
    loopComplete = (mMeasureCounter >= DcTestExaminForm.loopCount)
End Function

Private Function isRetuneMode() As Boolean
    '### グループアクションの終了判定 #####################
    With DcTestExaminForm
        Select Case .ControlStatus
            Case TEST_RETURN:
                isRetuneMode = True
                .ControlStatus = TEST_STEP
        End Select
    End With
End Function

Private Sub categoryLogger(ByVal testCategory As CDcCategory, ByVal writer As IActionStream)
    '### テストカテゴリパラメータのダンプ #################
    If Not writer Is Nothing Then
        writer.SetCategoryLocation testCategory.Name
        paramLogger testCategory, writer
    End If
End Sub

Private Sub measureActionLogger(ByVal action As IMeasure, ByVal lCount As Long, ByVal writer As IActionStream)
    '### メジャーアクションパラメータのダンプ #############
    If Not writer Is Nothing Then
        actionLogger action, writer
        Dim ParamCnt As CParamLong
        Set ParamCnt = CreateCParamLong
        With ParamCnt.AsIParameter
            .Name = REPEAT_COUNTER
            .AsLong = lCount
            .WriteOut writer
        End With
        measurePinLogger action, writer
    End If
End Sub

Private Sub actionLogger(ByVal action As IAction, ByVal writer As IActionStream)
    '### アクションパラメータのダンプ #####################
    If Not writer Is Nothing Then
        writer.SetGroupLocation action.GroupIndex
        paramLogger action, writer
    End If
End Sub

Private Sub measurePinLogger(ByVal mPins As IMeasure, ByVal writer As IActionStream)
    '### メジャーピンパラメータのダンプ ###################
    If Not writer Is Nothing Then
        Dim currPin As CMeasurePin
        For Each currPin In mPins.MeasurePins
            paramLogger currPin, writer
            writer.MoveNextData
        Next currPin
    End If
End Sub

Private Sub measurePinInfoReport(ByVal mPins As IMeasure, ByVal writer As IFileStream)
    '### メジャーピンパラメータのダンプ ###################
    If Not writer Is Nothing Then
        mPins.GetBoardInfo
        Dim currPin As CMeasurePin
        For Each currPin In mPins.MeasurePins
            currPin.ValidateRange
            paramLogger currPin, writer
            writer.MoveNext
        Next currPin
    End If
End Sub

Private Sub paramLogger(ByVal writable As IParameterWritable, ByVal writer As IParameterWriter)
    '### 型変換のためのプロシージャ #######################
    writable.WriteOut writer
End Sub

Private Sub timeStamp(ByVal action As IAction, ByVal stopTime As Double, ByVal writer As IActionStream)
    '### タイムスタンプログ出力 ###########################
    If Not writer Is Nothing Then
        writer.SetGroupLocation action.GroupIndex
        Dim ExecTime As CParamDouble
        Set ExecTime = CreateCParamDouble
        With ExecTime.AsIParameter
            .Name = EXAMIN_EXECTIME
            .AsDouble = stopTime
            .WriteOut writer
        End With
    End If
End Sub

Private Function queryNextStep(ByVal testStatus As status, ByVal action As IAction) As CONTROL_STATUS
    '### 各アクションパラメータ出力 #######################
    Dim writer As New CDcFormMessageWriter
    mCurrScenario.AsIParameterWritable.WriteOut writer
    If Not action Is Nothing Then
        Dim writable As IParameterWritable
        Set writable = action
        writable.WriteOut writer
        With writer.AsIParameterWriter
            Select Case testStatus
                Case FIRST_ACT:
                    .WriteAsString ACTION_LABEL, "Action Start"
                    .WriteAsBoolean MEASURE_FRM, True
                    .WriteAsBoolean RETURN_BTN, False
                Case END_OF_ACT:
                    .WriteAsString ACTION_LABEL, "Action Completed"
                    .WriteAsBoolean MEASURE_FRM, False
                    .WriteAsBoolean RETURN_BTN, True
            End Select
            If action.Name = DISCONNECT_CLASS Then
                .WriteAsBoolean MEASURE_FRM, False
                .WriteAsBoolean RETURN_BTN, True
            End If
        End With
    Else
        With writer.AsIParameterWriter
            Select Case testStatus
                Case TEST_START:
                    .WriteAsString ACTION_LABEL, "Before Device Drive"
                    .WriteAsLong GROUP_ID, NO_TEST
                    .WriteAsLong SITE_INDEX, NO_SITE
                    .WriteAsBoolean MEASURE_FRM, True
                    .WriteAsBoolean RETURN_BTN, False
                Case END_OF_TEST:
                    .WriteAsString ACTION_LABEL, "Scenario Completed"
                    .WriteAsLong GROUP_ID, NO_TEST
                    .WriteAsLong SITE_INDEX, NO_SITE
                    .WriteAsBoolean MEASURE_FRM, False
                    .WriteAsBoolean RETURN_BTN, True
            End Select
        End With
    End If
    Set writer = Nothing
    '### アクションステータス表示 #########################
    Dim sheetWriter As CDcScenarioSheetLogWriter
    On Error Resume Next
    Set sheetWriter = mLogWriter
    Dim actionID As Long
    If action Is Nothing Then
        actionID = NO_TEST
    Else
        actionID = action.GroupIndex
    End If
    sheetWriter.ShowCurrAction actionID
    '### フォームの表示 ###################################
    With DcTestExaminForm
        .ControlStatus = INITIAL
        .Show vbModeless
        Do While .ControlStatus = INITIAL
            '### パラメータ変更の確認 #####################
            If mScenarioReader.ReadAsBoolean(DATA_CHANGED) Then
                reconstructGroupAction mScenarioReader
            End If
            DoEvents
        Loop
        queryNextStep = .ControlStatus
        .Hide
    End With
End Function

'############# パラメータテスト用プロシージャ群 ###################################################
Private Sub writeOutInitialAction(ByVal scenario As Collection, ByVal SheetName As String)
'内容:
'   アクションパラメータ確認用テスト関数
'
'パラメータ:
'[scenario]   In  アクションオブジェクトの集合コレクション
'[sheetName]   In  パラメータ出力先ワークシート名
'
'注意事項:
'   アクション作成後に実行する
'   出力先ワークシートは予め任意のシート名で用意する
'
'    Dim testWriter As CActionParamWriter
'    Set testWriter = New CActionParamWriter
'    testWriter.Initialize sheetName
'    Dim currScenario As CDcCategory
'    Dim currGroup As Collection
'    Dim currAction As IAction
'    For Each currScenario In scenario
'        currScenario.WriteClassParameter testWriter
'        For Each currGroup In currScenario.Groups
'            For Each currAction In currGroup
'                currAction.WriteParameter testWriter
'                If currAction.Name = MEASURE_CLASS Then
'                    Dim measure As IMeasure
'                    Set measure = currAction
'                    Dim currPin As CMeasurePin
'                    For Each currPin In measure.MeasurePins
'                        currPin.WriteClassParameter testWriter
'                        testWriter.AsIFileStream.MoveNext
'                    Next currPin
'                Else
'                    testWriter.AsIFileStream.MoveNext
'                End If
'            Next currAction
'            testWriter.AsIFileStream.MoveNext
'        Next currGroup
'    Next currScenario
End Sub

'############# アクション構築プロシージャ群 #######################################################
Private Sub constructScenario(ByVal reader As IActionStream)
    '### アクションオブジェクト作成の実行##################
    Set mDcScenario = New Collection
    Dim dcCategory As CDcCategory
    Dim contentOfTest As Collection
    reader.Rewind
    Do While Not reader.IsEndOfCategory
        Set dcCategory = New CDcCategory
        Set contentOfTest = New Collection
        Do While Not reader.IsEndOfGroup
            contentOfTest.Add actionGroup(reader)
            reader.MoveNextGroup
        Loop
        dcCategory.Initialize contentOfTest, reader
        mDcScenario.Add dcCategory, Replace(dcCategory.Name, " ", "")
        reader.MoveNextCategory
    Loop
    Exit Sub
End Sub

Private Sub reconstructGroupAction(ByVal reader As IActionStream)
    '### リーダーから変更箇所の読み出し ###################
    reader.SeekChangedLocation
    Dim testName As CParamString
    Set testName = CreateCParamString
    Dim testKeyName As String
    With testName.AsIParameter
        .Name = TEST_CATEGORY
        .Read reader
        testKeyName = Replace(.AsString, " ", "")
    End With
    Dim groupNum As CParamLong
    Set groupNum = CreateCParamLong
    With groupNum.AsIParameter
        .Name = GROUP_ID
        .Read reader
    End With
    '### 更新対象のアクショングループ抽出 #################
    Dim currCItem As New Collection
    Set currCItem = mDcScenario(testKeyName).Groups
    Dim currGroup As New Collection
    Set currGroup = currCItem(groupNum.AsIParameter.AsLong + 1)
    '### アクショングループの再構築 #######################
    Dim newAction As New Collection
    On Error GoTo ErrHandler
    Set newAction = actionGroup(reader)
    Dim removeCount As Long
    Dim currAction As IAction
    removeCount = currGroup.Count
    For Each currAction In newAction
        currGroup.Add currAction
    Next currAction
    Dim dataIndex As Long
    For dataIndex = 1 To removeCount
        currGroup.Remove 1
    Next dataIndex
    '### カテゴリパラメータの更新 #########################
    mDcScenario(testKeyName).Initialize currCItem, reader
    reader.SetCategoryLocation mCurrScenario.Name
Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub
'#V21-release
Private Function actionGroup(ByVal reader As IParameterReader) As Collection
    '### リーダーからパラメータの読み出し #################
    Dim mainAction As CParamString
    Set mainAction = CreateCParamString
    With mainAction.AsIParameter
        .Name = TEST_ACTION
        .Read reader
    End With
    Dim postAction As CParamString
    Set postAction = CreateCParamString
    With postAction.AsIParameter
        .Name = TEST_POSTACTION
        .Read reader
    End With
    
    Dim measureSiteMode As CParamString
    Set measureSiteMode = CreateCParamString
    With measureSiteMode.AsIParameter
        .Name = MEASURE_SITE
        .Read reader
    End With
    Dim imputSitemode As String
    imputSitemode = measureSiteMode.AsIParameter.AsString
    
    '### グループアクションの作成 #########################
    Set actionGroup = New Collection
    With actionGroup
        If imputSitemode = "Concurrent" Or imputSitemode = "Serial" Then
            Select Case mainAction.AsIParameter.AsString
                Case "SET":
                    .Add setModeAction(reader)
                Case "MEASURE":
                    .Add measureAction(reader)
                Case "OPEN":
                    .Add disconnectAction(reader)
                Case "":
                    .Add setModeAction(reader)
                    .Add measureAction(reader)
                Case Else:
                    Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Action [" & mainAction.AsIParameter.AsString & "] !"
            End Select
            Select Case postAction.AsIParameter.AsString
                Case "OPEN":
                    .Add disconnectAction(reader)
                Case "":
                Case Else:
                    Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Post Action [" & postAction.AsIParameter.AsString & "] !"
            End Select
        Else 'Site指定用
            Select Case mainAction.AsIParameter.AsString
                Case "SET":
                    .Add setModeAction(reader, CLng(imputSitemode))
                Case "MEASURE":
                    .Add measureAction(reader, CLng(imputSitemode))
                Case "OPEN":
                    .Add disconnectAction(reader, CLng(imputSitemode))
                Case "":
                    .Add setModeAction(reader, CLng(imputSitemode))
                    .Add measureAction(reader, CLng(imputSitemode))
                Case Else:
                    Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Action [" & mainAction.AsIParameter.AsString & "] !"
            End Select
            Select Case postAction.AsIParameter.AsString
                Case "OPEN":
                    .Add disconnectAction(reader, CLng(imputSitemode))
                Case "":
                Case Else:
                    Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Post Action [" & postAction.AsIParameter.AsString & "] !"
            End Select
        End If
    End With
End Function
'#V21-Release
Private Function setModeAction(ByVal reader As IParameterReader, Optional ByVal measSite As Long = ALL_SITE) As IAction
    '### リーダーからパラメータの読み出し #################
    Dim SourceType As CParamString
    Set SourceType = CreateCParamString
    With SourceType.AsIParameter
        .Name = TEST_PINTYPE
        .Read reader
    End With
    Dim modeData As CParamString
    Set modeData = CreateCParamString
    With modeData.AsIParameter
        .Name = SET_MODE
        .Read reader
    End With
    Dim forceData As CParamString
    Set forceData = CreateCParamString
    With forceData.AsIParameter
        .Name = SET_FORCE
        .Read reader
    End With
    '### アクション内容の決定 #############################
    Dim setMode As ISetMode
 
    Dim measureSiteMode As CParamString
    Set measureSiteMode = CreateCParamString
    With measureSiteMode.AsIParameter
        .Name = MEASURE_SITE
        .Read reader
    End With
    
    Dim Boardchoice As String '暫定Board共用の変数
           
    Select Case SourceType.AsIParameter.AsString
        Case "BPMU":
            Boardchoice = mBPMUapichoice
            Select Case measureSiteMode.AsIParameter.AsString
                Case "Concurrent":
                    'Multi SiteのみBPMU利用制限Check ,Single Site
                    If CountExistSite > 0 Then
                        Call BPMUchcheck(reader)
                    End If
                    'Pins/Chans機能選択有
                    Select Case modeData.AsIParameter.AsString
                        Case "MI":
                            Select Case forceData.AsIParameter.AsString
                                Case "":
                                    Select Case Boardchoice
                                        Case "0":
                                            Set setMode = New CSetIRangeBPMUConcurrent 'PinsPins予定
                                        Case "1":
                                            Set setMode = New CSetIRangeBPMUConcurrent 'PinsChans予定
                                        Case "2":
                                            Set setMode = New CSetIRangeBPMUConcurrent 'ChansPins予定
                                        Case "3":
                                            Set setMode = New CSetIRangeBPMUConcurrent 'ChansChans予定
                                    End Select
                                Case Else
                                    Select Case Boardchoice
                                        Case "0":
                                            Set setMode = New CSetFVBPMUConCurrent
                                        Case "1":
                                            Set setMode = New CSetFVBPMUConCurrent
                                        Case "2":
                                            Set setMode = New CSetFVBPMUConCurrent
                                        Case "3":
                                            Set setMode = New CSetFVBPMUConCurrent
                                    End Select
                            End Select
                        Case "MV":
                            Select Case forceData.AsIParameter.AsString
                                Case "":
                                    Select Case Boardchoice
                                        Case "0":
                                            Set setMode = New CSetMVBPMUConCurrent
                                        Case "1":
                                            Set setMode = New CSetMVBPMUConCurrent
                                        Case "2":
                                            Set setMode = New CSetMVBPMUConCurrent
                                        Case "3":
                                            Set setMode = New CSetMVBPMUConCurrent
                                    End Select
                                Case Else
                                    Select Case Boardchoice
                                        Case "0":
                                            Set setMode = New CSetFIBPMUConCurrent
                                        Case "1":
                                            Set setMode = New CSetFIBPMUConCurrent
                                        Case "2":
                                            Set setMode = New CSetFIBPMUConCurrent
                                        Case "3":
                                            Set setMode = New CSetFIBPMUConCurrent
                                    End Select
                            End Select
                        Case Else
                            Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Test Mode [" & modeData.AsIParameter.AsString & "] !"
                    End Select
                Case "Serial":      'Serial
                    'Multi SiteのみBPMU利用制限Check ,Single Site
                    If CountExistSite > 0 Then
                        Call BPMUchcheck(reader)
                    End If
                    Select Case modeData.AsIParameter.AsString
                        Case "MI":
                            Select Case forceData.AsIParameter.AsString
                                Case "":
                                  Set setMode = New CSetIRangeBPMU
                                Case Else
                                  Set setMode = New CSetFVBPMU
                            End Select
                        Case "MV":
                            Select Case forceData.AsIParameter.AsString
                                Case "":
                                  Set setMode = New CSetMVBPMU
                                Case Else
                                  Set setMode = New CSetFIBPMU
                            End Select
                        Case Else
                            Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Test Mode [" & modeData.AsIParameter.AsString & "] !"
                    End Select
                
                Case Else '"サイト指定"
                    Call BPMUchcheckB(reader) 'BPMU Check pin複数
                    Select Case modeData.AsIParameter.AsString
                        Case "MI":
                            Select Case forceData.AsIParameter.AsString
                                Case "":
                                  Set setMode = New CSetIRangeBPMU
                                Case Else
                                  Set setMode = New CSetFVBPMU
                            End Select
                        Case "MV":
                            Select Case forceData.AsIParameter.AsString
                                Case "":
                                  Set setMode = New CSetMVBPMU
                                Case Else
                                  Set setMode = New CSetFIBPMU
                            End Select
                        Case Else
                            Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Test Mode [" & modeData.AsIParameter.AsString & "] !"
                    End Select
            End Select
        Case Else '"BPMU以外"
            Select Case measureSiteMode.AsIParameter.AsString
                Case "Concurrent":
                    Select Case SourceType.AsIParameter.AsString  '2重だが、Board判別用
                        Case "APMU": Boardchoice = mAPMUapichoice
                        Case "PPMU": Boardchoice = mPPMUapichoice
                        Case "DPS": Boardchoice = mDPSapichoice
                        Case "HDVIS": Boardchoice = mHDVISapichoice
                        Case "ICUL1G": Boardchoice = mICUL1Gapichoice
                    End Select
                    'Pins/Chans機能選択有
                    Select Case modeData.AsIParameter.AsString
                        Case "MI":
                            Select Case forceData.AsIParameter.AsString
                                Case "":
                                    Select Case Boardchoice
                                        Case "0":
                                            Set setMode = New CSetIRangeConcurrent 'PinsPins予定
                                        Case "1":
                                            Set setMode = New CSetIRangeConcurrent 'PinsChans予定
                                        Case "2":
                                            Set setMode = New CSetIRangeConcurrent 'ChansPins予定
                                        Case "3":
                                            Set setMode = New CSetIRangeConcurrent 'ChansChans予定
                                    End Select
                                Case Else
                                    Select Case Boardchoice
                                        Case "0":
                                            Set setMode = New CSetFVConcurrent
                                        Case "1":
                                            Set setMode = New CSetFVConcurrent
                                        Case "2":
                                            Set setMode = New CSetFVConcurrent
                                        Case "3":
                                            Set setMode = New CSetFVConcurrent
                                    End Select
                            End Select
                        Case "MV":
                            Select Case forceData.AsIParameter.AsString
                                Case "":
                                    Select Case Boardchoice
                                        Case "0":
                                            Set setMode = New CSetMVConCurrent
                                        Case "1":
                                            Set setMode = New CSetMVConCurrent
                                        Case "2":
                                            Set setMode = New CSetMVConCurrent
                                        Case "3":
                                            Set setMode = New CSetMVConCurrent
                                    End Select
                                Case Else
                                    Select Case Boardchoice
                                        Case "0":
                                            Set setMode = New CSetFIConCurrent
                                        Case "1":
                                            Set setMode = New CSetFIConCurrent
                                        Case "2":
                                            Set setMode = New CSetFIConCurrent
                                        Case "3":
                                            Set setMode = New CSetFIConCurrent
                                    End Select
                            End Select
                        Case Else
                            Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Test Mode [" & modeData.AsIParameter.AsString & "] !"
                    End Select
                Case Else '"Concurrent以外"
                    Select Case modeData.AsIParameter.AsString
                        Case "MI":
                            Select Case forceData.AsIParameter.AsString
                                Case "":
                                    Set setMode = New CSetIRange
                                Case Else
                                    Set setMode = New CSetFV
                            End Select
                        Case "MV":
                            Select Case forceData.AsIParameter.AsString
                                Case "":
                                    Set setMode = New CSetMV
                                Case Else
                                    Set setMode = New CSetFI
                            End Select
                        Case Else
                            Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Test Mode [" & modeData.AsIParameter.AsString & "] !"
                    End Select
            End Select
    End Select
          
    '### リーダーからアクションパラメータの読み込み #######
    setMode.Initialize measSite, mDcContainer, reader
    Set setModeAction = setMode
End Function
'#V21-Release
Private Function measureAction(ByVal reader As IParameterReader, Optional ByVal measSite As Long = ALL_SITE) As IAction
    '### リーダーからパラメータの読み出し #################
    Dim SourceType As CParamString
    Set SourceType = CreateCParamString
    With SourceType.AsIParameter
        .Name = TEST_PINTYPE
        .Read reader
    End With
    Dim modeData As CParamString
    Set modeData = CreateCParamString
    With modeData.AsIParameter
        .Name = SET_MODE
        .Read reader
    End With
    Dim measureSiteMode As CParamString
    Set measureSiteMode = CreateCParamString
    With measureSiteMode.AsIParameter
        .Name = MEASURE_SITE
        .Read reader
    End With
    '### アクション内容の決定 #############################
    Dim measure As IMeasure
    Select Case SourceType.AsIParameter.AsString
        Case "BPMU":
            Select Case modeData.AsIParameter.AsString
                Case "MI":
                    Set measure = New CMeasureIBPMU
                Case "MV":
                    Set measure = New CMeasureVBPMU
                Case Else
                    Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Test Mode [" & modeData.AsIParameter.AsString & "] !"
            End Select
        Case Else
            Select Case measureSiteMode.AsIParameter.AsString
                Case "Concurrent":
                    Select Case modeData.AsIParameter.AsString
                        Case "MI":
                            Set measure = New CMeasureIConcurrent
                        Case "MV":
                            Set measure = New CMeasureVConcurrent
                        Case Else
                            Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Test Mode [" & modeData.AsIParameter.AsString & "] !"
                    End Select
                Case Else
                    Select Case modeData.AsIParameter.AsString
                        Case "MI":
                            Set measure = New CMeasureI
                        Case "MV":
                            Set measure = New CMeasureV
                        Case Else
                            Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Test Mode [" & modeData.AsIParameter.AsString & "] !"
                    End Select
'                Case Else
'                    Err.Raise 9999, "CDcScenario.ConstructScenario", "This Is The Invalid Site Mode [" & measureSiteMode.AsIParameter.AsString & "] !"
            End Select
    End Select
    '### リーダーからアクションパラメータの読み込み #######
    measure.Initialize measSite, measurePinSet(reader), mDcContainer, reader
    Set measureAction = measure
End Function
'#V21-Release
Private Function disconnectAction(ByVal reader As IParameterReader, Optional ByVal measSite As Long = ALL_SITE) As IAction
    '### リーダーからパラメータの読み出し #################
    Dim SourceType As CParamString
    Set SourceType = CreateCParamString
    With SourceType.AsIParameter
        .Name = TEST_PINTYPE
        .Read reader
    End With
    '### アクション内容の決定 #############################
    Dim setOpen As IDisconnect

    Dim measureSiteMode As CParamString
    Set measureSiteMode = CreateCParamString
    With measureSiteMode.AsIParameter
        .Name = MEASURE_SITE
        .Read reader
    End With
     
    Select Case SourceType.AsIParameter.AsString
        Case "BPMU":
            Select Case measureSiteMode.AsIParameter.AsString
                Case "Concurrent":
                    Set setOpen = New CDisconnectBPMUConcurrent
                Case Else: 'concurrent以外
                    Set setOpen = New CDisconnectBPMU
            End Select
        Case Else: 'BPMU以外
            Select Case measureSiteMode.AsIParameter.AsString
                Case "Concurrent":
                    Set setOpen = New CDisConnectConcurrent
                Case Else: 'concurrent以外
                    Set setOpen = New CDisconnect
            End Select
    End Select
            
   
    '### リーダーからアクションパラメータの読み込み #######
    setOpen.Initialize measSite, mDcContainer, reader
    Set disconnectAction = setOpen
End Function

Private Function measurePinSet(ByVal reader As IParameterReader) As Collection
    Set measurePinSet = New Collection
    Dim fileStream As IActionStream
    Set fileStream = reader
    '### リーダーから測定ピンパラメータの読み込み #########
    Dim currPin As CMeasurePin
    Do While Not fileStream.IsEndOfData
        Set currPin = New CMeasurePin
        currPin.Initialize mDcContainer, reader
        currPin.ReadSpecInfo mInstanceReader
        measurePinSet.Add currPin
        fileStream.MoveNextData
    Loop
End Function
'#V21-Release
Private Function BPMUchcheck(ByVal reader As IParameterReader)
'BPMUのサイトシェアCheck
    
    'PinList取得
    Dim PinListBPMU As CParamString
    Set PinListBPMU = CreateCParamString
    With PinListBPMU.AsIParameter
        .Name = TEST_PINLIST
        .Read reader
    End With
    
    Dim SinglePinBPMU() As String 'Split(PinList,",")で切る
    Dim curSite As Long
    Dim i As Integer
    Dim j As Integer 'Site数
    Dim k As Integer 'Pin数
    Dim lp As Integer
    Dim n As Integer
    Dim g  As Integer
    Dim Channels() As Long
    Dim JudgeNumber() As Integer
    
    SinglePinBPMU = Split(PinListBPMU.AsIParameter.AsString, ",")
    k = UBound(SinglePinBPMU)
    j = CountExistSite
    ReDim JudgeNumber((j + 1) * (k + 1))
    g = -1
    For lp = 0 To k
        For i = 0 To j
            curSite = i '
            Call GetChanList(SinglePinBPMU(lp), curSite, chIO, Channels())
            'For chan = 0 To UBound(Channels)
            g = g + 1 'g++
            If Channels(0) < 64 Then
                JudgeNumber(g) = 0 'ボード0
            ElseIf (Channels(0) >= 64 And Channels(0) < 128) Then
                JudgeNumber(g) = 1 'ボード1
            ElseIf (Channels(0) >= 128 And Channels(0) < 192) Then
                JudgeNumber(g) = 2 'ボード2
            ElseIf (Channels(0) >= 192 And Channels(0) < 256) Then
                JudgeNumber(g) = 3 'ボード3
            ElseIf (Channels(0) >= 256 And Channels(0) < 320) Then
                JudgeNumber(g) = 4 'ボード4
            ElseIf (Channels(0) >= 320 And Channels(0) < 384) Then
                JudgeNumber(g) = 5 'ボード5
            ElseIf (Channels(0) >= 384 And Channels(0) < 448) Then
                JudgeNumber(g) = 6 'ボード6
            ElseIf (Channels(0) >= 448 And Channels(0) < 512) Then
                JudgeNumber(g) = 7 'ボード7
            End If
        Next i
    Next lp
    '力技
    For i = 0 To (g - 1)
        For lp = (i + 1) To g
            If JudgeNumber(i) = JudgeNumber(lp) Then
                Err.Raise 9999, " Can't use Multi Site on the one BPMU Board " & PinListBPMU.AsIParameter.AsString
            End If
        Next lp
    Next i
    
End Function
'#V21-Release
Private Function BPMUchcheckB(ByVal reader As IParameterReader)
'DC Scenario Sheetでの複数Pin登録禁止
    
    'PinList取得
    Dim PinListBPMU As CParamString
    Set PinListBPMU = CreateCParamString
    With PinListBPMU.AsIParameter
        .Name = TEST_PINLIST
        .Read reader
    End With
    
    Dim SinglePinBPMU() As String 'Split(PinList,",")で切る 'PinListはきってあげないと無理
    
    Dim k As Integer 'Pin数
    
    SinglePinBPMU = Split(PinListBPMU.AsIParameter.AsString, ",")
    k = UBound(SinglePinBPMU)
    
    If k >= 1 Then
        Err.Raise 9999, " Can't use PinGroups on the  BPMU Board "
    End If
    
End Function

