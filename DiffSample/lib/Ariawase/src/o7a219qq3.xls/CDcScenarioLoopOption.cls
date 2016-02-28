VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcScenarioLoopOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDcScenario

Private mCurrScenario As CDcCategory
Private mDcScenario As Collection

Private mSamplingMode As SAMPLING_PARAM
Private mSamplingLimit As CParamDouble
Private mSamplingCategories As Collection

Private mCurrLoopOption As Boolean
Private mWriter As CDcScenarioLoopOptionWriter

Private mFirstCategory As Boolean

Private Sub Class_Initialize()
    Set mCurrScenario = Nothing
    Set mDcScenario = Nothing
    Set mSamplingLimit = Nothing
    Set mSamplingCategories = Nothing
    Set mWriter = Nothing
    mFirstCategory = True
End Sub

Private Sub Class_Terminate()
    Set mCurrScenario = Nothing
    Set mDcScenario = Nothing
    Set mSamplingLimit = Nothing
    Set mSamplingCategories = Nothing
    Set mWriter = Nothing
End Sub

Private Sub IDcScenario_Apichoice()

End Sub

Private Sub IDcScenario_ClearContainer()

End Sub

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
    
    If mFirstCategory = True Then
        TheExec.Datalog.WriteComment "DC Scenario loop measure mode!"
        mFirstCategory = False
    End If
    
    If exist(mSamplingCategories, currTest) Then
        mCurrLoopOption = True
        Call mWriter.AsIActionStream.Rewind
    Else
       mCurrLoopOption = False
    End If
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
    If mCurrLoopOption = True Then
        Call loopExecute
    Else
        Call normalExecute
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CDcScenario.Execute", "Can Not Execute The Test Name [" & mCurrScenario.Name & "] !"

End Sub

Private Sub IDcScenario_GetTempResult(ByVal pName As String, ByRef pResult() As Double)

End Sub

Private Sub IDcScenario_Initialize(ByVal dcPerformer As IDcTest, ByVal scenarioReader As IParameterReader, ByVal instanceReader As IParameterReader, ByVal scenarioWriter As IParameterWriter, ByVal pinInfoWriter As IParameterWriter)

End Sub

Private Sub IDcScenario_Load()

End Sub

Private Property Let IDcScenario_ResultManager(ByVal resultKeeper As CResultManager)

End Property

Public Property Set Categories(ByRef pCategories As Collection)
    Set mDcScenario = pCategories
End Property

Public Sub SetLoopOption(ByRef pLoopCondition As LOOP_CONDITION)
    With pLoopCondition
        Set mSamplingCategories = .LOOP_CATEGORIES
        mSamplingMode = .Param
        
        Set mSamplingLimit = CreateCParamDouble
        If mSamplingMode = PARAM_COUNTS Then
            mSamplingLimit.AsIParameter.Name = SAMPLE_MODE_COUNTS
        Else
            mSamplingLimit.AsIParameter.Name = SAMPLE_MODE_TIME
        End If
        mSamplingLimit.AsIParameter.AsDouble = .PARAM_LIMIT
        
        Set mWriter = New CDcScenarioLoopOptionWriter
        Call mWriter.Initialize(.FOLDER)
    End With
End Sub

Private Sub normalExecute()
    Dim currGroup As Collection
    Dim actionIndex As Long
    Dim measures As Collection
    Dim subMeasure As Collection
    Dim currMeasure As IMeasure
    
    Set measures = New Collection
    For Each currGroup In mCurrScenario.Groups
        Set subMeasure = New Collection
        '### グループアクション実行開始 ###########
        For actionIndex = 1 To currGroup.Count
            Select Case currGroup.Item(actionIndex).Name
            Case MEASURE_CLASS:
                Dim tempMeasure As IMeasure
                '@@@ アクション実行 @@@@@@@
                currGroup.Item(actionIndex).Execute
                Set tempMeasure = currGroup.Item(actionIndex)
                '### MEASUREアクションの登録 ##
                subMeasure.Add tempMeasure
            Case Else
                '@@@ アクション実行 @@@@@@@@@@@
                currGroup.Item(actionIndex).Execute
            End Select
        Next actionIndex
        For Each currMeasure In subMeasure
            measures.Add currMeasure
        Next currMeasure
    Next currGroup
    '### 測定結果の格納 ###################################
    For Each currMeasure In measures
        currMeasure.WriteResult TheResult
    Next currMeasure
End Sub

Private Sub loopExecute()
    Dim currGroup As Collection
    Dim actionIndex As Long
    Dim measures As Collection
    Dim subMeasure As Collection
    Set measures = New Collection
    Set subMeasure = New Collection
    
    Dim currMeasure As IMeasure
    
    Dim measureLoop As IMeasureLoopOption
    Dim loopMeasures As Collection
    Set loopMeasures = New Collection
    
    Call categoryLogger(mCurrScenario, mWriter)
    
    For Each currGroup In mCurrScenario.Groups
        '### グループアクション実行開始 ###########
        For actionIndex = 1 To currGroup.Count
            Select Case currGroup.Item(actionIndex).Name
            Case MEASURE_CLASS:
                Set measureLoop = currGroup.Item(actionIndex)
                Call loopMeasures.Add(measureLoop)
                Set measureLoop = Nothing
            Case Else
                '@@@ LoopMeasure実行 @@@@@@@@@@@
                Call loopMeasure(loopMeasures, subMeasure)
                Set loopMeasures = New Collection
                '@@@ アクション実行 @@@@@@@@@@@
                currGroup.Item(actionIndex).Execute
            End Select
            For Each currMeasure In subMeasure
                measures.Add currMeasure
            Next currMeasure
            Set subMeasure = New Collection
        Next actionIndex
        
    Next currGroup
    '最終グループがMEASURE_CLASSの場合
    Set subMeasure = New Collection
    Call loopMeasure(loopMeasures, subMeasure)
    For Each currMeasure In subMeasure
        measures.Add currMeasure
    Next currMeasure

    Set loopMeasures = Nothing
    Set subMeasure = Nothing
    
    '### 測定結果の格納 ###################################
    For Each currMeasure In measures
        currMeasure.WriteResult TheResult
    Next currMeasure
    
    Set currMeasure = Nothing
    Set measures = Nothing
    
End Sub

Private Sub loopMeasure(ByRef pLoopMeasures As Collection, ByRef lastMeasures As Collection)
    
    If pLoopMeasures.Count = 0 Then Exit Sub
    Dim measure As IMeasureLoopOption
    
    Dim counter As Long
    Dim lastMeasureLapTime As Double
    Dim timer As CStopWatch
    Set timer = New CStopWatch
    
    'ファイルヘッダー設定
    Call limitLogger(mWriter)
    
    counter = 0
    timer.Start
    Do
        counter = counter + 1
        For Each measure In pLoopMeasures
            '@@@ アクション実行 @@@@@@@@@@@
            Call measure.Execute
            lastMeasureLapTime = timer.GetLapTime
            Call measureActionLogger(measure, counter, lastMeasureLapTime, mWriter)
        Next measure
      
    Loop While loopComplete(counter, lastMeasureLapTime) = False

    Dim lastMeasure As IMeasure
    For Each lastMeasure In pLoopMeasures
        lastMeasures.Add lastMeasure
    Next lastMeasure
        
    'ファイル出力
    Call mWriter.WriteFile(mCurrScenario.Name)
    'ライター初期化
    Call mWriter.AsIActionStream.Rewind
End Sub

Private Sub limitLogger(ByRef writer As IParameterWriter)
    Call mSamplingLimit.AsIParameter.WriteOut(writer)
'    Call writer.WriteAsString(DCLOOP_FILE_HEADER, "")
End Sub
Private Function loopComplete(ByVal pCounts As Long, ByVal pTime As Double) As Boolean
    Dim counter As Variant
    If mSamplingMode = PARAM_COUNTS Then
        counter = pCounts
    Else
        counter = pTime
    End If
    loopComplete = (counter >= mSamplingLimit.AsIParameter.AsDouble)
End Function
Private Sub categoryLogger(ByVal testCategory As CDcCategory, ByVal writer As IActionStream)
    '### テストカテゴリパラメータのダンプ #################
    If Not writer Is Nothing Then
        writer.SetCategoryLocation testCategory.Name
        paramLogger testCategory, writer
    End If
End Sub
Private Sub measureActionLogger(ByVal action As IMeasure, ByVal lCount As Long, ByVal stopTime As Double, ByVal writer As IActionStream)
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
        
        Dim ExecTime As CParamDouble
        Set ExecTime = CreateCParamDouble
        With ExecTime.AsIParameter
            .Name = EXAMIN_TIMESTAMP
            .AsDouble = stopTime
            .WriteOut writer
        End With
        measurePinLogger action, writer
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
Private Sub actionLogger(ByVal action As IAction, ByVal writer As IActionStream)
    '### アクションパラメータのダンプ #####################
    If Not writer Is Nothing Then
        writer.SetGroupLocation action.GroupIndex
        paramLogger action, writer
    End If
End Sub
Private Sub paramLogger(ByVal writable As IParameterWritable, ByVal writer As IParameterWriter)
    '### 型変換のためのプロシージャ #######################
    writable.WriteOut writer
End Sub

Private Function exist(ByRef pCollection As Collection, ByVal pKey As String) As Boolean
    On Error GoTo NOT_FOUND
    Call pCollection.Item(pKey)
    exist = True
    Exit Function
NOT_FOUND:
    exist = False
End Function
