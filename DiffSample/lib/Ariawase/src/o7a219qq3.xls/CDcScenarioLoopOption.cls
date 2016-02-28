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
'���e:
'   ���s����e�X�g�J�e�S���̌���Ǝ������[�h���̊e��ݒ���s��
'
'�p�����[�^:
'[currTest]   In  �e�X�g�C���X�^���X��
'
'���ӎ���:
'
    On Error GoTo ErrHandler
    '### �J�����g�e�X�g�̌��� #############################
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
'���e:
'   �w�肳�ꂽ�e�X�g�J�e�S���̃A�N�V���������s
'
'�p�����[�^:
'
'���ӎ���:
'   ���s�O�ɕK��SetScenario�Ŏ��s����J�e�S�������肷��K�v������
'
    On Error GoTo ErrHandler
    '### �V�i���I���s�J�n #################################
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
        '### �O���[�v�A�N�V�������s�J�n ###########
        For actionIndex = 1 To currGroup.Count
            Select Case currGroup.Item(actionIndex).Name
            Case MEASURE_CLASS:
                Dim tempMeasure As IMeasure
                '@@@ �A�N�V�������s @@@@@@@
                currGroup.Item(actionIndex).Execute
                Set tempMeasure = currGroup.Item(actionIndex)
                '### MEASURE�A�N�V�����̓o�^ ##
                subMeasure.Add tempMeasure
            Case Else
                '@@@ �A�N�V�������s @@@@@@@@@@@
                currGroup.Item(actionIndex).Execute
            End Select
        Next actionIndex
        For Each currMeasure In subMeasure
            measures.Add currMeasure
        Next currMeasure
    Next currGroup
    '### ���茋�ʂ̊i�[ ###################################
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
        '### �O���[�v�A�N�V�������s�J�n ###########
        For actionIndex = 1 To currGroup.Count
            Select Case currGroup.Item(actionIndex).Name
            Case MEASURE_CLASS:
                Set measureLoop = currGroup.Item(actionIndex)
                Call loopMeasures.Add(measureLoop)
                Set measureLoop = Nothing
            Case Else
                '@@@ LoopMeasure���s @@@@@@@@@@@
                Call loopMeasure(loopMeasures, subMeasure)
                Set loopMeasures = New Collection
                '@@@ �A�N�V�������s @@@@@@@@@@@
                currGroup.Item(actionIndex).Execute
            End Select
            For Each currMeasure In subMeasure
                measures.Add currMeasure
            Next currMeasure
            Set subMeasure = New Collection
        Next actionIndex
        
    Next currGroup
    '�ŏI�O���[�v��MEASURE_CLASS�̏ꍇ
    Set subMeasure = New Collection
    Call loopMeasure(loopMeasures, subMeasure)
    For Each currMeasure In subMeasure
        measures.Add currMeasure
    Next currMeasure

    Set loopMeasures = Nothing
    Set subMeasure = Nothing
    
    '### ���茋�ʂ̊i�[ ###################################
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
    
    '�t�@�C���w�b�_�[�ݒ�
    Call limitLogger(mWriter)
    
    counter = 0
    timer.Start
    Do
        counter = counter + 1
        For Each measure In pLoopMeasures
            '@@@ �A�N�V�������s @@@@@@@@@@@
            Call measure.Execute
            lastMeasureLapTime = timer.GetLapTime
            Call measureActionLogger(measure, counter, lastMeasureLapTime, mWriter)
        Next measure
      
    Loop While loopComplete(counter, lastMeasureLapTime) = False

    Dim lastMeasure As IMeasure
    For Each lastMeasure In pLoopMeasures
        lastMeasures.Add lastMeasure
    Next lastMeasure
        
    '�t�@�C���o��
    Call mWriter.WriteFile(mCurrScenario.Name)
    '���C�^�[������
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
    '### �e�X�g�J�e�S���p�����[�^�̃_���v #################
    If Not writer Is Nothing Then
        writer.SetCategoryLocation testCategory.Name
        paramLogger testCategory, writer
    End If
End Sub
Private Sub measureActionLogger(ByVal action As IMeasure, ByVal lCount As Long, ByVal stopTime As Double, ByVal writer As IActionStream)
    '### ���W���[�A�N�V�����p�����[�^�̃_���v #############
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
    '### ���W���[�s���p�����[�^�̃_���v ###################
    If Not writer Is Nothing Then
        Dim currPin As CMeasurePin
        For Each currPin In mPins.MeasurePins
            paramLogger currPin, writer
            writer.MoveNextData
        Next currPin
    End If
End Sub
Private Sub actionLogger(ByVal action As IAction, ByVal writer As IActionStream)
    '### �A�N�V�����p�����[�^�̃_���v #####################
    If Not writer Is Nothing Then
        writer.SetGroupLocation action.GroupIndex
        paramLogger action, writer
    End If
End Sub
Private Sub paramLogger(ByVal writable As IParameterWritable, ByVal writer As IParameterWriter)
    '### �^�ϊ��̂��߂̃v���V�[�W�� #######################
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