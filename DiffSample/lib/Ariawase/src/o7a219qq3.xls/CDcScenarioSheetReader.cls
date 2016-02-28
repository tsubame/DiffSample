VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcScenarioSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DC Test Scenarioワークシート用リーダークラス
'
'目的:
'   DC Test Scenarioワークシートに定義されたパラメータを読み込む
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterReader
Implements IActionStream

Dim mDcScenarioSht As Worksheet
Dim mScenarioTable As Collection
Dim mCurrScenario As Collection
Dim mCurrGroup As Range
Dim mCurrCell As Range
Dim mCategoryIndex As Long
Dim mGroupIndex As Long
Dim mDataIndex As Long

Const END_OF_RECORD = "END"

Private Sub Class_Initialize()
    Set mScenarioTable = New Collection
    Set mCurrScenario = New Collection
End Sub

Private Sub Class_Terminate()
    Set mScenarioTable = Nothing
    Set mCurrScenario = Nothing
End Sub

Public Sub Initialize(ByVal SheetName As String)
    On Error GoTo ErrorHandler
    Set mDcScenarioSht = Worksheets(SheetName)
    Me.AsIActionStream.Rewind
    Exit Sub
ErrorHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CDcScenarioSheetReader.Initialize", "Can Not Initialize The Dc Scenario Sheet Reader !"
End Sub

Private Function IParameterReader_ReadAsBoolean(ByVal dataName As String) As Boolean
    Select Case dataName:
        Case EXAMIN_FLAG:
            IParameterReader_ReadAsBoolean = dataCell(mCurrScenario.Item(1), dataName).Value
        Case DATA_CHANGED:
            Dim shtObject As Object
             Set shtObject = mDcScenarioSht
            IParameterReader_ReadAsBoolean = shtObject.IsChanged
        Case IS_VALIDATE:
            IParameterReader_ReadAsBoolean = mDcScenarioSht.Range("E1").Value
        Case Else
            Err.Raise 9999, "CDcScenarioSheetReader.AsIParameterReader.ReadAsBoolean", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Function IParameterReader_ReadAsDouble(ByVal dataName As String) As Double
    Err.Raise 9999, "CDcScenarioSheetReader.AsIParameterReader.ReadAsDouble", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsInteger(ByVal dataName As String) As Integer
    Err.Raise 9999, "CDcScenarioSheetReader.AsIParameterReader.ReadAsInteger", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsLong(ByVal dataName As String) As Long
    Select Case dataName
        Case CATEGORY_ID:
            IParameterReader_ReadAsLong = mCategoryIndex - 1
        Case GROUP_ID:
            IParameterReader_ReadAsLong = mGroupIndex - 1
        Case MEASURE_AVG:
            IParameterReader_ReadAsLong = dataCell(mCurrGroup, dataName).Value
        Case Else
            Err.Raise 9999, "CDcScenarioSheetReader.AsIParameterReader.ReadAsLong", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Function IParameterReader_ReadAsString(ByVal dataName As String) As String
    Select Case dataName:
        Case TEST_PINLIST:
            Dim currCell As Range
            Dim PinList As New Collection
            For Each currCell In mCurrGroup.Resize(mCurrGroup.Rows.Count - 1)
                PinList.Add currCell.Value
            Next currCell
            IParameterReader_ReadAsString = ComposeStringList(PinList)
        Case TEST_PINS:
            IParameterReader_ReadAsString = mCurrCell.Value
        Case MEASURE_LABEL:
            If IsEmpty(dataCell(mCurrCell, dataName)) Then
                IParameterReader_ReadAsString = NOT_DEFINE
            Else
                IParameterReader_ReadAsString = dataCell(mCurrCell, dataName).Value
            End If
        Case OPERATE_RESULT:
            IParameterReader_ReadAsString = dataCell(mCurrCell, dataName).Value
        Case TEST_CATEGORY, EXAMIN_MODE:
            IParameterReader_ReadAsString = dataCell(mCurrScenario.Item(1), dataName).Value
        Case TEST_POSTACTION:
            IParameterReader_ReadAsString = dataCell(mCurrGroup.Rows(mCurrGroup.Rows.Count), TEST_ACTION).Value
        Case TEST_ACTION, TEST_PINTYPE, SET_MODE, SET_RANGE, SET_FORCE, MEASURE_WAIT, MEASURE_SITE, OPERATE_FORCE:
            IParameterReader_ReadAsString = dataCell(mCurrGroup, dataName).Value
        Case Else
            Err.Raise 9999, "CDcScenarioSheetReader.AsIParameterReader.ReadAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Sub IActionStream_SetCategoryLocation(ByVal dataName As String)
    Dim currCategory As Collection
    mCategoryIndex = 1
    For Each currCategory In mScenarioTable
        If Replace(dataCell(currCategory.Item(1), TEST_CATEGORY).Value, " ", "") = Replace(dataName, " ", "") Then
            Exit For
        End If
        mCategoryIndex = mCategoryIndex + 1
    Next currCategory
    If Me.AsIActionStream.IsEndOfCategory Then
        Err.Raise 9999, "CDcScenarioSheetReader.IActionStream.SetCategoryLocation", "Can Not Find The Test Category [" & dataName & "] !"
    End If
    Set mCurrScenario = mScenarioTable.Item(mCategoryIndex)
    mGroupIndex = 1
    mDataIndex = 1
    Set mCurrGroup = mCurrScenario.Item(mGroupIndex)
    Set mCurrCell = mCurrGroup.Rows(mDataIndex)
End Sub

Private Sub IActionStream_SetGroupLocation(ByVal GroupIndex As Long)
    mGroupIndex = GroupIndex + 1
    Set mCurrGroup = mCurrScenario.Item(mGroupIndex)
    mDataIndex = 1
    Set mCurrCell = mCurrGroup.Rows(mDataIndex)
End Sub

Private Sub IActionStream_SetDataLocation(ByVal dataIndex As Long)
    mDataIndex = dataIndex + 1
    Set mCurrCell = mCurrGroup.Rows(mDataIndex)
End Sub

Private Sub IActionStream_MoveNextCategory()
    mCategoryIndex = mCategoryIndex + 1
    If Me.AsIActionStream.IsEndOfCategory Then Exit Sub
    Set mCurrScenario = mScenarioTable.Item(mCategoryIndex)
    mGroupIndex = 1
    mDataIndex = 1
    Set mCurrGroup = mCurrScenario.Item(mGroupIndex)
    Set mCurrCell = mCurrGroup.Rows(mDataIndex)
End Sub

Private Sub IActionStream_MoveNextGroup()
    mGroupIndex = mGroupIndex + 1
    If Me.AsIActionStream.IsEndOfGroup Then Exit Sub
    Set mCurrGroup = mCurrScenario.Item(mGroupIndex)
    mDataIndex = 1
    Set mCurrCell = mCurrGroup.Rows(mDataIndex)
End Sub

Private Sub IActionStream_MoveNextData()
    mDataIndex = mDataIndex + 1
    If Me.AsIActionStream.IsEndOfData Then Exit Sub
    Set mCurrCell = mCurrGroup.Rows(mDataIndex)
End Sub

Private Function IActionStream_IsEndOfCategory() As Boolean
    IActionStream_IsEndOfCategory = (mCategoryIndex > mScenarioTable.Count)
End Function

Private Function IActionStream_IsEndOfGroup() As Boolean
    IActionStream_IsEndOfGroup = (mGroupIndex > mCurrScenario.Count)
End Function

Private Function IActionStream_IsEndOfData() As Boolean
    IActionStream_IsEndOfData = (mDataIndex > mCurrGroup.Rows.Count - 1)
End Function

Private Sub IActionStream_Rewind()
    Set mScenarioTable = searchTrack
    mCategoryIndex = 1
    mGroupIndex = 1
    mDataIndex = 1
    Set mCurrScenario = mScenarioTable.Item(mCategoryIndex)
    Set mCurrGroup = mCurrScenario.Item(mGroupIndex)
    Set mCurrCell = mCurrGroup.Rows(mDataIndex)
    Dim shtObject As Object
    Set shtObject = mDcScenarioSht
    shtObject.IsChanged = False
End Sub

Private Sub IActionStream_SeekChangedLocation()
    Dim shtObject As Object
    Set shtObject = mDcScenarioSht
    Dim currScenario As Collection
    Dim currGroup As Range
    mCategoryIndex = 0
    For Each currScenario In mScenarioTable
        mCategoryIndex = mCategoryIndex + 1
        mGroupIndex = 0
        For Each currGroup In currScenario
            mGroupIndex = mGroupIndex + 1
            If currGroup.Rows(1).Row <= shtObject.ChangedCell.Row And _
                currGroup.Rows(currGroup.Rows.Count).Row > shtObject.ChangedCell.Row Then GoTo FindCell
        Next currGroup
    Next currScenario
FindCell:
    mDataIndex = 1
    Set mCurrScenario = mScenarioTable.Item(mCategoryIndex)
    Set mCurrGroup = mCurrScenario.Item(mGroupIndex)
    Set mCurrCell = mCurrGroup.Rows(mDataIndex)
    With shtObject
        .ChangedCell = Nothing
        .IsChanged = False
    End With
End Sub

Private Function dataCell(ByVal refCell As Range, ByVal dataName As String) As Range
    With mDcScenarioSht
        Set dataCell = .Cells(refCell.Rows(1).Row, .Range(dataName).Column)
    End With
End Function

Private Sub showFormatErrorCell(ByVal wsheet As Worksheet, ByVal errCell As Range)
    wsheet.Activate
    errCell.Select
End Sub

Public Function AsIParameterReader() As IParameterReader
    Set AsIParameterReader = Me
End Function

Public Function AsIActionStream() As IActionStream
    Set AsIActionStream = Me
End Function

Private Function searchTrack() As Collection
    Dim allTracks As Range
    Set allTracks = trackField
    Dim currCell As Range
    Dim topCell As Range
    Dim endCell As Range
    Dim currTrack As Range
    Set searchTrack = New Collection
    With allTracks
        Set topCell = mDcScenarioSht.Cells(.Row, .Column)
    End With
    For Each currCell In allTracks
        If currCell.Row <> topCell.Row And Not IsEmpty(currCell) Then
            Set endCell = currCell.offset(-1, 0)
            Set currTrack = mDcScenarioSht.Range(topCell.Cells, endCell.Cells)
            If currTrack.Count < 2 Then
                GoTo formatError
            Else
                searchTrack.Add sectionSearch(currTrack), topCell.Value
            End If
            Set topCell = currCell
        End If
    Next currCell
    Exit Function
formatError:
    Set searchTrack = Nothing
    showFormatErrorCell mDcScenarioSht, endCell
    Err.Raise 9999, "CDcScenarioSheetReader.Search", "Category Data Format Error On [" & mDcScenarioSht.Name & "] !"
End Function

Private Function sectionSearch(ByVal currTrack As Range) As Collection
    Dim dataField As Range
    Set dataField = sectionField(currTrack)
    Dim currCell As Range
    Dim topCell As Range
    Dim endCell As Range
    Dim currField As Range
    Set sectionSearch = New Collection
    With dataField
        Set topCell = mDcScenarioSht.Cells(.Row, .Column)
    End With
    For Each currCell In dataField
        If IsEmpty(currCell) Then
            If IsEmpty(topCell) Then GoTo formatError
            Set endCell = currCell
            Set currField = mDcScenarioSht.Range(topCell.Cells, endCell.Cells)
            sectionSearch.Add currField
            Set topCell = currCell.offset(1, 0)
        End If
    Next currCell
    Exit Function
formatError:
    Set sectionSearch = Nothing
    showFormatErrorCell mDcScenarioSht, topCell
    Err.Raise 9999, "CDcScenarioSheetReader.Search", "Measure Pin Data Format Error On [" & mDcScenarioSht.Name & "] !"
End Function

Private Function trackField() As Range
    Dim topCell As Range
    Dim endCell As Range
    Set trackField = Nothing
    With mDcScenarioSht
        Set topCell = .Range(TEST_CATEGORY).offset(1, 0)
        If IsEmpty(topCell) Then
            showFormatErrorCell mDcScenarioSht, topCell
            Err.Raise 9999, "CDcScenarioSheetReader.Search", "Can Not Find Top Of Category Data In [" & .Name & "] !"
        End If
        Set endCell = .Columns(topCell.Column).Find(END_OF_RECORD)
        If endCell Is Nothing Then
            Err.Raise 9999, "CDcScenarioSheetReader.Search", "Can Not Find End Of Category Data In [" & .Name & "] !"
        End If
        Set trackField = .Range(topCell.Cells, endCell.Cells)
    End With
End Function

Private Function sectionField(ByVal currTrack As Range) As Range
    Dim topCell As Range
    Dim endCell As Range
    Set sectionField = Nothing
    With mDcScenarioSht
        Set topCell = .Cells(currTrack.Row, .Range(TEST_PINS).Column)
        If IsEmpty(topCell) Then
            showFormatErrorCell mDcScenarioSht, topCell
            Err.Raise 9999, "CDcScenarioSheetReader.Search", "Can Not Find Top Of Measure Pin Data In [" & .Name & "] !"
        End If
        Set endCell = .Cells(currTrack.Rows(currTrack.Count).Row, .Range(TEST_PINS).Column)
        If Not IsEmpty(endCell) Then
            showFormatErrorCell mDcScenarioSht, endCell
            Err.Raise 9999, "CDcScenarioSheetReader.Search", "End Of Measure Pin Data Format Error In [" & .Name & "] !"
        End If
        Set sectionField = .Range(topCell.Cells, endCell.Cells)
    End With
End Function

