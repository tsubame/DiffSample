VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcScenarioSheetWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DC Test Scenarioワークシートパラメータ確認用ライタークラス
'
'目的:
'   アクションパラメータの確認用
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWriter
Implements IFileStream
Implements IActionStream

Dim mScenarioTable As Collection
Dim mCategoryIndex As Long
Dim mGroupIndex As Long
Dim mDataIndex As Long

Dim mDcScenarioSht As Worksheet
Dim mDataCell As Range

Const END_OF_RECORD = "END"
Const BOR = "B7"

Private Sub Class_Initialize()
    Set mScenarioTable = New Collection
End Sub

Private Sub Class_Terminate()
    Set mScenarioTable = Nothing
End Sub

Public Sub Initialize(ByVal SheetName As String)
    On Error GoTo ErrorHandler
    Set mDcScenarioSht = Worksheets(SheetName)
    Exit Sub
ErrorHandler:
    Err.Raise 9999, "CDcScenarioSheetWriter.Initialize", "Can Not Find The Sheet Name [" & SheetName & "] !"
End Sub

Public Sub SetGrouping()
    On Error Resume Next
    trackField.Rows.Ungroup
    Dim currTrack As Collection
    Me.AsIActionStream.Rewind
    For Each currTrack In mScenarioTable
        Dim topCell As Range
        Dim endCell As Range
        Dim currSector As Range
        Set topCell = currTrack.Item(1).Rows(1).offset(1, 0)
        For Each currSector In currTrack
            Set endCell = currSector.Rows(currSector.Rows.Count)
        Next currSector
        mDcScenarioSht.Range(topCell.Cells, endCell.Cells).Rows.group
    Next currTrack
'    mDcScenarioSht.Outline.ShowLevels RowLevels:=1
End Sub

Public Sub Validate()
    Dim currTrack As Collection
    Dim currGroup As Range
    On Error GoTo VERIFY_ERR
    Me.AsIActionStream.Rewind
    Application.ScreenUpdating = False
    For Each currTrack In mScenarioTable
        For Each currGroup In currTrack
            With currGroup
                ValidateActionParameter .Rows(1)
                ValidateActionParameter .Rows(.Rows.Count)
            End With
        Next currGroup
    Next currTrack
    Application.ScreenUpdating = True
    Exit Sub
VERIFY_ERR:
    Application.ScreenUpdating = True
End Sub

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    Err.Raise 9999, "CDcScenarioSheetWriter.IFileStream.SetLocation", "Does Not Support This Method !"
End Sub

Private Sub IFileStream_MoveNext()
    Set mDataCell = mDataCell.offset(1, 0)
End Sub

Private Sub IFileStream_Rewind()
    Set mDataCell = mDcScenarioSht.Range(BOR)
End Sub

Private Function IFileStream_IsEOR() As Boolean
    Err.Raise 9999, "CDcScenarioSheetWriter.IFileStream.IsEOR", "Does Not Support This Method !"
End Function

Private Sub IParameterWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
    Select Case dataName:
        Case EXAMIN_FLAG:
            dataCell(dataName).Value = UCase(CStr(blnData))
        Case DATA_CHANGED:
            Dim shtObject As Object
            Set shtObject = mDcScenarioSht
            shtObject.IsChanged = blnData
        Case IS_VALIDATE:
            mDcScenarioSht.Range("E1").Value = blnData
        Case Else
            Err.Raise 9999, "CDcScenarioSheetWriter.AsIParameterWriter.WriteAsBoolean", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Sub

Private Sub IParameterWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
    Err.Raise 9999, "CDcScenarioSheetWriter.AsIParameterWriter.WriteAsDouble", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
    Err.Raise 9999, "CDcScenarioSheetWriter.AsIParameterWriter.WriteAsInteger", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
    Select Case dataName
        Case MEASURE_AVG:
            dataCell(dataName).Value = lngData
        Case Else
            Err.Raise 9999, "CDcScenarioSheetWriter.AsIParameterWriter.WriteAsLong", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Sub

Private Sub IParameterWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
    Select Case dataName:
        Case TEST_PINS, MEASURE_LABEL, OPERATE_RESULT, TEST_CATEGORY, EXAMIN_MODE, TEST_ACTION, TEST_POSTACTION, TEST_PINTYPE, SET_MODE, SET_RANGE, SET_FORCE, MEASURE_WAIT, MEASURE_SITE, OPERATE_FORCE:
            dataCell(dataName).Value = strData
        Case EXAMIN_RESULT:
            Dim dataList As New Collection
            Dim currData As Variant
            Dim dataIndex As Long
            Set dataList = DecomposeStringList(strData)
            Application.EnableEvents = False
            For Each currData In dataList
                dataCell(dataName, dataIndex).Value = currData
                dataIndex = dataIndex + 1
            Next currData
            Application.EnableEvents = True
        Case EXAMIN_RESULTUNIT, EXAMIN_EXECTIME:
            Application.EnableEvents = False
            dataCell(dataName).Value = strData
            Application.EnableEvents = True
        Case Else
            Err.Raise 9999, "CDcScenarioSheetWriter.AsIParameterWriter.WriteAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Sub

Private Function dataCell(ByVal dataName As String, Optional ByVal offsetCol As Long = 0) As Range
    With mDcScenarioSht
        Set dataCell = .Cells(mDataCell.Row, .Range(dataName).offset(0, offsetCol).Column)
    End With
End Function

Private Sub IActionStream_SetCategoryLocation(ByVal dataName As String)
    mCategoryIndex = 1
    mGroupIndex = 1
    mDataIndex = 1
    Dim currCategory As Collection
    For Each currCategory In mScenarioTable
        Set mDataCell = mScenarioTable.Item(mCategoryIndex).Item(mGroupIndex).Rows(mDataIndex)
        If Replace(dataCell(TEST_CATEGORY).Value, " ", "") = Replace(dataName, " ", "") Then
            Exit For
        End If
        mCategoryIndex = mCategoryIndex + 1
    Next currCategory
    If Me.AsIActionStream.IsEndOfCategory Then
        Err.Raise 9999, "CDcScenarioSheetWriter.IActionStream.SetCategoryLocation", "Can Not Find The Test Category [" & dataName & "] !"
    End If
    Set mDataCell = mScenarioTable.Item(mCategoryIndex).Item(mGroupIndex).Rows(mDataIndex)
End Sub

Private Sub IActionStream_SetGroupLocation(ByVal GroupIndex As Long)
    mGroupIndex = GroupIndex + 1
    mDataIndex = 1
    Set mDataCell = mScenarioTable.Item(mCategoryIndex).Item(mGroupIndex).Rows(mDataIndex)
End Sub

Private Sub IActionStream_SetDataLocation(ByVal dataIndex As Long)
    mDataIndex = dataIndex + 1
    Set mDataCell = mScenarioTable.Item(mCategoryIndex).Item(mGroupIndex).Rows(mDataIndex)
End Sub

Private Sub IActionStream_MoveNextCategory()
    mCategoryIndex = mCategoryIndex + 1
    mGroupIndex = 1
    mDataIndex = 1
    If Me.AsIActionStream.IsEndOfCategory Then Exit Sub
    Set mDataCell = mScenarioTable.Item(mCategoryIndex).Item(mGroupIndex).Rows(mDataIndex)
End Sub

Private Sub IActionStream_MoveNextGroup()
    mGroupIndex = mGroupIndex + 1
    mDataIndex = 1
    If Me.AsIActionStream.IsEndOfGroup Then Exit Sub
    Set mDataCell = mScenarioTable.Item(mCategoryIndex).Item(mGroupIndex).Rows(mDataIndex)
End Sub

Private Sub IActionStream_MoveNextData()
    mDataIndex = mDataIndex + 1
    If Me.AsIActionStream.IsEndOfData Then Exit Sub
    Set mDataCell = mScenarioTable.Item(mCategoryIndex).Item(mGroupIndex).Rows(mDataIndex)
End Sub

Private Function IActionStream_IsEndOfCategory() As Boolean
    IActionStream_IsEndOfCategory = (mCategoryIndex > mScenarioTable.Count)
'    If IActionStream_IsEndOfCategory Then
'        Dim shtObject As Object
'        Set shtObject = mDcScenarioSht
'        shtObject.IsChanged = False
'    End If
End Function

Private Function IActionStream_IsEndOfGroup() As Boolean
    IActionStream_IsEndOfGroup = (mGroupIndex > mScenarioTable.Item(mCategoryIndex).Count)
End Function

Private Function IActionStream_IsEndOfData() As Boolean
    IActionStream_IsEndOfData = (mDataIndex > mScenarioTable.Item(mCategoryIndex).Item(mGroupIndex).Rows.Count)
End Function

Private Sub IActionStream_Rewind()
    Set mScenarioTable = searchTrack
    mCategoryIndex = 1
    mGroupIndex = 1
    mDataIndex = 1
    Set mDataCell = mScenarioTable.Item(mCategoryIndex).Item(mGroupIndex).Rows(mDataIndex)
End Sub

Private Sub IActionStream_SeekChangedLocation()
    Err.Raise 9999, "CDcScenarioSheetWriter.IActionStream.SeekChangedLocation", "Does Not Support This Method !"
End Sub

Private Sub showFormatErrorCell(ByVal wsheet As Worksheet, ByVal errCell As Range)
    wsheet.Activate
    errCell.Select
End Sub

Public Function AsIParameterWriter() As IParameterWriter
    Set AsIParameterWriter = Me
End Function

Public Function AsIFileStream() As IFileStream
    Set AsIFileStream = Me
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
    Err.Raise 9999, "CDcScenarioSheetWriter.Search", "Category Data Format Error On [" & mDcScenarioSht.Name & "] !"
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
    Err.Raise 9999, "CDcScenarioSheetWriter.Search", "Measure Pin Data Format Error On [" & mDcScenarioSht.Name & "] !"
End Function

Private Function trackField() As Range
    Dim topCell As Range
    Dim endCell As Range
    Set trackField = Nothing
    With mDcScenarioSht
        Set topCell = .Range(TEST_CATEGORY).offset(1, 0)
        If IsEmpty(topCell) Then
            showFormatErrorCell mDcScenarioSht, topCell
            Err.Raise 9999, "CDcScenarioSheetWriter.Search", "Can Not Find Top Of Category Data In [" & .Name & "] !"
        End If
        Set endCell = .Columns(topCell.Column).Find(END_OF_RECORD)
        If endCell Is Nothing Then
            Err.Raise 9999, "CDcScenarioSheetWriter.Search", "Can Not Find End Of Category Data In [" & .Name & "] !"
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
            Err.Raise 9999, "CDcScenarioSheetWriter.Search", "Can Not Find Top Of Measure Pin Data In [" & .Name & "] !"
        End If
        Set endCell = .Cells(currTrack.Rows(currTrack.Count).Row, .Range(TEST_PINS).Column)
        If Not IsEmpty(endCell) Then
            showFormatErrorCell mDcScenarioSht, endCell
            Err.Raise 9999, "CDcScenarioSheetWriter.Search", "End Of Measure Pin Data Format Error In [" & .Name & "] !"
        End If
        Set sectionField = .Range(topCell.Cells, endCell.Cells)
    End With
End Function
