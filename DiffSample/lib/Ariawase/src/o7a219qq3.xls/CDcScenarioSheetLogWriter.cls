VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcScenarioSheetLogWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DC Test Scenarioワークシート実験用ライタークラス
'
'目的:
'   DC Test Scenarioワークシートに各実験データを書き出す
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWriter
Implements IActionStream

Dim mScenarioTable As Collection
Dim mCategoryIndex As Long
Dim mGroupIndex As Long
Dim mDataIndex As Long

Dim mDcScenarioSht As Worksheet
Dim mDataCell As Range

Const END_OF_RECORD = "END"

Private Sub Class_Initialize()
    Set mScenarioTable = New Collection
End Sub

Private Sub Class_Terminate()
    Set mScenarioTable = Nothing
End Sub

Public Sub Initialize(ByVal SheetName As String, ByVal maxSiteNum As Long)
    On Error GoTo ErrorHandler
    Set mDcScenarioSht = Worksheets(SheetName)
    expandResultColumn maxSiteNum
    Me.AsIActionStream.Rewind
    Exit Sub
ErrorHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CDcScenarioSheetLogWriter.Initialize", "Can Not Initialize The Dc Scenario Sheet Log Writer !"
End Sub

Public Sub ShowCurrAction(ByVal actionID As Long)
    mDcScenarioSht.Activate
    If actionID = NO_TEST Then
        dataCell(TEST_CATEGORY).Select
    Else
        mScenarioTable.Item(mCategoryIndex).Item(actionID + 1).Select
    End If
End Sub

Private Sub IParameterWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
    Select Case dataName
        Case EXAMIN_EXECTIME:
            Application.EnableEvents = False
            dataCell(dataName).Value = dblData
            Application.EnableEvents = True
    End Select
End Sub

Private Sub IParameterWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
    Select Case dataName
        Case USERMACRO_UNIT:
            Application.EnableEvents = False
            dataCell(EXAMIN_RESULTUNIT).Value = strData
            Application.EnableEvents = True
        Case EXAMIN_RESULT:
            Application.EnableEvents = False
            Dim dataList As New Collection
            Dim currData As Variant
            Dim dataIndex As Long
            Set dataList = DecomposeStringList(strData)
            Select Case dataCell(EXAMIN_RESULTUNIT).Value
                Case NOT_DEFINE:
                    For Each currData In dataList
                        dataCell(dataName, dataIndex).Value = CDbl(currData)
                        dataIndex = dataIndex + 1
                    Next currData
                Case Else:
                Dim MainUnit As String
                Dim SubUnit As String
                Dim SubValue As Double
                SplitUnitValue "999" & dataCell(EXAMIN_RESULTUNIT).Value, MainUnit, SubUnit, SubValue
                For Each currData In dataList
                    dataCell(dataName, dataIndex).Value = CDbl(currData) / SubUnitToValue(SubUnit)
                    dataIndex = dataIndex + 1
                Next currData
            End Select
            Application.EnableEvents = True
    End Select
End Sub

Private Sub IParameterWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
End Sub

Private Sub IParameterWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
End Sub

Private Sub IParameterWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
End Sub

Private Function dataCell(ByVal dataName As String, Optional ByVal offsetCol As Long = 0) As Range
    With mDcScenarioSht
        Set dataCell = .Cells(mDataCell.Row, .Range(dataName).offset(0, offsetCol).Column)
    End With
End Function

Private Sub expandResultColumn(ByVal maxSiteNum As Long)
    Dim resultCell As Range
    Dim unitCell As Range
    With mDcScenarioSht
        Set resultCell = .Range(EXAMIN_RESULT)
        Set unitCell = .Range(EXAMIN_RESULTUNIT)
    End With
    Dim dataIndex As Long
    For dataIndex = 0 To maxSiteNum
        If unitCell.Column = resultCell.Column + dataIndex Then
            resultCell.offset(0, dataIndex).EntireColumn.Insert
            resultCell.offset(0, dataIndex).Value = "Result " & dataIndex
        End If
    Next dataIndex
End Sub

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
        Err.Raise 9999, "CDcScenarioSheetLogWriter.IActionStream.SetCategoryLocation", "Can Not Find The Test Category [" & dataName & "] !"
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
End Sub

Private Sub showFormatErrorCell(ByVal wsheet As Worksheet, ByVal errCell As Range)
    wsheet.Activate
    errCell.Select
End Sub

Public Function AsIParameterWriter() As IParameterWriter
    Set AsIParameterWriter = Me
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
    Err.Raise 9999, "CDcScenarioSheetLogWriter.Search", "Category Data Format Error On [" & mDcScenarioSht.Name & "] !"
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
    Err.Raise 9999, "CDcScenarioSheetLogWriter.Search", "Measure Pin Data Format Error On [" & mDcScenarioSht.Name & "] !"
End Function

Private Function trackField() As Range
    Dim topCell As Range
    Dim endCell As Range
    Set trackField = Nothing
    With mDcScenarioSht
        Set topCell = .Range(TEST_CATEGORY).offset(1, 0)
        If IsEmpty(topCell) Then
            showFormatErrorCell mDcScenarioSht, topCell
            Err.Raise 9999, "CDcScenarioSheetLogWriter.Search", "Can Not Find Top Of Category Data In [" & .Name & "] !"
        End If
        Set endCell = .Columns(topCell.Column).Find(END_OF_RECORD)
        If endCell Is Nothing Then
            Err.Raise 9999, "CDcScenarioSheetLogWriter.Search", "Can Not Find End Of Category Data In [" & .Name & "] !"
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
            Err.Raise 9999, "CDcScenarioSheetLogWriter.Search", "Can Not Find Top Of Measure Pin Data In [" & .Name & "] !"
        End If
        Set endCell = .Cells(currTrack.Rows(currTrack.Count).Row, .Range(TEST_PINS).Column)
        If Not IsEmpty(endCell) Then
            showFormatErrorCell mDcScenarioSht, endCell
            Err.Raise 9999, "CDcScenarioSheetLogWriter.Search", "End Of Measure Pin Data Format Error In [" & .Name & "] !"
        End If
        Set sectionField = .Range(topCell.Cells, endCell.Cells)
    End With
End Function
