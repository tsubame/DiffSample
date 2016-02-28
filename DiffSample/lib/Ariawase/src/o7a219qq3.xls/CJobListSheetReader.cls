VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CJobListSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   IG-XL Data Tool [Job List]用のワークシートリーダークラス
'
'目的:
'   Job Listシートで定義されるIG-XL Data Toolシート名を取得する
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterReader
Implements IFileStream

Const JOB_NAME = "B4"
Const PIN_MAP = "C4"
Const TEST_INSTANCE = "D4"
Const FLOW_TABLE = "E4"
Const AC_SPEC = "F4"
Const DC_SPEC = "G4"
Const PAT_SET = "H4"
Const PAT_GRP = "I4"
Const BIN_TABEL = "J4"
Const SHMOO_TABLE = "K4"
Const TEST_PROCEDURE = "L4"

Dim mJobListSheet As Worksheet
Dim mDataCell As Range
Dim mEorCell As Range

Public Sub Initialize(ByVal SheetName As String)
    On Error GoTo ErrorHandler
    Set mJobListSheet = Worksheets(SheetName)
    Me.AsIFileStream.Rewind
    updataEor
    Exit Sub
ErrorHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CJobListSheetReader.Initialize", "Can Not Initialize The Job List Sheet Reader !"
End Sub

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    With mJobListSheet
        Set mDataCell = .Columns(.Range(JOB_NAME).Column).Find(dataName)
    End With
    If mDataCell Is Nothing Then Err.Raise 9999, "CJobListSheetReader.AsIFileStream.SetLocation", "Can Not Find The JOB Name [" & dataName & "] !"
End Sub

Private Sub IFileStream_MoveNext()
    Set mDataCell = mDataCell.offset(1, 0)
End Sub

Private Sub IFileStream_Rewind()
    Set mDataCell = mJobListSheet.Range(JOB_NAME).offset(1, 0)
    If IsEmpty(mDataCell) Then Err.Raise 9999, "CJobListSheetReader.AsIFileStream.Rewind", "Can Not Find The Beginning Of Data !"
End Sub

Private Function IFileStream_IsEOR() As Boolean
    IFileStream_IsEOR = (mDataCell.Row > mEorCell.Row)
End Function

Private Function IParameterReader_ReadAsBoolean(ByVal dataName As String) As Boolean
    Err.Raise 9999, "CJobListSheetReader.AsIParameterReader.ReadAsBoolean", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsDouble(ByVal dataName As String) As Double
    Err.Raise 9999, "CJobListSheetReader.AsIParameterReader.ReadAsDouble", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsInteger(ByVal dataName As String) As Integer
    Err.Raise 9999, "CJobListSheetReader.AsIParameterReader.ReadAsInteger", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsLong(ByVal dataName As String) As Long
    Err.Raise 9999, "CJobListSheetReader.AsIParameterReader.ReadAsLong", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsString(ByVal dataName As String) As String
    Select Case dataName
        Case ACTIVE_JOB_NAME:
            IParameterReader_ReadAsString = mDataCell.Value
        Case PIN_MAP_TOOL:
            IParameterReader_ReadAsString = dataCell(PIN_MAP).Value
        Case TEST_INSTANCES_TOOL:
            IParameterReader_ReadAsString = dataCell(TEST_INSTANCE).Value
        Case FLOW_TABLE_TOOL:
            IParameterReader_ReadAsString = dataCell(FLOW_TABLE).Value
        Case AC_SPECS_TOOL:
            IParameterReader_ReadAsString = dataCell(AC_SPEC).Value
        Case DC_SPECS_TOOL:
            IParameterReader_ReadAsString = dataCell(DC_SPEC).Value
        Case PATTERN_SETS_TOOL:
            IParameterReader_ReadAsString = dataCell(PAT_SET).Value
        Case PATTERN_GROUPS_TOOL:
            IParameterReader_ReadAsString = dataCell(PAT_GRP).Value
        Case BIN_TABLE_TOOL:
            IParameterReader_ReadAsString = dataCell(BIN_TABEL).Value
        Case CHARACTERIZATION_TOOL:
            IParameterReader_ReadAsString = dataCell(SHMOO_TABLE).Value
        Case TEST_PROCEDURES_TOOL:
            IParameterReader_ReadAsString = dataCell(TEST_PROCEDURE).Value
        Case Else
           Err.Raise 9999, "CJobListSheetReader.AsIParameterReader.ReadAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Function dataCell(ByVal dataName As String) As Range
    With mJobListSheet
        Set dataCell = .Cells(mDataCell.Row, .Range(dataName).Column)
    End With
End Function

Private Sub updataEor()
    Dim endCell As Range
    Set endCell = mDataCell.End(xlDown)
    If WorksheetFunction.CountA(mJobListSheet.Range(mDataCell.Cells, endCell.Cells)) <= 1 Then
        Set mEorCell = mDataCell
    Else
        Set mEorCell = endCell
    End If
End Sub

Public Function AsIParameterReader() As IParameterReader
    Set AsIParameterReader = Me
End Function

Public Function AsIFileStream() As IFileStream
    Set AsIFileStream = Me
End Function
