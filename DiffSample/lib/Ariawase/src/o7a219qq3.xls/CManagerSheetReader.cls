VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CManagerSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   Project作成ワークシートの制御ワークシート [Manager Sheet]用 リーダークラス
'
'目的:
'   Managerシートで定義されるProjectシート名を取得する
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterReader
Implements IFileStream

Dim mManagerSheet As Worksheet
Dim mDataCell As Range
Dim mEorCell As Range

Public Sub Initialize(ByVal SheetName As String)
    On Error GoTo ErrorHandler
    Set mManagerSheet = Worksheets(SheetName)
    Me.AsIFileStream.Rewind
    updataEor
    Exit Sub
ErrorHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CManagerSheetReader.Initialize", "Can Not Initialize The Manager Sheet Reader !"
End Sub

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    With mManagerSheet
        Set mDataCell = .Columns(.Range(Replace(ACTIVE_JOB_NAME, " ", "")).Column).Find(dataName)
    End With
    If mDataCell Is Nothing Then Err.Raise 9999, "CManagerSheetReader.AsIFileStream.SetLocation", "Can Not Find The JOB Name [" & dataName & "] !"
End Sub

Private Sub IFileStream_MoveNext()
    Set mDataCell = mDataCell.offset(1, 0)
End Sub

Private Sub IFileStream_Rewind()
    Set mDataCell = mManagerSheet.Range(Replace(ACTIVE_JOB_NAME, " ", "")).offset(1, 0)
    If IsEmpty(mDataCell) Then Err.Raise 9999, "CManagerSheetReader.AsIFileStream.Rewind", "Can Not Find The Beginning Of Data !"
End Sub

Private Function IFileStream_IsEOR() As Boolean
    IFileStream_IsEOR = (mDataCell.Row > mEorCell.Row)
End Function

Private Function IParameterReader_ReadAsBoolean(ByVal dataName As String) As Boolean
    Err.Raise 9999, "CManagerSheetReader.AsIParameterReader.ReadAsBoolean", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsDouble(ByVal dataName As String) As Double
    Err.Raise 9999, "CManagerSheetReader.AsIParameterReader.ReadAsDouble", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsInteger(ByVal dataName As String) As Integer
    Err.Raise 9999, "CManagerSheetReader.AsIParameterReader.ReadAsInteger", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsLong(ByVal dataName As String) As Long
    Err.Raise 9999, "CManagerSheetReader.AsIParameterReader.ReadAsLong", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsString(ByVal dataName As String) As String
    Select Case dataName
        Case ACTIVE_JOB_NAME:
            IParameterReader_ReadAsString = mDataCell.Value
        Case DC_SCENARIO_TOOL, DC_PLAYBACK_TOOL, OFFSET_TOOL:
            IParameterReader_ReadAsString = dataCell(Replace(dataName, " ", "")).Value
        Case Else
           Err.Raise 9999, "CManagerSheetReader.AsIParameterReader.ReadAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Function dataCell(ByVal dataName As String) As Range
    With mManagerSheet
        Set dataCell = .Cells(mDataCell.Row, .Range(dataName).Column)
    End With
End Function

Private Sub updataEor()
    Dim endCell As Range
    Set endCell = mDataCell.End(xlDown)
    If WorksheetFunction.CountA(mManagerSheet.Range(mDataCell.Cells, endCell.Cells)) <= 1 Then
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
