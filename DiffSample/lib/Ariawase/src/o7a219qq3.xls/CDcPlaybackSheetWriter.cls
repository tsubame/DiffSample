VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcPlaybackSheetWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DC Playback Dataワークシート用ライタークラス
'
'目的:
'   DC Playback Dataワークシートに各パラメータを書き出す
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWriter
Implements IFileStream

Dim mDcPlaybackSht As Worksheet
Dim mDataCell As Range
Dim mEorCell As Range

Public Sub Initialize(ByVal SheetName As String)
    On Error GoTo ErrorHandler
    Set mDcPlaybackSht = Worksheets(SheetName)
    Me.AsIFileStream.Rewind
    Exit Sub
ErrorHandler:
    Err.Raise 9999, "CDcPlaybackSheetWriter.Initialize", "Can Not Initialize The Dc Playback Data Sheet Writer !"
End Sub

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    With mDcPlaybackSht
        Set mDataCell = .Columns(.Range(PB_LABEL).Column).Find(dataName)
    End With
    If mDataCell Is Nothing Then Err.Raise 9999, "CDcPlaybackSheetWriter.AsIFileStream.SetLocation", "Can Not Find The TEST Label [" & dataName & "] !"
End Sub

Private Sub IFileStream_MoveNext()
    Set mDataCell = mDataCell.offset(1, 0)
End Sub

Private Sub IFileStream_Rewind()
    Set mDataCell = mDcPlaybackSht.Range(PB_LABEL).offset(1, 0)
End Sub

Private Function IFileStream_IsEOR() As Boolean
    Err.Raise 9999, "CDcPlaybackSheetWriter.IFileStream.IsEOR", "Does Not Support This Method !"
End Function

Private Sub IParameterWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
    Err.Raise 9999, "CDcPlaybackSheetWriter.AsIParameterWriter.WriteAsBoolean", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
    Err.Raise 9999, "CDcPlaybackSheetWriter.AsIParameterWriter.WriteAsDouble", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
    Err.Raise 9999, "CDcPlaybackSheetWriter.AsIParameterWriter.WriteAsInteger", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
    Err.Raise 9999, "CDcPlaybackSheetWriter.AsIParameterWriter.WriteAsLong", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
    Select Case dataName
        Case PB_LABEL:
            mDataCell.Value = strData
        Case PB_CATEGORY, PB_LIMIT_HI, PB_LIMIT_LO, PB_REF_DATA:
            dataCell(dataName).Value = strData
        Case Else
            Err.Raise 9999, "CDcPlaybackSheetWriter.AsIParameterWriter.WriteAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Sub

Private Function dataCell(ByVal dataName As String) As Range
    With mDcPlaybackSht
        Set dataCell = .Cells(mDataCell.Row, .Range(dataName).Column)
    End With
End Function

Public Function AsIParameterWriter() As IParameterWriter
    Set AsIParameterWriter = Me
End Function

Public Function AsIFileStream() As IFileStream
    Set AsIFileStream = Me
End Function

Public Sub ClearCells()
    clearCell PB_CATEGORY
    clearCell PB_LABEL
    clearCell PB_LIMIT_HI
    clearCell PB_LIMIT_LO
End Sub

Private Sub clearCell(ByVal dataName As String)
    With mDcPlaybackSht
        .Range(.Cells(dataCell(dataName).Row, dataCell(dataName).Column), .Cells(65536, .Range(dataName).Column)).ClearContents
    End With
End Sub
