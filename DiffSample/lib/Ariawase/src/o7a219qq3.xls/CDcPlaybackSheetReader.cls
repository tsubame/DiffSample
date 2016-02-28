VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcPlaybackSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DC Playback Dataワークシート用リーダークラス
'
'目的:
'   DC Playback Dataワークシートに定義されたパラメータを読み込む
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterReader
Implements IFileStream

Dim mDcPlaybackSht As Worksheet
Dim mDataCell As Range
Dim mEorCell As Range
Dim mMaxSite As Long

Public Sub Initialize(ByVal SheetName As String, ByVal maxSiteNum As Long)
    On Error GoTo ErrorHandler
    Set mDcPlaybackSht = Worksheets(SheetName)
    mMaxSite = maxSiteNum
    Me.AsIFileStream.Rewind
    Exit Sub
ErrorHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CDcPlaybackSheetReader.Initialize", "Can Not Initialize The Dc Playback Data Sheet Reader !"
End Sub

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    With mDcPlaybackSht
        Set mDataCell = .Columns(.Range(PB_LABEL).Column).Find(dataName)
    End With
    If mDataCell Is Nothing Then Err.Raise 9999, "CDcPlaybackSheetReader.AsIFileStream.SetLocation", "Can Not Find The TEST Label [" & dataName & "] !"
End Sub

Private Sub IFileStream_MoveNext()
    Set mDataCell = mDataCell.offset(1, 0)
End Sub

Private Sub IFileStream_Rewind()
    Set mDataCell = mDcPlaybackSht.Range(PB_LABEL).offset(1, 0)
    updataEor
    If IsEmpty(mDataCell) Then
        showFormatErrorCell mDcPlaybackSht, mDataCell
        Err.Raise 9999, "CDcPlaybackSheetReader.AsIFileStream.Rewind", "Can Not Find The Beginning Of Data !"
    End If
End Sub

Private Function IFileStream_IsEOR() As Boolean
    IFileStream_IsEOR = (mDataCell.Row > mEorCell.Row)
    If IFileStream_IsEOR Then
        Dim shtObject As Object
        Set shtObject = mDcPlaybackSht
        shtObject.IsChanged = False
    End If
End Function

Private Function IParameterReader_ReadAsBoolean(ByVal dataName As String) As Boolean
    Select Case dataName:
        Case DATA_CHANGED:
            Dim shtObject As Object
            Set shtObject = mDcPlaybackSht
            IParameterReader_ReadAsBoolean = shtObject.IsChanged
        Case Else
            Err.Raise 9999, "CDcPlaybackSheetReader.AsIParameterReader.ReadAsBoolean", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Function IParameterReader_ReadAsDouble(ByVal dataName As String) As Double
    Err.Raise 9999, "CDcPlaybackSheetReader.AsIParameterReader.ReadAsDouble", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsInteger(ByVal dataName As String) As Integer
    Err.Raise 9999, "CDcPlaybackSheetReader.AsIParameterReader.ReadAsInteger", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsLong(ByVal dataName As String) As Long
    Err.Raise 9999, "CDcPlaybackSheetReader.AsIParameterReader.ReadAsLong", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsString(ByVal dataName As String) As String
    Select Case dataName
        Case PB_LABEL:
            IParameterReader_ReadAsString = mDataCell.Value
        Case PB_LIMIT_HI, PB_LIMIT_LO, PB_REF_DATA:
            IParameterReader_ReadAsString = dataCell(dataName).Value
        Case PB_DELTA_DATA:
            Dim dataList As New Collection
            Dim dataIndex As Long
            For dataIndex = 0 To mMaxSite
                Dim tempdata As String
                If IsEmpty(dataCell(dataName).offset(0, dataIndex)) Then
                    tempdata = NOT_DEFINE
                Else
                    tempdata = dataCell(dataName).offset(0, dataIndex).Value
                End If
                dataList.Add tempdata
            Next dataIndex
            IParameterReader_ReadAsString = ComposeStringList(dataList)
        Case Else
            Err.Raise 9999, "CDcPlaybackSheetReader.AsIParameterReader.ReadAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Function dataCell(ByVal dataName As String) As Range
    With mDcPlaybackSht
        Set dataCell = .Cells(mDataCell.Row, .Range(dataName).Column)
    End With
End Function

Private Sub updataEor()
    Dim endCell As Range
    Set endCell = mDataCell.End(xlDown)
    If WorksheetFunction.CountA(mDcPlaybackSht.Range(mDataCell.Cells, endCell.Cells)) <= 1 Then
        Set mEorCell = mDataCell
    Else
        Set mEorCell = endCell
    End If
End Sub

Private Sub showFormatErrorCell(ByVal wsheet As Worksheet, ByVal errCell As Range)
    wsheet.Activate
    errCell.Select
End Sub

Public Function AsIParameterReader() As IParameterReader
    Set AsIParameterReader = Me
End Function

Public Function AsIFileStream() As IFileStream
    Set AsIFileStream = Me
End Function
