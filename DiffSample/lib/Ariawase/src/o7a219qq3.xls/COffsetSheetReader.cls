VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COffsetSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   Offset Managerワークシート用リーダークラス
'
'目的:
'   Offset Managerワークシートに定義されたパラメータを読み込む
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterReader
Implements IFileStream

Dim mOffsetSht As Worksheet
Dim mDataCell As Range
Dim mEorCell As Range
Dim mMaxSite As Long
Dim mTesterIndex As Long

Public Sub Initialize(ByVal SheetName As String, ByVal testerNum As Long, ByVal maxSiteNum As Long)
    On Error GoTo ErrorHandler
    Set mOffsetSht = Worksheets(SheetName)
    mMaxSite = maxSiteNum
    Me.AsIFileStream.Rewind
    updataEor
    searchTester testerNum
    Exit Sub
ErrorHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "COffsetSheetReader.Initialize", "Can Not Initialize The Offset Manager Sheet Reader !"
End Sub

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    With mOffsetSht
        Set mDataCell = .Columns(.Range(OFFSET_LABEL).Column).Find(dataName)
    End With
    If mDataCell Is Nothing Then Err.Raise 9999, "COffsetSheetReader.AsIFileStream.SetLocation", "Can Not Find The TEST Label [" & dataName & "] !"
    Exit Sub
End Sub

Private Sub IFileStream_MoveNext()
    Set mDataCell = mDataCell.offset(mMaxSite + 1, 0)
End Sub

Private Sub IFileStream_Rewind()
    Set mDataCell = mOffsetSht.Range(OFFSET_LABEL).offset(1, 0)
    If IsEmpty(mDataCell) Then
        showFormatErrorCell mOffsetSht, mDataCell
        Err.Raise 9999, "COffsetSheetReader.AsIFileStream.Rewind", "Can Not Find The Beginning Of Data !"
    End If
End Sub

Private Function IFileStream_IsEOR() As Boolean
    IFileStream_IsEOR = (mDataCell.Row > mEorCell.Row)
    If IFileStream_IsEOR Then
        Dim shtObject As Object
        Set shtObject = mOffsetSht
        shtObject.IsChanged = False
    End If
End Function

Private Function IParameterReader_ReadAsBoolean(ByVal dataName As String) As Boolean
    Select Case dataName:
        Case DATA_CHANGED:
            Dim shtObject As Object
            Set shtObject = mOffsetSht
            IParameterReader_ReadAsBoolean = shtObject.IsChanged
        Case Else
            Err.Raise 9999, "COffsetSheetReader.AsIParameterReader.ReadAsBoolean", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Function IParameterReader_ReadAsDouble(ByVal dataName As String) As Double
    Err.Raise 9999, "COffsetSheetReader.AsIParameterReader.ReadAsDouble", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsInteger(ByVal dataName As String) As Integer
    Err.Raise 9999, "COffsetSheetReader.AsIParameterReader.ReadAsInteger", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsLong(ByVal dataName As String) As Long
    Err.Raise 9999, "COffsetSheetReader.AsIParameterReader.ReadAsLong", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsString(ByVal dataName As String) As String
    Select Case dataName
        Case OFFSET_LABEL:
            IParameterReader_ReadAsString = mDataCell.Value
        Case OFFSET_COEF, OFFSET_CONS:
            Dim dataList As String
            Dim dataIndex As Long
            For dataIndex = 0 To mMaxSite
                Dim tempdata As String
                If IsEmpty(dataCell(dataName).offset(dataIndex, 0)) Then
                    tempdata = NOT_DEFINE
                Else
                    tempdata = dataCell(dataName).offset(dataIndex, 0).Value
                End If
                If dataIndex = 0 Then
                    dataList = tempdata
                Else
                    dataList = dataList & "," & tempdata
                End If
            Next dataIndex
            IParameterReader_ReadAsString = dataList
        Case Else
            Err.Raise 9999, "COffsetSheetReader.AsIParameterReader.ReadAsString", "Can Not Find The Parameter Label [" & dataName & "] !"
    End Select
End Function

Private Function dataCell(ByVal dataName As String) As Range
    With mOffsetSht
        Set dataCell = .Cells(mDataCell.Row, .Range(dataName).offset(0, mTesterIndex).Column)
    End With
End Function

Private Sub updataEor()
    Dim currCell As Range
    Set currCell = mDataCell
    Do While Not IsEmpty(currCell)
        If mMaxSite > 0 Then
            Dim DataRange As Range
            Set DataRange = mOffsetSht.Range(currCell.offset(1, 0).Cells, currCell.offset(mMaxSite, 0).Cells)
            Dim dataCell As Range
            For Each dataCell In DataRange
                If Not IsEmpty(dataCell) Then
                    showFormatErrorCell mOffsetSht, dataCell
                    Err.Raise 9999, "COffsetSheetReader.UpdataEor", "Data Format Error On [" & mOffsetSht.Name & "] !"
                End If
            Next dataCell
        End If
        Set currCell = currCell.offset(mMaxSite + 1, 0)
    Loop
    Set mEorCell = currCell.offset(-1, 0)
End Sub

Private Sub searchTester(ByVal testerNum As Long)
'2010/12/7 0020207264 修正
    Dim currCell As Range
    Dim currCol As Range
    Dim dataIndex As Long
    mTesterIndex = 0
    With mOffsetSht
        Set currCol = .Range(TESTER_NUMBER)
        Do While currCol.Column < .Range(END_OF_TESTER_NUM).Column
            Set currCell = .Cells(.Range(TESTER_NUMBER).offset(1, 0).Row, currCol.Column)
            If currCell.Value = testerNum Then
                Exit Do
            End If
            Set currCol = currCol.offset(0, 3)
        Loop
        If currCol.Column > .Range(END_OF_TESTER_NUM).Column Then
            Err.Raise 9999, "COffsetSheetReader.SearchTester", "Can Not Find The Tester Number [" & testerNum & "] !"
        End If
        mTesterIndex = dataIndex
    End With
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
