VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcTextFileLogWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DCテスト実験用ライタークラス
'
'目的:
'   テキストファイルに各実験データを書き出す
'
'注意:
'   Datalogウインドウを開いていると
'   ファイルのリネーム時に「パス名が不正です」とのエラーになるかもしれない
'   その場合はPrintの度にOpen→Print→Closeをすれば回避できる可能性がある
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWriter
Implements IActionStream

Const WORK_FILE = "ResultLogWorkFile"
Dim mFileNo As Integer
Dim mFileName As String

Dim mCategoryName As String
Dim mGroupIndex As Long
Dim mCounter As Long
Dim mSiteMode As Long
Dim mTime As Double
Dim mAvgNum As Long
Dim mWait As String
Dim mParamData As Collection

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    Me.AsIActionStream.Rewind
    mFileNo = FreeFile
    mFileName = getCurrentDir & "\" & WORK_FILE
    Open mFileName For Output As mFileNo
    Exit Sub
ErrorHandler:
    Close mFileNo
    Err.Raise 9999, "CDcTextFileLogWriter.Initialize", "Opening The File Is Aborted!"
End Sub

Private Sub Class_Terminate()
    Close mFileNo
    Dim newFileName As String
    newFileName = Application.GetSaveAsFilename(InitialFileName:=makeFileName, fileFilter:="CSV File,*.csv,Txt File,*.txt,All Files,*.*", title:="Save DC Log File")
    If newFileName = "False" Then
        Kill mFileName
    Else
        Name mFileName As newFileName
    End If
    Set mParamData = Nothing
End Sub

Private Sub IParameterWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
    Select Case dataName:
        Case EXAMIN_TIMESTAMP:
            mTime = dblData
        Case Else
    End Select
End Sub

Private Sub IParameterWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
    Select Case dataName:
        Case TEST_PINS, MEASURE_LABEL, USERMACRO_UNIT:
            mParamData.Add strData, dataName
        Case TEST_CATEGORY:
            mCategoryName = strData
        Case EXAMIN_RESULT:
            Dim dataList As New Collection
            Dim currData As Variant
            Dim dataIndex As Long
            Set dataList = DecomposeStringList(strData)
            Select Case mParamData.Item(USERMACRO_UNIT)
                Case NOT_DEFINE:
                    For Each currData In dataList
                        printData CDbl(currData), dataIndex
                        dataIndex = dataIndex + 1
                    Next currData
                Case Else
                    Dim MainUnit As String
                    Dim SubUnit As String
                    Dim SubValue As Double
                    SplitUnitValue "999" & mParamData.Item(USERMACRO_UNIT), MainUnit, SubUnit, SubValue
                    For Each currData In dataList
                        printData CDbl(currData) / SubUnitToValue(SubUnit), dataIndex
                        dataIndex = dataIndex + 1
                    Next currData
            End Select
        Case MEASURE_WAIT:
            mWait = strData
        Case Else
    End Select
End Sub

Private Sub IParameterWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
    Select Case dataName:
        Case MEASURE_AVG:
            mAvgNum = lngData
        Case REPEAT_COUNTER:
            mCounter = lngData
            If mCounter = 1 Then
                Print #mFileNo, "TEST_CATEGORY"; ","; mCategoryName
                Print #mFileNo, "MEASURE_GROUP"; ","; mGroupIndex
                Print #mFileNo, "AVERAGING"; ","; mAvgNum
                Print #mFileNo, "REFERENCE_WAIT"; ","; mWait
                Print #mFileNo,
                Print #mFileNo, "TIME[S],COUNTS,SITE,PIN_NAME,LABEL,RESULT,UNIT"
            End If
        Case SITE_INDEX:
            mSiteMode = lngData
        Case Else
    End Select
End Sub

Private Sub IParameterWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
End Sub

Private Sub IParameterWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
End Sub

Private Sub IActionStream_SetCategoryLocation(ByVal dataName As String)
    mCategoryName = dataName
End Sub

Private Sub IActionStream_SetGroupLocation(ByVal GroupIndex As Long)
    mGroupIndex = GroupIndex
End Sub

Private Sub IActionStream_SetDataLocation(ByVal dataIndex As Long)
End Sub

Private Sub IActionStream_MoveNextCategory()
End Sub

Private Sub IActionStream_MoveNextGroup()
End Sub

Private Sub IActionStream_MoveNextData()
    Set mParamData = New Collection
End Sub

Private Sub IActionStream_Rewind()
    Set mParamData = New Collection
    mCounter = 0
End Sub

Private Sub IActionStream_SeekChangedLocation()
End Sub

Private Function IActionStream_IsEndOfCategory() As Boolean
End Function

Private Function IActionStream_IsEndOfGroup() As Boolean
End Function

Private Function IActionStream_IsEndOfData() As Boolean
End Function

Public Function AsIParameterWriter() As IParameterWriter
    Set AsIParameterWriter = Me
End Function

Public Function AsIActionStream() As IActionStream
    Set AsIActionStream = Me
End Function

Private Sub printData(ByVal resultValue As Double, ByVal dataIndex As Long)
    With mParamData
        Print #mFileNo, Format(mTime, "#0.00000");
        Print #mFileNo, ","; mCounter;
        Print #mFileNo, ","; dataIndex;
        Print #mFileNo, ","; .Item(TEST_PINS);
        Print #mFileNo, ","; .Item(MEASURE_LABEL);
        Print #mFileNo, ","; resultValue;
        Print #mFileNo, ","; .Item(USERMACRO_UNIT)
    End With
End Sub

Private Function makeFileName() As String
    Dim siteMode As String
    If mSiteMode = ALL_SITE Then
        siteMode = "ALL"
    Else
        siteMode = CStr(mSiteMode)
    End If
    makeFileName = getCurrentDir & "\" & getToday & "_" & Replace(mCategoryName, " ", "") & "_PinGroup" & mGroupIndex & "_Site" & siteMode & ".csv"
End Function

Private Function getToday() As String
    getToday = Format(DateTime.Now, "yymmdd") & "_" & Format(DateTime.Now, "hhnnss")
End Function

Private Function getCurrentDir() As String
    getCurrentDir = ActiveWorkbook.Path
End Function
