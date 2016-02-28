VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcLogReportWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   DCテスト測定レンジ適正判定結果出力用ライタークラス
'
'目的:
'   各ピンが持つプロパティやハードウェアから取得した設定情報
'   及び判定結果をCSVファイル形式で書き出す
'
'作成者:
'   0145206097
'
Option Explicit

#Const HSD200_USE = 1               'HSD200ボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

Implements IParameterWriter
Implements IFileStream

Const WORK_FILE = "ValidationWorkFile"

Dim mParamData As Collection
Dim mMainLabel As Collection
Dim mSubLabel As Collection
Dim mUnit As String
Dim mFileName As String

Private Const RETURNED_MEAS_RANGE_ICUL1G As Double = 1#

#If HSD200_USE = 0 Then
Private Const RETURNED_MEAS_RANGE_PPMU As Double = 2#
#Else
Private Const RETURNED_MEAS_RANGE_PPMU As Double = 1.5
#End If

Private Sub Class_Initialize()
    Set mMainLabel = New Collection
    With mMainLabel
        .Add MEASURE_LABEL
        .Add TEST_PINS
        .Add BOARD_NAME
        .Add BOARD_FORCE
        .Add BOARD_RANGE
        .Add USERMACRO_LOLIMIT
        .Add USERMACRO_HILIMIT
        .Add USERMACRO_JUDGE
        .Add VALIDATE_RESULT
        .Add EXAMIN_RESULT
    End With
    Set mSubLabel = New Collection
    With mSubLabel
        .Add "Label", MEASURE_LABEL
        .Add "Pin Name", TEST_PINS
        .Add "Board Name", BOARD_NAME
        .Add "Range", BOARD_RANGE
        .Add "Force", BOARD_FORCE
        .Add "Low", USERMACRO_LOLIMIT
        .Add "High", USERMACRO_HILIMIT
        .Add "Judge", USERMACRO_JUDGE
        .Add "Validate Result", VALIDATE_RESULT
        .Add "Result", EXAMIN_RESULT
    End With
    Set mParamData = New Collection
End Sub

Private Sub Class_Terminate()
    Set mMainLabel = Nothing
    Set mSubLabel = Nothing
    Set mParamData = Nothing
End Sub

Public Sub Initialize()
    Dim logFileNo As Integer
    On Error GoTo ErrorHandler
    logFileNo = FreeFile
    mFileName = getCurrentDir & "\" & WORK_FILE
    Open mFileName For Output As #logFileNo
    Dim currLabel As Variant
    For Each currLabel In mMainLabel
        Print #logFileNo, mSubLabel(currLabel) & ",";
    Next currLabel
    Print #logFileNo, ""
    Close #logFileNo
    Exit Sub
ErrorHandler:
    Close #logFileNo
    Err.Raise 9999, "CDcLogReportWriter.Initialize", "Opening The File Is Aborted!"
End Sub

Private Function IFileStream_IsEOR() As Boolean
    Dim newFileName As String
    newFileName = Application.GetSaveAsFilename(InitialFileName:=makeFileName, fileFilter:="CSV File,*.csv,Txt File,*.txt,All Files,*.*", title:="Save DC Range Validation Log File")
    If newFileName = "False" Then
        Kill mFileName
    Else
        Name mFileName As newFileName
    End If
    Set mParamData = New Collection
End Function

Private Sub IFileStream_SetLocation(ByVal dataName As String)
End Sub

Private Sub IFileStream_MoveNext()
    Dim logFileNo As Integer
    logFileNo = FreeFile
    Open mFileName For Append As #logFileNo
    Dim currData As Variant
    For Each currData In mMainLabel
        Print #logFileNo, mParamData(currData) & ",";
    Next currData
    Print #logFileNo, "EOR"
    Close #logFileNo
    Set mParamData = New Collection
End Sub

Private Sub IFileStream_Rewind()
    Set mParamData = New Collection
End Sub

Private Sub IParameterWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
    Dim MainUnit As String
    Dim SubUnit As String
    Dim SubValue As Double
    Select Case dataName
        Case USERMACRO_LOLIMIT, USERMACRO_HILIMIT:
            Select Case mUnit
                Case NOT_DEFINE:
                   mParamData.Add NOT_DEFINE, dataName
                Case Else
                    SplitUnitValue "999" & mUnit, MainUnit, SubUnit, SubValue
                    mParamData.Add dblData / SubUnitToValue(SubUnit) & SubUnit & MainUnit, dataName
            End Select
        Case BOARD_RANGE:
            If dblData = INVALIDATION_VALUE Then
                mParamData.Add "Disable", dataName
            Else
                Select Case mUnit
                'PPMUのMVモードはレンジ概念がないので（TheDcクラスからのGetMeasureRangeの取得結果は2V[HSD100]or1.5V[HSD200]となる）、
                'レンジ出力値は”None”固定に変更
                Case NOT_DEFINE:
                    'ここでは単位が未定義でMVモードかどうかの判断がつかないので
                    '力技で2V|1.5Vを意味する「2|1.5」の決め打ちで判断させている（2A|1.5Aは存在しないので）
                    '単位取得方法はいずれ変える必要がある
                    If dblData = RETURNED_MEAS_RANGE_PPMU And mParamData.Item(BOARD_NAME) = "PPMU" Then
                        mParamData.Add "None", dataName
                    ElseIf dblData = RETURNED_MEAS_RANGE_ICUL1G And mParamData.Item(BOARD_NAME) = "ICUL1G" Then
                        mParamData.Add "None", dataName
                    Else
                        mParamData.Add dblData, dataName
                    End If
                Case Else
                    SplitUnitValue "999" & mUnit, MainUnit, SubUnit, SubValue
                    If MainUnit = "V" And (mParamData.Item(BOARD_NAME) = "PPMU" Or mParamData.Item(BOARD_NAME) = "ICUL1G") Then
                        mParamData.Add "None", dataName
                    Else
                        NormalizeUnitValue dblData, SubUnit, SubValue
                        mParamData.Add SubValue & SubUnit & MainUnit, dataName
                    End If
                End Select
            End If
        Case BOARD_FORCE:
            If dblData = INVALIDATION_VALUE Then
                mParamData.Add "Disable", dataName
            Else
                Select Case mUnit
                Case NOT_DEFINE:
                    mParamData.Add dblData, dataName
                Case Else
                    Dim FUnit As String
                    Select Case Right(mUnit, 1)
                        Case "A": FUnit = "V"
                        Case "V": FUnit = "A"
                    End Select
                    NormalizeUnitValue dblData, SubUnit, SubValue
                    mParamData.Add SubValue & SubUnit & FUnit, dataName
                End Select
            End If
    End Select
End Sub

Private Sub IParameterWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
    Select Case dataName
        Case TEST_PINS, MEASURE_LABEL:
            mParamData.Add strData, dataName
        Case BOARD_NAME:
            mParamData.Add Replace(strData, "dc", ""), dataName
        Case USERMACRO_UNIT:
            mUnit = strData
        Case EXAMIN_RESULT:
            Dim dataList As New Collection
            Dim currData As Variant
            Set dataList = DecomposeStringList(strData)
            Dim covList As New Collection
            Select Case mUnit
                Case NOT_DEFINE:
                    For Each currData In dataList
                        covList.Add CDbl(currData)
                    Next currData
                Case Else
                    Dim MainUnit As String
                    Dim SubUnit As String
                    Dim SubValue As Double
                    SplitUnitValue "999" & mUnit, MainUnit, SubUnit, SubValue
                    For Each currData In dataList
                        NormalizeUnitValue currData, SubUnit, SubValue
                        covList.Add SubValue & SubUnit & MainUnit
                    Next currData
            End Select
            mParamData.Add ComposeStringList(covList), dataName
    End Select
End Sub

Private Sub IParameterWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
    Select Case dataName
        Case VALIDATE_RESULT:
            Select Case lngData
                Case VALIDATE_NG:
                    mParamData.Add "NG", dataName
                Case VALIDATE_NG_NO_JUDGE:
                    mParamData.Add "NG[NoBin]", dataName
                Case VALIDATE_WARNING:
                    mParamData.Add "WARNING", dataName
                Case VALIDATE_WARNING_NO_JUDGE:
                    mParamData.Add "WARNING[NoBin]", dataName
                Case VALIDATE_OK:
                    mParamData.Add "OK", dataName
                Case VALIDATE_OK_NO_JUDGE:
                    mParamData.Add "OK[NoBin]", dataName
                Case NO_JUDGE:
                    mParamData.Add "Disable", dataName
                Case DISABEL_TO_VALIDATION:
                    mParamData.Add "Disable", dataName
            End Select
        Case USERMACRO_JUDGE:
            Select Case mUnit
                Case NOT_DEFINE:
                   mParamData.Add NOT_DEFINE, dataName
                Case Else
                    mParamData.Add lngData, dataName
            End Select
    End Select
End Sub

Private Sub IParameterWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
End Sub

Private Sub IParameterWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
End Sub

Public Function AsIParameterWriter() As IParameterWriter
    Set AsIParameterWriter = Me
End Function

Public Function AsIFileStream() As IFileStream
    Set AsIFileStream = Me
End Function

Private Sub NormalizeUnitValue(ByVal dblData As Double, ByRef SubUnit As String, ByRef SubValue As Double)
    If Abs(dblData) >= 10 ^ 9 Then
        SubUnit = "G"
        SubValue = dblData / 10 ^ 9
    ElseIf Abs(dblData) >= 10 ^ 6 Then
        SubUnit = "M"
        SubValue = dblData / 10 ^ 6
    ElseIf Abs(dblData) >= 10 ^ 3 Then
        SubUnit = "k"
        SubValue = dblData / 10 ^ 3
    ElseIf Abs(dblData) >= 1 Then
        SubUnit = ""
        SubValue = dblData
    ElseIf Abs(dblData) = 0 Then
        SubUnit = ""
        SubValue = dblData
    ElseIf Abs(dblData) * 10 ^ 3 >= 1 Then
        SubUnit = "m"
        SubValue = dblData * 10 ^ 3
    ElseIf Abs(dblData) * 10 ^ 6 >= 1 Then
        SubUnit = "u"
        SubValue = dblData * 10 ^ 6
    ElseIf Abs(dblData) * 10 ^ 9 >= 1 Then
        SubUnit = "n"
        SubValue = dblData * 10 ^ 9
    ElseIf Abs(dblData) * 10 ^ 12 >= 1 Then
        SubUnit = "p"
        SubValue = dblData * 10 ^ 12
    End If
End Sub

Private Function makeFileName() As String
    makeFileName = getCurrentDir & "\" & getToday & "_" & Replace(ActiveWorkbook.Name, ".xls", "") & "_RangeValidateResult" & ".csv"
End Function

Private Function getToday() As String
    getToday = Format(DateTime.Now, "yymmdd") & "_" & Format(DateTime.Now, "hhnnss")
End Function

Private Function getCurrentDir() As String
    getCurrentDir = ActiveWorkbook.Path
End Function
