VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInstanceSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   IG-XL Data Tool [Test Instance Sheet]用のワークシートリーダークラス
'
'目的:
'   Test Instanceシートで定義されるIG-XL Data Toolシート名と、
'   Test Procedure Typeでユーザーマクロで指定している場合のスペック等のパラメータを取得する
'   ただしIG-XLテンプレートを使用している場合にこのパラメータを読み出すとエラーになる
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterReader
Implements IFileStream

Const T_NAME = "B4"
Const T_TYPE = "C4"
Const DC_CAT = "F4"
Const DC_SEL = "G4"
Const AC_CAT = "H4"
Const AC_SEL = "I4"
Const T_SETS = "J4"
Const E_SETS = "K4"
Const PIN_LEV = "L4"
Const LO_LIMIT = "N4"
Const HI_LIMIT = "O4"
Const LIMIT_JUDGE = "P4"
Const LIMIT_UNIT = "Q4"
Const USER_MACRO = "Other"

Dim mInstanceSht As Worksheet
Dim mDataCell As Range
Dim mEorCell As Range

Public Sub Initialize(ByVal SheetName As String)
    On Error GoTo ErrorHandler
    Set mInstanceSht = Worksheets(SheetName)
    Me.AsIFileStream.Rewind
    updataEor
    Exit Sub
ErrorHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "CInstanceSheetReader.Initialize", "Can Not Initialize The Test Instance Sheet Reader !"
End Sub

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    With mInstanceSht
        Set mDataCell = .Columns(.Range(T_NAME).Column).Find(dataName)
    End With
    If mDataCell Is Nothing Then Err.Raise 9999, "CInstanceSheetReader.AsIFileStream.SetLocation", "Can Not Find The JOB Name [" & dataName & "] !"
End Sub

Private Sub IFileStream_MoveNext()
    Set mDataCell = mDataCell.offset(1, 0)
End Sub

Private Sub IFileStream_Rewind()
    Set mDataCell = mInstanceSht.Range(T_NAME).offset(1, 0)
    If IsEmpty(mDataCell) Then Err.Raise 9999, "CInstanceSheetReader.AsIFileStream.Rewind", "Can Not Find The Beginning Of Data !"
End Sub

Private Function IFileStream_IsEOR() As Boolean
    IFileStream_IsEOR = (mDataCell.Row > mEorCell.Row)
End Function

Private Function IParameterReader_ReadAsBoolean(ByVal dataName As String) As Boolean
    Err.Raise 9999, "CInstanceSheetReader.AsIParameterReader.ReadAsBoolean", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsDouble(ByVal dataName As String) As Double
    Select Case dataName
        Case USERMACRO_HILIMIT:
            If isTemplate Then GoTo IsTemplateError
            IParameterReader_ReadAsDouble = dataCell(HI_LIMIT).Value
        Case USERMACRO_LOLIMIT:
            If isTemplate Then GoTo IsTemplateError
            IParameterReader_ReadAsDouble = dataCell(LO_LIMIT).Value
        Case Else
            Err.Raise 9999, "CInstanceSheetReader.AsIParameterReader.ReadAsDouble", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
    Exit Function
IsTemplateError:
    Err.Raise 9999, "CInstanceSheetReader.AsIParameterReader.ReadAsDouble", "Template Does Not Support The Parameter [" & dataName & "] !"
End Function

Private Function IParameterReader_ReadAsInteger(ByVal dataName As String) As Integer
    Err.Raise 9999, "CInstanceSheetReader.AsIParameterReader.ReadAsInteger", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsLong(ByVal dataName As String) As Long
    Select Case dataName
        Case USERMACRO_JUDGE:
            If isTemplate Then GoTo IsTemplateError
            IParameterReader_ReadAsLong = dataCell(LIMIT_JUDGE).Value
        Case Else
            Err.Raise 9999, "CInstanceSheetReader.AsIParameterReader.ReadAsLong", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
    Exit Function
IsTemplateError:
    Err.Raise 9999, "CInstanceSheetReader.AsIParameterReader.ReadAsLong", "Template Does Not Support The Parameter [" & dataName & "] !"
End Function

Private Function IParameterReader_ReadAsString(ByVal dataName As String) As String
    Select Case dataName
        Case TEST_NAME:
            IParameterReader_ReadAsString = mDataCell.Value
        Case DC_CATEGORY:
            IParameterReader_ReadAsString = dataCell(DC_CAT).Value
        Case DC_SELCTOR:
            IParameterReader_ReadAsString = dataCell(DC_SEL).Value
        Case AC_CATEGORY:
            IParameterReader_ReadAsString = dataCell(AC_CAT).Value
        Case AC_SELCTOR:
            IParameterReader_ReadAsString = dataCell(AC_SEL).Value
        Case TIME_SETS:
            IParameterReader_ReadAsString = dataCell(T_SETS).Value
        Case EDGE_SETS:
            IParameterReader_ReadAsString = dataCell(E_SETS).Value
        Case PIN_LEVELS:
            IParameterReader_ReadAsString = dataCell(PIN_LEV).Value
        Case USERMACRO_UNIT:
            If isTemplate Then GoTo IsTemplateError
            IParameterReader_ReadAsString = dataCell(LIMIT_UNIT).Value
        Case Else
            Err.Raise 9999, "CInstanceSheetReader.AsIParameterReader.ReadAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
    Exit Function
IsTemplateError:
    Err.Raise 9999, "CInstanceSheetReader.AsIParameterReader.ReadAsString", "IG-XL Template Does Not Support The Parameter [" & dataName & "] !"
End Function

Private Function dataCell(ByVal dataName As String) As Range
    With mInstanceSht
        Set dataCell = .Cells(mDataCell.Row, .Range(dataName).Column)
    End With
End Function

Private Sub updataEor()
    Dim endCell As Range
    Set endCell = mDataCell.End(xlDown)
    If WorksheetFunction.CountA(mInstanceSht.Range(mDataCell.Cells, endCell.Cells)) <= 1 Then
        Set mEorCell = mDataCell
    Else
        Set mEorCell = endCell
    End If
End Sub

Private Function isTemplate()
    isTemplate = (dataCell(T_TYPE).Value <> USER_MACRO)
End Function

Public Function AsIParameterReader() As IParameterReader
    Set AsIParameterReader = Me
End Function

Public Function AsIFileStream() As IFileStream
    Set AsIFileStream = Me
End Function
