Attribute VB_Name = "XEeeAuto_Offset"
Option Explicit
'概要:
'   SW_NODEを代替すべく、オフセットマネージャシートにアクセスするモジュール
'
'目的:
'
'
'作成者:
'   2012/02/03 Ver0.1 D.Maruyama
'   2012/03/05 Ver0.2 D.Maruyama 自動オフセット機能追加、運用的な問題があるので使用しない
'   2012/04/06 Ver0.3 D.Maruyama 自動オフセット機能削除、明示的なオフセットチェック機能を追加
'   2012/04/09 Ver0.4 D.Maruyama オフセットCSVファイルの名前が空白であるときはWriteOffsetMangerを省略するように変更
'   2013/03/08 Ver0.5 SLSI Offset ManagerシートへのテストLabel出力はサイト0のみに変更(@WriteOffsetSheet)、エラーメッセージ誤記修正。

Private Const OFFSET_SHEET As String = "Offset Manager"
Private Const SHEET_MANAGER_NAME As String = "Sheet Manager"
Private Const SHEET_MANAGER_JOBNAME As String = "B5"
Private Const SHEET_MANAGER_OFFSETSHEET As String = "E5"

Private Const TEST_INSTANCE_SHEET As String = "Test Instances"
Private Const OFFSET_SHEET_START_CELL As String = "C5"
Private Const TEST_INSTANCE_SHEET_START_CELL As String = "B5"

Private Const OFFSET_RESULT_FUNCTION As String = "ReturnResultEx_f"

Private m_colApplyOffsetTest As Collection

Private Type PositionInfo
    ParameterSlope_Column As Long
    ParameterOffset_Column As Long
    Sw_node_Row As Long
    Sw_node_Column As Long
    Test_Row As Long
    Test_Column As Long
    Unit_Column As Long
End Type

Public Sub WriteOffsetManager()

    Dim WriteWkstName As String
    'シートマネージャからオフセットシート名を取得
    WriteWkstName = GetCurrentOffsetManagerName
    Set m_colApplyOffsetTest = Nothing
    
    If WriteWkstName = "" Then
        TheExec.Datalog.WriteComment "[Check Comment] Offset.csv file is not use mode !"
        Exit Sub
    End If
    
    Dim ReadCsvWkst As Worksheet, WriteWkst As Worksheet
    
    Set m_colApplyOffsetTest = New Collection
    
    On Error GoTo ErrHandler
    
    '対象シートをオブジェクトで取得
    Set ReadCsvWkst = ThisWorkbook.Sheets("Read CSV")
    Set WriteWkst = ThisWorkbook.Sheets(WriteWkstName)
    Call ClearOffsetManagerSheet(WriteWkst)

    'ReadCSVシートから基準点をサーチ
    Dim Position As PositionInfo
    With ReadCsvWkst
        Position.Sw_node_Row = .Range("SW_NODE").Row
        Position.Test_Row = .Range("TEST").Row
        Position.Sw_node_Column = .Range("SW_NODE").Column
        Position.Test_Column = .Range("TEST").Column
        Position.Unit_Column = .Range("unit").Column
    End With
    
    'ReadCSVシートからオフセットをかけるべきテスト項目をリストアップ
    ReadCsvWkst.Activate
    Dim colTemp As New Collection
    Call ReadOffsetSheet(ReadCsvWkst, colTemp, Position)
    
    'ReadCSVシートからスロープと係数を読み込む
    Dim i As Long
    Dim objOffsetItem As CResultOffset
    Dim strItemName As Variant
    
    i = 1
    For Each strItemName In colTemp
        
        Set objOffsetItem = New CResultOffset
        
        Call GetOffsetParameterFromReadCSV(ReadCsvWkst, LCase(strItemName), objOffsetItem, Position)
        
        m_colApplyOffsetTest.Add objOffsetItem, objOffsetItem.Name
        Set objOffsetItem = Nothing
        
        i = i + 1
        
    Next strItemName
    
    '取得した値をオフセットシートに反映
    Call WriteOffsetSheet(WriteWkst, m_colApplyOffsetTest)
    
    Set WriteWkst = Nothing
    Set ReadCsvWkst = Nothing

ErrHandler:

    Set WriteWkst = Nothing
    Set ReadCsvWkst = Nothing
    
End Sub

Public Sub CheckAllOffsetExist()

    
   On Error GoTo ErrHandler
   
    If m_colApplyOffsetTest Is Nothing Then
        Exit Sub
    End If
    
    If m_colApplyOffsetTest.Count = 0 Then
        Exit Sub
    End If
    
    Dim mySht As Worksheet
    
    '対象シートをオブジェクトで取得
    Set mySht = ThisWorkbook.Sheets(TEST_INSTANCE_SHEET)
    
    '対象シートからテスト名を取得
        'TestConditionの全情報を配列に格納
    Dim rngStart As Range
    Dim rngEnd As Range
    Dim aryTest As Variant
    With mySht
        Set rngStart = .Range(TEST_INSTANCE_SHEET_START_CELL)
        Set rngEnd = .Range(TEST_INSTANCE_SHEET_START_CELL).End(xlDown)
        aryTest = .Range(rngStart, .Cells(rngEnd.Row, rngStart.Column + 2)).Value
    End With
    Set mySht = Nothing
    Set rngStart = Nothing
    Set rngEnd = Nothing
    
    
    'リザルトマネージャを基準にテストインスタンスを捜索
    Dim i As Long
    Dim IsFound As Boolean
    Dim IsOffsetCalled As Boolean
    Dim obj As CResultOffset
    Dim strTestName As String
    For Each obj In m_colApplyOffsetTest
        IsFound = False
        IsOffsetCalled = False
        For i = 1 To UBound(aryTest, 1)
            If obj.Name = aryTest(i, 1) Then
                If UCase(aryTest(i, 3)) = UCase(OFFSET_RESULT_FUNCTION) Then
                    IsFound = True
                    IsOffsetCalled = True
                    Exit For
                Else
                    IsFound = True
                    IsOffsetCalled = False
                    Exit For
                End If
            End If
        Next i
        If Not IsFound Or Not IsOffsetCalled Then
            strTestName = obj.Name
            Exit For
        End If
    Next
    
    On Error GoTo 0
    
    'そもそもSheetManagerにOffset Sheetの記載がない場合
    If TheOffsetResult Is Nothing Then
        MsgBox "Offset Sheet Name is not found Sheet Manager'", , "Offset Sheet"
        DisableAllTest
        First_Exec = 0
    End If
    
    '見つからずに抜けた場合
    If Not IsFound Then
        MsgBox "Offset Test " & strTestName & " is not existing in TestInstance", , "Offset Sheet"
        DisableAllTest
        First_Exec = 0
        Exit Sub
    End If
    
    '見つかったけど関数名がちがう場合
    If Not IsOffsetCalled Then
        MsgBox "Offset Test " & strTestName & " is not called by ReturnResultEx_f", , "Offset Sheet"
        DisableAllTest
        First_Exec = 0
        Exit Sub
    End If
    
    Exit Sub
    
ErrHandler:
   Err.Raise 9999, "CheckAllOffsetExist", "Internal Error"

End Sub

Public Function IsUseOffsetfile() As Boolean
    
    Dim tmpName As String
    IsUseOffsetfile = GetCurrentOffsetManagerName_impl(tmpName)
    
    If tmpName = "" Then
        IsUseOffsetfile = False
    End If
    
End Function

Private Function GetCurrentOffsetManagerName() As String
    
    Call GetCurrentOffsetManagerName_impl(GetCurrentOffsetManagerName)

End Function

Private Function GetCurrentOffsetManagerName_impl(ByRef strSheetName As String) As Boolean
    
    Dim strCurrentJobName As String
    strCurrentJobName = TheExec.CurrentJob
    
    On Error GoTo SHEET_MANAGER_NOT_FOUND
    Dim shtManager As Worksheet
    Set shtManager = ThisWorkbook.Worksheets.Item(SHEET_MANAGER_NAME)
    On Error GoTo 0
    
    Dim i As Long
    Dim IsFound As Boolean
    i = 0
    IsFound = False
    With shtManager
        While ((Not IsEmpty(.Range(SHEET_MANAGER_JOBNAME).offset(i, 0))) And IsFound = False)
            If (.Range(SHEET_MANAGER_JOBNAME).offset(i, 0) = strCurrentJobName) Then
                strSheetName = .Range(SHEET_MANAGER_OFFSETSHEET).offset(i, 0).Text
                IsFound = True
            End If
            i = i + 1
        Wend
    End With
    
    '見つからなかった場合デフォルトの名前をかく
    'おそらくエラーになるが、ユーザーのSheatManagerの存在を意識させる
    If Not IsFound Then
        strSheetName = ""
        GetCurrentOffsetManagerName_impl = False
        Exit Function
    End If
    
    GetCurrentOffsetManagerName_impl = True
    
    Exit Function

SHEET_MANAGER_NOT_FOUND:
    strSheetName = ""
    GetCurrentOffsetManagerName_impl = False
    
End Function


Private Sub ClearOffsetManagerSheet(ByRef shtOffsetManager As Worksheet)

    shtOffsetManager.Range(OFFSET_SHEET_START_CELL).offset(-1, 1) = ""

    Dim rngStart As Range
    Dim rngEnd As Range
    Dim rngClear As Range
    With shtOffsetManager
        Set rngStart = .Range(OFFSET_LABEL).offset(1, 0)
        Set rngEnd = .Range(OFFSET_LABEL).SpecialCells(xlCellTypeLastCell)
        Set rngClear = .Range(rngStart, rngEnd)
        rngClear.Clear
    End With
    
    Set rngStart = Nothing
    Set rngEnd = Nothing
    Set rngClear = Nothing
        
End Sub

Private Function ReadOffsetSheet(ByRef shtOffset As Worksheet, ByRef colItems As Collection, ByRef Position As PositionInfo) As Boolean

    If shtOffset Is Nothing Then
        ReadOffsetSheet = False
        Exit Function
    End If
    If colItems Is Nothing Then
        ReadOffsetSheet = False
        Exit Function
    End If
    
    Dim i As Long
    i = 1
    While (shtOffset.Cells(Position.Test_Row, Position.Test_Column).offset(i, 0).Text <> "")
        Call colItems.Add(shtOffset.Cells(Position.Test_Row, Position.Test_Column).offset(i, 0).Text)
        i = i + 1
    Wend
    
    ReadOffsetSheet = True
    
End Function

Private Function WriteOffsetSheet(ByRef shtOffset As Worksheet, ByRef colItems As Collection) As Boolean

    On Error GoTo ErrHandler

    If shtOffset Is Nothing Then
        WriteOffsetSheet = False
        Exit Function
    End If
    If colItems Is Nothing Then
        WriteOffsetSheet = False
        Exit Function
    End If
    
    Dim i As Long, j As Long
    Dim objOffsetItem As CResultOffset
    
    i = 1
    
    With shtOffset.Range(OFFSET_SHEET_START_CELL)
        .offset(-1, 1) = Sw_Node
        For Each objOffsetItem In colItems
            For j = 0 To nSite
                If j = 0 Then
                    .offset(i + j, 0) = objOffsetItem.Name
                End If
                .offset(i + j, 1) = j
                .offset(i + j, 2) = objOffsetItem.GetSlope(j)
                If objOffsetItem.GetUnit <> "" Then
                    .offset(i + j, 3) = CStr(objOffsetItem.GetOffset(j)) & objOffsetItem.GetUnit
                Else
                    .offset(i + j, 3) = CStr(objOffsetItem.GetOffset(j)) & "V"
                End If
            Next j
            i = i + (nSite + 1)
        Next objOffsetItem
    End With
    
    WriteOffsetSheet = True
    
    Exit Function
    
ErrHandler:
    MsgBox shtOffset.Range(OFFSET_LABEL).offset(i, 0)
End Function

Public Function GetOffsetParameterFromReadCSV(ByRef shtOffset As Worksheet, ByVal strItem As String, _
            ByRef objResultOffset As CResultOffset, ByRef Position As PositionInfo) As Double

    Dim i As Long
    Dim site As Long
    Dim arySlope(nSite) As Double
    Dim aryOffset(nSite) As Double
    Dim strUnit As String
   
    For site = 0 To nSite
        With Position
            i = 0
            While Sw_Node & "-" & site & "-Slope" <> shtOffset.Cells(.Sw_node_Row + 1, .Sw_node_Column + i)
                If shtOffset.Cells(.Sw_node_Row + 1, .Sw_node_Column + i) = "" Then
                    MsgBox "Not Find Sw_node= " & Sw_Node & "Slope @Offset Sheet"
                    Exit Function
                End If
                i = i + 1
            Wend
            .ParameterSlope_Column = .Sw_node_Column + i
        
            i = 0
            While Sw_Node & "-" & site & "-Offset" <> shtOffset.Cells(.Sw_node_Row + 1, .Sw_node_Column + i)
                If shtOffset.Cells(.Sw_node_Row + 1, .Sw_node_Column + i) = "" Then
                    MsgBox "Not Find Sw_node= " & Sw_Node & " Offset @Offset Sheet"
                    Exit Function
                End If
                i = i + 1
            Wend
            .ParameterOffset_Column = .Sw_node_Column + i
        End With
        
        Call GetSlope_FromSheet(arySlope(site), strItem, Position)
        Call GetOffset_FromSheet(aryOffset(site), strUnit, strItem, Position)
    Next site
    
    With objResultOffset
        .Name = strItem
        Call .SetParam(arySlope, aryOffset, strUnit)
    End With
    
End Function
    
'ここから下はもともとの流用

Private Sub GetOffset_FromSheet(ByRef OffsetVal As Double, ByRef strUnit As String, ByVal testName As String, ByRef Position As PositionInfo)

    Dim Row As Long
    Dim Unit As String
    Dim i As Long
        
    With Position
        i = 1
        While testName <> Cells(.Test_Row + i, .Test_Column)
            If Cells(.Test_Row + i, .Test_Column) = "" Then
                GoTo ErrHandler
            End If
            i = i + 1
        Wend
        Row = .Test_Row + i
        
        If Cells(Row, .ParameterOffset_Column) = "" Then
            MsgBox "Not Find Offset : " & testName & " @Offset Sheet"
        End If
        
        OffsetVal = Cells(Row, .ParameterOffset_Column)
        
    End With
                    
    Unit = Cells(Row, Position.Unit_Column)
'    Select Case Unit
'        Case "", "V", "A"
'            OffsetVal = OffsetVal
'        Case "%"
'            OffsetVal = OffsetVal * 0.01
'        Case "mV", "mA"
'            OffsetVal = OffsetVal * 0.001
'        Case "uV", "uA"
'            OffsetVal = OffsetVal * 0.000001
'        Case Else
'            MsgBox "Unkown Unit !!" & TestName & "Please Check!!"
'    End Select


    '%はVでだす、あとは係数を含めてOffsetManagerに記述
    If Unit = "%" Then
        strUnit = "V"
        OffsetVal = OffsetVal * 0.01
    Else
        strUnit = Unit
    End If
    Exit Sub
ErrHandler:
    MsgBox "Not Find Test : " & testName & " @Offset Sheet"
    Call DisableAllTest 'EeeJob関数
End Sub
Private Sub GetSlope_FromSheet(ByRef OffsetVal As Double, ByVal testName As String, Position As PositionInfo)

    Dim Row As Long
    Dim Unit As String
    Dim i As Long
        
    With Position
        i = 1
        While testName <> Cells(.Test_Row + i, .Test_Column)
            If Cells(.Test_Row + i, .Test_Column) = "" Then
                GoTo ErrHandler
            End If
            i = i + 1
        Wend
        Row = .Test_Row + i
        
        If Cells(Row, .ParameterSlope_Column) = "" Or Cells(Row, .ParameterSlope_Column) = 0 Then
            MsgBox "Not Find Slope : " & testName & " @Offset Sheet"
        End If
        
        OffsetVal = Cells(Row, .ParameterSlope_Column)
        
    End With
    Exit Sub
ErrHandler:
    MsgBox "Not Find Test : " & testName & " @Offset Sheet"
    Call DisableAllTest 'EeeJob関数
End Sub



