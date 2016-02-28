Attribute VB_Name = "ICUL1G_GetForReadCsv_Mod_v3_0"
Option Explicit
'Const 定数定義
Private Const VoltageReadCsvStartRow As Integer = 5, VoltageReadCsvStartCol As Integer = 2
Private Const VoltageReadCsvPinRow As Integer = 4, VoltageReadCsvPinCol As Integer = 6
Private Const VoltageReadCsvSiteCol As Integer = 5, VoltageReadCsvSecCol As Integer = 3
Private Const VoltageReadCsvCondCol As Integer = 2

Private Const VoltageWkstStartRow As Integer = 5, VoltageWkstStartCol As Integer = 2
Private Const VoltageWkstPinRow As Integer = 4, VoltageWkstPinCol As Integer = 6
Private Const VoltageWkstSiteCol As Integer = 5, VoltageWkstSecCol As Integer = 3
Private Const VoltageWkstCondCol As Integer = 2, VoltageWkstSwNodeCol As Integer = 4

'Long
Dim VoltageReadCsvEndPinCol As Long, VoltageReadCsvEndCondRow As Long
Dim VoltageWkstEndPinCol As Long, VoltageWkstEndCondRow As Long

Public Sub Get_Hard_data()

With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
End With

Call GetCSVData(Power_Supply_VoltageoffsetFileName)
Call GetVoltageForReadCSV("Power-Supply Voltage")

Call GetCSVData(Clock_VoltageOffsetFileName)
Call GetVoltageForReadCSV("Clock Voltage")

With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
End With

Call ShtPowerV.Initialize
Call ShtClockV.Initialize

End Sub
Public Sub GetCSVData(ByVal CsvFileName As String)
    Dim strArg, temp5 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim FileNo As Integer

    '======== CSV File Exist Check.
    If Dir(CsvFileName) = "" Then
        Flg_CsvFileFailSafe = False
        MsgBox "Error! [" & CsvFileName & "] is Not Found!"
        Exit Sub
    End If

    '=======Start_ReadPower_Supply_VoltageoffsetFileName============
    Worksheets("Read CSV").Range("A1:AZ1000").Clear     'Clear Sheet

    FileNo = FreeFile

    Open CsvFileName For Input As #FileNo              'CSV File OPEN

    On Error GoTo CloseFile                             'Error Check

    i = 0
    Do Until EOF(FileNo)                                     'Data Input to buffer
        Line Input #FileNo, temp5
        i = i + 1
        strArg = Split(temp5, ",")
        For j = 0 To UBound(strArg)
            Worksheets("Read CSV").Cells(i, j + 1) = strArg(j)                 'Data Input to sheet
        Next j
    Loop

    Close #FileNo                                       'CSV File Close
    On Error GoTo 0

Offset_csv_end:

Exit Sub

CloseFile:

    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    Close #FileNo
    MsgBox ("File Open Error! Please Check CSV File" & vbCrLf & Right(CsvFileName, Len(CsvFileName) - InStrRev(CsvFileName, "\")))
    GoTo Offset_csv_end

    '=======End_ReadPower_Supply_VoltageoffsetFile============

End Sub
Private Sub GetVoltageForReadCSV(ByVal WriteWkstName As String)
'Object
Dim ReadCsvWkst As Object, WriteWkst As Object, ObjAutoFil As Object, AutoFil As Object
Dim ObjCondArray As Object, ObjSecArray As Object, ObjSiteArray As Object, ObjPinArray As Object
Dim ObjDataArea As Range, ObjWriteArea As Range, targetCell As Range
Dim WriteBaseCell As Range, ReadBaseCell As Range
'Double
Dim OffsetVal() As Double
'Variant
Dim Cond_Arr As Variant, Sec_Arr As Variant, Site_Arr As Variant, Pin_Arr As Variant
'Integer
Dim i As Integer, j As Integer
'Boolean
Dim NG_Flg As Boolean
'エラー発生時定義
On Error GoTo ErrorRoutine
'対象シートをオブジェクトで取得
Set ReadCsvWkst = ThisWorkbook.Sheets("Read CSV")
Set WriteWkst = ThisWorkbook.Sheets(WriteWkstName)
'書き込みシートの最終行、最終列の取得
With WriteWkst
'書き込みシートをアクティブに
    .Activate
'データのあるセルを検索
    With .UsedRange
        VoltageWkstEndCondRow = .Find("*", , , , xlByRows, xlPrevious).Row
        VoltageWkstEndPinCol = .Find("*", , , , xlByColumns, xlPrevious).Column
    End With
End With
'書き込みシートのクリア
Call DeleteVoltageOffset(WriteWkst)
'READ CSVの最終行、最終列の取得
With ReadCsvWkst
'READCSVシートをアクティブに
    .Activate
    With .UsedRange
        VoltageReadCsvEndCondRow = .Find("*", , , , xlByRows, xlPrevious).Row
        VoltageReadCsvEndPinCol = .Find("*", , , , xlByColumns, xlPrevious).Column
    End With
'コンディションを配列に代入
Cond_Arr = .Range(.Cells(VoltageReadCsvStartRow, VoltageReadCsvCondCol), .Cells(VoltageReadCsvEndCondRow, VoltageReadCsvCondCol))
'セクションを配列に代入
Sec_Arr = .Range(.Cells(VoltageReadCsvStartRow, VoltageReadCsvSecCol), .Cells(VoltageReadCsvEndCondRow, VoltageReadCsvSecCol))
'サイトを配列に代入
Site_Arr = .Range(.Cells(VoltageReadCsvStartRow, VoltageReadCsvSiteCol), .Cells(VoltageReadCsvEndCondRow, VoltageReadCsvSiteCol))
'ピンを配列に代入
Pin_Arr = .Range(.Cells(VoltageReadCsvPinRow, VoltageReadCsvPinCol), .Cells(VoltageReadCsvPinRow, VoltageReadCsvEndPinCol))
'ディクショナリオブジェクトを設定
Set ObjCondArray = CreateObject("Scripting.Dictionary")
Set ObjSecArray = CreateObject("Scripting.Dictionary")
Set ObjSiteArray = CreateObject("Scripting.Dictionary")
Set ObjPinArray = CreateObject("Scripting.Dictionary")
'配列から重複分、空白を削除し連想配列に代入
Call GetSDObj(Cond_Arr, ObjCondArray)
Call GetSDObj(Sec_Arr, ObjSecArray)
Call GetSDObj(Site_Arr, ObjSiteArray)
Call GetSDObj(Pin_Arr, ObjPinArray)
'配列のクリア
Erase Cond_Arr, Sec_Arr, Site_Arr, Pin_Arr
'オフセット取得用変数のサイズ設定
ReDim Preserve OffsetVal(UBound(ObjCondArray.Keys), UBound(ObjSecArray.Keys), UBound(ObjSiteArray.Keys), UBound(ObjPinArray.Keys))
'読み込みエリアの設定
Set ObjDataArea = .Range(Cells(VoltageReadCsvStartRow, VoltageReadCsvPinCol), Cells(VoltageReadCsvEndCondRow, VoltageReadCsvEndPinCol))
'セル単位でオフセット値を取得し変数に代入
For Each targetCell In ObjDataArea
    OffsetVal(ObjCondArray.Item(LCase(.Cells(targetCell.Row, VoltageReadCsvStartCol))), _
        ObjSecArray.Item(LCase(.Cells(targetCell.Row, VoltageReadCsvSecCol))), _
        ObjSiteArray.Item(LCase(.Cells(targetCell.Row, VoltageReadCsvSiteCol))), _
        ObjPinArray.Item(LCase(.Cells(VoltageReadCsvPinRow, targetCell.Column)))) = targetCell
Next
'読み込みエリアのクリア
Set ObjDataArea = Nothing
'CSVとJOBの比較処理
'基準セルの取得
Set WriteBaseCell = WriteWkst.Cells(VoltageWkstPinRow, VoltageWkstCondCol)
Set ReadBaseCell = ReadCsvWkst.Cells(VoltageReadCsvPinRow, VoltageReadCsvCondCol)
'CSVを基準にコンディション、セクション、サイトを比較
i = 1
Do Until ReadBaseCell.offset(i, 0).Row > VoltageReadCsvEndCondRow
    NG_Flg = True
    For j = VoltageWkstStartRow To VoltageWkstEndCondRow
        If ReadBaseCell.offset(i, 0) = WriteWkst.Cells(j, VoltageWkstCondCol) And _
            ReadBaseCell.offset(i, 1) = WriteWkst.Cells(j, VoltageWkstSecCol) And _
            ReadBaseCell.offset(i, 3) = WriteWkst.Cells(j, VoltageWkstSiteCol) Then
            NG_Flg = False
            Exit For
        End If
    Next j
    If NG_Flg = True Then
        MsgBox "Sheet = " & .Name & vbCrLf & _
            "Row = " & ReadBaseCell.offset(i, 0).Row & vbCrLf & _
            "Item Doesn't Match , Please Check!!", vbOKOnly, "Error!!"
        GoTo ErrorRoutine
    End If
i = i + 1
Loop
'CSVを基準にピンを比較
i = 0
Do Until ReadBaseCell.offset(0, i).Column > VoltageReadCsvEndPinCol
    NG_Flg = True
    For j = VoltageWkstCondCol To VoltageWkstEndPinCol
        If ReadBaseCell.offset(0, i) = WriteWkst.Cells(VoltageWkstPinRow, j) Then
            NG_Flg = False
            Exit For
        End If
    Next j
    If NG_Flg = True Then
        MsgBox "Sheet = " & .Name & vbCrLf & _
            "Row = " & ReadBaseCell.offset(0, i).Row & vbCrLf & _
            "Column = " & ReadBaseCell.offset(0, i).Column & vbCrLf & _
            "Item Doesn't Match , Please Check!!", vbOKOnly, "Error!!"
        GoTo ErrorRoutine
    End If
i = i + 1
Loop
End With
'変数から対象シートに書き込み
With WriteWkst
'対象シートをアクティブに
.Activate
'オートフィルタチェック
On Error Resume Next
Set ObjAutoFil = .AutoFilter.Filters()
On Error GoTo 0
If Not ObjAutoFil Is Nothing Then
    For Each AutoFil In ObjAutoFil
        If AutoFil.On Then
            .ShowAllData
        End If
    Next
End If
Set ObjAutoFil = Nothing: Set AutoFil = Nothing
'書き込みエリアの設定
Set ObjWriteArea = .Range(Cells(VoltageWkstStartRow, VoltageWkstPinCol), Cells(VoltageWkstEndCondRow, VoltageWkstEndPinCol))
'セル単位で対象シートにオフセット値を書き込み
For Each targetCell In ObjWriteArea
    If IsEmpty(.Cells(targetCell.Row, VoltageWkstSiteCol)) = False Then
        If ObjCondArray.Exists(LCase(.Cells(targetCell.Row, VoltageWkstStartCol))) = True And _
            ObjSecArray.Exists(LCase(.Cells(targetCell.Row, VoltageWkstSecCol))) = True And _
            ObjSiteArray.Exists(LCase(.Cells(targetCell.Row, VoltageWkstSiteCol))) = True And _
            ObjPinArray.Exists(LCase(.Cells(VoltageWkstPinRow, targetCell.Column))) = True Then
            .Cells(targetCell.Row, VoltageWkstSwNodeCol) = Sw_Node
            targetCell.Value = _
                OffsetVal(ObjCondArray.Item(LCase(.Cells(targetCell.Row, VoltageWkstStartCol))), _
                ObjSecArray.Item(LCase(.Cells(targetCell.Row, VoltageWkstSecCol))), _
                ObjSiteArray.Item(LCase(.Cells(targetCell.Row, VoltageWkstSiteCol))), _
                ObjPinArray.Item(LCase(.Cells(VoltageWkstPinRow, targetCell.Column))))
        Else
        MsgBox "Sheet = " & .Name & vbCrLf & _
            "Row = " & targetCell.Row & vbCrLf & _
            "Column = " & targetCell.Column & vbCrLf & _
            "Item Don't Exist , Please Check!!", vbOKOnly, "Error!!"
            Exit For
        End If
    End If
Next
'読み込みエリアのクリア
Set ObjWriteArea = Nothing
End With
'オブジェクトのクリア
Set ReadCsvWkst = Nothing: Set WriteWkst = Nothing
Set ObjCondArray = Nothing: Set ObjSecArray = Nothing: Set ObjSiteArray = Nothing: Set ObjPinArray = Nothing
'配列のクリア
Erase OffsetVal
'プロシージャの終了
Exit Sub
'エラー処理
ErrorRoutine:
        MsgBox "Error!! Check GetVoltageForReadCSV!!", vbOKOnly, "Error!!"
        Call Break
        End
End Sub
Private Sub DeleteVoltageOffset(ByVal OffsetWkstName As Object)  '書き込みシートのクリア
Dim i As Integer
With OffsetWkstName
    For i = VoltageWkstStartRow To VoltageWkstEndCondRow
        If IsEmpty(.Cells(i, VoltageWkstSiteCol)) = False Then
            .Cells(i, VoltageWkstSwNodeCol).ClearContents
            .Range(.Cells(i, VoltageWkstPinCol), .Cells(i, VoltageWkstEndPinCol)).ClearContents
        End If
    Next
End With
End Sub
Private Sub GetSDObj(ByVal BaseArray As Variant, OutSDObj As Object) '連想配列の作成
Dim i As Integer
Dim ArrayVal As Variant
i = 0
For Each ArrayVal In BaseArray
    If IsEmpty(ArrayVal) = False Then
        If OutSDObj.Exists(LCase(ArrayVal)) = False Then
            OutSDObj.Add LCase(ArrayVal), i
            i = i + 1
        End If
    End If
Next
End Sub


