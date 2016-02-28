VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PatDriveFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 対象シート名
Private Const SheetName = "PatDriveFormat"

' 解析開始位置
Private Const BeginCol = 2
Private Const BeginRow = 4

' 最終行判定
Private Const SpaceRowMax = 3

' テーブルの配列
Private Tables_() As PatDriveFormatTable

Private Sub Class_Initialize()
    ReDim Tables_(0)
End Sub

Private Sub Class_Terminate()
    Erase Tables_
End Sub

Private Function GetChansPinNames(ByRef wbook As Workbook) As String()
    
    Dim wsheet As Worksheet
    Dim names() As String
    Dim Name As String
    Dim ii As Long
    
On Error GoTo ReleaseObjects:

    ' ChansシートのPinNameを取得
    For ii = 1 To wbook.Sheets.Count
        If "Chans" = wbook.Worksheets(ii).Name Then
            Set wsheet = wbook.Worksheets(ii)
            Exit For
        End If
    Next
    
    If wsheet Is Nothing Then
        GetChansPinNames = names
        Exit Function
    End If
    
    ' ピン数を求める
    ' セルB7からPinNameが入力されているものとする
    For ii = 7 To wsheet.Rows.Count
        Name = Trim(wsheet.Cells(ii, 2))
        If Name = "" Then
            Exit For
        End If
    Next
    
    ReDim names(ii - 8)
    For ii = 0 To UBound(names)
        names(ii) = wsheet.Cells(ii + 7, 2)
    Next
    
    GetChansPinNames = names
    
ReleaseObjects:
    Set wsheet = Nothing
    
End Function

Public Sub Initialize(ByRef wbook As Workbook)

On Error GoTo ReleaseObjects

    Dim wsheet As Worksheet
    
    Dim ii As Long
    Dim ci As Long ' Column Index
    Dim ri As Long ' Row Index
    Dim cirs As Long ' Column Index Range Start
    Dim rirs As Long ' Row Index Range Start
    Dim cire As Long ' Column Index Range End
    Dim rire As Long ' Row Index Range End
    Dim space_count As Integer
    Dim table_name As String
    Dim str_tmp As String
    Dim pin_names() As String
    
    ' ChansシートのPinName一覧を取得
    ' もし存在しない場合はアクセス時に例外となる
    pin_names = GetChansPinNames(wbook)
    
    ' 対象シートの取得
    For ii = 1 To wbook.Sheets.Count
        If SheetName = wbook.Worksheets(ii).Name Then
            Set wsheet = wbook.Worksheets(ii)
            Exit For
        End If
    Next
    
    ' 対象シートがなければ処理を終了
    If wsheet Is Nothing Then
        Exit Sub
    End If
    
    ' Tableの作成
    ci = BeginCol
    cirs = -1
    cire = -1
    rirs = -1
    rire = -1
    space_count = 0
    table_name = ""
    For ri = BeginRow To wsheet.Rows.Count
    
        ' シート名(テーブル名)取得
        str_tmp = Trim(wsheet.Cells(ri, ci))
        
        ' 空行チェック
        If Len(str_tmp) > 0 Then
        
            ' シート名の入力有り
            space_count = 0
            
            ' Range開始セルの設定
            If table_name = "" Then
                        
                table_name = str_tmp
                cirs = ci
                rirs = ri
                
                ' 文字列が'_'で始まる行の検索
                Dim ri_ As Long
                For ri_ = ri To wsheet.Rows.Count
                    If InStr(1, Trim(wsheet.Cells(ri_, ci + 1)), "_") = 1 Then
                        Exit For
                    End If
                Next
                                
                ' 最終列の検索
                For cire = ci + 1 To wsheet.Columns.Count
                    ' 空白セルの直前を最終列とする
                    If Trim(wsheet.Cells(ri_, cire)) = "" Then
                        cire = cire - 1
                        Exit For
                    End If
                Next
                
            ' シート名が一つ上の行と異なる場合はテーブルの作成
            ElseIf table_name <> str_tmp Then
            
                ' テーブルの最終行を現在値の-1とする
                rire = ri - 1
                
                ' テーブルの追加
                Call AddTable(wsheet.Range( _
                    wsheet.Cells(rirs, cirs), wsheet.Cells(rire, cire)), _
                    pin_names)
            
            End If
                        
        Else
        
            ' 空行
            
            ' 直前までのデータでテーブルの作成
            If Not table_name = "" Then
                rire = ri - 1
            End If
            
            space_count = space_count + 1
            
        End If
        
        ' rireが設定されていればテーブルを作成
        If rire > 0 Then
        
            Call AddTable(wsheet.Range( _
                wsheet.Cells(rirs, cirs), wsheet.Cells(rire, cire)), _
                pin_names)
            
            ' 変数初期化
            table_name = ""
            rirs = -1
            cirs = -1
            cire = -1
            rire = -1
            
        End If
        
        ' 空行が指定行数より連続した場合はループを抜ける
        If space_count >= SpaceRowMax Then
            Exit For
        End If
    
    Next
        
ReleaseObjects:

    Set wsheet = Nothing

End Sub

Private Sub AddTable( _
    values As Range, _
    pin_names() As String)

On Error GoTo ReleaseObjects

    ' Table作成
    Dim table As PatDriveFormatTable
    Set table = New PatDriveFormatTable
    Call table.Initialize(values, pin_names)

    ' 作成したTableを格納
    ' 初期値は先頭がNothingのため、Nothingの場合はRedimを行わない
    If Not Tables_(0) Is Nothing Then
        ReDim Preserve Tables_(UBound(Tables_) + 1)
    End If
        
    Set Tables_(UBound(Tables_)) = table
            
ReleaseObjects:
    Set table = Nothing
    
End Sub

' 値の取得
' @param sheet_name シート名(テーブル名)
' @param event_name イベント名
' @param identifier 識別子
' @param event_index イベントインデックス
' @return 値
Public Function GetValue( _
    sheet_name As String, _
    event_name As String, _
    identifier As String, _
    Optional event_index As Integer = 0) As String
    
    Dim ti As Integer
    For ti = 0 To UBound(Tables_)
        If sheet_name = Tables_(ti).GetName() Then
            GetValue = Tables_(ti).GetValue(event_name, identifier, event_index)
            Exit Function
        End If
    Next
        
    GetValue = ""

End Function

' Tableの取得
' @param sheet_name シート名(テーブル名)
' @return !Nothing PatDriveFormatTable
' @return Nothing 一致するTable無し
Public Function GetTable(sheet_name As String) As PatDriveFormatTable

    Dim ii As Long
    For ii = 0 To UBound(Tables_)
        If Tables_(ii).GetName() = sheet_name Then
            Set GetTable = Tables_(ii)
            Exit Function
        End If
    Next
    
    Set GetTable = Nothing
    
End Function

' 全Tableの取得
' @return 全Table
Public Function GetTables() As PatDriveFormatTable()
    GetTables = Tables_
End Function

' 端子名(Outputのみ)の取得
' @param sheet_name シート名(テーブル名)
' @param identifier 識別子
' @param pin_number ピン番号
' @return 端子名
Public Function GetPinName( _
    sheet_name As String, _
    identifier As String, _
    pin_number As Long) As String
    
    Dim table As PatDriveFormatTable
    Set table = GetTable(sheet_name)
    If table Is Nothing Then
        GetPinName = ""
    Else
        GetPinName = table.GetPinName(identifier, pin_number)
    End If
    
    Set table = Nothing
    
End Function
