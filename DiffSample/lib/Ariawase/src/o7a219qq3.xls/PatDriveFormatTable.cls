VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PatDriveFormatTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NoHeader = "_No_"
Private Const EventHeader = "_event_"

Private Name_ As String
Private Columns_() As PatDriveFormatColumn
Private ColumnGroups_() As PatDriveFormatColumnGroup
Private Events_() As PatDriveFormatEvent
Private EventGroups_() As PatDriveFormatEventGroup
Private Values_() As String

Private Sub Class_Initialize()
    Name_ = ""
    ReDim Columns_(0)
    ReDim ColumnGroups_(0)
    ReDim Events_(0)
    ReDim EventGroups_(0)
End Sub

Private Sub Class_Terminate()
    Erase Columns_
    Erase ColumnGroups_
    Erase Events_
    Erase EventGroups_
End Sub

' 初期化
' @param values 解析対象となるセル範囲
' @param pin_names ChansシートのPinNameリスト
Public Sub Initialize( _
    ByRef values As Range, _
    pin_names() As String)

    Dim pi As Long
    Dim ri As Long
    Dim ci As Long
    Dim ni As Long ' No Index
    Dim ei As Long ' Event Index
    Dim hri As Long ' Header Row Index
    Dim str_tmp As String
    Dim chans_pin_name As String
    
On Error GoTo ReleaseObjects
    
    If values Is Nothing Then
        Exit Sub
    End If
    
    ' 名称
    Name_ = Trim(values(1, 1))
            
    ' 見出しの位置を取得
    ' 2列目が"_"で始まる行の取得
    For ri = 1 To values.Rows.Count
        If InStr(1, Trim(values.Cells(ri, 2)), "_", vbBinaryCompare) = 1 Then
            Exit For
        End If
    Next
    
    ' Column作成処理
    For ci = 2 To values.Columns.Count
    
        ' Event作成用のインデックスを取得
        str_tmp = Trim(values.Cells(ri, ci))
        If StrComp(str_tmp, NoHeader, vbBinaryCompare) = 0 Then
            ni = ci
        ElseIf StrComp(str_tmp, EventHeader, vbBinaryCompare) = 0 Then
            ei = ci
        Else
            
On Error GoTo ReleaseColumn

            Dim pos As Integer
            Dim pin_name As String
            pin_name = values.Cells(ri - 1, ci)
            
            ' Columnの端子名と一致するChansのPinNameを取得
            For pi = 0 To UBound(pin_names)
                ' '_'以降の文字列を取得
                chans_pin_name = pin_names(pi)
                pos = InStr(1, chans_pin_name, "_", vbTextCompare)
                If pos > 0 Then
                    If Mid(chans_pin_name, pos + 1) = pin_name Then
                        pin_name = chans_pin_name
                        Exit For
                    End If
                End If
            Next
    
            ' Columnの作成
            ' 名称は識別子の一つ上のセル
            ' インデックスは0始まりにするため-2をしている
            Dim pdf_column As PatDriveFormatColumn
            Set pdf_column = New PatDriveFormatColumn
            Call pdf_column.Initialize( _
                values.Cells(ri, ci), _
                pin_name, _
                ci)
            
            If Not Columns_(0) Is Nothing Then
                ReDim Preserve Columns_(UBound(Columns_) + 1)
            End If
            Set Columns_(UBound(Columns_)) = pdf_column
                
            ' グループへ追加
            Dim pdf_group As PatDriveFormatColumnGroup
            Dim group_name As String
            group_name = GetColumnGroupName(pdf_column.GetIdentifier())
            If group_name <> "" Then
                ' グループ化対象であれば対象グループの取得
                Set pdf_group = GetColumnGroup(group_name)
                If pdf_group Is Nothing Then
                    ' グループが無い場合は作成
                    Set pdf_group = New PatDriveFormatColumnGroup
                    Call pdf_group.Initialize(group_name)
                    
                    If Not ColumnGroups_(0) Is Nothing Then
                        ReDim Preserve ColumnGroups_(UBound(ColumnGroups_) + 1)
                    End If
                    Set ColumnGroups_(UBound(ColumnGroups_)) = pdf_group
                End If
                
                ' グループへ追加
                Call pdf_group.AddColumn(pdf_column)
                            
            End If
            
ReleaseColumn:
            Set pdf_group = Nothing
            Set pdf_column = Nothing
            
        End If
    Next
    
    ' Event作成処理
    For ri = ri + 1 To values.Rows.Count
    
On Error GoTo ReleaseEvent

        ' Eventの作成
        Dim pdf_event As PatDriveFormatEvent
        Set pdf_event = New PatDriveFormatEvent
        Call pdf_event.Initialize( _
            CLng(values.Cells(ri, ni)), values.Cells(ri, ei), ri)
            
        If Not Events_(0) Is Nothing Then
            ReDim Preserve Events_(UBound(Events_) + 1)
        End If
        Set Events_(UBound(Events_)) = pdf_event
            
        ' EventGroupの取得
        Dim group As PatDriveFormatEventGroup
        Set group = GetEventGroup(pdf_event.GetName())
        If group Is Nothing Then
            ' 存在しない場合は新規作成
            Set group = New PatDriveFormatEventGroup
            Call group.Initialize(pdf_event.GetName())
            
            ' 配列に追加
            If Not EventGroups_(0) Is Nothing Then
                ReDim Preserve EventGroups_(UBound(EventGroups_) + 1)
            End If
            Set EventGroups_(UBound(EventGroups_)) = group
        End If
        
        ' EventGroupにEventを追加
        Call group.AddEvent(pdf_event)
                
ReleaseEvent:
        Set pdf_event = Nothing
        Set group = Nothing
            
    Next
    
    ' 値の取得
    ReDim Values_(UBound(Events_), UBound(Columns_))
    For ri = 0 To UBound(Events_)
        For ci = 0 To UBound(Columns_)
            Values_(ri, ci) = values.Cells( _
                Events_(ri).GetRangeIndex(), _
                Columns_(ci).GetRangeIndex())
        Next
    Next
        
ReleaseObjects:
    ' Nothing
    
End Sub

' EventGroupの取得
' @param group_name グループ名(イベント名と同じ)
' @return !Nothing PatDriveFormatEventGroup
' @return Nothing グループが見つからなかった
Public Function GetEventGroup(group_name As String) As PatDriveFormatEventGroup

    ' 先頭がNothingであればGroupは存在しない
    If EventGroups_(0) Is Nothing Then
        Set GetEventGroup = Nothing
        Exit Function
    End If

    Dim ii As Integer
    For ii = 0 To UBound(EventGroups_)
        If EventGroups_(ii).GetName() = group_name Then
            Set GetEventGroup = EventGroups_(ii)
            Exit Function
        End If
    Next
    
    Set GetEventGroup = Nothing
    
End Function

' 全EventGroupの取得
' @return 全EventGroup
Public Function GetEventGroups() As PatDriveFormatEventGroup()
    GetEventGroups = EventGroups_
End Function

' 名称取得
' @return 名称
Public Function GetName() As String
    GetName = Name_
End Function

' ColumnGroupの取得
' @param group_name グループ名
' @return !Nothing PatDriveFormatColumnGroup
' @return Nothing グループが見つからなかった
Public Function GetColumnGroup(group_name As String) As PatDriveFormatColumnGroup

    If ColumnGroups_(0) Is Nothing Then
        Set GetColumnGroup = Nothing
        Exit Function
    End If
    
    Dim ii As Long
    For ii = 0 To UBound(ColumnGroups_)
        If group_name = ColumnGroups_(ii).GetName() Then
            Set GetColumnGroup = ColumnGroups_(ii)
            Exit Function
        End If
    Next
    
    Set GetColumnGroup = Nothing
    
End Function

' 全ColumnGroupの取得
' @return 全ColumnGroup
Public Function GetColumnGroups() As PatDriveFormatColumnGroup()
    GetColumnGroups = ColumnGroups_
End Function

' グループ名の取得
' 識別子からグループ名を取得する
' @param column_identifier 識別子
' @return グループ名(空白時はグループ対象無し)
Private Function GetColumnGroupName(column_identifier As String) As String
    
    ' 識別子が特定のパターンであるかチェック
    If Not column_identifier Like "_*[0-9]_" Then
        GetColumnGroupName = ""
        Exit Function
    End If
    
    ' 後ろ2文字を削ったものをグループ名とする
    GetColumnGroupName = Mid(column_identifier, 1, Len(column_identifier) - 2)
    
End Function

' 値の取得
' @param event_name イベント名
' @param identifier 識別子
' @param event_index イベントのインデックス
' @return 値
Public Function GetValue( _
    event_name As String, _
    identifier As String, _
    Optional event_index As Integer = 0)
    
    Dim ri As Long
    Dim ci As Long
    Dim no As Long
    Dim ii As Long
    
    no = -1

    ' event_nameが数字のみであればイベント番号として解釈する
    For ii = 1 To Len(event_name)
        If Not Mid(event_name, ii, 1) Like "[0-9]" Then
            Exit For
        End If
    Next
    If ii > Len(event_name) Then
        no = CLng(event_name)
    End If
    
    ' 一致するイベントの検索
    If no = -1 Then
        Dim ei As Integer
        ei = 0
        ' イベント名をキーとする場合は同名イベント時のインデックスを考慮
        For ri = 0 To UBound(Events_)
            If event_name = Events_(ri).GetName() Then
                If event_index <= ei Then
                    Exit For
                End If
                ei = ei + 1
            End If
        Next
    Else
        For ri = 0 To UBound(Events_)
            If event_name = Events_(ri).GetNo() Then
                Exit For
            End If
        Next
    End If
    
    If ri > UBound(Events_) Then
        GetValue = ""
        Exit Function
    End If
    
    ' 一致する列の検索
    For ci = 0 To UBound(Columns_)
        If identifier = Columns_(ci).GetIdentifier() Then
            Exit For
        End If
    Next
    
    If ci > UBound(Columns_) Then
        GetValue = ""
        Exit Function
    End If
    
    GetValue = Values_(ri, ci)
    
End Function

' 端子名の取得
' @param identifier 識別子
' @param pin_number ピン番号
' @return 端子名
Public Function GetPinName( _
    identifier As String, _
    pin_number As Long) As String
    
    Dim id As String
    Dim ci As Long
    
    id = "_" + identifier + CStr(pin_number) + "_"
    
    ' 一致する列の検索
    For ci = 0 To UBound(Columns_)
        If id = Columns_(ci).GetIdentifier() Then
            GetPinName = Columns_(ci).GetName()
            Exit Function
        End If
    Next
    
    GetPinName = ""
    
End Function
