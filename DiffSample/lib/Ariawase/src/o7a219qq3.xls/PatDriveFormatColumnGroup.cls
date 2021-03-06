VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PatDriveFormatColumnGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Name_ As String
Private Columns_() As PatDriveFormatColumn

Private Sub Class_Initialize()
    Name_ = ""
    ReDim Columns_(0)
End Sub

Private Sub Class_Terminate()
    Erase Columns_
End Sub

' 初期化
' @param group_name 名称
Public Sub Initialize(group_name As String)
    
    Name_ = group_name
    
End Sub

' Columnの追加
' @param pdf_column 追加するColumn
Public Sub AddColumn( _
    ByRef pdf_column As PatDriveFormatColumn)

    If pdf_column Is Nothing Then
        Exit Sub
    End If
    
    If Not Columns_(0) Is Nothing Then
        ReDim Preserve Columns_(UBound(Columns_) + 1)
    End If
    Set Columns_(UBound(Columns_)) = pdf_column
    
End Sub

' Columnの取得
' @param identifier 識別子
' @return !Nothing Column
' @return Nothing 一致するColumn無し
Public Function GetColumnByIdentifier(identifier As String)

    Dim ii As Long
    
    If Columns_(0) Is Nothing Then
        Set GetColumnByIdentifier = Nothing
        Exit Function
    End If
        
    For ii = 0 To UBound(Columns_)
        
        If identifier = Columns_(ii).GetIdentifier() Then
            Set GetColumnByIdentifier = Columns_(ii)
            Exit Function
        End If
        
    Next
    
    Set GetColumnByIdentifier = Nothing
    
End Function

' Columnの取得
' @param column_name 名称
' @return !Nothing Column
' @return Nothing 一致するColumn無し
Public Function GetColumnByName(column_name As String)

    Dim ii As Long

    If Columns_(0) Is Nothing Then
        Set GetColumnByName = Nothing
        Exit Function
    End If
    
    For ii = 0 To UBound(Columns_)
    
        If column_name = Columns_(ii).GetName() Then
            Set GetColumnByName = Columns_(ii)
            Exit Function
        End If
    Next
    
    Set GetColumnByName = Nothing
    
End Function

' 名称取得
' @return 名称
Public Function GetName() As String
    GetName = Name_
End Function

' 全Columnの取得
Public Function GetColumns() As PatDriveFormatColumn()
    GetColumns = Columns_
End Function
