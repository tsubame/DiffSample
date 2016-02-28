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

' ‰Šú‰»
' @param group_name –¼Ì
Public Sub Initialize(group_name As String)
    
    Name_ = group_name
    
End Sub

' Column‚Ì’Ç‰Á
' @param pdf_column ’Ç‰Á‚·‚éColumn
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

' Column‚Ìæ“¾
' @param identifier ¯•Êq
' @return !Nothing Column
' @return Nothing ˆê’v‚·‚éColumn–³‚µ
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

' Column‚Ìæ“¾
' @param column_name –¼Ì
' @return !Nothing Column
' @return Nothing ˆê’v‚·‚éColumn–³‚µ
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

' –¼Ìæ“¾
' @return –¼Ì
Public Function GetName() As String
    GetName = Name_
End Function

' ‘SColumn‚Ìæ“¾
Public Function GetColumns() As PatDriveFormatColumn()
    GetColumns = Columns_
End Function
