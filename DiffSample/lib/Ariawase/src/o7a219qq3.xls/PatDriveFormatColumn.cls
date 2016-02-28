VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PatDriveFormatColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Name_ As String
Private Identifier_ As String
Private RangeIndex_ As Long

Private Sub Class_Initialize()
    Name_ = ""
    Identifier_ = ""
    RangeIndex_ = 0
End Sub
 
Private Sub Class_Terminate()
    ' Nothing
End Sub

' 初期化
' @param identifier 識別子
' @param column_name 名称(空文字可)
' @param range_index 列番号
Public Sub Initialize( _
    identifier As String, _
    column_name As String, _
    range_index As Long)
    
    Identifier_ = identifier
    Name_ = column_name
    RangeIndex_ = range_index
    
End Sub

' 名称の取得
' @return 名称
Public Function GetName() As String
    GetName = Name_
End Function

' 識別子の取得
' @return 識別子
Public Function GetIdentifier() As String
    GetIdentifier = Identifier_
End Function

' 列番号の取得
' シート名から始まるセル範囲内での列番号
' @return 列番号
Public Function GetRangeIndex() As Long
    GetRangeIndex = RangeIndex_
End Function
