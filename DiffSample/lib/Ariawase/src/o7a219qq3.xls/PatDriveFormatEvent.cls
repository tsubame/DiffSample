VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PatDriveFormatEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private No_ As Long
Private Name_ As String
Private RangeIndex_ As Long

Private Sub Class_Initialize()
    No_ = 0
    Name_ = ""
End Sub

Private Sub Class_Terminate()
    ' Nothing
End Sub

' 初期化
' @param no 番号
' @param event_name イベント名
' @param range_index 行番号
Public Sub Initialize( _
    no As Long, _
    event_name As String, _
    range_index As Long)
    
    No_ = no
    Name_ = event_name
    RangeIndex_ = range_index
    
End Sub
    
' イベントNoの取得
' @return イベントNo
Public Function GetNo() As Long
    GetNo = No_
End Function

' イベント名の取得
' @return イベント名
Public Function GetName() As String
    GetName = Name_
End Function

' 行番号の取得
' シート名から始まるセル範囲内での行番号
' @return 行番号
Public Function GetRangeIndex() As Long
    GetRangeIndex = RangeIndex_
End Function
