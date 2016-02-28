VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PatDriveFormatEventGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Name_ As String
Private Events_() As PatDriveFormatEvent

Private Sub Class_Initialize()
    Name_ = ""
    ReDim Events_(0)
End Sub

Private Sub Class_Terminate()
    Erase Events_
End Sub

' 初期化
' @param name 名称
Public Sub Initialize(Name As String)
    Name_ = Name
End Sub

' 全Eventの取得
' @return 全Event
Public Function GetEvents() As PatDriveFormatEvent()
    GetEvents = Events_
End Function

' Eventの取得
' @param no 取得を行うEventの番号
' @return !Nothing Event
' @return Nothing Eventが見つからなかった
Public Function GetEvent(no As Long) As PatDriveFormatEvent
    
    Dim ii As Long
    For ii = 0 To UBound(Events_)
        If Events_(ii).GetNo() = no Then
            Set GetEvent = Events_(ii)
            Exit Function
        End If
    Next
    
    Set GetEvent = Nothing
    
End Function

' 名称の取得
' @return 名称
Public Function GetName() As String
    GetName = Name_
End Function

' Eventの追加
' @param pdf_event 追加するEvent
Public Sub AddEvent(pdf_event As PatDriveFormatEvent)
    
    If pdf_event Is Nothing Then
        Exit Sub
    End If
    
    ' 初期値は先頭がNothingのため、ReDimは行わず上書きする
    If Not Events_(0) Is Nothing Then
        ReDim Preserve Events_(UBound(Events_) + 1)
    End If
    
    Set Events_(UBound(Events_)) = pdf_event
    
End Sub
