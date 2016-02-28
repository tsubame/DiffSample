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

' ������
' @param name ����
Public Sub Initialize(Name As String)
    Name_ = Name
End Sub

' �SEvent�̎擾
' @return �SEvent
Public Function GetEvents() As PatDriveFormatEvent()
    GetEvents = Events_
End Function

' Event�̎擾
' @param no �擾���s��Event�̔ԍ�
' @return !Nothing Event
' @return Nothing Event��������Ȃ�����
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

' ���̂̎擾
' @return ����
Public Function GetName() As String
    GetName = Name_
End Function

' Event�̒ǉ�
' @param pdf_event �ǉ�����Event
Public Sub AddEvent(pdf_event As PatDriveFormatEvent)
    
    If pdf_event Is Nothing Then
        Exit Sub
    End If
    
    ' �����l�͐擪��Nothing�̂��߁AReDim�͍s�킸�㏑������
    If Not Events_(0) Is Nothing Then
        ReDim Preserve Events_(UBound(Events_) + 1)
    End If
    
    Set Events_(UBound(Events_)) = pdf_event
    
End Sub