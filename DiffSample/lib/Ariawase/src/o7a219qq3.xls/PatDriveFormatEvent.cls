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

' ������
' @param no �ԍ�
' @param event_name �C�x���g��
' @param range_index �s�ԍ�
Public Sub Initialize( _
    no As Long, _
    event_name As String, _
    range_index As Long)
    
    No_ = no
    Name_ = event_name
    RangeIndex_ = range_index
    
End Sub
    
' �C�x���gNo�̎擾
' @return �C�x���gNo
Public Function GetNo() As Long
    GetNo = No_
End Function

' �C�x���g���̎擾
' @return �C�x���g��
Public Function GetName() As String
    GetName = Name_
End Function

' �s�ԍ��̎擾
' �V�[�g������n�܂�Z���͈͓��ł̍s�ԍ�
' @return �s�ԍ�
Public Function GetRangeIndex() As Long
    GetRangeIndex = RangeIndex_
End Function