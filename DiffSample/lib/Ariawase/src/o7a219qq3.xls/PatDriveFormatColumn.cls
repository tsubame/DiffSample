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

' ������
' @param identifier ���ʎq
' @param column_name ����(�󕶎���)
' @param range_index ��ԍ�
Public Sub Initialize( _
    identifier As String, _
    column_name As String, _
    range_index As Long)
    
    Identifier_ = identifier
    Name_ = column_name
    RangeIndex_ = range_index
    
End Sub

' ���̂̎擾
' @return ����
Public Function GetName() As String
    GetName = Name_
End Function

' ���ʎq�̎擾
' @return ���ʎq
Public Function GetIdentifier() As String
    GetIdentifier = Identifier_
End Function

' ��ԍ��̎擾
' �V�[�g������n�܂�Z���͈͓��ł̗�ԍ�
' @return ��ԍ�
Public Function GetRangeIndex() As Long
    GetRangeIndex = RangeIndex_
End Function