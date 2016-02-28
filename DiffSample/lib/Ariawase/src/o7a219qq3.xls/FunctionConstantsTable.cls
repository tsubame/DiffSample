VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctionConstantsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �����o�ϐ�
Private Name_ As String
Private Dic_ As Object

' �w�b�_���
Private Const NameHeader = "_ConstantDefTable_"
Private Const ValueHeader = "_Value_"

Private Sub Class_Initialize()
    Name_ = ""
    Set Dic_ = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set Dic_ = Nothing
End Sub

' ������
' @param values �e�[�u���̓��e���L�q���ꂽ�Z���̈�
Public Sub Initialize(values As Range)
    
    Dim ri As Integer
    Dim ci As Integer
    Dim ni As Integer ' Name(Key) Index
    Dim vi As Integer ' Value Index
    Dim str_tmp As String
    
    On Error GoTo ReleaseObjects
    
    ' ���O
    Name_ = Trim(values.Cells(1, 1))
    
    ' ���o���̈ʒu���擾
    ' 2��ڂ�"_"�Ŏn�܂�s�̎擾
    For ri = 1 To values.Rows.Count
        If InStr(1, Trim(values.Cells(ri, 2)), "_", vbBinaryCompare) = 1 Then
            Exit For
        End If
    Next
    
    For ci = 2 To values.Columns.Count
        str_tmp = Trim(values.Cells(ri, ci))
        If StrComp(str_tmp, NameHeader, vbBinaryCompare) = 0 Then
            ni = ci
        ElseIf StrComp(str_tmp, ValueHeader, vbBinaryCompare) = 0 Then
            vi = ci
        End If
    Next
    
    ' ���o���s�̎��s����f�[�^�Ƃ���
    For ri = ri + 1 To values.Rows.Count
        ' Dictionary�ɒǉ�
        Call Dic_.Add(values.Cells(ri, ni).Value, values.Cells(ri, vi).Value)
    Next
    
ReleaseObjects:
    ' Nothing
        
End Sub

' ���O�̎擾
' @return ���O
Public Function GetName() As String
    GetName = Name_
End Function

' �L�[�̈ꗗ���擾
' @return �L�[(�萔��)�ꗗ
Public Function GetNames() As String()

    Dim ii As Integer
    Dim values() As String
    
    If IsNull(Dic_) Or Dic_.Count <= 0 Then
        Exit Function
    End If
    
    ReDim values(Dic_.Count - 1)
    
    ' �R�s�[
    For ii = 0 To Dic_.Count - 1
        values(ii) = Dic_.Keys()(ii)
    Next
    
    GetNames = values
        
End Function

' �l�̎擾
' @param key �L�[(�萔��)
' @return �萔
Public Function GetValue(key As String) As Variant
    
    If IsNull(Dic_) Or Dic_.Count <= 0 Or Not Dic_.Exists(key) Then
        Exit Function
    End If
    
    GetValue = Dic_.Item(key)
        
End Function