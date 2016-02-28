VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CheckInputForm 
   Caption         =   "PARAMETER"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   OleObjectBlob   =   "CheckInputForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CheckInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private querycount As Integer
Private defaultvalue As Boolean

Private usermode As Boolean

Private Sub CancelButton_Click()
    
    CheckValue = -1
    Unload Me
    
End Sub

Private Sub CheckBox_Click()
    Dim Count As Integer
    Dim result As Integer
    
    If usermode Then
        If CheckBox.Value <> defaultvalue And 0 < querycount Then
            For Count = querycount To 1 Step -1
                If 1 < Count Then
                    result = MsgBox("DO YOU WANT TO REALLY CHANGE THE CHECK?", vbYesNo + vbDefaultButton2, "CHANGE SELECTION")
                Else
                    result = MsgBox("ARE YOU SURE?", vbYesNo + vbDefaultButton2, "CHANGE SELECTION")
                End If
                
                If result = vbNo Then
                    usermode = False
                    CheckBox.Value = defaultvalue
                    usermode = True
                    Exit Sub
                End If
            Next Count
        End If
    End If
    
End Sub

Private Sub OKButton_Click()

    Select Case CheckBox.Value
      Case True
        CheckValue = 1
      Case False
        CheckValue = 0
      Case Else
        CheckValue = -1
    End Select
    
    Unload Me
    
End Sub

Private Sub UserForm_Activate()
    
    CheckBox.SetFocus

End Sub

Private Sub UserForm_Initialize()

    querycount = 0
    defaultvalue = False
    usermode = False
    
End Sub

Public Function Setup(Label As String, default As Boolean, query As Integer) As String
    
    defaultvalue = default
    querycount = query
    
    CheckBox.Caption = Label
    CheckBox.Value = default
    
    usermode = True
    
End Function

