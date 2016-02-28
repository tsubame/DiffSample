VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DcTestExaminForm 
   Caption         =   "DC TEST EXECUTION CONTROLLER"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   OleObjectBlob   =   "DcTestExaminForm.frx":0000
End
Attribute VB_Name = "DcTestExaminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Dim mStatus As CONTROL_STATUS
Dim mLoop As Long

Private Sub LoopNumber_Change()
    If LoopNumber.Value <> "" Then
        mLoop = LoopNumber.Value
    Else
        mLoop = 1
    End If
End Sub

Private Sub ReturnButton_Click()
    mStatus = TEST_RETURN
End Sub

Private Sub StepButton_Click()
    mStatus = TEST_STEP
End Sub

Private Sub ContinueButton_Click()
    mStatus = TEST_CONTINUE
End Sub

Private Sub MeasureButton_Click()
    mStatus = TEST_REPEAT
End Sub

Private Sub EndButton_Click()
    mStatus = TEST_END
End Sub

Public Property Get ControlStatus() As CONTROL_STATUS
    ControlStatus = mStatus
End Property

Public Property Let ControlStatus(ByVal setStatus As CONTROL_STATUS)
    mStatus = setStatus
End Property

Public Property Get loopCount() As Long
    loopCount = mLoop
End Property

Public Property Let loopCount(ByVal setLoop As Long)
    LoopNumber.Value = setLoop
End Property

Private Sub UserForm_Terminate()
    EndButton_Click
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Dim answer As Integer
        answer = MsgBox("Do You Really Exit The Examination Controller?" & vbCrLf _
                        & "'Yes' To Exit Examination And Continue Subsequent DC Tests." _
                        , vbYesNo + vbQuestion + vbDefaultButton2, "EXIT EXAMINATION CONTROLLER")
        Select Case answer:
            Case vbYes:
                ContinueButton_Click
            Case vbNo:
        End Select
        Cancel = True
    End If
End Sub
