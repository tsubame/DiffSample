VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DcTestContinuousForm 
   Caption         =   "DC TEST CONTINUOUS MODE"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   OleObjectBlob   =   "DcTestContinuousForm.frx":0000
End
Attribute VB_Name = "DcTestContinuousForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Dim mMeasInterpose As MEASURE_STATUS

Private Sub UserForm_Initialize()
    mMeasInterpose = MEAS_INITIAL
    Me.MeasureCounter = 1
End Sub

Private Sub ExitButton_Click()
    mMeasInterpose = MEAS_EXIT
End Sub

Private Sub RestartButton_Click()
    mMeasInterpose = MEAS_RESTART
End Sub

Private Sub StopButton_Click()
    mMeasInterpose = MEAS_STOP
End Sub

Public Property Get MeasureStatus() As MEASURE_STATUS
    MeasureStatus = mMeasInterpose
End Property

Public Property Let MeasureStatus(ByVal setIntarpose As MEASURE_STATUS)
    mMeasInterpose = setIntarpose
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Select Case mMeasInterpose:
            Case MEAS_STOP:
                Dim answer As Integer
                answer = MsgBox("Do You Really Exit Continuous Measurement?" & vbCrLf _
                                & "'Yes' To Return To The Examination Controller." _
                                , vbYesNo + vbQuestion + vbDefaultButton2, "INTERRUPT CONTINUOUS MEASUREMENT")
                Select Case answer:
                    Case vbYes:
                        ExitButton_Click
                    Case vbNo:
                End Select
            Case Else:
        End Select
        Cancel = True
    End If
End Sub
