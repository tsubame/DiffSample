VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcFormMessageWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   DCテスト実験フォーム用ライタークラス
'
'目的:
'   実験フォームの各ラベルにステイタスを書き出す
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWriter

Private Sub IParameterWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
End Sub

Private Sub IParameterWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
    With DcTestExaminForm
        Select Case dataName:
            Case TEST_CATEGORY: .Instance.Caption = strData
            Case ACTION_LABEL: .Condition.Caption = strData
        End Select
    End With
End Sub

Private Sub IParameterWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
    With DcTestExaminForm
        Select Case dataName:
            Case GROUP_ID::
                If lngData = NO_TEST Then
                    .GroupNumber.Caption = ""
                Else
                    .GroupNumber.Caption = lngData
                End If
            Case SITE_INDEX:
                If lngData = ALL_SITE Then
                    .SiteNumber.Caption = "ALL"
                ElseIf lngData = NO_SITE Then
                    .SiteNumber.Caption = ""
                Else
                    .SiteNumber.Caption = lngData
                End If
        End Select
    End With
End Sub

Private Sub IParameterWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
End Sub

Private Sub IParameterWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
    With DcTestExaminForm
        Select Case dataName:
            Case MEASURE_FRM:
                .MeasureFrame.enabled = blnData
                .MeasureButton.enabled = blnData
                .LoopNumberLabel.enabled = blnData
                .LoopNumber.enabled = blnData
                .OutPutFileCheck.enabled = blnData
            Case RETURN_BTN: .ReturnButton.enabled = blnData
        End Select
    End With
End Sub

Public Function AsIParameterWriter() As IParameterWriter
    Set AsIParameterWriter = Me
End Function
