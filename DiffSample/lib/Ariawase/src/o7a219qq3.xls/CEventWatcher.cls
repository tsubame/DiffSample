VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEventWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   汎用ワークシートReaderのイベント監視クラス
'
'目的:
'   Readerはワークシート更新時にキャッシュを更新しないといけない
'   ワークシートのChangeイベントを監視し、Readerに通知する(EventをRaiseする)
'
'作成者:
'   a_oshima
'

Option Explicit

Private WithEvents mSheet As Excel.Worksheet
Attribute mSheet.VB_VarHelpID = -1
Private WithEvents mButton As MSForms.CommandButton
Attribute mButton.VB_VarHelpID = -1

Public Event WorksheetChange(rng As Range)
Public Event ButtonClick(Msg As String)

Private Sub Class_Initialize()
    Set mSheet = Nothing
    Set mButton = Nothing
End Sub

Private Sub Class_Terminate()
    Set mSheet = Nothing
    Set mButton = Nothing
End Sub

Property Set SetSheet(ByRef pSht As Excel.Worksheet)
    Set mSheet = pSht
    If mSheet.OLEObjects.Count > 0 Then
        Dim obj As OLEObject
        For Each obj In mSheet.OLEObjects
            If obj.Name Like "CommandButton1" Then
                Set SetButton = ThisWorkbook.Worksheets(mSheet.Name).CommandButton1
                'Set SetButton = obj   '型不一致エラー
                Exit For
            End If
        Next obj
        Set obj = Nothing
    Else
        Set SetButton = Nothing
    End If
End Property
Property Set SetButton(ByRef pButton As MSForms.CommandButton)
    Set mButton = pButton
End Property

Private Sub mSheet_Change(ByVal pTarget As Range)
    'MsgBox "Sheet Changed: " & pTarget.Address
    RaiseEvent WorksheetChange(pTarget)
End Sub
Private Sub mButton_Click()
    'MsgBox "Button Clicked!"
    RaiseEvent ButtonClick("Yeah!")
End Sub


