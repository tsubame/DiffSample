VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlaneBankDisplay 
   Caption         =   "Plane Bank Display"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "PlaneBankDisplay.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "PlaneBankDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit

Private m_Active As Boolean

'=======2010/03/26 Add Maruyama For V202ここから==============
'VBランタイムの影響によってListViewが消える件について
'動的作成で対応する
Private WithEvents LV_Display As MSComctlLib.ListView
Attribute LV_Display.VB_VarHelpID = -1
Private Const LV_DISPLAY_IDENTIFIER As String = "LV_Display"
'=======2010/03/26 Add Maruyama For V202 ここまで==============

Private Sub CB_OK_Click()
    m_Active = False
'    Unload Me
End Sub

Public Sub Display()
    Show vbModeless
    m_Active = True
    While m_Active = True
        DoEvents
    Wend
    Unload Me
End Sub

Private Sub LV_Display_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With LV_Display
        .SortKey = ColumnHeader.index - 1
        .SortOrder = .SortOrder Xor lvwDescending
        .Sorted = True
    End With
End Sub

Private Sub LV_Display_DblClick()
    
    If LV_Display.SelectedItem Is Nothing Then Exit Sub
    
    With LV_Display.SelectedItem
        Call theidv.OpenForm
        theidv.PlaneNameGreen = .SubItems(1)
        Call theidv.Refresh
    End With
End Sub

Private Sub UserForm_Initialize()
    
    Call AddListViewControl '2010/03/26 Add Maruyama For V202
    
    With LV_Display
'        .View = lvwReport '->AddListViewControlへ移動
        Call .ColumnHeaders.Add(, "Name", "Name")
        Call .ColumnHeaders.Add(, "Plane Name", "Plane Name")
    End With
    
    Call Refresh
End Sub

Private Sub UserForm_Terminate()
    
    Call DeleteListViewContorl '2010/03/26 Add Maruyama For V202
    m_Active = False

End Sub

Private Sub Refresh()

    Dim listArr()
    Dim PlaneList As Variant
        
    LV_Display.ListItems.Clear
    
    If TheIDP.PlaneBank.Count = 0 Then Exit Sub
    
    PlaneList = Split(Replace(TheIDP.PlaneBank.List, vbCrLf, ","), ",")
    ReDim listArr(TheIDP.PlaneBank.Count - 1, 1)
    
    Dim i As Long
    For i = 0 To UBound(listArr, 1)
        listArr(i, 0) = PlaneList(2 * i + 0)
        listArr(i, 1) = PlaneList(2 * i + 1)
    Next i
    
    For i = 0 To UBound(listArr, 1)
        With LV_Display.ListItems.Add
            .Text = PlaneList(2 * i + 0)
            .SubItems(1) = PlaneList(2 * i + 1)
        End With
    Next i
End Sub

'=======2010/03/26 Add Maruyama For V202ここから==============
Private Sub AddListViewControl()

    Dim tmpCntrl As Control
    Set tmpCntrl = Me.Controls.Add("MSComctlLib.ListViewCtrl", LV_DISPLAY_IDENTIFIER, True)

    Set LV_Display = tmpCntrl
    With LV_Display
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
    End With
    
    With tmpCntrl
        .Top = 18
        .width = 210
        .Left = 12.75
        .height = 90
    End With
    
    Set tmpCntrl = Nothing
    
End Sub

Private Sub DeleteListViewContorl()

    Call Me.Controls.Remove(LV_DISPLAY_IDENTIFIER)
    Set LV_Display = Nothing

End Sub
'=======2010/03/26 Add Maruyama For V202 ここまで==============


