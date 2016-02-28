VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UsingPlaneDisplay 
   Caption         =   "Using Plane List"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   OleObjectBlob   =   "UsingPlaneDisplay.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UsingPlaneDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












Option Explicit

Private m_List() As String
Private m_Active As Boolean

Private m_IsCmBoxBaseEnable As Boolean

'=======2010/03/26 Add Maruyama For V202ここから==============
'VBランタイムの影響によってListViewが消える件について
'動的作成で対応する
Private WithEvents LV_Planes As MSComctlLib.ListView
Attribute LV_Planes.VB_VarHelpID = -1
Private WithEvents LV_Bit As MSComctlLib.ListView
Attribute LV_Bit.VB_VarHelpID = -1
Private Const LV_PLANES_IDENTIFIER As String = "LV_Planes"
Private Const LV_BIT_IDENTIFIER As String = "LV_Bit"
Private Const LV_BIT_WIDTH As Long = 150
'=======2010/03/26 Add Maruyama For V202 ここまで==============

Public Sub Display()
    Show vbModeless
    m_Active = True
    While m_Active = True
        DoEvents
    Wend
    Unload Me
End Sub

Private Sub CB_OK_Click()
    m_Active = False
End Sub

Private Sub LV_Planes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With LV_Planes
        .SortKey = ColumnHeader.index - 1
        .SortOrder = .SortOrder Xor lvwDescending
        .Sorted = True
    End With
End Sub

'#FlagExpansion
Private Sub LV_Planes_DblClick()

    If LV_Planes.SelectedItem Is Nothing Then Exit Sub
    
    With LV_Planes.SelectedItem
        Call theidv.OpenForm
        theidv.PlaneNameGreen = .Text
        theidv.PMDName = .SubItems(2)
'        Call Theidv.Refresh
    End With
    
    Dim tmp As ListItem
    Set tmp = LV_Bit.SelectedItem
    If Not (tmp Is Nothing) Then
        If tmp.Checked = True Then
            Dim planeGroup As String
            planeGroup = Left(CmBoxBase.Value, InStr(CmBoxBase.Value, vbTab) - 1)
            
            Dim FlagPlaneName As String '選択されているフラグプレーン名を作成
            FlagPlaneName = Right(CmBoxBase.Value, Len(CmBoxBase.Value) - InStr(CmBoxBase.Value, vbTab))
            FlagPlaneName = Replace(Replace(FlagPlaneName, ")", vbNullString), "(", vbNullString)
            Dim FlagPlane As IImgFlag
            
            For Each FlagPlane In TheIDP.PlaneManager(planeGroup).GetSharedFlagPlanes
                If planeGroup <> "" And FlagPlane.Name = FlagPlaneName Then
                    theidv.PlaneNameFlag = FlagPlane.Name
                    theidv.MsbFlag = LV_Bit.SelectedItem.Text
                    Exit For
                End If
            Next
        End If
    End If
    
    theidv.HilightColor = "Red" '2009/04/16 Maruyama デフォルト色変更
    
    Call theidv.Refresh
    
End Sub

Private Sub UserForm_Initialize()

    Call AddListViewControl '2010/03/26 Add Maruyama For V202
    
    With LV_Planes
'        .View = lvwReport '->AddListViewControlへ移動
        Call .ColumnHeaders.Add(, "Plane Name", "Plane Name")
        Call .ColumnHeaders.Add(, "Group", "Group")
        Call .ColumnHeaders.Add(, "Current PMD", "Current PMD")
        Call .ColumnHeaders.Add(, "Comment", "Comment")
    End With
    '=======2009/04/02 Add Maruyama ここから==============
    With LV_Bit
'        .View = lvwReport '->AddListViewControlへ移動
'        .CheckBoxes = True '->AddListViewControlへ移動
        Call .ColumnHeaders.Add(, "Bit", "Bit")
        Call .ColumnHeaders.Add(, "Flag Name", "Flag Name")
        
        '2009/04/16 Maruyama 列幅変更
        .ColumnHeaders(1).width = LV_BIT_WIDTH / 5
        .ColumnHeaders(2).width = LV_BIT_WIDTH / 5 * 4
        
    End With
    m_IsCmBoxBaseEnable = False
    '=======2009/04/02 Add Maruyama ここまで==============
    
    Call CB_Refresh_Click
End Sub

Private Sub UserForm_Terminate()
    
    Call DeleteListViewContorl '2010/03/26 Add Maruyama For V202
    
    m_Active = False
End Sub

'#FlagExpansion
Private Sub CB_Refresh_Click()

    '*****Using Plane List *******************************************************************
    LV_Planes.ListItems.Clear
    
    Dim arrUsingPlanes As Collection
    Set arrUsingPlanes = TheIDP.DumpUsingPlane
    If arrUsingPlanes.Count <= 0 Then Exit Sub
    
    '=======2009/03/31 Add Maruyama ここから==============
    'PlaneBankからリストを取得
    Dim PlaneList As Variant
    If TheIDP.PlaneBank.Count <> 0 Then
        PlaneList = Split(Replace(TheIDP.PlaneBank.List, vbCrLf, ","), ",")
    Else
        Set PlaneList = Nothing
    End If
    '=======2009/03/31 Add Maruyama ここまで==============
    
    Dim p As CImgPlane
    For Each p In arrUsingPlanes
        If IsAddLvItem(p, PlaneList) Then '2009/03/31 Add Maruyama ==============
            With LV_Planes.ListItems.Add
                .Text = p.Name
                .SubItems(1) = p.planeGroup
                .SubItems(2) = p.CurrentPmdName
                .SubItems(3) = p.Comment
            End With
        End If '2009/03/31 Add Maruyama ==============
    Next p
    
    
    '*****Using FlagBit List *****************************************************************
    '=======2009/04/02 Add Maruyama ここから==============
    m_IsCmBoxBaseEnable = False
    
    CmBoxBase.Clear
    LV_Bit.ListItems.Clear
    
    Dim bases As Collection
    Set bases = TheIDP.DumpPlaneGroup
    
    'プレングループがなければお手上げ
    If bases.Count = 0 Then Exit Sub
    
    'DropDownListに追加
    'SharedFlagがないプレングループは追加しない
    Dim Base As Variant
'    Dim flg As CImgFlag    '型を Interface に変更
    
    Dim flg As IImgFlag
    For Each Base In bases
        For Each flg In TheIDP.PlaneManager(Base).GetSharedFlagPlanes
            If flg.Name <> "" Then
                CmBoxBase.AddItem Base & vbTab & "(" & flg.Name & ")" '2009/04/12 Maruyama 変更
            End If
        Next
    Next Base
    
    m_IsCmBoxBaseEnable = True
    
    '先頭にあわせる
    If CmBoxBase.ListCount <> 0 Then
        CmBoxBase.ListIndex = 0
    End If
    
    '=======2009/04/02 Add Maruyama ここまで==============
    
    
End Sub

'=======2009/03/31 Add Maruyama ここから==============
Private Sub CBox_ShowPlaneBankPlane_Click()
    CB_Refresh_Click
End Sub

'#FlagExpansion
Private Function IsAddLvItem(ByVal plane As CImgPlane, ByRef PlaneList As Variant)
    
    'チェックされていないときは全部表示
    If CBox_ShowPlaneBankPlane.Value = True Then
        IsAddLvItem = True
        Exit Function
    End If
    
    'オブジェクトの時はPlaneBankにひとつもないときなので心おきなく追加
    If IsObject(PlaneList) Then
        IsAddLvItem = True
        Exit Function
    End If
    
    'PlaneBankに同じやつが見つかったら、追加しない
    Dim i As Long
    For i = 1 To UBound(PlaneList) Step 2
        If plane.Name = PlaneList(i) Then
            IsAddLvItem = False
            Exit Function
        End If
    Next i
    
    'SharedFlagも追加しない
    Dim FlagPlane As IImgFlag
    For Each FlagPlane In plane.Manager.GetSharedFlagPlanes
        If plane.Name = FlagPlane.Name Then
            IsAddLvItem = False
            Exit Function
        End If
    Next
    
    IsAddLvItem = True
    
End Function
'=======2009/03/31 Add Maruyama ここまで==============

'=======2009/04/02 Add Maruyama ここから==============
Private Sub LV_Bit_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With LV_Bit
        .SortKey = ColumnHeader.index - 1
        .SortOrder = .SortOrder Xor lvwDescending
        .Sorted = True
    End With
End Sub

'#FlagExpansion
Private Sub CmBoxBase_Change()
    
    If Not m_IsCmBoxBaseEnable Then Exit Sub
    
    'ListViewをクリアする
    LV_Bit.ListItems.Clear
    
    'SharedFlagが存在している状態から　存在していない状態に遷移すると
    'コンボボックスが空になるかもしれない
    If CmBoxBase.Value = "" Then Exit Sub
    
    'プレングループがSharedFlagを持っているかのチェック。この前にやってるはずなので不要か？
    Dim flg As IImgFlag
    Dim planeGroup As String
    planeGroup = Left(CmBoxBase.Value, InStr(CmBoxBase.Value, vbTab) - 1)
            
    Dim FlagPlaneName As String '選択されているフラグプレーン名を作成
    FlagPlaneName = Right(CmBoxBase.Value, Len(CmBoxBase.Value) - InStr(CmBoxBase.Value, vbTab))
    FlagPlaneName = Replace(Replace(FlagPlaneName, ")", vbNullString), "(", vbNullString)
                        
    For Each flg In TheIDP.PlaneManager(planeGroup).GetSharedFlagPlanes
        If flg.Name <> "" And flg.Name = FlagPlaneName Then
            'FlagListを取得し配列化
            Dim PlaneList As Variant
            PlaneList = Split(Replace(flg.FlagBitList, vbCrLf, ":"), ":")
            'ListViewに追加
            Dim i As Long
            For i = 0 To UBound(PlaneList) - 1 Step 2
                With LV_Bit.ListItems.Add
                    .Text = PlaneList(i)
                    .SubItems(1) = PlaneList(i + 1)
                End With
            Next i
        End If
    Next
        
End Sub

Private Sub LV_Bit_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    '同じアイテムに対してチェックボタンを押してFlaseになった場合
    'イベントはチェックボタンを変更してから呼ばれる
    If (Item Is LV_Bit.SelectedItem) And (LV_Bit.SelectedItem.Checked = False) Then
        Item.Selected = False
        Item.Checked = False
        Exit Sub
    End If

    Dim tmp As ListItem
    For Each tmp In LV_Bit.ListItems
        tmp.Checked = False
        tmp.Selected = False
    Next tmp
    
    Item.Selected = True
    Item.Checked = True

End Sub

Private Sub LV_Bit_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    '同じアイテムに対してクリックを押した場合
    '本来アイテムクリックの動作はチェックボックスとは何の関係もない。
    'なのでチェックされていたら、はずすという処理を行う
    If (Item Is LV_Bit.SelectedItem) And (LV_Bit.SelectedItem.Checked = True) Then
        Item.Checked = False
        Exit Sub
    End If
    
    Dim tmp As ListItem
    For Each tmp In LV_Bit.ListItems
        tmp.Checked = False
    Next tmp
    
    Item.Checked = True
    
End Sub
'=======2009/04/02 Add Maruyama ここまで==============

'=======2010/03/26 Add Maruyama For V202ここから==============
Private Sub AddListViewControl()

    Dim tmpLvPlane As Control
    Set tmpLvPlane = Me.Controls.Add("MSComctlLib.ListViewCtrl", LV_PLANES_IDENTIFIER, True)

    Set LV_Planes = tmpLvPlane
    With LV_Planes
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
    End With
    
    With tmpLvPlane
        .Top = 24
        .width = 336
        .Left = 24
        .height = 192
    End With
    
    Set tmpLvPlane = Nothing


    Dim tmpLvBit As Control
    Set tmpLvBit = Me.Controls.Add("MSComctlLib.ListViewCtrl", LV_BIT_IDENTIFIER, True)
    
    Set LV_Bit = tmpLvBit
    With LV_Bit
        .View = lvwReport
        .FullRowSelect = True
        .CheckBoxes = True
        .LabelEdit = lvwManual
    End With
    
    With tmpLvBit
        .Top = 54
        .width = LV_BIT_WIDTH
        .Left = 378
        .height = 156
    End With
    
    Set tmpLvBit = Nothing

End Sub

Private Sub DeleteListViewContorl()

    Call Me.Controls.Remove(LV_PLANES_IDENTIFIER)
    Call Me.Controls.Remove(LV_BIT_IDENTIFIER)
    Set LV_Planes = Nothing
    Set LV_Bit = Nothing

End Sub

'=======2010/03/26 Add Maruyama For V202 ここまで==============

