VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UsingPlaneDisplay 
   Caption         =   "Using Plane List"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   OleObjectBlob   =   "UsingPlaneDisplay.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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

'=======2010/03/26 Add Maruyama For V202��������==============
'VB�����^�C���̉e���ɂ����ListView�������錏�ɂ���
'���I�쐬�őΉ�����
Private WithEvents LV_Planes As MSComctlLib.ListView
Attribute LV_Planes.VB_VarHelpID = -1
Private WithEvents LV_Bit As MSComctlLib.ListView
Attribute LV_Bit.VB_VarHelpID = -1
Private Const LV_PLANES_IDENTIFIER As String = "LV_Planes"
Private Const LV_BIT_IDENTIFIER As String = "LV_Bit"
Private Const LV_BIT_WIDTH As Long = 150
'=======2010/03/26 Add Maruyama For V202 �����܂�==============

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
            
            Dim FlagPlaneName As String '�I������Ă���t���O�v���[�������쐬
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
    
    theidv.HilightColor = "Red" '2009/04/16 Maruyama �f�t�H���g�F�ύX
    
    Call theidv.Refresh
    
End Sub

Private Sub UserForm_Initialize()

    Call AddListViewControl '2010/03/26 Add Maruyama For V202
    
    With LV_Planes
'        .View = lvwReport '->AddListViewControl�ֈړ�
        Call .ColumnHeaders.Add(, "Plane Name", "Plane Name")
        Call .ColumnHeaders.Add(, "Group", "Group")
        Call .ColumnHeaders.Add(, "Current PMD", "Current PMD")
        Call .ColumnHeaders.Add(, "Comment", "Comment")
    End With
    '=======2009/04/02 Add Maruyama ��������==============
    With LV_Bit
'        .View = lvwReport '->AddListViewControl�ֈړ�
'        .CheckBoxes = True '->AddListViewControl�ֈړ�
        Call .ColumnHeaders.Add(, "Bit", "Bit")
        Call .ColumnHeaders.Add(, "Flag Name", "Flag Name")
        
        '2009/04/16 Maruyama �񕝕ύX
        .ColumnHeaders(1).width = LV_BIT_WIDTH / 5
        .ColumnHeaders(2).width = LV_BIT_WIDTH / 5 * 4
        
    End With
    m_IsCmBoxBaseEnable = False
    '=======2009/04/02 Add Maruyama �����܂�==============
    
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
    
    '=======2009/03/31 Add Maruyama ��������==============
    'PlaneBank���烊�X�g���擾
    Dim PlaneList As Variant
    If TheIDP.PlaneBank.Count <> 0 Then
        PlaneList = Split(Replace(TheIDP.PlaneBank.List, vbCrLf, ","), ",")
    Else
        Set PlaneList = Nothing
    End If
    '=======2009/03/31 Add Maruyama �����܂�==============
    
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
    '=======2009/04/02 Add Maruyama ��������==============
    m_IsCmBoxBaseEnable = False
    
    CmBoxBase.Clear
    LV_Bit.ListItems.Clear
    
    Dim bases As Collection
    Set bases = TheIDP.DumpPlaneGroup
    
    '�v�����O���[�v���Ȃ���΂���グ
    If bases.Count = 0 Then Exit Sub
    
    'DropDownList�ɒǉ�
    'SharedFlag���Ȃ��v�����O���[�v�͒ǉ����Ȃ�
    Dim Base As Variant
'    Dim flg As CImgFlag    '�^�� Interface �ɕύX
    
    Dim flg As IImgFlag
    For Each Base In bases
        For Each flg In TheIDP.PlaneManager(Base).GetSharedFlagPlanes
            If flg.Name <> "" Then
                CmBoxBase.AddItem Base & vbTab & "(" & flg.Name & ")" '2009/04/12 Maruyama �ύX
            End If
        Next
    Next Base
    
    m_IsCmBoxBaseEnable = True
    
    '�擪�ɂ��킹��
    If CmBoxBase.ListCount <> 0 Then
        CmBoxBase.ListIndex = 0
    End If
    
    '=======2009/04/02 Add Maruyama �����܂�==============
    
    
End Sub

'=======2009/03/31 Add Maruyama ��������==============
Private Sub CBox_ShowPlaneBankPlane_Click()
    CB_Refresh_Click
End Sub

'#FlagExpansion
Private Function IsAddLvItem(ByVal plane As CImgPlane, ByRef PlaneList As Variant)
    
    '�`�F�b�N����Ă��Ȃ��Ƃ��͑S���\��
    If CBox_ShowPlaneBankPlane.Value = True Then
        IsAddLvItem = True
        Exit Function
    End If
    
    '�I�u�W�F�N�g�̎���PlaneBank�ɂЂƂ��Ȃ��Ƃ��Ȃ̂ŐS�����Ȃ��ǉ�
    If IsObject(PlaneList) Then
        IsAddLvItem = True
        Exit Function
    End If
    
    'PlaneBank�ɓ����������������A�ǉ����Ȃ�
    Dim i As Long
    For i = 1 To UBound(PlaneList) Step 2
        If plane.Name = PlaneList(i) Then
            IsAddLvItem = False
            Exit Function
        End If
    Next i
    
    'SharedFlag���ǉ����Ȃ�
    Dim FlagPlane As IImgFlag
    For Each FlagPlane In plane.Manager.GetSharedFlagPlanes
        If plane.Name = FlagPlane.Name Then
            IsAddLvItem = False
            Exit Function
        End If
    Next
    
    IsAddLvItem = True
    
End Function
'=======2009/03/31 Add Maruyama �����܂�==============

'=======2009/04/02 Add Maruyama ��������==============
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
    
    'ListView���N���A����
    LV_Bit.ListItems.Clear
    
    'SharedFlag�����݂��Ă����Ԃ���@���݂��Ă��Ȃ���ԂɑJ�ڂ����
    '�R���{�{�b�N�X����ɂȂ邩������Ȃ�
    If CmBoxBase.Value = "" Then Exit Sub
    
    '�v�����O���[�v��SharedFlag�������Ă��邩�̃`�F�b�N�B���̑O�ɂ���Ă�͂��Ȃ̂ŕs�v���H
    Dim flg As IImgFlag
    Dim planeGroup As String
    planeGroup = Left(CmBoxBase.Value, InStr(CmBoxBase.Value, vbTab) - 1)
            
    Dim FlagPlaneName As String '�I������Ă���t���O�v���[�������쐬
    FlagPlaneName = Right(CmBoxBase.Value, Len(CmBoxBase.Value) - InStr(CmBoxBase.Value, vbTab))
    FlagPlaneName = Replace(Replace(FlagPlaneName, ")", vbNullString), "(", vbNullString)
                        
    For Each flg In TheIDP.PlaneManager(planeGroup).GetSharedFlagPlanes
        If flg.Name <> "" And flg.Name = FlagPlaneName Then
            'FlagList���擾���z��
            Dim PlaneList As Variant
            PlaneList = Split(Replace(flg.FlagBitList, vbCrLf, ":"), ":")
            'ListView�ɒǉ�
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

    '�����A�C�e���ɑ΂��ă`�F�b�N�{�^����������Flase�ɂȂ����ꍇ
    '�C�x���g�̓`�F�b�N�{�^����ύX���Ă���Ă΂��
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
    
    '�����A�C�e���ɑ΂��ăN���b�N���������ꍇ
    '�{���A�C�e���N���b�N�̓���̓`�F�b�N�{�b�N�X�Ƃ͉��̊֌W���Ȃ��B
    '�Ȃ̂Ń`�F�b�N����Ă�����A�͂����Ƃ����������s��
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
'=======2009/04/02 Add Maruyama �����܂�==============

'=======2010/03/26 Add Maruyama For V202��������==============
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

'=======2010/03/26 Add Maruyama For V202 �����܂�==============

