VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNavigationCommander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �i�r�Q�[�V����GUI�̃��b�p�[�N���X
'   �i�r�Q�[�V�������j���[�̋@�\�g���ƃ��j���[�N���b�N���̃C�x���g���󂯎������������s��
'
'�ړI:
'   �@�i�r�Q�[�V�������j���[�̍\�z�ƃ}�N���֐��̓o�^
'   �A�f�[�^�c���[�̕\���y�уf�[�^�q�X�g���̕\��
'   �B���j���[����̃��[�U�[�v�����󂯕t���O���֗v�����o��
'
'   Revision History:
'   Data        Description
'   2008/12/11  �쐬
'   2009/02/20�@���d�l�ύX
'               ���c���[�\���̃��W�b�N��ύX
'               �@�f�[�^�V�[�g�������Ȃ��c�[���J�e�S���̓c���[�r���[�ɔ�\��
'               �A�f�[�^�V�[�g��������Ȃ��c�[���������ӂɌ��܂�ꍇ�̓c�[���J�e�S�����C���[�Ƀc�[����+�V�[�g���Ń{�^���\��
'               �B�c�[���J�e�S������\���ݒ�̏ꍇ�̓c���[�r���[�ɕ\�����Ȃ�
'               �C�c�[���J�e�S���̃J�e�S���C�Y�ݒ肪�I�t�̏ꍇ�̓c�[���J�e�S�����C���[�ɃV�[�g���Ń{�^���\��
'               ���q�X�g���\���̃��W�b�N��ύX
'               �@�f�[�^�V�[�g���폜����Ă���ꍇ�̕\����ǉ�
'               �A�A�N�e�B�u�v���p�e�B�擾�ƃi�r�Q�[�^�̃C���f�b�N�X�擾�ɂ��J�����g�f�[�^�V�[�g�𔻕ʂ��郍�W�b�N�ɕύX
'               �B�\������f�[�^�V�[�g�Ƀ��[�N�V�[�g�ȊO�̃V�[�g�I�u�W�F�N�g���܂߂�
'               ���G�N�Z���I�u�W�F�N�g�̎擾
'               �@�����ŃV���[�g�J�b�g�L�[�Ƀ}�N���o�^���s���i�i�ރ{�^���A�߂�{�^���j
'               �A�@�ɔ����A�}�N���o�^�p�ɐi�ރ{�^���A�߂�{�^������p���\�b�h���O�����J
'   2009/05/11�@���d�l�ύX
'               �@Create���\�b�h�̒ǉ��FInitialize�ōs���Ă����J�X�^���o�[�쐬�������ڊ�
'               �ADestroy���\�b�h�̒ǉ��FCreate���\�b�h�ɑ΂��郁�j���[�o�[�폜�@�\�Ƃ��Ēǉ�
'                                        �V���[�g�J�b�g�L�[�̍폜�͂����ł̂ݍs����
'
'�쐬��:
'   0145206097
'
Option Explicit

Private oXL As Excel.Application

Private mNavigator As IDataNavigator
Private mExplorer As IDataExplorer

Private mNaviCommandBar As Office.CommandBar
Private mDataTreeMenu As Office.CommandBarPopup
Private mHistoryMenu As Office.CommandBarPopup
Private WithEvents mForeButton As Office.CommandBarButton
Attribute mForeButton.VB_VarHelpID = -1
Private WithEvents mBackButton As Office.CommandBarButton
Attribute mBackButton.VB_VarHelpID = -1

Private Const NAVI_BAR_NAME = "EeeNavi"

Private Const DATATREE_BAR_MENU_CAPTION = "EeeNavi(&X)"
Private Const DATATREE_BAR_MENU_POPUP = "Data Sheet Browser"
Private Const DATATREE_BAR_MENU_ACTION = "Commander_DataTreeMenu_Events"

Private Const DATATREE_BAR_MENU_LIST_ACTION = "OpenDataFolder"
Private Const DATATREE_BAR_MENU_LIST_ICON = 32
Private Const DATATREE_BAR_MENU_BUTTON_ACTION = "Commander_DataTreeMenuButton_Events"
Private Const DATATREE_BAR_MENU_BUTTON_ICON = 142

Private Const HISTORY_BAR_FORE_BUTTON_POPUP = "Forward"
Private Const HISTORY_BAR_FORE_BUTTON_ICON = 133
Private Const HISTORY_BAR_BACK_BUTTON_POPUP = "Back"
Private Const HISTORY_BAR_BACK_BUTTON_ICON = 132

Private Const HISTORY_BAR_MENU_CAPTION = " "
Private Const HISTORY_BAR_MENU_POPUP = "History Browser"
Private Const HISTORY_BAR_MENU_ACTION = "Commander_HistoryMenu_Events"

Private Const HISTORY_BAR_MENU_BUTTON_CURRENT_CAPTION = "Current Browsing"
Private Const HISTORY_BAR_MENU_BUTTON_ACTION = "Commander_HistoryMenuButton_Events"
Private Const HISTORY_BAR_MENU_BUTTON_CURRENT_ICON = 184
Private Const HISTORY_BAR_MENU_BUTTON_FORE_ICON = 156
Private Const HISTORY_BAR_MENU_BUTTON_BACK_ICON = 155
Private Const HISTORY_BAR_MENU_BUTTON_DELETE_ICON = 358

Private Const FORE_BUTTON_SHORTCUT_KEY = "%{RIGHT}"
Private Const BACK_BUTTON_SHORTCUT_KEY = "%{LEFT}"
Private Const HISTORY_BAR_FORE_BUTTON_MACRO = "ShortCut_HistoryForeButton_Events"
Private Const HISTORY_BAR_BACK_BUTTON_MACRO = "ShortCut_HistoryBackButton_Events"

Private Sub Class_Initialize()
    Set mNavigator = New IDataNavigator
    Set mExplorer = New IDataExplorer
End Sub

Public Sub Initialize(ByVal excelApp As Excel.Application, ByVal navigator As IDataNavigator, ByVal explorer As IDataExplorer)
'���e:
'   �i�r�Q�[�V����GUI�I�u�W�F�N�g�̏�����
'
'�p�����[�^:
'[excelApp]   In  �G�N�Z���I�u�W�F�N�g
'[navigator]  In  �i�r�Q�[�V�����I�u�W�F�N�g
'[explorer]   In  �G�N�X�v���[���[�I�u�W�F�N�g
'
'���ӎ���:
'
    Set oXL = excelApp
    Set mNavigator = navigator
    Set mExplorer = explorer
End Sub

Public Sub Create()
'���e:
'   �i�r�Q�[�V����GUI�̍쐬
'
'���ӎ���:
'
    '### ���Ƀ��j���[�o�[�����݂���ꍇ�͍폜���� ###################
    Me.Destroy
    oXL.ScreenUpdating = False
    '### �i�r�Q�[�V�����p���j���[�o�[��V�K�ō쐬 ###################
    Set mNaviCommandBar = oXL.CommandBars.Add(Name:=NAVI_BAR_NAME, Position:=msoBarTop, temporary:=True)
    mNaviCommandBar.Visible = True
    '### �f�[�^�c���[�p�A�q�X�g���p�e���j���[��ǉ� #################
    With mNaviCommandBar.Controls
        Set mDataTreeMenu = .Add(Type:=msoControlPopup)
        Set mBackButton = .Add(Type:=msoControlButton)
        Set mForeButton = .Add(Type:=msoControlButton)
        Set mHistoryMenu = .Add(Type:=msoControlPopup)
    End With
    '### �f�[�^�c���[���j���̐ݒ�Ǝ��s�}�N���ǉ� ###################
    With mDataTreeMenu
         .Caption = DATATREE_BAR_MENU_CAPTION
        .TooltipText = DATATREE_BAR_MENU_POPUP
        .OnAction = DATATREE_BAR_MENU_ACTION
    End With
    '### �q�X�g���{�^���i�i�ށj�̐ݒ� ###############################
    With mForeButton
        .FaceId = HISTORY_BAR_FORE_BUTTON_ICON
        .Style = msoButtonIcon
        .TooltipText = HISTORY_BAR_FORE_BUTTON_POPUP
        .enabled = False
    End With
    '### �q�X�g���{�^���i�߂�j�̐ݒ� ###############################
    With mBackButton
        .BeginGroup = True
        .FaceId = HISTORY_BAR_BACK_BUTTON_ICON
        .Style = msoButtonIcon
        .TooltipText = HISTORY_BAR_BACK_BUTTON_POPUP
        .enabled = False
    End With
    '### �q�X�g�����j���[�̐ݒ�Ǝ��s�}�N���ǉ� #####################
    With mHistoryMenu
        .Caption = HISTORY_BAR_MENU_CAPTION
        .TooltipText = HISTORY_BAR_MENU_POPUP
        .OnAction = HISTORY_BAR_MENU_ACTION
    End With
    '### �f�[�^�c���[�\���p�̃��j���[�����ݒ� #######################
    createDataTreeMenu
    '### �V���[�g�J�b�g�L�[�Ǝ��s�}�N���̓o�^ #######################
    With oXL
        .OnKey BACK_BUTTON_SHORTCUT_KEY, HISTORY_BAR_BACK_BUTTON_MACRO
        .OnKey FORE_BUTTON_SHORTCUT_KEY, HISTORY_BAR_FORE_BUTTON_MACRO
    End With
    oXL.ScreenUpdating = True
End Sub

Public Sub Destroy()
'���e:
'   �i�r�Q�[�V����GUI�̔j��
'
'���ӎ���:
'
    With oXL
        On Error Resume Next
        .CommandBars(NAVI_BAR_NAME).Delete
        .OnKey BACK_BUTTON_SHORTCUT_KEY
        .OnKey FORE_BUTTON_SHORTCUT_KEY
        On Error GoTo 0
    End With
End Sub

Public Sub DisplayDataTreeMenu()
'���e:
'   �f�[�^�c���[�̃��j���[�\��
'
'���ӎ���:
'
    '### �f�[�^�c���[�̃f�[�^�V�[�g�{�^�����ŏ��ɍ폜 ###############
    clearDataTreeMenu
    '### �G�N�X�v���[���[����f�[�^�c���[�̎擾 #####################
    Dim treeData As Collection
    Set treeData = mExplorer.DataTree
    '### �f�[�^�c���[�̃g�b�v���j���[�̐ݒ� #########################
    Dim topLayer As CToolsupplier
    For Each topLayer In treeData
        Dim topMenu As Office.CommandBarPopup
        Set topMenu = mDataTreeMenu.Controls(topLayer.Name)
        '### �f�[�^���Ȃ��ꍇ�̓g�b�v���j���[��\�����Ȃ� ###########
        topMenu.Visible = topLayer.IsContain
        If topMenu.Visible Then
            '### �f�[�^�c���[�̃T�u���j���[�̐ݒ� ###################
            Dim subLayer As CToolCategory
            For Each subLayer In topLayer.DataTools
                Dim subMenu As Office.CommandBarPopup
                Set subMenu = topMenu.Controls(subLayer.Name)
                '### �f�[�^���Ȃ��ꍇ�A�܂��͕\���ݒ肪������ #######
                '### �Ȃ��Ă���ꍇ�̓T�u���j���[��\�����Ȃ� #######
                With subLayer
                    subMenu.Visible = (.IsContain And .Visible)
                End With
                If subMenu.Visible Then
                    '### ���j���[�ȉ��Ƀ{�^���A�C�e�����쐬 #########
                    Dim dataItem As CDataSheet
                    For Each dataItem In subLayer.DataSheets
                        Dim menuItem As Office.CommandBarButton
                        Set menuItem = subMenu.Controls.Add(Type:=msoControlButton, Before:=1)
                        '### �{�^���̐ݒ�Ǝ��s�}�N���̒ǉ� #########
                        With menuItem
                            .Caption = dataItem.Name
                            .FaceId = DATATREE_BAR_MENU_BUTTON_ICON
                            .OnAction = "'Commander_DataTreeMenuButton_Events " & Chr(34) & dataItem.Name & Chr(34) & "'"
                            If dataItem.IsActive Then
                                .state = msoButtonDown
                            Else
                                .state = msoButtonUp
                            End If
                        End With
                    Next dataItem
                    '### �f�[�^�V�[�g���J�e�S�����Ń��j�[�N�Ɍ��܂�ꍇ�A
                    '### �܂��̓J�e�S���C�Y���s��Ȃ��ꍇ�͊K�w�\�����Ȃ�
                    With subLayer
                        If .IsUnique Or Not .Categorize Then
                            For Each menuItem In subMenu.Controls
                                If .IsUnique And .Categorize Then
                                    menuItem.Caption = .Name
                                    menuItem.ShortcutText = "[" & .DataSheets(1).Name & "]"
                                End If
                                menuItem.Move topMenu.CommandBar, subMenu.index
                            Next menuItem
                            subMenu.Visible = False
                        End If
                    End With
                End If
            Next subLayer
        End If
    Next topLayer
End Sub

Public Sub DisplayHistoryMenu()
'���e:
'   �f�[�^�q�X�g���̃��j���[�\��
'
'���ӎ���:
'
    '### �f�[�^�q�X�g���̃f�[�^�V�[�g�{�^�����ŏ��ɍ폜 #############
    clearHistoryMenu
    '### �i�r�Q�[�^�[����f�[�^�q�X�g���̎擾 #######################
    Dim dataList As Collection
    Dim histIdx As Long
    With mNavigator
        Set dataList = .History
        histIdx = .HistoryIndex
    End With
    '### �q�X�g�����Ȃ��ꍇ�̓{�^�����쐬���Ȃ� #####################
    If dataList Is Nothing Then Exit Sub
    '### �q�X�g���̃f�[�^�V�[�g�{�^���̍쐬 #########################
    Dim itemIdx As Long
    itemIdx = 1
    Dim dataItem As CDataSheet
    For Each dataItem In dataList
        Dim menuItem As Office.CommandBarButton
        Set menuItem = mHistoryMenu.Controls.Add(Type:=msoControlButton, Before:=1)
        '### �f�[�^�V�[�g�{�^���̐ݒ� ###############################
        With menuItem
            '### �A�N�e�B�u�̏ꍇ�̓{�^���𖳌��ɐݒ� ###############
            If (dataItem.IsActive And histIdx = itemIdx) Then
                .Caption = HISTORY_BAR_MENU_BUTTON_CURRENT_CAPTION
                .FaceId = HISTORY_BAR_MENU_BUTTON_CURRENT_ICON
                .state = msoButtonDown
                .enabled = False
            '### ����ȊO�͎��s�}�N���̒ǉ� #########################
            Else
                .Caption = dataItem.Name
                .OnAction = "'Commander_HistoryMenuButton_Events " & Chr(34) & (itemIdx) & Chr(34) & "'"
                .state = msoButtonUp
                If itemIdx > histIdx Then
                    .FaceId = HISTORY_BAR_MENU_BUTTON_FORE_ICON
                Else
                    .FaceId = HISTORY_BAR_MENU_BUTTON_BACK_ICON
                End If
                '### �f�[�^�V�[�g���폜����Ă���ꍇ�͖��� #########
                If dataItem.IsDeleted Then
                    .FaceId = HISTORY_BAR_MENU_BUTTON_DELETE_ICON
                    .enabled = False
                Else
                    .enabled = True
                End If
            End If
        End With
        itemIdx = itemIdx + 1
    Next dataItem
End Sub

Private Sub createDataTreeMenu()
'####################################################################
'�i�r�Q�[�V�����c���[�r���[�|�b�v�A�b�v���j���[�̏����ݒ���s��
'####################################################################
    '### �G�N�X�v���[���[����f�[�^�c���[�̎擾 #####################
    Dim treeData As Collection
    Set treeData = mExplorer.DataTree
    Dim topLayer As CToolsupplier
    '### �f�[�^�c���[�̃g�b�v���j���[�̍쐬 #########################
    For Each topLayer In treeData
        Dim topMenu As Office.CommandBarPopup
        Set topMenu = mDataTreeMenu.Controls.Add(Type:=msoControlPopup)
        topMenu.Caption = topLayer.Name
        '### �f�[�^�c���[�̃T�u���j���[�̍쐬 #######################
        Dim subLayer As CToolCategory
        For Each subLayer In topLayer.DataTools
            Dim subMenu As Office.CommandBarPopup
            Set subMenu = topMenu.Controls.Add(Type:=msoControlPopup)
            subMenu.Caption = subLayer.Name
        Next subLayer
    Next topLayer
End Sub

Private Sub clearDataTreeMenu()
'####################################################################
'�i�r�Q�[�V�����c���[�r���[���j���[�{�^�����폜����
'####################################################################
    '### �f�[�^�c���[�g�b�v���j���[�̎擾 ###########################
    Dim topMenu As Office.CommandBarPopup
    For Each topMenu In mDataTreeMenu.Controls
        With topMenu
            Dim menuCount As Long
            menuCount = .Controls.Count
            Dim subIdx As Long
            For subIdx = menuCount To 1 Step -1
                Dim itemType As Long
                '### �T�u���j���[�̃^�C�v�擾 #######################
                itemType = .Controls(subIdx).Type
                If itemType = Office.msoControlPopup Then
                    '### ����ɃT�u���j���[������Ύ擾 #############
                    Dim subMenu As Office.CommandBarPopup
                    Set subMenu = .Controls(subIdx)
                    Dim itemCount As Long
                    itemCount = subMenu.Controls.Count
                    '### �T�u���j���[�ȉ��̃{�^���폜 ###############
                    Dim itemIdx As Long
                    For itemIdx = itemCount To 1 Step -1
                        subMenu.Controls(itemIdx).Delete
                    Next itemIdx
                '### �T�u���j���[���{�^����������폜 ###############
                ElseIf itemType = Office.msoControlButton Then
                    .Controls(subIdx).Delete
                End If
            Next subIdx
        End With
    Next topMenu
End Sub

Private Sub clearHistoryMenu()
'####################################################################
'�i�r�Q�[�V�����q�X�g�����j���[�{�^�����폜����
'####################################################################
    Dim menuItem As Office.CommandBarButton
    For Each menuItem In mHistoryMenu.Controls
        menuItem.Delete
    Next menuItem
End Sub

Private Sub mForeButton_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'####################################################################
'�i�r�Q�[�V�������j���[�o�[�̐i�ރ{�^���������ꂽ���ɔ�������C�x���g
'####################################################################
    '### �i�r�Q�[�^�����O�̃f�[�^�擾 ###########################
    Dim dataItem As CDataSheet
    Set dataItem = mNavigator.ForwardDataSheet
    '### �擾�ɐ���������C�x���g�}�N���Ăяo�� #####################
    If Not dataItem Is Nothing Then
        Commander_HistoryButton_Events dataItem.Name
    End If
    '### �������g�̃q�X�g�����j���[�X�e�[�^�X�̐ݒ� #################
    Me.SetHistoryButtonEnable
End Sub

Private Sub mBackButton_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'####################################################################
'�i�r�Q�[�V�������j���[�o�[�̖߂�{�^���������ꂽ���ɔ�������C�x���g
'####################################################################
    '### �i�r�Q�[�^������̃f�[�^�擾 ###########################
    Dim dataItem As CDataSheet
    Set dataItem = mNavigator.BackwardDataSheet
    '### �擾�ɐ���������C�x���g�}�N���Ăяo�� #####################
    If Not dataItem Is Nothing Then
        Commander_HistoryButton_Events dataItem.Name
    End If
    '### �������g�̃q�X�g�����j���[�X�e�[�^�X�̐ݒ� #################
    Me.SetHistoryButtonEnable
End Sub

Public Sub HistoryMenuButton_Click(ByVal hIndex As Long)
'���e:
'   �R�}���h���C������q�X�g�����j���[�̃{�^��������s�����\�b�h
'
'�p�����[�^:
'[hIndex]  In  �R���g���[���{�^���̃C���f�b�N�X�ԍ�
'
'���ӎ���:
'
    '### �i�r�Q�[�^����w��C���f�b�N�X�̃f�[�^�擾 #################
    Dim dataItem As CDataSheet
    Set dataItem = mNavigator.DataSheet(hIndex)
    '### �擾�ɐ���������C�x���g�}�N���Ăяo�� #####################
    If Not dataItem Is Nothing Then
        Commander_HistoryButton_Events dataItem.Name
    End If
    '### �������g�̃q�X�g�����j���[�X�e�[�^�X�̐ݒ� #################
    Me.SetHistoryButtonEnable
End Sub

Public Sub HistoryForeButton_Click()
'���e:
'   �R�}���h���C������u�i�ށv�{�^��������s�����\�b�h
'
'���ӎ���:
'
    mForeButton.Execute
End Sub

Public Sub HistoryBackButton_Click()
'���e:
'   �R�}���h���C������u�߂�v�{�^��������s�����\�b�h
'
'���ӎ���:
'
    mBackButton.Execute
End Sub

Public Sub SetHistoryButtonEnable()
'���e:
'   �u�i�ށv�u�߂�v�R���g���[���{�^���̃C�l�[�u���X�e�[�^�X��ݒ肷�郁�\�b�h
'
'���ӎ���:
'
    With mNavigator
        '### �u�i�ށv�{�^���̃C�l�[�u���ݒ� #########################
        mForeButton.enabled = .IsEnableForward
        '### �u�߂�v�{�^���̃C�l�[�u���ݒ� #########################
        mBackButton.enabled = .IsEnableBackward
    End With
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    oXL.CommandBars(NAVI_BAR_NAME).Delete
    On Error GoTo 0
    Set mNavigator = Nothing
    Set mExplorer = Nothing
End Sub