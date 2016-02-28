Attribute VB_Name = "XLibEeeNaviConstructor"
'�T�v:
'   Eee-JOB���[�N�V�[�g�i�r�Q�[�V�����N���X�̃R���X�g���N�^���W���[��
'
'�ړI:
'   �u�b�N�N�����Ƀi�r�Q�[�V�����̃X�^�[�g�A�b�v�J�X�^�����j���[�𐶐�����
'   �@�G�N�X�v���[���[�̃f�[�^�c���[���\�����邽�߂̊e�c�[�����₻�̔z�����̏�����
'   �@���G�N�X�v���[���[�̂��߂̃��[�_�[�I�u�W�F�N�g�𐶐�����[2009/02/20�ύX]
'   �A�i�r�Q�[�V���������s���邽�߂̃I�u�W�F�N�g�𐶐������������s��
'   �B�i�r�Q�[�V�������J�n���邽�߂̃��[�U�[�C���^�[�t�F�[�X�����[�N�V�[�g���j���[�o�[�֓o�^����
'   �C�c���[�r���[�o�͂̂��߂̃��C�^�[�I�u�W�F�N�g�𐶐�
'
'   Revision History:
'   Data        Description
'   2008/12/11�@�쐬
'   2008/12/15�@���@�\�ǉ�
'               �@�c���[�r���[�o�͋@�\�y�у��j���[�ւ̒ǉ��o�^
'   2009/01/16�@���s��C��
'               �@IG-XL�̃C�j�V�����C�Y�ɂ��G�N�Z���I�����̃A�v���P�[�V�������삪�s�\�ɂȂ�s����
'               �@���C�j�V�����C�Y���̃Z�b�g�A�b�v���j���[�č\�z�̍ۂɃi�r�Q�[�V�����I�u�W�F�N�g��j������
'   2009/02/20�@���d�l�ύX
'               �@�c���[�\���̒�`�����[�N�V�[�g�ֈړ��E���[�_�[�̏������݂̂��s��
'               �A�X�^�[�g�A�b�v�J�X�^�����j���[���`���[�g���j���[�o�[�ɂ��\��������
'   2009/04/07�@Ver1.00�����[�X [EeeNavigationVer1_0.xla]
'   2009/04/21�@���d�l�ύX
'               �@�A�h�C���J���K�C�h���C���ɏ]���t�@�C�����ύX [EeeNavigationAddIn.xla]
'               �A�A�h�C���J���K�C�h���C���ɏ]���v���W�F�N�g���ύX [EeeNavigationAddIn]
'               �B�o�[�W���������t�@�C���̃J�X�^���v���p�e�B�ɐݒ�E�o�[�W�����ԍ������̃v���p�e�B����擾
'               �C�v���W�F�N�g���ύX�ɂ��R�[�h�C��
'   2009/04/22�@���s��C��
'               �@�`���[�g���j���[�o�[�փR�s�[�����c�[���o�[���e���|�����R���g���[���ɂȂ�Ȃ��s�����
'               �@���R�s�[�ł͂Ȃ��e���|�����ŐV�K�쐬���A�T�u���j���[�݂̂��R�s�[������@�ɕύX
'   2009/05/11�@���s��C��
'               �@�i�r�Q�[�V�����N������JOB�A�����[�h�ŃG�N�Z���G���[����������s����
'               �@��CNavigationCommander�N���X�̎d�l�ύX�Ƃ���ɔ����Ăяo�����̕ύX
'               ���d�l�ύX
'               �@�Z�b�g�A�b�v���j���[�̃i�r�Q�[�V�����I���{�^���̃X�e�[�^�X�ݒ��ǉ�
'   2009/05/12  Ver1.01�����[�X
'             �@���d�l�ύX
'               �@�Z�b�g�A�b�v���j���[�̃c���[�r���[�o�̓{�^���̃X�e�[�^�X�ݒ��ύX
'                 ���A�N�e�B�u�ȃV�[�g���O���t�`���[�g�̏ꍇ�͖����ɐݒ�iIG-XL�G���[����̂��߁j
'               �A�c���[�r���[�o�̓{�^���̃L���v�V�����ƃA�C�R���̕ύX�i�v�����^�[�o�͂�A�z�����邽�߁j
'   2009/06/15  Ver1.01�A�h�C���W�J�ɔ����ύX
'             �@���d�l�ύX
'               �@�c���[�r���[��`�V�[�g��JOB���ɓW�J���A�V�[�g���œǍ��惏�[�N�V�[�g����肷��
'               �A�c���[�r���[�o�͗p�̃��C�^�[���e�L�X�g�p�ɕύX
'               �B�o�[�W�����C���t�H���[�V�������v���p�e�B�擾����Œ�ɕύX
'
'�쐬��:
'   0145206097
'
Option Explicit

Private Const WORKSHEET_MENU_ID = "Worksheet Menu Bar"
Private Const CHART_MENU_ID = "Chart Menu Bar"
Private Const EEENAVI_SETUP_MENU_CAPTION = "EeeNavi SetUp(&N)"
Private Const EEENAVI_SETUP_MENU_ACTION = "SetMenuButtonStatus"
Private Const EEENAVI_SETUP_MENU_STARTUP_BUTTON_CAPTION = "Start Navigation(&S)"
Private Const EEENAVI_SETUP_MENU_STARTUP_BUTTON_POPUP = "Start EeeNavi Tool Bar"
Private Const EEENAVI_SETUP_MENU_STARTUP_BUTTON_ICON = 140
Private Const EEENAVI_SETUP_MENU_STARTUP_BUTTON_ACTION = "ConstructEeeNavigation"
Private Const EEENAVI_SETUP_MENU_PRINT_BUTTON_CAPTION = "Make TreeView(&M)"
Private Const EEENAVI_SETUP_MENU_PRINT_BUTTON_POPUP = "Make Tree View On Worksheet"
Private Const EEENAVI_SETUP_MENU_PRINT_BUTTON_ICON = 512
Private Const EEENAVI_SETUP_MENU_PRINT_BUTTON_ACTION = "CreateTreeView"
Private Const EEENAVI_SETUP_MENU_CLOSE_BUTTON_CAPTION = "End Navigation(&E)"
Private Const EEENAVI_SETUP_MENU_CLOSE_BUTTON_POPUP = "Terminate EeeNavi Tool Bar"
Private Const EEENAVI_SETUP_MENU_CLOSE_BUTTON_ICON = 358
Private Const EEENAVI_SETUP_MENU_CLOSE_BUTTON_ACTION = "TerminateEeeNavigation"
Private Const EEENAVI_SETUP_MENU_INFO_BUTTON_CAPTION = "Infomation(&I)"
Private Const EEENAVI_SETUP_MENU_INFO_BUTTON_POPUP = "Show EeeNavi Infomation"
Private Const EEENAVI_SETUP_MENU_INFO_BUTTON_ICON = 984
Private Const EEENAVI_SETUP_MENU_INFO_BUTTON_ACTION = "LoadVersionInfomation"

Private Const PROPERTY_NAME = "File Version"
Private Const APPLICATION_NAME = "EeeNavigation"

Public mEeeNaviBar As CNavigationCommander
Public mDataFolder As CBookEventsAcceptor

Private Const MAX_HISTORY = 15

Public Sub CreateEeeNaviSetUpMenu()
'���e:
'   �i�r�Q�[�V�������J�n���邽�߂̃J�X�^�����j���[�o�[���G�N�Z�����j���[�o�[�֓o�^����
'   �u�b�N���J�����Ƃ��Ɏ��s����K�v������
'
'���ӎ���:
'   IG-XL�����ŃV�X�e�����������s���ƃ��[�N�u�b�N�I�u�W�F�N�g�Ȃǂ̎Q�Ƃ��؂�邽�߁A
'   �X�^�[�g�A�b�v�͎����ōs�킸���̃��j���[�o�[���烆�[�U�[���蓮�ŋN������
'   ���j���[�o�[�̓e���|�����ɐݒ肵�A�u�b�N�����Ǝ����I�ɍ폜�����悤�ɂ���
'
    '### �i�r�Q�[�V�����I�u�W�F�N�g�̔j�� ###########################
    TerminateEeeNavigation
    '### ���[�N�V�[�g���j���[�o�[�I�u�W�F�N�g�̎擾 #################
    Dim wkShtMenuBar As Office.CommandBar
    Set wkShtMenuBar = Application.CommandBars(WORKSHEET_MENU_ID)
    '### �`���[�g���j���[�o�[�I�u�W�F�N�g�̎擾 #####################
    Dim chartMenuBar As Office.CommandBar
    Set chartMenuBar = Application.CommandBars(CHART_MENU_ID)
    '### ���Ƀ��j���[�o�[�����݂���ꍇ�͍폜���� ###################
    On Error Resume Next
    wkShtMenuBar.Controls(EEENAVI_SETUP_MENU_CAPTION).Delete
    chartMenuBar.Controls(EEENAVI_SETUP_MENU_CAPTION).Delete
    On Error GoTo 0
    '### ���[�N�V�[�g���j���[�o�[��EeeNavi�c�[���o�[�̒ǉ� ##########
    Dim eeeNaviMenu As Office.CommandBarPopup
    Set eeeNaviMenu = wkShtMenuBar.Controls.Add(Type:=msoControlPopup, temporary:=True)
    '### �`���[�g���j���[�o�[��EeeNavi�c�[���o�[�̒ǉ� ##############
    Dim eeeNaviCMenu As Office.CommandBarPopup
    Set eeeNaviCMenu = chartMenuBar.Controls.Add(Type:=msoControlPopup, temporary:=True)
    '### �ǉ������c�[���o�[�̃L���v�V�����ݒ�ƃT�u���j���[�ǉ� #####
    Dim startNaviBtn As Office.CommandBarButton
    Dim printTreeBtn As Office.CommandBarButton
    Dim endNaviBtn As Office.CommandBarButton
    Dim helpMenu As Office.CommandBarButton
    With eeeNaviMenu
        .Caption = EEENAVI_SETUP_MENU_CAPTION
        .OnAction = EEENAVI_SETUP_MENU_ACTION
        With .Controls
            Set startNaviBtn = .Add(Type:=msoControlButton)
            Set printTreeBtn = .Add(Type:=msoControlButton)
            Set endNaviBtn = .Add(Type:=msoControlButton)
            Set helpMenu = .Add(Type:=msoControlButton)
        End With
    End With
    With eeeNaviCMenu
        .Caption = eeeNaviMenu.Caption
        .OnAction = eeeNaviMenu.OnAction
    End With
    '### EeeNavi�X�^�[�g�A�b�v�{�^���̐ݒ�Ǝ��s�}�N���ǉ� ##########
    With startNaviBtn
        .Caption = EEENAVI_SETUP_MENU_STARTUP_BUTTON_CAPTION
        .TooltipText = EEENAVI_SETUP_MENU_STARTUP_BUTTON_POPUP
        .FaceId = EEENAVI_SETUP_MENU_STARTUP_BUTTON_ICON
        .OnAction = EEENAVI_SETUP_MENU_STARTUP_BUTTON_ACTION
        '### �`���[�g���j���[�o�[�փR�s�[ ###########################
        .Copy eeeNaviCMenu.CommandBar
    End With
    '### EeeNavi�v�����g�{�^���̐ݒ�Ǝ��s�}�N���ǉ� ################
    With printTreeBtn
        .Caption = EEENAVI_SETUP_MENU_PRINT_BUTTON_CAPTION
        .TooltipText = EEENAVI_SETUP_MENU_PRINT_BUTTON_POPUP
        .FaceId = EEENAVI_SETUP_MENU_PRINT_BUTTON_ICON
        .OnAction = EEENAVI_SETUP_MENU_PRINT_BUTTON_ACTION
        '### �`���[�g���j���[�o�[�փR�s�[ ###########################
        .Copy eeeNaviCMenu.CommandBar
    End With
    '### EeeNavi�I���{�^���̐ݒ�Ǝ��s�}�N���ǉ� ####################
    With endNaviBtn
        .Caption = EEENAVI_SETUP_MENU_CLOSE_BUTTON_CAPTION
        .TooltipText = EEENAVI_SETUP_MENU_CLOSE_BUTTON_POPUP
        .FaceId = EEENAVI_SETUP_MENU_CLOSE_BUTTON_ICON
        .OnAction = EEENAVI_SETUP_MENU_CLOSE_BUTTON_ACTION
        .BeginGroup = True
        '### �`���[�g���j���[�o�[�փR�s�[ ###########################
        .Copy eeeNaviCMenu.CommandBar
    End With
    '### EeeNavi�w���v�{�^���̐ݒ�Ǝ��s�}�N���ǉ� ##################
    With helpMenu
        .Caption = EEENAVI_SETUP_MENU_INFO_BUTTON_CAPTION
        .TooltipText = EEENAVI_SETUP_MENU_INFO_BUTTON_POPUP
        .FaceId = EEENAVI_SETUP_MENU_INFO_BUTTON_ICON
        .OnAction = EEENAVI_SETUP_MENU_INFO_BUTTON_ACTION
        .BeginGroup = True
        '### �`���[�g���j���[�o�[�փR�s�[ ###########################
        .Copy eeeNaviCMenu.CommandBar
    End With
End Sub

Public Sub SetMenuButtonStatus()
'���e:
'   EeeNavi���j���[�N���b�N�Ŏ��s�����}�N���֐�
'   �v�����g�{�^���ƏI���{�^���̃X�e�[�^�X��ݒ肷��
'
'���ӎ���:
'   ���[�N�u�b�N�I�u�W�F�N�g��Nothing�̏ꍇ���A�N�e�B�u�ȃV�[�g��
'   ���[�N�V�[�g�ȊO�ł���ꍇ�̓v�����g�{�^���͖����ɂȂ�
'   �iIG-XL�G���[����̂��߁j
'
    '### EeeNavi���j���[�o�[�I�u�W�F�N�g�̎擾 ######################
    Dim eeeNaviMenu As Office.CommandBarPopup
    Set eeeNaviMenu = Application.CommandBars.ActionControl
    '### EeeNavi�v�����g�{�^���I�u�W�F�N�g�̎擾 ####################
    Dim printTreeBtn As Office.CommandBarButton
    Set printTreeBtn = eeeNaviMenu.Controls(EEENAVI_SETUP_MENU_PRINT_BUTTON_CAPTION)
    '### EeeNavi�v�����g�{�^���C�l�[�u���̐ݒ� ######################
    Dim shType As Excel.XlSheetType
    shType = Application.ActiveWorkbook.ActiveSheet.Type
    printTreeBtn.enabled = ((Not mDataFolder Is Nothing) And (shType = Excel.xlWorksheet))
    '### EeeNavi�I���{�^���I�u�W�F�N�g�̎擾 ########################
    Dim endNaviBtn As Office.CommandBarButton
    Set endNaviBtn = eeeNaviMenu.Controls(EEENAVI_SETUP_MENU_CLOSE_BUTTON_CAPTION)
    '### EeeNavi�I���{�^���C�l�[�u���̐ݒ� ##########################
    endNaviBtn.enabled = (Not mEeeNaviBar Is Nothing)
End Sub

Public Sub ConstructEeeNavigation()
'���e:
'   EeeNavi�X�^�[�g�A�b�v�{�^���Ŏ��s�����}�N���֐�
'   �i�r�Q�[�V�����̃I�u�W�F�N�g�𐶐������������s��
'
'���ӎ���:
'
    On Error GoTo ErrHandler
    '### �c���[�r���[��`���[�N�V�[�g�̎擾 #########################
    Dim wsSheet As Excel.Worksheet
    Set wsSheet = getWsSheet("TreeViewDefinition")
    '### �i�r�Q�[�V�����I�u�W�F�N�g�̐��� ###########################
    Dim eeeNaviCore As CDataHistoryController
    Set eeeNaviCore = New CDataHistoryController
    eeeNaviCore.Initialize MAX_HISTORY
    '### �c���[��`�f�[�^�̃��[�_�[���� #############################
    Dim treeReader As CWsTreeDataReader
    Set treeReader = New CWsTreeDataReader
    treeReader.Initialize wsSheet
    '### �G�N�X�v���[���[�I�u�W�F�N�g�̐��� #########################
    Dim eeeExplCore As CDataTreeComposer
    Set eeeExplCore = New CDataTreeComposer
    eeeExplCore.Initialize treeReader
    '### ���[�N�u�b�N�I�u�W�F�N�g�̐��� #############################
    Set mDataFolder = New CBookEventsAcceptor
    mDataFolder.Initialize Application, eeeNaviCore, eeeExplCore
    '### �i�r�Q�[�V����GUI�I�u�W�F�N�g�̐��� ########################
    Set mEeeNaviBar = New CNavigationCommander
    With mEeeNaviBar
        .Initialize Application, eeeNaviCore, eeeExplCore
        .Create
    End With
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Private Function getWsSheet(ByVal shName As String) As Excel.Worksheet
    '### ���[�N�V�[�g�I�u�W�F�N�g�擾�p�v���V�[�W�� #################
    On Error GoTo NotExist
    Set getWsSheet = ActiveWorkbook.Worksheets(shName)
    Exit Function
NotExist:
    Err.Raise 9999, "Start EeeNavigation", shName & " Worksheet Can Not Find !"
End Function

Public Sub CreateTreeView()
'���e:
'   EeeNavi�v�����g�{�^���Ŏ��s�����}�N���֐�
'   �f�[�^�c���[�̈ꗗ���e�L�X�g�Ƀv�����g�A�E�g����
'
'���ӎ���:
'
    '### �c���[�r���[���C�^�[�I�u�W�F�N�g�̐��� #####################
    Dim treeViewWriter As CTextTreeViewWriter
    Set treeViewWriter = New CTextTreeViewWriter
    On Error GoTo ErrHandler
    treeViewWriter.OpenFile ActiveWorkbook.Path
    '### ���[�N�u�b�N�I�u�W�F�N�g�ɑ΂��c���[�r���[�쐬�����s #######
    mDataFolder.WriteTreeView treeViewWriter
    treeViewWriter.CloseFile
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Exit Sub
End Sub

Public Sub TerminateEeeNavigation()
'���e:
'   EeeNavi�I���{�^���Ŏ��s�����}�N���֐�
'   �i�r�Q�[�V�����̃f�R���X�g���N�^���s��
'
'���ӎ���:
'
    '### ���[�N�u�b�N�I�u�W�F�N�g�̔j�� #############################
    Set mDataFolder = Nothing
    '### �i�r�Q�[�V����GUI�p���j���[�o�[�̑|�� ######################
    '�I�u�W�F�N�g���s���Ԃ̏ꍇ�ɃG���[�ɂȂ�Ȃ��悤�ɉ���i�s�v�H�j
    On Error Resume Next
    If Not mEeeNaviBar Is Nothing Then mEeeNaviBar.Destroy
    On Error GoTo 0
    '### �i�r�Q�[�V�����I�u�W�F�N�g�̔j�� ###########################
    Set mEeeNaviBar = Nothing
End Sub

Public Sub LoadVersionInfomation()
'���e:
'   EeeNavi�C���t�H���[�V�����{�^���Ŏ��s�����}�N���֐�
'   �i�r�Q�[�V���������t�H�[���\������
'
'���ӎ���:
'
    Dim revNum As String
'    revNum = ThisWorkbook.CustomDocumentProperties.Item(PROPERTY_NAME).Value
    revNum = "1.01"
    With EeeNaviVerFrm
        .VersionLabel = APPLICATION_NAME & " Ver." & revNum
        .Show
    End With
End Sub
