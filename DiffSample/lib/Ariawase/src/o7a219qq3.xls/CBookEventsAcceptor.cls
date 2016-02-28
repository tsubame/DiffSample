VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBookEventsAcceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ���[�N�u�b�N�̃��b�p�[�N���X
'   ���[�N�u�b�N�̋@�\�g���ƃ��[�N�V�[�g����̃C�x���g���󂯎������������s��
'
'�ړI:
'   �@���[�N�V�[�g�V�K�쐬���̃f�[�^�c���[�X�V
'   �A���[�N�V�[�g���O�ύX���̃f�[�^�c���[�X�V
'   �B���[�N�V�[�g�폜���̃f�[�^�c���[�X�V
'   �C���[�N�V�[�g�{�����̃i�r�Q�[�V�����q�X�g���ǉ�
'
'   ���d�l�ύX[2008/12/15�ȍ~]
'
'   [�O�����J�p]
'   �@�f�[�^�c���[�X�V�@�\
'   �A�f�[�^�V�[�g�\���@�\
'   �B�f�[�^�c���[�̃c���[�r���[���[�ւ̏o�͋@�\
'   [���[�N�V�[�g����̃C�x���g�擾]
'   �@���[�N�V�[�g�폜���̃i�r�Q�[�V�����q�X�g���̃��Z�b�g
'   �A���[�N�V�[�g�{�����̃i�r�Q�[�V�����q�X�g���ǉ�
'
'   Revision History:
'   Data        Description
'   2008/11/18  �쐬�E�S�R�[�h���s�E����OK
'   2008/12/10�@���@�\�ǉ�
'             �@�@�f�[�^�c���[�X�V�E�f�[�^�V�[�g�{���̒ǉ��@�\���O���֌��J
'             �@���C�x���g�擾�ǉ�
'             �@�@IG-XL�̃G�N�Z���C�x���g�����Ή��̂��߃V�[�g�폜�̃��j���[�C�x���g���擾
'   2008/12/15�@���@�\�ǉ�
'             �@�@�c���[�r���[�o�͂̂��߂̊O�����J�@�\��ǉ�
'   2009/02/20�@���d�l�ύX
'               �@�G�N�Z���I�u�W�F�N�g�̎擾
'               �A�V�[�g�؂�ւ����̃}�N���Ăяo����ǉ�
'               �B�f�[�^�V�[�g�̃A�N�e�B�u�X�e�[�^�X�ݒ��p�~
'               �C�i�r�Q�[�V�����ւ̃f�[�^�ǉ��Ƀ��[�N�V�[�g�I�u�W�F�N�g�ȊO���܂߂�悤�ύX
'
'�쐬��:
'   0145206097
'
Option Explicit

Private oXL As Excel.Application

Private WithEvents mWorkbook As Excel.Workbook
Attribute mWorkbook.VB_VarHelpID = -1
Private WithEvents mDeleteMenu As Office.CommandBarButton
Attribute mDeleteMenu.VB_VarHelpID = -1
Private mNavigator As IDataNavigator
Private mExplorer As IDataExplorer
Private mPreSheet As CDataSheet

Private Const DELETE_MENU_CAPTION = "Edit"
Private Const DELETE_BUTTON_ID = 847

Private Sub Class_Initialize()
    Set mNavigator = New IDataNavigator
    Set mExplorer = New IDataExplorer
    Set mPreSheet = New CDataSheet
End Sub

Public Sub Initialize(ByVal excelApp As Excel.Application, ByVal navigator As IDataNavigator, ByVal explorer As IDataExplorer)
'���e:
'   ���[�N�u�b�NGUI�I�u�W�F�N�g�̏�����
'
'�p�����[�^:
'[excelApp]   In  �G�N�Z���I�u�W�F�N�g
'[navigator]  In  �i�r�Q�[�V�����I�u�W�F�N�g
'[explorer]   In  �G�N�X�v���[���[�I�u�W�F�N�g
'
'���ӎ���:
'
    Set oXL = excelApp
    Set mWorkbook = oXL.ActiveWorkbook
    Set mNavigator = navigator
    Set mExplorer = explorer
    '### �A�N�e�B�u�V�[�g���i�r�Q�[�V�����q�X�g���ɒǉ��@############
    mNavigator.Add getDataSheet(mWorkbook.ActiveSheet)
    On Error GoTo ErrGetMenu
    '### �E�B���h�E�Y���j���[�o�[�R���g���[���擾 ###################
    Set mDeleteMenu = oXL.CommandBars(DELETE_MENU_CAPTION).FindControl(id:=DELETE_BUTTON_ID)
    Exit Sub
ErrGetMenu:
    Err.Raise 999, "CBookEventsAcceptor.Initialize", "Can Not Get Edit Menu Button Object !"
End Sub

'### �i�r�Q�[�V����GUI�Ɍ��J���郏�[�N�u�b�N�̊g���@�\ ##########################
Public Sub ExplorerDataSheet()
'���e:
'   ���[�N�u�b�N���̃��[�N�V�[�g�������W���ăG�N�X�v���[���[�̃f�[�^�c���[���X�V
'
'�p�����[�^:
'
'���ӎ���:
'
    '### �G�N�X�v���[���[�Ƀf�[�^��n���ăf�[�^�c���[���\�z���� #####
    mExplorer.ConstructDataTree collectDataSheet
End Sub

Public Sub ShowDataSheet(ByVal SheetName As String)
'���e:
'   ���[�N�u�b�N���̃��[�N�V�[�g�̕\�����\�b�h
'
'�p�����[�^:
'[sheetName]   In  �A�N�e�B�u�ɂ��郏�[�N�V�[�g��
'
'���ӎ���:
'   ���[�N�V�[�g�����ɃA�N�e�B�u�ȏꍇ�͂��̑���͖��������
'   ���̃��\�b�h���s��̓V�[�g����ɂ��}�N���C�x���g����������
'
    '### ���[�N�V�[�g�����ɃA�N�e�B�u�ȏꍇ�͖��� ###################
    If SheetName = mWorkbook.ActiveSheet.Name Then Exit Sub
    '### �V�[�g���A�N�e�B�u�ɐݒ� ###################################
    On Error Resume Next
    mWorkbook.Sheets(SheetName).Activate
    On Error GoTo 0
End Sub

Public Sub ShowDataSheetWithEventCancel(ByVal SheetName As String)
'���e:
'   �u�b�N�C�x���g�𔺂�Ȃ����[�N�u�b�N���̃��[�N�V�[�g�̕\�����\�b�h
'
'�p�����[�^:
'[sheetName]   In  �A�N�e�B�u�ɂ��郏�[�N�V�[�g��
'
'���ӎ���:
'   ���̃��\�b�h�ł̓��[�N�V�[�g����ɂ��C�x���g�̓L�����Z������
'   �܂����[�N�V�[�g�����ɃA�N�e�B�u�ȏꍇ�͂��̑���͖��������
'   �C�x���g�ɂ�郏�[�N�u�b�N�̓��������𔺂�Ȃ��P���ȋ@�\������
'
    '### ���[�N�V�[�g�����ɃA�N�e�B�u�ȏꍇ�͖��� ###################
    If SheetName = mWorkbook.ActiveSheet.Name Then Exit Sub
    '### �}�N���C�x���g�𖳌��ɐݒ� #################################
    oXL.EnableEvents = False
    '### �V�[�g���A�N�e�B�u�ɐݒ� ###################################
    On Error Resume Next
    mWorkbook.Sheets(SheetName).Activate
    On Error GoTo 0
    '### �}�N���C�x���g��L���ɐݒ� #################################
    oXL.EnableEvents = True
End Sub

Public Sub WriteTreeView(ByVal writer As IParameterWriter)
'���e:
'   �c���[�f�[�^���X�V���ăf�[�^�o�͂̎��s��v������
'
'�p�����[�^:
'[writer]   In  �c���[�f�[�^�̏o�͐�ƂȂ郉�C�^�[
'
'���ӎ���:
'
    '### �G�N�X�v���[���[�̍X�V #####################################
    Me.ExplorerDataSheet
    oXL.ScreenUpdating = False
    '### �G�N�X�v���[���[�ɑ΂��f�[�^�o�͂����s������ ###############
    On Error GoTo ErrHandler
    mExplorer.WriteDataTree writer
    oXL.ScreenUpdating = True
    Exit Sub
ErrHandler:
    oXL.ScreenUpdating = True
    Err.Raise 999, "CBookEventsAcceptor.WriteTreeView", "Can Not Print Tree View"
End Sub

'################################################################################

'#### ���[�N�u�b�NGUI�̑���ɂ��N������C�x���g�֐��Q #########################
Private Sub mDeleteMenu_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'####################################################################
'���[�N�V�[�g���j���[��̃��[�N�V�[�g�폜�{�^���������ꂽ���ɔ�������C�x���g�֐�
'�i�V�[�g�^�u��̉E�N���b�N���j���[���܂�)
'
'IG-XL���Ń��[�N�V�[�g�폜���s���ƁA�폜���j���[�̃R���g���[���{�^����
'IG-XL���}�N����o�^���Ă��邽�߁A�G�N�Z���C�x���g������肳��Ă��܂�
'���̂��ߍ폜�{�^���N���b�N�̃C�x���g���擾����IG-XL�C�x���g����U�L�����Z�����A
'�蓮�œo�^���Ă���IG-XL�}�N�������s����
'####################################################################

    '### �A�N�e�B�u�ȃV�[�g����f�[�^�V�[�g�I�u�W�F�N�g�𐶐� #######
    Dim workSh As CDataSheet
    Set workSh = New CDataSheet
    workSh.Initialize mWorkbook.ActiveSheet

    '### "Test Procedures"�V�[�g���폜�����IG-XL�}�N������ #########
    '### �Ȃ����G���[�ɂȂ� ���̃G���[�Ɋ������܂�Ȃ�����  #########
    '### EeeNavi�̏������X���[���� �@�@�@ �@�@�@�@�@�@�@�@�@#########
    If workSh.ParentName = "Test Procedures" Then GoTo ErrIGXLMacro

    '### IGXL�̃}�N���C�x���g�����s #################################
    CancelDefault = True
    On Error GoTo ErrIGXLMacro
    oXL.Run (mDeleteMenu.OnAction)

    '### �V�[�g���폜����Ă�����i�r�Q�[�V���������Z�b�g ###########
    If workSh.IsDeleted Then
        mNavigator.Clear
        mNavigator.Add getDataSheet(mWorkbook.ActiveSheet)
        BookEventsAcceptor_History_Events
    End If
    Exit Sub
    '### �}�N�����o�^����Ă��Ȃ��ꍇ�͒ʏ�̍폜����� #############
ErrIGXLMacro:
    CancelDefault = False
End Sub

Private Sub mWorkbook_SheetDeactivate(ByVal sh As Object)
'####################################################################
'���[�N�V�[�g���f�B�A�N�e�B�u�ɂȂ����Ƃ��ɔ�������C�x���g�֐�
'####################################################################
    '### �e���|�����̃f�[�^�V�[�g�I�u�W�F�N�g���� ###################
    Set mPreSheet = getDataSheet(sh)
End Sub

Private Sub mWorkbook_SheetActivate(ByVal sh As Object)
'####################################################################
'���[�N�V�[�g���A�N�e�B�u�ɂȂ����Ƃ��ɔ�������C�x���g
'####################################################################
    '### �V�[�g���폜����Ă�����i�r�Q�[�V���������Z�b�g ###########
    If mPreSheet.IsDeleted Then mNavigator.Clear
    '### �A�N�e�B�u���[�N�V�[�g���i�r�Q�[�V�����q�X�g���ɒǉ��@######
    mNavigator.Add getDataSheet(sh)
    BookEventsAcceptor_History_Events
End Sub
'################################################################################

Private Function collectDataSheet() As Collection
'####################################################################
'���[�N�u�b�N���̃��[�N�V�[�g�����W
'####################################################################
    Set collectDataSheet = New Collection
    Dim workSh As Object
    Dim currData As CDataSheet
    '### ���[�N�u�b�N���̃��[�N�V�[�g������     #####################
    '### �����Ώۂ̓��[�N�V�[�g�I�u�W�F�N�g�̂� #####################
    For Each workSh In mWorkbook.Worksheets
        '### �f�[�^�V�[�g�I�u�W�F�N�g���� ###########################
        Set currData = New CDataSheet
        currData.Initialize workSh
        '### �f�[�^���X�g�ւ̒ǉ� ###################################
        collectDataSheet.Add currData
    Next workSh
    '### ���[�N�V�[�g���Ȃ��ꍇ��Nothing��Ԃ� ######################
    If collectDataSheet.Count = 0 Then
        Set collectDataSheet = Nothing
    End If
End Function

Private Function getDataSheet(ByVal workSh As Object) As CDataSheet
'####################################################################
'�f�[�^�V�[�g�I�u�W�F�N�g�̃R���X�g���N�^
'####################################################################
    Set getDataSheet = New CDataSheet
    With workSh
        getDataSheet.Initialize workSh
    End With
End Function

Private Sub Class_Terminate()
    Set mNavigator = Nothing
    Set mExplorer = Nothing
    Set mWorkbook = Nothing
    On Error Resume Next
    Set mDeleteMenu = Nothing
    On Error GoTo 0
End Sub