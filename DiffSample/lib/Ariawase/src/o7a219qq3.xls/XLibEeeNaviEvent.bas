Attribute VB_Name = "XLibEeeNaviEvent"
'�T�v:
'   EeeNavigation�����p����C�x���g�}�N���֐��Q
'
'�ړI:
'   �@�i�r�Q�[�V����GUI�ɓo�^����}�N���֐� [Commander_***]
'   �A�i�r�Q�[�V�����q�X�g�����甭������}�N���֐�
'   �@�����[�N�u�b�N�I�u�W�F�N�g����̌Ăяo���ɕύX[2009/02/20] [BookEventsAcceptor_***]
'   �B�i�r�Q�[�V�����q�X�g���ɑ΂���v���p�e�B������s�����߂̃}�N���֐�
'     ��???�폜
'   �C�V���[�g�J�b�g�L�[�ւ̓o�^�p�}�N���֐�
'   �@���ǉ� [2009/02/20] [ShortCut_***]
'
'   Revision History:
'   Data        Description
'   2008/12/11�@�쐬
'   2009/02/20  ���@�\�ǉ�
'               �@�V���[�g�J�b�g�L�[�ւ̓o�^�p�}�N���֐���ǉ��i�i�ރ{�^���A�߂�{�^���j
'               ���d�l�ύX
'               �@�i�r�Q�[�V�����q�X�g������Ăяo�����}�N�������[�N�u�b�N�I�u�W�F�N�g����̌Ăяo���ɕύX
'
'�쐬��:
'   0145206097
'
Option Explicit

Public Sub Commander_DataTreeMenu_Events()
'���e:
'   �f�[�^�c���[���j���[���N���b�N�����Ƃ��ɔ�������C�x���g�}�N��
'   �G�N�X�v���[���[�̍X�V���s���f�[�^�c���[���j���[��\������
'
'���ӎ���:
'
    On Error GoTo MenuError
    '### ���[�N�u�b�N�I�u�W�F�N�g�փf�[�^�c���[�̍X�V��v������ #####
    mDataFolder.ExplorerDataSheet
    '### �i�r�Q�[�V����GUI�I�u�W�F�N�g�փf�[�^�c���[�̕\����v������
    mEeeNaviBar.DisplayDataTreeMenu
    Exit Sub
MenuError:
    MsgBox "Error Occured !! " & CStr(999) & " - " & "EeeNavi Tool Bar" & Chr(13) & Chr(13) & "Can Not Display Data Tree Menu!"
End Sub

Public Sub Commander_HistoryMenu_Events()
'���e:
'   �q�X�g�����j���[���N���b�N�����Ƃ��ɔ�������C�x���g�}�N��
'   �q�X�g���ꗗ�����j���[�\������
'
'���ӎ���:
'
    On Error GoTo MenuError
    '### �i�r�Q�[�V����GUI�I�u�W�F�N�g�փq�X�g���ꗗ�̕\����v������
    mEeeNaviBar.DisplayHistoryMenu
    Exit Sub
MenuError:
    MsgBox "Error Occured !! " & CStr(999) & " - " & "EeeNavi Tool Bar" & Chr(13) & Chr(13) & "Can Not Display Data History Menu!"
End Sub

Public Sub Commander_DataTreeMenuButton_Events(ByVal SheetName As String)
'���e:
'   �f�[�^�c���[���j���[�̃f�[�^�V�[�g���N���b�N�����Ƃ��ɔ�������C�x���g�}�N��
'   �f�[�^�V�[�g��\�����q�X�g���ւ̒ǉ����s��
'
'�p�����[�^:
'[sheetName]   In  �\�����郏�[�N�V�[�g��
'
'���ӎ���:
'
    '### ���[�N�u�b�N�I�u�W�F�N�g�փf�[�^�V�[�g�̕\����v������ #####
    mDataFolder.ShowDataSheet SheetName
End Sub

Public Sub Commander_HistoryButton_Events(ByVal SheetName As String)
'���e:
'   �q�X�g�����j���[�̊e�{�^���N���b�N�C�x���g����Ăяo�����}�N��
'   ���[�N�u�b�N�I�u�W�F�N�g�֎w�肵���f�[�^�V�[�g�̕\����v������
'   �f�[�^�V�[�g��\�����q�X�g���ւ̒ǉ��͍s��Ȃ�
'
'�p�����[�^:
'[sheetName]   In  �\�����郏�[�N�V�[�g��
'
'���ӎ���:
'
    '### ���[�N�u�b�N�I�u�W�F�N�g�փf�[�^�V�[�g�̕\����v������ #####
    mDataFolder.ShowDataSheetWithEventCancel SheetName
End Sub

Public Sub Commander_HistoryMenuButton_Events(ByVal hIndex As Long)
'���e:
'   �q�X�g�����j���[�̃f�[�^�V�[�g���N���b�N�����Ƃ��ɔ�������C�x���g�}�N��
'   ���j���[���̃C���f�b�N�ԍ����i�r�Q�[�V����GUI�I�u�W�F�N�g�֓n���ړI�ł̂ݎg�p����
'
'�p�����[�^:
'[hIndex]   In  �N���b�N���ꂽ���j���[�̃C���f�b�N�X�ԍ�
'
'���ӎ���:
'
    '### �i�r�Q�[�V����GUI�I�u�W�F�N�g�փC���f�b�N�X�ԍ���n�� ######
    mEeeNaviBar.HistoryMenuButton_Click hIndex
End Sub

Public Sub BookEventsAcceptor_History_Events()
'���e:
'   ���[�N�u�b�N�I�u�W�F�N�g����q�X�g�����j���[��
'   �v���p�e�B����̂��߂ɌĂяo�����C�x���g�}�N��
'
'���ӎ���:
'
    '### �q�X�g�����j���[�̃X�e�[�^�X�𓮓I�ɐݒ肷�� ###############
    mEeeNaviBar.SetHistoryButtonEnable
End Sub

Public Sub ShortCut_HistoryForeButton_Events()
'���e:
'   �V���[�g�J�b�g�ɓo�^����i�r�Q�[�V�����u�i�ށv�{�^������̃}�N��
'
'���ӎ���:
'
    '### �u�i�ށv�{�^���̃N���b�N��������s���� #####################
    On Error Resume Next
    mEeeNaviBar.HistoryForeButton_Click
    On Error GoTo 0
End Sub

Public Sub ShortCut_HistoryBackButton_Events()
'���e:
'   �V���[�g�J�b�g�ɓo�^����i�r�Q�[�V�����u�߂�v�{�^������̃}�N��
'
'���ӎ���:
'
    '### �u�߂�v�{�^���̃N���b�N��������s���� #####################
    On Error Resume Next
    mEeeNaviBar.HistoryBackButton_Click
    On Error GoTo 0
End Sub

