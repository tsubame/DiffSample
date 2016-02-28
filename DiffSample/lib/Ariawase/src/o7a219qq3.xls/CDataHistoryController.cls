VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataHistoryController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �f�[�^�V�[�g�{�������Ǘ��p�N���X
'
'�ړI:
'   IDataNavigator�C���^�[�t�F�C�X���������i�r�Q�[�V�����̎��̂Ƃ��đ��݂���
'
'   Revision History:
'   Data        Description
'   2008/11/10  �쐬
'   2008/11/11�@���S�R�[�h���s�E����OK
'   2008/12/11�@���d�l�ύX
'             �@�@�O��q�X�g���Ǘ��̔p�~�E�C���f�b�N�X�Ǘ��ւ̕ύX
'             �@�A�q�X�g���L�����E���Z�b�g���̃C�x���g�����̒ǉ�
'             �@�B�C���^�[�t�F�C�X�@�\�ǉ��ɂ��ύX
'   2009/02/20�@���d�l�ύX
'               �@�q�X�g���L�����E���Z�b�g���̃C�x���g�}�N���Ăяo���͍폜�iCBookEventsAcceptor�ֈړ��j
'               �A�����C���f�b�N�X�O��y�юw��̃f�[�^�V�[�g�I�u�W�F�N�g��Ԃ����\�b�h�̃��W�b�N��ύX
'               �@�i�f�[�^�V�[�g���폜����Ă���ꍇ���l�����Ď����I�ɗL���ȃf�[�^�V�[�g�܂őO��ɃX�L�b�v����j
'               �B�����C���f�b�N�X�̗L�����胁�\�b�h�̃��W�b�N��ύX
'               �@�i�C���f�b�N�X�ʒu�����łȂ��A�C���f�b�N�X�O��ɗL���ȃf�[�^�V�[�g���܂܂�Ă��邩�ǂ������l������j
'               �C�C���^�[�t�F�C�X�@�\�ǉ��ɂ��ύX
'               �D�f�[�^�V�[�g�̃A�N�e�B�u�X�e�[�^�X�ݒ��p�~
'�쐬��:
'   0145206097
'
Option Explicit

Implements IDataNavigator

Private mHistory As Collection
Private mHistIndex As Long
Private mMaxHistory As Long

Private Sub Class_Initialize()
    Set mHistory = New Collection
    mHistIndex = 0
End Sub

Public Sub Initialize(ByVal maxHistory As Long)
'���e:
'   �i�r�Q�[�V�����I�u�W�F�N�g�̏�����
'
'�p�����[�^:
'[maxHistory]   In  �ő�q�X�g����
'
'���ӎ���:
'
    If maxHistory < 1 Then GoTo WrongData
    mMaxHistory = maxHistory
    Exit Sub
    '### �ő�q�X�g�����������ȏꍇ�̓G���[���� #####################
WrongData:
    Err.Raise 999, "CDataHistoryController.Initialize", "Please Input Max History Number That Are Larger Than 1 !"
End Sub

Private Sub IDataNavigator_Add(ByVal dataSht As CDataSheet)
    '### �f�[�^�V�[�g�I�u�W�F�N�g�̃q�X�g���ւ̒ǉ� #################
    FILO mHistory, dataSht
End Sub

Private Property Get IDataNavigator_History() As Collection
    '### �q�X�g�����Ȃ��ꍇ��Nothing��Ԃ� ##########################
    If mHistory.Count = 0 Then
        Set IDataNavigator_History = Nothing
    '### �q�X�g���R���N�V������Ԃ� #################################
    Else
        Set IDataNavigator_History = mHistory
    End If
End Property

Private Function IDataNavigator_DataSheet(ByVal hIndex As Long) As CDataSheet
    '### �����ݒ� ###################################################
    Set IDataNavigator_DataSheet = Nothing
    '### �C���f�b�N�X���L���ȏꍇ ###################################
    If hIndex >= 1 And hIndex <= mHistory.Count Then
        '### �w�肵���f�[�^�V�[�g���폜����Ă��Ȃ���� #############
        '### �f�[�^�V�[�g�I�u�W�F�N�g��Ԃ�             #############
        If Not mHistory(hIndex).IsDeleted Then
            mHistIndex = hIndex
            Set IDataNavigator_DataSheet = mHistory(mHistIndex)
        End If
    End If
End Function

Private Property Get IDataNavigator_ForwardDataSheet() As CDataSheet
    '### �����ݒ� ###################################################
    Set IDataNavigator_ForwardDataSheet = Nothing
    '### �C���f�b�N�X���C���N�������g���f�[�^�I�u�W�F�N�g��Ԃ� #####
    '### �w�肵���f�[�^�V�[�g���폜����Ă���ꍇ�̓X�L�b�v��   #####
    '### �u�i�ށv�X�e�[�^�X�������ɂȂ�܂ŌJ��Ԃ�             #####
    Do While Me.AsIDataNavigator.IsEnableForward
        mHistIndex = mHistIndex + 1
        If Not mHistory(mHistIndex).IsDeleted Then
            Set IDataNavigator_ForwardDataSheet = mHistory(mHistIndex)
            Exit Do
        End If
    Loop
End Property

Private Property Get IDataNavigator_BackwardDataSheet() As CDataSheet
    '### �����ݒ� ###################################################
    Set IDataNavigator_BackwardDataSheet = Nothing
    '### �C���f�b�N�X���f�N�������g���f�[�^�I�u�W�F�N�g��Ԃ� #######
    '### �w�肵���f�[�^�V�[�g���폜����Ă���ꍇ�̓X�L�b�v�� #######
    '### �u�߂�v�X�e�[�^�X�������ɂȂ�܂ŌJ��Ԃ�           #######
    Do While Me.AsIDataNavigator.IsEnableBackward
        mHistIndex = mHistIndex - 1
        If Not mHistory(mHistIndex).IsDeleted Then
            Set IDataNavigator_BackwardDataSheet = mHistory(mHistIndex)
            Exit Do
        End If
    Loop
End Property

Private Property Get IDataNavigator_IsEnableForward() As Boolean
    '### �����ݒ� ###################################################
    IDataNavigator_IsEnableForward = False
    '### �C���f�b�N�X���q�X�g�����𒴂��Ă��Ȃ��ꍇ ##################
    If mHistIndex < mHistory.Count Then
        '### ���݂̃C���f�b�N�X�ȍ~�̃f�[�^���S�� ####################
        '### �폜����Ă�����False                ####################
        Dim dataIdx As Long
        For dataIdx = mHistIndex + 1 To mHistory.Count
            If Not mHistory(dataIdx).IsDeleted Then
                IDataNavigator_IsEnableForward = True
                Exit For
            End If
        Next dataIdx
    End If
End Property

Private Property Get IDataNavigator_IsEnableBackward() As Boolean
    '### �����ݒ� ###################################################
    IDataNavigator_IsEnableBackward = False
    '### �C���f�b�N�X��1�ȉ��łȂ��ꍇ ###############################
    If mHistIndex - 1 > 0 Then
        '### ���݂̃C���f�b�N�X�ȑO�̃f�[�^���S�� ####################
        '### �폜����Ă�����False                ####################
        Dim dataIdx As Long
        For dataIdx = mHistIndex - 1 To 1 Step -1
            If Not mHistory(dataIdx).IsDeleted Then
                IDataNavigator_IsEnableBackward = True
                Exit For
            End If
        Next dataIdx
    End If
End Property

Private Property Get IDataNavigator_HistoryIndex() As Long
    '### �J�����g�C���f�b�N�X��Ԃ� #################################
    IDataNavigator_HistoryIndex = mHistIndex
End Property

Private Sub IDataNavigator_Clear()
    '### �q�X�g���ƃC���f�b�N�X�̃��Z�b�g ###########################
    Set mHistory = New Collection
    mHistIndex = 0
End Sub

Public Function AsIDataNavigator() As IDataNavigator
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IDataNavigator�^�I�u�W�F�N�g
'
'���ӎ���:
'
    Set AsIDataNavigator = Me
End Function

Private Function IsMaxCount(ByVal History As Collection) As Boolean
'####################################################################
'   �o�^�����ő�q�X�g�����𒴂��Ă��邩�ǂ����̔���
'####################################################################
    IsMaxCount = (History.Count >= mMaxHistory)
End Function

Private Sub FILO(ByVal History As Collection, ByVal dataSht As CDataSheet)
'####################################################################
'   �f�[�^�V�[�g�I�u�W�F�N�g�̃q�X�g���ւ̒ǉ��i������o���j
'####################################################################
    With History
        '### �q�X�g�����ő�ł�������Â��f�[�^����폜 #############
        If IsMaxCount(History) Then
            .Remove 1
            mHistIndex = mHistIndex - 1
        End If
        '### �ǉ�����f�[�^����O�Ɠ����ł���Ζ������� ###########
        If .Count > 0 Then
            If .Item(.Count).Name = dataSht.Name Then
                mHistIndex = .Count
                Exit Sub
            End If
        End If
        '### �f�[�^�̒ǉ��ƃC���f�b�N�X�̍X�V #######################
        mHistIndex = .Count + 1
        .Add dataSht
    End With
End Sub

Private Sub Class_Terminate()
    Set mHistory = Nothing
End Sub
