VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CToolCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ���̃J�e�S���ɑ�����f�[�^�V�[�g�I�u�W�F�N�g��ێ�����N���X
'
'�ړI:
'   �f�[�^�c���[�\�z�̂��߂̒��ԊK�w�Ƃ��đ��݂���
'
'   Revision History:
'   Data        Description
'   2008/11/10  �쐬�E�S�R�[�h���s�E����OK
'   2008/12/15  ���@�\�ǉ�
'               �p�����[�^�̊O���o�͋@�\�̃C���^�[�t�F�C�X������
'   2009/02/06  ���@�\�ǉ�
'               �@Visible�v���p�e�B�̒ǉ��i�f�[�^�c���[�ɕ\�����邩���Ȃ������w��j
'               ���@�\�폜
'               �@Initialize���\�b�h��p�~�EName�v���p�e�B�ݒ�ɕύX
'               ���S�R�[�h���s�E����OK
'   2009/02/20  ���@�\�ǉ�
'               �@Categorize�v���p�e�B�̒ǉ� (�f�[�^�c���[�ɃJ�e�S���C�Y���ĕ\�����邩�ǂ������w��)
'               �ANameMap�v���p�e�B�̒ǉ��i�f�[�^�V�[�g��̃J�e�S�����̒�`�ꏊ���w�肷��j
'               �BAdd���\�b�h�̒ǉ��i�f�[�^�V�[�g���R���N�V�����ɒǉ�����j
'               �CSort���\�b�h�̒ǉ��i�f�[�^���X�g�̃\�[�g��Ƃ��s���ECDataTreeComposer�N���X����̈ڊǋ@�\�j
'               �DIsContain�v���p�e�B�̒ǉ��i�f�[�^�V�[�g���܂�ł��邩�ǂ������擾�j
'               �EIsUnique�v���p�e�B�̒ǉ��i�f�[�^�V�[�g�����j�[�N���ǂ������擾�j
'   2009/06/15  Ver1.01�A�h�C���W�J�ɔ����ύX
'             �@���d�l�ύX
'               �����C���^�[�t�F�[�X��JOB���̊����C���^�[�t�F�[�X�Ƌ��L���邽�߃��\�b�h����ύX
'
'�쐬��:
'   0145206097
'
Option Explicit

Implements IParameterWritable

Private mDataSheets As Collection
Private mName As String
Private mNameMap As String
Private mIsVisible As Boolean
Private mCategorize As Boolean

Private Sub Class_Initialize()
    Set mDataSheets = New Collection
End Sub

Public Sub Clear()
'���e:
'   �f�[�^�V�[�g�R���N�V�����̏�����
'
'���ӎ���:
'
    Set mDataSheets = New Collection
End Sub

Public Sub Add(ByVal dataSht As CDataSheet)
'���e:
'   �f�[�^�V�[�g�̃R���N�V�����ǉ�
'
'�p�����[�^:
'[dataSht]   In  �f�[�^�V�[�g�I�u�W�F�N�g
'
'���ӎ���:
'
    On Error GoTo IsDuplicate
    mDataSheets.Add dataSht, dataSht.Name
    Exit Sub
IsDuplicate:
    Err.Raise 999, "CToolCategory.Add", "This Is Duplicate Name [ " & dataSht.Name & " ] In Tree View Definition"
End Sub

Public Sub Sort(ByVal dataList As Collection)
'���e:
'   �f�[�^���X�g���玩���̃c�[���ɑ�����f�[�^�V�[�g�𔲂����
'
'�p�����[�^:
'[dataList]   In/Out  �f�[�^�V�[�g�I�u�W�F�N�g�̃R���N�V����
'
'���ӎ���:
'   ����������f�[�^�V�[�g�̓R���N�V��������폜�����
'   �������g�����C���h�J�[�h�̏ꍇ�͂��̍�Ƃ͍s���Ȃ�
'
    '### �c�[���J�e�S�������C���h�J�[�h�̏ꍇ�̓\�[�g���s��Ȃ� #####
    If mName = WILD_CARD Then Exit Sub
    '### ��������f�[�^���X�g�̃\�[�g��� ###########################
    Dim shtIdx As Long
    Dim dataSht As CDataSheet
    For shtIdx = dataList.Count To 1 Step -1
        Set dataSht = dataList.Item(shtIdx)
        '### �c�[��������v����f�[�^�V�[�g�𔲂���� ###############
        If mName = dataSht.ParentName(mNameMap) Then
            Me.Add dataSht
            dataList.Remove (shtIdx)
        End If
    Next shtIdx
End Sub

Public Property Get DataSheets() As Collection
'���e:
'   �f�[�^�V�[�g�R���N�V�����̎擾
'
'�߂�l�F
'   CDataSheet�^�I�u�W�F�N�g�̃R���N�V����
'
'���ӎ���:
'
    Set DataSheets = mDataSheets
End Property

Public Property Let Name(ByVal toolName As String)
'���e:
'   �c�[�����̐ݒ�
'
'�p�����[�^:
'[toolName]   In  �c�[����
'
'���ӎ���:
'
    mName = toolName
End Property

Public Property Get Name() As String
'���e:
'   �c�[�����̎擾
'
'�߂�l�F
'   �c�[����
'
'���ӎ���:
'
    Name = mName
End Property

Public Property Let Visible(ByVal isVisible As Boolean)
'���e:
'   �f�[�^�c���[�\���v���p�e�B�̐ݒ�
'
'�p�����[�^:
'[isVisible]   In  True/False
'
'���ӎ���:
'
    mIsVisible = isVisible
End Property

Public Property Get Visible() As Boolean
'���e:
'   �f�[�^�c���[�\���v���p�e�B�̎擾
'
'�߂�l�F
'   True/False
'
'���ӎ���:
'
    Visible = mIsVisible
End Property

Public Property Let Categorize(ByVal isCategorize As Boolean)
'���e:
'   �f�[�^�c���[�̃J�e�S���C�Y�\���v���p�e�B�̐ݒ�
'
'�p�����[�^:
'[isCategorize]   In  True/False
'
'���ӎ���:
'
    mCategorize = isCategorize
End Property

Public Property Get Categorize() As Boolean
'���e:
'   �f�[�^�c���[�̃J�e�S���C�Y�\���v���p�e�B�̎擾
'
'�߂�l�F
'   True/False
'
'���ӎ���:
'
    Categorize = mCategorize
End Property

Public Property Let NameMap(ByVal location As String)
'���e:
'   �J�e�S�����擾����Z���ʒu�v���p�e�B�̐ݒ�
'
'�p�����[�^:
'[location]   In  �Z���ʒu��\��������iA1,B1�`���j
'
'���ӎ���:
'
    mNameMap = location
End Property

Public Property Get IsContain() As Boolean
'���e:
'   �f�[�^�V�[�g��ێ����Ă��邩�ǂ����̃v���p�e�B�擾
'
'�߂�l�F
'   True/False
'
'���ӎ���:
'
    IsContain = (mDataSheets.Count > 0)
End Property

Public Property Get IsUnique() As Boolean
'���e:
'   �f�[�^�V�[�g���c�[���J�e�S�����ň�ӂɌ��܂邩�̃v���p�e�B�擾
'
'�߂�l�F
'   True/False
'
'���ӎ���:
'
    IsUnique = (mDataSheets.Count = 1)
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    '### �p�����[�^�o�͂̎��s #######################################
    With writer
        .WriteAsString TOOL_NAME, mName
        .WriteAsString NAME_LOCATION, mNameMap
        .WriteAsBoolean IS_CATEGORIZE, mCategorize
        .WriteAsBoolean IS_VISIBLE_TOOL, mIsVisible
        .WriteAsBoolean IS_SHT_CONTAIN, Me.IsContain
        .WriteAsBoolean IS_SHT_UNIQUE, Me.IsUnique
    End With
End Sub

Private Property Get IParameterWritable_ParameterList() As Collection
    Err.Raise 9999, "CToolCategory.IParameterWritable.ParameterList", "Does Not Support This Method !"
End Property

Public Function AsIParameterWritable() As IParameterWritable
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IParameterWritable�^�I�u�W�F�N�g
'
'���ӎ���:
'
    Set AsIParameterWritable = Me
End Function

Private Sub Class_Terminate()
    Set mDataSheets = Nothing
End Sub