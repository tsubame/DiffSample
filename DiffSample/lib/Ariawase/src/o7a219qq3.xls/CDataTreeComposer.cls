VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataTreeComposer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �f�[�^�V�[�g�̃��X�g����f�[�^�c���[���\�����ێ�����N���X
'
'�ړI:
'   IDataExplorer�C���^�[�t�F�C�X���������G�N�X�v���[���[�I�u�W�F�N�g�̎��ԂƂ��đ��݂���
'
'   Revision History:
'   Data        Description
'   2008/11/10  �쐬
'   2008/11/11�@���S�R�[�h���s�E����OK
'   2008/12/11�@���d�l�ύX
'             �@�@�c�[���J�e�S���̏�ʊK�w��ǉ��Ǘ��֕ύX
'             �@�A�f�[�^�c���[���ނ���R�ꂽ�f�[�^�V�[�g���J�X�^���c�[���Ƃ��ĕʊǗ��֕ύX
'             �@�B�C���^�[�t�F�C�X�@�\�ǉ��ɂ��ύX
'   2008/12/15�@���@�\�ǉ�
'               �@�C���^�[�t�F�C�X�@�\�ǉ��ɂ��f�[�^�c���[�̊O���o�͋@�\������
'   2008/12/17  ���d�l�ύX
'               12/11�A�̋@�\���C���E���ނ���R���f�[�^�V�[�g�̊Ǘ��@�\�̓T�v���C���[�N���X�ֈړ�
'   2009/02/20  ���d�l�ύX
'               �@�c���[�\���̒�`�f�[�^�����[�_�[����擾����d�l�ɕύX
'               �A�f�[�^���X�g�̃\�[�g�@�\��CToolCategory�N���X�ֈڊ�
'               �@���C���h�J�[�h�̃c�[���J�e�S���݂̂��̃N���X���f�[�^�V�[�g�̒ǉ���Ƃ𒼐ڍs��
'
'�쐬��:
'   0145206097
'
Option Explicit

Implements IDataExplorer

Private mDataTree As Collection
Private mWildCard As CToolCategory

Private Sub Class_Initialize()
    Set mDataTree = New Collection
    Set mWildCard = Nothing
End Sub

Public Sub Initialize(ByVal reader As IParameterReader)
'���e:
'   �G�N�X�v���[���[�I�u�W�F�N�g�̏�����
'
'�p�����[�^:
'[reader]   In  �f�[�^�c���[�\������`���Ă���t�@�C���̃��[�_�[
'
'���ӎ���:
'
    '### ���[�_�[��Nothing�̏ꍇ�̓G���[���� ########################
    If reader Is Nothing Then GoTo ReaderErr
    '### �t�@�C���X�g���[���̏��� ###################################
    Dim fileStream As IFileStream
    Set fileStream = reader
    '### ���O�d���`�F�b�N�p�R���N�V���� #############################
    Dim errChecker As New Collection
    Dim errName As String
    '### ���[�_�[����p�����[�^�̓ǂݍ��� ###########################
    Dim tempSupp As CToolsupplier
    Do While Not fileStream.IsEOR
        Set tempSupp = New CToolsupplier
        '### ���[�_�[����g�b�v���C���[���̓ǂݍ��� #################
        tempSupp.Name = reader.ReadAsString(SUPPLIER_NAME)
        '### ���C���[�����󗓂̏ꍇ�̓G���[���� #####################
        If tempSupp.Name = NOT_DEFINE Then GoTo BlankErr
        '### ���[�_�[����T�u���C���[�p�����[�^�̓ǂݍ��� ###########
        Dim tempTool As CToolCategory
        Do While (reader.ReadAsString(TOOL_NAME) <> END_OF_DATA)
            Set tempTool = New CToolCategory
            With tempTool
                '### ���[�_�[����T�u���C���[���̓ǂݍ��� ###########
                .Name = reader.ReadAsString(TOOL_NAME)
                '### ���C���[�����󗓂̏ꍇ�̓G���[���� #############
                If .Name = NOT_DEFINE Then GoTo BlankErr
                '### ���̑��T�u���C���[�̃v���p�e�B�ݒ� #############
                .NameMap = reader.ReadAsString(NAME_LOCATION)
                .Visible = reader.ReadAsBoolean(IS_VISIBLE_TOOL)
                .Categorize = reader.ReadAsBoolean(IS_CATEGORIZE)
            End With
            '### �T�u���C���[�̓o�^ #################################
            errName = tempTool.Name
            On Error GoTo AssignErr
            tempSupp.Add tempTool
            errChecker.Add tempTool, tempTool.Name
            On Error GoTo 0
            '### ���C���h�J�[�h�̃J�e�S���͕ێ����Ă��� #############
            If tempTool.Name = WILD_CARD Then
                Set mWildCard = tempTool
            End If
            fileStream.MoveNext
        Loop
        '### �g�b�v���C���[�̓o�^ ###################################
        errName = tempSupp.Name
        On Error GoTo AssignErr
        mDataTree.Add tempSupp, tempSupp.Name
        On Error GoTo 0
        fileStream.MoveNext
    Loop
    Exit Sub
    '### �G���[���� #################################################
ReaderErr:
    Err.Raise 999, "CDataTreeComposer.Initialize", "Reader Object Is Necessary For Initialization"
AssignErr:
    Err.Raise 999, "CDataTreeComposer.Initialize", "This Is Duplicate Name [ " & errName & " ] In Data Table"
BlankErr:
    Err.Raise 999, "CDataTreeComposer.Initialize", "Wrong Blank Is Include In Data Table"
End Sub

Private Sub IDataExplorer_ConstructDataTree(ByVal dataList As Collection)
    '### �^����ꂽ�f�[�^�V�[�g��Nothing�̏ꍇ�͖������� ############
    If dataList Is Nothing Then Exit Sub
    '### �c�[���J�e�S���N���X�Ƀ��X�g�̃\�[�g��Ƃ��s�킹�� #########
    Dim toolSupp As CToolsupplier
    For Each toolSupp In mDataTree
        Dim DataTool As CToolCategory
        For Each DataTool In toolSupp.DataTools
            With DataTool
                .Clear
                .Sort dataList
            End With
        Next DataTool
    Next toolSupp
    '### ���C���h�J�[�h�̃J�e�S���Ɏc��̃f�[�^�V�[�g������ #######
    If Not mWildCard Is Nothing Then
        Dim shtIdx As Long
        Dim dataSht As CDataSheet
        For shtIdx = dataList.Count To 1 Step -1
            Set dataSht = dataList.Item(shtIdx)
            mWildCard.Add dataSht
        Next shtIdx
    End If
End Sub

Private Property Get IDataExplorer_DataTree() As Collection
    '### �c�[���T�v���C���[�I�u�W�F�N�g�̃R���N�V������Ԃ� #########
    Set IDataExplorer_DataTree = mDataTree
End Property

Private Sub IDataExplorer_WriteDataTree(ByVal writer As IParameterWriter)
    '### �e�f�[�^�K�w�Ƀf�[�^�o�͂�v������ #########################
    '### ���̃N���X�̖����͂����̎��s���߂ƃ��C�^�[�𐧌䂷�邱�� #
    Dim fileStream As IFileStream
    Set fileStream = writer
    With writer
        Dim toolSupp As CToolsupplier
        For Each toolSupp In mDataTree
            '### ���C�^�[�փg�b�v���C���[�p�����[�^�̏����o�� #######
            toolSupp.AsIParameterWritable.WriteOut writer
            fileStream.MoveNext
            Dim DataTool As CToolCategory
            For Each DataTool In toolSupp.DataTools
                '### ���C�^�[�փT�u���C���[�p�����[�^�̏����o�� #####
                DataTool.AsIParameterWritable.WriteOut writer
                fileStream.MoveNext
                Dim dataSht As CDataSheet
                For Each dataSht In DataTool.DataSheets
                    '### ���C�^�[�փf�[�^�V�[�g�p�����[�^�̏����o�� #
                    dataSht.AsIParameterWritable.WriteOut writer
                    fileStream.MoveNext
                Next dataSht
            Next DataTool
        Next toolSupp
        fileStream.IsEOR
    End With
End Sub

Public Function AsIDataExplorer() As IDataExplorer
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IDataExplorer�^�I�u�W�F�N�g
'
'���ӎ���:
'
    Set AsIDataExplorer = Me
End Function

Private Sub Class_Terminate()
    Set mDataTree = Nothing
    Set mWildCard = Nothing
End Sub