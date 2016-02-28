VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CToolsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ���̃T�v���C���[�ɑ�����c�[���J�e�S���I�u�W�F�N�g��ێ�����N���X
'
'�ړI:
'   �f�[�^�c���[�\�z�̂��߂̍ŏ�ʊK�w�Ƃ��đ��݂���
'
'   Revision History:
'   Data        Description
'   2008/12/11  �쐬
'   2008/12/15  ���@�\�ǉ�
'               �p�����[�^�̊O���o�͋@�\�̃C���^�[�t�F�C�X������
'   2008/12/17  ���d�l�ύX
'               �@�f�[�^�V�[�g�I�u�W�F�N�g�̃v���p�e�B�ǉ�
'               �@�i�T�v���C���[�I�u�W�F�N�g���J�e�S���ɕ��ނ��Ȃ��f�[�^�V�[�g�𒼐ڊǗ�����j
'   2009/02/06  ���@�\�폜
'               �@Initialize���\�b�h��p�~�EName�v���p�e�B�ݒ�ɕύX
'   2009/02/20  ���@�\�ǉ�
'               �@Add���\�b�h�̒ǉ��i�c�[���J�e�S�����R���N�V�����ɒǉ�����j
'               �AIsContain�v���p�e�B�̒ǉ��i�c�[���J�e�S�����܂�ł��邩�ǂ������擾�j
'               ���d�l�ύX
'               �@�J�X�^���c�[���Ƃ��ĕʊǗ����Ă����c�[���J�e�S���͔p�~
'                 ���C���h�J�[�h�̃J�e�S���Ƃ��đ��̃c�[���J�e�S���ƈꌳ�Ǘ��Ƃ���
'   2009/06/15  Ver1.01�A�h�C���W�J�ɔ����ύX
'             �@���d�l�ύX
'               �����C���^�[�t�F�[�X��JOB���̊����C���^�[�t�F�[�X�Ƌ��L���邽�߃��\�b�h����ύX
'
'�쐬��:
'   0145206097
'
Option Explicit

Implements IParameterWritable

Private mDataTools As Collection
Private mName As String

Private Sub Class_Initialize()
    Set mDataTools = New Collection
End Sub

Public Sub Add(ByVal DataTool As CToolCategory)
'���e:
'   �c�[���J�e�S���̃R���N�V�����ǉ�
'
'�p�����[�^:
'[dataTool]   In  �c�[���J�e�S���I�u�W�F�N�g
'
'���ӎ���:
'
    On Error GoTo AssignErr
    mDataTools.Add DataTool, DataTool.Name
    Exit Sub
AssignErr:
    Err.Raise 999, "CToolSupplier.Add", "This Is Duplicate Name [ " & DataTool.Name & " ] In Data Table"
End Sub

Public Property Get DataTools() As Collection
'���e:
'   �c�[���J�e�S���R���N�V�����̎擾
'
'�߂�l�F
'   CDataCategory�^�I�u�W�F�N�g�̃R���N�V����
'
'���ӎ���:
'
    Set DataTools = mDataTools
End Property

Public Property Let Name(ByVal suppName As String)
'���e:
'   �T�v���C���[���̐ݒ�
'
'�p�����[�^:
'[suppName]   In  �T�v���C���[��
'
'���ӎ���:
'
    mName = suppName
End Property

Public Property Get Name() As String
'���e:
'   �T�v���C���[���̎擾
'
'�߂�l�F
'   �T�v���C���[��
'
'���ӎ���:
'
    Name = mName
End Property

Public Property Get IsContain() As Boolean
'���e:
'   �c�[���J�e�S����ێ����Ă��邩�ǂ����̃v���p�e�B�擾
'
'�߂�l�F
'   True/False
'
'���ӎ���:
'
    IsContain = mDataTools.Count > 0
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    '### �p�����[�^�o�͂̎��s #######################################
    With writer
        .WriteAsString SUPPLIER_NAME, mName
        .WriteAsBoolean IS_TOOL_CONTAIN, Me.IsContain
    End With
End Sub

Private Property Get IParameterWritable_ParameterList() As Collection
    Err.Raise 9999, "CToolsupplier.IParameterWritable.ParameterList", "Does Not Support This Method !"
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
    Set mDataTools = Nothing
End Sub
