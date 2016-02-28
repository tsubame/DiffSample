VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �ΏۂƂȂ�f�[�^�V�[�g�̃I�u�W�F�N�g��ێ�����N���X
'
'�ړI:
'   ���[�N�V�[�g�i�r�Q�[�V�����ƃ��[�U�[�C���^�[�t�F�[�X��
'   ���̃I�u�W�F�N�g�𗘗p���ď�����������
'
'   Revision History:
'   Data        Description
'   2008/11/10  �쐬�E�S�R�[�h���s�E����OK
'   2008/12/11�@���d�l�ύX
'             �@�@�V�[�g���v���p�e�B�����[�N�V�[�g�I�u�W�F�N�g�v���p�e�B�֕ύX�i�V�[�g���ύX�ɑΉ����邽�߁j
'   2008/12/15  ���@�\�ǉ�
'               �@�p�����[�^�̊O���o�͋@�\�̃C���^�[�t�F�C�X������
'   2009/02/06  ���d�l�ύX
'               �@�c�[�����̃v���p�e�B��Read Only�ɕύX
'               �A���[�N�V�[�g���j������Ă����ꍇ�̃V�[�g���ƃc�[�����ɍ폜����Ă��鎖���������x����Ԃ��悤�ύX
'               ���S�R�[�h���s�E����OK
'   2009/02/20�@���d�l�ύX
'               �@�c�[�����擾�Z���ʒu���I�v�V�����w��o����悤�C��
'               �A�A�N�e�B�u�X�e�[�^�X�͂��̃N���X���ێ����Ă��郏�[�N�V�[�g�I�u�W�F�N�g����擾����悤�ύX
'               �B���̃N���X���ێ�����f�[�^�̎��̂����[�N�V�[�g����V�[�g�I�u�W�F�N�g�ɕύX
'   2009/06/15  Ver1.01�A�h�C���W�J�ɔ����ύX
'             �@���d�l�ύX
'               �����C���^�[�t�F�[�X��JOB���̊����C���^�[�t�F�[�X�Ƌ��L���邽�߃��\�b�h����ύX
'
'�쐬��:
'   0145206097
'
Option Explicit

Implements IParameterWritable

Private mDataSheet As Object
Private mIsActive As Boolean

Public Sub Initialize(ByVal excelApp As Object)
'���e:
'   �f�[�^�V�[�g�I�u�W�F�N�g�̏�����
'
'�p�����[�^:
'[excelApp]   In  �f�[�^�V�[�g�̎��̂ƂȂ�V�[�g�I�u�W�F�N�g
'
'���ӎ���:
'
    Set mDataSheet = excelApp
End Sub

Public Function IsDeleted() As Boolean
'���e:
'   �V�[�g�I�u�W�F�N�g�̗L���𔻕�
'
'�߂�l�F
'   �V�[�g�I�u�W�F�N�g���j������Ă�����True��Ԃ�
'
'���ӎ���:
'   ���[�U�[�̃V�[�g����ɂ�肱�̃N���X���ێ����Ă���V�[�g�I�u�W�F�N�g��
'   �j�����ꂽ�ꍇ�A���̃v���p�e�B�𗘗p���ăV�[�g�̑��݂̗L�����`�F�b�N�o����
'
    On Error GoTo NotFound
    Dim MyName As String
    MyName = mDataSheet.Name
    IsDeleted = False
    Exit Function
NotFound:
    IsDeleted = True
End Function

Public Property Get Name() As String
'���e:
'   �V�[�g���̎擾
'
'�߂�l�F
'   �V�[�g��
'
'���ӎ���:
'   �V�[�g���j������Ă���ꍇ��"<Deleted>"��Ԃ�
'
    If Not Me.IsDeleted Then
        Name = mDataSheet.Name
    Else
        Name = SHEET_MISSING
    End If
End Property

Public Function ParentName(Optional ByVal location As String = DEF_NAME_MAP) As String
'���e:
'   ���[�N�V�[�g��������c�[�����̎擾
'
'�p�����[�^:
'[location]   In  �c�[��������`����Ă��郏�[�N�V�[�g�̃Z���ʒu
'
'�߂�l�F
'   �c�[����
'
'���ӎ���:
'   ���̃v���p�e�B���g����̂͑Ώۂ����[�N�V�[�g�I�u�W�F�N�g�ł���Ƃ�����
'   ����ȊO�̃V�[�g�I�u�W�F�N�g�̏ꍇ��"NotWorkSheet"��Ԃ�
'   ���[�N�V�[�g�̃c�[�����Z���ʒu�̓f�t�H���g"B1"
'   ���[�N�V�[�g���j������Ă���ꍇ��"<Deleted>"��Ԃ�
'
    If Not Me.IsDeleted Then
        If mDataSheet.Type = Excel.xlWorksheet Then
            ParentName = mDataSheet.Range(location).Value
        Else
            ParentName = NOT_WORKSHEET
        End If
    Else
        ParentName = SHEET_MISSING
    End If
End Function

Public Property Get IsActive() As Boolean
'���e:
'   �V�[�g�̃A�N�e�B�u�X�e�[�^�X�̎擾
'
'�߂�l�F
'   True/False
'
'���ӎ���:
'   �V�[�g���j������Ă���ꍇ��False��Ԃ�
'
    If Not Me.IsDeleted Then
        With mDataSheet
            IsActive = (.Name = .parent.ActiveSheet.Name)
        End With
    Else
        IsActive = False
    End If
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    '### �p�����[�^�o�͂̎��s #######################################
    With writer
        .WriteAsString sheet_name, Me.Name
        .WriteAsBoolean IS_SHT_ACTIVE, Me.IsActive
        .WriteAsBoolean IS_SHT_DELETED, Me.IsDeleted
    End With
End Sub

Private Property Get IParameterWritable_ParameterList() As Collection
    Err.Raise 9999, "CDataSheet.IParameterWritable.ParameterList", "Does Not Support This Method !"
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
    Set mDataSheet = Nothing
End Sub