VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWsTreeDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �c���[�r���[�̊K�w�\����`���[�N�V�[�g�̃��[�_�[�N���X
'
'�ړI:
'   �c���[�r���[�̊K�w�\����`�����[�N�V�[�g����ǂݍ���
'
'   Revision History:
'   Data        Description
'   2009/02/03  �쐬
'   2009/06/15  Ver1.01�A�h�C���W�J�ɔ����ύX
'             �@���d�l�ύX
'               �����C���^�[�t�F�[�X��JOB���̊����C���^�[�t�F�[�X�Ƌ��L���邽�߃��\�b�h����ύX
'
'�쐬��:
'   0145206097
'
Option Explicit

Implements IParameterReader
Implements IFileStream

Private mTreeDataSh As Excel.Worksheet
Private mDataCell As Excel.Range
Private Const START_CELL = "A4"
Private Const SUBLAYER_COL = 1
Private Const LOCATION_COL = 2
Private Const CATEGORIZE_COL = 3
Private Const VISIBLE_COL = 4

Public Sub Initialize(ByVal wksht As Excel.Worksheet)
'���e:
'   ���[�_�[�I�u�W�F�N�g�̏�����
'
'�p�����[�^:
'[wkSht]  In  �ǂݍ��ݐ�̃��[�N�V�[�g�I�u�W�F�N�g
'
'���ӎ���:
'
    '### ���[�N�V�[�g�I�u�W�F�N�g�̎擾 #############################
    Set mTreeDataSh = wksht
    '### �t�@�C���X�g���[���̃��Z�b�g ###############################
    Me.AsIFileStream.Rewind
    '### �擪�̃f�[�^����`����Ă��Ȃ��ꍇ�̓G���[���� #############
    With mTreeDataSh
        If IsEmpty(mDataCell) Then
            Err.Raise 999, "CWsTreeDataReader.Initialize", "Start Of Data Is Not Defined in  [" & .Name & "] !"
        End If
    '### EOF����`����Ă��Ȃ��ꍇ�̓G���[���� ######################
        Dim endCell As Excel.Range
        Set endCell = .Columns(mDataCell.Column).Find(END_OF_FILE)
        If endCell Is Nothing Then
            Err.Raise 999, "CWsTreeDataReader.Initialize", "End Of File Is Not Defined in  [" & .Name & "] !"
        End If
    End With
End Sub

Private Function IParameterReader_ReadAsBoolean(ByVal dataName As String) As Boolean
    '### �Z���f�[�^�̓Ǎ��� #########################################
    Select Case dataName:
        Case IS_CATEGORIZE:
            IParameterReader_ReadAsBoolean = mDataCell.offset(0, CATEGORIZE_COL).Value
        Case IS_VISIBLE_TOOL:
            IParameterReader_ReadAsBoolean = mDataCell.offset(0, VISIBLE_COL).Value
        Case Else
            Err.Raise 999, "CWsTreeDataReader.AsIParameterReader.ReadAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Function IParameterReader_ReadAsDouble(ByVal dataName As String) As Double
    '### ���̃��[�_�[�͂��̃C���^�[�t�F�C�X�@�\��ۏႵ�Ȃ� #########
    Err.Raise 999, "CWsTreeDataReader.AsIParameterReader.ReadAsDouble", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsInteger(ByVal dataName As String) As Integer
    '### ���̃��[�_�[�͂��̃C���^�[�t�F�C�X�@�\��ۏႵ�Ȃ� #########
    Err.Raise 999, "CWsTreeDataReader.AsIParameterReader.ReadAsInteger", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsLong(ByVal dataName As String) As Long
    '### ���̃��[�_�[�͂��̃C���^�[�t�F�C�X�@�\��ۏႵ�Ȃ� #########
    Err.Raise 999, "CWsTreeDataReader.AsIParameterReader.ReadAsLong", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsString(ByVal dataName As String) As String
    '### �Z���f�[�^�̓Ǎ��� #########################################
    Select Case dataName:
        Case SUPPLIER_NAME:
            If IsEmpty(mDataCell) Then
                IParameterReader_ReadAsString = NOT_DEFINE
            Else
                IParameterReader_ReadAsString = mDataCell.Value
            End If
        Case TOOL_NAME:
            If IsEmpty(mDataCell.offset(0, SUBLAYER_COL)) Then
                IParameterReader_ReadAsString = NOT_DEFINE
            Else
                IParameterReader_ReadAsString = mDataCell.offset(0, SUBLAYER_COL).Value
            End If
        Case NAME_LOCATION:
            '### �󔒂̏ꍇ�̓f�t�H���g�l��Ԃ� #####################
            If IsEmpty(mDataCell.offset(0, LOCATION_COL)) Then
                IParameterReader_ReadAsString = DEF_NAME_MAP
            Else
                IParameterReader_ReadAsString = mDataCell.offset(0, LOCATION_COL).Value
            End If
        Case Else
            Err.Raise 999, "CWsTreeDataReader.AsIParameterReader.ReadAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Sub IFileStream_MoveNext()
    '### �f�[�^�Z���C���f�b�N�X�̃C���N�������g #####################
    Set mDataCell = mDataCell.offset(1, 0)
End Sub

Private Function IFileStream_IsEOR() As Boolean
    '### �f�[�^�Z���̏I������ #######################################
    IFileStream_IsEOR = (mDataCell.Value = END_OF_FILE)
End Function

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    '### ���̃��[�_�[�͂��̃C���^�[�t�F�C�X�@�\��ۏႵ�Ȃ� #########
    Err.Raise 999, "CWsTreeDataReader.AsIFileStream.SetLocation", "Does Not Support This Method !"
End Sub

Private Sub IFileStream_Rewind()
    '### �X�^�[�g�f�[�^�Z���I�u�W�F�N�g�̎擾 #######################
    Set mDataCell = mTreeDataSh.Range(START_CELL)
End Sub

Public Function AsIParameterReader() As IParameterReader
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IParameterReader�^�I�u�W�F�N�g
'
'���ӎ���:
'
    Set AsIParameterReader = Me
End Function

Public Function AsIFileStream() As IFileStream
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IFileStreamr�^�I�u�W�F�N�g
'
'���ӎ���:
'
    Set AsIFileStream = Me
End Function