VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetConditionHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   �����ݒ藚�����p�N���X
'
'�ړI:
'   �ݒ藚�����̌^�Ƃ��Ďg�p
'
'�쐬��:
'   SLSI����
'
'���ӓ_:
'   CSetFunctionInfo ���K�v�ł��B
'   �e�v���p�e�B�ւ̒l�̏������݂�Write at once�ł�
'   TheError�Ƃ��Č��J����Ă���G���[�}�l�[�W��Object���K�v�ł�
'
Option Explicit

Private m_RecordDate As String                        '�����̋L�^��
Private m_RecordTime As String                        '�����̋L�^����
Private m_ExecutionTime As Double                     '�Z�b�g�֐��̎��s���ԕۑ��p[Sec]
Private m_ExecTestInstancesName As String             '�Z�b�g�֐������s���ꂽ�e�X�g�C���X�^���X����
Private m_SetFunctionParameterObj As CSetFunctionInfo '�Z�b�g�֐��̃p�����[�^OBJ�p
Private Const ERR_NUMBER As Long = 9999               '�G���[���ɓn���G���[�ԍ�

Public Property Let RecordTime(ByVal pTime As String)
'���e:
'   �����L�^���̎���
'
'�p�����[�^:
'   [pTime]  In  String�^:   �L�^���̎���
'
'�߂�l:
'
'���ӎ���:
'   �l�̐ݒ��Write at once�ł�
'
    Static WriteFlg As Boolean
    If WriteFlg = False Then
        m_RecordTime = pTime
        WriteFlg = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get RecordTime() As String
'���e:
'   �����L�^���̎���
'
'�p�����[�^:
'
'�߂�l:
'   �����L�^���̎���
'
'���ӎ���:
'
    RecordTime = m_RecordTime

End Property

Public Property Let RecordDate(ByVal pRecDate As String)
'���e:
'   �����L�^���̓��t
'
'�p�����[�^:
'   [RecDate]  In  String�^:   �L�^���̓��t
'
'�߂�l:
'
'���ӎ���:
'   �l�̐ݒ��Write at once�ł�
'
    Static WriteFlg As Boolean
    If WriteFlg = False Then
        m_RecordDate = pRecDate
        WriteFlg = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get RecordDate() As String
'���e:
'   �����L�^���̓��t
'
'�p�����[�^:
'
'�߂�l:
'   �����L�^���̓��t
'
'���ӎ���:
'
    RecordDate = m_RecordDate

End Property

Public Property Let ExecTime(ByVal pTime As Double)
'���e:
'   �Z�b�g�֐��̎��s����
'
'�p�����[�^:
'   [pTime]  In  Double�^:   ���s����
'
'�߂�l:
'
'���ӎ���:
'   �l�̐ݒ��Write at once�ł�
'
    Static WriteFlg As Boolean
    If WriteFlg = False Then
        m_ExecutionTime = pTime
        WriteFlg = True
    Else
        Call RaiseError("It has already set it")
    End If
End Property

Public Property Get ExecTime() As Double
'���e:
'   �Z�b�g�֐��̎��s����
'
'�p�����[�^:
'
'�߂�l:
'   �Z�b�g�֐��̎��s����
'
'���ӎ���:
'
    ExecTime = m_ExecutionTime

End Property

Public Property Let ExecTestInstanceName(ByVal pName As String)
'���e:
'   �Z�b�g�֐����s���̃e�X�g�C���X�^���X��
'
'�p�����[�^:
'   [pName]  In  String�^:   �e�X�g�C���X�^���X����
'
'�߂�l:
'
'���ӎ���:
'   �l�̐ݒ��Write at once�ł�
'
    Static WriteFlg As Boolean
    If WriteFlg = False Then
        m_ExecTestInstancesName = pName
        WriteFlg = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get ExecTestInstanceName() As String
'���e:
'   �Z�b�g�֐����s���̃e�X�g�C���X�^���X��
'
'�p�����[�^:
'
'�߂�l:
'   �Z�b�g�֐����s���̃e�X�g�C���X�^���X��
'
'���ӎ���:
'
    ExecTestInstanceName = m_ExecTestInstancesName

End Property

Public Property Set SetFunctionInfo(ByVal pParameterObj As CSetFunctionInfo)
'���e:
'   ���s���̃Z�b�g�֐����
'
'�p�����[�^:
'   [ParameterObj]  In  CSetFunctionInfo�^:  �Z�b�g�֐����Object
'
'�߂�l:
'
'���ӎ���:
'   �l�̐ݒ��Write at once�ł�
'
    If m_SetFunctionParameterObj Is Nothing Then
        Set m_SetFunctionParameterObj = pParameterObj
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get SetFunctionInfo() As CSetFunctionInfo
'���e:
'   ���s���̃Z�b�g�֐����
'
'�p�����[�^:
'
'�߂�l:
'   ���s���̃Z�b�g�֐����Object
'
'���ӎ���:
'
    Set SetFunctionInfo = m_SetFunctionParameterObj

End Property

'�R���X�g���N�^
Private Sub Class_Initialize()
    Call ClearInfomation
End Sub

'���������񏉊���
Private Sub ClearInfomation()
    m_RecordDate = ""
    m_RecordTime = ""
    m_ExecutionTime = 0#
    m_ExecTestInstancesName = ""
    Set m_SetFunctionParameterObj = Nothing
End Sub

'�G���[���C�Y�p�֐�
Private Sub RaiseError(ByVal pErrMsg As String)
    If Not (TheError Is Nothing) Then
        '�G���[�}�l�[�W����
        Call TheError.Raise(ERR_NUMBER, TypeName(Me), pErrMsg)
    Else
        'VB��Err��
        Call Err.Raise(ERR_NUMBER, TypeName(Me), pErrMsg)
    End If
End Sub
