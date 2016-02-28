VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetFunctionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   �Z�b�g�֐����N���X
'
'�ړI:
'   �Z�b�g�֐��p�����[�^�̌^�Ƃ��Ďg�p
'
'�쐬��:
'   SLSI����
'
'���ӓ_:
'   �e�v���p�e�B�ւ̒l�̏������݂�Write at once�ł�
'   �p�����[�^�̒l��Object�͐ݒ�ł��܂���
'   TheError�Ƃ��Č��J����Ă���G���[�}�l�[�W��Object���K�v�ł�
'
Option Explicit

Private m_ConditionName As String              '�����������
Private m_SetFunctionName As String            '�����ݒ�@�\����
Private m_ConditionParameters As Collection    '�����ݒ�@�\�̃p�����[�^�l�̕ۊǗp
Private m_IsEnableMacro As Boolean             '�}�N���̗L�����
Private m_MacroController As IMacroController  '�}�N�����s�@�\��Object
Private m_ConditionSheetName As String         '�����ݒ�V�[�g�̖��O
Private m_ExecuteSettingName As String         '�}�N���O���[�v��`��
Private Const ERR_NUMBER As Long = 9999        '�G���[���ɓn���G���[�ԍ�

Public Sub RunMacro()
'���e:
'   Excel�}�N���̎��s����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'   MacroController��Object���ݒ肳��Ă��Ȃ��ꍇ�̓G���[�ƂȂ�܂�
'
    
    Call m_MacroController.RunMacro(Me)

End Sub

Public Property Get MacroController() As IMacroController
'���e:
'   Excel�}�N�����s�@�\��Object
'
'�p�����[�^:
'
'�߂�l:
'   Excel�}�N�����s�@�\��Object
'
'���ӎ���:
'
    Set MacroController = m_MacroController
    
End Property

Public Property Set MacroController(ByRef pController As IMacroController)
'���e:
'   Excel�}�N�����s�@�\��Object
'
'�p�����[�^:
'   [pController]  In  IMacroController�^: Excel�}�N�����s�@�\��Object
'
'�߂�l:
'
'���ӎ���:
'
    
    Set m_MacroController = pController
    
End Property

Public Property Let ConditionSheetName(ByVal pWorksheetName As String)
'���e:
'   TestCondition���[�N�V�[�g�̖��O
'
'�p�����[�^:
'   [pWorksheetName]  In  string�^:�@���[�N�V�[�g��
'
'�߂�l:
'
'���ӎ���:
'   �l�̐ݒ��Write at once�ł�
'
    
    Static WriteFlag As Boolean
    If WriteFlag = False Then
        m_ConditionSheetName = pWorksheetName
        WriteFlag = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get ConditionSheetName() As String
'���e:
'   TestCondition���[�N�V�[�g�̖��O
'
'�p�����[�^:
'
'�߂�l:
'   TestCondition���[�N�V�[�g��
'
'���ӎ���:
'
    ConditionSheetName = m_ConditionSheetName

End Property

Public Property Let IsEnableMacro(ByVal pSetValue As Boolean)
'���e:
'   �}�N���L�������ݒ�l
'
'�p�����[�^:
'   [pSetValue]  In  Boolean�^:   �ݒ�l(True:�L��)
'
'�߂�l:
'
'���ӎ���:
'   �l�̐ݒ��Write at once�ł�
'
    
    Static WriteFlag As Boolean
    If WriteFlag = False Then
        m_IsEnableMacro = pSetValue
        WriteFlag = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get IsEnableMacro() As Boolean
'���e:
'   �}�N���L�������ݒ�l
'
'�p�����[�^:
'
'�߂�l:
'   �ݒ�l (True:�L��)
'
'���ӎ���:
'
    
    IsEnableMacro = m_IsEnableMacro

End Property

Public Property Let ConditionName(ByVal pName As String)
'���e:
'   �����������
'
'�p�����[�^:
'   [pName]  In  String�^:   ���������������
'
'�߂�l:
'
'���ӎ���:
'   �l�̐ݒ��Write at once�ł�
'
    Static WriteFlag As Boolean
    If WriteFlag = False Then
        m_ConditionName = pName
        WriteFlag = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get ConditionName() As String
'���e:
'   �����������
'
'�p�����[�^:

'�߂�l:
'   ���������������
'
'���ӎ���:
'
    ConditionName = m_ConditionName

End Property

Public Property Let FunctionName(ByVal pName As String)
'���e:
'   �Z�b�g�֐�����
'
'�p�����[�^:
'   [pName]  In  String�^:   �Z�b�g�֐���������
'
'�߂�l:
'
'���ӎ���:
'   �l�̐ݒ��Write at once�ł�
'
    Static WriteFlag As Boolean
    If WriteFlag = False Then
        m_SetFunctionName = pName
        WriteFlag = True
    Else
        Call RaiseError("It has already set it")
    End If

End Property

Public Property Get FunctionName() As String
'���e:
'   �Z�b�g�֐�����
'
'�p�����[�^:
'
'�߂�l:
'   �Z�b�g�֐���������
'
'���ӎ���:
'
    FunctionName = m_SetFunctionName

End Property

Public Property Get ArgParameterCount() As Long
'���e:
'   �Z�b�g�֐��̓o�^�ς݃p�����[�^��
'
'�p�����[�^:
'
'�߂�l:
'   �Z�b�g�֐��̓o�^�ς݃p�����[�^��
'
'���ӎ���:
'   Read only�ł�
'
    ArgParameterCount = m_ConditionParameters.Count

End Property

Public Sub AddParameter(ByVal pParameter As Variant)
'���e:
'   �Z�b�g�֐��̃p�����[�^��o�^����
'
'�p�����[�^:
'   [Parameter]  In  �Z�b�g�֐��̃p�����[�^�l
'
'�߂�l:
'
'���ӎ���:
'   Object�͒l�Ƃ��ēo�^�ł��܂���
'
    If IsObject(pParameter) Then
        '�p�����[�^��Object�̎��͓o�^�����Ȃ��ŃG���[�Ƃ���
        Call RaiseError("The object cannot be registered as a parameter")
    Else
        Call m_ConditionParameters.Add(pParameter)
    End If

End Sub

Public Property Get Arg(ByVal pArgNumber As Long) As Variant
'���e:
'   �o�^����Ă���Z�b�g�֐��̃p�����[�^�l���擾
'
'�p�����[�^:
'   [pArgNumber]  In  �擾����p�����[�^�ԍ��l
'
'�߂�l:
'   �w�肳�ꂽ�ԍ��̃p�����[�^�l
'
'���ӎ���:
'   �w�肳�ꂽ�ԍ��̃p�����[�^�����݂��Ȃ��ꍇ�A
'   �擾�l��Empty�ƂȂ�܂�
'
    Arg = GetParameter(pArgNumber, m_ConditionParameters)

End Property
Public Property Let ExecuteSettingName(ByVal pExecuteSettingName As String)
'���e:
'   �}�N���O���[�v��`��
'
'�p�����[�^:
'   [pExecuteSettingName]  In  String�^:   �Z�b�g�}�N���O���[�v��`��������
'
'�߂�l:
'
'���ӎ���:
'   �l�̐ݒ��Write at once�ł�
'

    Static WriteFlag As Boolean
    If WriteFlag = False Then
        m_ExecuteSettingName = pExecuteSettingName
        WriteFlag = True
    Else
        Call RaiseError("It has been already set.")
    End If

End Property

Public Property Get ExecuteSettingName() As String
'���e:
'   �}�N���O���[�v��`��
'
'�p�����[�^:
'
'�߂�l:
'   �}�N���O���[�v��`��
'
'���ӎ���:
'
    ExecuteSettingName = m_ExecuteSettingName

End Property

Public Function Clone() As CSetFunctionInfo
'���e:
'   �N���[��Object��񋟂���
'
'�p�����[�^:
'
'�߂�l:
'   �N���[��Object
'
'���ӎ���:
'
    
    '�N���[���pObj����
    Dim cloneobj As CSetFunctionInfo
    Set cloneobj = New CSetFunctionInfo

    '���̕�������
    With cloneobj
        .ConditionName = m_ConditionName                  '�������̂̓o�^
        .FunctionName = m_SetFunctionName                 '�Z�b�g�֐����̂̓o�^
        .IsEnableMacro = m_IsEnableMacro                  '�}�N���L���ݒ�l
        .ConditionSheetName = m_ConditionSheetName        '�����\���[�N�V�[�g��
        .ExecuteSettingName = m_ExecuteSettingName        '�}�N���O���[�v��`��
        Set .MacroController = m_MacroController          '�}�N�����s�@�\��Object
        Dim tmpRetParameter As Variant
        For Each tmpRetParameter In m_ConditionParameters '�Z�b�g�֐��̃p�����[�^�̓o�^
            .AddParameter (tmpRetParameter)
        Next tmpRetParameter
    End With

    '���ݒ�ς݂�Object��Ԃ�
    Set Clone = cloneobj

End Function

'�R���X�g���N�^
Private Sub Class_Initialize()
    Call ClearInfomation
End Sub

'���������񏉊���
Private Sub ClearInfomation()
    m_ConditionName = ""
    m_SetFunctionName = ""
    m_IsEnableMacro = False
    m_ConditionSheetName = ""
    m_ExecuteSettingName = ""
    Set m_MacroController = Nothing
    Set m_ConditionParameters = New Collection
End Sub

'�p�����[�^�܂��w��ԍ��̃p�����[�^�l�����o���֐��i�w��ԍ��̃p�����[�^�����݂��Ȃ�����Empty�l��Ԃ��j
Private Function GetParameter(ByVal pArgNo As Long, ByVal pParemeterCollection As Collection) As Variant
    
    On Error GoTo DATA_GET_ERROR
    GetParameter = pParemeterCollection.Item(pArgNo + 1)
    On Error GoTo 0
    Exit Function

DATA_GET_ERROR:
    GetParameter = Empty
    Call RaiseError("Arg=" & pArgNo & " parameter is nothing")

End Function

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

