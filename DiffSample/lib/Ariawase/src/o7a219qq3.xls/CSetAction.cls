VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   AcquireFrameWork�iFWSetCondition�j�̏���ێ������s����N���X
'
'   Revision History:
'       Data        Description
'       2010/04/28  AcquireFrameWork�iFWSetCondition�j�����s����@�\����������
'       2010/06/11  �f�[�^�\���������̂��߁A�v���O�����R�[�h��ύX����
'       2010/06/22  ���s�O��̏���ǉ������im_Status��ύX�j
'                   ���s��񃊃Z�b�g�@�\����������
'       2011/11/22  UserMacro�̎��s�𓮓I/�ÓI�ɐ؂�ւ���@�\����������
'
'�쐬��:
'   tomoyoshi.takase
'   0145184346
'

Option Explicit

Implements IAcquireAction

Private Const ERR_NUMBER = 9999         ' �G���[�ԍ���ێ�����
Private Const CLASS_NAME = "CSetAction" ' Class���̂�ێ�����
Private Const DYNAMIC_WORD As String = "debug" ' ���I�����ݒ�p�̃L�[���[�h

Private m_FunctionInfo As CAcquireFunctionInfo ' ImageACQTBL�V�[�g1�s�̃f�[�^��ێ�����
'Private m_ActionLogger As CActionLogger        ' ���sLog�f�[�^���_���v����ActionLogger��ێ�����
Private m_InstanceName As String               ' TestInstance���̂�ێ�����
Private m_ActionName As String                 ' AcquireFrameWork�����s����UserMacro���̂�ێ�����
Private m_Status As Long                       ' ���s���ʂ�ێ�����i-1�F���s�O�@0�F����I���@1�F�G���[�����j
Private m_ISetCondition As ISetCondition       ' �����ݒ�N���X�̃C���^�[�t�F�[�X

Public Property Get AsIAcquireAction() As IAcquireAction
    Set AsIAcquireAction = Me
End Property

Private Function IAcquireAction_Execute() As Integer
'���e:
'   SetCondition��UserMacro�����s����
'
'�p�����[�^:
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'
    Call m_ISetCondition.Execute(m_FunctionInfo)
    m_Status = TL_SUCCESS

    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    m_Status = TL_ERROR
    Call TheError.Raise(ERR_NUMBER, "TheCondition.SetCondition", "Can not Execute Set Condition Macro: " & m_InstanceName)


End Function

Private Function IAcquireAction_ToptStart() As Integer
'���e:
'   TestOptimizer�@�\���g���āAAcquireFrameWork�����s����
'
'�p�����[�^:
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    '#####  AcquireFrameWork�����s����  #####
    With m_FunctionInfo
        If .macroName <> "" Then
            On Error GoTo IGXLErrHandler
            Call TheHdw.TOPT.Start(.GroupName, .macroName, .InstanceName)
        Else
            On Error GoTo ErrHandler
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".ToptStart", "Can Not Found Acquire Frame Work Name ! ( " _
                                & .GroupName & " / " & .InstanceName & " )")
        End If
    End With


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
IGXLErrHandler:
    IAcquireAction_ToptStart = TL_ERROR
    m_Status = IAcquireAction_ToptStart
    Exit Function
ErrHandler:
    IAcquireAction_ToptStart = TL_ERROR
    m_Status = IAcquireAction_ToptStart
    Exit Function


End Function

Private Function IAcquireAction_ClearStatus() As Integer
'���e:
'   AcquireFrameWork�C���X�^���X�̎��s�������폜����
'
'�p�����[�^:
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWork�C���X�^���X�̎��s�������폜����  #####
    Call TheHdw.TOPT.ClearStatus(m_FunctionInfo.GroupName)


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    IAcquireAction_ClearStatus = TL_ERROR
    Exit Function


End Function

Private Property Get IAcquireAction_GetFrameWorkName() As String
    IAcquireAction_GetFrameWorkName = m_FunctionInfo.macroName
End Property

Private Property Get IAcquireAction_GetStatus() As Integer
    IAcquireAction_GetStatus = m_Status
End Property

Private Sub IAcquireAction_Reset()
    m_Status = -1
End Sub

Private Property Get IAcquireAction_AcquireFunctionInfo() As CAcquireFunctionInfo
    Set IAcquireAction_AcquireFunctionInfo = m_FunctionInfo
End Property

Private Sub IAcquireAction_Initialize(reader As IParameterReader)
    Set m_FunctionInfo = New CAcquireFunctionInfo
    Call m_FunctionInfo.Initialize(reader)
    m_InstanceName = reader.ReadAsString("Arg0@Parameters")
    
    '�����œ��Ior�ÓI�̏����ݒ�C���X�^���X���쐬
    Dim DebugWord As String
    DebugWord = LCase(reader.ReadAsString("Arg4@Parameters"))
    If StrComp(DebugWord, DYNAMIC_WORD, 0) = 0 Then

        '���I�����ݒ�N���X���쐬
        Dim DynamicSetCondition As CDynamicSetCondition
        Set DynamicSetCondition = New CDynamicSetCondition
        Set m_ISetCondition = DynamicSetCondition
    
    ElseIf StrComp(DebugWord, "", 0) = 0 Then
    
        '�ÓI�����ݒ�N���X���쐬
        Dim StaticSetCondition As CStaticSetCondition
        Set StaticSetCondition = New CStaticSetCondition
        Set m_ISetCondition = StaticSetCondition
        
    Else
        'Arg4��"debug"�A�܂���""(��)�łȂ��ꍇ
        
        m_Status = TL_ERROR
        Call TheError.Raise(ERR_NUMBER, "CSetAction", "Please input ""debug"" or Null Character into Arg4." & vbCr & vbCr & "Input Char = " & DebugWord)

    End If
    
    m_Status = -1
End Sub

Private Sub Class_Initialize()
    Set m_FunctionInfo = Nothing
    Set m_ISetCondition = Nothing
    m_InstanceName = ""
    m_ActionName = ""
    m_Status = -1
End Sub

Private Sub Class_Terminate()
    Set m_FunctionInfo = Nothing
    Set m_ISetCondition = Nothing
    m_InstanceName = ""
    m_ActionName = ""
    m_Status = -1
End Sub

