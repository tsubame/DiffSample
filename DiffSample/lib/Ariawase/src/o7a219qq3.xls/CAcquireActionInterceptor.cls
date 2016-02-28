VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcquireActionInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   CAcquireAction,CPostAcquireAction�Ɋ��荞��ŁA���s�󋵂����O�Ɏc���܂��B
'�ړI:
'   ���O�o�͗p
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��8��

Option Explicit

Implements IAcquireAction
Implements IAcquireActionInstance

Private Const ERR_NUMBER = 9999             ' �G���[�ԍ���ێ�����
Private Const CLASS_NAME = "CAcquireActionInterceptor" ' Class���̂�ێ�����

Private m_IAcquireAction As IAcquireAction

Private m_ActionLogger As CActionLogger

Public Property Get AsIAcquireAction() As IAcquireAction
    Set AsIAcquireAction = Me
End Property

Public Property Set actionLogger(pActionlog As CActionLogger)
    Set m_ActionLogger = pActionlog
End Property

Private Property Set IAcquireActionInstance_AcquireActionInstance(pInstance As IAcquireAction)
'���e:
'   ���̃N���X��IAcquireAction�C���^�[�t�F�C�X�̃C���X�^���X���Z�b�g���܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   IAcquireAction
'�߂�l:
'
'���ӎ���:
'
    Set m_IAcquireAction = pInstance

End Property

Private Property Get IAcquireActionInstance_AcquireActionInstance() As IAcquireAction
'���e:
'   ���̃N���X������IAcquireAction�C���^�[�t�F�C�X�̃C���X�^���X��Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   IAcquireActionInstance
'
'���ӎ���:
'
    Set IAcquireActionInstance_AcquireActionInstance = m_IAcquireAction

End Property

Private Function IAcquireAction_Execute() As Integer
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Integer
'
'���ӎ���:
'
'
    Dim pStatus As Long
    
    '#####  ActionLogger�̏��𐶐�����  #####
    Dim p_ActionInfo As CActionInfo
    Set p_ActionInfo = New CActionInfo

    With p_ActionInfo
        .AcquireGroupName = m_IAcquireAction.AcquireFunctionInfo.GroupName
        .AcquireInstanceName = m_IAcquireAction.AcquireFunctionInfo.InstanceName
        .FrameworkName = m_IAcquireAction.AcquireFunctionInfo.macroName
        .TestInstanceName = TheExec.DataManager.InstanceName
        Set .Parameters = m_IAcquireAction.AcquireFunctionInfo.Clone
    End With

    Call m_ActionLogger.LogBeforeExecution(p_ActionInfo)
    
    pStatus = m_IAcquireAction.Execute

    Call m_ActionLogger.LogAfterExecution(p_ActionInfo, pStatus)
    
    IAcquireAction_Execute = pStatus

End Function

Private Function IAcquireAction_ToptStart() As Integer
    IAcquireAction_ToptStart = m_IAcquireAction.ToptStart
End Function

Private Function IAcquireAction_ClearStatus() As Integer
    IAcquireAction_ClearStatus = m_IAcquireAction.ClearStatus
End Function

Private Property Get IAcquireAction_GetFrameWorkName() As String
    IAcquireAction_GetFrameWorkName = m_IAcquireAction.GetFrameWorkName
End Property

Private Property Get IAcquireAction_GetStatus() As Integer
    IAcquireAction_GetStatus = m_IAcquireAction.GetStatus
End Property

Private Sub IAcquireAction_Reset()
    Call m_IAcquireAction.Reset
End Sub

Private Sub IAcquireAction_Initialize(reader As IParameterReader)
    Call m_IAcquireAction.Initialize(reader)
End Sub

Private Property Get IAcquireAction_AcquireFunctionInfo() As CAcquireFunctionInfo
    Set IAcquireAction_AcquireFunctionInfo = m_IAcquireAction.AcquireFunctionInfo
End Property

Private Sub Class_Initialize()
    Set m_IAcquireAction = Nothing
    Set m_ActionLogger = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_IAcquireAction = Nothing
    Set m_ActionLogger = Nothing
End Sub




