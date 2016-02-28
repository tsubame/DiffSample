VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMacroExecInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �}�N�����s���Ƀ��O���擾���邽�߂̃��b�p�[�N���X
'
'�ړI:
'   ���O���擾���Ȃ���}�N�����s���s��
'

Option Explicit

Implements IMacroController '�}�N�����s�@�\�pI/F

Private m_Logger As CSetConditionLogger        '���K�[Object
Private m_MacroController As IMacroController  '�}�N�����s�@�\��Object

Public Sub Initialize()
'���e:
'   �S�̋@�\�̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    
    Set m_Logger = Nothing
    Set m_MacroController = Nothing

End Sub

Public Property Set SetConditionLogger(ByRef pLoggerObject As CSetConditionLogger)
'���e:
'   �����ݒ藚���L�^�p���K�[Object
'
'�p�����[�^:
'   [pLoggerObject]  In  CSetConditionLogger�^: ���K�[Object

'�߂�l:
'
'���ӎ���:
'
    
    Set m_Logger = pLoggerObject

End Property

Public Property Get SetConditionLogger() As CSetConditionLogger
'���e:
'   �����ݒ藚���L�^�p���K�[Object
'
'�p�����[�^:

'�߂�l:
'   ���K�[Object
'
'���ӎ���:
'
    
    Set SetConditionLogger = m_Logger

End Property

Public Property Get MacroController() As IMacroController
'���e:
'   �}�N�����s�@�\Object
'
'�p�����[�^:

'�߂�l:
'   �}�N�����s�@�\��Object
'
'���ӎ���:
'
    
    Set MacroController = m_MacroController

End Property

Public Property Set MacroController(ByRef pController As IMacroController)
'���e:
'   �}�N�����s�@�\Object
'
'�p�����[�^:
'   [pController]  In  IMacroController�^: �}�N�����s�@�\��Object
'
'�߂�l:
'
'���ӎ���:
'
    
    Set m_MacroController = pController

End Property

Private Sub IMacroController_RunMacro(ByRef pParameters As CSetFunctionInfo)
'���e:
'   �}�N���̎��s����
'
'�p�����[�^:
'   [pParameters]  In  CSetFunctionInfo�^: �}�N�����s�Ɏg�p����p�����[�^
'
'�߂�l:
'
'���ӎ���:
'
    
    Dim ExecutionTime As Double
    ExecutionTime = 0#
    
    '���s���Ԃ��v�����Ȃ���}�N�������s����
    If pParameters.IsEnableMacro = True Then
        Dim TimerObj As CEeeTimer
        Set TimerObj = New CEeeTimer
        With TimerObj
            Call .Initialize
            Call .StartTimer
            'Excel�}�N���̎��s
            Call m_MacroController.RunMacro(pParameters)
            ExecutionTime = .StopTimer
        End With
    End If

    '���s���O�L�^����
    Call m_Logger.AddRecord(ExecutionTime, pParameters)
    
End Sub