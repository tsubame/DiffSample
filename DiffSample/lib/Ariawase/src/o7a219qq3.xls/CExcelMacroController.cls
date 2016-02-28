VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExcelMacroController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ��������ݒ�p�̃}�N�������s����
'
'�ړI:
'   �}�N���L�������؂�ւ��Ή�
'

Option Explicit

Implements IMacroController                           '�}�N�����s�@�\�pI/F
Private m_TargetBookName As String                    '�}�N���̑��݂��郏�[�N�u�b�N��

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
    
    m_TargetBookName = ""

End Sub

Public Property Let TargetBookName(ByVal pBookName As String)
'���e:
'   �}�N���̑��݂���WorkBook��
'
'�p�����[�^:
'   [pBookName]  In  String�^:   ���[�N�u�b�N��
'
'�߂�l:
'
'���ӎ���:
'
    
    m_TargetBookName = pBookName

End Property

Public Property Get TargetBookName() As String
'���e:
'   �}�N���̑��݂���WorkBook��
'
'�p�����[�^:
'
'�߂�l:
'   ���[�N�u�b�N��
'
'���ӎ���:
'
    
    TargetBookName = m_TargetBookName

End Property

'I/F�p�̎���
Private Sub IMacroController_RunMacro(ByRef pParameters As CSetFunctionInfo)
'���e:
'   Excel�}�N���̎��s
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    
    Call Excel.Application.Run("'" & m_TargetBookName & "'!" & pParameters.FunctionName, pParameters)

End Sub