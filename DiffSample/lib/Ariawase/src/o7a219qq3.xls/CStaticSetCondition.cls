VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStaticSetCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �ÓI�����ݒ�N���X
'
'�ړI:
'   TestCondition�V�[�g����擾����Condition���]���̏����Ŏ��s����N���X
'
'�쐬��:
'   0145184935 K.Hara
'

Option Explicit

Implements ISetCondition

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

Public Sub IsetCondition_Execute(ByVal pFunctionInfo As CAcquireFunctionInfo)
'���e:
'   �]���̏����Ń}�N�����s���s��
'
'[pFunctionInfo] IN CAcquireFunctionInfo�^: FWSetCondition�s
'
'���l:
'
    '�]���̏����Ń}�N�����s���s��
    Call TheCondition.SetCondition(pFunctionInfo.InstanceName)

End Sub