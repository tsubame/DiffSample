VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAcquireActionInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �N���X�����C���X�^���X��Ԃ��܂��B
'�ړI:
'   AcquireAction�N���X�����C���X�^���X��Setter�AGetter�B
'   IAcquireAction��I/F�Ɏ��N���X��p�Ŏg���Ă��������B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��10��

Option Explicit

Public Property Set AcquireActionInstance(pInstance As IAcquireAction)
End Property

Public Property Get AcquireActionInstance() As IAcquireAction
End Property

