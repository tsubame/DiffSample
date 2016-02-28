VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' �N���X��  �F  csPALS.cls
' �T�v      �F  PALS�S�̂𐧌䂷��N���X
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/09/01�@�V�K�쐬        K.Sumiyashiki
'==========================================================================

Public CommonInfo As csPALS_Common      'csPALS_Common�N���X�̒�`

Public LoopParams As csPALS_LoopMain    'csPALS_LoopMain�N���X�̒�`


'*********************************************************************************
' �֐���    �F  Class_Initialize
' �T�v      �F  �C���X�^���X�������̏�����
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/09/01�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()

On Error GoTo errPALSClass_Initialize

    Set CommonInfo = New csPALS_Common
        
    Set LoopParams = New csPALS_LoopMain

Exit Sub

errPALSClass_Initialize:
    Call sub_errPALS("csPALS not formed at 'Class_Initialize'", "0-4-01-0-13")

End Sub


'*********************************************************************************
' �֐���    �F  Class_Terminate
' �T�v      �F  �C���X�^���X�J�����̏���
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/09/01�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Terminate()
    Set CommonInfo = Nothing
    Set LoopParams = Nothing
End Sub
