VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestConditionController 
   Caption         =   "TestConditionController"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   OleObjectBlob   =   "TestConditionController.frx":0000
End
Attribute VB_Name = "TestConditionController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit


Public Event QueryClose(ByRef Cancel As Integer, ByVal CloseMode As Integer)   'TestConditionController�t�H�[���I�����m�C�x���g

Private Sub AbortButton_Click()
'Abort�{�^�����������ꂽ���̏���
    '�����ł͉������Ȃ�
    'CTestConditionController�N���X�ł��̃C�x���g���擾���A�����ŏ�������
    
End Sub

Private Sub ContinueButton_Click()
'Continue�{�^�����������ꂽ���̏���
    '�����ł͉������Ȃ�
    'CTestConditionController�N���X�ł��̃C�x���g���擾���A�����ŏ�������
    
End Sub

Private Sub ExecuteButton_Click()
'Execute�{�^�����������ꂽ���̏���
    '�����ł͉������Ȃ�
    'CTestConditionController�N���X�ł��̃C�x���g���擾���A�����ŏ�������

End Sub

Private Sub ReloadButton_Click()
'Reload�{�^�����������ꂽ���̏���
    '�����ł͉������Ȃ�
    'CTestConditionController�N���X�ł��̃C�x���g���擾���A�����ŏ�������
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'�~�{�^�����������ꂽ���̏���

    RaiseEvent QueryClose(Cancel, CloseMode)   '�~�{�^���ł͏I���ł��Ȃ��|��MsgBox�ŕ\������

End Sub
