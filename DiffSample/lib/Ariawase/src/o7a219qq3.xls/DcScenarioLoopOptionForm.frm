VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DcScenarioLoopOptionForm 
   Caption         =   "DC Test Scenario Looping Option"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7395
   OleObjectBlob   =   "DcScenarioLoopOptionForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "DcScenarioLoopOptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'���X�g�{�b�N�X��multiSelect��2-fmMultiSelectExtended�ɕύX
'�t�H���_�I���{�^���̃e�L�X�g�𔼊p�s���I�h�ɕύX
'�J�e�S���A�C�e���ړ��{�^���̃��C�A�E�g�z�u��ύX
'Form�N�����̃t�H���_�p�X���A���̃u�b�N�̃p�X(��JOB�t�@�C����Path)�Ɏw��



Option Explicit
Public Event QueryClose(Cancel As Integer, CloseMode As Integer)

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    RaiseEvent QueryClose(Cancel, CloseMode)
End Sub
