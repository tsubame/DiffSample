VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMeasure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �d���E�d������N���X����������C���^�[�t�F�C�X
'
'�ړI:
'   �d���E�d������N���X�����p�����[�^�̏��������s��
'   ���茋�ʂ�Result�N���X�I�u�W�F�N�g�֏����o��
'
'�쐬��:
'   0145206097
'
Option Explicit

Public Sub Initialize(ByVal measSite As Long, ByVal MeasurePins As Collection, ByVal setContainer As CContainer, ByVal reader As IParameterReader)
End Sub

Public Sub WriteResult(ByVal resultKeeper As CResultManager)
End Sub

Public Sub GetBoardInfo()
End Sub

Public Property Get MeasurePins() As Collection
End Property