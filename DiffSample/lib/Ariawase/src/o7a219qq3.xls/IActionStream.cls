VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IActionStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   DC Test Scenario�p�̃��[�_�[/���C�^�[����������X�g���[���C���^�[�t�F�C�X
'
'�ړI:
'   DC Test Scenario���e�f�[�^�����[�h/���C�g����ۂ̃f�[�^�A�N�Z�X������s��
'
'�쐬��:
'   0145206097

Option Explicit

Public Sub MoveNextCategory()
End Sub

Public Sub MoveNextGroup()
End Sub

Public Sub MoveNextData()
End Sub

Public Sub SetCategoryLocation(ByVal dataName As String)
End Sub

Public Sub SetGroupLocation(ByVal GroupIndex As Long)
End Sub

Public Sub SetDataLocation(ByVal dataIndex As Long)
End Sub

Public Function IsEndOfCategory() As Boolean
End Function

Public Function IsEndOfGroup() As Boolean
End Function

Public Function IsEndOfData() As Boolean
End Function

Public Sub Rewind()
End Sub

Public Sub SeekChangedLocation()
End Sub