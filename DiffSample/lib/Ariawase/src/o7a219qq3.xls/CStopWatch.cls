VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   �^�C�}�[�N���X
'
'�ړI:
'   ���Ԃ��v������
'
'�쐬��:
'   0145206097
'
Option Explicit

Dim mTime As Double

Public Sub Start()
    mTime = TheExec.timer
End Sub

Public Sub Reset()
    mTime = 0
End Sub

Public Function GetLapTime() As Double
    GetLapTime = TheExec.timer(mTime)
End Function