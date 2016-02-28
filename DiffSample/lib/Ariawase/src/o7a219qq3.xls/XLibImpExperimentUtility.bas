Attribute VB_Name = "XLibImpExperimentUtility"
'�T�v:
'   ImSceExperimentController��Utility
'
'�ړI:
'   ImSceExperimentController�̏�����/�j���̃��[�e���e�B���`����
'
'�쐬��:
'   0145184306
'
Option Explicit

Public Const RESTART_ERROR_NUMBER As Long = 8000

Private mEnableExperimentMode As Boolean

Public Property Get EnableExperimentMode() As Boolean
'���e:
'   �����@�\�̏�Ԏ擾
'
'���l:
'
'
    EnableExperimentMode = mEnableExperimentMode
End Property

Public Property Let EnableExperimentMode(ByVal pEnable As Boolean)
'���e:
'   �����@�\�̏�Ԑݒ�
'
'���l:
'
'
    mEnableExperimentMode = pEnable
End Property

Public Function GetSubParamLabel(ByVal pPath As String, ByVal pCurPath As String) As String
'���e:
'   �p�����[�^�N���X�̃����o�ϐ��̃��x�����擾����
'
'[pPath]       IN String�^:     �����o�ϐ��̐�΃p�X
'[pCurPath]    IN String�^:     �p�����[�^�N���X�̃p�X
'
'���l:
'
'
    Dim myLabel As String
    myLabel = Mid$(pPath, Len(pCurPath) + 1)
    If myLabel Like "\*" Then
        myLabel = Mid$(myLabel, 2)
    End If
    Dim myIndex As Long
    myIndex = InStr(myLabel, "\")
    If myIndex > 0 Then
        myLabel = Left$(myLabel, myIndex - 1)
    End If
    GetSubParamLabel = myLabel
End Function

Public Function GetSubParamIndex(ByVal pPath As String, ByVal pCurPath As String) As Long
    GetSubParamIndex = CLng(Mid$(Strings.Left$(pPath, InStr(Len(pCurPath) + 1, pPath, ")") - 1), InStr(Len(pCurPath) + 1, pPath, "(") + 1))
End Function
