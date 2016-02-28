VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScenarioParameterViewer 
   Caption         =   "ScenarioParameterViewer"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4875
   OleObjectBlob   =   "ScenarioParameterViewer.frx":0000
End
Attribute VB_Name = "ScenarioParameterViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








'�T�v:
'   �p�����[�^�\���p�t�H�[��
'
'�ړI:
'   ���C�^�[�Ń_���v���ꂽ����\������
'
'�쐬��:
'   0145184306
'
Option Explicit

Private m_Active As Boolean
Private m_EndStatus As Boolean

Public Sub Display()
'���e:
'   ���C�^�[�Ń_���v���ꂽ����\������B
'
'���l:
'
'
    Show vbModeless
    m_Active = True
    m_EndStatus = False
    While m_Active = True
        DoEvents
    Wend
End Sub

Private Sub btnEnd_Click()
'���e:
'   �����I���{�^��
'   �����ꂽ�ꍇ�́A�����I���t���O��True�ɂ���
'
'���l:
'
'
    m_Active = False
    m_EndStatus = True
    Me.ScenarioParamView.Value = ""

End Sub

Private Sub btnContinue_Click()
'���e:
'   OK�{�^��
'
'���l:
'
'
    m_Active = False
    m_EndStatus = False
    Me.ScenarioParamView.Value = ""
    
End Sub


Private Sub QuitEnable_Change()
'���e:
'   �����I���{�^����ON/OFF��؂�ւ���
'
'���l:
'
'
    If QuitEnable = True Then
        btnEnd.enabled = True
    Else
        btnEnd.enabled = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'���e:
'   "X"�{�^���������ꂽ���̓���
'
'���l:
'
'
    If CloseMode = vbFormControlMenu Then
        btnContinue_Click
        Cancel = True
    End If
    Me.ScenarioParamView.Value = ""

End Sub

Property Get EndStatus() As Boolean
'���e:
'   �����I���t���O�v���p�e�B
'
'�߂�l:
'   �����I���t���O(Boolean�^)
'
'���l:
'
'
    EndStatus = m_EndStatus
End Property

Property Let EndStatus(pStatus As Boolean)
'���e:
'   �����I���t���O�v���p�e�B
'
'����:
'   �����I���t���O(Boolean�^)
'
'���l:
'
'
    m_EndStatus = pStatus
End Property

Private Sub UserForm_Initialize()
'���e:
'   �R���X�g���N�^
'
'���l:
'
'
    With Me
        .ScenarioParamView.Text = ""
        .ScenarioParamView.Locked = True
        .QuitEnable = False
        btnEnd.enabled = False
    End With
    m_Active = True
    m_EndStatus = False
End Sub

Private Sub UserForm_Terminate()
'���e:
'   �f�X�g���N�^
'
'���l:
'
'
    With Me
        .ScenarioParamView.Text = ""
        .QuitEnable = False
        btnEnd.enabled = False
    End With
    m_Active = False
    m_EndStatus = False
End Sub



