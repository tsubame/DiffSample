Attribute VB_Name = "xEeeAuto_Rank"
'�T�v:
'   Rank Sheet�̑��݊m�F
'
'�ړI:
'   Rank�V�[�g�����ɍs���Ă�������True��Ԃ�
'
'�쐬��:
'   2012/03/12 D.Maruyama
'   2012/11/12 H.Arikawa  RankSheet�ɍ��ڋL�ڂ����邩�`�F�b�N���郋�[�`����ǉ��B

Private Enum EeeAutoRankState
    UNKOWN
    INITALIZED
End Enum

Private m_IsRankSheet_Exist As Boolean
Private m_State As EeeAutoRankState

Private Const RANK_SHEET_NAME As String = "rank_sheet"

Option Explicit

'���e:
'   ���̃��W���[���̏�����
'
'���l:
'   RANKSHEET�̂���Ȃ��𔻒f����
'
Public Sub InitializeEeeAutoRank()

    m_IsRankSheet_Exist = False

    Dim mySheet As Worksheet
    
    For Each mySheet In ThisWorkbook.Worksheets
        If mySheet.Name = RANK_SHEET_NAME Then
            m_IsRankSheet_Exist = False
            Exit For
        End If
    Next mySheet
    
    m_State = INITALIZED
        
End Sub

'���e:
'   RANK���������邩�ǂ�����Ԃ�
'
'���l:
'   RANKSHEET�̂���ꍇ True�A�Ȃ��ꍇ False
'
'�L��R�����g
'   RANKSHEET�͑S�^�C�v�}�������̂ŁA���g�̋L�ڂ����邩���`�F�b�N����B
'   �Z���w��Œl�������Ă��邩���`�F�b�N������B

Public Function IsRankEnable() As Boolean

    Dim tname As String

    If m_State <> INITALIZED Then
        Err.Raise 9999, "IsRankEnable", "xEeeAutoRank is not Initialized!"
        IsRankEnable = False
        Exit Function
    End If
        
    IsRankEnable = m_IsRankSheet_Exist
    
    tname = Worksheets("Tenken").Range("B9").Value
    
    If tname = "" Then
        IsRankEnable = False
    End If
    
End Function
