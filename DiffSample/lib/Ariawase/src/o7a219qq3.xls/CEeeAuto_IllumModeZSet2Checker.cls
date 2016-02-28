VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEeeAuto_IllumModeZSet2Checker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �R���f�B�V�����`�F�b�J�@�����ޔ��̏ȗ��ۂ𔻒f����
'
'�ړI:
'   ���̂Ƃ��̏����ݒ����������擾����(OptModOrModZ2�p)
'�쐬��:
'   �ێR Create 2013/02/05
'   �L�� �ҏW Ver1.1 2013/02/25
'   �ێR/�L�� Add 2013/02/08 2013/03/11�ǉ�
'
Option Explicit

Private Const THIS_CLASS_NAME As String = "CEeeAuto_IllumModeZSet2Checker"

Private m_IsEnable As Boolean 'FW_OptModOrModZ2�����s���ׂ����ǂ���

Private m_eEndDownPos As DownPosi '�ŏI�I�Ȉړ���̓�����
Private m_iEndFNumberTurret As Integer '�ŏI�I�Ȉړ���̃^�[���b�g�ʒu


Implements IEeeAuto_TestConditionChecker

'�N���X�̏�����
Private Sub Class_Initialize()
    m_IsEnable = True
End Sub

'OptModOrModZ2�ݒ�O�̏������擾����
'���{���̈Ӗ��ł��邪�AOptModOrModZ1���s�����ǂ����́A���݈ʒu�ƍŏI�I��
'�s���悾���Ō��܂��Ă��܂����߁A���̊֐��Ŕ��f�܂ōs���B
'OptModOrModZ2���Ă΂��^�C�~���O�ł́A�u�ޔ��v-�u�ޔ��ɕK�v�Ȑݒ�v�͊������Ă���
'�P���ɍŏI�ȍs����Ƃ����Ă��邩�ǂ����Ŕ��f���Ă悢�B
'�ŏI�I�ȍs�����[IEeeAuto_TestConditionChecker->SetEndPosition]�Őݒ肵�Ă�������
Private Sub IEeeAuto_TestConditionChecker_CheckBeforeCondition()

    '�����l�͐�΂��
    m_IsEnable = True
    
    Dim eCurrentPos As DownPosi
    Dim iFNumberTurret As Integer
    
    If OptCond.IllumMaker = NIKON Then
    
        '���݂�F�l�A�������擾
        eCurrentPos = GetCurrentDownPosition
        iFNumberTurret = GetCurrentFNumberTurret
        
        Select Case OptCond.IllumModel
            Case "N-SIS3", "N-SIS5", "N-SIS5KAI"
                'F�l�A�������Ƃ��Ɉړ���Ɠ����ȏꍇ�̓K�`�����ŏȗ��Ƃ���
                If m_eEndDownPos = eCurrentPos And m_iEndFNumberTurret = iFNumberTurret Then
                    m_IsEnable = False
                    Exit Sub
                End If
                
                '���̎��_�Ń��j�b�g�����͓���ς�
                '�������ꏏ�̏ꍇ�͏�́u�ړ���Ɠ����ꍇ�v��FALSE�ɂȂ�B
                m_IsEnable = True
            
            Case "N-SIS3KAI"
                'F�l�A�������Ƃ��Ɉړ���Ɠ����ȏꍇ�̓K�`�����ŏȗ��Ƃ���
                If m_eEndDownPos = eCurrentPos And m_iEndFNumberTurret = iFNumberTurret Then
                    m_IsEnable = False
                    Exit Sub
                End If
                
                '���̎��_�ō��������͓���ς�
                'F�l���ꏏ�̏ꍇ�͏�́u�ړ���Ɠ����ꍇ�v��FALSE�ɂȂ�B
                m_IsEnable = True
                
            Case Else
                'Nikon���������� "N-SIS3", "N-SIS5", "N-SIS5KAI","N-SIS3KAI"�łȂ��ꍇ
                '������
                
        End Select
        
    Else
        'NIKON�����łȂ��ꍇ
        '������
    End If
    
    
End Sub

'CheckBeforeCondition�̌��ʂ�Ԃ��̂�
Private Function IEeeAuto_TestConditionChecker_CheckAfterCondition() As Boolean
    
    '���̃N���X��Before�����ōs���ׂ����ǂ������m�肷��B
    '�l��Ԃ��̂�
    IEeeAuto_TestConditionChecker_CheckAfterCondition = m_IsEnable

End Function

Private Sub IEeeAuto_TestConditionChecker_SetOperationMode(ByVal strMode As String)
    
    '���[�h�͂��̃N���X�ɂ͂Ȃ�
    
End Sub

'�����ݒ�̍ŏI�I�ȍs��������̃N���X�ɋ�����
'CheckBeforeCondition���ĂԑO�ɋ����Ă����邱��
Private Sub IEeeAuto_TestConditionChecker_SetEndPosition(ByVal strPos As String)
    
    'PALS��GlobalObject����s����́A�������AF�l���Ƃ�
    With OptCond.CondInfo(strPos)
        m_eEndDownPos = .OptModDownPosition
        m_iEndFNumberTurret = .OptModFnumberTurret
    End With
    
End Sub

'�������ʒu�擾�̃��b�p�֐�
'�̂��̂����u����Ƃ�悤�ɕς���ׂ�
Private Function GetCurrentDownPosition() As DownPosi

    If OptCond.IllumMaker = "NIKON" Then
        'Nikon�����̏ꍇ�A���u�ɂ�炸PALS����Ƃ�B
        '���኱�Ǝ�B�{���͑��u����Ƃ�ׂ��B
        GetCurrentDownPosition = PALS_IlluminatorMod.Flg_DownPosi
    End If
    
End Function

'�^�[���b�g�ʒu�擾�̃��b�p�֐�
'�̂��̂����u����Ƃ�悤�ɕς���ׂ�
Private Function GetCurrentFNumberTurret() As Integer

    If OptCond.IllumMaker = "NIKON" Then
        'Nikon�����̏ꍇ�A���u�ɂ�炸PALS����Ƃ�B
        '���኱�Ǝ�B�{���͑��u����Ƃ�ׂ��B
        GetCurrentFNumberTurret = PALS_IlluminatorMod.Flg_FUnit
    End If
    
End Function