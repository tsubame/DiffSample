Attribute VB_Name = "XEeeAuto_PowerCondition"
'�T�v:
'   PowerCondition�̊T�O��񋟂���
'
'�ړI:
'   PowerSequence�V�[�g�ƁAPowerCondition�V�[�g��ǂݍ���
'   PowerSupplyVoltage�̓d����Sequence�ɂ��������Ĉ������B
'
'�쐬��:
'   2011/12/05 Ver0.1 D.Maruyama
'   2011/12/07 Ver0.2 D.Maruyama
'       �EPowerCondition������PowerSequence���APowerSupplyVoltage����Ԃ���悤�ɂ���
'       �E�R�����g���ꕔ�C��
'   2012/04/06 Ver0.3 D.Maruyma  ApplyPowerCondition���������ƃV�[�P���X����2�����悤�ɕύX
'                                 ����ɂ��킹�āA�s�v�Ȋ֐��A�ϐ����폜����

Option Explicit

'�Œ�l
Private Const POWER_SEQUENCE_SHEET_COND_INDEX_CELL = "B4"
Private Const POWER_SEQUENCE_SHEET_NAME = "PowerSequence"

'���W���[�����ϐ�
Private m_colPowerSequence As Collection


'���e:
'   ���W���[����������
'
'���l:
'   ���W���[�����ϐ������������ɂ��āA�����V�[�g����ǂݒ����B
'
Public Sub InitializePowerCondition()

    '���������ɂ���
    Set m_colPowerSequence = Nothing

    '�R���N�V�����̐���
    Set m_colPowerSequence = New Collection
    
    '�V�[�g�̓ǂݍ���
    Call ReadPowerSequenceSheet
    
End Sub

Public Sub UninitializePowerCondition()

    '���������ɂ���
    Set m_colPowerSequence = Nothing

End Sub

'���e:
'   PowerCondition��ݒ�
'
'�p�����[�^:
'[strPowerConditionName]    IN   String:    �ݒ肷��PowerCondition��
'
'���l:
'   �w�肵��PowerCondition�����s����
'
Public Sub ApplyPowerCondition(ByVal strPowerConditionName As String, ByVal strSequenceName As String)
   
    Dim pPowerSequence As CPowerSequence
    
On Error GoTo SEQUENCE_NOT_FOUND
    Set pPowerSequence = m_colPowerSequence.Item(strSequenceName)
On Error GoTo 0

On Error GoTo ErrHandler
    Call pPowerSequence.Execute(strPowerConditionName)
On Error GoTo 0
    
    Exit Sub
        
ErrHandler:
    Call MsgBox("ApplyPowerCondition Fucntion Error Detect! @ " & strPowerConditionName & " @ " & strSequenceName)
    Exit Sub

SEQUENCE_NOT_FOUND:
    Call MsgBox("ApplyPowerCondition Fucntion Sequence not found! @ " & strPowerConditionName & " @ " & strSequenceName)
    Exit Sub

End Sub

'���e:
'   PowerCondition��ݒ�BFor APMU UnderShoot
'
'�p�����[�^:
'[strPowerConditionName]    IN   String:    �ݒ肷��PowerCondition��
'
'���l:
'   �w�肵��PowerCondition�����s����
'
Public Sub ApplyPowerConditionForUS(ByVal strPowerConditionName As String, ByVal strSequenceName As String)
   
    Dim pPowerSequence As CPowerSequence
    
On Error GoTo SEQUENCE_NOT_FOUND
    Set pPowerSequence = m_colPowerSequence.Item(strSequenceName)
On Error GoTo 0

On Error GoTo ErrHandler
    Call pPowerSequence.ExcecuteForUS(strPowerConditionName)
On Error GoTo 0
    
    Exit Sub
        
ErrHandler:
    Call MsgBox("ApplyPowerCondition Fucntion Error Detect! @ " & strPowerConditionName & " @ " & strSequenceName)
    Exit Sub

SEQUENCE_NOT_FOUND:
    Call MsgBox("ApplyPowerCondition Fucntion Sequence not found! @ " & strPowerConditionName & " @ " & strSequenceName)
    Exit Sub

End Sub


'���e:
'   PowerSequenceSheet����ǂݍ��݂��s��
'
'���l:
'
'
Private Sub ReadPowerSequenceSheet()

    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(POWER_SEQUENCE_SHEET_NAME)
    
    Dim NumOfItem As Long
    NumOfItem = sht.Range(POWER_SEQUENCE_SHEET_COND_INDEX_CELL).End(xlToRight).Column _
        - sht.Range(POWER_SEQUENCE_SHEET_COND_INDEX_CELL).Column - 1
    
    Dim i As Long
    Dim j As Long

    With sht.Range(POWER_SEQUENCE_SHEET_COND_INDEX_CELL)
        For i = 0 To NumOfItem
            Dim tempPowerSeq As CPowerSequence
            Set tempPowerSeq = New CPowerSequence
            j = 0
            Call tempPowerSeq.InitializeThisClass(.offset(j, i + 1))
            While .offset(j + 1, i + 1) <> ""
                Dim tempPowerSequenceItem As IPowerSequenceItem
                If IsNumeric(.offset(j + 1, i + 1)) Then
                    Set tempPowerSequenceItem = New CPowerSquenceWait
                    Call tempPowerSequenceItem.SetParam(.offset(j + 1, i + 1))
                Else
                    Set tempPowerSequenceItem = New CPowerSequencePin
                    Call tempPowerSequenceItem.SetParam(.offset(j + 1, i + 1).Text)
                End If
                Call tempPowerSeq.Add(tempPowerSequenceItem)
                j = j + 1
            Wend
            Call m_colPowerSequence.Add(tempPowerSeq, tempPowerSeq.Name)
        Next i
    End With

End Sub

'���e:SetVoltage(US�΍���)
Public Sub PowerDown4ApmuUnderShoot() '2012/11/16 175Debug Arikawa

        '�p�^����~
    Call StopPattern 'EeeJob�֐�
   
    Dim pPowerSequence As CPowerSequence
    If getPowerDownSequence(pPowerSequence) = False Then Exit Sub
    
    On Error GoTo ERROR_DETECTION1
    Call pPowerSequence.ExcecuteForUS("ZERO")
    Exit Sub
ERROR_DETECTION1:
    Call pPowerSequence.ExcecuteForUS("ZERO_V")
End Sub

'���e
'   �d��Off����Power Sequence�����擾����B
'   Gang�̗L���ɂ���āA�\��"PowerSequence"�V�[�g�ɐ��������A���W�X�^�ʐMI/F
'   �Ɉˑ����Ȃ��d��Off�V�[�P���X�����قȂ邽�߁A
'       1. APMU Gang�̉\���̂���[�q���܂܂��ꍇ
'       2. APMU Gang�̉\���̂���[�q���܂܂�Ȃ��ꍇ
'   �̏��ɁA�V�[�P���X�����擾����B
'Description
'   To return power sequence object for power down.
'   In order to support both pin assigns with APMU Gang and without APMU Gang,
'   it is done in the following order.
'       1. Get power sequence with name "ANY_SeqOff_GangOff"
Private Function getPowerDownSequence(ByRef pPowerSequence As CPowerSequence) As Boolean
    Const POWER_DOWN_SEQUENCE_GANG As String = "ANY_SeqOff_GangOff"
    Const POWER_DOWN_SEQUENCE_NOGANG As String = "ANY_SeqOff"

    '1. APMU Gang OFF���̃V�[�P���X�BGang�̎�s����OFF����΁A�S��OFF�ɂȂ�͂��B
    On Error GoTo ErrorGangNotFound
    Set pPowerSequence = m_colPowerSequence.Item(POWER_DOWN_SEQUENCE_GANG)
    getPowerDownSequence = True
    Exit Function
    
ErrorGangNotFound:
    '2. APMU Gang�̉\���̂���[�q���܂܂�Ȃ��ꍇ��OFF�V�[�P���X�B
    On Error GoTo ErrorSeqNotFound
    Set pPowerSequence = m_colPowerSequence.Item(POWER_DOWN_SEQUENCE_NOGANG)
    getPowerDownSequence = True
    Exit Function

ErrorSeqNotFound:
    Err.Raise 9999, "getPowerDownSequence", "Power down sequence not found [" & GetInstanceName & "] !"
    Call DisableAllTest
End Function

'���e:SetVoltage(US�΍���)
Public Sub Set_Voltage(ByVal strPowerConditionName As String, ByVal strSequenceName As String) '2012/11/16 175Debug Arikawa

        '�p�^����~
    Call StopPattern 'EeeJob�֐�
   
    Dim pPowerSequence As CPowerSequence
    Set pPowerSequence = m_colPowerSequence.Item(strSequenceName)

    Call pPowerSequence.ExcecuteForUS(strPowerConditionName)

End Sub


