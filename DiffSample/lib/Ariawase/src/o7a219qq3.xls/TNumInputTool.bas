Attribute VB_Name = "TNumInputTool"
'�T�v:
'   Flow Table���[�N�V�[�g�ցA�e�X�g�ԍ����[���ɏ]�����e�X�g�ԍ�����͂���
'
'�ړI:
'   �e�X�g�ԍ����[���̏���
'   �e�X�g�ԍ����͋֎~���ւ̔ԍ����̓g���u���h�~
'
'�쐬��:
'   SLSI ����
'
'�J������:
'   2009-05-25 ���쏉�Ŋ����B
'
'   2009-05-27 �e�X�g�ԍ��̍폜�������A�蓮���̑���ɍ��킹�� ���͒l��""����Empty�ɕύX�B
'              �e�X�g�ԍ����͋֎~���ɒl�������Ă����Ƃ��̃��b�Z�[�WWindow��ύX�B
'
'   2009-06-15 ���s���Ƀe�X�g�ԍ����[����Version���w�肷�邱�Ƃŕ����̃e�X�g�ԍ����[���ɑΉ��B
'              2008-12-17�ɍ��ӂ��ꂽ�e�X�g�ԍ����[����Version2�������B
'              �e�X�g�ԍ����[��Version1��dc���x����p�~(Job�`�F�b�NTool�̌��o�d�l�ɍ��킹��)�B
'              TOOL��Version�ԍ����J�@�\�ǉ�
'
'   2009-06-18 �����Ō��J�ɂނ��ă��W���[�����ύX(ZZZ_InputTestNumTool-->TNumInputTool)
'   2012-12-26 H.Arikawa ��������TNumber�ɍ��킹��VER3��ǉ�

Option Explicit

'Flow Table�V�[�g�Ɋւ����`�Ȃ�
Private Const LABEL_INDEX = "B5"
Private Const OPCODE_INDEX = "G5"
Private Const PARAMETER_INDEX = "H5"
Private Const RESULT_INDEX = "O5"
Private Const LAST_OPCODE_VALUE = "set-device"
Private Const FLOW_TABLE_SHT_NAME = "Flow Table"
Private Const TEST_CATEGORY_HEADER_NAME = "SEQ"
Private Const PROHIBITION_PLACE_MSG = "It is a test number input prohibition place"

'�e�X�g�ԍ����[�������������֐����̒�`(���V���[���ǉ����ɂ̓R�R�������e)
Private Const TNUM_RULE_SUPPLIER_V1 = "RetStartTestNumber_V1"
Private Const TNUM_RULE_SUPPLIER_V2 = "RetStartTestNumber_V2"
Private Const TNUM_RULE_SUPPLIER_V3 = "RetStartTestNumber_V3"

Public Enum TNumRuleVer
    VER_1 = 1#
    VER_2 = 2#
    VER_3 = 3#
End Enum

'Parameter �Z����Index�p
Private m_ParameterRng As Range

'TOOL��Version���J�p
Private Const TOOL_VERSION = "1.00"


Public Sub InputTestNumber(ByVal pRuleVerNo As TNumRuleVer, _
Optional ByVal pTgtFlowTable As Worksheet = Nothing)
'���e:
'   Flow Table�V�[�g�Ƀe�X�g�ԍ�����͂���
'
'�p�����[�^:
'   [pRuleVerNo]       In   �e�X�g�ԍ����[����Version�ԍ�
'   [pTgtFlowTable]    In   �ԍ�����͂���FlowTabel�V�[�gObject�i�I�v�V�����j
'
'�߂�l:
'
'���ӎ���:
'
    Dim Label As String
    Dim OpcodeRng As Range
    Dim LabelRng As Range
    Dim ResultRng As Range
    Dim TestNumber As Long
    Dim answer As Long
    Dim GetTestNumber As String
        
    '�g�p����e�X�g�ԍ����[���̑I������(���V���[���ǉ����ɂ̓R�R�������e)
    Select Case pRuleVerNo
        Case VER_1
            GetTestNumber = TNUM_RULE_SUPPLIER_V1
        Case VER_2
            GetTestNumber = TNUM_RULE_SUPPLIER_V2
        Case VER_3
            GetTestNumber = TNUM_RULE_SUPPLIER_V3
        Case Else
            Call Err.Raise(9999#, "InputTestNumber", "Rule Version = " & pRuleVerNo & " is unknown test number rule !")
            Exit Sub
    End Select
            
    '���[�N�V�[�g�̎w��ȗ����́AFlow Table������Ƀ^�[�Q�b�g�Ƃ���
    '(�{���̓V�[�g�������Ƃ��̃G���[����������ق����悢�j
    If pTgtFlowTable Is Nothing Then
        Set pTgtFlowTable = Worksheets(FLOW_TABLE_SHT_NAME)
    End If
                
    '����̃C���f�b�N�X�ݒ�
    Set LabelRng = pTgtFlowTable.Range(LABEL_INDEX)
    Set OpcodeRng = pTgtFlowTable.Range(OPCODE_INDEX)
    Set m_ParameterRng = pTgtFlowTable.Range(PARAMETER_INDEX)
    Set ResultRng = pTgtFlowTable.Range(RESULT_INDEX)
        
    '��ԍŏ��̃��x��
    Label = LabelRng.Value
    TestNumber = Application.Run(GetTestNumber, Label)
    
    While OpcodeRng <> LAST_OPCODE_VALUE
        '���x���̒l�m�F(�V�������x�����o�ꂵ����)
        If (LabelRng.Value <> "") And (Label <> LabelRng.Value) Then
            '�V�������x���p�̃e�X�g�ԍ��擾
            Label = LabelRng.Value
            TestNumber = Application.Run(GetTestNumber, Label)
        End If
        
        'Enable Word���󗓁AResult���󗓁AParameter�̒l��SEQ�̂Ƃ����
        ' �e�X�g�ԍ��͓��͂���Ȃ���OK�H
        If (LabelRng.offset(0#, 1#).Value <> "") And (ResultRng.Value <> "") _
        And (m_ParameterRng.Value <> TEST_CATEGORY_HEADER_NAME) Then
                
            '�p�����[�^���󗓂łȂ��A����TName���󗓂łȂ�
            If (m_ParameterRng.Value <> "") And (TNameRng.Value <> "") Then
                'Test�ԍ����Z���ɓ���
                TNumberRng.Value = TestNumber
                '�e�X�g�ԍ���1�C���N�������g����
                TestNumber = TestNumber + 1#
            Else
                'Test�ԍ�����͂��Ȃ��i�e�X�g�ԍ����͋֎~�̕����Ȃ̂ŋ󗓂ɂ���j�������̏����͕s�v����
                If TNumberRng.Value <> "" Then
                    answer = MsgBox(MakeEraseConfirmMsg, vbYesNo + vbExclamation, PROHIBITION_PLACE_MSG)
                    If answer = vbYes Then
                       TNumberRng.Value = Empty
                    End If
                End If
            End If
        Else
            'Test�ԍ�����͂��Ȃ��i�e�X�g�ԍ����͋֎~�̕����Ȃ̂ŋ󗓂ɂ���j
                If TNumberRng.Value <> "" Then
                    answer = MsgBox(MakeEraseConfirmMsg, vbYesNo + vbExclamation, PROHIBITION_PLACE_MSG)
                    If answer = vbYes Then
                        TNumberRng.Value = Empty
                    End If
                End If
        End If
    
        'Index���ЂƂ��ɐi�߂�
        Set LabelRng = LabelRng.offset(1#, 0#)
        Set m_ParameterRng = m_ParameterRng.offset(1#, 0#)
        Set OpcodeRng = OpcodeRng.offset(1#, 0#)
        Set ResultRng = ResultRng.offset(1#, 0#)
    
    Wend
    
    Call MsgBox("Done", vbInformation, "InputTestNumber")

End Sub

Public Function TNumInputToolVer() As String
'���e:
'   TNumInputTool��Version�ԍ���Ԃ�
'
'�p�����[�^:
'
'�߂�l:
'   TNumInputTool��Version�ԍ�
'
'���ӎ���:
'
    TNumInputToolVer = TOOL_VERSION

End Function

'--------------------------------------------------------------------------------
'�ȉ� Private Function

'#Pass
'TName Range�̎擾
Private Function TNameRng() As Range
    Set TNameRng = m_ParameterRng.offset(0#, 1#)
End Function

'#Pass
'TNum Range�̎擾
Private Function TNumberRng() As Range
    Set TNumberRng = m_ParameterRng.offset(0#, 2#)
End Function

'#Pass
'�e�X�g�ԍ��̏����m�F�p�̃��b�Z�[�W�쐬�p
Private Function MakeEraseConfirmMsg() As String
    MakeEraseConfirmMsg = "Address= " & m_ParameterRng.Address & vbCrLf & _
    "Parameter= " & m_ParameterRng.Value & vbCrLf & _
    "Test number= " & m_ParameterRng.offset(0#, 2#).Value & vbCrLf & _
    "Do you erase a test number?"
End Function


'--------------------------------------------------------------------------------
'�ȉ� �e�X�g�ԍ����[���̎���(���V���[���ǉ����ɂ̓R�R�������e)

'#Pass
'�e�X�g�ԍ����[�� Version1
'2009/05/22���_�ł�Legacy���[���̎���(�\�[�X=�ؑ�����̏����)
Private Function RetStartTestNumber_V1(ByVal pLabel As String) As Long
    
    Select Case pLabel
        Case "dcpar"
            RetStartTestNumber_V1 = 2#
            Exit Function
        Case "image"
            RetStartTestNumber_V1 = 1002#
            Exit Function
        Case "color"
            RetStartTestNumber_V1 = 2002#
            Exit Function
        Case "flmura"
            RetStartTestNumber_V1 = 3002#
            Exit Function
        Case "grade"
            RetStartTestNumber_V1 = 4002#
            Exit Function
        Case "shiroten"
            RetStartTestNumber_V1 = 5002#
            Exit Function
        Case "nashiji"
            RetStartTestNumber_V1 = 6002#
            Exit Function
        Case "margin"
            RetStartTestNumber_V1 = 7002#
            Exit Function
        Case Else
            Call Err.Raise(9999#, "RetStartTestNumber_V1", pLabel & " is UnKnown Label !")
    End Select

End Function

'#Pass
'�e�X�g�ԍ����[�� Version2
'2008/12/17�Ɍ��؁A�F�{�A����ō��ӂ��ꂽ�V���[���̎���(�\�[�X=�ēc����̏����)
Private Function RetStartTestNumber_V2(ByVal pLabel As String) As Long
    
    Select Case pLabel
        Case "dcpar"
            RetStartTestNumber_V2 = 2#
            Exit Function
        Case "image"
            RetStartTestNumber_V2 = 1002#
            Exit Function
        Case "color"
            RetStartTestNumber_V2 = 3002#
            Exit Function
        Case "grade"
            RetStartTestNumber_V2 = 5002#
            Exit Function
        Case "shiroten"
            RetStartTestNumber_V2 = 6002#
            Exit Function
        Case "margin"
            RetStartTestNumber_V2 = 8002#
            Exit Function
        Case Else
            Call Err.Raise(9999#, "RetStartTestNumber_V2", pLabel & " is UnKnown Label !")
    End Select

End Function

'#Pass
'�e�X�g�ԍ����[�� Version3
'2012/11/12 �������ɑΉ��������[������
Private Function RetStartTestNumber_V3(ByVal pLabel As String) As Long
    
    Select Case pLabel
        Case "dcpar"
            RetStartTestNumber_V3 = 2#
            Exit Function
        Case "image"
            RetStartTestNumber_V3 = 1002#
            Exit Function
        Case "grade"
            RetStartTestNumber_V3 = 5002#
            Exit Function
        Case "shiroten"
            RetStartTestNumber_V3 = 6002#
            Exit Function
        Case "margin"
            RetStartTestNumber_V3 = 8002#
            Exit Function
        Case Else
            Call Err.Raise(9999#, "RetStartTestNumber_V3", pLabel & " is UnKnown Label !")
    End Select

End Function

