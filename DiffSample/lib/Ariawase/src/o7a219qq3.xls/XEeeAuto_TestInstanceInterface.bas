Attribute VB_Name = "XEeeAuto_TestInstanceInterface"
Option Explicit

'�T�v:
'   TestInstance�����Arg�擾�Ɋւ���֐��Q
'
'�ړI:
'
'
'�쐬��:
'   2011/12/21 Ver0.1 D.Maruyama
'   2012/04/09 Ver0.2 D.Maruyama�@�@dcsetup�֐��̈����ύX�ɔ����A�uEEE_AUTO_DCSETUP_PARAM�v�̒�`���C��
'   2012/04/09 Ver0.2 D.Maruyama�@�@dcsetup�֐��̈����ύX�ɔ����A�uEEE_AUTO_DCSETUP_PARAM�v�̒�`���C��


Private Const EEE_AUTO_TEST_INSTANCE_ARG_START As Long = 20

Public Const EEE_AUTO_VARIABLE_PARAM As Long = -1

Public Const EEE_AUTO_DCSETUP_PARAM As Long = 3
Public Const EEE_AUTO_ENDOFTEST_PARAM As Long = 1

'���e:
'   TestInstance�����Arg�̓ǂݎ������b�v����B
'   ���r���[�ȂƂ����Arg�̊J�n�ʒu�Ƃ������߁A���b�v�֐��ŉǐ����悭����B
'
'
'�p�����[�^:
'[arystrParam]         Out  ���ʔz��
'[lNumOfParam]         In�@ Arg�̐�
'
'���ӎ���:
'   arystrParam�͊m�ۂ��ĂȂ����I�z���n������
'
Public Function EeeAutoGetArgument(ByRef arystrParam() As String, ByVal lNumOfParam As Long) As Boolean

    EeeAutoGetArgument = False

    Dim ArgArr() As String
    Dim Argnum As Long
    
    Call TheExec.DataManager.GetArgumentList(ArgArr, Argnum)

    '���҂���p�����[�^���قȂ�ꍇ�̓G���[�Ƃ���B�ψ����̏ꍇ�͖���
    If lNumOfParam <> (Argnum - EEE_AUTO_TEST_INSTANCE_ARG_START) And _
            lNumOfParam <> EEE_AUTO_VARIABLE_PARAM Then
        EeeAutoGetArgument = False
        Exit Function
    End If
    
    '�K�v�Ȑ����������̔z����m�ۂ���
    Dim lUsedNum As Long
    lUsedNum = Argnum - EEE_AUTO_TEST_INSTANCE_ARG_START
    ReDim arystrParam(lUsedNum)
    
    '�R�s�[����
    Dim i As Long
    For i = 0 To lUsedNum - 1
        arystrParam(i) = ArgArr(EEE_AUTO_TEST_INSTANCE_ARG_START + i)
    Next i
    
    EeeAutoGetArgument = True
    
End Function


