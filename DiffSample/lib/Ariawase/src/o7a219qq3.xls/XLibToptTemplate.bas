Attribute VB_Name = "XLibToptTemplate"
'�T�v:
'   ToptFrameWork�e���v���[�g�����p����TheImageTest�I�u�W�F�N�g�̃��b�p�[�֐��Q
'
'   Revision History:
'       Data        Description
'       2010/04/28  �B���e�X�g�����s����@�\����������
'       2010/05/12  �v���O�����R�[�h�𐮗�����
'       2010/05/31  Error������ύX����
'       2010/06/11  �v���O�����R�[�h�𐮗�����
'
'�쐬��:
'   0145184346
'

Option Explicit

Public Function SetScenario(argc As Long, argv() As String) As Integer
'���e:
'   �B���e�X�g�����s���邽�߂̏���������
'
'�p�����[�^:
'   [argc]    In  �w�肳�ꂽ�����̐�
'   [argv()]  In  �w�肳�ꂽ�����̔z��f�[�^
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    On Error GoTo ErrHandler


    '#####  �B���e�X�g�����s���邽�߂̏���������  #####
    SetScenario = TheImageTest.SetScenario


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    SetScenario = TL_ERROR
    Exit Function


End Function

Public Function Execute(argc As Long, argv() As String) As Integer
'���e:
'   �B���e�X�g�����s����
'
'�p�����[�^:
'   [argc]    In  �w�肳�ꂽ�����̐�
'   [argv()]  In  �w�肳�ꂽ�����̔z��f�[�^
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    On Error GoTo ErrHandler


    '#####  �B���e�X�g�����s����  #####
    Execute = TheImageTest.Execute


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    Execute = TL_ERROR
    Exit Function


End Function
