Attribute VB_Name = "DCtestSampleMod"
'�T�v:
'   �e�X�g�C���X�^���X���쐬����ۂ̃T���v���v���V�[�W���Q
'
'�ړI:
'   DC�e�X�g�V�i���I�V�[�g�𗘗p����ۂ̐��`�C���X�^���X�����[�U�Ɍ��J����
'
'
'�쐬��:
'   SLSI��J
'
Option Explicit

Private Function MultiDcTest_f() As Long
'���e:
'   DC�e�X�g�C���X�^���X���쐬����ۂ̃T���v���v���V�[�W��
'   ���[�U�[�͂��̃T���v������C�ӂ�DC�e�X�g�C���X�^���X���쐬�E�������邱�Ƃ��\
'   �܂��V�i���I���s�O��ɔC�ӂ̏����̒ǉ����\
'
'���l:
'   �C���X�^���X�V�[�g�̃C���X�^���X���ƁA�e�X�g�V�i���I�V�[�g��
'   �J�e�S�����i�X�y�[�X�͋�����j�͐���������Ă����K�v������

    Call SiteCheck

    '@@@ ����V�i���I������ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    TheDcTest.SetScenario GetInstanceName
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    '##########  TEST CONDITION SET UP ####################
'    Call XXXXX_Setup
'    Call SET_XXXXX_CONDITION
'    Call SetVoltage(XXXXX)
'    Call SetVRL(XXXXX)
'    Call PatSet(XXXXX)
'    TheHdw.Wait XXXXX * mS
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    '@@@ ����V�i���I���s @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    TheDcTest.Execute
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

End Function


Private Function ReturnResult_ShirotenMargin_f() As Long
'���e:
'   �C���X�^���X�����L�[�Ƃ��ăe�X�g���ʃR���N�V��������v�f�����o��Test�֐��ɓn��
'   ��̊֐��Ƃ̈Ⴂ�́A�e�X�g���ʂ��Ȃ��ꍇ�́A�G���[�ł͂Ȃ��A0��Ԃ��B
'
'���l:
'   �C���X�^���X�V�[�g�̃C���X�^���X���ƁA�e�X�g�V�i���I�V�[�g��
'   �e�X�g���x���͐���������Ă����K�v������
'
    Dim resultTest() As Double
    
    Call SiteCheck

    If TheResult.IsExist(GetInstansNameAsUCase) Then
        Call TheResult.GetResult(GetInstansNameAsUCase, resultTest)
    Else
        ReDim resultTest(GetSiteCount)
        Dim SiteIndex As Long
        For SiteIndex = 0 To GetSiteCount
            resultTest(SiteIndex) = 0
        Next SiteIndex
    End If
    Call test(resultTest)
    
End Function
