Attribute VB_Name = "XEeeAuto_Result"
'�T�v:
'  �@���ʂ������n�֐�
'
'�ړI:
'
'
'�쐬��:
'   2012/01/27 Ver0.1 D.Maruyama
'   2013/03/15 Ver0.2 H.Arikawa �s�v�֐��폜

Option Explicit

Private Function ReturnResult_f() As Long
'���e:
'   �C���X�^���X�����L�[�Ƃ��ăe�X�g���ʃR���N�V��������v�f�����o��Test�֐��ɓn��
'
'���l:
'   �C���X�^���X�V�[�g�̃C���X�^���X���ƁA�e�X�g�V�i���I�V�[�g��
'   �e�X�g���x���͐���������Ă����K�v������
'

    Call SiteCheck

    On Error GoTo DATA_ERR
    Dim resultTest() As Double
    TheResult.GetResult GetInstansNameAsUCase, resultTest
    Call test(resultTest)

    Exit Function
DATA_ERR:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    ReDim resultTest(GetSiteCount)
    Dim SiteIndex As Long
    For SiteIndex = 0 To GetSiteCount
        resultTest(SiteIndex) = 0
    Next SiteIndex
    Call test(resultTest)
    Break
End Function


Private Function ReturnResultEx_f() As Long
'���e:
'   �C���X�^���X�����L�[�Ƃ��ăI�t�Z�b�g�V�[�g����f�[�^���擾��
'   �e�X�g���ʂɃI�[�o�[���C�g���Ă���Test�֐��ɓn��
'
'���l:
'   �C���X�^���X�V�[�g�̃C���X�^���X���ƁA�e�X�g�V�i���I�V�[�g��
'   �e�X�g���x���͐���������Ă����K�v������
'

    Call SiteCheck
    On Error GoTo ErrHandler
    If TheOffsetResult Is Nothing Then
        Err.Raise 9999, "ReturnResultEx_f", "Can Not Implement This Instance In Function [" & GetInstanceName & "] !"
    End If
    '@@@ ���茋�ʃR���o�[�g @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    TheOffsetResult.Calculate GetInstansNameAsUCase, TheResult
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    Dim resultTest() As Double
    TheResult.GetResult GetInstansNameAsUCase, resultTest
    Call test(resultTest)
    Exit Function
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    ReDim resultTest(GetSiteCount)
    Dim SiteIndex As Long
    For SiteIndex = 0 To GetSiteCount
        resultTest(SiteIndex) = 0
    Next SiteIndex
    Call test(resultTest)
End Function
