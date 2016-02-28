Attribute VB_Name = "XLibDcTestTemplate"
'�T�v:
'   DCTestScenario�e���v���[�g�����p����TheDcTest�I�u�W�F�N�g�̃��b�p�[�֐��Q
'
'�ړI:
'   DCTestScenario�e���v���[�g���A�h�C���Œ񋟂��邽�߂̎�i�Ƃ��č쐬
'   �e���v���[�g����TheDcTest�I�u�W�F�N�g�̎Q�Ƃ��o���Ȃ����߁A
'   ���b�p�[�֐���p�ӂ��e���v���[�g���炱�̊֐����Ăяo��
'
'   Revision History:
'   Data        Description
'   2008/09/25�@�]���Ń����[�X
'   2009/04/07�@V2.0���C�u�����Z�b�g�p�Ƀ����[�X
'               ���d�l�ύX
'               �@Eee-JOB���C�u�����Z�b�g�̃v���p�e�B����\�b�h���̃K�C�h���C���{�s�ɔ����֐����̕ύX
'               �A�@�̗��R��TheDcTest�I�u�W�F�N�g�̊e���\�b�h�Ăяo����ύX
'               �B�I�u�W�F�N�g���̕ύX
'
'�쐬��:
'   0145206097
'
'
'   Ver1.1 2013/02/01 H.Arikawa Ex�p��Execute��dumpPPMUreg��ǉ��B

Option Explicit

Public Function SetScenario(argc As Long, argv() As String) As Long
    On Error GoTo ErrHandler
    If argc > 1 Then
        Err.Raise 9999, "XLibDcTestTemplate.SetScenario", "Two Or More Arguments Are Not Supported !"
        GoTo ErrHandler
    End If
    TheDcTest.SetScenario argv(0)
    SetScenario = TL_SUCCESS
    Exit Function
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    SetScenario = TL_ERROR
End Function

Public Function Execute(argc As Long, argv() As String) As Long
    On Error GoTo ErrHandler
    TheDcTest.Execute
    Execute = TL_SUCCESS
    Exit Function
ErrHandler:
    dumpPPMUreg
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Execute = TL_ERROR
End Function
