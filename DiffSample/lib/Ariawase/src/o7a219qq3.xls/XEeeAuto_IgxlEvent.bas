Attribute VB_Name = "XEeeAuto_IgxlEvent"
Option Explicit
'   2012/12/21  H.Arikawa
'               IGXL_OnProgramLoaded��ǉ��B
'   2013/10/22  H.Arikawa
'               �I�t���C�����[�h�œ��삷��ۂ�csv�ǂݍ��݂ɂ����Ȃ��悤�ɕύX�B


'Validation�̏I���^�C�~���O�Ɏ��s�����C���^�[�|�[�Y�t�@���N�V����

Public Sub IGXL_OnProgramValidated()

    If APMU_CheckFailSafe_f = False Then
        ThisWorkbook.Saved = True
        Application.Quit
    End If
    
    '�e�X�g�J�n���̃C���^�[�|�[�Y�t�@���N�V������EeeJob���̏����������s�����
    'OffsetSheet����̏ꍇ�AEeeJob�̏������ŃG���[�ɂȂ邽�߁A�o���f�[�V�����̃^�C�~���O�Ŏ��s����
    If TheExec.TesterMode = testModeOffline Then Flg_Simulator = 1
    Call JobEnvInit
    If Not TheExec.TesterMode = testModeOffline Then
        Call GetCsvFileName
        Call ReadOffsetFile
        Call WriteOffsetManager
        Call ReadOptFile '�㏑�����Ă��܂��̂�
    End If
        
End Sub

Public Sub IGXL_OnProgramLoaded()
'PatGrps�V�[�g��"TSBName"�̃Z���̐F�Â����s��
    Call PatGrpsColorMake
    
End Sub

