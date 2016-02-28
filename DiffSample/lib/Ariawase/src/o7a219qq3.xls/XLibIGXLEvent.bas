Attribute VB_Name = "XLibIGXLEvent"
'�T�v:
'   IG-XL�C�x���g�ŌĂяo�����}�N���֐�
'
'�ړI:
'   �e�^�C�~���O�ł̃��C�u���������Ɏg�p
'   ���[�U�[�ւ͕ʃ}�N���֐���񋟂���
'
'   Revision History:
'   Data        Description
'   2008/05/20�@Eee-JOB V1.21�Ń����[�X [XLibIGXLEvent.bas]
'   2009/04/07  Eee-JOB V2.00�Ń����[�X [XLibIGXLEvent.bas]
'               ���d�l�ύX
'               �v���p�e�B�E���\�b�h���̃K�C�h���C���{�s�ɔ������̕ύX�ɑΉ�
'               ���@�\�ǉ�
'               OnProgramLoaded�C�x���g��EeeNavigation���j���[�R���X�g���N�^�p�֐��ǉ�
'   2009/04/21�@���d�l�ύX
'               �A�h�C���쐬�K�C�h���C���ɔ����A�h�C���v���W�F�N�g���ύX�ɑΉ�
'   2009/06/15  EeeNavigation Ver1.01�A�h�C���W�J�ɔ����ύX
'             �@���d�l�ύX
'               �i�r�Q�[�V�������j���[�̃R���X�g���N�^�p�}�N���֐�����ύX
'
'�쐬��:
'   0145206097
'

Option Explicit
'Tool�Ή���ɃR�����g�O���Ď��������ɂ���B�@2013/03/07 H.Arikawa
#Const CUB_UB_USE = 0    'CUB UB�̐ݒ�          0�F���g�p�A0�ȊO�F�g�p

'TESTER��Initial���̃C�x���g�AStartDatatool���s���ɃR�R�����s�����B
Public Function OnTesterInitialized() As Long
    
    #If CUB_UB_USE <> 0 Then
    Call XLibJob.InitCub
    #End If

    InitControlShtReader

    '### ���[�U�[�֒񋟂���}�N�� #########################
    On Error Resume Next
    Application.Run ("IGXL_OnTesterInitialized")
    On Error GoTo 0

End Function

'�v���O�������[�h��̃C�x���g
Public Function OnProgramLoaded() As Long

    If TheExec.RunMode = runModeProduction Then
        CheckExaminationMode
    End If

    '### EeeNavi�Z�b�g�A�b�v���j���[�R���X�g���N�^ ########
    On Error Resume Next
    Application.Run ("XLibEeeNaviConstructor.CreateEeeNaviSetUpMenu")
    On Error GoTo 0

    '### TOPT�t���[�����[�N�p #############################
    Call XLibToptFrameWorkUtility.ResetEeeJobObjects

    '### ���[�U�[�֒񋟂���}�N�� #########################
    On Error Resume Next
    Application.Run ("IGXL_OnProgramLoaded")
    On Error GoTo 0

End Function

'�o���f�[�V�����J�n���̃C�x���g
Public Function OnValidationStart() As Long
    
    '### OnValidationStart�Ŏ��s����֐��Q ################
    XLibJobUtility.RunAtValidationStart
    
    '### ���[�U�[�֒񋟂���}�N�� #########################
    On Error Resume Next
    Application.Run ("IGXL_OnValidationStart")
    On Error GoTo 0

End Function

'�o���f�[�V������̃C�x���g�AValidate Job���s��ɃR�R�����s�����B
Public Function OnProgramValidated() As Long
        
    InitControlShtReader
    ValidateDCTestSenario

    '### TOPT�t���[�����[�N�p #############################
    Call XLibToptFrameWorkUtility.ResetEeeJobSheetObjects
    Call XLibToptFrameWorkUtility.RunAtValidated

    '### ���[�U�[�֒񋟂���}�N�� #########################
    On Error Resume Next
    Application.Run ("IGXL_OnProgramValidated")
    On Error GoTo 0

End Function

'TDR�L�����u���[�V������̃C�x���g
Public Function OnTDRCalibrated() As Long

    '### ���[�U�[�֒񋟂���}�N�� #########################
    On Error Resume Next
    Application.Run ("IGXL_OnTDRCalibrated")
    On Error GoTo 0

End Function

'JOB�v���O�������s�J�n����̃C�x���g
Public Function OnProgramStarted() As Long

    '### TOPT�t���[�����[�N�p #############################
    Call XLibToptFrameWorkUtility.RunAtJobStart

    '### ���[�U�[�֒񋟂���}�N�� #########################
    On Error Resume Next
    Application.Run ("IGXL_OnProgramStarted")
    On Error GoTo 0

End Function

'JOB�v���O�������s�I������̃C�x���g
Public Function OnProgramEnded() As Long

    '### TOPT�t���[�����[�N�p #############################
    Call XLibToptFrameWorkUtility.RunAtJobEnd

    '### ���[�U�[�֒񋟂���}�N�� #########################
    On Error Resume Next
    Application.Run ("IGXL_OnProgramEnded")
    On Error GoTo 0

End Function
