VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVISVISource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   IP750 �d��Board I/F
'
'�ړI:
'   �e��d���{�[�h�̐�����@�̓���
'
'�쐬��:
'   SLSI����
'
'
'Code Checked
'Comment Checked
'

Option Explicit

Private Const ALL_SITE = -1

Public Property Get BoardName() As String
'���e:
'   IP750�{�[�h�̖��O
'
'�p�����[�^:
'
'�߂�l:
'   IP750�{�[�h��
'
'���ӎ���:
'   Read Only
'
End Property

Public Property Let PinList(ByVal PinList As String)
'���e:
'   ����Ώۃs�����X�g
'
'�p�����[�^:
'    [PinList]   In  ����Ώۃs�����X�g�B
'
'�߂�l:
'
'���ӎ���:
'
End Property

Public Property Get PinList() As String
'���e:
'   ����Ώۃs�����X�g
'
'�p�����[�^:
'
'�߂�l:
'   ����Ώۃs�����X�g
'
'���ӎ���:
'
End Property

Public Sub SetFVMI(ByVal ForceV As Variant, _
    ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True)
'���e:
'   PinList�ɐݒ肳�ꂽ�s����d�������Ԃɐݒ肷��
'
'�p�����[�^:
'    [ForceV]      In   ����d���B�z��w��\�B
'    [ClampI]      In   �N�����v�d���l�B
'    [Site]        In   �T�C�g�w��B�I�v�V����(Default -1)
'    [ConnectOn]   In   �f�o�C�X�ɐڑ����邩�ǂ����B�I�v�V����(Default True)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s����d�������Ԃɂ���
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��ForceV�ň���d�����w��BForceV�͐��l or �T�C�g�����̔z��B
'    �����l�̏ꍇ�S�T�C�g�����l��ݒ衔z��̏ꍇ�e�T�C�g�ɑΉ�����z��v�f�̒l��ݒ�
'    ��ClampI�ŃN�����v�d����ݒ�
'    ��Site�őΏۂ̃T�C�g���w��B-1(�f�t�H���g)������ƁA�S�T�C�g���ΏہB
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'    ���d����������W��ForceV��d�����背���W��ClampI���画�肵�Đݒ肷��
'    ��ConnectOn��True(�f�t�H���g)�ɂ���ƁA�ݒ�ƃR�l�N�g����x�ɂ���B
'    ��ConnectOn��False�ɂ���Ƥ�ݒ�݂̂ŃR�l�N�g�͍s�Ȃ�Ȃ�� (���ɃR�l�N�g����Ă���ꍇ�͂��̂܂�)
'
End Sub
'V21-Release
Public Sub SetFVMIMulti(ByVal ForceV As Variant, _
    ByVal ClampI As Double, _
    Optional ByVal ConnectOn As Boolean = True)
'���e:
'   PinList�ɐݒ肳�ꂽ�s����d�������Ԃɐݒ肷��(Site����)
'
'�p�����[�^:
'    [ForceV]      In   ����d���B�z��w��\�B
'    [ClampI]      In   �N�����v�d���l�B
'    [ConnectOn]   In   �f�o�C�X�ɐڑ����邩�ǂ����B�I�v�V����(Default True)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s����d�������Ԃɂ���
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��ForceV�ň���d�����w��BForceV�͐��l or �T�C�g�����̔z��B
'    �����l�̏ꍇ�S�T�C�g�����l��ݒ衔z��̏ꍇ�e�T�C�g�ɑΉ�����z��v�f�̒l��ݒ�
'    ��ClampI�ŃN�����v�d����ݒ�
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'    ���d����������W��ForceV��d�����背���W��ClampI���画�肵�Đݒ肷��
'    ��ConnectOn��True(�f�t�H���g)�ɂ���ƁA�ݒ�ƃR�l�N�g����x�ɂ���B
'    ��ConnectOn��False�ɂ���Ƥ�ݒ�݂̂ŃR�l�N�g�͍s�Ȃ�Ȃ�� (���ɃR�l�N�g����Ă���ꍇ�͂��̂܂�)
'
End Sub
Public Sub SetFIMV(ByVal ForceI As Variant, _
    ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True)
'���e:
'   PinList�ɐݒ肳�ꂽ�s����d�������Ԃɐݒ肷��
'
'�p�����[�^:
'    [ForceI]       In    ����d���B�z��w��\�B
'    [ClampV]       In    �N�����v�d���l�B
'    [Site]         In    �T�C�g�w��B�I�v�V����(Default -1)
'    [ConnectOn]    In    �f�o�C�X�ɐڑ����邩�ǂ����B�I�v�V����(Default True)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s����d�������Ԃɂ���
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��ForceI�ň���d�����w��BForceV�͐��l or �T�C�g�����̔z��B
'    �����l�̏ꍇ�S�T�C�g�����l��ݒ衔z��̏ꍇ�e�T�C�g�ɑΉ�����z��v�f�̒l��ݒ�
'    ��ClampV�ŃN�����v�d����ݒ�
'    ��Site�őΏۂ̃T�C�g���w��B-1(�f�t�H���g)������ƁA�S�T�C�g���ΏہB
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'    ���d����������W��ForceI��d�����背���W��ClampV���画�肵�Đݒ肷��
'    ��ConnectOn��True(�f�t�H���g)�ɂ���ƁA�ݒ�ƃR�l�N�g����x�ɂ���B
'    ��ConnectOn��False�ɂ���Ƥ�ݒ�݂̂ŃR�l�N�g�͍s�Ȃ�Ȃ�� (���ɃR�l�N�g����Ă���ꍇ�͂��̂܂�)
'
End Sub
'V21-Release
Public Sub SetFIMVMulti(ByVal ForceI As Variant, _
    ByVal ClampV As Double, _
    Optional ByVal ConnectOn As Boolean = True)
'���e:
'   PinList�ɐݒ肳�ꂽ�s����d�������Ԃɐݒ肷��(Site����)
'
'�p�����[�^:
'    [ForceI]       In    ����d���B�z��w��\�B
'    [ClampV]       In    �N�����v�d���l�B
'    [ConnectOn]    In    �f�o�C�X�ɐڑ����邩�ǂ����B�I�v�V����(Default True)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s����d�������Ԃɂ���
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��ForceI�ň���d�����w��BForceV�͐��l or �T�C�g�����̔z��B
'    �����l�̏ꍇ�S�T�C�g�����l��ݒ衔z��̏ꍇ�e�T�C�g�ɑΉ�����z��v�f�̒l��ݒ�
'    ��ClampV�ŃN�����v�d����ݒ�
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'    ���d����������W��ForceI��d�����背���W��ClampV���画�肵�Đݒ肷��
'    ��ConnectOn��True(�f�t�H���g)�ɂ���ƁA�ݒ�ƃR�l�N�g����x�ɂ���B
'    ��ConnectOn��False�ɂ���Ƥ�ݒ�݂̂ŃR�l�N�g�͍s�Ȃ�Ȃ�� (���ɃR�l�N�g����Ă���ꍇ�͂��̂܂�)
'
End Sub
Public Sub SetMV(ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True)
'���e:
'   PinList�ɐݒ肳�ꂽ�s����d�������Ԃɐݒ肷��B
'
'�p�����[�^:
'    [ClampV]       In    �N�����v�d���l�B
'    [Site]         In    �T�C�g�w��B�I�v�V����(Default -1)
'    [ConnectOn]    In    �f�o�C�X�ɐڑ����邩�ǂ����B�I�v�V����(Default True)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��ClampV�ŃN�����v�d����ݒ�(APMU���\�[�X�̂݃N�����v�@�\�L��)
'    ��Site�őΏۂ̃T�C�g���w��B-1(�f�t�H���g)������ƁA�S�T�C�g���ΏہB
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'    ���d�����背���W��ClampV���画�肵�Đݒ肷��
'    ��ConnectOn��True(�f�t�H���g)�ɂ���ƁA�ݒ�ƃR�l�N�g����x�ɂ���B
'    ��ConnectOn��False�ɂ���Ƥ�ݒ�݂̂ŃR�l�N�g�͍s�Ȃ�Ȃ�� (���ɃR�l�N�g����Ă���ꍇ�͂��̂܂�)
'
End Sub
'V21-Release
Public Sub SetMVMulti(ByVal ClampV As Double, _
    Optional ByVal ConnectOn As Boolean = True)
'���e:
'   PinList�ɐݒ肳�ꂽ�s����d�������Ԃɐݒ肷��B
'
'�p�����[�^:
'    [ClampV]       In    �N�����v�d���l�B
'    [ConnectOn]    In    �f�o�C�X�ɐڑ����邩�ǂ����B�I�v�V����(Default True)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��ClampV�ŃN�����v�d����ݒ�(APMU���\�[�X�̂݃N�����v�@�\�L��)
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'    ���d�����背���W��ClampV���画�肵�Đݒ肷��
'    ��ConnectOn��True(�f�t�H���g)�ɂ���ƁA�ݒ�ƃR�l�N�g����x�ɂ���B
'    ��ConnectOn��False�ɂ���Ƥ�ݒ�݂̂ŃR�l�N�g�͍s�Ȃ�Ȃ�� (���ɃR�l�N�g����Ă���ꍇ�͂��̂܂�)
'
End Sub
Public Sub ConnectPins(Optional ByVal site As Long = ALL_SITE)
'���e:
'   PinList�ɐݒ肳�ꂽ�s�����f�o�C�X�ɐڑ�����
'
'�p�����[�^:
'    [Site]    In   �T�C�g�w��B�I�v�V����(Default -1)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s�����f�o�C�X�ɐڑ�����
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��Site�őΏۂ̃T�C�g���w��B-1(�f�t�H���g)������ƁA�S�T�C�g���ΏہB
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'
End Sub

Public Sub DisconnectPins(Optional ByVal site As Long = ALL_SITE)
'���e:
'   PinList�ɐݒ肳�ꂽ�s�����f�o�C�X����؂藣���
'
'�p�����[�^:
'    [Site]    In   �T�C�g�w��B�I�v�V����(Default -1)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s�����f�o�C�X����؂藣���
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��Site�őΏۂ̃T�C�g���w��B-1(�f�t�H���g)������ƁA�S�T�C�g���ΏہB
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'    ���؂藣������A0V�����Ԃɐݒ肷��(�ڑ��͂��Ȃ�)�B
'
End Sub
'V21-Release
Public Sub DisconnectPinsMulti()
'���e:
'   PinList�ɐݒ肳�ꂽ�s�����f�o�C�X����؂藣���(Site����)
'
'�p�����[�^:
'
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s�����f�o�C�X����؂藣���
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ���؂藣������A0V�����Ԃɐݒ肷��(�ڑ��͂��Ȃ�)�B
'
End Sub
Public Sub SetGND(Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True)
'���e:
'   PinList�ɐݒ肳�ꂽ�s����ڒn(0V���)��Ԃɂ���
'
'�p�����[�^:
'    [Site]         In    �T�C�g�w��B�I�v�V����(Default -1)
'    [ConnectOn]    In    �f�o�C�X�ɐڑ����邩�ǂ����B�I�v�V����(Default True)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s����ڒn����(0V������)
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��Site�őΏۂ̃T�C�g���w��B-1(�f�t�H���g)������ƁA�S�T�C�g���ΏہB
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'    ��ConnectOn��True(�f�t�H���g)�ɂ���ƁA�ݒ�ƃR�l�N�g����x�ɂ���B
'    ��ConnectOn��False�ɂ���Ƥ�ݒ�݂̂ŃR�l�N�g�͍s�Ȃ�Ȃ�� (���ɃR�l�N�g����Ă���ꍇ�͂��̂܂�)
'
End Sub

Public Sub ChangeMIRange(ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE)
'���e:
'   PinList�ɐݒ肳�ꂽ�s���̓d�����背���W��ύX����
'
'�p�����[�^:
'    [ClampI]    In    �N�����v�d���l�B
'    [Site]      In    �T�C�g�w��B�I�v�V����(Default -1)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s���̈���d����ς�����d�����背���W��ύX����
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��ClampI�ŃN�����v�d����ݒ�
'    ���d�����背���W��ClampI���画�肵�Đݒ肷��
'    ��Site�őΏۂ̃T�C�g���w��B-1(�f�t�H���g)������ƁA�S�T�C�g���ΏہB
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'
End Sub
'V21-Release
Public Sub ChangeMIRangeMulti(ByVal ClampI As Double)
'���e:
'   PinList�ɐݒ肳�ꂽ�s���̓d�����背���W��ύX����
'
'�p�����[�^:
'    [ClampI]    In    �N�����v�d���l�B

'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s���̈���d����ς�����d�����背���W��ύX����
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ��ClampI�ŃN�����v�d����ݒ�
'    ���d�����背���W��ClampI���画�肵�Đݒ肷��
'    ����A�N�e�B�u�T�C�g�ɑ΂��Ă͉������Ȃ��
'
End Sub

Public Sub MeasureV(ByRef retResult() As Double, _
    ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE)
'���e:
'   PinList�ɐݒ肳�ꂽ�s���̓d���l�𑪒肷��
'
'�p�����[�^:
'    [retResult()]   Out   ���ʂ��i�[����z��B(�T�C�g�����̗v�f���K�v)
'    [AvgNum]        In    ���ω񐔁B
'    [Site]          In    �T�C�g�w��B�I�v�V����(Default -1)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s���̓d���l�𑪒肷��
'    ��PinList�͒P�ƃs���̂݁B�����s���A�s���O���[�v(�M�����O�s��������)�̎w��͕s�B
'    ��retResult�Ɍ��ʂ�Ԃ��
'    ��AvgNum�̐������T���v�����O���Ĥ���̕��ς�Ԃ��
'    ��Site�őΏۂ̃T�C�g���w��B-1(�f�t�H���g)������ƁA�S�T�C�g���ΏہB
'    ��Site��-1�ȊO�̏ꍇ�A�ΏۃT�C�g�ȊO��retResult()�̒l�͕ω������Ȃ��B
'    ����A�N�e�B�u�T�C�g�̌��ʂ�0��Ԃ��BSite��-1�̎������l�B
'
End Sub

Public Sub MeasureI(ByRef retResult() As Double, _
    ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE)
'���e:
'   PinList�ɐݒ肳�ꂽ�s���̓d���l�𑪒肷��
'
'�p�����[�^:
'    [retResult()]    Out   ���ʂ��i�[����z��B(�T�C�g�����̗v�f���K�v)
'    [AvgNum]         In    ���ω񐔁B
'    [Site]           In    �T�C�g�w��B�I�v�V����(Default -1)
'
'�߂�l:
'
'�ڍד���
'    ��PinList�ɑΉ�����s���̓d���l�𑪒肷��
'    ��PinList�͒P�ƃs���̂݁B�����s���A�s���O���[�v(�M�����O�s��������)�̎w��͕s�B
'    ��retResult�Ɍ��ʂ�Ԃ��
'    ��AvgNum�̐������T���v�����O���Ĥ���̕��ς�Ԃ��
'    ��Site�őΏۂ̃T�C�g���w��B-1(�f�t�H���g)������ƁA�S�T�C�g���ΏہB
'    ��Site��-1�ȊO�̏ꍇ�A�ΏۃT�C�g�ȊO��retResult()�̒l�͕ω������Ȃ��B
'    ����A�N�e�B�u�T�C�g�̌��ʂ�0��Ԃ��BSite��-1�̎������l�B
'
End Sub

Public Sub MeasureIMulti(ByVal avgNum As Long)
'���e:
'   PinList�ɐݒ肳�ꂽ�s���̓d���l�𑪒肷��
'
'�p�����[�^:
'    [AvgNum]    In   ���ω񐔁B
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s���̓d���l�𑪒肷�題��ʂ�ReadMIMultiResult���g���ēǂݏo���
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ���M�����O�s���͑Ή����Ȃ��
'    ��AvgNum�̐������T���v�����O���Ĥ���̕��ς�Ԃ��
'    ���S�T�C�g���ΏۂɂȂ顔�A�N�e�B�u�T�C�g�̌��ʂ�0�ɂȂ�
'
End Sub

Public Sub MeasureVMulti(ByVal avgNum As Long)
'���e:
'   PinList�ɐݒ肳�ꂽ�s���̓d���l�𑪒肷��
'
'�p�����[�^:
'    [AvgNum]    In    ���ω񐔁B
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�ɑΉ�����s���̓d���l�𑪒肷�題��ʂ�ReadMVMultiResult���g���ēǂݏo���
'    ��PinList�̓J���}��؂�̃s�����X�g��s���O���[�v�w��\�
'    ���M�����O�s���͑Ή����Ȃ��
'    ��AvgNum�̐������T���v�����O���Ĥ���̕��ς�Ԃ��
'    ���S�T�C�g���ΏۂɂȂ顔�A�N�e�B�u�T�C�g�̌��ʂ�0�ɂȂ�
'
End Sub

Public Sub ReadMVMultiResult(ByRef retResult() As Double)
'���e:
'   MeasureVMulti�ő��肵���d���l��ǂݏo���
'
'�p�����[�^:
'    [retResult()]    Out   ���ʂ��i�[����z��B(�T�C�g�����̗v�f���K�v)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�͍Ō�Ɏg����MeasureVMulti�̃s�����X�g�Ɋ܂܂��s�������w��
'    ��retResult�Ɍ��ʂ�Ԃ��
'
End Sub

Public Sub ReadMIMultiResult(ByRef retResult() As Double)
'���e:
'   MeasureIMulti�ő��肵���d���l��ǂݏo���
'
'�p�����[�^:
'    [retResult()]    Out   ���ʂ��i�[����z��B(�T�C�g�����̗v�f���K�v)
'
'�߂�l:
'
'���ӎ���:
'    ��PinList�͍Ō�Ɏg����MeasureIMulti�̃s�����X�g�Ɋ܂܂��s�������w��
'    ��retResult�Ɍ��ʂ�Ԃ��
'
End Sub

Public Function GetMeasureRange() As Collection
'���e:
'   PinList�ɐݒ肳�ꂽ�s���̃��W���[�����W���擾
'
'�p�����[�^:
'
'�߂�l:
'   ���W���[�����W�擾����
'
'���ӎ���:
'���ʂ̎��o����:
'   �擾���ʃR���N�V�������A
'   PinList�̃s�������L�[�ɁA�s������
'   ���背���W�l�����o�����Ƃ��ł��܂�
'
'MI���[�h����GANG(MEARGE)�ڑ��̑Ή�:
'   APMU/HDVIS���\�[�X�Ɋւ��āAGANG(MEARGE)�ڑ�
'   ���s�����߂�PinGP��PinList�Ɏw�肵�����A
'   MI���[�h�̏ꍇ�́A����ڑ����l������
'   �d�������W�l��Ԃ��܂��B
'   (DPS���\�[�X��GANG�ڑ��ɂ͑Ή����Ă��܂���)
'
'GangGp�s���̎擾��:
'   FVMI���[�h�ɐݒ肳�ꂽ3CH��GangGP(PowerType��PinGp)�ł���
'   GANG_APMU�iG_APMU1,G_APMU2,G_APMU3�j���w�肷���
'
'   �擾���� TheDC.Pins("GANG_APMU").GetMeasureRange
'   �擾���� 150[mA] (IRange50[mA]�~����ڑ���3)
'
'GANG�ڑ�GP�̃}���`�w��:
'   GANG�ڑ��p�̃}���`PinGP�w��͑Ή����܂���
'   GANG�ڑ��p��PinGP�𕡐�������PinList�Ɏw�肷���
'   �G���[�ƂȂ�܂�
'
'   (�G���[��)
'    �~ TheDC.Pins("GANG_ApmuGp1,GANG_ApmuGp2).GetMeasureRange
'
'FAIL�T�C�g������ꍇ:
'   ���݂���T�C�g�̂���1�ł�Active�łȂ��T�C�g�����݂���
'   ��ԂŖ��߂����s�����Ƃ��̓G���[�ƂȂ�܂��B
'   ���݃T�C�g�����ׂ�Active�ȏ�ԂŎ��s���Ă�������
'
'IG-XL�V�~�����[�^���ł̎��s:
'   IG-XL�V�~�����[�^���ɂ����Ă�BPMU�s���̃����W�͎擾�ł��܂���B
'   ���s����ƃG���[���b�Z�[�W�o�͌�A�擾���ʂ�Nothing�ƂȂ�܂�

End Function

Public Function GetForceVLimit() As Double()
'���e:
'   �d������̎��ɋ������ő�l�ƍŏ��l���擾
'
'�p�����[�^:
'
'�߂�l:
'   �d��������̐ݒ�\�ȍő�l�A�ŏ��l
'   �v�f��0-->�ŏ��l�A�v�f��1-->�ő�l
'
'���ӎ���:
'
'
End Function

Public Function GetForceILimit() As Double()
'���e:
'   �d������̎��ɋ������ő�l�ƍŏ��l���擾
'
'�p�����[�^:
'
'�߂�l:
'   �d��������̐ݒ�\�ȍő�l�A�ŏ��l
'   �v�f��0-->�ŏ��l�A�v�f��1-->�ő�l
'
'���ӎ���:
'GANG�iMEARGE�j�ڑ��̃T�|�[�g:
'   APMU/HDVIS���\�[�X�Ɋւ��āAGANG(MEARGE)�ڑ�
'   ���s�����߂�PinGP��PinList�Ɏw�肵�����́A
'   ����ڑ������l������Force�d���l��Ԃ��܂��B
'
'GangGp�s���̎擾��:
'   3CH��GangGP(PowerType��PinGp)�ł���
'   GANG_APMU�iG_APMU1,G_APMU2,G_APMU3�j���w�肷���
'
'   �擾���� TheDC.Pins("GANG_APMU").GetForceILimit
'   �擾���� 150[mA] (50[mA]�~����ڑ���3)
'
'GANG�ڑ�GP�̃}���`�w��:
'   GANG�ڑ��p�̃}���`PinGP�w��͑Ή����܂���
'   GANG�ڑ��p��PinGP�𕡐�������PinList�Ɏw�肷���
'   �G���[�ƂȂ�܂�
'
'   (�G���[��)
'   �~ TheDC.Pins("GANG_ApmuGp1,GANG_ApmuGp2).GetForceILimit
'
'DPS���\�[�X:
'   DPS���\�[�X��FI���[�h�����݂��Ȃ�����
'   ���߂����s����ƃG���[�ƂȂ�܂�
'
End Function

Public Function GetForceValue() As Collection
'���e:
'   PinList�ɐݒ肳�ꂽ�s����Force�l���擾
'
'�p�����[�^:
'
'�߂�l:
'   Force�l�A�擾����
'
'���ӎ���:
'���ʂ̎��o����:
'   �擾���ʃR���N�V�������A
'   PinList�̃s�������L�[�ɁA�s������
'   Force�l�����o�����Ƃ��ł��܂�
'
'GANG�ڑ�GP�̃}���`�w��:
'   GANG�ڑ��p�̃}���`PinGP�w��͑Ή����܂���
'   GANG�ڑ��p��PinGP�𕡐�������PinList�Ɏw�肷���
'   �G���[�ƂȂ�܂�
'
'   (�G���[��)
'    �~ TheDC.Pins("GANG_ApmuGp1,GANG_ApmuGp2).GetMeasureRange
'
'FAIL�T�C�g������ꍇ:
'   ���݂���T�C�g�̂���1�ł�Active�łȂ��T�C�g�����݂���
'   ��ԂŎ��s�����Ƃ��̓G���[�ƂȂ�܂��B
'   ���݃T�C�g�����ׂ�Active�ȏ�ԂŎ��s���Ă�������
'
'DPS���\�[�X�̎擾:
'   DPS��Force�l�A�擾���ʂ͏�ɁudpsPrimaryVoltage�v�̒l�ƂȂ�܂�
'
'IG-XL�V�~�����[�^���ł̎��s:
'   IG-XL�V�~�����[�^���ɂ����Ă�BPMU�s����Force�l�͎擾�ł��܂���B
'   ���s����ƃG���[���b�Z�[�W�o�͌�A�擾���ʂ�Nothing�ƂȂ�܂�
'
End Function

Public Property Get ForceIRangeList() As String
'���e:
'   PinList�ɐݒ肳�ꂽ�s���̃��\�[�X��
'   �ݒ�\��Force�d�������W���X�g��Ԃ�
'
'�p�����[�^:
'
'�߂�l:
'   �J���}��؂�`���̓d�������W���X�g
'   (��) 40uA,200uA,1mA,5mA,50mA
'   �������W�p�����[�^�����݂��Ȃ��Ƃ���"None"�ƂȂ�܂��B
'   ��APMU��GANG�Ώۃs����50mA��Ԃ��܂�
'
'���ӎ���:
'   DPS�ɂ�FI���[�h�͑��݂��܂���
'   Auto�ASmart�͏o�͂��܂���
'
End Property

Public Property Get ForceVRangeList() As String
'���e:
'   PinList�ɐݒ肳�ꂽ�s���̃��\�[�X��
'   �ݒ�\��Force�d�������W���X�g��Ԃ�
'
'�p�����[�^:
'
'�߂�l:
'   �J���}��؂�`���̓d�������W���X�g
'   (��) 2V,5V,10V,35V
'   �������W�p�����[�^�����݂��Ȃ��Ƃ���"None"�ƂȂ�܂��B
'
'���ӎ���:
'   DPS�APPMU�ɂ͓d�������W�̃p�����[�^�͑��݂��܂���B
'   APMU��40V�͏o�͂��܂���
'   Auto�ASmart�͏o�͂��܂���
'
End Property

Public Property Get MeasIRangeList() As String
'���e:
'   PinList�ɐݒ肳�ꂽ�s���̃��\�[�X��
'   �ݒ�\��Measure�d�������W���X�g��Ԃ�
'
'�p�����[�^:
'
'�߂�l:
'   �J���}��؂�`���̓d�������W���X�g
'   (��) 200nA,2uA,10uA,40uA,200uA,1mA,5mA,50mA
'   �������W�p�����[�^�����݂��Ȃ��Ƃ���"None"�ƂȂ�܂��B
'   ��APMU��GANG�Ώۃs����50mA��Ԃ��܂�
'
'���ӎ���:
'   BPMU��2uA�͏o�͂��܂���
'   Auto�ASmart�͏o�͂��܂���
'
End Property

Public Property Get MeasVRangeList() As String
'���e:
'   PinList�ɐݒ肳�ꂽ�s�����\�[�X��
'   �ݒ�\��Measure�d�������W���X�g��Ԃ�
'
'�p�����[�^:
'
'�߂�l:
'   �J���}��؂�`���̓d�������W���X�g
'   (��) 2V,5V,10V,35V
'   �������W�p�����[�^�����݂��Ȃ��Ƃ���"None"�ƂȂ�܂��B
'
'���ӎ���:
'   APMU��40V�͏o�͂��܂���
'   Auto�ASmart�͏o�͂��܂���
'
End Property
