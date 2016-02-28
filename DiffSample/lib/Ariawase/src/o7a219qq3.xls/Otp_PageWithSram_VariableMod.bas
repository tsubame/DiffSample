Attribute VB_Name = "Otp_PageWithSram_VariableMod"
Option Explicit

'������ Adjust Variable ������

Public Const RejiIn As String = "Ph_SDA"                      'Blow����ۂ�Blow�����̓s��
Public Const RejiOut As String = "Ph_SDA"                     'RollCall����ۂ�RollCall���o�̓s��
Public Const OtpPageStart As Integer = 0                 'OTPMAP�̃X�^�[�g�y�[�W�i���o�[
Public Const OtpPageEnd As Integer = 11                   'OTPMAP�̃��X�g�y�[�W�i���o�[
Public Const OtpPixOffset_X As Long = -1                  'OTP�A�h���X�ƕ����A�h���X�̍��فiX�����j
Public Const OtpPixOffset_Y As Long = 25                  'OTP�A�h���X�ƕ����A�h���X�̍��فiY�����j

Public Const BitWidthAll_Lot1  As Long = 4                  'Bit��
Public Page_Lot1(BitWidthAll_Lot1 - 1) As Long                 'OTP�̊eBit���A���y�[�W�ڂȂ̂��������ϐ�
Public Bit_Lot1(BitWidthAll_Lot1 - 1) As Long                  'OTP�̊eBit���A��Bit�ڂȂ̂��������ϐ�

Public Const BitWidthAll_Lot2  As Long = 4                  'Bit��
Public Page_Lot2(BitWidthAll_Lot2 - 1) As Long                 'OTP�̊eBit���A���y�[�W�ڂȂ̂��������ϐ�
Public Bit_Lot2(BitWidthAll_Lot2 - 1) As Long                  'OTP�̊eBit���A��Bit�ڂȂ̂��������ϐ�

Public Const BitWidthAll_Lot7  As Long = 4                  'Bit��
Public Page_Lot7(BitWidthAll_Lot7 - 1) As Long                 'OTP�̊eBit���A���y�[�W�ڂȂ̂��������ϐ�
Public Bit_Lot7(BitWidthAll_Lot7 - 1) As Long                  'OTP�̊eBit���A��Bit�ڂȂ̂��������ϐ�

Public Const BitWidthAll_Lot8  As Long = 4                  'Bit��
Public Page_Lot8(BitWidthAll_Lot8 - 1) As Long                 'OTP�̊eBit���A���y�[�W�ڂȂ̂��������ϐ�
Public Bit_Lot8(BitWidthAll_Lot8 - 1) As Long                  'OTP�̊eBit���A��Bit�ڂȂ̂��������ϐ�

Public Const BitWidthAll_Lot9  As Long = 4                  'Bit��
Public Page_Lot9(BitWidthAll_Lot9 - 1) As Long                 'OTP�̊eBit���A���y�[�W�ڂȂ̂��������ϐ�
Public Bit_Lot9(BitWidthAll_Lot9 - 1) As Long                  'OTP�̊eBit���A��Bit�ڂȂ̂��������ϐ�

Public Const BitWidthAll_Wafer  As Long = 8                  'Bit��
Public Page_Wafer(BitWidthAll_Wafer - 1) As Long                 'OTP�̊eBit���A���y�[�W�ڂȂ̂��������ϐ�
Public Bit_Wafer(BitWidthAll_Wafer - 1) As Long                  'OTP�̊eBit���A��Bit�ڂȂ̂��������ϐ�

Public Const BitWidthAll_Chip  As Long = 16                  'Bit��
Public Page_Chip(BitWidthAll_Chip - 1) As Long                 'OTP�̊eBit���A���y�[�W�ڂȂ̂��������ϐ�
Public Bit_Chip(BitWidthAll_Chip - 1) As Long                  'OTP�̊eBit���A��Bit�ڂȂ̂��������ϐ�

Public Const MaxRepair_Single_CP_FD As Integer = 10               '���ו␳�\�ő吔
Public Const BitWidthN_Single_CP_FD As Integer = 8               '���o��  bit��
Public Const BitWidthX_Single_CP_FD As Integer = 12               'X�A�h���X bit��
Public Const BitWidthY_Single_CP_FD As Integer = 12               'Y�A�h���X bit��
Public Const BitWidthS_Single_CP_FD As Integer = 2               'Source��� bit��
Public Const BitWidthD_Single_CP_FD As Integer = 2               'Direction bit��
Public Const DefRep_SrcType_Single_CP_FD As String = "SrcType1"           'Source���^�C�v
Public Const NgAddress_LeftS_Single_CP_FD As Long = 9            'Couplet ��p�[NG�A�h���X��(Left Start)
Public Const NgAddress_LeftE_Single_CP_FD As Long = 12            'Couplet ��p�[NG�A�h���X��(Left End)
Public Const NgAddress_RightS_Single_CP_FD As Long = 3285           'Couplet ��p�[NG�A�h���X��(Right Start)
Public Const NgAddress_RightE_Single_CP_FD As Long = 3288           'Couplet ��p�[NG�A�h���X��(Right End)
Public Const BitWidthAll_Single_CP_FD As Long = BitWidthN_Single_CP_FD + (MaxRepair_Single_CP_FD * (BitWidthX_Single_CP_FD + BitWidthY_Single_CP_FD + BitWidthS_Single_CP_FD + BitWidthD_Single_CP_FD))               '���ו␳�ɕK�v��bit��
Public Page_Single_CP_FD(BitWidthAll_Single_CP_FD - 1) As Long              '�␳�ɕK�v��OTP�̊eBit���A���y�[�W�ڂȂ̂��������ϐ�
Public Bit_Single_CP_FD(BitWidthAll_Single_CP_FD - 1) As Long               '�␳�ɕK�v��OTP�̊eBit���A��Bit�ڂȂ̂��������ϐ�


Public Const BitWidth_O_TEMP  As Long = 10                 '���x�v�I�t�Z�b�g bit��
Public Const BitWidth_S_TEMP  As Long = 10                 '���x�v�X�� bit��
Public Const BitWidthAll_TEMP As Long = BitWidth_O_TEMP + BitWidth_S_TEMP                '���ו␳�ɕK�v��bit��
Public Page_TEMP(BitWidthAll_TEMP - 1) As Long               '�␳�ɕK�v��OTP�̊eBit���A���y�[�W�ڂȂ̂��������ϐ�
Public Bit_TEMP(BitWidthAll_TEMP - 1) As Long                '�␳�ɕK�v��OTP�̊eBit���A��Bit�ڂȂ̂��������ϐ�

Public Const BitWidthAll_SRAM  As Long = 72                  'Bit��
Public Page_SRAM(BitWidthAll_SRAM - 1) As Long                 'OTP�̊eBit���A���y�[�W�ڂȂ̂��������ϐ�
Public Bit_SRAM(BitWidthAll_SRAM - 1) As Long                  'OTP�̊eBit���A��Bit�ڂȂ̂��������ϐ�

'��Fix Variable
'����������������������������������������������������������
Public Const OtpMaxBitParPage As Integer = 512                                          '�e�y�[�W�������Ă���Bit���̒��̍ő�Bit��(1�n�܂�ŃJ�E���g)
Public Const OtpInfoSheet_Row_Page As Long = 2                                          'OTP_Infomation�V�[�g �y�[�W����Row���X�^�[�g�I�t�Z�b�g
Public Const OtpInfoSheet_Column_Page As Long = 2                                       'OTP_Infomation�V�[�g �y�[�W����Column���X�^�[�g�I�t�Z�b�g
Public Const OtpInfoSheet_Row_Bit As Long = 2                                           'OTP_Infomation�V�[�g Bit(Dec)����Row���X�^�[�g�I�t�Z�b�g
Public Const OtpInfoSheet_Column_Bit As Long = 5                                        'OTP_Infomation�V�[�g Bit(Dec)����Column���X�^�[�g�I�t�Z�b�g
Public Const OtpInfoSheet_Row_BlowInfo As Long = 2                                      'OTP_Infomation�V�[�g Blow���e�i�ϓ��l���j��Row���X�^�[�g�I�t�Z�b�g
Public Const OtpInfoSheet_Column_BlowInfo As Long = 8                                   'OTP_Infomation�V�[�g Blow���e�i�ϓ��l���jColumn���X�^�[�g�I�t�Z�b�g
Public Const OtpInfoSheet_Row_Value As Long = 2                                         'OTP_Infomation�V�[�g Value(Bin)��Row���X�^�[�g�I�t�Z�b�g
Public Const OtpInfoSheet_Column_Value As Long = 9                                      'OTP_Infomation�V�[�g Value(Bin)��Column���X�^�[�g�I�t�Z�b�g
Public Const OtpInfoSheet_Row_FF As Long = 2                                            'OTP_Infomation�V�[�g FF�������ݏ���Row���X�^�[�g�I�t�Z�b�g
Public Const OtpInfoSheet_Column_FF As Long = 11                                        'OTP_Infomation�V�[�g FF�������ݏ���Column���X�^�[�g�I�t�Z�b�g
Public Const BitParHex As Integer = 8                                                   '1Hex��񂪉�bit��
Public Const OtpPageSize As Integer = OtpPageEnd + 1                                    'OTPMAP�̃y�[�W�����B�ԍ��̃X�^�[�g��0����ł��邱�ƑO��
Public AddrParPage(OtpPageSize - 1) As Integer                                          '�ePage����Address��
Public BitParPage(OtpPageSize - 1) As Integer                                           '�ePage����bit��
Public BlowDataAllBin(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String           '�SBlow���(Bin)
Public BlowDataAllBin2(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String          '�SBlow���(Bin) �Œ�l�p
Public ReadDataAllBin(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String           '�SRead���(Bin)
Public FFBlowInfo(OtpPageSize - 1, OtpMaxBitParPage - 1) As String                      'FFBlow���(Bin)

Public Const Label_Page_OtpBlow As String = "OTP_Blow_BlowLabel_Page_OtpBlowPage"                                    'BLOW�p�^�[�� Page���X�^�[�g���x����
Public Const Label_Page_OtpBlow_Break As String = "OTP_Blow_BlowLabel_Page_OtpBlow_Break"                            'FFBLOW�p�^�[�� Page���X�^�[�g���x����
Public Const Label_Page_OtpVerify As String = "OTP_Verify_VerifyLabel_Page_OtpVerifyPage"                            'VERIFY�p�^�[�� Page���X�^�[�g���x����
Public Const Label_Page_BlankCheck As String = "OTP_Verify_VerifyLabel_Page_BlankCheckPage"                          'BLANK�`�F�b�N�p�^�[�� Page���X�^�[�g���x����
Public Const Label_Page_OtpFixedValueCheck As String = "OTP_Verify_VerifyLabel_Page_OtpFixedValueCheckPage"          '�Œ�l�`�F�b�N�p�^�[�� Page���X�^�[�g���x����

Public Const Label_OtpBlow As String = "OTP_Blow_BlowLabel_OtpBlowPage"                                              'BLOW�p�^�[�� Blow���X�^�[�g���x����
Public Const Label_OtpBlowAuto As String = "OTP_Blow_BlowLabelAuto_OtpBlowPage"                                      'BLOW�p�^�[�� AutoBlow���X�^�[�g���x����
Public Const Label_OtpBlow_Break As String = "OTP_Blow_BlowLabel_OtpBlow_Break"                                      'FFBLOW�p�^�[�� Blow���X�^�[�g���x����
Public Const Label_OtpBlowAuto_Break As String = "OTP_Blow_BlowLabelAuto_OtpBlow_Break"                              'FFBLOW�p�^�[�� AutoBlow���X�^�[�g���x����
Public Const Label_OtpVerify As String = "OTP_Verify_VerifyLabel_OtpVerifyPage"                                      'VERIFY�p�^�[�� Verify���X�^�[�g���x����
Public Const Label_BlankCheck As String = "OTP_Verify_VerifyLabel_BlankCheckPage"                                    'BLANK�`�F�b�N�p�^�[�� Verify���X�^�[�g���x����
Public Const Label_OtpFixedValueCheck As String = "OTP_Verify_VerifyLabel_OtpFixedValueCheckPage"                    '�Œ�l�`�F�b�N�p�^�[�� Verify���X�^�[�g���x����

Public Vector_OtpRead(OtpMaxBitParPage - 1) As Integer                            'RollCall�p�^�[���eVector���i�[�ϐ�
Public Const ByteParVector_VerifyPat As Long = 9

'----- SRAM Repair Blow Infomation (Fix) -----------------
Public SramBlowDataAllBin(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String           '�SBlow���(Bin) SRAM�璷��p�̕ϐ�
Public SramReadDataAllBin(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String           '�SRead���(Bin) SRAM�璷��p�̕ϐ�

'----- Flag -----------------------------------------------
Public Flg_OTP_BLOW As Integer                                                          '���蒆��OTPBLOW�����s���邩�̑I��
Public Flg_OtpBlowPage(OtpPageSize - 1) As Boolean                                      'OTPBLOW�����s����p�^�[���i�y�[�W�j�̑I���iTrue�F�p�^�[�����s�@False�F�p�^�[�����s�����j
Public Flg_OtpBlowFixValPage As Integer                                                 '�Œ�l��Blow����Page������������t���O�B�Œ�lPage������Page����΍Ō�̌Œ�lPage�݂̂�ێ��B
Public Flg_ModifyPage(OtpPageSize - 1) As Boolean                                       'OTPBLOW�̕ϓ��lModify���s���p�^�[���i�y�[�W�j�̑I���iTrue�FModify���s�@False�FModify���s�����j
Public Flg_ActiveSite_OTP(nSite) As Double
Public FFBlowPage As Long                                                               'FFBLOW�����s����Page���
Public Flg_ModifyPageSRAM(OtpPageSize - 1) As Boolean                               'SRAM�璷��Modify���s���p�^�[���i�y�[�W�j�̑I���iTrue�FModify���s�@False�FModify���s�����j

