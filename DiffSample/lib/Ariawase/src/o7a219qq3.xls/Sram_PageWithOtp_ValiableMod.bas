Attribute VB_Name = "Sram_PageWithOtp_ValiableMod"
Option Explicit

'----- SRAM Infomation (Adjust) ---------------------------
Public Const Ef_Bist_Rd_En_Width As Integer = 1                                     'OTP�������Ă���1�璷�C�l�[�u������BIT��
Public Const Ef_Bist_Rd_Addr_Width As Integer = 9                                   'OTP�������Ă���1�璷�A�h���X����BIT��
Public Const Ef_Bist_Rd_Data_Width As Integer = 8                                   'OTP�������Ă���1�璷�f�[�^����BIT��
Public Const MAX_EF_BIST_RD_BIT As Byte = 4                                         '�ő�␳Memory��
Public Const BitWidth_OtpRep_Parity As Integer = 4                                  'OTP�������Ă���FUSE�璷�p���e�B�[����BIT��
Public Const BitWidth_OtpRep_Addr As Integer = 7                                    'OTP�������Ă���FUSE�璷�A�h���X����BIT��
Public Const BitWidth_OtpRep_EN As Integer = 1                                      'OTP�������Ă���FUSE�璷�C�l�[�u������BIT��
Public Const BitWidth_OtpRep_Data As Integer = 1                                    'OTP�������Ă���FUSE�璷Data����BIT��
Public Const Bist_Num_Mem As Integer = 12                                           'Memory��
Public Const Bist_Max_Num_Io As Integer = 50                                       '�ő�IO���i�eMemory�������Ă���IO���̒��̍ő�IO�����`�j
Public Const RCON_START_Addr As Integer = 41                                       'RCON�̍ŏ��̓]���f�[�^Bit�ԍ�
Public Const RCON_END_Addr As Integer = 1                                           'RCON�̍Ō�̓]���f�[�^Bit�ԍ�
Public Const SRAMRD_OUTPUT_PIN As String = "Ph_FSTROBE"                      '�p�^�[��OUTPUT(BIST) PIN
Public Const RCON_ChainType As String = "Descending"           'RCON�̃`�F�[���i���o�[������(Ascending)���~��(Descending)��
Public Const RCON_FirstInfoType As String = "MEMID_1st"   'RCON��Memory���̍ŏ���Bit��MEMID(MEMID_1st)��FAILINFO(FAILINFO_1st)��
Public Const RCON_FailInfoType As String = "Ascending"     'RCON��FAILINFO�i���o�[������(Ascending)���~��(Descending)��

'----- SRAM Infomation (Fix) ------------------------------
Public BIST_NUM_IO(Bist_Num_Mem + 1) As Integer                                     '�eMemory�������Ă���IO�����i�[����ϐ��@(�i�[��Ƃ�SRAM�����̍ŏ��ōs��)
Public BIST_RED_TYPE(Bist_Num_Mem + 1) As Integer                                   '�eMemory���璷IO��ێ����Ă��邩���i�[����ϐ��@(�i�[��Ƃ�SRAM�����̍ŏ��ōs��)
Public BIST_IO_EN_NO(Bist_Num_Mem + 1) As Integer                                   '�eMemory�̐ڑ��`�F�[���iRCON�j�i���o�[���i�[����ϐ��@(�i�[��Ƃ�SRAM�����̍ŏ��ōs��)
Public BIST_FAIL_REG(nSite, Bist_Num_Mem + 1, Bist_Max_Num_Io + 1) As Byte          '�SPreSRAM�����̕s��Memory�ƕs��I/O�����܂Ƃ߂Ċi�[����ϐ�
Public BIST_FAIL_IO_NO(nSite, Bist_Num_Mem + 1) As Integer                          '�SPreSRAM������̏璷�\�Ɣ��肵���s��I/O�i���o�[���i�[����ϐ�
Public EF_BIST_REPAIR_DATA(nSite, 2 ^ Ef_Bist_Rd_Addr_Width) As Byte                'RCON�̊eMemory�ɑ΂���璷���(EN�ƕs��I/O�̏��)���i�[����ϐ��i�z��ԍ��͂��̂܂�RCON�ԍ��ƂȂ�j
Public RepairMemoryCount(nSite) As Long                                             '�璷�K�v�������[���i�[�ϐ�
Public Ef_Enbl_Addr(nSite, MAX_EF_BIST_RD_BIT) As Integer                           '�璷�C�l�[�u���f�[�^[Dec]
Public Ef_Rcon_Addr(nSite, MAX_EF_BIST_RD_BIT) As Integer                           '�璷�A�h���X�f�[�^[Dec]
Public Ef_Repr_Data(nSite, MAX_EF_BIST_RD_BIT) As Integer                           '�璷�f�[�^�̃f�[�^[Dec]
Public EF_BIST_ROLLCALL_INDEX() As Integer                                          'SRAM�璷RollCall�p�^�[���̊e���Ғl�����Vector���i�[�ϐ�
Public Const Rep_Bist_Data_Len As Integer = (RCON_START_Addr + RCON_END_Addr) + 1   'RCON�̃f�[�^Bit��(�L��RCON��+2��)
Public Const SramRepBlowLabel As String = "START"                                   'SRAM�璷BLOW�p�^�[���X�^�[�g���x����
Public Const SramRepVerifyLabel As String = "START"                                 'SRAM�璷VERIFY�p�^�[���X�^�[�g���x����
Public Const BitWidth_OtpRep_Total As Integer = BitWidth_OtpRep_Parity + _
                                                BitWidth_OtpRep_Addr + _
                                                BitWidth_OtpRep_EN + _
                                                BitWidth_OtpRep_Data                'OTP�������Ă���FUSE�璷Total��BIT��

Public Const TBL_CELL_PAT As Long = 2
Public Const TBL_CELL_LIST_STRow As Long = 5
Public Const TBL_DIR_CELL_Row As Long = 2
Public Const TBL_DIR_CELL_Col As Long = 3
Public Const TBL_INDEX_CYCLE As Long = 0        'TBL�t�@�C��1�s���X�y�[�X�ŋ�؂������́A1�ڂ̕�����i�J�E���g��0�n�܂�j
Public Const TBL_INDEX_BIT As Long = 2          'TBL�t�@�C��1�s���X�y�[�X�ŋ�؂������́A3�ڂ̕�����i�J�E���g��0�n�܂�j
Public Const TBL_INDEX_MACRO As Long = 5        'TBL�t�@�C��1�s���X�y�[�X�ŋ�؂������́A6�ڂ̕�����i�J�E���g��0�n�܂�j
Public dirTblFile As String                     'TBL�t�@�C����u���Ă���f�B���N�g��



'##### TBL INFOMATION VARIABLE ###################################################################
Public Type FAIL_CYCLE_INFO
    CycleNo As Long         'FailCycle�ԍ�
    MemoryNo As Long        'FailCycle�ɑΉ�����BIT(I/O)�ԍ�
    IoNo As Long            'FailCycle�ɑΉ�����}�N���ԍ�
End Type

Public Type TBL_LIST_INFO
    PatFileName As String   'PatGp�Ŏw�肳���p�^�[���t�@�C����
    TblFileName As String   'PatFile�ɑΉ�����TBL�t�@�C����
    FailInfo() As FAIL_CYCLE_INFO   'TBL�t�@�C���̒��g
End Type

Public TblInfo() As TBL_LIST_INFO


'    �ϐ�            �\���̇@          �\���̇A
' TblInfo(X0)                                    : SRAMBIST�̎�ސ�(���p�^�[������TBL�t�@�C����)���A�z�񂪗p�ӂ����
'        (X1)
'        (X2)-------PatFileName                  : ���̔z��Ɋ��蓖�Ă���p�^�[���̖��O
'               |---TblFileName                  : ���̃p�^�[���̏�񂪋L�ڂ���Ă���TBL�t�@�C���̖��O
'               |---FailInfo(Y0)                 : ����TBL�t�@�C���̏ڍ׏��B�璷�Ɋւ�����ҒlVector�����A�z�񂪗p�ӂ����
'               |---FailInfo(Y1)
'               |---FailInfo(Y2)
'               |---FailInfo(Y3)-------CycleNo   : ����Vector�Ɋ��蓖�Ă���Cycle�ԍ�
'                                  |---MemoryNo  : ����Vector�Ɋ��蓖�Ă���SRAM�������ԍ�
'                                  |---IoNo      : ����Vector�Ɋ��蓖�Ă���SRAM����������IO�ԍ�

'#################################################################################################


'----- Flag -----------------------------------------------
Public Flg_SRAM_BLOW As Integer                                                     'SRAM�璷ON/OFF �t���O
Public Flg_SramDebug As Integer                                                     'SRAM�f�o�b�O���O�f���o���t���O�i0:���O�f���o�������@1:���O�f���o���L��j
Public Bist_Alpg_Fail_Flag(nSite) As Byte                                           'ALPG Fail�t���O
Public Bist_Repairable_Flag(nSite) As Byte                                          'Repair ON/OFF �t���O
Public Flg_ActiveSite_SramRep(nSite) As Double                                      'SRAM�璷Blow����
Public Flg_PostSramRun As Boolean
Public Flg_SramBlankNg(nSite) As Long


Public Sub ValiableSet_SramDesignInfo_IO()

'+++ Test Infomation +++++++++++++++++++++++++++++++
'�eMemory�������Ă���IO���Ə璷�Z���̗L�������i�[
'RED_TYPE = 1 : �璷IO�L��
'RED_TYPE = 0 : �璷IO����
'+++++++++++++++++++++++++++++++++++++++++++++++++++
    
'!!!!! Must Const "-1" !!!!!!!!!!!!!!!!!!!!!!!!!!!!!
BIST_NUM_IO(0) = -1
BIST_RED_TYPE(0) = -1
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

BIST_NUM_IO(1) = 50: BIST_RED_TYPE(1) = 1
BIST_NUM_IO(2) = 50: BIST_RED_TYPE(2) = 1
BIST_NUM_IO(3) = 50: BIST_RED_TYPE(3) = 1
BIST_NUM_IO(4) = 38: BIST_RED_TYPE(4) = 0
BIST_NUM_IO(5) = 38: BIST_RED_TYPE(5) = 0
BIST_NUM_IO(6) = 38: BIST_RED_TYPE(6) = 0
BIST_NUM_IO(7) = 38: BIST_RED_TYPE(7) = 0
BIST_NUM_IO(8) = 48: BIST_RED_TYPE(8) = 0
BIST_NUM_IO(9) = 48: BIST_RED_TYPE(9) = 0
BIST_NUM_IO(10) = 29: BIST_RED_TYPE(10) = 1
BIST_NUM_IO(11) = 46: BIST_RED_TYPE(11) = 1
BIST_NUM_IO(12) = 46: BIST_RED_TYPE(12) = 1



End Sub
     
Public Sub ValiableSet_SramDesignInfo_RCON()
  
'+++ Test Infomation +++++++++++++++++++++++++++++++
'�eMemory�̐ڑ��`�F�[���iRCON�j�i���o�[���i�[
BIST_IO_EN_NO(1) = 41
BIST_IO_EN_NO(2) = 34
BIST_IO_EN_NO(3) = 27
BIST_IO_EN_NO(10) = 20
BIST_IO_EN_NO(11) = 14
BIST_IO_EN_NO(12) = 7
BIST_IO_EN_NO(4) = 0
BIST_IO_EN_NO(5) = 0
BIST_IO_EN_NO(6) = 0
BIST_IO_EN_NO(7) = 0
BIST_IO_EN_NO(8) = 0
BIST_IO_EN_NO(9) = 0
End Sub

Public Sub SramValiableClear()
  
'+++ Test Infomation +++++++++++++++++++++++++++++++
'SRAM���ϐ��̏����l�ݒ�
'+++++++++++++++++++++++++++++++++++++++++++++++++++

    Dim site As Long
    Dim NowPage As Long
    Dim NowBit As Long

    For site = 0 To nSite
        For NowPage = OtpPageStart To OtpPageEnd
            For NowBit = 0 To OtpMaxBitParPage - 1
                SramBlowDataAllBin(site, NowPage, NowBit) = "0"
                SramReadDataAllBin(site, NowPage, NowBit) = "L"
            Next NowBit
        Next NowPage
    Next site
    
    
    'Variable Clear
    Erase Ef_Enbl_Addr
    Erase Ef_Rcon_Addr
    Erase Ef_Repr_Data
    Erase BIST_FAIL_REG
    Erase BIST_FAIL_IO_NO
    Erase Bist_Alpg_Fail_Flag
    Erase Bist_Repairable_Flag
    Erase RepairMemoryCount
    Flg_PostSramRun = False
    Erase EF_BIST_REPAIR_DATA

End Sub
