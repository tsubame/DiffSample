VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MIFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If EEE_AUTO_CAPTURE = 0 Then
Option Explicit

'==========================================================================
'
'   MIFC(MIPI/SMIA/CMOS Interface FPGA Controler)
'   Class Module Fuction :  This Class Modude is control tool of MSC Interface FPGA
'                           And ICUD2
'
'
'
'   First Created By T.Minoda(Teradyne)  2008/07/21
'
'   ///Version 1.2///
'
'   Update history
'   Ver0.5  first edition for debugging.
'   Ver0.6  Adding mipi function for debugging. By T.Minoda 2008/08/01
'   Ver0.7  Adding mipi data line RFSW control method. By T.Minoda 2008/08/07
'           --> MSC IF FPGA Rev06
'   Ver0.8  Adding input data select method of cmos through. By T.Minoda 2008/08/28
'           --> MSC IF FPGA Rev0F
'   Ver0.9  Adding capture mode of MIPI 3/4 Lane, and Frame skip function. By H.Miyamoto 2009/03/24
'           --> MSC IF FPGA Rev10
'   Ver1.0  Adding DecodeTimeOut function. By T.Minoda 2009/06/29
'           --> MSC IF FPGA Rev12
'   Ver1.0a Modify DecodeTimeOut function. By T.Minoda 2009/07/01
'           --> MSC IF FPGA Rev13
'   Ver1.1  Modify DecodeTimeOut function, adding CMOS12bit, assign select function, and changed InitTapValue sequence. By T.Minoda 2009/09/17
'           --> MSC IF FPGA Rev14
'   Ver1.1a Correct enum value of LVDS_PN_REVERSE. By T.Minoda 2009/10/06
'           --> MSC IF FPGA Rev16
'   Ver1.2  Adding MIPI raw mode select, Unsupport SMIA Capture if decoder FPGA rev=2.1. By Y.Kimura 2011/03/23
'           --> MSC IF FPGA Rev21
'
'==========================================================================


'+++ Define API Fucntion +++++
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


'+++ Define Const Value +++++
Private Const COMMAND_SETCAPTUREMODE As Long = 1
Private Const COMMAND_SETCLOCKMODE As Long = 2
Private Const COMMAND_SETCMOSATTRIBUTEMODE As Long = 3
Private Const COMMAND_SETAVENUM As Long = 4
Private Const COMMAND_GETDETAILERROR As Long = 5
Private Const COMMAND_SETSKEWADJMODE As Long = 6
Private Const COMMAND_SETCOUNTERMODE As Long = 7
Private Const COMMAND_SETEXPANDDATA As Long = 8
Private Const COMMAND_GETMODEDATA As Long = 9
Private Const COMMAND_SETFRAMESKIPMODE As Long = 10                 'add H.Miyamoto 2009/3/24
Private Const COMMAND_SETCHECKERATTRIBUTE As Long = 13
Private Const COMMAND_SETDECODETIMEOUT As Long = 14                 'add T.Minoda 2009/06/29
Private Const VDCOMMAND_SETCHECKERATTRIBUTE As Long = 8
Private Const VDCOMMAND_SETCMOSATTRIBUTEMODE As Long = 9
Private Const COMMAND_VIRTULDEVISE_CTRL As Long = 12
Private Const VDCOMMAND_SETEXPANDDATA As Long = 13

Private Const SUBCOMMAND_SETCMOSBITORDERMODE As Long = 1
Private Const SUBCOMMAND_SETCMOSPIXELSKIPMODE As Long = 2
Private Const SUBCOMMAND_SETCMOSLATCHSW_DCK As Long = 3
Private Const SUBCOMMAND_SETCMOSLATCHSW_XVS As Long = 4
Private Const SUBCOMMAND_SETCMOSLATCHSW_XHS As Long = 5
Private Const SUBCOMMAND_SETPIXELNUM_MASTER As Long = 6
Private Const SUBCOMMAND_SETLINENUM_MASTER As Long = 7
Private Const SUBCOMMAND_SELTHROUGHDATABIT As Long = 8
Private Const SUBCOMMAND_SELCMOS_SDRDDR As Long = 9                 'add T.Minoda 2009/09/17
Private Const SUBCOMMAND_SELCMOSDATAASSIGN As Long = 10             'add T.Minoda 2009/09/17
Private Const SUBCOMMAND_SETSKEW_MIPI As Long = 1
Private Const SUBCOMMAND_SETSKEW_SMIA As Long = 2
Private Const SUBCOMMAND_SETSKEW_CMOS_A As Long = 3
Private Const SUBCOMMAND_SETSKEW_CMOS_B As Long = 4
Private Const SUBCOMMAND_SETSKEW_CMOS_C As Long = 5
Private Const SUBCOMMAND_SETCHECKER_SITE As Long = 1
Private Const SUBCOMMAND_SETCHECKER_PATTERN As Long = 2
Private Const SUBCOMMAND_SETLVDS_SKIPMODE As Long = 3
Private Const SUBCOMMAND_SETCHECKER_CRCERROR As Long = 4
Private Const SUBCOMMAND_SETCHECKER_PATTERNSIZE As Long = 5
Private Const SUBCOMMAND_SELFRAMESTART_SYNCCODE As Long = 3
Private Const SUBCOMMAND_SELCRCOPERATION As Long = 4
Private Const SUBCOMMAND_SELECCOPERATION As Long = 5
Private Const SUBCOMMAND_SELMIPICLKRFSW_CTRL As Long = 6
Private Const SUBCOMMAND_SELMIPIDATARFSW_CTRL As Long = 7
Private Const SUBCOMMAND_SELMIPIDATALANE_ASSIGN As Long = 8         'add H.Miyamoto 2009/3/24
Private Const SUBCOMMAND_SELSMIA_ASSIGN As Long = 9                 'add T.Minoda 2009/09/17
Private Const SUBCOMMAND_SETLVDS_PNREVERSE As Long = 10             'add T.Minoda 2009/09/17
Private Const SUBCOMMAND_SETMIPI_RAWMODE As Long = 12               'add Y.Kimura 2011/03/23
Private Const SUBCOMMAND_COUNT_SET As Long = 1
Private Const SUBCOMMAND_COUNT_READ_RESAULT As Long = 2
Private Const SUBCOMMAND_SETCHECKER_STRAT_RESET As Long = 1
Private Const SUBCOMMAND_SETCHECKER_LVDS_ON_OFF As Long = 2
Private Const SUBCOMMAND_SETCHECKER_CMOS_ON_OFF As Long = 3
Private Const SUBCOMMAND_SETTIMEOUT_TIME As Long = 1                'add T.Minoda 2009/06/29
Private Const SUBCOMMAND_SETTIMEOUT_VPULSE As Long = 2              'add T.Minoda 2009/06/29
Private Const SUBCOMMAND_SETTIMEOUT_HPULSE As Long = 3              'add T.Minoda 2009/06/29


Private Const DEFAULT_STATUS_PINNAME As String = "Status"
Private Const DEFAULT_PARAM0Bit_PINNAME As String = "PARAM0"
Private Const DEFAULT_PARAM1Bit_PINNAME As String = "PARAM1"
Private Const DEFAULT_PARAM2Bit_PINNAME As String = "PARAM2"
Private Const DEFAULT_PARAM3Bit_PINNAME As String = "PARAM3"
Private Const DEFAULT_MODE_PINNAME As String = "MODE"
Private Const DEFAULT_STROB_PINNAME As String = "STROB"
Private Const DEFAULT_ENABLE_PINNAME As String = "ENABLE"
Private Const DEFAULT_RESET_PINNAME As String = "RST"
Private Const DEFAULT_STATUS_PINNAMEFORVIRTUALFPGA As String = "StatusVirtualFpag"
Private Const DEFAULT_PARAM0Bit_PINNAMEFORVIRTUALFPGA As String = "PARAM0VirtualFpag"
Private Const DEFAULT_PARAM1Bit_PINNAMEFORVIRTUALFPGA As String = "PARAM1VirtualFpag"
Private Const DEFAULT_PARAM2Bit_PINNAMEFORVIRTUALFPGA As String = "PARAM2VirtualFpag"
Private Const DEFAULT_PARAM3Bit_PINNAMEFORVIRTUALFPGA As String = "PARAM3VirtualFpag"
Private Const DEFAULT_MODE_PINNAMEFORVIRTUALFPGA As String = "MODEVirtualFpag"
Private Const DEFAULT_STROB_PINNAMEFORVIRTUALFPGA As String = "STROBVirtualFpag"
Private Const DEFAULT_ENABLE_PINNAMEFORVIRTUALFPGA As String = "ENABLEVirtualFpag"
Private Const DEFAULT_RESET_PINNAMEFORVIRTUALFPGA As String = "RSTVirtualFpag"
Private Const DEFAULT_SAMPLINGCOUNT_STATUSMEAS As Long = 3
Private Const DEFAULT_LIMIT_STATUSLEVEL As Double = 1.5
Private Const DEFAULT_SETTING_TIME_PARAM As Double = 0.00001        '10us
Private Const DEFAULT_UTIL_SLEW_RATE As Double = 0.00001            '10uS
Private Const DEFAULT_STROBPULS_WIDTH As Double = 0.0001            '100uS
Private Const DEFAULT_RSTPULS_WIDTH As Double = 0.0001              '100uS
Private Const DEFAULT_ENBPULS_WIDTH As Double = 0.0001              '100uS
Private Const LOGING_FILENAME As String = "runlog.txt"
Private Const LOGING_DIR As String = "tmp"
'Private Const LOGING_FILEMAXSIZE As Long = 10485760
'Private Const LOGING_FILEMAXCNT As Long = 10
Private Const DEFAULT_LOGING_FILEMAXSIZE As Long = 10485760         '10M
Private Const DEFAULT_LOGING_FILEMAXCNT As Long = 10

Private Const FPGA_MINOR_VERSION_INDEX As Long = 4
Private Const FPGA_MAJOR_VERSION_INDEX As Long = 5

Private Const CLASS_VERSION As String = "1.2"                      'change Y.Kimura 2011/04/01

'+++ Define IDelay Tap Value +++++
'Default
''Private Const DEFAULT_TAPVALUE_MIPIDATA As Long = 12
''Private Const DEFAULT_TAPVALUE_MIPICLK As Long = 0
''Private Const DEFAULT_TAPVALUE_SMIADATA As Long = 2
''Private Const DEFAULT_TAPVALUE_SMIASTRB As Long = 0
''Private Const DEFAULT_TAPVALUE_CMOSDATA1 As Long = 0
''Private Const DEFAULT_TAPVALUE_CMOSDATA2 As Long = 0
''Private Const DEFAULT_TAPVALUE_CMOSDATA3 As Long = 0
''Private Const DEFAULT_TAPVALUE_CMOSXVHS As Long = 20
''Private Const DEFAULT_TAPVALUE_CMOSDCK As Long = 20

Public DEFAULT_TAPVALUE_MIPIDATA As Long
Public DEFAULT_TAPVALUE_MIPICLK As Long
Public DEFAULT_TAPVALUE_SMIADATA As Long
Public DEFAULT_TAPVALUE_SMIASTRB As Long
Public DEFAULT_TAPVALUE_CMOSDATA1 As Long
Public DEFAULT_TAPVALUE_CMOSDATA2 As Long
Public DEFAULT_TAPVALUE_CMOSDATA3 As Long
Public DEFAULT_TAPVALUE_CMOSXVHS As Long
Public DEFAULT_TAPVALUE_CMOSDCK As Long



'+++ Define Enum Type +++++
Public Enum SIGNAL_OUT_MODE
    SO_ON = 1
    SO_OFF = 2
End Enum

Public Enum VIRTULDEV_ENABLE
    VE_RESET = 1
    VE_START = 2
End Enum

Public Enum LOGIC_STATE
    LS_LOW = 0
    LS_HIGH = 1
End Enum

Public Enum CAP_STATE
    CS_SUCCESS = 0
    CS_ERROR = 1
End Enum

Public Enum COUNTER_MODE
''    CM_PIXEL_COUNT = 1
''    CM_PIXEL_COUNT_OF_HLINE = 2
''    CM_H_COUNT_OF_VLINE = 3
''    CM_F_COUNT = 4
    CM_E1PIXEL_COUNT_OF_HLINE = 5
    CM_E2PIXEL_COUNT_OF_HLINE = 6
    CM_EVLINE_INDEX_OF_ERR4 = 7
    CM_CRCERR_NUM_OF_LINE = 8
    CM_CRCERR_NUM_OF_FRAME = 9
    CM_ECCERR_NUM_OF_LINE = 10
    CM_ECCERR_NUM_OF_FRAME = 11
End Enum

Public Enum CAPTURE_MODE
    CM_MIPI_1LANE_RAW10 = 1
    CM_MIPI_2LANE_RAW10 = 2
    CM_SMIA_1CH_RAW8_COMPRESSION = 3
    CM_SMIA_1CH_RAW10 = 4
    CM_SMIA_2CH_RAW10 = 5
    CM_CMOS_MASTERMODE_P10BIT_1CH = 6
    CM_CMOS_MASTERMODE_P10BIT_2CH = 7
    CM_CMOS_THROUGH = 8
    CM_MIPI_3LANE_RAW10 = 9             'add H.miyamoto 2009/03/24
    CM_MIPI_4LANE_RAW10 = 10            'add H.miyamoto 2009/03/24
    CM_CMOS_MASTERMODE_P12BIT_1CH = 11  'add T.Minoda 2009/09/17
End Enum

Public Enum CAPTURE_MODE_VD
    CM_1CH_COMPRESSION_RAW8 = 1
    CM_1CH_RAW8_NONSMIA = 2
    CM_1CH_RAW10_NONSMIA = 3
    CM_2CH_RAW10_NONSMIA = 4
    CM_CMOS_OUT = 5
    CM_CRC_1CH_COMPRESSION_RAW8 = 6
    CM_CRC_1CH_RAW8_NONSMIA = 7
    CM_CRC_1CH_RAW10_NONSMIA = 8
    CM_CRC_2CH_RAW10_NONSMIA = 9
    CM_CMOS_BYPASS = 10
    CM_CMOS_OUT_MASTERMODE = 11
    CM_2CH_RAW10_CLASS2 = 12
    CM_CRC_2CH_RAW10_CLASS2 = 13
End Enum

Public Enum CLOCK_MODE
    CLM_DATA_STROB = 1
    CLM_DATA_CLOCK_COM = 2
    CLM_DATA_CLOCK_SEP = 3
End Enum

Public Enum LVDS_SKIP_MODE
    LS_NON_SKIP = 1
    LS_1_2_SKIP = 2
    LS_1_4_SKIP = 3
End Enum

Public Enum AVENUM_MODE
    AV_1 = 0
    AV_2 = 1
    AV_4 = 2
    AV_6 = 3
    AV_8 = 4
    AV_12 = 5
    AV_16 = 6
    AV_24 = 7
    AV_32 = 8
    AV_48 = 9
    AV_64 = 10
    AV_96 = 11
    AV_128 = 12
    AV_192 = 13
    AV_256 = 14
    AV_511 = 15
End Enum

Public Enum CHECKER_PATTERN
    CP_PATTERN1 = 1
    CP_PATTERN2 = 2
    CP_PATTERN3 = 3
    CP_PATTERN4 = 4
    CP_PATTERN5 = 5
    CP_PATTERN6 = 6
    CP_PATTERN7 = 7
    CP_PATTERN8 = 8
End Enum

Public Enum READ_MODE
    RM_CAPTUREMODE = 1
    RM_AVEMODE = 2
    RM_CLOCK_MODE = 3
    RM_FPGA_MINOR_VERSION = 4
    RM_MAJOR_VERSION = 5
    RM_BUFFER_MODE = 6
'    RM_EXPAND_NUM = 8
End Enum

Public Enum CMOS_BITORDER_MODE
    CB_LSB_MSB_NONREVERSE = 1
    CB_LSB_MSB_REVERSE = 2
End Enum

Public Enum CMOS_LATCH_CH
    CL_DCK = 3
    CL_XVS = 4
    CL_XHS = 5
    CL_DCK_XVS_XHS = 10
End Enum

Public Enum CMOS_LATCH_MODE
    CL_FALL_LATCH = 1
    CL_RISE_LATCH = 2
End Enum

Public Enum CMOS_DATARATE               'add T.Minoda 2009/09/17
    CD_SDR = 1
    CD_DDR = 2
End Enum

Public Enum CMTH_INPUT_DATABIT
    CT_DATA0_9 = 1
    CT_DATA10_19 = 2
    CT_DATA20_29 = 3
End Enum

Public Enum FPGA_MODE
    FM_DECODE_FPGA = 0
    FM_VIRTUAL_FPGA = 1
    FM_BOTH_FPGA = 2
End Enum

Public Enum LOGGINGOUT_MODE
    LM_ERROR_ONLY = 0
    LM_SUCCESS_ONLY = 1
    LM_BOTH = 2
End Enum

Public Enum LOGGINGOUT_LOCATION
    LL_JOB_DIRECTORY = 0
    LL_SYSTEM_TEMP_DIRECTORY = 1
End Enum

Public Enum TAPMODE
    TAP_RST = 0
    MIPI = 1
    SMIA = 2
    CMOS_10 = 3
    CMOS_20 = 4
End Enum

Public Enum DELAYSET_MIPICH
    MIPI_DATA_BOTH = 0
    MIPI_DATA_1 = 1
    MIPI_DATA_2 = 2
    MIPI_CLK = 3
    MIPI_DATA_3 = 4                 'add H.Miyamoto 2009/3/24
    MIPI_DATA_4 = 5                 'add H.Miyamoto 2009/3/24
End Enum

Public Enum DELAYSET_SMIACH
    SMIA_DATA_BOTH = 0
    SMIA_DATA_1 = 1
    SMIA_DATA_2 = 2
    SMIA_STRB_BOTH = 3
    SMIA_STRB_1 = 4
    SMIA_STRB_2 = 5
    SMIA_CLK_BOTH = 6
    SMIA_CLK_1 = 7
    SMIA_CLK_2 = 8
End Enum

Public Enum DELAYSET_CMOSCH
    CMOS_DATA_0 = 0
    CMOS_DATA_1 = 1
    CMOS_DATA_2 = 2
    CMOS_DATA_3 = 3
    CMOS_DATA_4 = 4
    CMOS_DATA_5 = 5
    CMOS_DATA_6 = 6
    CMOS_DATA_7 = 7
    CMOS_DATA_8 = 8
    CMOS_DATA_9 = 9
    CMOS_DATA_10 = 10
    CMOS_DATA_11 = 11
    CMOS_DATA_12 = 12
    CMOS_DATA_13 = 13
    CMOS_DATA_14 = 14
    CMOS_DATA_15 = 15
    CMOS_DATA_16 = 16
    CMOS_DATA_17 = 17
    CMOS_DATA_18 = 18
    CMOS_DATA_19 = 19
    CMOS_DATA_20 = 20
    CMOS_DATA_21 = 21
    CMOS_DATA_22 = 22
    CMOS_DATA_23 = 23
    CMOS_DATA_24 = 24
    CMOS_DATA_25 = 25
    CMOS_DATA_26 = 26
    CMOS_DATA_27 = 27
    CMOS_DATA_28 = 28
    CMOS_DATA_29 = 29
    CMOS_XVS = 30
    CMOS_XHS = 31
    CMOS_DCK = 32
    CMOS_DATA_0_9 = 33
    CMOS_DATA_10_19 = 34
    CMOS_DATA_20_29 = 35
    CMOS_XVS_XHS = 36
End Enum

Public Enum EXDATA_CHAN
    EXPANDDATA_SET_CH1 = 1
    EXPANDDATA_SET_CH2 = 2
End Enum

Public Enum EXDATA_NUM
    EN_0 = 0
    EN_2 = 1
    EN_4 = 2
    EN_6 = 3
    EN_8 = 4
    EN_10 = 5
    EN_12 = 6
    EN_14 = 7
    EN_16 = 8
End Enum

Public Enum CRC_ERROR_MODE
    CRC_ERROR_OFF = 1
    CRC_ERROR_ON = 2
End Enum

Public Enum CRC_ONOFF
    CRC_ON = 1
    CRC_OFF = 2
End Enum

Public Enum ECC_ONOFF
    ECC_ON = 1
    ECC_OFF = 2
End Enum

Public Enum MIPICLK_RFSW_CTRL
    RC_LP_STATE = 1
    RC_ALL_HS = 2
    RC_ALL_LP = 3
End Enum

Public Enum MIPIDATA_RFSW_CTRL
    RD_LP_STATE = 1
    RD_ALL_HS = 2
    RD_ALL_LP = 3
End Enum

Public Enum FRAME_START_CODE
    FS_CODE_SOD = 1
    FS_CODE_SOF = 2
End Enum

Public Enum PATTERN_SIZE
    H2664xV1992 = 1
    H2664xV4928 = 2
    H6304xV2400 = 3
End Enum

Private Type LASTERRORINFO
    lngErrorNum As Long
    strErrorName As String
    strErrorFunction As String
    strErrortmp As String
End Type

Public Enum MIPI_DATALANE_ASSIGN    'add H.Miyamoto 2009/3/24
    LANE_1234_Default = 0
    LANE_1243 = 1
    LANE_1324 = 2
    LANE_1342 = 3
    LANE_1423 = 4
    LANE_1432 = 5
    LANE_2134 = 6
    LANE_2143 = 7
    LANE_2314 = 8
    LANE_2341 = 9
    LANE_2413 = 10
    LANE_2431 = 11
    LANE_3124 = 12
    LANE_3142 = 13
    LANE_3214 = 14
    LANE_3241 = 15
    LANE_3412 = 16
    LANE_3421 = 17
    LANE_4123 = 18
    LANE_4132 = 19
    LANE_4213 = 20
    LANE_4231 = 21
    LANE_4312 = 22
    LANE_4321 = 23
End Enum

Public Enum SMIA_ASSIGN             'add T.Minoda 2009/09/17
    SMLINE_D1S1S2D2_Default = 0
    SMLINE_D2S2S1D1 = 1
    MPLINE_D1S1S2D2 = 2
    MPLINE_D2S2S1D1 = 3
End Enum

Public Enum CMOS_DATA_ASSIGN        'add T.Minoda 2009/09/17
    CASSIGN_0_Default = 0
    CASSIGN_1 = 1
    CASSIGN_2 = 2
    CASSIGN_3 = 3
End Enum

Public Enum LVDS_PN_REVERSE         'add T.Minoda 2009/09/17
    LN_PN_NONREVERSE = 1            'correct T.Minoda 2009/10/06
    LN_PN_REVERSE = 2
End Enum

Public Enum FRAMESKIP_MODE          'add H.Miyamoto 2009/3/24
    FS_SKIP_0 = 0
    FS_SKIP_1 = 1
    FS_SKIP_2 = 2
    FS_SKIP_3 = 3
    FS_SKIP_4 = 4
    FS_SKIP_5 = 5
    FS_SKIP_6 = 6
    FS_SKIP_7 = 7
    FS_SKIP_8 = 8
    FS_SKIP_9 = 9
    FS_SKIP_10 = 10
    FS_SKIP_11 = 11
    FS_SKIP_12 = 12
    FS_SKIP_13 = 13
    FS_SKIP_14 = 14
    FS_SKIP_15 = 15
End Enum

'add T.Minoda 2009/06/29
Public Enum TIMEOUT
    TO_0SEC = 0
    TO_1SEC = 1
    TO_2SEC = 2
    TO_3SEC = 3
    TO_4SEC = 4
    TO_6SEC = 5
    TO_8SEC = 6
    TO_10SEC = 7
    TO_12SEC = 8
    TO_14SEC = 9
    TO_16SEC = 10
    TO_18SEC = 11
    TO_30SEC = 12                   'change T.Minoda 2009/09/17
    TO_60SEC = 13                   'change T.Minoda 2009/09/17
    TO_180SEC = 14                  'change T.Minoda 2009/09/17
    TO_300SEC = 15                  'change T.Minoda 2009/09/17
End Enum

'add T.Minoda 2009/06/29
Public Enum DUMMYPULSE
    DP_1_2E1 = 0
    DP_1_2E2 = 1
    DP_1_2E3 = 2
    DP_1_2E4 = 3
    DP_1_2E5 = 4
    DP_1_2E6 = 5
    DP_1_2E7 = 6
    DP_1_2E8 = 7
    DP_1_2E9 = 8
    DP_1_2E10 = 9
    DP_1_2E11 = 10
    DP_1_2E12 = 11
    DP_1_2E13 = 12
    DP_1_2E14 = 13
    DP_1_2E15 = 14
    DP_1_2E16 = 15
End Enum

'add Y.Kimura 2011/03/23
Public Enum mipi_rawmode
    MIPI_RAWMODE_RAW10 = 1
    MIPI_RAWMODE_RAW12 = 2
End Enum

'+++ Define Member Variable +++++
Private m_strStatusPinName As String
Private m_strParamPinName(3) As String
Private m_strModePinName As String
Private m_strStrobPinName As String
Private m_strEnablePinName As String
Private m_strResetPinName As String
Private m_strStatusPinNameForVirtualFpga As String
Private m_strParamPinNameForVirtualFpga(3) As String
Private m_strModePinNameForVirtualFpga As String
Private m_strStrobPinNameForVirtualFpga As String
Private m_strEnablePinNameForVirtualFpga As String
Private m_strResetPinNameForVirtualFpga As String
Private m_lngSampCountOfStatusMeas As Long
Private m_dblLimitOfStatusLev As Double
Private m_eLogicState() As LOGIC_STATE
Private m_eLogicState_Vir() As LOGIC_STATE
Private m_dblSettingTimeOfParam As Double
Private m_dblStrobPulsWidth As Double
Private m_dblRstPulsWidth As Double
Private m_dblEnbPulsWidth As Double
Private m_strLastErrorMsg As String
Private m_strLastErrorTmpMsg As String
Private m_dblSlewRateOfUtility As Double
Private m_eCaptureMode As CAPTURE_MODE
Private m_eFSkipMode As FRAMESKIP_MODE          'add H.Miyamoto 2009/3/24
Private m_sLastErrorInfo As LASTERRORINFO
Private m_eAveNumber As AVENUM_MODE
Private m_eClockMode As CLOCK_MODE
Private m_strUtilityPins() As String
Private m_strPins() As String
Private m_blPinsTableCreadedFlg As Boolean
Private m_eEnable As LOGIC_STATE
Private m_eMode As LOGIC_STATE
Private m_eReset As LOGIC_STATE
Private m_eStrob As LOGIC_STATE
Private m_blVirtualFpgaEnabled As Boolean
Private m_strLogDirPath As String
Private m_strLogFilePath As String
Private m_blLoggingEnabled As Boolean
Private m_eLoggingMode As LOGGINGOUT_MODE
Private m_strFunction As String
Private m_blRunTimeErrorEnable As Boolean
Private m_eLoggingLocation As LOGGINGOUT_LOCATION
Private m_lngLoggingFileSize As Long
Private m_lngLoggingFileNum As Long
Private m_lngVirEnabSite As Long
Private m_lngTapMode As Long
Private m_blTimeOutEnable As Boolean            'add T.Minoda 2009/06/29
Private m_dblUserDecTimeOut As Double           'add T.Minoda 2009/06/29
Private m_eDummyVpulse As DUMMYPULSE            'add T.Minoda 2009/06/29
Private m_eDummyHpulse As DUMMYPULSE            'add T.Minoda 2009/06/29
Private m_blInitTapFlag As Boolean              'add T.Minoda 2009/09/17
Private m_blSMIACapture_ignore_Revision As Boolean 'add Y.Kimura 2011/03/23






'==========================================================================
'
'
'               Public Property
'
'
'==========================================================================

Public Property Get SwVersion() As String

    SwVersion = CLASS_VERSION

End Property

Public Property Get CaptureMode() As CAPTURE_MODE

    CaptureMode = m_eCaptureMode

End Property

Public Property Let RunTimeErrorEnable(ByVal blRunTimeErrorEnable As Boolean)

    m_blRunTimeErrorEnable = blRunTimeErrorEnable

End Property

Public Property Let LoggingFileSize(ByVal lngLoggingFileSize As Long)

    m_lngLoggingFileSize = lngLoggingFileSize

End Property

Public Property Let LoggingFileNum(ByVal lngLoggingFileNum As Long)

    m_lngLoggingFileNum = lngLoggingFileNum

End Property

Public Property Let LoggingMode(ByVal eLoggingMode As LOGGINGOUT_MODE)

    m_eLoggingMode = eLoggingMode

End Property

Public Property Let LoggingLocation(ByVal eLoggingLocation As LOGGINGOUT_LOCATION)

    m_eLoggingLocation = eLoggingLocation
    
    Call LogFileInit
    
End Property

Public Property Let LoggingEnabled(ByVal blLoggingEnabled As Boolean)
    
    m_blLoggingEnabled = blLoggingEnabled
    
End Property

Public Property Let VirtualFpgaEnabled(ByVal blVirtualFpgaEnabled As Boolean)

    m_blVirtualFpgaEnabled = blVirtualFpgaEnabled
    
End Property

Public Property Let StatusPinName(ByVal strStatusPinName As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If IsDchPin(strStatusPinName) Then
        m_strStatusPinName = strStatusPinName
    Else
        MsgBox strStatusPinName & " is not Dch Pin!"
    End If
End Property

Public Property Let ModePinName(ByVal strModePinName As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If IsUtilityPin(strModePinName) Then
        m_strModePinName = strModePinName
    Else
        MsgBox strModePinName & " is not Utility Pin!"
    End If
End Property

Public Property Let ResetPinName(ByVal strResetPinName As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If IsUtilityPin(strResetPinName) Then
        m_strResetPinName = strResetPinName
    Else
        MsgBox strResetPinName & " is not Utility Pin!"
    End If
End Property

Public Property Let EnablePinName(ByVal strEnablePinName As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If IsUtilityPin(strEnablePinName) Then
        m_strEnablePinName = strEnablePinName
    Else
        MsgBox strEnablePinName & " is not Utility Pin!"
    End If
End Property

Public Property Let StrobPinName(ByVal strStrobPinName As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If IsUtilityPin(strStrobPinName) Then
        m_strStrobPinName = strStrobPinName
    Else
        MsgBox strStrobPinName & " is not Utility Pin!"
    End If
End Property

Public Property Let ParamPinName(lngIndex As Long, ByVal strParamPinName As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If lngIndex >= 0 And lngIndex < 4 Then
        If IsUtilityPin(strParamPinName) Then
            m_strParamPinName(lngIndex) = strParamPinName
        Else
            MsgBox strParamPinName & " is not Utility Pin!"
        End If
    End If
End Property

Public Property Let StatusPinNameofVirtualDevice(ByVal strStatusPinNameForVirtualFpga As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If IsDchPin(strStatusPinNameForVirtualFpga) Then
        m_strStatusPinNameForVirtualFpga = strStatusPinNameForVirtualFpga
    Else
        MsgBox strStatusPinNameForVirtualFpga & " is not Dch Pin!"
    End If
End Property

Public Property Let ModePinNameOfVirtualDevice(ByVal strModePinNameForVirtualFpga As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If IsUtilityPin(strModePinNameForVirtualFpga) Then
        m_strModePinNameForVirtualFpga = strModePinNameForVirtualFpga
    Else
        MsgBox strModePinNameForVirtualFpga & " is not Utility Pin!"
    End If
End Property

Public Property Let ResetPinNameOfVirtualDevice(ByVal strResetPinNameForVirtualFpga As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If IsUtilityPin(strResetPinNameForVirtualFpga) Then
        m_strResetPinNameForVirtualFpga = strResetPinNameForVirtualFpga
    Else
        MsgBox strResetPinNameForVirtualFpga & " is not Utility Pin!"
    End If
End Property

Public Property Let EnablePinNameOfVirtualDevice(ByVal strEnablePinNameForVirtualFpga As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If IsUtilityPin(strEnablePinNameForVirtualFpga) Then
        m_strEnablePinNameForVirtualFpga = strEnablePinNameForVirtualFpga
    Else
        MsgBox strEnablePinNameForVirtualFpga & " is not Utility Pin!"
    End If
End Property

Public Property Let StrobPinNameOfVirtualDevice(ByVal strStrobPinNameForVirtualFpga As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If IsUtilityPin(strStrobPinNameForVirtualFpga) Then
        m_strStrobPinNameForVirtualFpga = strStrobPinNameForVirtualFpga
    Else
        MsgBox strStrobPinNameForVirtualFpga & " is not Utility Pin!"
    End If
End Property

Public Property Let ParamPinNameOfVirtualDevice(lngIndex As Long, ByVal strParamPinNameForVirtualFpga As String)
    
    If m_blPinsTableCreadedFlg = False Then
        Call InitVariable
        m_blPinsTableCreadedFlg = True
    End If
    
    If lngIndex >= 0 And lngIndex < 4 Then
        If IsUtilityPin(strParamPinNameForVirtualFpga) Then
            m_strParamPinNameForVirtualFpga(lngIndex) = strParamPinNameForVirtualFpga
        Else
            MsgBox strParamPinNameForVirtualFpga & " is not Utility Pin!"
        End If
    End If
End Property

'add T.Minoda 2009/06/29
Public Property Let DecodeTimeOutEnable(ByVal blTimeOutEnable As Boolean)

    m_blTimeOutEnable = blTimeOutEnable
    
End Property

'add T.Minoda 2009/06/29
Public Property Let DecodeTimeOut(ByVal dblUserDecTimeOut As Double)

    If dblUserDecTimeOut < 0 Then
        MsgBox "Please specify correct time out time! [dblUserDecTimeOut>=0]"
        Exit Property
    End If
    m_dblUserDecTimeOut = dblUserDecTimeOut
    
End Property

'add T.Minoda 2009/06/29
Public Property Let DummyVsyncPulse(ByVal eDummyVpulse As DUMMYPULSE)

    m_eDummyVpulse = eDummyVpulse
    
End Property

'add T.Minoda 2009/06/29
Public Property Let DummyHsyncPulse(ByVal eDummyHpulse As DUMMYPULSE)

    m_eDummyHpulse = eDummyHpulse
    
End Property

'add T.Minoda 2009/09/17
Public Property Let InitTapFlag(ByVal blInitTapFlag As Boolean)

    m_blInitTapFlag = blInitTapFlag
    
End Property





'==========================================================================
'
'
'               Public Method
'
'
'==========================================================================


Public Function Reset(Optional strSMIAignoreVersion As String = "2.1") As Boolean

    Dim strFpgaVersion() As String   ' add Y.Kimura 2011/04/01
    
    LoggingSection = "Reset"
    
    Rst = LS_LOW
    TheHdw.WAIT m_dblRstPulsWidth
    Strob = LS_LOW
    Enable = LS_LOW
    TheHdw.WAIT m_dblRstPulsWidth
    Rst = LS_HIGH
    TheHdw.WAIT m_dblRstPulsWidth
    Rst = LS_LOW
    
    If m_blVirtualFpgaEnabled = True Then
        RstVD = LS_LOW
        TheHdw.WAIT m_dblRstPulsWidth
        StrobVD = LS_LOW
        EnableVD = LS_LOW
        TheHdw.WAIT m_dblRstPulsWidth
        RstVD = LS_HIGH
        TheHdw.WAIT m_dblRstPulsWidth
        RstVD = LS_LOW
    End If
    
    If WaitReady(FM_DECODE_FPGA) = True Then
        Call SetDefaultMode
        Call ClearLastInfo
        If m_blInitTapFlag = False Then     'add T.Minoda 2009/09/17
            m_blSMIACapture_ignore_Revision = False   ' add Y.Kimura 2011/04/01
            Call InitTapValue(TAP_RST)
            Call GetFpgaVersion(strFpgaVersion, , strSMIAignoreVersion) ' add Y.Kimura 2011/04/01
            m_blInitTapFlag = True
        End If
        Reset = True
    Else
        Call SetLastError(1, "Reset")
        Reset = False
    End If
    
End Function


Public Sub DecodeStart()
    
    LoggingSection = "DecodeStart"
    
    SetDecTimeOut       'add T.Minoda 2009/06/29
    
    Enable = LS_LOW
    TheHdw.WAIT m_dblEnbPulsWidth
    Enable = LS_HIGH
        
End Sub


Public Sub DecodeStop()

    Enable = LS_LOW
    
End Sub


Public Sub InitTapValue(eTapMode As TAPMODE)

    Select Case eTapMode
        Case TAP_RST
            SetDelayTapForMIPI MIPI_DATA_BOTH, DEFAULT_TAPVALUE_MIPIDATA
            SetDelayTapForMIPI MIPI_CLK, DEFAULT_TAPVALUE_MIPICLK
            SetDelayTapForSMIA SMIA_DATA_BOTH, DEFAULT_TAPVALUE_SMIADATA
            SetDelayTapForSMIA SMIA_STRB_BOTH, DEFAULT_TAPVALUE_SMIASTRB
            SetDelayTapForCMOS CMOS_DATA_0_9, DEFAULT_TAPVALUE_CMOSDATA1
            SetDelayTapForCMOS CMOS_DATA_10_19, DEFAULT_TAPVALUE_CMOSDATA2
            SetDelayTapForCMOS CMOS_DATA_20_29, DEFAULT_TAPVALUE_CMOSDATA3
            SetDelayTapForCMOS CMOS_XVS_XHS, DEFAULT_TAPVALUE_CMOSXVHS
            SetDelayTapForCMOS CMOS_DCK, DEFAULT_TAPVALUE_CMOSDCK
        
        Case MIPI
            SetDelayTapForMIPI MIPI_DATA_BOTH, DEFAULT_TAPVALUE_MIPIDATA
            SetDelayTapForMIPI MIPI_CLK, DEFAULT_TAPVALUE_MIPICLK
            
        Case SMIA
            SetDelayTapForSMIA SMIA_DATA_BOTH, DEFAULT_TAPVALUE_SMIADATA
            SetDelayTapForSMIA SMIA_STRB_BOTH, DEFAULT_TAPVALUE_SMIASTRB
            
        Case CMOS_10
            SetDelayTapForCMOS CMOS_DATA_0_9, DEFAULT_TAPVALUE_CMOSDATA1
            SetDelayTapForCMOS CMOS_XVS_XHS, DEFAULT_TAPVALUE_CMOSXVHS
            SetDelayTapForCMOS CMOS_DCK, DEFAULT_TAPVALUE_CMOSDCK
            
        Case CMOS_20
            SetDelayTapForCMOS CMOS_DATA_10_19, DEFAULT_TAPVALUE_CMOSDATA2
            SetDelayTapForCMOS CMOS_DATA_20_29, DEFAULT_TAPVALUE_CMOSDATA3
            SetDelayTapForCMOS CMOS_XVS_XHS, DEFAULT_TAPVALUE_CMOSXVHS
            SetDelayTapForCMOS CMOS_DCK, DEFAULT_TAPVALUE_CMOSDCK
            
        Case Else
            MsgBox "Missing : setting error at InitTapValue!!"
            Exit Sub
            
    End Select
    
    m_lngTapMode = eTapMode

End Sub


Public Sub Capture(Pins As String, Dest As String, FullPad As String, acqMode As IdpAcquireMode, _
                            Optional numFrame As Long = 1, _
                            Optional averagMode As IdpAverageMode = idpAverage, _
                            Optional NumFields As Long = 2)

    LoggingSection = "Capture"

    TheHdw.IDP.SetPMD Dest, FullPad
    TheHdw.IDP.Acquire Pins, Dest, 1, idpNonAverage, acqMode, NumFields

End Sub


Public Sub MultiCapture(Pins As String, Dest As String, FullPad As String, acqMode As IdpAcquireMode, acqIndex As Long, _
                            Optional numFrame As Long = 1, _
                            Optional averagMode As IdpAverageMode = idpAverage, _
                            Optional NumFields As Long = 2)
                            
    Dim cPins(0) As String
    Dim cDest(0) As String
    
    LoggingSection = "MultiCapture"
    
    cPins(0) = Pins
    cDest(0) = Dest
    
    TheHdw.IDP.SetPMD Dest, FullPad
    TheHdw.IDP.MultiCapture cPins, cDest, 1, 1, idpNonAverage, acqMode, acqIndex, NumFields
    
End Sub


Public Sub MultiCaptureEx(Pins As String, Dest As String, FullPad As String, acqMode As IdpAcquireMode, acqIndex As Long, _
                            Optional numFrame As Long = 1, _
                            Optional averagMode As IdpAverageMode = idpAverage, _
                            Optional NumFields As Long = 2)
                            
    Dim cPins(0) As String
    Dim cDest(0) As String
    
    LoggingSection = "MultiCaptureEx"
    
    cPins(0) = Pins
    cDest(0) = Dest
    
    TheHdw.IDP.SetPMD Dest, FullPad
    TheHdw.IDP.MultiCapture cPins, cDest, 1, 1, idpNonAverage, acqMode, acqIndex, NumFields
    
End Sub


Public Sub SetCaptureMode(eCaptureMode As CAPTURE_MODE)
    Dim lngBuf() As Long
    Dim lngAve() As Long
    Dim lngSiteNum As Long
    
    LoggingSection = "SetCaptureMode"
    
    ''' add Y.Kimura 2011/03/23
    If ((m_blSMIACapture_ignore_Revision = True) And _
        ((eCaptureMode = CM_SMIA_1CH_RAW8_COMPRESSION) Or _
         (eCaptureMode = CM_SMIA_1CH_RAW10) Or _
         (eCaptureMode = CM_SMIA_2CH_RAW10) _
        )) Then
       
        Call SetLastError(1, "SetCaptureMode:SMIA couldn't be captured by current FPGA Rev.!")
        MsgBox "SetCaptureMode:SMIA couldn't be captured by current FPGA Rev."
        Exit Sub
    End If
    ''' END OF add Y.Kimura 2011/03/23
    
    Enable = LS_LOW
    If WaitReady(FM_BOTH_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETCAPTUREMODE)
        If WaitReady(FM_BOTH_FPGA) = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(eCaptureMode))
            If WaitReady(FM_BOTH_FPGA) = True Then
                m_eCaptureMode = eCaptureMode           'delete InitTapValue T.Minoda 2009/09/17
                Exit Sub
            Else
                Call SetLastError(1, "SetCaptureMode:WaitReady After Sending Value is done")
            End If
        Else
            Call SetLastError(1, "SetCaptureMode:WaitReady After Sending Command is done")
        End If
    Else
        
        Call SetLastError(1, "SetCaptureMode:First WaitReady")
        
    End If

End Sub


Public Sub SetCaptureModeForVD(eCaptureMode As CAPTURE_MODE_VD)
    Dim lngBuf() As Long
    Dim lngAve() As Long
    Dim lngSiteNum As Long
    
    LoggingSection = "SetCaptureMode"
    
    EnableVD = LS_LOW
    If WaitReady(FM_VIRTUAL_FPGA) = True Then
    
        ModeVD = LS_HIGH
        ParamVD = Convert10to2(COMMAND_SETCAPTUREMODE)
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
            ModeVD = LS_LOW
            ParamVD = Convert10to2(CLng(eCaptureMode))
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                m_eCaptureMode = eCaptureMode
                Exit Sub
            Else
                Call SetLastError(1, "SetCaptureMode:WaitReady After Sending Value is done")
            End If
        Else
            Call SetLastError(1, "SetCaptureMode:WaitReady After Sending Command is done")
        End If
    Else
        
        Call SetLastError(1, "SetCaptureMode:First WaitReady")
        
    End If

End Sub


Public Sub SetAveragNum(eAveNum As AVENUM_MODE)
    Dim lngCap() As Long
    Dim lngSiteNum As Long
    
    LoggingSection = "SetAveragNum"
    
    Enable = LS_LOW
'    If m_eAveNumber <> eAveNum Then
    
        If WaitReady = True Then
        
            mode = LS_HIGH
            Param = Convert10to2(COMMAND_SETAVENUM)
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eAveNum))
                If WaitReady = True Then
                    m_eAveNumber = eAveNum
                    Exit Sub
                Else
                    Call SetLastError(1, "SetAveragNum:WaitReady After Sending Value is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetAveragNum:WaitReady After Sending Command is done")
                Exit Sub
            End If
        Else
            
            Call SetLastError(1, "SetAveragNum:First WaitReady")
            Exit Sub
        End If
'    Else
'        Exit Sub
'    End If

End Sub


Public Sub SetClockMode(eClockMod As CLOCK_MODE)
    
    LoggingSection = "SetClockMode"
    
    Enable = LS_LOW
'    If m_eClockMode <> eClockMod Then
    
        If WaitReady(FM_DECODE_FPGA) = True Then
        
            mode = LS_HIGH
            Param = Convert10to2(COMMAND_SETCLOCKMODE)
            If WaitReady(FM_DECODE_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eClockMod))
                If WaitReady(FM_DECODE_FPGA) = True Then
                    m_eClockMode = eClockMod
                    Exit Sub
                Else
                    Call SetLastError(1, "SetClockMode:WaitReady After Sending Value is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetClockMode:WaitReady After Sending Command is done")
                Exit Sub
            End If
            
        Else
            
            Call SetLastError(1, "SetClockMode:First WaitReady")
            Exit Sub
            
        End If
'    Else
'        Exit Sub
'    End If
    
End Sub


Public Sub SetClockModeForVD(eClockMod As CLOCK_MODE)
    
    LoggingSection = "SetClockMode"
    
    EnableVD = LS_LOW
'    If m_eClockMode <> eClockMod Then
    
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
        
            ModeVD = LS_HIGH
            ParamVD = Convert10to2(COMMAND_SETCLOCKMODE)
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                ModeVD = LS_LOW
                ParamVD = Convert10to2(CLng(eClockMod))
                If WaitReady(FM_VIRTUAL_FPGA) = True Then
                    m_eClockMode = eClockMod
                    Exit Sub
                Else
                    Call SetLastError(1, "SetClockMode:WaitReady After Sending Value is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetClockMode:WaitReady After Sending Command is done")
                Exit Sub
            End If
            
        Else
            
            Call SetLastError(1, "SetClockMode:First WaitReady")
            Exit Sub
            
        End If
'    Else
'        Exit Sub
'    End If
    
End Sub


Public Sub GetDetailErrorInfo(lngErrorword() As Long)
    Dim lngCnt As Long
    Dim eLogicState() As LOGIC_STATE
    Dim lngSiteNum As Long
    Dim strTmp As String
        
    ReDim lngErrorword(TheExec.sites.ExistingCount - 1)
    
    LoggingSection = "GetDetailErrorInfo"
    
    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_GETDETAILERROR)
        If WaitReady = True Then
            For lngCnt = 1 To 15
                mode = LS_LOW
                Param = Convert10to2(lngCnt)
                If WaitReady = False Then
                    Call SetLastError(1, "GetDetailErrorInfo:WaitReady After Addr[" & Trim(str(lngCnt)) & "] Command is done")
                    Exit Sub
                End If
                Call GetStatusLogic(eLogicState, m_strStatusPinName)
                For lngSiteNum = 0 To TheExec.sites.ExistingCount - 1
                    
                    If TheExec.sites.site(lngSiteNum).Active Then
                        
                        lngErrorword(lngSiteNum) = lngErrorword(lngSiteNum) Or IIf(eLogicState(lngSiteNum) = LS_HIGH, 2 ^ (lngCnt - 1), 0)
                        
                    End If
                    
                Next lngSiteNum
                
            Next lngCnt
            
            For lngSiteNum = 0 To TheExec.sites.ExistingCount - 1
                
                If TheExec.sites.site(lngSiteNum).Active Then
                    
                   If lngErrorword(lngSiteNum) <> 0 Then
                        strTmp = "ErrorInfo:" & Convert10to2(lngErrorword(lngSiteNum), 16)
                        Call SetLastError(2, "GetDetailErrorInfo", strTmp, lngSiteNum)
                   End If
                    
                End If
                
            Next lngSiteNum
            Exit Sub
        Else
            Call SetLastError(1, "GetDetailErrorInfo:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "GetDetailErrorInfo:First WaitReady")
        Exit Sub
        
    End If
        
End Sub


Public Sub GetDetailErrorInfoWithOutErrorRise(lngErrorword() As Long)

    Dim lngCnt As Long
    Dim eLogicState() As LOGIC_STATE
    Dim lngSiteNum As Long
    Dim strTmp As String
    
    
    
    ReDim lngErrorword(TheExec.sites.ExistingCount - 1)
    
    LoggingSection = "GetDetailErrorInfo"
    
    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_GETDETAILERROR)
        If WaitReady = True Then
            For lngCnt = 1 To 15
                mode = LS_LOW
                Param = Convert10to2(lngCnt)
                If WaitReady = False Then
                    Call SetLastError(1, "GetDetailErrorInfoWithOutErrorRise:WaitReady After Addr[" & Trim(str(lngCnt)) & "] Command is done")
                    Exit Sub
                End If
                Call GetStatusLogic(eLogicState, m_strStatusPinName)
                For lngSiteNum = 0 To TheExec.sites.ExistingCount - 1
                    
                    If TheExec.sites.site(lngSiteNum).Active Then
                        
                        lngErrorword(lngSiteNum) = lngErrorword(lngSiteNum) Or IIf(eLogicState(lngSiteNum) = LS_HIGH, 2 ^ (lngCnt - 1), 0)
                        
                    End If
                    
                Next lngSiteNum
                
            Next lngCnt
            
'            For lngSiteNum = 0 To TheExec.Sites.ExistingCount - 1
'
'                If TheExec.Sites.Site(lngSiteNum).Active Then
'
'                   If lngErrorword(lngSiteNum) <> 0 Then
'                        strTmp = "ErrorInfo:" & Convert10to2(lngErrorword(lngSiteNum), 16)
'                        Call SetLastError(2, "GetDetailErrorInfo", strTmp, lngSiteNum)
'                   End If
'
'                End If
'
'            Next lngSiteNum
            Exit Sub
        Else
            Call SetLastError(1, "GetDetailErrorInfoWithOutErrorRise:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "GetDetailErrorInfoWithOutErrorRise:First WaitReady")
        Exit Sub
        
    End If
        
End Sub


Public Sub GetModeInfo(eReadMode As READ_MODE, lngMode() As Long)
    Dim lngCnt As Long
    Dim eLogicState() As LOGIC_STATE
    Dim lngSiteNum As Long
            
    ReDim lngMode(TheExec.sites.ExistingCount - 1)
    
    LoggingSection = "GetModeInfo"
    
    
    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_GETMODEDATA)
        If WaitReady = True Then
            'Send Sub Command
            mode = LS_LOW
            Param = Convert10to2(CLng(eReadMode))
            If WaitReady = False Then
                Call SetLastError(1, "GetModeInfo:WaitReady After Sending SubCommand[" & Convert10to2(CLng(eReadMode)) & "] is done")
                Exit Sub
            End If
            
            For lngCnt = 1 To 4
            
                mode = LS_LOW
                Param = Convert10to2(1)
                If WaitReady = False Then
                    Call SetLastError(1, "GetModeInfo:WaitReady After Sending Value is done")
                    Exit Sub
                End If
                Call GetStatusLogic(eLogicState, m_strStatusPinName)
                For lngSiteNum = 0 To TheExec.sites.ExistingCount - 1
                    
                    If TheExec.sites.site(lngSiteNum).Active Then
                        
                        lngMode(lngSiteNum) = lngMode(lngSiteNum) Or IIf(eLogicState(lngSiteNum) = LS_HIGH, 2 ^ (lngCnt - 1), 0)
                        
                    End If
                    
                Next lngSiteNum
                
            Next lngCnt
            
            Exit Sub
        Else
            Call SetLastError(1, "GetModeInfo:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "GetModeInfo:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


Public Sub GetFpgaVersion(strFpgaVersion() As String, Optional lngLngValue As Long = -1, Optional strSMIAignoreVersion As String = "2.1")
    Dim lngCnt As Long
    Dim lngSiteNum As Long
    Dim lngMinorVersion() As Long
    Dim lngMajorVersion() As Long
    Dim strVersion As String
    
    LoggingSection = "GetFpgaVersion"
    
    ReDim strFpgaVersion(TheExec.sites.ExistingCount - 1)
    ReDim lngMinorVersion(TheExec.sites.ExistingCount - 1)
    ReDim lngMajorVersion(TheExec.sites.ExistingCount - 1)
        
    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_GETMODEDATA)
        
        If WaitReady = True Then
            'Send Sub Command Mainor Version
            mode = LS_LOW
            Param = Convert10to2(FPGA_MINOR_VERSION_INDEX)
            If WaitReady = False Then
                Call SetLastError(1, "FpgaVersion[Minor]:WaitReady After Sending SubCommand[" & Convert10to2(FPGA_MINOR_VERSION_INDEX) & "] is done")
                Exit Sub
            End If
            
            'Read Minor Version Bit
            If GetRegister(4, lngMinorVersion, "FpgaVersion[Minor]") = False Then
                Exit Sub
            End If
            
            
            mode = LS_HIGH
            Param = Convert10to2(COMMAND_GETMODEDATA)
            
            If WaitReady = True Then
                'Send Sub Command Major Version
                mode = LS_LOW
                Param = Convert10to2(FPGA_MAJOR_VERSION_INDEX)
                If WaitReady = False Then
                    Call SetLastError(1, "FpgaVersion[Major]:WaitReady After Sending SubCommand[" & Convert10to2(FPGA_MAJOR_VERSION_INDEX) & "] is done")
                    Exit Sub
                End If
                
                'Read Major Version Bit
                If GetRegister(4, lngMajorVersion, "FpgaVersion[Major]") = False Then
                    Exit Sub
                End If
                
                For lngSiteNum = 0 To TheExec.sites.ExistingCount - 1
                    
                    If TheExec.sites.site(lngSiteNum).Active Then
                        
                        If lngLngValue = -1 Then
                            strFpgaVersion(lngSiteNum) = Convert10to16(lngMajorVersion(lngSiteNum)) & "." & Convert10to16(lngMinorVersion(lngSiteNum))
                            Call SMIACapture_ignore_RevisionCheck(strFpgaVersion(lngSiteNum), strSMIAignoreVersion)   ' add Y.Kimura 2011/04/01
                        Else
                            strFpgaVersion(lngSiteNum) = lngMajorVersion(lngSiteNum) * 16 + lngMinorVersion(lngSiteNum)
                            Call SMIACapture_ignore_RevisionCheck(strFpgaVersion(lngSiteNum), strSMIAignoreVersion)   ' add Y.Kimura 2011/04/01
                        End If
                    End If
                    
                Next lngSiteNum
                Exit Sub
            Else
                Call SetLastError(1, "FpgaVersion[Major]:WaitReady After Sending Command is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "FpgaVersion[Minor]:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "FpgaVersion:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


Public Sub SetCmosBitOrderMode(eCmosBitOrderMode As CMOS_BITORDER_MODE)
    
    LoggingSection = "SetCmosBitOrderMode"
    
    Enable = LS_LOW
    If WaitReady(FM_BOTH_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETCMOSATTRIBUTEMODE)
        If WaitReady(FM_BOTH_FPGA) = True Then
        
            mode = LS_LOW
            Param = Convert10to2(SUBCOMMAND_SETCMOSBITORDERMODE)
            If WaitReady(FM_BOTH_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eCmosBitOrderMode))
                If WaitReady(FM_BOTH_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetCmosBitOrderMode:WaitReady After Sending Value of SubCommand [BitOrder] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetCmosBitOrderMode:WaitReady After Sending SubCommand [BitOrder] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetCmosBitOrderMode:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "SetCmosBitOrderMode:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


'Public Sub SetCmosPixelSkipMode(eCmosPixelSkipMode As CMOS_PIXELSKIP_MODE)
'    Dim lngCap() As Long
'    Dim lngSiteNum As Long
'
'    LoggingSection = "SetCmosPixelSkipMode"
'
'    If eCmosPixelSkipMode <> 1 Then
'        Call GetModeInfo(RM_CAPTUREMODE, lngCap)
'        For lngSiteNum = 0 To TheExec.Sites.ExistingCount - 1
'            If TheExec.Sites.site(lngSiteNum).Active Then
'                If lngCap(lngSiteNum) = 11 Then
'                    MsgBox "Does not support the Skip function when CMOS Master Mode!", vbOKOnly, "Missing"
'                    eCmosPixelSkipMode = 1
'                End If
'                Exit For
'            End If
'        Next lngSiteNum
'    End If
'
'    Enable = LS_LOW
'    If WaitReady(FM_BOTH_FPGA) = True Then
'
'        Mode = LS_HIGH
'        Param = Convert10to2(COMMAND_SETCMOSATTRIBUTEMODE)
'        If WaitReady(FM_BOTH_FPGA) = True Then
'
'            Mode = LS_LOW
'            Param = Convert10to2(SUBCOMMAND_SETCMOSPIXELSKIPMODE)
'            If WaitReady(FM_BOTH_FPGA) = True Then
'                Mode = LS_LOW
'                Param = Convert10to2(CLng(eCmosPixelSkipMode))
'                If WaitReady(FM_BOTH_FPGA) = True Then
'                    Exit Sub
'                Else
'                    Call SetLastError(1, "SetCmosPixelSkipMode:WaitReady After Sending Value of SubCommand [PixelSkip] is done")
'                    Exit Sub
'                End If
'            Else
'                Call SetLastError(1, "SetCmosPixelSkipMode:WaitReady After Sending SubCommand [PixelSkip] is done")
'                Exit Sub
'            End If
'        Else
'            Call SetLastError(1, "SetCmosPixelSkipMode:WaitReady After Sending Command is done")
'            Exit Sub
'        End If
'
'    Else
'        Call SetLastError(1, "SetCmosPixelSkipMode:First WaitReady")
'        Exit Sub
'    End If
'
'End Sub


Public Sub SetCmosClockLatchMode(eCmosClockLatchChan As CMOS_LATCH_CH, eCmosClockLatchMode As CMOS_LATCH_MODE)
    Dim lngCnt As Long
    
    LoggingSection = "SetCmosClockLatchMode"
    
    Enable = LS_LOW
    If WaitReady(FM_BOTH_FPGA) = True Then
    
        If eCmosClockLatchChan = 10 Then
            For lngCnt = 3 To 5
                mode = LS_HIGH
                Param = Convert10to2(COMMAND_SETCMOSATTRIBUTEMODE)
                If WaitReady(FM_BOTH_FPGA) = True Then
                    mode = LS_LOW
                    Param = Convert10to2(CLng(lngCnt))
                    If WaitReady(FM_BOTH_FPGA) = True Then
                        mode = LS_LOW
                        Param = Convert10to2(CLng(eCmosClockLatchMode))
                        If WaitReady(FM_BOTH_FPGA) = True Then
                        Else
                            Call SetLastError(1, "SetCmosClockLatchMode:WaitReady After Sending Value of SubCommand [LatchMode] is done")
                            Exit Sub
                        End If
                    Else
                        Call SetLastError(1, "SetCmosClockLatchMode:WaitReady After Sending SubCommand [LatchChannel] is done")
                        Exit Sub
                    End If
                Else
                    Call SetLastError(1, "SetCmosClockLatchMode:WaitReady After Sending Command is done")
                    Exit Sub
                End If
            Next lngCnt
            Exit Sub
        
        ElseIf (eCmosClockLatchChan >= 3 And eCmosClockLatchChan <= 5) Then
            mode = LS_HIGH
            Param = Convert10to2(COMMAND_SETCMOSATTRIBUTEMODE)
            If WaitReady(FM_BOTH_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eCmosClockLatchChan))
                If WaitReady(FM_BOTH_FPGA) = True Then
                    mode = LS_LOW
                    Param = Convert10to2(CLng(eCmosClockLatchMode))
                    If WaitReady(FM_BOTH_FPGA) = True Then
                        Exit Sub
                    Else
                        Call SetLastError(1, "SetCmosClockLatchMode:WaitReady After Sending Value of SubCommand [LatchMode] is done")
                        Exit Sub
                    End If
                Else
                    Call SetLastError(1, "SetCmosClockLatchMode:WaitReady After Sending SubCommand [LatchChannel] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetCmosClockLatchMode:WaitReady After Sending Command is done")
                Exit Sub
            End If
        
        Else
            MsgBox "Please select channel to set from list!"
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SetCmosClockLatchMode:First WaitReady")
        Exit Sub
    End If
    
End Sub


Public Sub SelectCMOSDataRate(eCmosDataRate As CMOS_DATARATE)       'add T.Minoda 2009/09/17

    LoggingSection = "SelectCMOSDataRate"

    Enable = LS_LOW

    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETCMOSATTRIBUTEMODE)
        If WaitReady = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SELCMOS_SDRDDR))
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eCmosDataRate))
                If WaitReady = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectCMOSDataRate:WaitReady After Sending Value of SubCommand [CMOS_SelectSDRorDDR] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectCMOSDataRate:WaitReady After Sending SubCommand [CMOS_SelectSDRorDDR] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectCMOSDataRate:WaitReady After Sending Command is done")
            Exit Sub
        End If

    Else
        Call SetLastError(1, "SelectCMOSDataRate:First WaitReady")
        Exit Sub
    End If

End Sub


'Public Sub SelectCMOSMasterCircuit(eMasterCircuit As MASTER_CIRCUIT)
'
'    LoggingSection = "SelectCMOSMasterCircuit"
'
'    Enable = LS_LOW
'
'    If WaitReady = True Then
'        Mode = LS_HIGH
'        Param = Convert10to2(COMMAND_SETCMOSATTRIBUTEMODE)
'        If WaitReady = True Then
'            Mode = LS_LOW
'            Param = Convert10to2(CLng(SUBCOMMAND_SELCIRCUIT_MASTER))
'            If WaitReady = True Then
'                Mode = LS_LOW
'                Param = Convert10to2(CLng(eMasterCircuit))
'                If WaitReady = True Then
'                    Exit Sub
'''                    If eMasterCircuit = MC_LINE_ONLY Then
'''                        Exit Sub
'''                    Else
'''                        SetCMOSMasterCapSize mcPixelNum, mcLineNum
'''                    End If
'                Else
'                    Call SetLastError(1, "SelectCMOSMasterCircuit:WaitReady After Sending Value of SubCommand [SelectCircuitForCMOSMaster] is done")
'                    Exit Sub
'                End If
'            Else
'                Call SetLastError(1, "SelectCMOSMasterCircuit:WaitReady After Sending SubCommand [SelectCircuitForCMOSMaster] is done")
'                Exit Sub
'            End If
'        Else
'            Call SetLastError(1, "SelectCMOSMasterCircuit:WaitReady After Sending Command is done")
'            Exit Sub
'        End If
'
'    Else
'        Call SetLastError(1, "SelectCMOSMasterCircuit:First WaitReady")
'        Exit Sub
'    End If
'
'End Sub


Public Sub SetCMOSMasterCapSize(mcPixelNum As Long, mcLineNum As Long)
    Dim lngPix(3) As Long
    Dim lngLin(3) As Long
    
    If (mcPixelNum < 100) Or (mcPixelNum > 6400) Or (mcLineNum < 100) Or (mcLineNum > 5000) Then
        MsgBox "Please set it in more than 'H100xV100' or 'H6400xV5000' following."
        Exit Sub
    End If
    
    lngPix(0) = (mcPixelNum And &HF)
    lngPix(1) = (mcPixelNum And &HF0) / (2 ^ 4)
    lngPix(2) = (mcPixelNum And &HF00) / (2 ^ 8)
    lngPix(3) = (mcPixelNum And &HF000) / (2 ^ 12)
    lngLin(0) = (mcLineNum And &HF)
    lngLin(1) = (mcLineNum And &HF0) / (2 ^ 4)
    lngLin(2) = (mcLineNum And &HF00) / (2 ^ 8)
    lngLin(3) = (mcLineNum And &HF000) / (2 ^ 12)


    Enable = LS_LOW
    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETCMOSATTRIBUTEMODE)
        If WaitReady = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SETPIXELNUM_MASTER))
            If WaitReady = True Then
                'Send Msb of Pixel Number
                mode = LS_LOW
                Param = Convert10to2(lngPix(3))
                If WaitReady = False Then
                    Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending Value of SubCommand [Pixel Number] is done")
                    Exit Sub
                End If
                
                'Send Msb-1 of Pixel Number
                mode = LS_LOW
                Param = Convert10to2(lngPix(2))
                If WaitReady = False Then
                    Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending Value of SubCommand [Pixel Number] is done")
                    Exit Sub
                End If
                
                'Send Lsb+1 of Pixel Number
                mode = LS_LOW
                Param = Convert10to2(lngPix(1))
                If WaitReady = False Then
                    Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending Value of SubCommand [Pixel Number] is done")
                    Exit Sub
                End If
                
                'Send Lsb of Pixel Number
                mode = LS_LOW
                Param = Convert10to2(lngPix(0))
                If WaitReady = False Then
                    Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending Value of SubCommand [Pixel Number] is done")
                    Exit Sub
                End If
                
            Else
                Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending SubCommand [SetPixelNumber for CMOSMaster] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SetCMOSMasterCapSize:First WaitReady")
        Exit Sub
    End If
    
    
    Enable = LS_LOW
    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETCMOSATTRIBUTEMODE)
        If WaitReady = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SETLINENUM_MASTER))
            If WaitReady = True Then
                'Send Msb of Line Number
                mode = LS_LOW
                Param = Convert10to2(lngLin(3))
                If WaitReady = False Then
                    Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending Value of SubCommand [Line Number] is done")
                    Exit Sub
                End If
                
                'Send Msb-1 of Line Number
                mode = LS_LOW
                Param = Convert10to2(lngLin(2))
                If WaitReady = False Then
                    Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending Value of SubCommand [Line Number] is done")
                    Exit Sub
                End If
                
                'Send Lsb+1 of Line Number
                mode = LS_LOW
                Param = Convert10to2(lngLin(1))
                If WaitReady = False Then
                    Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending Value of SubCommand [Line Number] is done")
                    Exit Sub
                End If
                
                'Send Lsb of Line Number
                mode = LS_LOW
                Param = Convert10to2(lngLin(0))
                If WaitReady = False Then
                    Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending Value of SubCommand [Line Number] is done")
                    Exit Sub
                End If
                
            Else
                Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending SubCommand [SetLineNumber for CMOSMaster] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetCMOSMasterCapSize:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SetCMOSMasterCapSize:First WaitReady")
        Exit Sub
    End If

End Sub


Public Sub SelectCMOSThroughDataBit(eInputDataBit As CMTH_INPUT_DATABIT)

    LoggingSection = "SelectCMOSThroughDataBit"

    Enable = LS_LOW

    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETCMOSATTRIBUTEMODE)
        If WaitReady = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SELTHROUGHDATABIT))
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eInputDataBit))
                If WaitReady = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectCMOSThroughDataBit:WaitReady After Sending Value of SubCommand [CMOSThroughInputDataBit] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectCMOSThroughDataBit:WaitReady After Sending SubCommand [CMOSThroughInputDataBit] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectCMOSThroughDataBit:WaitReady After Sending Command is done")
            Exit Sub
        End If

    Else
        Call SetLastError(1, "SelectCMOSThroughDataBit:First WaitReady")
        Exit Sub
    End If

End Sub


Public Sub SetExpandDataCount(eExpandDataChan As EXDATA_CHAN, eExpandDataNum As EXDATA_NUM)
    
    LoggingSection = "SetExpandDataCount"
    
    Enable = LS_LOW
    If WaitReady(FM_DECODE_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady(FM_DECODE_FPGA) = True Then
        
            mode = LS_LOW
            Param = Convert10to2(CLng(eExpandDataChan))
            If WaitReady(FM_DECODE_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eExpandDataNum))
                If WaitReady(FM_DECODE_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetExpandDataCount:WaitReady After Sending Value of SubCommand [ExpandDataNumber] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetExpandDataCount:WaitReady After Sending SubCommand [ExpandDataChannel] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetExpandDataCount:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "SetExpandDataCount:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


Public Sub SelectFrameStartSyncCode(eFrameStartSyncCode As FRAME_START_CODE)
    
    LoggingSection = "SelectFrameStartSyncCode"
    
    Enable = LS_LOW
    If WaitReady(FM_BOTH_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady(FM_BOTH_FPGA) = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SELFRAMESTART_SYNCCODE))
            If WaitReady(FM_BOTH_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eFrameStartSyncCode))
                If WaitReady(FM_BOTH_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectFrameStartSyncCode:WaitReady After Sending Value of SubCommand [FrameStartSyncCode] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectFrameStartSyncCode:WaitReady After Sending SubCommand [SelectFrameStartCode] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectFrameStartSyncCode:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SelectFrameStartSyncCode:First WaitReady")
        Exit Sub
    End If
    
End Sub


Public Sub SelectFrameStartSyncCodeForVD(eFrameStartSyncCode As FRAME_START_CODE)
    
    LoggingSection = "SelectFrameStartSyncCode"
    
    EnableVD = LS_LOW
    If WaitReady(FM_VIRTUAL_FPGA) = True Then
    
        ModeVD = LS_HIGH
        ParamVD = Convert10to2(VDCOMMAND_SETEXPANDDATA)
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
            ModeVD = LS_LOW
            ParamVD = Convert10to2(CLng(SUBCOMMAND_SELFRAMESTART_SYNCCODE))
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                ModeVD = LS_LOW
                ParamVD = Convert10to2(CLng(eFrameStartSyncCode))
                If WaitReady(FM_VIRTUAL_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectFrameStartSyncCode:WaitReady After Sending Value of SubCommand [FrameStartSyncCode] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectFrameStartSyncCode:WaitReady After Sending SubCommand [SelectFrameStartCode] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectFrameStartSyncCode:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SelectFrameStartSyncCode:First WaitReady")
        Exit Sub
    End If
    
End Sub


Public Sub SelectCRCOperation(eCRCOperation As CRC_ONOFF)
    
    LoggingSection = "SelectCRCOperation"
    
    Enable = LS_LOW
    If WaitReady(FM_DECODE_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady(FM_DECODE_FPGA) = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SELCRCOPERATION))
            If WaitReady(FM_DECODE_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eCRCOperation))
                If WaitReady(FM_DECODE_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectCRCOperation:WaitReady After Sending Value of SubCommand [CRCOnOffSelectValue] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectCRCOperation:WaitReady After Sending SubCommand [CRCOnOffSelect] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectCRCOperation:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SelectCRCOperation:First WaitReady")
        Exit Sub
    End If
    
End Sub


Public Sub SelectECCOperation(eECCOperation As ECC_ONOFF)
    
    LoggingSection = "SelectECCOperation"
    
    Enable = LS_LOW
    If WaitReady(FM_DECODE_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady(FM_DECODE_FPGA) = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SELECCOPERATION))
            If WaitReady(FM_DECODE_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eECCOperation))
                If WaitReady(FM_DECODE_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectECCOperation:WaitReady After Sending Value of SubCommand [ECCOnOffSelectValue] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectECCOperation:WaitReady After Sending SubCommand [ECCOnOffSelect] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectECCOperation:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SelectECCOperation:First WaitReady")
        Exit Sub
    End If
    
End Sub


Public Sub SelectMPCLKRfswControl(eMpclkRfswCtrl As MIPICLK_RFSW_CTRL)
    
    LoggingSection = "SelectMPCLKRfswControl"
    
    Enable = LS_LOW
    If WaitReady(FM_DECODE_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady(FM_DECODE_FPGA) = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SELMIPICLKRFSW_CTRL))
            If WaitReady(FM_DECODE_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eMpclkRfswCtrl))
                If WaitReady(FM_DECODE_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectMPCLKRfswControl:WaitReady After Sending Value of SubCommand [RFSWControlSelectValue] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectMPCLKRfswControl:WaitReady After Sending SubCommand [RFSWControlSelect] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectMPCLKRfswControl:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SelectMPCLKRfswControl:First WaitReady")
        Exit Sub
    End If
    
End Sub


Public Sub SelectMPDATARfswControl(eMpdatRfswCtrl As MIPIDATA_RFSW_CTRL)
    
    LoggingSection = "SelectMPDATARfswControl"
    
    Enable = LS_LOW
    If WaitReady(FM_DECODE_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady(FM_DECODE_FPGA) = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SELMIPIDATARFSW_CTRL))
            If WaitReady(FM_DECODE_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eMpdatRfswCtrl))
                If WaitReady(FM_DECODE_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectMPDATARfswControl:WaitReady After Sending Value of SubCommand [RFSWControlSelectValue] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectMPDATARfswControl:WaitReady After Sending SubCommand [RFSWControlSelect] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectMPDATARfswControl:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SelectMPDATARfswControl:First WaitReady")
        Exit Sub
    End If
    
End Sub


Public Sub SelectMPDATALaneAssign(eMpDatAssign As MIPI_DATALANE_ASSIGN)           'add H.Miyamoto 2009/3/24
    Dim lngLAssignLsb As Long
    Dim lngLAssignMsb As Long
    
    LoggingSection = "SelectMPDATALaneAssign"
    
    lngLAssignLsb = CLng(eMpDatAssign) And &HF
    lngLAssignMsb = (CLng(eMpDatAssign) And &HF0) / (2 ^ 4)
    
    Enable = LS_LOW
    If WaitReady(FM_DECODE_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady(FM_DECODE_FPGA) = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SELMIPIDATALANE_ASSIGN))
            If WaitReady(FM_DECODE_FPGA) = True Then
                'Send Msb of LaneAssign
                mode = LS_LOW
                Param = Convert10to2(lngLAssignMsb)
                If WaitReady(FM_DECODE_FPGA) = False Then
                    Call SetLastError(1, "SelectMPDATALaneAssign:WaitReady After Sending Value of SubCommand [LaneAssignSelectValue Msb] is done")
                    Exit Sub
                End If

                'Send Lsb of LaneAssign
                mode = LS_LOW
                Param = Convert10to2(lngLAssignLsb)
                If WaitReady(FM_DECODE_FPGA) = False Then
                    Call SetLastError(1, "SelectMPDATALaneAssign:WaitReady After Sending Value of SubCommand [LaneAssignSelectValue Lsb] is done")
                    Exit Sub
                End If
                Exit Sub
            Else
                Call SetLastError(1, "SelectMPDATALaneAssign:WaitReady After Sending SubCommand [LaneAssignSelect] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectMPDATALaneAssign:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SelectMPDATALaneAssign:First WaitReady")
        Exit Sub
    End If
    
End Sub


Public Sub SelectSMIAAssign(eSmiaAssign As SMIA_ASSIGN)                 'add T.Minoda 2009/09/17

    LoggingSection = "SelectSMIAAssign"

    Enable = LS_LOW

    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SELSMIA_ASSIGN))
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eSmiaAssign))
                If WaitReady = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectSMIAAssign:WaitReady After Sending Value of SubCommand [SMIA_AssignNo] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectSMIAAssign:WaitReady After Sending SubCommand [SMIA_AssignSelect] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectSMIAAssign:WaitReady After Sending Command is done")
            Exit Sub
        End If

    Else
        Call SetLastError(1, "SelectSMIAAssign:First WaitReady")
        Exit Sub
    End If

End Sub


Public Sub SelectCMOSDataAssign(eCmosDataAssign As CMOS_DATA_ASSIGN)        'add T.Minoda 2009/09/17

    LoggingSection = "SelectCMOSDataAssign"

    Enable = LS_LOW

    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETCMOSATTRIBUTEMODE)
        If WaitReady = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SELCMOSDATAASSIGN))
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eCmosDataAssign))
                If WaitReady = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectCMOSDataAssign:WaitReady After Sending Value of SubCommand [CMOS_DataAssignSelectValue] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectCMOSDataAssign:WaitReady After Sending SubCommand [CMOS_DataAssignSelect] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectCMOSDataAssign:WaitReady After Sending Command is done")
            Exit Sub
        End If

    Else
        Call SetLastError(1, "SelectCMOSDataAssign:First WaitReady")
        Exit Sub
    End If

End Sub


Public Sub SetLvdsPNReverse(eLvdsPNreverse As LVDS_PN_REVERSE)                 'add T.Minoda 2009/09/17

    LoggingSection = "SetLvdsPNReverse"

    Enable = LS_LOW

    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SETLVDS_PNREVERSE))
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eLvdsPNreverse))
                If WaitReady = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetLvdsPNReverse:WaitReady After Sending Value of SubCommand [MIPI/SMIA_PNreverse] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetLvdsPNReverse:WaitReady After Sending SubCommand [MIPI/SMIA_PNreverse] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetLvdsPNReverse:WaitReady After Sending Command is done")
            Exit Sub
        End If

    Else
        Call SetLastError(1, "SetLvdsPNReverse:First WaitReady")
        Exit Sub
    End If

End Sub


Public Sub SetDelayTapForMIPI(eDelaySetChan As DELAYSET_MIPICH, Optional lngDelayTap As Long = 0, Optional lngSetSite As Long = -1)
    Dim lngCnt As Long
    Dim lngTapLsb As Long
    Dim lngTapMsb As Long

    LoggingSection = "SetDelayTapForMIPI"

    If (lngDelayTap < 0) Or (lngDelayTap > 63) Then
        MsgBox "Please Set Range Of DelayTap to 0 - 63!"
        Exit Sub
    End If

    lngTapLsb = lngDelayTap And &HF
    lngTapMsb = (lngDelayTap And &HF0) / (2 ^ 4)

    Enable = LS_LOW
    If WaitReady = True Then
    
        If (eDelaySetChan >= 0 And eDelaySetChan <= 5) Then
            mode = LS_HIGH
            Param = Convert10to2(COMMAND_SETSKEWADJMODE, , lngSetSite)
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(SUBCOMMAND_SETSKEW_MIPI, , lngSetSite)
                If WaitReady = True Then
                    'Send Adjust Channel
                    mode = LS_LOW
                    Param = Convert10to2(eDelaySetChan, , lngSetSite)
                    If WaitReady = False Then
                        Call SetLastError(1, "SetDelayTapForMIPI:WaitReady After Sending Value of SubCommand [SkewForLvds Channel] is done")
                        Exit Sub
                    End If

                    'Send Msb of Tap
                    mode = LS_LOW
                    Param = Convert10to2(lngTapMsb, , lngSetSite)
                    If WaitReady = False Then
                        Call SetLastError(1, "SetDelayTapForMIPI:WaitReady After Sending Value of SubCommand [SkewForLvds MsbTap] is done")
                        Exit Sub
                    End If

                    'Send Lsb of Tap
                    mode = LS_LOW
                    Param = Convert10to2(lngTapLsb, , lngSetSite)
                    If WaitReady = False Then
                        Call SetLastError(1, "SetDelayTapForMIPI:WaitReady After Sending Value of SubCommand [SkewForLvds LsbTap] is done")
                        Exit Sub
                    End If
                    Exit Sub
                Else
                    Call SetLastError(1, "SetDelayTapForMIPI:WaitReady After Sending SubCommand [SkewForLvds] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetDelayTapForMIPI:WaitReady After Sending Command is done")
                Exit Sub
            End If

        Else
            MsgBox "Please select channel to set from list!"
            Exit Sub
        End If

    Else
        Call SetLastError(1, "SetDelayTapForMIPI:First WaitReady")
        Exit Sub
    End If

End Sub


Public Sub SetDelayTapForSMIA(eDelaySetChan As DELAYSET_SMIACH, Optional lngDelayTap As Long = 0, Optional lngSetSite As Long = -1)
    Dim lngCnt As Long
    Dim lngTapLsb As Long
    Dim lngTapMsb As Long

    LoggingSection = "SetDelayTapForSMIA"

    If (lngDelayTap < 0) Or (lngDelayTap > 63) Then
        MsgBox "Please Set Range Of DelayTap to 0 - 63!"
        Exit Sub
    End If

    lngTapLsb = lngDelayTap And &HF
    lngTapMsb = (lngDelayTap And &HF0) / (2 ^ 4)

    Enable = LS_LOW
    If WaitReady = True Then
    
        If (eDelaySetChan >= 0 And eDelaySetChan <= 8) Then
            mode = LS_HIGH
            Param = Convert10to2(COMMAND_SETSKEWADJMODE, , lngSetSite)
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(SUBCOMMAND_SETSKEW_SMIA, , lngSetSite)
                If WaitReady = True Then
                    'Send Adjust Channel
                    mode = LS_LOW
                    Param = Convert10to2(eDelaySetChan, , lngSetSite)
                    If WaitReady = False Then
                        Call SetLastError(1, "SetDelayTapForSMIA:WaitReady After Sending Value of SubCommand [SkewForLvds Channel] is done")
                        Exit Sub
                    End If

                    'Send Msb of Tap
                    mode = LS_LOW
                    Param = Convert10to2(lngTapMsb, , lngSetSite)
                    If WaitReady = False Then
                        Call SetLastError(1, "SetDelayTapForSMIA:WaitReady After Sending Value of SubCommand [SkewForLvds MsbTap] is done")
                        Exit Sub
                    End If

                    'Send Lsb of Tap
                    mode = LS_LOW
                    Param = Convert10to2(lngTapLsb, , lngSetSite)
                    If WaitReady = False Then
                        Call SetLastError(1, "SetDelayTapForSMIA:WaitReady After Sending Value of SubCommand [SkewForLvds LsbTap] is done")
                        Exit Sub
                    End If
                    Exit Sub
                Else
                    Call SetLastError(1, "SetDelayTapForSMIA:WaitReady After Sending SubCommand [SkewForLvds] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetDelayTapForSMIA:WaitReady After Sending Command is done")
                Exit Sub
            End If

        Else
            MsgBox "Please select channel to set from list!"
            Exit Sub
        End If

    Else
        Call SetLastError(1, "SetDelayTapForSMIA:First WaitReady")
        Exit Sub
    End If

End Sub


'Public Sub SetDelayTapForSMIAFromSheet(Optional lngSetSite As Long = -1, Optional strTapSetName As String = "Default")
'    Dim SkewMapSht As Worksheet
'    Dim FindTapSet As Long
'    Dim FindModeLine As Long
'
'    Dim lngCnt As Long
'    Dim lngTapLsb As Long
'    Dim lngTapMsb As Long
'
'    LoggingSection = "SetDelayTapForSMIAFromSheet"
'
''''''    Set SkewMapSht = ThisWorkbook.Worksheets("SkewMap")
''''''    With SkewMapSht
''''''        For FindTapSet = 1 To 100
''''''            If .Cells(1, FindTapSet) = strtapset Then
''''''                For FindModeLine = 3 To 10
''''''
''''''            Else
''''''            End If
''''''        Next FindTapSet
'
'
'    If (lngDelayTap < 0) Or (lngDelayTap > 63) Then
'        MsgBox "Please Set Range Of DelayTap to 0 - 63!"
'        Exit Sub
'    End If
'
'    lngTapLsb = lngDelayTap And &HF
'    lngTapMsb = (lngDelayTap And &HF0) / (2 ^ 4)
'
'    Enable = LS_LOW
'    If WaitReady = True Then
'
'        If (eDelaySetChan >= 0 And eDelaySetChan <= 8) Then
'            Mode = LS_HIGH
'            Param = Convert10to2(COMMAND_SETSKEWADJMODE, , lngSetSite)
'            If WaitReady = True Then
'                Mode = LS_LOW
'                Param = Convert10to2(SUBCOMMAND_SETSKEW_SMIA, , lngSetSite)
'                If WaitReady = True Then
'                    'Send Adjust Channel
'                    Mode = LS_LOW
'                    Param = Convert10to2(eDelaySetChan, , lngSetSite)
'                    If WaitReady = False Then
'                        Call SetLastError(1, "SetDelayTapForLvds:WaitReady After Sending Value of SubCommand [SkewForLvds Channel] is done")
'                        Exit Sub
'                    End If
'
'                    'Send Msb of Tap
'                    Mode = LS_LOW
'                    Param = Convert10to2(lngTapMsb, , lngSetSite)
'                    If WaitReady = False Then
'                        Call SetLastError(1, "SetDelayTapForLvds:WaitReady After Sending Value of SubCommand [SkewForLvds MsbTap] is done")
'                        Exit Sub
'                    End If
'
'                    'Send Lsb of Tap
'                    Mode = LS_LOW
'                    Param = Convert10to2(lngTapLsb, , lngSetSite)
'                    If WaitReady = False Then
'                        Call SetLastError(1, "SetDelayTapForLvds:WaitReady After Sending Value of SubCommand [SkewForLvds LsbTap] is done")
'                        Exit Sub
'                    End If
'                    Exit Sub
'                Else
'                    Call SetLastError(1, "SetDelayTapForLvds:WaitReady After Sending SubCommand [SkewForLvds] is done")
'                    Exit Sub
'                End If
'            Else
'                Call SetLastError(1, "SetDelayTapForLvds:WaitReady After Sending Command is done")
'                Exit Sub
'            End If
'
'        Else
'            MsgBox "Please select channel to set from list!"
'            Exit Sub
'        End If
'
'    Else
'        Call SetLastError(1, "SetDelayTapForLvds:First WaitReady")
'        Exit Sub
'    End If
'
'End Sub


Public Sub SetDelayTapForCMOS(eDelaySetChan As DELAYSET_CMOSCH, Optional lngDelayTap As Long = 0, Optional lngSetSite As Long = -1)
    Dim lngTapLsb As Long
    Dim lngTapMsb As Long
    Dim ActualSubCom As Long
    Dim ActualSetCh As Long

    LoggingSection = "SetDelayTapForCMOS"

    If lngDelayTap < 0 Or lngDelayTap > 63 Then
        MsgBox "Please Set Range Of DelayTap to 0 - 63!"
        Exit Sub
    End If

    lngTapLsb = lngDelayTap And &HF
    lngTapMsb = (lngDelayTap And &HF0) / (2 ^ 4)

    Enable = LS_LOW
    If WaitReady = True Then

        If (eDelaySetChan >= 0 And eDelaySetChan <= 15) Then
            ActualSubCom = SUBCOMMAND_SETSKEW_CMOS_A
            ActualSetCh = eDelaySetChan
        ElseIf (eDelaySetChan >= 16 And eDelaySetChan <= 31) Then
            ActualSubCom = SUBCOMMAND_SETSKEW_CMOS_B
            ActualSetCh = eDelaySetChan - 16
        ElseIf (eDelaySetChan >= 32 And eDelaySetChan <= 36) Then
            ActualSubCom = SUBCOMMAND_SETSKEW_CMOS_C
            ActualSetCh = eDelaySetChan - 32
        Else
            MsgBox "Please select channel to set from list!"
            Exit Sub
        End If
        
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETSKEWADJMODE, , lngSetSite)
        If WaitReady = True Then
            mode = LS_LOW
            Param = Convert10to2(ActualSubCom, , lngSetSite)
            If WaitReady = True Then
                'Send Adjust Channel
                mode = LS_LOW
                Param = Convert10to2(ActualSetCh, , lngSetSite)
                If WaitReady = False Then
                    Call SetLastError(1, "SetDelayTapForCMOS:WaitReady After Sending Value of SubCommand [SkewForCmos Bit] is done")
                    Exit Sub
                End If

                'Send Msb of Tap
                mode = LS_LOW
                Param = Convert10to2(lngTapMsb, , lngSetSite)
                If WaitReady = False Then
                    Call SetLastError(1, "SetDelayTapForCMOS:WaitReady After Sending Value of SubCommand [SkewForCmosData MsbTap] is done")
                    Exit Sub
                End If

                'Send Lsb of Tap
                mode = LS_LOW
                Param = Convert10to2(lngTapLsb, , lngSetSite)
                If WaitReady = False Then
                    Call SetLastError(1, "SetDelayTapForCMOS:WaitReady After Sending Value of SubCommand [SkewForCmosData LsbTap] is done")
                    Exit Sub
                End If
                Exit Sub
            Else
                Call SetLastError(1, "SetDelayTapForCMOS:WaitReady After Sending SubCommand [SkewForCmosData] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetDelayTapForCMOS:WaitReady After Sending Command is done")
            Exit Sub
        End If

    Else
        Call SetLastError(1, "SetDelayTapForCMOS:First WaitReady")
        Exit Sub
    End If


End Sub


Public Sub SetVirtualDeviceCheckPattern(eVirtulDevPattern As CHECKER_PATTERN)
    
    LoggingSection = "SetVirtualDeviceCheckPattern"
    
    EnableVD = LS_LOW
    If WaitReady(FM_VIRTUAL_FPGA) = True Then
    
        ModeVD = LS_HIGH
        ParamVD = Convert10to2(VDCOMMAND_SETCHECKERATTRIBUTE)
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
            ModeVD = LS_LOW
            ParamVD = Convert10to2(SUBCOMMAND_SETCHECKER_PATTERN)
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                ModeVD = LS_LOW
                ParamVD = Convert10to2(CLng(eVirtulDevPattern))
                If WaitReady(FM_VIRTUAL_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetVirtualDeviceCheckPattern:WaitReady After Sending Pattern Value is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetVirtualDeviceCheckPattern:WaitReady After Sending SubCommand Pattern is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetVirtualDeviceCheckPattern:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "SetVirtualDeviceCheckPattern:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


Public Sub SetVirtualDeviceEnableSite(lngVirEnabSite As Long)
    
    LoggingSection = "SetVirtualDeviceEnableSite"
    
    If lngVirEnabSite < 1 Or lngVirEnabSite > 16 Then
        MsgBox "Not Support Site!!"
        Exit Sub
    End If
    
    m_lngVirEnabSite = lngVirEnabSite
    
    EnableVD = LS_LOW
    If WaitReady(FM_VIRTUAL_FPGA) = True Then
    
        ModeVD = LS_HIGH
        ParamVD = Convert10to2(VDCOMMAND_SETCHECKERATTRIBUTE)
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
            ModeVD = LS_LOW
            ParamVD = Convert10to2(SUBCOMMAND_SETCHECKER_SITE)
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                ModeVD = LS_LOW
                ParamVD = Convert10to2(lngVirEnabSite)
                If WaitReady(FM_VIRTUAL_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetVirtualDeviceEnableSite:WaitReady After Sending SiteSelect Value is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetVirtualDeviceEnableSite:WaitReady After Sending SubCommand SiteSelect is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetVirtualDeviceEnableSite:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "SetVirtualDeviceEnableSite:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


Public Sub SetLvdsSkipMode(eLvdsSkipMode As LVDS_SKIP_MODE)
    
    LoggingSection = "SetLvdsSkipMode"
    
    EnableVD = LS_LOW
    If WaitReady(FM_VIRTUAL_FPGA) = True Then
    
        ModeVD = LS_HIGH
        ParamVD = Convert10to2(VDCOMMAND_SETCHECKERATTRIBUTE)
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
            ModeVD = LS_LOW
            ParamVD = Convert10to2(SUBCOMMAND_SETLVDS_SKIPMODE)
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                ModeVD = LS_LOW
                ParamVD = Convert10to2(CLng(eLvdsSkipMode))
                If WaitReady(FM_VIRTUAL_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetLvdsSkipMode:WaitReady After Sending Pattern Value is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetLvdsSkipMode:WaitReady After Sending SubCommand Pattern is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetLvdsSkipMode:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "SetLvdsSkipMode:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


Public Sub SetVirtualDeviceCRCError(eVirtulDevCRCErr As CRC_ERROR_MODE)
    
    LoggingSection = "SetVirtualDeviceCRCError"
    
    EnableVD = LS_LOW
    If WaitReady(FM_VIRTUAL_FPGA) = True Then
    
        ModeVD = LS_HIGH
        ParamVD = Convert10to2(VDCOMMAND_SETCHECKERATTRIBUTE)
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
            ModeVD = LS_LOW
            ParamVD = Convert10to2(SUBCOMMAND_SETCHECKER_CRCERROR)
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                ModeVD = LS_LOW
                ParamVD = Convert10to2(CLng(eVirtulDevCRCErr))
                If WaitReady(FM_VIRTUAL_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetVirtualDeviceCRCError:WaitReady After Sending Pattern Value is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetVirtualDeviceCRCError:WaitReady After Sending SubCommand Pattern is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetVirtualDeviceCRCError:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "SetVirtualDeviceCRCError:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


Public Sub SetVirtualDevicePatternSize(eVirtulDevPSize As PATTERN_SIZE)
    
    LoggingSection = "SetVirtualDevicePatternSize"
    
    EnableVD = LS_LOW
    If WaitReady(FM_VIRTUAL_FPGA) = True Then
    
        ModeVD = LS_HIGH
        ParamVD = Convert10to2(VDCOMMAND_SETCHECKERATTRIBUTE)
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
            ModeVD = LS_LOW
            ParamVD = Convert10to2(SUBCOMMAND_SETCHECKER_PATTERNSIZE)
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                ModeVD = LS_LOW
                ParamVD = Convert10to2(CLng(eVirtulDevPSize))
                If WaitReady(FM_VIRTUAL_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetVirtualDevicePatternSize:WaitReady After Sending Pattern Value is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetVirtualDevicePatternSize:WaitReady After Sending SubCommand Pattern is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetVirtualDevicePatternSize:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "SetVirtualDevicePatternSize:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


Public Sub GetStatus(eStatus() As CAP_STATE)
    Dim lngCnt As Long
    Dim eLogicState() As LOGIC_STATE
    Dim lngSiteNum As Long
    
    ReDim eStatus(TheExec.sites.ExistingCount - 1)
    
    m_strFunction = "GetStatus"
    
    mode = LS_HIGH
    If WaitReady = True Then
'    Mode = LS_LOW
        Call GetStatusLogic(eLogicState, m_strStatusPinName)
        For lngSiteNum = 0 To TheExec.sites.ExistingCount - 1
            
            If TheExec.sites.site(lngSiteNum).Active Then
                
                eStatus(lngSiteNum) = IIf(eLogicState(lngSiteNum) = LS_HIGH, CS_ERROR, CS_SUCCESS)
                
            End If
            
        Next lngSiteNum
        Exit Sub
    Else
        Call SetLastError(1, "GetStatus:First WaitReady")
        Exit Sub
    End If
    
End Sub


Public Sub RiseRunTimeError()
    Dim strDate As String
    Dim strMsg As String
    
    strDate = Format(Now, "yyyy/mm/dd_hh:mm:ss") & ","
    strMsg = strDate & "ErrNum:" & Trim(str(m_sLastErrorInfo.lngErrorNum)) & _
                            ",ErrName:" & m_sLastErrorInfo.strErrorName & _
                            ",ErrFunc:" & m_sLastErrorInfo.strErrorFunction & _
                            ",ErrNote:" & m_sLastErrorInfo.strErrortmp

    Call TheExec.ErrorLogMessage(strMsg)
    Call TheExec.ErrorReport

End Sub


Public Sub SetCountMode(eCounterMode As COUNTER_MODE)
    
    LoggingSection = "SetCountMode"
    
    Enable = LS_LOW
    If WaitReady = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETCOUNTERMODE)
        If WaitReady = True Then
        
            mode = LS_LOW
            Param = Convert10to2(SUBCOMMAND_COUNT_SET)
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eCounterMode))
                If WaitReady = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetCountMode:WaitReady After Sending Value of SubCommand [SetCountMode] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetCountMode:WaitReady After Sending SubCommand [SetCountMode] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetCountMode:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "SetCountMode:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


Public Sub GetCounterResult(lngMode() As Long, Optional lngMask As Long = &HFFF)
    Dim lngCnt As Long
    Dim eLogicState() As LOGIC_STATE
    Dim lngSiteNum As Long
        
    ReDim lngMode(TheExec.sites.ExistingCount - 1)
    
    LoggingSection = "GetCounterResult"
    
    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETCOUNTERMODE)
        If WaitReady = True Then
        
            'Send Sub Command
            mode = LS_LOW
            Param = Convert10to2(SUBCOMMAND_COUNT_READ_RESAULT)
            If WaitReady = False Then
                Call SetLastError(1, "GetCounterResult:WaitReady After Sending SubCommand[Read Counter Result] is done")
                Exit Sub
            End If
            
            'Read Minor Version Bit
            If GetRegister(16, lngMode, "GetCounterResult") = False Then
                Exit Sub
            End If
                        
            For lngSiteNum = 0 To TheExec.sites.ExistingCount - 1
                If TheExec.sites.site(lngSiteNum).Active Then
                    lngMode(lngSiteNum) = lngMode(lngSiteNum) And lngMask
                End If
            Next lngSiteNum
            
            Exit Sub
        Else
            Call SetLastError(1, "GetCounterResult:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        
        Call SetLastError(1, "GetCounterResult:First WaitReady")
        Exit Sub
        
    End If
    
End Sub


Public Sub VirtualDeviseCtrl(eVirtualdevCtrl As VIRTULDEV_ENABLE)
    
    LoggingSection = "VirtualDeviseCtrl"
    
    EnableVD = LS_LOW
    If WaitReady(FM_VIRTUAL_FPGA) = True Then
    
        ModeVD = LS_HIGH
        ParamVD = Convert10to2(COMMAND_VIRTULDEVISE_CTRL)
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
        
            If ConvertRevHexStrToLng(CLASS_VERSION) >= 0 Then
                'Send Sub Command
                ModeVD = LS_LOW
                ParamVD = Convert10to2(SUBCOMMAND_SETCHECKER_STRAT_RESET)
                If WaitReady(FM_VIRTUAL_FPGA) = False Then
                    Call SetLastError(1, "VirtualDeviseCtrl:WaitReady After Sending SubCommand[Start/Reset] is done")
                    Exit Sub
                End If
            End If
            
            ModeVD = LS_LOW
            ParamVD = Convert10to2(CLng(eVirtualdevCtrl))
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                Exit Sub
            Else
                Call SetLastError(1, "VirtualDeviseCtrl:WaitReady After Sending Value is done")
            End If
        Else
            Call SetLastError(1, "VirtualDeviseCtrl:WaitReady After Sending Command is done")
        End If
    Else
        
        Call SetLastError(1, "VirtualDeviseCtrl:First WaitReady")
        
    End If

End Sub


Public Sub SetLvdsOutMode(eSignalOutMode As SIGNAL_OUT_MODE)
    
    LoggingSection = "SetLvdsOutMode"
    
    Enable = LS_LOW
    If WaitReady(FM_VIRTUAL_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_VIRTULDEVISE_CTRL)
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
        
            'Send Sub Command
            mode = LS_LOW
            Param = Convert10to2(SUBCOMMAND_SETCHECKER_LVDS_ON_OFF)
            If WaitReady = False Then
                Call SetLastError(1, "SetLvdsOutMode:WaitReady After Sending SubCommand[LVDS ON/OFF] is done")
                Exit Sub
            End If
            
            mode = LS_LOW
            Param = Convert10to2(CLng(eSignalOutMode))
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                Exit Sub
            Else
                Call SetLastError(1, "SetLvdsOutMode:WaitReady After Sending Value is done")
            End If
        Else
            Call SetLastError(1, "SetLvdsOutMode:WaitReady After Sending Command is done")
        End If
    Else
        
        Call SetLastError(1, "SetLvdsOutMode:First WaitReady")
        
    End If

End Sub


Public Sub SetCmosOutMode(eSignalOutMode As SIGNAL_OUT_MODE)
    
    LoggingSection = "SetCmosOutMode"
    
    Enable = LS_LOW
    If WaitReady(FM_VIRTUAL_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_VIRTULDEVISE_CTRL)
        If WaitReady(FM_VIRTUAL_FPGA) = True Then
        
            'Send Sub Command
            mode = LS_LOW
            Param = Convert10to2(SUBCOMMAND_SETCHECKER_CMOS_ON_OFF)
            If WaitReady = False Then
                Call SetLastError(1, "SetCmosOutMode:WaitReady After Sending SubCommand[CMOS ON/OFF] is done")
                Exit Sub
            End If
            
            mode = LS_LOW
            Param = Convert10to2(CLng(eSignalOutMode))
            If WaitReady(FM_VIRTUAL_FPGA) = True Then
                Exit Sub
            Else
                Call SetLastError(1, "SetCmosOutMode:WaitReady After Sending Value is done")
            End If
        Else
            Call SetLastError(1, "SetCmosOutMode:WaitReady After Sending Command is done")
        End If
    Else
        
        Call SetLastError(1, "SetCmosOutMode:First WaitReady")
        
    End If

End Sub


Public Sub SetFrameSkipMode(eFSkipMode As FRAMESKIP_MODE)           'add H.Miyamoto 2009/3/24
    Dim lngBuf() As Long
    Dim lngAve() As Long
    Dim lngSiteNum As Long
    
    LoggingSection = "SetFrameSkipMode"
    
    Enable = LS_LOW
    If WaitReady(FM_DECODE_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETFRAMESKIPMODE)
        If WaitReady(FM_DECODE_FPGA) = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(eFSkipMode))
            If WaitReady(FM_DECODE_FPGA) = True Then
                m_eFSkipMode = eFSkipMode
                Exit Sub
            Else
                Call SetLastError(1, "SetFrameSkipMode:WaitReady After Sending Value is done")
            End If
        Else
            Call SetLastError(1, "SetFrameSkipMode:WaitReady After Sending Command is done")
        End If
    Else
        Call SetLastError(1, "SetCaptureMode:First WaitReady")
    End If

End Sub



'==========================================================================
'
'
'               Private Property
'
'
'==========================================================================


Private Sub InitVariable()
    Dim strUtiltyPins As String
    Dim strPins As String
    Dim lngCnt As Long
    Dim strToken() As String

    strUtiltyPins = tl_tm_GetUtilPinsAndGroups
    strToken = VBA.Split(Trim(strUtiltyPins), ",")
    ReDim m_strUtilityPins(UBound(strToken))
    
    For lngCnt = 0 To UBound(strToken)
        m_strUtilityPins(lngCnt) = strToken(lngCnt)
    Next lngCnt
    
    strPins = tl_tm_GetDigitalPinsAndGroups
    strToken = VBA.Split(Trim(strPins), ",")
    ReDim m_strPins(UBound(strToken))
    
    For lngCnt = 0 To UBound(strToken)
        m_strPins(lngCnt) = strToken(lngCnt)
    Next lngCnt
    
End Sub


Private Property Let LoggingSection(strFunction As String)
    Dim fn As Long
    Dim strDate As String
    
    m_strFunction = strFunction
    
    If m_blLoggingEnabled = True And m_eLoggingMode <> LM_ERROR_ONLY Then
        
        strDate = Format(Now, "yyyy/mm/dd_hh:mm:ss") & ","
        fn = FreeFile
        
        
        Open m_strLogFilePath For Append As #fn
        
            Print #fn, strDate & "########## InTo Function:" & strFunction & " ##########"
            
        Close #fn
        
        Call RefreshLogFile
        
    End If
    
End Property


Private Property Get SystemTempDirectory() As String
    Dim lngLeng As Long
    Dim strName As String
    Dim lngRet As Long

    strName = String(250, Chr(0))
    lngLeng = Len(strName)

    lngRet = GetTempPath(lngLeng, strName)
    
    strName = VBA.Left(strName, VBA.InStr(strName, Chr(0)) - 1)
    
    SystemTempDirectory = strName
    
End Property


Private Property Let LoggingParam(strParam As String)
    Dim fn As Long
    Dim strDate As String
        
    If m_blLoggingEnabled Then
        If m_eLoggingMode <> LM_ERROR_ONLY Then
            strDate = Format(Now, "yyyy/mm/dd_hh:mm:ss") & ","
            fn = FreeFile
            
            
            Open m_strLogFilePath For Append As #fn
            
                Print #fn, strDate & "Send Parameter --------> Param:" & strParam & _
                                            ",{Mode:" & IIf(m_eMode = LS_HIGH, "High", "Low") & "," & _
                                            "Enable:" & IIf(m_eEnable = LS_HIGH, "High", "Low") & "," & _
                                            "Reset:" & IIf(m_eReset = LS_HIGH, "High", "Low") & "," & _
                                            "Strob:" & IIf(m_eStrob = LS_HIGH, "High", "Low") & "}"
            Close #fn
            
            Call RefreshLogFile

        End If
        
    End If
    
End Property


Private Property Let mode(eLogicVal As LOGIC_STATE)
    
    If eLogicVal = LS_HIGH Then
        
        TheHdw.Utility.Pins(m_strModePinName).state = bitHigh
        m_eMode = LS_HIGH
        
    Else
    
        TheHdw.Utility.Pins(m_strModePinName).state = bitLow
        m_eMode = LS_LOW
    
    End If

    TheHdw.WAIT m_dblSlewRateOfUtility

End Property


Private Property Let Rst(eLogicVal As LOGIC_STATE)
    
    If eLogicVal = LS_HIGH Then
        
        TheHdw.Utility.Pins(m_strResetPinName).state = bitHigh
        m_eReset = LS_HIGH
    Else
    
        TheHdw.Utility.Pins(m_strResetPinName).state = bitLow
        m_eReset = LS_LOW
    End If

    TheHdw.WAIT m_dblSlewRateOfUtility

End Property


Private Property Let Enable(eLogicVal As LOGIC_STATE)
    
    If eLogicVal = LS_HIGH Then
        
        TheHdw.Utility.Pins(m_strEnablePinName).state = bitHigh
        m_eEnable = LS_HIGH
        
    Else
    
        TheHdw.Utility.Pins(m_strEnablePinName).state = bitLow
        m_eEnable = LS_LOW
    
    End If

    TheHdw.WAIT m_dblSlewRateOfUtility

End Property


Private Property Let Strob(eLogicVal As LOGIC_STATE)
    
    If eLogicVal = LS_HIGH Then
        
        TheHdw.Utility.Pins(m_strStrobPinName).state = bitHigh
        m_eStrob = LS_HIGH
    Else
    
        TheHdw.Utility.Pins(m_strStrobPinName).state = bitLow
        m_eStrob = LS_LOW
    End If

    TheHdw.WAIT m_dblSlewRateOfUtility

End Property


Private Property Let Param(ByVal strParam As String)
    Dim lngCnt As Long
    Dim lngIndex As Long
    Dim site As Long
    Dim RetChan() As Long
    Dim RetChanCnt As Long
    Dim RetSiteCnt As Long
    Dim Err As String


    If Len(strParam) = 4 Then

        LoggingParam = strParam
    
        lngIndex = 0
        For lngCnt = 4 To 1 Step -1
            TheHdw.Utility.Pins(m_strParamPinName(lngIndex)).state = IIf(Mid(strParam, lngCnt, 1) = "1", bitHigh, bitLow)
            lngIndex = lngIndex + 1
        Next lngCnt

        TheHdw.WAIT m_dblSlewRateOfUtility
        TheHdw.WAIT m_dblSettingTimeOfParam
    
        Call SendStrob
    
    ElseIf Len(strParam) = 5 Then
    
        site = Right(strParam, 1)
        strParam = Mid(strParam, 1, 4)
    
        LoggingParam = strParam
    
        lngIndex = 0
        For lngCnt = 4 To 1 Step -1
            TheExec.DataManager.GetChanList m_strParamPinName(lngIndex), site, chUtil, RetChan, RetChanCnt, RetSiteCnt, Err
            TheHdw.Utility.Chans(RetChan(0)).state = IIf(Mid(strParam, lngCnt, 1) = "1", bitHigh, bitLow)
            lngIndex = lngIndex + 1
        Next lngCnt

        TheHdw.WAIT m_dblSlewRateOfUtility
        TheHdw.WAIT m_dblSettingTimeOfParam
    
        TheExec.DataManager.GetChanList m_strStrobPinName, site, chUtil, RetChan, RetChanCnt, RetSiteCnt, Err
        TheHdw.Utility.Chans(RetChan(0)).state = bitHigh
        m_eStrob = LS_HIGH
        TheHdw.WAIT m_dblStrobPulsWidth
        TheHdw.Utility.Chans(RetChan(0)).state = bitLow
        m_eStrob = LS_LOW
    
    Else
        Exit Property
    End If

End Property


Private Property Let ModeVD(eLogicVal As LOGIC_STATE)
    
    If eLogicVal = LS_HIGH Then
        
        TheHdw.Utility.Pins(m_strModePinNameForVirtualFpga).state = bitHigh
        m_eMode = LS_HIGH
        
    Else
    
        TheHdw.Utility.Pins(m_strModePinNameForVirtualFpga).state = bitLow
        m_eMode = LS_LOW
    
    End If

    TheHdw.WAIT m_dblSlewRateOfUtility

End Property


Private Property Let RstVD(eLogicVal As LOGIC_STATE)
    
    If eLogicVal = LS_HIGH Then
        
        TheHdw.Utility.Pins(m_strResetPinNameForVirtualFpga).state = bitHigh
        m_eReset = LS_HIGH
    Else
    
        TheHdw.Utility.Pins(m_strResetPinNameForVirtualFpga).state = bitLow
        m_eReset = LS_LOW
    End If

    TheHdw.WAIT m_dblSlewRateOfUtility

End Property


Private Property Let EnableVD(eLogicVal As LOGIC_STATE)
    
    If eLogicVal = LS_HIGH Then
        
        TheHdw.Utility.Pins(m_strEnablePinNameForVirtualFpga).state = bitHigh
        m_eEnable = LS_HIGH
        
    Else
    
        TheHdw.Utility.Pins(m_strEnablePinNameForVirtualFpga).state = bitLow
        m_eEnable = LS_LOW
    
    End If

    TheHdw.WAIT m_dblSlewRateOfUtility

End Property


Private Property Let StrobVD(eLogicVal As LOGIC_STATE)
    
    If eLogicVal = LS_HIGH Then
        
        TheHdw.Utility.Pins(m_strStrobPinNameForVirtualFpga).state = bitHigh
        m_eStrob = LS_HIGH
    Else
    
        TheHdw.Utility.Pins(m_strStrobPinNameForVirtualFpga).state = bitLow
        m_eStrob = LS_LOW
    End If

    TheHdw.WAIT m_dblSlewRateOfUtility

End Property


Private Property Let ParamVD(ByVal strParam As String)
    Dim lngCnt As Long
    Dim lngIndex As Long
    Dim site As Long
    Dim RetChan() As Long
    Dim RetChanCnt As Long
    Dim RetSiteCnt As Long
    Dim Err As String


    If Len(strParam) = 4 Then

        LoggingParam = strParam
    
        lngIndex = 0
        For lngCnt = 4 To 1 Step -1
            TheHdw.Utility.Pins(m_strParamPinNameForVirtualFpga(lngIndex)).state = IIf(Mid(strParam, lngCnt, 1) = "1", bitHigh, bitLow)
            lngIndex = lngIndex + 1
        Next lngCnt

        TheHdw.WAIT m_dblSlewRateOfUtility
        TheHdw.WAIT m_dblSettingTimeOfParam
    
        StrobVD = LS_HIGH
        TheHdw.WAIT m_dblStrobPulsWidth
        StrobVD = LS_LOW
    
    ElseIf Len(strParam) = 5 Then
    
        site = Right(strParam, 1)
        strParam = Mid(strParam, 1, 4)
    
        LoggingParam = strParam
    
        lngIndex = 0
        For lngCnt = 4 To 1 Step -1
            TheExec.DataManager.GetChanList m_strParamPinNameForVirtualFpga(lngIndex), site, chUtil, RetChan, RetChanCnt, RetSiteCnt, Err
            TheHdw.Utility.Chans(RetChan(0)).state = IIf(Mid(strParam, lngCnt, 1) = "1", bitHigh, bitLow)
            lngIndex = lngIndex + 1
        Next lngCnt

        TheHdw.WAIT m_dblSlewRateOfUtility
        TheHdw.WAIT m_dblSettingTimeOfParam
    
        TheExec.DataManager.GetChanList m_strStrobPinNameForVirtualFpga, site, chUtil, RetChan, RetChanCnt, RetSiteCnt, Err
        TheHdw.Utility.Chans(RetChan(0)).state = bitHigh
        m_eStrob = LS_HIGH
        TheHdw.WAIT m_dblStrobPulsWidth
        TheHdw.Utility.Chans(RetChan(0)).state = bitLow
        m_eStrob = LS_LOW
    
    Else
        Exit Property
    End If

End Property



'==========================================================================
'
'
'               Private Method
'
'
'==========================================================================

Private Function WaitReady(Optional eFpgaMode As FPGA_MODE = FM_DECODE_FPGA, Optional dblTimeOutmSec As Double = 10, Optional eExectVal As LOGIC_STATE = LS_LOW) As Boolean
    Dim dblNowTime As Double
           
    If (eFpgaMode = FM_DECODE_FPGA) Or (eFpgaMode = FM_BOTH_FPGA) Then
        If WaitReadyByPin(FM_DECODE_FPGA) = False Then
            WaitReady = False
            Exit Function
        End If
    End If
    
    WaitReady = True
    
End Function


Private Function WaitReadyByPin(Optional eFpgaMode As FPGA_MODE, Optional dblTimeOutmSec As Double = 10, Optional eExectVal As LOGIC_STATE = LS_LOW) As Boolean
    Dim dblNowTime As Double
    
    mode = LS_HIGH
    
    WaitReadyByPin = False
    
    dblNowTime = TheExec.timer
    
    Do
        
        If eFpgaMode = FM_DECODE_FPGA Then
            Call GetStatusLogic(m_eLogicState, m_strStatusPinName)
            If IsAllValueEqual(m_eLogicState, eExectVal) = True Then
                mode = LS_LOW
                WaitReadyByPin = True
                Exit Function
            End If
        Else
            Call GetStatusLogic(m_eLogicState_Vir, m_strStatusPinNameForVirtualFpga, 1, eExectVal)
            If IsAllValueEqual(m_eLogicState_Vir, eExectVal) = True Then
                mode = LS_LOW
                WaitReadyByPin = True
                Exit Function
            End If
        End If
        
    
    Loop While (TheExec.timer - dblNowTime) < dblTimeOutmSec

    mode = LS_LOW

End Function

'add T.Minoda 2009/06/29
'change T.Minoda 2009/07/01
Private Sub SetDecTimeOut()

    Dim eDecTimeOut As TIMEOUT
    Dim dblDFrmRate As Double
    Dim dblOfstTime As Double
    Dim strRetActSec As String
    
    LoggingSection = "SetDecTimeOut"
    
    dblDFrmRate = 0.3       'For EeeAuto 145
    dblOfstTime = 1#        'For EeeAuto 145
    If m_blTimeOutEnable = True Then
        If m_dblUserDecTimeOut = 0 Then
            eDecTimeOut = SelectTimeOutBit(dblDFrmRate * (DecAve(m_eAveNumber) + m_eFSkipMode) + dblOfstTime, strRetActSec)
        Else
            eDecTimeOut = SelectTimeOutBit(m_dblUserDecTimeOut, strRetActSec)
        End If
        TheHdw.IDP.CaptureTimeOut = CDbl(strRetActSec) + 3
    Else
        eDecTimeOut = 0
    End If
    
    If m_eDummyVpulse <> 0 Then SetDummyPulse "V", m_eDummyVpulse
    If m_eDummyHpulse <> 0 Then SetDummyPulse "H", m_eDummyHpulse
  
    
    Enable = LS_LOW
    If WaitReady(FM_DECODE_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETDECODETIMEOUT)
        If WaitReady(FM_DECODE_FPGA) = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SETTIMEOUT_TIME))
            If WaitReady(FM_DECODE_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eDecTimeOut))
                If WaitReady(FM_DECODE_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetDecTimeOut:WaitReady After Sending Value of SubCommand [SelectTimeOutValue] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetDecTimeOut:WaitReady After Sending SubCommand [SetDecodeTimeOut] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetDecTimeOut:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SetDecTimeOut:First WaitReady")
        Exit Sub
    End If
    
End Sub

'add T.Minoda 2009/06/29
Private Sub SetDummyPulse(VorH As String, eDummyPulse As DUMMYPULSE)
    
    Dim SubComOfTimeOutPulse As Long
    
    LoggingSection = "SetDummyPulse"

    If VorH = "V" Then
        SubComOfTimeOutPulse = SUBCOMMAND_SETTIMEOUT_VPULSE
    Else
        SubComOfTimeOutPulse = SUBCOMMAND_SETTIMEOUT_HPULSE
    End If
    
    Enable = LS_LOW
    If WaitReady(FM_DECODE_FPGA) = True Then
    
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETDECODETIMEOUT)
        If WaitReady(FM_DECODE_FPGA) = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SubComOfTimeOutPulse))
            If WaitReady(FM_DECODE_FPGA) = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eDummyPulse))
                If WaitReady(FM_DECODE_FPGA) = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetDummyPulse:WaitReady After Sending Value of SubCommand [SelectTimeOutPulse] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetDummyPulse:WaitReady After Sending SubCommand [SelectVPulseOrHPulse] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetDummyPulse:WaitReady After Sending Command is done")
            Exit Sub
        End If
        
    Else
        Call SetLastError(1, "SetDummyPulse:First WaitReady")
        Exit Sub
    End If
    
End Sub

Public Sub SetMIPIrawmode(eMIPIrawmode As mipi_rawmode)                 'add Y.Kimura 2011/03/23

    LoggingSection = "SetMIPIrawmode"

    Enable = LS_LOW

    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(SUBCOMMAND_SETMIPI_RAWMODE))
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(eMIPIrawmode))
                If WaitReady = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SetMIPIrawmode:WaitReady After Sending Value of SubCommand [MIPI RAWMODE SET] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SetMIPIrawmode:WaitReady After Sending SubCommand [MIPI RAWMODE SET] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SetMIPIrawmode:WaitReady After Sending Command is done")
            Exit Sub
        End If

    Else
        Call SetLastError(1, "SetMIPIrawmode:First WaitReady")
        Exit Sub
    End If

End Sub


'add T.Minoda 2009/06/29
'change T.Minoda 2009/07/01
Private Function SelectTimeOutBit(dblDecTime As Double, Optional strRetActSec As String) As TIMEOUT

    If (dblDecTime <= 0) Then
        SelectTimeOutBit = TO_0SEC      'None
        strRetActSec = 0
    ElseIf (dblDecTime > 0 And dblDecTime <= 1) Then
        SelectTimeOutBit = TO_1SEC      '1sec
        strRetActSec = 1
    ElseIf (dblDecTime > 1 And dblDecTime <= 2) Then
        SelectTimeOutBit = TO_2SEC      '2sec
        strRetActSec = 2
    ElseIf (dblDecTime > 2 And dblDecTime <= 3) Then
        SelectTimeOutBit = TO_3SEC      '3sec
        strRetActSec = 3
    ElseIf (dblDecTime > 3 And dblDecTime <= 4) Then
        SelectTimeOutBit = TO_4SEC      '4sec
        strRetActSec = 4
    ElseIf (dblDecTime > 4 And dblDecTime <= 6) Then
        SelectTimeOutBit = TO_6SEC      '6sec
        strRetActSec = 6
    ElseIf (dblDecTime > 6 And dblDecTime <= 8) Then
        SelectTimeOutBit = TO_8SEC      '8sec
        strRetActSec = 8
    ElseIf (dblDecTime > 8 And dblDecTime <= 10) Then
        SelectTimeOutBit = TO_10SEC     '10sec
        strRetActSec = 10
    ElseIf (dblDecTime > 10 And dblDecTime <= 12) Then
        SelectTimeOutBit = TO_12SEC     '12sec
        strRetActSec = 12
    ElseIf (dblDecTime > 12 And dblDecTime <= 14) Then
        SelectTimeOutBit = TO_14SEC     '14sec
        strRetActSec = 14
    ElseIf (dblDecTime > 14 And dblDecTime <= 16) Then
        SelectTimeOutBit = TO_16SEC     '16sec
        strRetActSec = 16
    ElseIf (dblDecTime > 16 And dblDecTime <= 18) Then
        SelectTimeOutBit = TO_18SEC     '18sec
        strRetActSec = 18
    ElseIf (dblDecTime > 18 And dblDecTime <= 30) Then      'change T.Minoda 2009/09/17
        SelectTimeOutBit = TO_30SEC     '30sec
        strRetActSec = 30
    ElseIf (dblDecTime > 30 And dblDecTime <= 60) Then      'change T.Minoda 2009/09/17
        SelectTimeOutBit = TO_60SEC     '60sec
        strRetActSec = 60
    ElseIf (dblDecTime > 60 And dblDecTime <= 180) Then     'change T.Minoda 2009/09/17
        SelectTimeOutBit = TO_180SEC    '180sec
        strRetActSec = 180
    Else                                                    'change T.Minoda 2009/09/17
        SelectTimeOutBit = TO_300SEC    '300sec
        strRetActSec = 300
    End If

End Function


'add T.Minoda 2009/07/01
Private Function DecAve(eDecAveNum As AVENUM_MODE) As Double

    Select Case eDecAveNum
        Case 0
            DecAve = 1
        Case 1
            DecAve = 2
        Case 2
            DecAve = 4
        Case 3
            DecAve = 6
        Case 4
            DecAve = 8
        Case 5
            DecAve = 12
        Case 6
            DecAve = 16
        Case 7
            DecAve = 24
        Case 8
            DecAve = 32
        Case 9
            DecAve = 48
        Case 10
            DecAve = 64
        Case 11
            DecAve = 96
        Case 12
            DecAve = 128
        Case 13
            DecAve = 192
        Case 14
            DecAve = 256
        Case Else
            DecAve = 511
    End Select
    
End Function


Private Function IsUtilityPin(strPin As String) As Boolean
    Dim lngCnt As Long

    For lngCnt = 0 To UBound(m_strUtilityPins)
        If strPin = m_strUtilityPins(lngCnt) Then
            IsUtilityPin = True
            Exit Function
        End If
    Next lngCnt
    
    IsUtilityPin = False

End Function


Private Function IsDchPin(strPin As String) As Boolean
    Dim lngCnt As Long


    For lngCnt = 0 To UBound(m_strPins)
        If strPin = m_strPins(lngCnt) Then
            IsDchPin = True
            Exit Function
        End If
    Next lngCnt
    
    IsDchPin = False
    
End Function


Private Sub SendStrob()

    Strob = LS_HIGH
    TheHdw.WAIT m_dblStrobPulsWidth
    Strob = LS_LOW
    
End Sub


Private Function bitHigh() As UtilBitState
    
    bitHigh = utilBitState0

End Function

Private Function bitLow() As UtilBitState
    
    bitLow = utilBitState1
    
End Function


Private Function IsAllValueEqual(eLogicState() As LOGIC_STATE, eExpectVal As LOGIC_STATE) As Boolean
    Dim lngSiteNum As Long
    Dim blAllMatch As Boolean
    
    blAllMatch = True
    For lngSiteNum = 0 To (TheExec.sites.ExistingCount - 1)
    
        If TheExec.sites.site(lngSiteNum).Active Then
                        
            If eLogicState(lngSiteNum) <> eExpectVal Then blAllMatch = False
            
        End If
    
    Next lngSiteNum
    
    IsAllValueEqual = blAllMatch
    
End Function


Private Sub GetStatusLogic(eLogicState() As LOGIC_STATE, strStatusPinName As String, Optional lngMask As Long = -1, Optional eMaskVal As LOGIC_STATE = LS_LOW)
    Dim adblMeasureAvg() As Double
    Dim lngSiteNum As Long
    
    ReDim eLogicState(TheExec.sites.ExistingCount - 1)
    
''    Call SetPPMUForceCurrent(strStatusPinName, ppmuAutoRange, 0)
    Call SetPPMUForceCurrent(strStatusPinName, ppmu2mA, 0)
    Call MeasurePMU(strStatusPinName, m_lngSampCountOfStatusMeas, adblMeasureAvg)
        
    For lngSiteNum = 0 To (TheExec.sites.ExistingCount - 1)
    
        If TheExec.sites.site(lngSiteNum).Active Then
                        
            If (lngMask = -1) Then
                eLogicState(lngSiteNum) = IIf(adblMeasureAvg(lngSiteNum) > m_dblLimitOfStatusLev, LS_HIGH, LS_LOW)
            Else
                If lngMask And 2 ^ lngSiteNum Then
                    eLogicState(lngSiteNum) = IIf(adblMeasureAvg(lngSiteNum) > m_dblLimitOfStatusLev, LS_HIGH, LS_LOW)
                Else
                    eLogicState(lngSiteNum) = eMaskVal
                End If
            
            End If
        End If
    
    Next lngSiteNum
    
End Sub


Private Sub SetDefaultMode()

    m_eAveNumber = AV_1
    m_eClockMode = CLM_DATA_STROB
    m_eCaptureMode = CM_SMIA_2CH_RAW10
    m_lngVirEnabSite = 1
    m_lngTapMode = 0
    m_eFSkipMode = FS_SKIP_0
    
End Sub


Private Sub ClearLastInfo()

    m_sLastErrorInfo.lngErrorNum = 0
    m_sLastErrorInfo.strErrorName = ""
    m_sLastErrorInfo.strErrortmp = ""
    
End Sub


Private Sub SetLastError(ByVal lngErrIndex As Long, ByVal strFunction As String, Optional strTmp As String = "", Optional lngsite As Long = -1)

    Select Case lngErrIndex
            Case 1
                m_sLastErrorInfo.lngErrorNum = 1
                m_sLastErrorInfo.strErrorName = "Busy TimeOut!"
                m_sLastErrorInfo.strErrorFunction = strFunction
                m_sLastErrorInfo.strErrortmp = GetStatusInfo
            Case 2
                m_sLastErrorInfo.lngErrorNum = 2
                m_sLastErrorInfo.strErrorName = "Capture Error!"
                m_sLastErrorInfo.strErrorFunction = strFunction
                m_sLastErrorInfo.strErrortmp = "Site:" & Trim(str(lngsite)) & "," & strTmp
    End Select

    Call LoggingData

End Sub


Private Function GetStatusInfo() As String
    Dim lngSiteNum As Long
    Dim strInfo As String
    Dim strInfo_vir As String
    
    On Error GoTo err_handle
    
    strInfo = "[Decoder Fpga] -> :"
    For lngSiteNum = 0 To (UBound(m_eLogicState))
    
        If TheExec.sites.site(lngSiteNum).Active Then
            strInfo = strInfo & "Site" & Trim(str(lngSiteNum)) & "=" & IIf(m_eLogicState(lngSiteNum) = LS_HIGH, "Busy", "Ready")
        Else
            strInfo = strInfo & "Site" & Trim(str(lngSiteNum)) & "=NonActive"
        End If
        
        If lngSiteNum < (UBound(m_eLogicState)) Then
            strInfo = strInfo & " : "
        End If
        
    Next lngSiteNum

    If m_blVirtualFpgaEnabled = True Then
        strInfo = strInfo & " / [Virtual Dev Fpga] -> :"
        For lngSiteNum = 0 To (UBound(m_eLogicState))
        
            If TheExec.sites.site(lngSiteNum).Active Then
                strInfo = strInfo & "Site" & Trim(str(lngSiteNum)) & "=" & IIf(m_eLogicState_Vir(lngSiteNum) = LS_HIGH, "Busy", "Ready")
            Else
                strInfo = strInfo & "Site" & Trim(str(lngSiteNum)) & "=NonActive"
            End If
            
        Next lngSiteNum
    
        If lngSiteNum < (UBound(m_eLogicState)) Then
            strInfo = strInfo & " : "
        End If
        
    End If

    GetStatusInfo = strInfo
    
    Exit Function
    
err_handle:
    strInfo = VBA.String(TheExec.sites.ExistingCount, "?")
    GetStatusInfo = strInfo
    
End Function


Private Function Convert10to2(Value As Long, Optional lngPatting As Long = 4, Optional site As Long = -1) As String
    Dim lngBit As Long
    Dim strData As String
    Dim lngPatCnt As Long

    Do Until (Value < 2 ^ lngBit)
        If (Value And 2 ^ lngBit) <> 0 Then
            strData = "1" & strData
        Else
            strData = "0" & strData
        End If

        lngBit = lngBit + 1
    Loop

    lngPatCnt = (lngPatting - Len(strData))

    If lngPatCnt > 0 Then
        strData = VBA.String(lngPatCnt, "0") & strData
    End If

    If site <> -1 Then
        Convert10to2 = strData & CStr(site)
    Else
        Convert10to2 = strData
    End If

End Function


Private Function Convert10to16(Value As Long) As String
    Dim strData As String
    
    Select Case Value
        Case 0 To 9
            strData = Trim(str(Value))
        Case 10 To 15
            strData = VBA.Chr(VBA.Asc("A") + Value - 10)
        Case Else
            strData = "?"
    End Select
    
    Convert10to16 = strData

End Function


Private Function ConvertRevHexStrToLng(ByVal strVer As String) As Long
    Dim strToken() As String
    Dim strBuff As String
    Dim lngCnt As Long
    Dim lngVer As Long
    
    strToken = VBA.Split(strVer, ".")
    
    lngVer = 0
    strBuff = ""
    For lngCnt = 0 To UBound(strToken)
        strBuff = strBuff & strToken(lngCnt)
    Next lngCnt

    ConvertRevHexStrToLng = val("&H" & strBuff)

End Function


Private Sub LogFileInit()
    Dim lngCnt As Long
    Dim strFnam As String
    Dim strFnam_sr As String
    Dim strFnam_de As String
    Dim strSDir As String
    
    Select Case m_eLoggingLocation = LL_SYSTEM_TEMP_DIRECTORY
        Case LL_JOB_DIRECTORY
            strSDir = IIf(VBA.Right(ThisWorkbook.Path, 1) = "\", ThisWorkbook.Path, ThisWorkbook.Path & "\") & LOGING_DIR
            m_strLogDirPath = IIf(VBA.Right(ThisWorkbook.Path, 1) = "\", ThisWorkbook.Path, ThisWorkbook.Path & "\") & LOGING_DIR & "\" & TheHdw.Computer.Name
        Case LL_SYSTEM_TEMP_DIRECTORY
            m_strLogDirPath = SystemTempDirectory
            strSDir = IIf(VBA.Right(m_strLogDirPath, 1) = "\", m_strLogDirPath, m_strLogDirPath & "\") & LOGING_DIR
            m_strLogDirPath = IIf(VBA.Right(m_strLogDirPath, 1) = "\", m_strLogDirPath, m_strLogDirPath & "\") & LOGING_DIR & "\" & TheHdw.Computer.Name
        Case Else
            strSDir = IIf(VBA.Right(ThisWorkbook.Path, 1) = "\", ThisWorkbook.Path, ThisWorkbook.Path & "\") & LOGING_DIR
            m_strLogDirPath = IIf(VBA.Right(ThisWorkbook.Path, 1) = "\", ThisWorkbook.Path, ThisWorkbook.Path & "\") & LOGING_DIR & "\" & TheHdw.Computer.Name
    End Select
    
    m_strLogFilePath = m_strLogDirPath & "\" & LOGING_FILENAME
    strFnam = VBA.Left(m_strLogFilePath, VBA.InStrRev(m_strLogFilePath, ".") - 1)
    
    
    If VBA.Dir(strSDir, vbDirectory) = "" Then
        MkDir strSDir
    End If
    
    If VBA.Dir(m_strLogDirPath, vbDirectory) = "" Then
        MkDir m_strLogDirPath
    End If
    
    Call RefreshLogFile

End Sub


Private Sub RefreshLogFile()
    Dim lngCnt As Long
    Dim strFnam As String
    Dim strFnam_sr As String
    Dim strFnam_de As String
    
    If VBA.Dir(m_strLogFilePath, vbNormal) <> "" Then
        
        If VBA.FileLen(m_strLogFilePath) > m_lngLoggingFileSize Then
            
            
            strFnam = Left(m_strLogFilePath, VBA.InStrRev(m_strLogFilePath, ".") - 1)
            
            For lngCnt = m_lngLoggingFileNum - 1 To 0 Step -1
                    
                If lngCnt > 0 Then
                    strFnam_sr = strFnam & "_" & Trim(str(lngCnt)) & ".txt"
                    strFnam_de = strFnam & "_" & Trim(str(lngCnt + 1)) & ".txt"
                Else
                    strFnam_sr = m_strLogFilePath
                    strFnam_de = strFnam & "_" & Trim(str(lngCnt + 1)) & ".txt"
                End If
                If VBA.Dir(strFnam_sr, vbNormal) <> "" Then
                    
                    FileCopy strFnam_sr, strFnam_de
                    
                End If
                
            Next lngCnt
            
            Kill m_strLogFilePath
            
        End If
        
    End If

End Sub


Private Sub LoggingData()
    Dim fn As Long
    Dim strDate As String
    Dim strMsg As String
    
    If m_blLoggingEnabled Then
        strDate = Format(Now, "yyyy/mm/dd_hh:mm:ss") & ","
        strMsg = strDate & "ErrNum:" & Trim(str(m_sLastErrorInfo.lngErrorNum)) & _
                                ",ErrName:" & m_sLastErrorInfo.strErrorName & _
                                ",ErrFunc:" & m_sLastErrorInfo.strErrorFunction & _
                                ",ErrTestInstance:" & TheExec.DataManager.InstanceName & _
                                ",ErrNote:" & m_sLastErrorInfo.strErrortmp
        If m_eLoggingMode <> LM_SUCCESS_ONLY Then
            
            fn = FreeFile
            
            Open m_strLogFilePath For Append As #fn
            
                Print #fn, strMsg
                
            Close #fn
            
        End If
        
        Call RefreshLogFile
    
    End If
    
    If m_blRunTimeErrorEnable Then
        Call TheExec.ErrorLogMessage(strMsg)
        Call TheExec.ErrorReport
    Else
        Call TheExec.ErrorMessageToLogfile(strMsg)
    End If
        
End Sub


Private Function GetRegister(ByVal lngBitLength As Long, ByRef lngRegisterVal() As Long, ByVal strFunction As String) As Boolean
    Dim lngSiteNum As Long
    Dim lngCnt As Long
    Dim eLogicState() As LOGIC_STATE
    
    GetRegister = False
    
    For lngSiteNum = 0 To TheExec.sites.ExistingCount - 1
        
        If TheExec.sites.site(lngSiteNum).Active Then lngRegisterVal(lngSiteNum) = 0
                                
    Next lngSiteNum
    
    For lngCnt = 1 To lngBitLength
        
        mode = LS_LOW               'Chang SubCommand
        
        Param = Convert10to2(1)     'Send Index.it Increment Register bit(Lsb First)
        If WaitReady = False Then
            Call SetLastError(1, strFunction & ":WaitReady After Sending Value is done")
            Exit Function
        End If
                
'        Mode = LS_LOW               'Chang content of Status to Register
        
        Call GetStatusLogic(eLogicState, m_strStatusPinName)    'Read Register bit
        For lngSiteNum = 0 To TheExec.sites.ExistingCount - 1
            
            If TheExec.sites.site(lngSiteNum).Active Then
                
                lngRegisterVal(lngSiteNum) = lngRegisterVal(lngSiteNum) + IIf(eLogicState(lngSiteNum) = LS_HIGH, 2 ^ (lngCnt - 1), 0)
                
            End If
            
        Next lngSiteNum
    
    Next lngCnt
    
    GetRegister = True
    
End Function


Private Sub MeasurePMU(astrPinName As String, alngSamples As Long, ByRef adblMeasureAvg() As Double)
    Dim lngRetChanArray() As Long
    Dim lngRetChanCnt As Long
    Dim lngRetSiteCnt As Long
    Dim strRetErrMsg As String
    
    Dim dblRetValue() As Double
    
    
    Call TheExec.DataManager.GetChanListForSelectedSites(astrPinName, chIO, lngRetChanArray(), lngRetChanCnt, lngRetSiteCnt, strRetErrMsg)
    
    TheHdw.PPMU.Samples = alngSamples
    
    Call TheHdw.PPMU.Chans(lngRetChanArray()).MeasureVoltages(dblRetValue())
    
    TheHdw.PPMU.Samples = 1
    
    Call AverageValue(lngRetChanArray(), alngSamples, dblRetValue(), adblMeasureAvg())

End Sub


Private Sub SetPPMUForceCurrent(astrPinName As String, aenmIRange As PpmuIRange, adblForceCurrent As Double)

    TheHdw.PPMU.Pins(astrPinName).ForceCurrent(aenmIRange) = adblForceCurrent

End Sub


Private Sub AverageValue(alngChans() As Long, alngSamples As Long, adblMeasVal() As Double, ByRef adblAvgVal() As Double)
    Dim lngSiteNum As Long
    Dim lngLoopCounter As Long
    Dim i As Long
    
    ReDim adblAvgVal(TheExec.sites.ExistingCount - 1)
    
    For lngSiteNum = 0 To (TheExec.sites.ExistingCount - 1)
    
        If TheExec.sites.site(lngSiteNum).Active Then
            
            For i = 0 To alngSamples - 1
                adblAvgVal(lngSiteNum) = adblAvgVal(lngSiteNum) + adblMeasVal((UBound(alngChans()) + 1) * i + lngLoopCounter)
            Next i
            
            adblAvgVal(lngSiteNum) = adblAvgVal(lngSiteNum) / alngSamples
            
            lngLoopCounter = lngLoopCounter + 1
        End If
    
    Next lngSiteNum

End Sub


Private Sub Class_Initialize()

    'Set Pin Name
    m_strStatusPinName = DEFAULT_STATUS_PINNAME
    m_strParamPinName(0) = DEFAULT_PARAM0Bit_PINNAME
    m_strParamPinName(1) = DEFAULT_PARAM1Bit_PINNAME
    m_strParamPinName(2) = DEFAULT_PARAM2Bit_PINNAME
    m_strParamPinName(3) = DEFAULT_PARAM3Bit_PINNAME
    m_strModePinName = DEFAULT_MODE_PINNAME
    m_strStrobPinName = DEFAULT_STROB_PINNAME
    m_strEnablePinName = DEFAULT_ENABLE_PINNAME
    m_strResetPinName = DEFAULT_RESET_PINNAME
    m_strStatusPinNameForVirtualFpga = DEFAULT_STATUS_PINNAMEFORVIRTUALFPGA
    m_strParamPinNameForVirtualFpga(0) = DEFAULT_PARAM0Bit_PINNAMEFORVIRTUALFPGA
    m_strParamPinNameForVirtualFpga(1) = DEFAULT_PARAM1Bit_PINNAMEFORVIRTUALFPGA
    m_strParamPinNameForVirtualFpga(2) = DEFAULT_PARAM2Bit_PINNAMEFORVIRTUALFPGA
    m_strParamPinNameForVirtualFpga(3) = DEFAULT_PARAM3Bit_PINNAMEFORVIRTUALFPGA
    m_strModePinNameForVirtualFpga = DEFAULT_MODE_PINNAMEFORVIRTUALFPGA
    m_strStrobPinNameForVirtualFpga = DEFAULT_STROB_PINNAMEFORVIRTUALFPGA
    m_strEnablePinNameForVirtualFpga = DEFAULT_ENABLE_PINNAMEFORVIRTUALFPGA
    m_strResetPinNameForVirtualFpga = DEFAULT_RESET_PINNAMEFORVIRTUALFPGA
    
    'Set Sampling Count In Meas Status
    m_lngSampCountOfStatusMeas = DEFAULT_SAMPLINGCOUNT_STATUSMEAS
    
    'Set AC Timing
    m_dblLimitOfStatusLev = DEFAULT_LIMIT_STATUSLEVEL
    m_dblSettingTimeOfParam = DEFAULT_SETTING_TIME_PARAM
    m_dblSlewRateOfUtility = DEFAULT_UTIL_SLEW_RATE
    m_dblStrobPulsWidth = DEFAULT_STROBPULS_WIDTH
    m_dblRstPulsWidth = DEFAULT_RSTPULS_WIDTH
    m_dblEnbPulsWidth = DEFAULT_ENBPULS_WIDTH
    
    'Default Mode
    Call SetDefaultMode
    Call ClearLastInfo
    
    ReDim m_strUtilityPins(0)
    ReDim m_strPins(0)
    
    m_blPinsTableCreadedFlg = False
    m_blVirtualFpgaEnabled = False
    m_blLoggingEnabled = True
    m_blRunTimeErrorEnable = False
    m_eLoggingMode = LM_ERROR_ONLY
    m_eLoggingLocation = LL_JOB_DIRECTORY
    m_lngLoggingFileSize = DEFAULT_LOGING_FILEMAXSIZE
    m_lngLoggingFileNum = DEFAULT_LOGING_FILEMAXCNT
    m_blTimeOutEnable = False           'add T.Minoda 2009/06/29
    m_dblUserDecTimeOut = 0             'add T.Minoda 2009/06/29
    m_eDummyVpulse = 0                  'add T.Minoda 2009/06/29
    m_eDummyHpulse = 0                  'add T.Minoda 2009/06/29
    m_blInitTapFlag = False             'add T.Minoda 2009/09/17
        
    Call LogFileInit
    
End Sub


Private Sub Class_Terminate()
    
    On Error Resume Next
    
    Erase m_eLogicState
    Erase m_eLogicState_Vir
    Erase m_strUtilityPins
    Erase m_strPins
    
    On Error GoTo 0
    
End Sub

' add Y.Kimura 2011/04/01
Private Sub SMIACapture_ignore_RevisionCheck(strFpgaVersion As String, strSMIAignoreVersion As String)

    Dim lngVersion_index As Long
    Dim strSMIAignoreVersionArray() As String
    
    ''' Convert comma to array
    strSMIAignoreVersionArray = Split(strSMIAignoreVersion, ",")
    
    For lngVersion_index = 0 To UBound(strSMIAignoreVersionArray)
        If strFpgaVersion = strSMIAignoreVersionArray(lngVersion_index) Then
            m_blSMIACapture_ignore_Revision = True
'            MsgBox "SMIA ignore revision MATCH! FPGA Rev:" & strFpgaVersion
            Exit For
        End If
    Next lngVersion_index
        
End Sub


Public Sub SelectTxR(Optional txmode As Long = 3)

    LoggingSection = "SelectTxR"

    Enable = LS_LOW

    If WaitReady = True Then
        mode = LS_HIGH
        Param = Convert10to2(COMMAND_SETEXPANDDATA)
        If WaitReady = True Then
            mode = LS_LOW
            Param = Convert10to2(CLng(11))
            If WaitReady = True Then
                mode = LS_LOW
                Param = Convert10to2(CLng(txmode))
                If WaitReady = True Then
                    Exit Sub
                Else
                    Call SetLastError(1, "SelectTxR:WaitReady After Sending Value of SubCommand [MIPI RAWMODE SET] is done")
                    Exit Sub
                End If
            Else
                Call SetLastError(1, "SelectTxR:WaitReady After Sending SubCommand [MIPI RAWMODE SET] is done")
                Exit Sub
            End If
        Else
            Call SetLastError(1, "SelectTxR:WaitReady After Sending Command is done")
            Exit Sub
        End If

    Else
        Call SetLastError(1, "SelectTxR:First WaitReady")
        Exit Sub
    End If

End Sub

#End If

