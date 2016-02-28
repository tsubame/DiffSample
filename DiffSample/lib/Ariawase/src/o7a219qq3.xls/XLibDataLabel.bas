Attribute VB_Name = "XLibDataLabel"
'概要:
'   リーダー/ライター上のデータラベル定義
'
'目的:
'   リーダー/ライターがデータをやり取りする際のデータラベルを定義する
'
'作成者:
'   0145206097
'
Option Explicit

'### Control Sheet名 ######################################
Public Const JOB_LIST_TOOL = "Job List"
Public Const SHEET_MANAGER_TOOL = "Data Sheet Manager"

'### Job List Sheetで定義するDataTool名 ###################
Public Const FLOW_TABLE_TOOL = "Flow Table"
Public Const TEST_INSTANCES_TOOL = "Test Instances"
Public Const PIN_MAP_TOOL = "Pin Map"
Public Const CHAN_MAP_TOOL = "Channel Map"
Public Const AC_SPECS_TOOL = "AC Specs"
Public Const DC_SPECS_TOOL = "DC Specs"
Public Const PATTERN_SETS_TOOL = "Pattern Sets"
Public Const PATTERN_GROUPS_TOOL = "Pattern Groups"
Public Const BIN_TABLE_TOOL = "Bin Table"
Public Const CHARACTERIZATION_TOOL = "Characterization"
Public Const TEST_PROCEDURES_TOOL = "Test Procedures"

'### Test Instance Sheetで定義するDataTool名 ##############
Public Const TIME_SETSB_TOOL = "Time Sets (Basic)"
Public Const TIME_SETS_TOOL = "Time Sets"
Public Const EDGE_SETS_TOOL = "Edge Sets"
Public Const PIN_LEVELS_TOOL = "Pin Levels"

'### Sheet Manager Sheetで定義するDataTool名 ##############
Public Const DC_SCENARIO_TOOL = "DC Test Scenario"
Public Const DC_PLAYBACK_TOOL = "DC Playback Data"
Public Const OFFSET_TOOL = "Offset Manager"

'### DC Test Scenarioのデータラベル #######################
Public Const TEST_CATEGORY = "TestCategory"
Public Const TEST_ACTION = "TestAction"
Public Const TEST_POSTACTION = "TestPostAction"
Public Const TEST_PINS = "TestPins"
Public Const TEST_PINLIST = "TestPinList"
Public Const TEST_PINTYPE = "TestPinType"

Public Const SET_MODE = "SetMode"
Public Const SET_RANGE = "SetRange"
Public Const SET_FORCE = "SetForce"

Public Const MEASURE_WAIT = "MeasureWait"
Public Const MEASURE_AVG = "MeasureAverage"
Public Const MEASURE_SITE = "MeasureSite"
Public Const MEASURE_LABEL = "MeasureLabel"

Public Const OPERATE_FORCE = "OperateForce"
Public Const OPERATE_RESULT = "OperateResult"

Public Const EXAMIN_FLAG = "ExaminFlag"
Public Const EXAMIN_MODE = "ExaminMode"
Public Const EXAMIN_RANGECHECK = "ExaminRangeCheck"
Public Const EXAMIN_TIMESTAMP = "ExaminTimeStamp"
Public Const EXAMIN_EXECTIME = "ExaminExecTime"
Public Const EXAMIN_RESULT = "ExaminResult"
Public Const EXAMIN_RESULTUNIT = "ExaminResultUnit"
Public Const REPEAT_COUNTER = "RepeatCounter"
Public Const IS_VALIDATE = "RangeValidation"

'### MeasureRange適正チェック用のデータラベル ############
Public Const NO_JUDGE = 0
Public Const VALIDATE_OK = 1
Public Const VALIDATE_NG = 2
Public Const VALIDATE_WARNING = 3
Public Const VALIDATE_OK_NO_JUDGE = 4
Public Const VALIDATE_NG_NO_JUDGE = 5
Public Const VALIDATE_WARNING_NO_JUDGE = 6
Public Const DISABEL_TO_VALIDATION = 7

Public Const INVALIDATION_VALUE = 99999

Public Const BOARD_NAME = "BoardName"
Public Const BOARD_RANGE = "BoardRange"
Public Const BOARD_FORCE = "BoardForce"
Public Const VALIDATE_RESULT = "ValidateResult"

Public Const END_OF_REPORT_DATA = "ReportComment"

'### DC Playback Dataのデータラベル #######################
Public Const PB_CATEGORY = "PbCategory"
Public Const PB_LABEL = "PbLabel"
Public Const PB_LIMIT_HI = "PbLimitHigh"
Public Const PB_LIMIT_LO = "PbLimitLow"
Public Const PB_REF_DATA = "PbRefData"
Public Const PB_DELTA_DATA = "DCDeltaData"

'### Offset Managerのデータラベル #########################
Public Const OFFSET_LABEL = "OffsetLabel"
Public Const OFFSET_COEF = "Coefficient"
Public Const OFFSET_CONS = "Constant"
Public Const TESTER_NUMBER = "TesterNumber"
Public Const END_OF_TESTER_NUM = "OffsetComment"

'### DcTestExaminFormのデータラベル #######################
Public Const MEASURE_FRM = "MeasureFrame"
Public Const RETURN_BTN = "ReturnBtn"
Public Const ACTION_LABEL = "ActionLabel"
Public Const CATEGORY_ID = "CategoryID"
Public Const GROUP_ID = "GroupID"
Public Const SITE_INDEX = "SiteIndex"

'### Job List Sheetのデータラベル #########################
Public Const ACTIVE_JOB_NAME = "Job Name"

'### Test Instance Sheetのデータラベル ####################
Public Const TEST_NAME = "Test Name"
Public Const DC_CATEGORY = "Category @DC Specs"
Public Const DC_SELCTOR = "Selector @DC Specs"
Public Const AC_CATEGORY = "Category @ADC Specs"
Public Const AC_SELCTOR = "Selector @AC Specs"
Public Const TIME_SETS = "Time Sets"
Public Const EDGE_SETS = "Edge Sets"
Public Const PIN_LEVELS = "Pin Levels"
Public Const USERMACRO_LOLIMIT = "Arg0"
Public Const USERMACRO_HILIMIT = "Arg1"
Public Const USERMACRO_JUDGE = "Arg2"
Public Const USERMACRO_UNIT = "Arg3"

'### Data Sheetの名前・バージョン管理用 ###################
Public Const TOOL_NAME_CELL = "$B$1"
Public Const VERSION_CELL = "$A$1"
Public Const CURR_VERSION = "V1.2"

'### その他 ###############################################
Public Const DATA_CHANGED = "DataChanged"
Public Const NOT_DEFINE = "N.D"
Public Const MEASURE_CLASS = "MEASURE_CLASS"
Public Const SETMODE_CLASS = "SETMODE_CLASS"
Public Const DISCONNECT_CLASS = "DISCONNECT_CLASS"
'Public Const ALL_SITE = -1
Public Const ALL_TEST = -1
Public Const NO_SITE = -2
Public Const NO_TEST = -2
Public Const DIGIT_NUMBER = 12
