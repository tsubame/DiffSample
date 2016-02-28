Attribute VB_Name = "XLibEeeNaviConst"
'概要:
'   パラメータのデータラベル定義モジュール
'
'目的:
'   パラメータのリード/ライトで利用するデータラベル
'
'   Revision History:
'   Data        Description
'   2008/12/12　作成
'   2009/06/15  Ver1.01アドイン展開に伴う変更
'             　■仕様変更
'               NOT_DEFINEラベルをJOB側の既存ラベルと共有
'
'作成者:
'   0145206097
'
Option Explicit

Public Const SUPPLIER_NAME = "SupplierName"
Public Const IS_TOOL_CONTAIN = "IsToolContain"

Public Const TOOL_NAME = "ToolName"
Public Const IS_SHT_CONTAIN = "IsSheetContain"
Public Const IS_SHT_UNIQUE = "IsSheetUnique"
Public Const IS_VISIBLE_TOOL = "IsVisibleTool"
Public Const IS_CATEGORIZE = "IsCategorize"
Public Const NAME_LOCATION = "NameLocation"

Public Const sheet_name = "SheetName"
Public Const SHEET_PARENT_NAME = "SheetParentName"
Public Const IS_SHT_ACTIVE = "IsSheetActive"
Public Const IS_SHT_DELETED = "IsSheetDeleted"

Public Const END_OF_FILE = "EOF"
Public Const END_OF_DATA = "EOD"

Public Const WILD_CARD = "*"
Public Const SHEET_MISSING = "<Deleted>"
'Public Const NOT_DEFINE = "NotDefine"
Public Const DEF_NAME_MAP = "B1"
Public Const NOT_WORKSHEET = "NotWorksheet"
