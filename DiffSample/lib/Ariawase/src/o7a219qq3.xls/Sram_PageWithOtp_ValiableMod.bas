Attribute VB_Name = "Sram_PageWithOtp_ValiableMod"
Option Explicit

'----- SRAM Infomation (Adjust) ---------------------------
Public Const Ef_Bist_Rd_En_Width As Integer = 1                                     'OTPが持っている1冗長イネーブル分のBIT数
Public Const Ef_Bist_Rd_Addr_Width As Integer = 9                                   'OTPが持っている1冗長アドレス分のBIT数
Public Const Ef_Bist_Rd_Data_Width As Integer = 8                                   'OTPが持っている1冗長データ分のBIT数
Public Const MAX_EF_BIST_RD_BIT As Byte = 4                                         '最大補正Memory数
Public Const BitWidth_OtpRep_Parity As Integer = 4                                  'OTPが持っているFUSE冗長パリティー分のBIT数
Public Const BitWidth_OtpRep_Addr As Integer = 7                                    'OTPが持っているFUSE冗長アドレス分のBIT数
Public Const BitWidth_OtpRep_EN As Integer = 1                                      'OTPが持っているFUSE冗長イネーブル分のBIT数
Public Const BitWidth_OtpRep_Data As Integer = 1                                    'OTPが持っているFUSE冗長Data分のBIT数
Public Const Bist_Num_Mem As Integer = 12                                           'Memory数
Public Const Bist_Max_Num_Io As Integer = 50                                       '最大IO数（各Memoryが持っているIO数の中の最大IO数を定義）
Public Const RCON_START_Addr As Integer = 41                                       'RCONの最初の転送データBit番号
Public Const RCON_END_Addr As Integer = 1                                           'RCONの最後の転送データBit番号
Public Const SRAMRD_OUTPUT_PIN As String = "Ph_FSTROBE"                      'パターンOUTPUT(BIST) PIN
Public Const RCON_ChainType As String = "Descending"           'RCONのチェーンナンバーが昇順(Ascending)か降順(Descending)か
Public Const RCON_FirstInfoType As String = "MEMID_1st"   'RCONのMemory情報の最初のBitがMEMID(MEMID_1st)かFAILINFO(FAILINFO_1st)か
Public Const RCON_FailInfoType As String = "Ascending"     'RCONのFAILINFOナンバーが昇順(Ascending)か降順(Descending)か

'----- SRAM Infomation (Fix) ------------------------------
Public BIST_NUM_IO(Bist_Num_Mem + 1) As Integer                                     '各Memoryが持っているIO数を格納する変数　(格納作業はSRAM試験の最初で行う)
Public BIST_RED_TYPE(Bist_Num_Mem + 1) As Integer                                   '各Memoryが冗長IOを保持しているかを格納する変数　(格納作業はSRAM試験の最初で行う)
Public BIST_IO_EN_NO(Bist_Num_Mem + 1) As Integer                                   '各Memoryの接続チェーン（RCON）ナンバーを格納する変数　(格納作業はSRAM試験の最初で行う)
Public BIST_FAIL_REG(nSite, Bist_Num_Mem + 1, Bist_Max_Num_Io + 1) As Byte          '全PreSRAM試験の不良Memoryと不良I/O情報をまとめて格納する変数
Public BIST_FAIL_IO_NO(nSite, Bist_Num_Mem + 1) As Integer                          '全PreSRAM試験後の冗長可能と判定した不良I/Oナンバーを格納する変数
Public EF_BIST_REPAIR_DATA(nSite, 2 ^ Ef_Bist_Rd_Addr_Width) As Byte                'RCONの各Memoryに対する冗長情報(ENと不良I/Oの情報)を格納する変数（配列番号はそのままRCON番号となる）
Public RepairMemoryCount(nSite) As Long                                             '冗長必要メモリー個数格納変数
Public Ef_Enbl_Addr(nSite, MAX_EF_BIST_RD_BIT) As Integer                           '冗長イネーブルデータ[Dec]
Public Ef_Rcon_Addr(nSite, MAX_EF_BIST_RD_BIT) As Integer                           '冗長アドレスデータ[Dec]
Public Ef_Repr_Data(nSite, MAX_EF_BIST_RD_BIT) As Integer                           '冗長データのデータ[Dec]
Public EF_BIST_ROLLCALL_INDEX() As Integer                                          'SRAM冗長RollCallパターンの各期待値判定のVector情報格納変数
Public Const Rep_Bist_Data_Len As Integer = (RCON_START_Addr + RCON_END_Addr) + 1   'RCONのデータBit長(有効RCON個数+2個)
Public Const SramRepBlowLabel As String = "START"                                   'SRAM冗長BLOWパターンスタートラベル名
Public Const SramRepVerifyLabel As String = "START"                                 'SRAM冗長VERIFYパターンスタートラベル名
Public Const BitWidth_OtpRep_Total As Integer = BitWidth_OtpRep_Parity + _
                                                BitWidth_OtpRep_Addr + _
                                                BitWidth_OtpRep_EN + _
                                                BitWidth_OtpRep_Data                'OTPが持っているFUSE冗長TotalのBIT数

Public Const TBL_CELL_PAT As Long = 2
Public Const TBL_CELL_LIST_STRow As Long = 5
Public Const TBL_DIR_CELL_Row As Long = 2
Public Const TBL_DIR_CELL_Col As Long = 3
Public Const TBL_INDEX_CYCLE As Long = 0        'TBLファイル1行をスペースで区切った時の、1個目の文字列（カウントが0始まり）
Public Const TBL_INDEX_BIT As Long = 2          'TBLファイル1行をスペースで区切った時の、3個目の文字列（カウントが0始まり）
Public Const TBL_INDEX_MACRO As Long = 5        'TBLファイル1行をスペースで区切った時の、6個目の文字列（カウントが0始まり）
Public dirTblFile As String                     'TBLファイルを置いているディレクトリ



'##### TBL INFOMATION VARIABLE ###################################################################
Public Type FAIL_CYCLE_INFO
    CycleNo As Long         'FailCycle番号
    MemoryNo As Long        'FailCycleに対応するBIT(I/O)番号
    IoNo As Long            'FailCycleに対応するマクロ番号
End Type

Public Type TBL_LIST_INFO
    PatFileName As String   'PatGpで指定されるパターンファイル名
    TblFileName As String   'PatFileに対応するTBLファイル名
    FailInfo() As FAIL_CYCLE_INFO   'TBLファイルの中身
End Type

Public TblInfo() As TBL_LIST_INFO


'    変数            構造体①          構造体②
' TblInfo(X0)                                    : SRAMBISTの種類数(＝パターン数＝TBLファイル数)分、配列が用意される
'        (X1)
'        (X2)-------PatFileName                  : この配列に割り当てられるパターンの名前
'               |---TblFileName                  : このパターンの情報が記載されているTBLファイルの名前
'               |---FailInfo(Y0)                 : このTBLファイルの詳細情報。冗長に関する期待値Vector数分、配列が用意される
'               |---FailInfo(Y1)
'               |---FailInfo(Y2)
'               |---FailInfo(Y3)-------CycleNo   : このVectorに割り当てられるCycle番号
'                                  |---MemoryNo  : このVectorに割り当てられるSRAMメモリ番号
'                                  |---IoNo      : このVectorに割り当てられるSRAMメモリ内のIO番号

'#################################################################################################


'----- Flag -----------------------------------------------
Public Flg_SRAM_BLOW As Integer                                                     'SRAM冗長ON/OFF フラグ
Public Flg_SramDebug As Integer                                                     'SRAMデバッグログ吐き出しフラグ（0:ログ吐き出し無し　1:ログ吐き出し有り）
Public Bist_Alpg_Fail_Flag(nSite) As Byte                                           'ALPG Failフラグ
Public Bist_Repairable_Flag(nSite) As Byte                                          'Repair ON/OFF フラグ
Public Flg_ActiveSite_SramRep(nSite) As Double                                      'SRAM冗長Blow時の
Public Flg_PostSramRun As Boolean
Public Flg_SramBlankNg(nSite) As Long


Public Sub ValiableSet_SramDesignInfo_IO()

'+++ Test Infomation +++++++++++++++++++++++++++++++
'各Memoryが持っているIO数と冗長セルの有無情報を格納
'RED_TYPE = 1 : 冗長IO有り
'RED_TYPE = 0 : 冗長IO無し
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
'各Memoryの接続チェーン（RCON）ナンバーを格納
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
'SRAM書変数の初期値設定
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
