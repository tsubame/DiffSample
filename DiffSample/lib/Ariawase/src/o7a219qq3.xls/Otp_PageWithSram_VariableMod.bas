Attribute VB_Name = "Otp_PageWithSram_VariableMod"
Option Explicit

'■■■ Adjust Variable ■■■

Public Const RejiIn As String = "Ph_SDA"                      'Blowする際のBlow情報入力ピン
Public Const RejiOut As String = "Ph_SDA"                     'RollCallする際のRollCall情報出力ピン
Public Const OtpPageStart As Integer = 0                 'OTPMAPのスタートページナンバー
Public Const OtpPageEnd As Integer = 11                   'OTPMAPのラストページナンバー
Public Const OtpPixOffset_X As Long = -1                  'OTPアドレスと物理アドレスの差異（X方向）
Public Const OtpPixOffset_Y As Long = 25                  'OTPアドレスと物理アドレスの差異（Y方向）

Public Const BitWidthAll_Lot1  As Long = 4                  'Bit幅
Public Page_Lot1(BitWidthAll_Lot1 - 1) As Long                 'OTPの各Bitが、何ページ目なのかを示す変数
Public Bit_Lot1(BitWidthAll_Lot1 - 1) As Long                  'OTPの各Bitが、何Bit目なのかを示す変数

Public Const BitWidthAll_Lot2  As Long = 4                  'Bit幅
Public Page_Lot2(BitWidthAll_Lot2 - 1) As Long                 'OTPの各Bitが、何ページ目なのかを示す変数
Public Bit_Lot2(BitWidthAll_Lot2 - 1) As Long                  'OTPの各Bitが、何Bit目なのかを示す変数

Public Const BitWidthAll_Lot7  As Long = 4                  'Bit幅
Public Page_Lot7(BitWidthAll_Lot7 - 1) As Long                 'OTPの各Bitが、何ページ目なのかを示す変数
Public Bit_Lot7(BitWidthAll_Lot7 - 1) As Long                  'OTPの各Bitが、何Bit目なのかを示す変数

Public Const BitWidthAll_Lot8  As Long = 4                  'Bit幅
Public Page_Lot8(BitWidthAll_Lot8 - 1) As Long                 'OTPの各Bitが、何ページ目なのかを示す変数
Public Bit_Lot8(BitWidthAll_Lot8 - 1) As Long                  'OTPの各Bitが、何Bit目なのかを示す変数

Public Const BitWidthAll_Lot9  As Long = 4                  'Bit幅
Public Page_Lot9(BitWidthAll_Lot9 - 1) As Long                 'OTPの各Bitが、何ページ目なのかを示す変数
Public Bit_Lot9(BitWidthAll_Lot9 - 1) As Long                  'OTPの各Bitが、何Bit目なのかを示す変数

Public Const BitWidthAll_Wafer  As Long = 8                  'Bit幅
Public Page_Wafer(BitWidthAll_Wafer - 1) As Long                 'OTPの各Bitが、何ページ目なのかを示す変数
Public Bit_Wafer(BitWidthAll_Wafer - 1) As Long                  'OTPの各Bitが、何Bit目なのかを示す変数

Public Const BitWidthAll_Chip  As Long = 16                  'Bit幅
Public Page_Chip(BitWidthAll_Chip - 1) As Long                 'OTPの各Bitが、何ページ目なのかを示す変数
Public Bit_Chip(BitWidthAll_Chip - 1) As Long                  'OTPの各Bitが、何Bit目なのかを示す変数

Public Const MaxRepair_Single_CP_FD As Integer = 10               '欠陥補正可能最大数
Public Const BitWidthN_Single_CP_FD As Integer = 8               '検出個数  bit幅
Public Const BitWidthX_Single_CP_FD As Integer = 12               'Xアドレス bit幅
Public Const BitWidthY_Single_CP_FD As Integer = 12               'Yアドレス bit幅
Public Const BitWidthS_Single_CP_FD As Integer = 2               'Source情報 bit幅
Public Const BitWidthD_Single_CP_FD As Integer = 2               'Direction bit幅
Public Const DefRep_SrcType_Single_CP_FD As String = "SrcType1"           'Source情報タイプ
Public Const NgAddress_LeftS_Single_CP_FD As Long = 9            'Couplet 画角端NGアドレス幅(Left Start)
Public Const NgAddress_LeftE_Single_CP_FD As Long = 12            'Couplet 画角端NGアドレス幅(Left End)
Public Const NgAddress_RightS_Single_CP_FD As Long = 3285           'Couplet 画角端NGアドレス幅(Right Start)
Public Const NgAddress_RightE_Single_CP_FD As Long = 3288           'Couplet 画角端NGアドレス幅(Right End)
Public Const BitWidthAll_Single_CP_FD As Long = BitWidthN_Single_CP_FD + (MaxRepair_Single_CP_FD * (BitWidthX_Single_CP_FD + BitWidthY_Single_CP_FD + BitWidthS_Single_CP_FD + BitWidthD_Single_CP_FD))               '欠陥補正に必要なbit数
Public Page_Single_CP_FD(BitWidthAll_Single_CP_FD - 1) As Long              '補正に必要なOTPの各Bitが、何ページ目なのかを示す変数
Public Bit_Single_CP_FD(BitWidthAll_Single_CP_FD - 1) As Long               '補正に必要なOTPの各Bitが、何Bit目なのかを示す変数


Public Const BitWidth_O_TEMP  As Long = 10                 '温度計オフセット bit幅
Public Const BitWidth_S_TEMP  As Long = 10                 '温度計傾き bit幅
Public Const BitWidthAll_TEMP As Long = BitWidth_O_TEMP + BitWidth_S_TEMP                '欠陥補正に必要なbit数
Public Page_TEMP(BitWidthAll_TEMP - 1) As Long               '補正に必要なOTPの各Bitが、何ページ目なのかを示す変数
Public Bit_TEMP(BitWidthAll_TEMP - 1) As Long                '補正に必要なOTPの各Bitが、何Bit目なのかを示す変数

Public Const BitWidthAll_SRAM  As Long = 72                  'Bit幅
Public Page_SRAM(BitWidthAll_SRAM - 1) As Long                 'OTPの各Bitが、何ページ目なのかを示す変数
Public Bit_SRAM(BitWidthAll_SRAM - 1) As Long                  'OTPの各Bitが、何Bit目なのかを示す変数

'■Fix Variable
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Public Const OtpMaxBitParPage As Integer = 512                                          '各ページが持っているBit数の中の最大Bit数(1始まりでカウント)
Public Const OtpInfoSheet_Row_Page As Long = 2                                          'OTP_Infomationシート ページ情報のRow情報スタートオフセット
Public Const OtpInfoSheet_Column_Page As Long = 2                                       'OTP_Infomationシート ページ情報のColumn情報スタートオフセット
Public Const OtpInfoSheet_Row_Bit As Long = 2                                           'OTP_Infomationシート Bit(Dec)情報のRow情報スタートオフセット
Public Const OtpInfoSheet_Column_Bit As Long = 5                                        'OTP_Infomationシート Bit(Dec)情報のColumn情報スタートオフセット
Public Const OtpInfoSheet_Row_BlowInfo As Long = 2                                      'OTP_Infomationシート Blow内容（変動値情報）のRow情報スタートオフセット
Public Const OtpInfoSheet_Column_BlowInfo As Long = 8                                   'OTP_Infomationシート Blow内容（変動値情報）Column情報スタートオフセット
Public Const OtpInfoSheet_Row_Value As Long = 2                                         'OTP_Infomationシート Value(Bin)のRow情報スタートオフセット
Public Const OtpInfoSheet_Column_Value As Long = 9                                      'OTP_Infomationシート Value(Bin)のColumn情報スタートオフセット
Public Const OtpInfoSheet_Row_FF As Long = 2                                            'OTP_Infomationシート FF書き込み情報のRow情報スタートオフセット
Public Const OtpInfoSheet_Column_FF As Long = 11                                        'OTP_Infomationシート FF書き込み情報のColumn情報スタートオフセット
Public Const BitParHex As Integer = 8                                                   '1Hex情報が何bitか
Public Const OtpPageSize As Integer = OtpPageEnd + 1                                    'OTPMAPのページ枚数。番号のスタートが0からであること前提
Public AddrParPage(OtpPageSize - 1) As Integer                                          '各Pageが何Addressか
Public BitParPage(OtpPageSize - 1) As Integer                                           '各Pageが何bitか
Public BlowDataAllBin(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String           '全Blow情報(Bin)
Public BlowDataAllBin2(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String          '全Blow情報(Bin) 固定値用
Public ReadDataAllBin(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String           '全Read情報(Bin)
Public FFBlowInfo(OtpPageSize - 1, OtpMaxBitParPage - 1) As String                      'FFBlow情報(Bin)

Public Const Label_Page_OtpBlow As String = "OTP_Blow_BlowLabel_Page_OtpBlowPage"                                    'BLOWパターン Page情報スタートラベル名
Public Const Label_Page_OtpBlow_Break As String = "OTP_Blow_BlowLabel_Page_OtpBlow_Break"                            'FFBLOWパターン Page情報スタートラベル名
Public Const Label_Page_OtpVerify As String = "OTP_Verify_VerifyLabel_Page_OtpVerifyPage"                            'VERIFYパターン Page情報スタートラベル名
Public Const Label_Page_BlankCheck As String = "OTP_Verify_VerifyLabel_Page_BlankCheckPage"                          'BLANKチェックパターン Page情報スタートラベル名
Public Const Label_Page_OtpFixedValueCheck As String = "OTP_Verify_VerifyLabel_Page_OtpFixedValueCheckPage"          '固定値チェックパターン Page情報スタートラベル名

Public Const Label_OtpBlow As String = "OTP_Blow_BlowLabel_OtpBlowPage"                                              'BLOWパターン Blow情報スタートラベル名
Public Const Label_OtpBlowAuto As String = "OTP_Blow_BlowLabelAuto_OtpBlowPage"                                      'BLOWパターン AutoBlow情報スタートラベル名
Public Const Label_OtpBlow_Break As String = "OTP_Blow_BlowLabel_OtpBlow_Break"                                      'FFBLOWパターン Blow情報スタートラベル名
Public Const Label_OtpBlowAuto_Break As String = "OTP_Blow_BlowLabelAuto_OtpBlow_Break"                              'FFBLOWパターン AutoBlow情報スタートラベル名
Public Const Label_OtpVerify As String = "OTP_Verify_VerifyLabel_OtpVerifyPage"                                      'VERIFYパターン Verify情報スタートラベル名
Public Const Label_BlankCheck As String = "OTP_Verify_VerifyLabel_BlankCheckPage"                                    'BLANKチェックパターン Verify情報スタートラベル名
Public Const Label_OtpFixedValueCheck As String = "OTP_Verify_VerifyLabel_OtpFixedValueCheckPage"                    '固定値チェックパターン Verify情報スタートラベル名

Public Vector_OtpRead(OtpMaxBitParPage - 1) As Integer                            'RollCallパターン各Vector情報格納変数
Public Const ByteParVector_VerifyPat As Long = 9

'----- SRAM Repair Blow Infomation (Fix) -----------------
Public SramBlowDataAllBin(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String           '全Blow情報(Bin) SRAM冗長専用の変数
Public SramReadDataAllBin(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String           '全Read情報(Bin) SRAM冗長専用の変数

'----- Flag -----------------------------------------------
Public Flg_OTP_BLOW As Integer                                                          '測定中のOTPBLOWを実行するかの選択
Public Flg_OtpBlowPage(OtpPageSize - 1) As Boolean                                      'OTPBLOWを実行するパターン（ページ）の選択（True：パターン実行　False：パターン実行無し）
Public Flg_OtpBlowFixValPage As Integer                                                 '固定値をBlowするPage情報を持たせるフラグ。固定値Pageが複数Pageあれば最後の固定値Pageのみを保持。
Public Flg_ModifyPage(OtpPageSize - 1) As Boolean                                       'OTPBLOWの変動値Modifyを行うパターン（ページ）の選択（True：Modify実行　False：Modify実行無し）
Public Flg_ActiveSite_OTP(nSite) As Double
Public FFBlowPage As Long                                                               'FFBLOWを実行するPage情報
Public Flg_ModifyPageSRAM(OtpPageSize - 1) As Boolean                               'SRAM冗長のModifyを行うパターン（ページ）の選択（True：Modify実行　False：Modify実行無し）

