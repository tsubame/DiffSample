Attribute VB_Name = "PALS_Common_Mod"
Option Explicit

'==========================================================================
' モジュール名：  PALS_Common_mod.bas
' 概要        ：  PALS全体で共通に使用する関数群
' 備考        ：  なし
' 更新履歴    ：  Rev1.0      2010/09/30　新規作成        K.Sumiyashiki
'==========================================================================

'###########debug!!!!!!!!!!!!
'Public Const nSite As Long = 3
'Public Const Sw_Node As Long = 65
'Public Const g_MaxPalsCount As Long = 100
'###########debug!!!!!!!!!!!!


Public Declare Sub mSecSleep Lib "kernel32" Alias "Sleep" (ByVal lngmSec As Long)
Public Declare Function sub_PalsCopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public PALS As csPALS
Public blnRunPals As Boolean

Public Const PALSNAME As String = "PALS     ParameterAuto-adjustLinkSystem"
Public Const PALSVER As String = "1.50beta"

Public g_ErrorFlg_PALS As Boolean       'パルスでのエラーを示すフラグ(エラー発生時にTrueに変更)
Public Const PALS_ERRORTITLE As String = "- PALS Error -"

Public g_RunAutoFlg_PALS As Boolean       '自動連携機能動作中フラグ　True:自動起動中、False：通常起動

Public blnPALS_ANI  As Boolean
Public blnPALS_STOP As Boolean

Public Const PALS_OPTTARGET      As String = "OptTarget"
Public Const PALS_OPTIDENTIFIER  As String = "OptIdentifier"
Public Const PALS_OPTJUDGELIMIT  As String = "OptJudgeLimit"

Public Const PALS_LOOPCATEGORY1  As String = "CapCategory1"
Public Const PALS_LOOPCATEGORY2  As String = "CapCategory2"
Public Const PALS_LOOPJUDGELIMIT As String = "LoopJudgeLimit"

Public Const PALS_WAITADJFLG     As String = "WaitAdjFlg"

Public Const PALS_CHECKROW   As Integer = 1
Public Const PALS_CONTENTROW As Integer = 2
Public Const EXCEL_MAXCOLUMN As Integer = 256

Public PALS_ParamFolder As String
Public Const PALS_PARAMFOLDERNAME As String = "PALS_Params"
Public Const PALS_PARAMFOLDERNAME_VOLT As String = "PALS_Volt"
Public Const PALS_PARAMFOLDERNAME_WAVE As String = "PALS_Wave"
Public Const PALS_PARAMFOLDERNAME_WAIT As String = "PALS_Wait"
Public Const PALS_PARAMFOLDERNAME_OPT As String = "PALS_Opt"
Public Const PALS_PARAMFOLDERNAME_LOOP As String = "PALS_Loop"
Public Const PALS_PARAMFOLDERNAME_BIAS As String = "PALS_Bias"
Public Const PALS_PARAMFOLDERNAME_TRACE As String = "PALS_Trace"

Type PALS_TOOL_LIST
    PalsAdj     As Boolean
    VoltageAdj  As Boolean
    WaveAdj     As Boolean
    WaitAdj     As Boolean
    OptAdj      As Boolean
    LoopAdj     As Boolean
    BiasAdj     As Boolean
    TraceAdj     As Boolean
End Type


Public FLG_PALS_DISABLE As PALS_TOOL_LIST
Public FLG_PALS_RUN     As PALS_TOOL_LIST

'Public objLoadedJob     As Object

'****************************************
'****************  定数  ****************
'****************************************
'シート名の定義
Public Const FLOW_TABLE     As String = "Flow Table"
Public Const TEST_INSTANCES As String = "Test Instances"
Public Const TESTCONDITION  As String = "TestCondition"

'>>> 2011/5/6 M.Imamura
Public PinSheetname As String                               'Chans or ChannelMap SheetName
Public Const PinSheetnameChans = "Chans"                    'Chans SheetName
Public Const PinSheetnameChannel = "Channel Map"            'ChannelMap SheetName
'<<< 2011/5/6 M.Imamura

Public Const ReadSheetName = "Power-Supply Voltage"         'Read  SheetName
Public Const ReadSheetNameInfo = "Power-Supply Pin Info"    'Read  SheetName
Public Const OutPutSheetname = "Voltage Backup"             'Write SheetName

Public Const WaveSetupSheetName   As String = "WaveAdjustSetup"
Public Const WaveResultSheetName  As String = "WaveAdjustResult"

Public Const OptResultSheetName  As String = "OptAdjustResult"

Public Const WaitResultSheetName  As String = "WaitAdjustResult"

Public Const CONDSHTNAME As String = "ConditionSetTable"
Public Const ACQTBLSHTNAME As String = "Image ACQTBL"

Public Const intOscAdd            As Integer = 11

'FlowTable読み取り用の定数
'以降のFTはFlowTableの略
Public Const FT_LABEL_X      As Integer = 2
Public Const FT_START_Y      As Integer = 7
Public Const FT_TNAME_X      As Integer = 9
Public Const FT_OPCODE_X     As Integer = 7
Public Const FT_PARAMETER_X  As Integer = 8
Public Const FT_BIN_X        As Integer = 12
Public Const FT_TNUM_X       As Integer = 10
Public Const FT_LASTROW_NAME As String = "set-device"
Public Const FT_SURGE_NAME   As String = "D_SURGE"

'TestInstances読み取り用の定数
'以降のTIはTestInstancesの略
Public Const TI_START_Y      As Integer = 6
Public Const TI_TESTNAME_X   As Integer = 2
Public Const TI_LOWLIMIT_X   As Integer = 14
Public Const TI_HIGHLIMIT_X  As Integer = 15
Public Const TI_UNIT_X       As Integer = 17
Public Const TI_CATEGORY1_X  As Integer = 19
Public Const TI_CATEGORY2_X  As Integer = 20
Public Const TI_JUDGELIMIT_X As Integer = 21
Public Const TI_ARG2_X       As Integer = 16

'TestCondition読み取り用の定数
'以降のTCはTestConditionの略
Public Const TC_START_Y         As Integer = 5
Public Const TC_CONDINAME_X     As Integer = 2
Public Const TC_PROCEDURENAME_X As Integer = 3
Public Const TC_ARG1_X          As Integer = 4
Public Const TC_SWNODE_X        As Integer = 1

'データログの最終行を示す文字列
Public Const DATALOG_END As String = "========================================================================="

'データログの項目一覧を示す文字列
Public Const DATALOG_INDEX As String = " Number  Site Result   Test Name       Pin       Channel Low            Measured       High           Force          Loc"
Public Const DATALOG_INDEX2 As String = " Number  Site Result   Test Name       Pin        Channel Low            Measured       High           Force          Loc"

'データログの各項目位置を検索する際に仕様する文字列
Public Const SITE_POSI      As String = "Site"
Public Const RESULT_POSI    As String = "Result"
Public Const TESTNAME_POSI  As String = "Test Name"
Public Const PIN_POSI       As String = "Pin"
Public Const MEASURED_POSI  As String = "Measured"
Public Const HIGH_POSI      As String = "High"
Public Const CHAN_POSI      As String = "Channel"

'>>>2011/10/3 M.IMAMURA コンピュータネームの検索INDEX
Public Const TESTERNAME_INDEX      As String = "      Node Name:"
'<<<2011/10/3 M.IMAMURA コンピュータネームの検索INDEX

'データログの各項目位置の保存を行う構造体
Public Type DatalogPosition
    SiteStart     As Integer    'site情報スタート位置
    SiteCount     As Integer    'site情報の最大文字数
    TestNameStart As Integer    'テスト名のスタート位置
    TestNameCount As Integer    'テスト名の最大文字数
    MeasuredStart As Integer    '特性値のスタート位置
    MeasuredCount As Integer    '特性値の最大文字数
    PinNameStart  As Integer    '[Pin]のスタート位置
    PinNameCount  As Integer    '[Pin]の最大文字数
'>>>2011/05/12 K.SUMIYASHIKI ADD
    ResultStart  As Integer     '[Result(PASS/FAIL情報)]のスタート位置
    ResultCount  As Integer     '[Result(PASS/FAIL情報)]の最大文字数
'<<<2011/05/12 K.SUMIYASHIKI ADD
End Type

'単位換算用係数
'Private Const TERA As Double = 1000000000000#      'テラ
'Private Const GIGA As Long = 1000000000            'ギガ
Private Const MEGA   As Long = 1000000              'メガ
Private Const KIRO   As Long = 1000                 'キロ
Private Const MILLI  As Double = 0.001              'ミリ
Private Const MAICRO As Double = 0.000001           'マイクロ
Private Const NANO   As Double = 0.000000001        'ナノ
Private Const PIKO   As Double = 0.000000000001     'ピコ
Private Const FEMTO  As Double = 0.000000000000001  'フェムト

'データログ名(フルパス)
Public g_strOutputDataText As String


Private Const GC_NUMBER   As String = "Number"
Private Const GC_SITE     As String = "Site"
Private Const GC_RESULT   As String = "Result"
Private Const GC_TESTNAME As String = "Test Name"
Private Const GC_PIN      As String = "Pin"
Private Const GC_CHANNEL  As String = "Channel"
Private Const GC_LOW      As String = "Low"
Private Const GC_MEASURED As String = "Measured"
Private Const GC_HIGH     As String = "High"
Private Const GC_FORCE    As String = "Force"
Private Const GC_LOC      As String = "Loc"

Public Const SET_WAIT     As String = "xxSetWait"
Public Const SET_AVERAGE  As String = "xxSetAverage"
Public Const ACQUIRE_MODE As String = "xxAcquireMode"

Private colDataIndex As New Collection        'データログ特性値のインデックス検索用文字列を格納するコレクション

'>>>2011/05/12 K.SUMIYASHIKI ADD
Public Type ActiveCheck
    Enable As Boolean                '各サイトの状態(Activeかどうか)を格納している変数(Active⇒True)の定義
End Type

Public Type ActiveSiteInformation
    site(nSite) As ActiveCheck       '各サイトの状態を格納する構造体(サイト数分の配列で定義)の定義
End Type

Public g_ActiveSiteInfo As ActiveSiteInformation    '各サイトの状態を格納する構造体
'<<<2011/05/12 K.SUMIYASHIKI ADD

Public CategoryData As csPALS_LoopMain    'csPALS_LoopMainクラスの定義
'>>>2011/08/29 M.IMAMURA ADD
Public Const gblnForCis As Boolean = True
Public Flg_StopPMC_PALS As Boolean
Public Enum Enm_ErrFileBank
    Enm_ErrFileBank_LOCAL
    Enm_ErrFileBank_SERVER
End Enum
Public Const FILEBANK_LOCALPATH     As String = "C:\ERROR_LOG_"
'<<<2011/08/29 M.IMAMURA ADD

Public Sub Pause(interval As Single)
    Dim T1 As Single
    T1 = timer
    Do
        DoEvents
    Loop While timer - T1 < interval
End Sub

Public Sub RunPALS(Optional PalsRunNormal As Boolean = True)
    
On Error GoTo errPALSRunPALS
    
    PALS_ParamFolder = ThisWorkbook.Path & "\" & PALS_PARAMFOLDERNAME
    blnPALS_ANI = PalsRunNormal

    Call sub_PalsFileCheck

    If Sw_Node = 0 Then
        Call sub_errPALS("Sw_Node=0!! Please Check Your Condition!!", "0-2-01-5-02")
        Exit Sub
    End If

    'パルスのエラーフラグ初期化
    'パルス内でエラーが発生した場合、Trueに変更
    g_ErrorFlg_PALS = False
    
    'TOOL CHECK
    FLG_PALS_DISABLE.BiasAdj = True
    FLG_PALS_DISABLE.WaitAdj = True
    FLG_PALS_DISABLE.LoopAdj = True
    FLG_PALS_DISABLE.OptAdj = True
    FLG_PALS_DISABLE.TraceAdj = True
    FLG_PALS_DISABLE.VoltageAdj = True
    FLG_PALS_DISABLE.WaveAdj = True
    
    Set PALS = Nothing
    Set PALS = New csPALS
    
    If g_ErrorFlg_PALS Then
        End
    End If
    
    frm_PALS.Show
    
    g_ErrorFlg_PALS = False

Exit Sub

errPALSRunPALS:
    Call sub_errPALS("PALS initialize Failed at 'RunPALS'", "0-2-01-0-03")

End Sub

'Start Tester Run
Public Sub sub_exec_run()
    TheExec.RunTestProgram
End Sub

'Set Measure Condition [Do All]
Public Sub sub_exec_DoAll(blnDoAll As Boolean)
    TheExec.RunOptions.DoAll = blnDoAll
End Sub

'Do Optini
Public Sub sub_run_Optini()
    Call OptIni
End Sub



'********************************************************************************************
' 名前: sub_set_datalog
' 内容: IG-XLのデータログ設定を行う
' 引数: blnSetLog      :True⇒関数内でデータログ名の設定を行う
'       blnSetLog      :False⇒関数内でデータログ名の初期化を行う
'       strFileHeader  :設定ファイル名の先頭に付与する文字列(例:LoopAdjData)
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_set_datalog(ByVal blnSetLog As Boolean, Optional strFileHeader1 As String = vbNullString, Optional strFileHeader2 As String = vbNullString)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_set_datalog

    With TheExec.Datalog
        If blnSetLog = True Then
            'データログ名の設定(例:LoopAdjData_q7a163xa2_tool_debug_#65_20100927_170600.txt)
            g_strOutputDataText = PALS_ParamFolder & "\" & strFileHeader1 & "\" & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node) & "\" & strFileHeader2 & "_" & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & _
                                    "_#" & CStr(Sw_Node) & "_" & Format(Date, "yyyymmdd") & "_" & Format(TIME, "hhmmss") & ".txt"
            'Set Output txt Log
            .Setup.DatalogSetup.TextOutput = True
            'Set Output File
            .Setup.DatalogSetup.TextOutputFile = g_strOutputDataText
        Else
            'Set Output txt Log
            .Setup.DatalogSetup.TextOutput = False
            'Set Output File
            .Setup.DatalogSetup.TextOutputFile = vbNullString
            'Set EndLot
            .Setup.LotSetup.EndLot = True
            'Data Apply
            .ApplySetup
        End If
    End With

Exit Sub

errPALSsub_set_datalog:
    Call sub_errPALS("Set datalog name error at 'sub_set_datalog'", "0-2-02-0-04")

End Sub


'********************************************************************************************
' 名前: sub_ReadDatalog
' 内容: 一回分の測定データの読み取りを行う
' 引数: lngNowLoopCnt :測定回数
'       intFileNo     :オープンファイルのファイルNo
'       DatalogPosi   :データログの各特性値記入位置を格納する構造体
'       blnContFail   :Continue On FailかStop On Failかを判断するフラグ
'                      True ⇒FAIL項目データも読み取る
'                      False⇒FAIL項目データは除外する
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/20　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_ReadDatalog(ByVal lngNowLoopCnt As Long, ByVal intFileNo As Integer, _
                            ByRef DatalogPosi As DatalogPosition, ByVal blnContFail As Boolean)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_ReadDatalog

    Dim strbuf As String           'テキストファイルから読み込んだ文字列を格納
    Dim blnFlgRead As Boolean      '読込制御用フラグ
    Dim blnFlgGetPos As Boolean    'ポジション制御用フラグ
    
    'フラグ初期化
    blnFlgRead = False
    blnFlgGetPos = False

    'DATALOG_ENDで設定された行が来るまで繰り返し
    Do Until blnFlgRead
        'ファイルから１行読み込む
        Line Input #intFileNo, strbuf
        
        '一回目の測定時のみ、データログの各項目位置を確認
        If (lngNowLoopCnt = 1) And (blnFlgGetPos = False) Then
            
            Call sub_InputDataIndex
            
            'データログ項目一覧の行の場合、strBufに一行分の文字列を格納
'            Do While (strBuf <> DATALOG_INDEX And strBuf <> DATALOG_INDEX2)
            Do While Not sub_CheckDatalogIndex(strbuf)
                '>>>2011/10/3 M.IMAMURA コンピュータネームの取得
                If InStr(1, strbuf, TESTERNAME_INDEX) > 0 Then
                    PALS.CommonInfo.g_strTesterName = Trim$(Mid(strbuf, InStr(1, strbuf, ":") + 1))
                End If
                '<<<2011/10/3 M.IMAMURA コンピュータネームの取得
                Line Input #intFileNo, strbuf
            Loop
            
'            '取り込み項目の位置検索
'            Call sub_GetDataPosition(strbuf, DatalogPosi)
            
            'ポジション制御用フラグをTrueに変更
            blnFlgGetPos = True
            
            '次行をstrBufに格納
            Line Input #intFileNo, strbuf

            '取り込み項目の位置検索
            Call sub_GetDataPosition(strbuf, DatalogPosi)
        End If
        
        'データログの値を取得
        If (Mid(strbuf, DatalogPosi.PinNameStart, 5) = "Empty") And (InStr(1, strbuf, "NGTEST") = 0) And _
                (InStr(1, strbuf, "WATCHS") = 0) And Len(strbuf) > 0 Then
            '特性値取得関数
            Call sub_GetDatalogData(lngNowLoopCnt, strbuf, DatalogPosi, blnContFail)
        End If
        
        'DATALOG_ENDで設定された行に到達したら、読込制御用フラグをTrueに変更
        If strbuf = DATALOG_END Then
            blnFlgRead = True
        End If
    Loop

Exit Sub

errPALSsub_ReadDatalog:
    Call sub_errPALS("Read datalog data error at 'sub_ReadDatalog'", "0-2-03-0-05")

End Sub


'********************************************************************************************
' 名前: sub_InputDataIndex
' 内容: データログ特性値のインデックスを検索する際に使用する文字列をコレクションに追加。
'       コレクションは、データログから特性値を読み取る際に、インデックスを検出する為に使用する。
' 引数: なし
' 戻値: なし
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_InputDataIndex()

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

    'データログ特性値のインデックス検索用文字列をコレクションに追加
    With colDataIndex
        .Add GC_NUMBER
        .Add GC_SITE
        .Add GC_RESULT
        .Add GC_TESTNAME
        .Add GC_PIN
        .Add GC_CHANNEL
        .Add GC_LOW
        .Add GC_MEASURED
        .Add GC_HIGH
        .Add GC_FORCE
        .Add GC_LOC
    End With

End Sub


'********************************************************************************************
' 名前: sub_CheckDatalogIndex
' 内容: 引数で渡された文字列に、コレクション内の全文字列が含まれているかチェックする。
'       データログ特性値のインデックス行を検出する為に使用する。
' 引数: strBuf  :データログの1行
' 戻値: True    :コレクション内の文字列が全て含まれている場合
'     : False   :コレクション内の文字列が一つでも欠けている場合
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_CheckDatalogIndex(ByRef strbuf As String) As Boolean

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_CheckDatalogIndex

    '初期化
    sub_CheckDatalogIndex = False
        
    Dim varIndexName As Variant     'ForEachで使用する変数
    
    'データログのインデックスに含まれている文字列分繰り返し
    For Each varIndexName In colDataIndex
        
        'データログインデックスに含まれている文字列が一つでも欠けていればFalseを返し、関数を抜ける
        If InStr(1, strbuf, varIndexName) = 0 Then
            Exit Function
        End If
    Next varIndexName

    '全て含まれている場合、Trueを返す
    sub_CheckDatalogIndex = True

Exit Function

errPALSsub_CheckDatalogIndex:
    Call sub_errPALS("Check DatalogIndex error at 'sub_CheckDatalogIndex'", "0-2-04-0-06")

End Function



'********************************************************************************************
' 名前: sub_GetDataPosition
' 内容: 測定データ位置の取得を行う
' 引数: strBuf      :一行分のデータログ(データのインデックスが含まれているデータ)
'       DatalogPosi :データログの各特性値記入位置を格納する構造体
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/20　新規作成   K.Sumiyashiki
'            Rev1.1      2011/05/12　処理変更   K.Sumiyashiki
'                                    ⇒IG-XLのバージョンでResultFormatのデータ位置が変化する点に対応
'********************************************************************************************
Private Sub sub_GetDataPosition(ByRef strbuf As String, ByRef DatalogPosi As DatalogPosition)
    
    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALSsub_GetDataPosition
    
    '各項目の位置検索
    
    Dim index(10) As Integer        '各データの開始位置を格納する配列
    Dim IndexCnt As Integer         '配列番号をインクリメントする為の変数
    '初期化
    IndexCnt = 0

    Dim blnCheckStart As Boolean    '各データが開始しているか判断するフラグ
    '初期化
    blnCheckStart = False

    Dim i As Long   'strbufの文字を検索する際に、文字数をインクリメントしていく為のLOOPカウンタ
    
    'strbufに格納されている文字列を1文字ずつ、順番に検索
    For i = 1 To Len(strbuf) - 1
        '各インデックスの開始位置を判断
        If blnCheckStart = False Then
            '検索した文字位置が空白でない箇所を、各データが始まった位置と判断
            If Mid(strbuf, i, 1) <> " " Then
                blnCheckStart = True
                index(IndexCnt) = i
                IndexCnt = IndexCnt + 1
            End If

        '連続2文字が空白の場合、各データが終わったと判断
        ElseIf blnCheckStart = True Then
            If Mid(strbuf, i, 1) = " " And Mid(strbuf, i + 1, 1) = " " Then
                blnCheckStart = False
            End If
        End If
    Next i

'Index配列の中身
'Index(0) -> Number
'Index(1) -> Site
'Index(2) -> Result
'Index(3) -> Test Name
'Index(4) -> Pin
'Index(5) -> Channel
'Index(6) -> Low
'Index(7) -> Measured
'Index(8) -> High
'Index(9) -> Force
'Index(10)-> Loc

    With DatalogPosi
        .SiteStart = index(1)                               'site情報スタート位置
        .SiteCount = index(2) - .SiteStart                  'site情報の最大文字数
        .TestNameStart = index(3)                           'テスト名のスタート位置
        .TestNameCount = index(4) - .TestNameStart          'テスト名の最大文字数
        .MeasuredStart = index(7)                           '特性値のスタート位置
        .MeasuredCount = index(8) - .MeasuredStart          '特性値の最大文字数
        .PinNameStart = index(4)                            'Pinのスタート位置
        .PinNameCount = index(5) - .PinNameStart            'Pinの最大文字数
'>>>2011/05/12 K.SUMIYASHIKI ADD
        .ResultStart = index(2)                             'Result(PASS/FAIL情報)のスタート位置
        .ResultCount = .TestNameStart - .ResultStart        'Result(PASS/FAIL情報)の最大文字数
'<<<2011/05/12 K.SUMIYASHIKI ADD
    End With

Exit Sub

errPALSsub_GetDataPosition:
    Call sub_errPALS("Get Data Position error at 'sub_GetDataPosition'", "0-2-05-0-07")

End Sub


'********************************************************************************************
' 名前: sub_GetDatalogData
' 内容: データログから特性値の抽出を行い、単位変換後、変数に入力
' 引数: lngNowLoopCnt  :測定回数
'       strBuf         :一行分のデータログ(特性値が含まれているデータ)
'       DatalogPosi    :データログの各特性値記入位置を格納する構造体
'       blnContFail   :Continue On FailかStop On Failかを判断するフラグ
'                      True ⇒FAIL項目データも読み取る
'                      False⇒FAIL項目データは除外する
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/20　新規作成   K.Sumiyashiki
'            Rev1.1      2011/05/12　処理追加   K.Sumiyashiki
'                                    ⇒PASS/FAILの情報取得処理追加
'********************************************************************************************
Private Sub sub_GetDatalogData(ByVal lngNowLoopCnt As Long, ByRef strbuf As String, ByRef DatalogPosi As DatalogPosition, ByVal blnContFail As Boolean)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_GetDatalogData

    Dim strTestName As String           'TestName(項目名)
    Dim strSiteNo As String             'サイト番号
    Dim dblMeasured As Double           '特性値
    Dim intIndex As Integer             'LoopTestInfoのインデックス
'>>>2011/05/12 K.SUMIYASHIKI ADD
    Dim strPassFail As String           'PASS/FAILの情報
'<<<2011/05/12 K.SUMIYASHIKI ADD

    With DatalogPosi
        'テスト名の読み取り
        strTestName = RTrim$(Mid$(strbuf, .TestNameStart, .TestNameCount))
        If strTestName = FT_SURGE_NAME Then
            Exit Sub
        End If

        'サイト情報の読み取り
        strSiteNo = RTrim$(Mid$(strbuf, .SiteStart, .SiteCount))
                
        '特性値の読み取り
        'sub_ConvertUnitで単位換算を行っている(ex:"510 m" -> 0.51)
        dblMeasured = sub_ConvertUnit(RTrim$(Mid$(strbuf, .MeasuredStart, .MeasuredCount)))
        
        'TestnameInfoListコレクションを使用し、該当テスト項目のインデックスを取得
        intIndex = PALS.CommonInfo.TestnameInfoList(strTestName)
        
'>>>2011/05/12 K.SUMIYASHIKI UPDATE
'>>>2011/06/16 M.IMAMURA blnContFail Add.
        'Stop On Failの場合、そのサイトの有効フラグをFalseに変更し、以降のデータを読み取らない
        If blnContFail = False Then
            If g_ActiveSiteInfo.site(CInt(strSiteNo)).Enable = False Then
                PALS.CommonInfo.TestInfo(intIndex).site(val(strSiteNo)).Enable(lngNowLoopCnt) = False
                Exit Sub
            End If
        End If
'<<<2011/06/16 M.IMAMURA blnContFail Add.
        
        '指定項目のPASS判断フラグを一旦Trueに初期化
        PALS.CommonInfo.TestInfo(intIndex).site(val(strSiteNo)).Enable(lngNowLoopCnt) = True

        '特性値書き込み
        PALS.CommonInfo.TestInfo(intIndex).site(val(strSiteNo)).Data(lngNowLoopCnt) = val(dblMeasured)

        strPassFail = RTrim$(Mid$(strbuf, .ResultStart, .ResultCount))
        '指定項目がFAILしていた場合、有効判断フラグをFalseへ変更
        If strPassFail = "FAIL" Then
            g_ActiveSiteInfo.site(CInt(strSiteNo)).Enable = False
'''            PALS.CommonInfo.TestInfo(intIndex).Site(Val(strSiteNo)).Enable(lngNowLoopCnt) = False
        Else
'''            PALS.CommonInfo.TestInfo(intIndex).Site(Val(strSiteNo)).Enable(lngNowLoopCnt) = True
        End If
'<<<2011/05/12 K.SUMIYASHIKI UPDATE
    End With

Exit Sub

errPALSsub_GetDatalogData:
    Call sub_errPALS("Get datalog data error at 'sub_GetDatalogData'", "0-2-06-0-08")

End Sub


'********************************************************************************************
' 名前: sub_ConvertUnit
' 内容: 特性値の単位変換を行う
'       "0.51 m" => "0.00051"
' 引数: strBuf  :特性値データ　ex)"0.51 m"
' 戻値: 単位変換後の特性値
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/20　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_ConvertUnit(ByRef strbuf As String) As Double

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_ConvertUnit

    Dim SplitData As Variant    '分割された文字列を格納

    'strBufを空白で分割
    'ex)"0.51 m" => SplitData(0)="0.51",SplitData(1)="m"
    SplitData = Split(strbuf, " ")
    
    '単位換算
    '(UBoundにすることで、空白が2つ以上あった場合でも対応可)
    Select Case SplitData(UBound(SplitData))
        Case "M"
            sub_ConvertUnit = val(SplitData(0)) * MEGA
        
        Case "MV"
            sub_ConvertUnit = val(SplitData(0)) * MEGA
            
        Case "K"
            sub_ConvertUnit = val(SplitData(0)) * KIRO
        
        Case "KV"
            sub_ConvertUnit = val(SplitData(0)) * KIRO
        
        Case "V"
            sub_ConvertUnit = val(SplitData(0))
        
        Case "mV", "mW"
            sub_ConvertUnit = val(SplitData(0)) * MILLI
        
        Case "m"
            sub_ConvertUnit = val(SplitData(0)) * MILLI
                
        Case "u"
            sub_ConvertUnit = val(SplitData(0)) * MAICRO
        
        Case "uV", "uW"
            sub_ConvertUnit = val(SplitData(0)) * MAICRO
        
        Case "n"
            sub_ConvertUnit = val(SplitData(0)) * NANO
        
        Case "nV", "nW"
            sub_ConvertUnit = val(SplitData(0)) * NANO
        
        Case "p"
            sub_ConvertUnit = val(SplitData(0)) * PIKO
        
        Case "pV", "pW"
            sub_ConvertUnit = val(SplitData(0)) * PIKO
        
        Case "f"
            sub_ConvertUnit = val(SplitData(0)) * FEMTO
        
        Case "fV"
            sub_ConvertUnit = val(SplitData(0)) * FEMTO
        
        Case Else
            sub_ConvertUnit = val(val(SplitData(0)))
        
    End Select

Exit Function

errPALSsub_ConvertUnit:
    Call sub_errPALS("Convert Unit error at 'sub_ConvertUnit'", "0-2-07-0-09")

End Function


Public Function all_mod_export()
    Call make_moddir
    
    Call export_module("frm_PALS.frm")
    Call export_module("frm_PALS_BiasAdj_Main.frm")
    Call export_module("frm_PALS_LoopAdj_Main.frm")
    Call export_module("frm_PALS_OptAdj_Main.frm")
    Call export_module("frm_PALS_TraceAdj_Main.frm")
    Call export_module("frm_PALS_VoltAdj_Main.frm")
    Call export_module("frm_PALS_WaitAdj_Main.frm")
    Call export_module("frm_PALS_WaveAdj_Confirm.frm")
    Call export_module("frm_PALS_WaveAdj_Doing.frm")
    Call export_module("frm_PALS_WaveAdj_Main.frm")
    Call export_module("frm_PALS_WaveAdj_Warning.frm")
    
    Call export_module("Conditionset_Mod_ShutOnly.bas")
    Call export_module("PALS_BiasAdj_Mod.bas")
    Call export_module("PALS_Common_Mod.bas")
    Call export_module("PALS_LoopAdj_Mod.bas")
    Call export_module("PALS_OptAdj_Mod.bas")
    Call export_module("PALS_Sub_Mod.bas")
    Call export_module("PALS_TraceAcq_Mod.bas")
    Call export_module("PALS_TraceAdj_Mod.bas")
    Call export_module("PALS_VoltAdj_Mod.bas")
    Call export_module("PALS_WaitAdj_Mod.bas")

    Call export_module("PALS_WaveAdj_mod_Common.bas")
    Call export_module("PALS_WaveAdj_mod_GetWave.bas")
    Call export_module("PALS_WaveAdj_mod_H.bas")
    Call export_module("PALS_WaveAdj_mod_HShared.bas")
    Call export_module("PALS_WaveAdj_mod_LH.bas")
    Call export_module("PALS_WaveAdj_mod_RG.bas")
    Call export_module("PALS_WaveAdj_mod_Shutter.bas")
    Call export_module("PALS_WaveAdj_mod_TVCfunctions.bas")
    Call export_module("PALS_WaveAdj_mod_VVT.bas")

    Call export_module("PALS_IlluminatorMod.bas")

    Call export_module("csPALS.cls")
    Call export_module("csPALS_Common.cls")
    Call export_module("csPALS_LoopCategoryParams.cls")
    Call export_module("csPALS_LoopMain.cls")
    Call export_module("csPALS_OptCond.cls")
    Call export_module("csPALS_OptCondParams.cls")
    Call export_module("csPALS_TestInfo.cls")
    Call export_module("csPALS_TestInfoParams.cls")

    Call export_module("csPALS_WaveACPSet.cls")
    Call export_module("csPALS_WaveACSet.cls")
    Call export_module("csPALS_WaveAdjust.cls")
    Call export_module("csPALS_WaveDCPSet.cls")
    Call export_module("csPALS_WaveDcSet.cls")
    Call export_module("csPALS_WaveDevicePin.cls")
    Call export_module("csPALS_WaveOscPSet.cls")
    Call export_module("csPALS_WaveOscSet.cls")
    Call export_module("csPALS_WaveResource.cls")


End Function


Public Sub make_moddir()
    Dim modfir As String

    On Error Resume Next

    modfir = ActiveWorkbook.Path & "\bas"
    MkDir modfir

    modfir = ActiveWorkbook.Path & "\frm"
    MkDir modfir

    modfir = ActiveWorkbook.Path & "\cls"
    MkDir modfir

    On Error GoTo 0
End Sub
Public Sub export_module(mymodule As String)

    Dim waveI As Long
    Dim mytype As Integer
    Dim modfir As String
    
    With Workbooks(ActiveWorkbook.Name).VBProject
        'check all project
        For waveI = 1 To .VBComponents.Count
            'delete pin point!!
            'check type & name
            If Right(mymodule, 3) = "bas" Then
                mytype = 1
                modfir = ActiveWorkbook.Path & "\bas"
            End If
            If Right(mymodule, 3) = "cls" Then
                mytype = 2
                modfir = ActiveWorkbook.Path & "\cls"
            End If
            If Right(mymodule, 3) = "frm" Then
                mytype = 3
                modfir = ActiveWorkbook.Path & "\frm"
            End If
            If .VBComponents(waveI).Name = Left(mymodule, Len(mymodule) - 4) And .VBComponents(waveI).Type = mytype Then
                'delete compornents
                If mytype = 1 Then .VBComponents(waveI).Export modfir & "\" & .VBComponents(waveI).Name & ".bas"
                If mytype = 2 Then .VBComponents(waveI).Export modfir & "\" & .VBComponents(waveI).Name & ".cls"
                If mytype = 3 Then .VBComponents(waveI).Export modfir & "\" & .VBComponents(waveI).Name & ".frm"
                Exit For
            End If
        Next waveI
    End With
    
End Sub

'********************************************************************************************
' 名前: sub_ModuleCheck
' 内容: 引数で渡されたモジュールが存在するかチェックを行う
' 引数: mymodule :拡張子付きモジュール名
' 戻値: True  : 一致あり
'       False : 一致なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/10/29　新規作成   M.Imamura
'********************************************************************************************
Public Function sub_ModuleCheck(mymodule As String) As Boolean

    Dim vbc As Object
    Dim lngLoopCnt As Long
    Dim mytype As Integer
    
    sub_ModuleCheck = False
    
    If Right(mymodule, 3) = "bas" Then mytype = 1
    If Right(mymodule, 3) = "cls" Then mytype = 2
    If Right(mymodule, 3) = "frm" Then mytype = 3

    With Workbooks(ActiveWorkbook.Name).VBProject
        For lngLoopCnt = 1 To .VBComponents.Count
            'check type & name
            If .VBComponents(lngLoopCnt).Name = Left(mymodule, Len(mymodule) - 4) And .VBComponents(lngLoopCnt).Type = mytype Then
                sub_ModuleCheck = True
                Exit Function
            End If
        Next lngLoopCnt
    End With
    
    
End Function

'********************************************************************************************
' 名前: sub_SheetNameCheck
' 内容: 引数で渡された名前のシートが存在するかチェックを行う
' 引数: strSheetName :検索シート名
' 戻値: True  : 一致シートあり
'       False : 一致シートなし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Function sub_SheetNameCheck(ByVal strSheetName As String) As Boolean

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_SheetNameCheck
    
    sub_SheetNameCheck = False

    Dim objWorkSheet As Worksheet
    
    For Each objWorkSheet In Worksheets
        If strSheetName = objWorkSheet.Name Then
            sub_SheetNameCheck = True
            Exit For
        End If
    Next

Exit Function

errPALSsub_SheetNameCheck:
    Call sub_errPALS("SheetName check error at 'sub_SheetNameCheck'", "0-2-08-0-10")

End Function


'********************************************************************************************
' 名前: sub_InitCollection
' 内容: コレクションデータの初期化
' 引数: col:コレクション
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_InitCollection(ByRef col As Collection)

    If g_ErrorFlg_PALS Or col.Count = 0 Then
        Exit Sub
    End If

On Error GoTo errPALSsub_InitCollection
    
    Dim i As Long   'ループカウンタ

    'コレクションデータの削除
    For i = col.Count To 1 Step -1
        col.Remove (i)
    Next i

Exit Sub

errPALSsub_InitCollection:
    Call sub_errPALS("Collection remove error at 'sub_InitCollection'", "0-2-09-0-11")

End Sub

'********************************************************************************************
' 名前: sub_errPALS
' 内容: エラー表示及びエラーログ作成
' 引数: strPalsErrMsg:エラー詳細情報
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/10/19　新規作成   M.Imamura
' 更新履歴： Rev1.1      2011/05/06　エラーコードを追加   M.Imamura
' 更新履歴： Rev1.2      2012/03/14　ローカルPCへの保存を追加   M.Imamura
'********************************************************************************************

Public Sub sub_errPALS(ByVal strPalsErrMsg As String, Optional strPalsErrorCode As String = "", Optional enumFileBank As Enm_ErrFileBank = Enm_ErrFileBank_SERVER)
    Dim strPalsErrDescription As String
    
    'PALSエラーフラグ発生
    g_ErrorFlg_PALS = True
    
    'メッセージボックス表示
    If Len(strPalsErrorCode) > 0 Then strPalsErrorCode = " " & strPalsErrorCode
    If Err.Number = 0 Then
        'PALSセルフチェック時のコメント
        strPalsErrDescription = "PALS Check Error"
    Else
        'VBAエラーの詳細
        strPalsErrDescription = Err.Description
    End If
    
    '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    If g_RunAutoFlg_PALS = False Then
        MsgBox "Error" & strPalsErrorCode & " : " & strPalsErrMsg & vbCrLf & "Description : " & strPalsErrDescription, vbExclamation, PALS_ERRORTITLE
    End If
    '<<<2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    
    'エラーログ吐き出し
    Dim intFileNo As Integer
    Dim strOutputDataText As String
    intFileNo = FreeFile

    If enumFileBank = Enm_ErrFileBank_SERVER Then
        If PALS_ParamFolder = "" Then PALS_ParamFolder = ThisWorkbook.Path
        strOutputDataText = PALS_ParamFolder & "\PALS_ERROR_LOG_" & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & _
                                        "_#" & CStr(Sw_Node) & ".txt"
    ElseIf enumFileBank = Enm_ErrFileBank_LOCAL Then 'オート測定停止用
        Flg_StopPMC_PALS = True 'StopPMCで止める
        strOutputDataText = FILEBANK_LOCALPATH & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node) & ".txt"
    End If
    
    Open strOutputDataText For Append As #intFileNo

    Print #intFileNo, "--------------------------------------------"
    Print #intFileNo, "Date         : " & Format(Date, "yyyy/mm/dd") & " " & Format(TIME, "hh:mm:ss")
    Print #intFileNo, "Message      : " & strPalsErrMsg
    Print #intFileNo, "Description  : " & strPalsErrDescription
    Print #intFileNo, "ErrorCode    : " & strPalsErrorCode
    
    Print #intFileNo, ""
    
    Close #intFileNo

End Sub
Public Sub sub_PalsFileCheck(Optional ByVal strPalsCheckDir As String = "")
    Dim strPalsCheckFile As String
    Dim intFileCount As Integer
    Dim intFileNo As Integer

On Error Resume Next
    strPalsCheckFile = PALS_ParamFolder
    If strPalsCheckDir <> "" Then strPalsCheckFile = strPalsCheckFile & "\" & strPalsCheckDir
    
'    With Application.FileSearch
'        .LookIn = strPalsCheckFile
'        .filename = "*.*"
'        .Execute
'        intFileCount = .FoundFiles.count
'
'        If .FoundFiles.count = 0 Then
            MkDir strPalsCheckFile
'        End If
'
'    End With
        
    'Write RunningLogData
    intFileNo = FreeFile
    If strPalsCheckDir <> "" Then
        Open strPalsCheckFile & "\" & strPalsCheckDir & ".log" For Append As #intFileNo
        Print #intFileNo, Format(Date, "yyyy/mm/dd") & " " & Format(TIME, "hh:mm:ss") & " " & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node)
        Close #intFileNo
        strPalsCheckFile = strPalsCheckFile & "\" & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node)
        MkDir strPalsCheckFile
    Else
        Open strPalsCheckFile & "\" & PALS_PARAMFOLDERNAME & ".log" For Append As #intFileNo
        Print #intFileNo, Format(Date, "yyyy/mm/dd") & " " & Format(TIME, "hh:mm:ss") & " " & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node)
        Close #intFileNo
    End If
    

On Error GoTo 0

End Sub


'********************************************************************************************
' 名前: sub_TestingStatusOutPals
' 内容: フォームのフォント色を変更
' 引数: objPalsForm:フォームオブジェクト
'       strPalsMsg:表示する文字列
'       RedColor:Trueでフォント赤(デフォルト:False)
'       BlueColor:Trueでフォント青(デフォルト:False)
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_TestingStatusOutPals(objPalsForm As Object, strPalsMsg As String, Optional RedColor As Boolean = False, Optional BlueColor As Boolean = False)

    On Error GoTo errPALSsub_TestingStatusOutPals
    
    If RedColor = True Then
        objPalsForm.lblProcess.ForeColor = vbRed
    ElseIf BlueColor = True Then
        objPalsForm.lblProcess.ForeColor = vbBlue
    Else
        objPalsForm.lblProcess.ForeColor = vbBlack
    End If
    
    objPalsForm.lblProcess.Caption = strPalsMsg
    DoEvents

    Exit Sub

errPALSsub_TestingStatusOutPals:
    Call sub_errPALS("Status change error at 'sub_TestingStatusOutPals'", "0-2-10-0-12")

End Sub


'********************************************************************************************
' 名前: sub_InitActiveSiteInfo
' 内容: 各サイトのPASS/FAIL情報を初期化。全てTrue(Active状態)に初期化。
' 引数: なし
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2011/05/16　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_InitActiveSiteInfo()

    Dim sitez As Long

    With g_ActiveSiteInfo
        '全サイトを繰り返し
        For sitez = LBound(.site) To UBound(.site)
            '各サイトをActive状態で初期化
            .site(sitez).Enable = True
        Next sitez
    End With

End Sub

'********************************************************************************************
' 名前: ReadCategoryData
' 内容: csPALS_LoopMainのインスタンスを生成
' 引数: なし
' 戻値: なし
' 備考: なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
' 更新履歴： Rev1.1      2011/07/29　PALS_Subから移動   M.Imamura
'********************************************************************************************
Public Sub ReadCategoryData()

'>>>2011/06/02 K.SUMIYASHIKI ADD
    g_ErrorFlg_PALS = False
'<<<2011/06/02 K.SUMIYASHIKI ADD

    Set CategoryData = Nothing
    Set CategoryData = New csPALS_LoopMain

End Sub
'********************************************************************************************
' 名前: ResetPals
' 内容: PALSパラメータを再読み込み
' 引数: なし
' 戻値: なし
' 備考: なし
' 更新履歴： Rev1.0      2011/12/06　新規作成   M.Imamura
'********************************************************************************************
Public Sub ResetPals(Optional ByVal strResetMode As String = "ALL")

    
    Select Case strResetMode
        Case "PALS"
            Set PALS = Nothing
            Set PALS = New csPALS
        Case "ALL"
            Set OptCond = Nothing
            Set OptCond = New csPALS_OptCond
            
            Set CategoryData = Nothing
            Set CategoryData = New csPALS_LoopMain
        
            Call Get_Power_Condition
        
            Call Excel.Application.Run("TimingInit_PALS")
        
        Case "OPT"
            Set OptCond = Nothing
            Set OptCond = New csPALS_OptCond

        Case "TESTCOND"
            Set CategoryData = Nothing
            Set CategoryData = New csPALS_LoopMain

        Case "VOLT"
            Call Get_Power_Condition
        
        Case "TIME"
            Call Excel.Application.Run("TimingInit_PALS")

    End Select
End Sub

Public Function sub_CheckResultFormat() As Boolean
  sub_CheckResultFormat = True
  If TheExec.Datalog.Setup.DatalogSetup.SelectSetupFile = False Then
    sub_CheckResultFormat = False
    Call sub_errPALS("ResultFormat is NotChecked!! Please Select ResultFormat", "0-2-11-2-12")
  End If
End Function
Public Sub sub_OutPutCsv(ByVal InputWorkSheetName As String, ByVal OutPutCSVFName As String, Optional ByVal bln_ShowMsg As Boolean = True)

' Excel Application/Book/Sheetオブジェクト定義
    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlWS As Excel.Worksheet

    Dim intMaxRow As Integer
    Dim intMaxColumn As Integer
    
    Dim intWriteRow As Integer
    Dim intWriteColumn As Integer
    
    '2012/7/9 FullPath Get. M.Imamura
    Dim strOutPutCSVFName_Full As String
    
    If Left(OutPutCSVFName, 2) = "./" Or Left(OutPutCSVFName, 2) = ".\" Then
        strOutPutCSVFName_Full = ThisWorkbook.Path & Mid(OutPutCSVFName, 2, Len(OutPutCSVFName) - 1)
    Else
        strOutPutCSVFName_Full = OutPutCSVFName
    End If
    
    intMaxRow = Worksheets(InputWorkSheetName).UsedRange.Rows.Count
    intMaxColumn = Worksheets(InputWorkSheetName).UsedRange.Columns.Count
    
    On Error GoTo errPALSsub_OutPutCsv

    'BackUp CSV FIle
    '2012/7/9 FunctionNameChanged. M.Imamura FileCopy -> sub_PalsFileCopy
    If sub_PalsFileCopy(strOutPutCSVFName_Full, strOutPutCSVFName_Full & "_" & Format(Date, "yyyymmdd") & "_" & Format(TIME, "hhmmss")) = False Then
        GoTo errPALSsub_OutPutCsv
    End If

    Set xlApp = CreateObject("Excel.Application")

    xlApp.DisplayAlerts = False
    xlApp.Visible = False
    xlApp.ScreenUpdating = False

    Set xlWB = xlApp.Workbooks.Open(strOutPutCSVFName_Full)
    Set xlWS = xlWB.Worksheets(1)

    xlWB.Worksheets(1).Cells.Select
    xlApp.Selection.ClearContents

    For intWriteRow = 1 To intMaxRow
        For intWriteColumn = 1 To intMaxColumn
            xlWS.Cells(intWriteRow, intWriteColumn).Value = Worksheets(InputWorkSheetName).Cells(intWriteRow, intWriteColumn).Value
        Next intWriteColumn
    Next intWriteRow
    xlWS.Cells(1, 1).Font.color = vbRed
    xlWS.Cells(3, 2).Value = "location:"

    xlWB.SaveAs strOutPutCSVFName_Full, xlCSV

    xlWB.Close
    xlApp.Quit

    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    
    If bln_ShowMsg = True Then
        MsgBox "PALS saved [" & InputWorkSheetName & "]Sheet to " & vbCrLf & "  CSV[" & strOutPutCSVFName_Full & "]", vbOKOnly, PALSNAME
    End If
Exit Sub
    
errPALSsub_OutPutCsv:
    Call sub_errPALS("CsvOutPutError " & strOutPutCSVFName_Full & " at 'sub_OutPutCsv'", "0-5-05-6-37")

    If Not (xlWB Is Nothing) Then xlWB.Close
    Set xlWS = Nothing
    Set xlWB = Nothing
    If Not (xlApp Is Nothing) Then xlApp.Quit
    Set xlApp = Nothing
    
End Sub

'2012/7/9 FunctionNameChanged. M.Imamura FileCopy -> sub_PalsFileCopy
Public Function sub_PalsFileCopy(tgtFile As String, newFile As String) As Boolean
    Dim RetNum As Long
    RetNum = sub_PalsCopyFile(tgtFile, newFile, True) 'Arg3 True means overwrite, False means preserve.
    If RetNum Then
        sub_PalsFileCopy = True
    Else
        sub_PalsFileCopy = False
    End If
End Function
