Attribute VB_Name = "XLibIGXLEvent"
'概要:
'   IG-XLイベントで呼び出されるマクロ関数
'
'目的:
'   各タイミングでのライブラリ処理に使用
'   ユーザーへは別マクロ関数を提供する
'
'   Revision History:
'   Data        Description
'   2008/05/20　Eee-JOB V1.21版リリース [XLibIGXLEvent.bas]
'   2009/04/07  Eee-JOB V2.00版リリース [XLibIGXLEvent.bas]
'               ■仕様変更
'               プロパティ・メソッド名称ガイドライン施行に伴う名称変更に対応
'               ■機能追加
'               OnProgramLoadedイベントにEeeNavigationメニューコンストラクタ用関数追加
'   2009/04/21　■仕様変更
'               アドイン作成ガイドラインに伴うアドインプロジェクト名変更に対応
'   2009/06/15  EeeNavigation Ver1.01アドイン展開に伴う変更
'             　■仕様変更
'               ナビゲーションメニューのコンストラクタ用マクロ関数名を変更
'
'作成者:
'   0145206097
'

Option Explicit
'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const CUB_UB_USE = 0    'CUB UBの設定          0：未使用、0以外：使用

'TESTERのInitial時のイベント、StartDatatool実行時にココが実行される。
Public Function OnTesterInitialized() As Long
    
    #If CUB_UB_USE <> 0 Then
    Call XLibJob.InitCub
    #End If

    InitControlShtReader

    '### ユーザーへ提供するマクロ #########################
    On Error Resume Next
    Application.Run ("IGXL_OnTesterInitialized")
    On Error GoTo 0

End Function

'プログラムロード後のイベント
Public Function OnProgramLoaded() As Long

    If TheExec.RunMode = runModeProduction Then
        CheckExaminationMode
    End If

    '### EeeNaviセットアップメニューコンストラクタ ########
    On Error Resume Next
    Application.Run ("XLibEeeNaviConstructor.CreateEeeNaviSetUpMenu")
    On Error GoTo 0

    '### TOPTフレームワーク用 #############################
    Call XLibToptFrameWorkUtility.ResetEeeJobObjects

    '### ユーザーへ提供するマクロ #########################
    On Error Resume Next
    Application.Run ("IGXL_OnProgramLoaded")
    On Error GoTo 0

End Function

'バリデーション開始時のイベント
Public Function OnValidationStart() As Long
    
    '### OnValidationStartで実行する関数群 ################
    XLibJobUtility.RunAtValidationStart
    
    '### ユーザーへ提供するマクロ #########################
    On Error Resume Next
    Application.Run ("IGXL_OnValidationStart")
    On Error GoTo 0

End Function

'バリデーション後のイベント、Validate Job実行後にココが実行される。
Public Function OnProgramValidated() As Long
        
    InitControlShtReader
    ValidateDCTestSenario

    '### TOPTフレームワーク用 #############################
    Call XLibToptFrameWorkUtility.ResetEeeJobSheetObjects
    Call XLibToptFrameWorkUtility.RunAtValidated

    '### ユーザーへ提供するマクロ #########################
    On Error Resume Next
    Application.Run ("IGXL_OnProgramValidated")
    On Error GoTo 0

End Function

'TDRキャリブレーション後のイベント
Public Function OnTDRCalibrated() As Long

    '### ユーザーへ提供するマクロ #########################
    On Error Resume Next
    Application.Run ("IGXL_OnTDRCalibrated")
    On Error GoTo 0

End Function

'JOBプログラム実行開始直後のイベント
Public Function OnProgramStarted() As Long

    '### TOPTフレームワーク用 #############################
    Call XLibToptFrameWorkUtility.RunAtJobStart

    '### ユーザーへ提供するマクロ #########################
    On Error Resume Next
    Application.Run ("IGXL_OnProgramStarted")
    On Error GoTo 0

End Function

'JOBプログラム実行終了直後のイベント
Public Function OnProgramEnded() As Long

    '### TOPTフレームワーク用 #############################
    Call XLibToptFrameWorkUtility.RunAtJobEnd

    '### ユーザーへ提供するマクロ #########################
    On Error Resume Next
    Application.Run ("IGXL_OnProgramEnded")
    On Error GoTo 0

End Function
