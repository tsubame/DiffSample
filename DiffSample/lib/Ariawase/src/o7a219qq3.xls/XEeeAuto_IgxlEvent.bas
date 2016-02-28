Attribute VB_Name = "XEeeAuto_IgxlEvent"
Option Explicit
'   2012/12/21  H.Arikawa
'               IGXL_OnProgramLoadedを追加。
'   2013/10/22  H.Arikawa
'               オフラインモードで動作する際にcsv読み込みにいかないように変更。


'Validationの終了タイミングに実行されるインターポーズファンクション

Public Sub IGXL_OnProgramValidated()

    If APMU_CheckFailSafe_f = False Then
        ThisWorkbook.Saved = True
        Application.Quit
    End If
    
    'テスト開始時のインターポーズファンクションでEeeJob側の初期化が実行される
    'OffsetSheetが空の場合、EeeJobの初期化でエラーになるため、バリデーションのタイミングで実行する
    If TheExec.TesterMode = testModeOffline Then Flg_Simulator = 1
    Call JobEnvInit
    If Not TheExec.TesterMode = testModeOffline Then
        Call GetCsvFileName
        Call ReadOffsetFile
        Call WriteOffsetManager
        Call ReadOptFile '上書きしてしまうので
    End If
        
End Sub

Public Sub IGXL_OnProgramLoaded()
'PatGrpsシートの"TSBName"のセルの色づけを行う
    Call PatGrpsColorMake
    
End Sub

