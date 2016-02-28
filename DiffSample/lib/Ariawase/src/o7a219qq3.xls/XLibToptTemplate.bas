Attribute VB_Name = "XLibToptTemplate"
'概要:
'   ToptFrameWorkテンプレートが利用するTheImageTestオブジェクトのラッパー関数群
'
'   Revision History:
'       Data        Description
'       2010/04/28  撮像テストを実行する機能を実装した
'       2010/05/12  プログラムコードを整理した
'       2010/05/31  Error処理を変更した
'       2010/06/11  プログラムコードを整理した
'
'作成者:
'   0145184346
'

Option Explicit

Public Function SetScenario(argc As Long, argv() As String) As Integer
'内容:
'   撮像テストを実行するための準備をする
'
'パラメータ:
'   [argc]    In  指定された引数の数
'   [argv()]  In  指定された引数の配列データ
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  撮像テストを実行するための準備をする  #####
    SetScenario = TheImageTest.SetScenario


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    SetScenario = TL_ERROR
    Exit Function


End Function

Public Function Execute(argc As Long, argv() As String) As Integer
'内容:
'   撮像テストを実行する
'
'パラメータ:
'   [argc]    In  指定された引数の数
'   [argv()]  In  指定された引数の配列データ
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  撮像テストを実行する  #####
    Execute = TheImageTest.Execute


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    Execute = TL_ERROR
    Exit Function


End Function
