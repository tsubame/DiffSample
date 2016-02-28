Attribute VB_Name = "XLibDcTestTemplate"
'概要:
'   DCTestScenarioテンプレートが利用するTheDcTestオブジェクトのラッパー関数群
'
'目的:
'   DCTestScenarioテンプレートをアドインで提供するための手段として作成
'   テンプレートからTheDcTestオブジェクトの参照が出来ないため、
'   ラッパー関数を用意しテンプレートからこの関数を呼び出す
'
'   Revision History:
'   Data        Description
'   2008/09/25　評価版リリース
'   2009/04/07　V2.0ライブラリセット用にリリース
'               ■仕様変更
'               ①Eee-JOBライブラリセットのプロパティ･メソッド名称ガイドライン施行に伴う関数名の変更
'               ②①の理由でTheDcTestオブジェクトの各メソッド呼び出しを変更
'               ③オブジェクト名の変更
'
'作成者:
'   0145206097
'
'
'   Ver1.1 2013/02/01 H.Arikawa Ex用にExecuteにdumpPPMUregを追加。

Option Explicit

Public Function SetScenario(argc As Long, argv() As String) As Long
    On Error GoTo ErrHandler
    If argc > 1 Then
        Err.Raise 9999, "XLibDcTestTemplate.SetScenario", "Two Or More Arguments Are Not Supported !"
        GoTo ErrHandler
    End If
    TheDcTest.SetScenario argv(0)
    SetScenario = TL_SUCCESS
    Exit Function
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    SetScenario = TL_ERROR
End Function

Public Function Execute(argc As Long, argv() As String) As Long
    On Error GoTo ErrHandler
    TheDcTest.Execute
    Execute = TL_SUCCESS
    Exit Function
ErrHandler:
    dumpPPMUreg
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Execute = TL_ERROR
End Function
