Attribute VB_Name = "XEeeAuto_Result"
'概要:
'  　結果かえし系関数
'
'目的:
'
'
'作成者:
'   2012/01/27 Ver0.1 D.Maruyama
'   2013/03/15 Ver0.2 H.Arikawa 不要関数削除

Option Explicit

Private Function ReturnResult_f() As Long
'内容:
'   インスタンス名をキーとしてテスト結果コレクションから要素を取り出しTest関数に渡す
'
'備考:
'   インスタンスシートのインスタンス名と、テストシナリオシートの
'   テストラベルは整合を取っておく必要がある
'

    Call SiteCheck

    On Error GoTo DATA_ERR
    Dim resultTest() As Double
    TheResult.GetResult GetInstansNameAsUCase, resultTest
    Call test(resultTest)

    Exit Function
DATA_ERR:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    ReDim resultTest(GetSiteCount)
    Dim SiteIndex As Long
    For SiteIndex = 0 To GetSiteCount
        resultTest(SiteIndex) = 0
    Next SiteIndex
    Call test(resultTest)
    Break
End Function


Private Function ReturnResultEx_f() As Long
'内容:
'   インスタンス名をキーとしてオフセットシートからデータを取得し
'   テスト結果にオーバーライトしてからTest関数に渡す
'
'備考:
'   インスタンスシートのインスタンス名と、テストシナリオシートの
'   テストラベルは整合を取っておく必要がある
'

    Call SiteCheck
    On Error GoTo ErrHandler
    If TheOffsetResult Is Nothing Then
        Err.Raise 9999, "ReturnResultEx_f", "Can Not Implement This Instance In Function [" & GetInstanceName & "] !"
    End If
    '@@@ 測定結果コンバート @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    TheOffsetResult.Calculate GetInstansNameAsUCase, TheResult
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    Dim resultTest() As Double
    TheResult.GetResult GetInstansNameAsUCase, resultTest
    Call test(resultTest)
    Exit Function
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    ReDim resultTest(GetSiteCount)
    Dim SiteIndex As Long
    For SiteIndex = 0 To GetSiteCount
        resultTest(SiteIndex) = 0
    Next SiteIndex
    Call test(resultTest)
End Function
