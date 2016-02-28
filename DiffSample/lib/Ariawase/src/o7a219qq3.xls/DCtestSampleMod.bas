Attribute VB_Name = "DCtestSampleMod"
'概要:
'   テストインスタンスを作成する際のサンプルプロシージャ群
'
'目的:
'   DCテストシナリオシートを利用する際の雛形インスタンスをユーザに公開する
'
'
'作成者:
'   SLSI大谷
'
Option Explicit

Private Function MultiDcTest_f() As Long
'内容:
'   DCテストインスタンスを作成する際のサンプルプロシージャ
'   ユーザーはこのサンプルから任意のDCテストインスタンスを作成・実装することが可能
'   またシナリオ実行前後に任意の処理の追加が可能
'
'備考:
'   インスタンスシートのインスタンス名と、テストシナリオシートの
'   カテゴリ名（スペースは許可する）は整合を取っておく必要がある

    Call SiteCheck

    '@@@ 測定シナリオ初期化 @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    TheDcTest.SetScenario GetInstanceName
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    '##########  TEST CONDITION SET UP ####################
'    Call XXXXX_Setup
'    Call SET_XXXXX_CONDITION
'    Call SetVoltage(XXXXX)
'    Call SetVRL(XXXXX)
'    Call PatSet(XXXXX)
'    TheHdw.Wait XXXXX * mS
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    '@@@ 測定シナリオ実行 @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    TheDcTest.Execute
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

End Function


Private Function ReturnResult_ShirotenMargin_f() As Long
'内容:
'   インスタンス名をキーとしてテスト結果コレクションから要素を取り出しTest関数に渡す
'   上の関数との違いは、テスト結果がない場合は、エラーではなく、0を返す。
'
'備考:
'   インスタンスシートのインスタンス名と、テストシナリオシートの
'   テストラベルは整合を取っておく必要がある
'
    Dim resultTest() As Double
    
    Call SiteCheck

    If TheResult.IsExist(GetInstansNameAsUCase) Then
        Call TheResult.GetResult(GetInstansNameAsUCase, resultTest)
    Else
        ReDim resultTest(GetSiteCount)
        Dim SiteIndex As Long
        For SiteIndex = 0 To GetSiteCount
            resultTest(SiteIndex) = 0
        Next SiteIndex
    End If
    Call test(resultTest)
    
End Function
