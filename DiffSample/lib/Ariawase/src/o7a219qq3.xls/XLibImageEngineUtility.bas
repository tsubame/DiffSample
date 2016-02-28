Attribute VB_Name = "XLibImageEngineUtility"
'概要:
'   TheImageTestのユーティリティ
'
'目的:
'   TheImageTest:CImageEngineの初期化/破棄のユーテリティを定義する
'
'作成者:
'   a_oshima

Option Explicit

Public TheImageTest As CImageEngine

Public Sub CreateTheImageTestIfNothing()
'内容:
'   As Newの代替として初回に呼ばせるインスタンス生成処理
'
'パラメータ:
'   なし
'
'注意事項:
'
    On Error GoTo ErrHandler
    If TheImageTest Is Nothing Then
        '### TheImageTestの初期化 ###################
        Set TheImageTest = New CImageEngine
        With TheImageTest
            .Initialize GetActionLoggerInstance, GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeTestInstances)
            If .CreateScenario = TL_ERROR Then
                TheError.Raise 9999, "XLibImageEngineUtility.CreateTheImageTestIfNothing", "CreateScenario returned TL_ERROR"
            End If
        End With
    End If
    Exit Sub
ErrHandler:
    Set TheImageTest = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub DestroyTheImageTest()
    Set TheImageTest = Nothing
End Sub

Public Function RunAtJobEnd() As Long

End Function

Public Sub EnableInterceptor(pFlag As Boolean, pLogger As CActionLogger)
    Call TheImageTest.EnableInterceptor(pFlag, pLogger)
End Sub
