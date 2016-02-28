Attribute VB_Name = "XLibWkShtReaderUtility"
'概要:
'   ReaderManagerのユーティリティ
'
'目的:
'   ReaderManagerの初期化/破棄のユーテリティを定義する
'
'作成者:
'   a_oshima

Option Explicit

Private mReaderManager As CWorkSheetReaderManager

Public Sub CreateReaderManagerIfNothing()
'内容:
'   As Newの代替として初回に呼ばせるインスタンス生成処理
'
'パラメータ:
'   なし
'
'注意事項:
'
    On Error GoTo ErrHandler
    If mReaderManager Is Nothing Then
        Set mReaderManager = New CWorkSheetReaderManager
        Call mReaderManager.GetReaderInstance(eSheetType.shtTypeDeviceConfigurations)
#If ITS <> 0 Then
        Call mReaderManager.GetReaderInstance(eSheetType.shtTypeImgTestScenario)
#End If
    End If
    Exit Sub
ErrHandler:
    Set mReaderManager = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function GetWkShtReaderManagerInstance() As CWorkSheetReaderManager
'内容:
'   ReaderManagerのインスタンスを返す
'
'パラメータ:
'   なし
'
'戻り値:
'   ReaderManagerのインスタンス
'
'例外:
'   未初期化時に呼ばれるとVBA例外発生
'  （パフォーマンス改善のためAsNewの代替として用意してあり、Nothingチェックは行わない）
'
'注意事項:
'   初期化処理を先に呼び、インスタンスが生成されていること

    Set GetWkShtReaderManagerInstance = mReaderManager
End Function

Public Sub DestroyWkShtReaderManager()
    Set mReaderManager = Nothing
End Sub


