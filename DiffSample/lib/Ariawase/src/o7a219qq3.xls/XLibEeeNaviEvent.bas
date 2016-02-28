Attribute VB_Name = "XLibEeeNaviEvent"
'概要:
'   EeeNavigationが利用するイベントマクロ関数群
'
'目的:
'   ①ナビゲーションGUIに登録するマクロ関数 [Commander_***]
'   ②ナビゲーションヒストリから発生するマクロ関数
'   　→ワークブックオブジェクトからの呼び出しに変更[2009/02/20] [BookEventsAcceptor_***]
'   ③ナビゲーションヒストリに対するプロパティ操作を行うためのマクロ関数
'     →???削除
'   ④ショートカットキーへの登録用マクロ関数
'   　→追加 [2009/02/20] [ShortCut_***]
'
'   Revision History:
'   Data        Description
'   2008/12/11　作成
'   2009/02/20  ■機能追加
'               　ショートカットキーへの登録用マクロ関数を追加（進むボタン、戻るボタン）
'               ■仕様変更
'               　ナビゲーションヒストリから呼び出されるマクロをワークブックオブジェクトからの呼び出しに変更
'
'作成者:
'   0145206097
'
Option Explicit

Public Sub Commander_DataTreeMenu_Events()
'内容:
'   データツリーメニューをクリックしたときに発生するイベントマクロ
'   エクスプローラーの更新を行いデータツリーメニューを表示する
'
'注意事項:
'
    On Error GoTo MenuError
    '### ワークブックオブジェクトへデータツリーの更新を要求する #####
    mDataFolder.ExplorerDataSheet
    '### ナビゲーションGUIオブジェクトへデータツリーの表示を要求する
    mEeeNaviBar.DisplayDataTreeMenu
    Exit Sub
MenuError:
    MsgBox "Error Occured !! " & CStr(999) & " - " & "EeeNavi Tool Bar" & Chr(13) & Chr(13) & "Can Not Display Data Tree Menu!"
End Sub

Public Sub Commander_HistoryMenu_Events()
'内容:
'   ヒストリメニューをクリックしたときに発生するイベントマクロ
'   ヒストリ一覧をメニュー表示する
'
'注意事項:
'
    On Error GoTo MenuError
    '### ナビゲーションGUIオブジェクトへヒストリ一覧の表示を要求する
    mEeeNaviBar.DisplayHistoryMenu
    Exit Sub
MenuError:
    MsgBox "Error Occured !! " & CStr(999) & " - " & "EeeNavi Tool Bar" & Chr(13) & Chr(13) & "Can Not Display Data History Menu!"
End Sub

Public Sub Commander_DataTreeMenuButton_Events(ByVal SheetName As String)
'内容:
'   データツリーメニューのデータシートをクリックしたときに発生するイベントマクロ
'   データシートを表示しヒストリへの追加を行う
'
'パラメータ:
'[sheetName]   In  表示するワークシート名
'
'注意事項:
'
    '### ワークブックオブジェクトへデータシートの表示を要求する #####
    mDataFolder.ShowDataSheet SheetName
End Sub

Public Sub Commander_HistoryButton_Events(ByVal SheetName As String)
'内容:
'   ヒストリメニューの各ボタンクリックイベントから呼び出されるマクロ
'   ワークブックオブジェクトへ指定したデータシートの表示を要求する
'   データシートを表示しヒストリへの追加は行わない
'
'パラメータ:
'[sheetName]   In  表示するワークシート名
'
'注意事項:
'
    '### ワークブックオブジェクトへデータシートの表示を要求する #####
    mDataFolder.ShowDataSheetWithEventCancel SheetName
End Sub

Public Sub Commander_HistoryMenuButton_Events(ByVal hIndex As Long)
'内容:
'   ヒストリメニューのデータシートをクリックしたときに発生するイベントマクロ
'   メニュー内のインデック番号をナビゲーションGUIオブジェクトへ渡す目的でのみ使用する
'
'パラメータ:
'[hIndex]   In  クリックされたメニューのインデックス番号
'
'注意事項:
'
    '### ナビゲーションGUIオブジェクトへインデックス番号を渡す ######
    mEeeNaviBar.HistoryMenuButton_Click hIndex
End Sub

Public Sub BookEventsAcceptor_History_Events()
'内容:
'   ワークブックオブジェクトからヒストリメニューの
'   プロパティ操作のために呼び出されるイベントマクロ
'
'注意事項:
'
    '### ヒストリメニューのステータスを動的に設定する ###############
    mEeeNaviBar.SetHistoryButtonEnable
End Sub

Public Sub ShortCut_HistoryForeButton_Events()
'内容:
'   ショートカットに登録するナビゲーション「進む」ボタン操作のマクロ
'
'注意事項:
'
    '### 「進む」ボタンのクリック動作を実行する #####################
    On Error Resume Next
    mEeeNaviBar.HistoryForeButton_Click
    On Error GoTo 0
End Sub

Public Sub ShortCut_HistoryBackButton_Events()
'内容:
'   ショートカットに登録するナビゲーション「戻る」ボタン操作のマクロ
'
'注意事項:
'
    '### 「戻る」ボタンのクリック動作を実行する #####################
    On Error Resume Next
    mEeeNaviBar.HistoryBackButton_Click
    On Error GoTo 0
End Sub

