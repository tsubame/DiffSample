Attribute VB_Name = "XLibTheFlagBankUtility"
'概要:
'   TheFlagBankのUtilityモジュール
'
'   Revision History:
'       Data        Description
'       2010/10/07  FlagBankのUtility機能を実装した
'       2010/10/28  コメント文を追加＆変更した
'       2011/03/04　CFlagBank不具合修正に伴う変更(by 0145206097)
'                   ダンプモード状態の判断処理をクラスへ移動
'
'作成者:
'   0145184346
'

Option Explicit

'/** パブリックフラグバンクオブジェクト **/
Public TheFlagBank As CFlagBank
'/** ログファイル名 **/
Private mSaveFileName As String

Public Sub CreateTheFlagBankIfNothing()
'内容:
'   TheFlagBankの初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    On Error GoTo ErrHandler
    If TheFlagBank Is Nothing Then Set TheFlagBank = New CFlagBank
    Exit Sub
ErrHandler:
    Set TheFlagBank = Nothing
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

Public Sub SaveModeTheFlagBank(ByVal pDump As Boolean, Optional saveFileName As String)
'内容:
'   TheFlagBankのログ取得を行なう
'
'パラメータ:
'   [pDump]         In ログ取得モード指定
'   [SaveFileName]  In ログファイル名
'
'戻り値:
'
'注意事項:
'
    If TheFlagBank Is Nothing Then Exit Sub
    mSaveFileName = saveFileName
    TheFlagBank.Dump pDump
End Sub

Public Sub DestroyTheFlagBank()
'内容:
'   TheFlagBankを破棄する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Set TheFlagBank = Nothing
    mSaveFileName = ""
End Sub

Public Function RunAtJobEnd() As Long
'内容:
'   テスト実行終了時に、LogFileを保存する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    If TheFlagBank Is Nothing Then Exit Function
    TheFlagBank.Save mSaveFileName
End Function
