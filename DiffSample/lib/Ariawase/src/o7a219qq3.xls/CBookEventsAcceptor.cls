VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBookEventsAcceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ワークブックのラッパークラス
'   ワークブックの機能拡張とワークシート操作のイベントを受け取り内部処理を行う
'
'目的:
'   ①ワークシート新規作成時のデータツリー更新
'   ②ワークシート名前変更時のデータツリー更新
'   ③ワークシート削除時のデータツリー更新
'   ④ワークシート閲覧時のナビゲーションヒストリ追加
'
'   ↓仕様変更[2008/12/15以降]
'
'   [外部公開用]
'   ①データツリー更新機能
'   ②データシート表示機能
'   ③データツリーのツリービューワーへの出力機能
'   [ワークシート操作のイベント取得]
'   ①ワークシート削除時のナビゲーションヒストリのリセット
'   ②ワークシート閲覧時のナビゲーションヒストリ追加
'
'   Revision History:
'   Data        Description
'   2008/11/18  作成・全コード実行・動作OK
'   2008/12/10　■機能追加
'             　　データツリー更新・データシート閲覧の追加機能を外部へ公開
'             　■イベント取得追加
'             　　IG-XLのエクセルイベント横取り対応のためシート削除のメニューイベントを取得
'   2008/12/15　■機能追加
'             　　ツリービュー出力のための外部公開機能を追加
'   2009/02/20　■仕様変更
'               ①エクセルオブジェクトの取得
'               ②シート切り替え時のマクロ呼び出しを追加
'               ③データシートのアクティブステータス設定を廃止
'               ④ナビゲーションへのデータ追加にワークシートオブジェクト以外を含めるよう変更
'
'作成者:
'   0145206097
'
Option Explicit

Private oXL As Excel.Application

Private WithEvents mWorkbook As Excel.Workbook
Attribute mWorkbook.VB_VarHelpID = -1
Private WithEvents mDeleteMenu As Office.CommandBarButton
Attribute mDeleteMenu.VB_VarHelpID = -1
Private mNavigator As IDataNavigator
Private mExplorer As IDataExplorer
Private mPreSheet As CDataSheet

Private Const DELETE_MENU_CAPTION = "Edit"
Private Const DELETE_BUTTON_ID = 847

Private Sub Class_Initialize()
    Set mNavigator = New IDataNavigator
    Set mExplorer = New IDataExplorer
    Set mPreSheet = New CDataSheet
End Sub

Public Sub Initialize(ByVal excelApp As Excel.Application, ByVal navigator As IDataNavigator, ByVal explorer As IDataExplorer)
'内容:
'   ワークブックGUIオブジェクトの初期化
'
'パラメータ:
'[excelApp]   In  エクセルオブジェクト
'[navigator]  In  ナビゲーションオブジェクト
'[explorer]   In  エクスプローラーオブジェクト
'
'注意事項:
'
    Set oXL = excelApp
    Set mWorkbook = oXL.ActiveWorkbook
    Set mNavigator = navigator
    Set mExplorer = explorer
    '### アクティブシートをナビゲーションヒストリに追加　############
    mNavigator.Add getDataSheet(mWorkbook.ActiveSheet)
    On Error GoTo ErrGetMenu
    '### ウィンドウズメニューバーコントロール取得 ###################
    Set mDeleteMenu = oXL.CommandBars(DELETE_MENU_CAPTION).FindControl(id:=DELETE_BUTTON_ID)
    Exit Sub
ErrGetMenu:
    Err.Raise 999, "CBookEventsAcceptor.Initialize", "Can Not Get Edit Menu Button Object !"
End Sub

'### ナビゲーションGUIに公開するワークブックの拡張機能 ##########################
Public Sub ExplorerDataSheet()
'内容:
'   ワークブック内のワークシート情報を収集してエクスプローラーのデータツリーを更新
'
'パラメータ:
'
'注意事項:
'
    '### エクスプローラーにデータを渡してデータツリーを構築する #####
    mExplorer.ConstructDataTree collectDataSheet
End Sub

Public Sub ShowDataSheet(ByVal SheetName As String)
'内容:
'   ワークブック内のワークシートの表示メソッド
'
'パラメータ:
'[sheetName]   In  アクティブにするワークシート名
'
'注意事項:
'   ワークシートが既にアクティブな場合はこの操作は無視される
'   このメソッド実行後はシート操作によるマクロイベントが発生する
'
    '### ワークシートが既にアクティブな場合は無視 ###################
    If SheetName = mWorkbook.ActiveSheet.Name Then Exit Sub
    '### シートをアクティブに設定 ###################################
    On Error Resume Next
    mWorkbook.Sheets(SheetName).Activate
    On Error GoTo 0
End Sub

Public Sub ShowDataSheetWithEventCancel(ByVal SheetName As String)
'内容:
'   ブックイベントを伴わないワークブック内のワークシートの表示メソッド
'
'パラメータ:
'[sheetName]   In  アクティブにするワークシート名
'
'注意事項:
'   このメソッドではワークシート操作によるイベントはキャンセルする
'   またワークシートが既にアクティブな場合はこの操作は無視される
'   イベントによるワークブックの内部処理を伴わない単純な機能を持つ
'
    '### ワークシートが既にアクティブな場合は無視 ###################
    If SheetName = mWorkbook.ActiveSheet.Name Then Exit Sub
    '### マクロイベントを無効に設定 #################################
    oXL.EnableEvents = False
    '### シートをアクティブに設定 ###################################
    On Error Resume Next
    mWorkbook.Sheets(SheetName).Activate
    On Error GoTo 0
    '### マクロイベントを有効に設定 #################################
    oXL.EnableEvents = True
End Sub

Public Sub WriteTreeView(ByVal writer As IParameterWriter)
'内容:
'   ツリーデータを更新してデータ出力の実行を要求する
'
'パラメータ:
'[writer]   In  ツリーデータの出力先となるライター
'
'注意事項:
'
    '### エクスプローラーの更新 #####################################
    Me.ExplorerDataSheet
    oXL.ScreenUpdating = False
    '### エクスプローラーに対しデータ出力を実行させる ###############
    On Error GoTo ErrHandler
    mExplorer.WriteDataTree writer
    oXL.ScreenUpdating = True
    Exit Sub
ErrHandler:
    oXL.ScreenUpdating = True
    Err.Raise 999, "CBookEventsAcceptor.WriteTreeView", "Can Not Print Tree View"
End Sub

'################################################################################

'#### ワークブックGUIの操作により起動するイベント関数群 #########################
Private Sub mDeleteMenu_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'####################################################################
'ワークシートメニュー上のワークシート削除ボタンが押された時に発生するイベント関数
'（シートタブ上の右クリックメニューを含む)
'
'IG-XL環境でワークシート削除を行うと、削除メニューのコントロールボタンに
'IG-XLがマクロを登録しているため、エクセルイベントが横取りされてしまう
'そのため削除ボタンクリックのイベントを取得してIG-XLイベントを一旦キャンセルし、
'手動で登録してあるIG-XLマクロを実行する
'####################################################################

    '### アクティブなシートからデータシートオブジェクトを生成 #######
    Dim workSh As CDataSheet
    Set workSh = New CDataSheet
    workSh.Initialize mWorkbook.ActiveSheet

    '### "Test Procedures"シートを削除するとIG-XLマクロ内で #########
    '### なぜかエラーになる そのエラーに巻き込まれないため  #########
    '### EeeNaviの処理をスルーする 　　　 　　　　　　　　　#########
    If workSh.ParentName = "Test Procedures" Then GoTo ErrIGXLMacro

    '### IGXLのマクロイベントを実行 #################################
    CancelDefault = True
    On Error GoTo ErrIGXLMacro
    oXL.Run (mDeleteMenu.OnAction)

    '### シートが削除されていたらナビゲーションをリセット ###########
    If workSh.IsDeleted Then
        mNavigator.Clear
        mNavigator.Add getDataSheet(mWorkbook.ActiveSheet)
        BookEventsAcceptor_History_Events
    End If
    Exit Sub
    '### マクロが登録されていない場合は通常の削除動作へ #############
ErrIGXLMacro:
    CancelDefault = False
End Sub

Private Sub mWorkbook_SheetDeactivate(ByVal sh As Object)
'####################################################################
'ワークシートがディアクティブになったときに発生するイベント関数
'####################################################################
    '### テンポラリのデータシートオブジェクト生成 ###################
    Set mPreSheet = getDataSheet(sh)
End Sub

Private Sub mWorkbook_SheetActivate(ByVal sh As Object)
'####################################################################
'ワークシートがアクティブになったときに発生するイベント
'####################################################################
    '### シートが削除されていたらナビゲーションをリセット ###########
    If mPreSheet.IsDeleted Then mNavigator.Clear
    '### アクティブワークシートをナビゲーションヒストリに追加　######
    mNavigator.Add getDataSheet(sh)
    BookEventsAcceptor_History_Events
End Sub
'################################################################################

Private Function collectDataSheet() As Collection
'####################################################################
'ワークブック中のワークシートを収集
'####################################################################
    Set collectDataSheet = New Collection
    Dim workSh As Object
    Dim currData As CDataSheet
    '### ワークブック中のワークシートを検索     #####################
    '### 検索対象はワークシートオブジェクトのみ #####################
    For Each workSh In mWorkbook.Worksheets
        '### データシートオブジェクト生成 ###########################
        Set currData = New CDataSheet
        currData.Initialize workSh
        '### データリストへの追加 ###################################
        collectDataSheet.Add currData
    Next workSh
    '### ワークシートがない場合はNothingを返す ######################
    If collectDataSheet.Count = 0 Then
        Set collectDataSheet = Nothing
    End If
End Function

Private Function getDataSheet(ByVal workSh As Object) As CDataSheet
'####################################################################
'データシートオブジェクトのコンストラクタ
'####################################################################
    Set getDataSheet = New CDataSheet
    With workSh
        getDataSheet.Initialize workSh
    End With
End Function

Private Sub Class_Terminate()
    Set mNavigator = Nothing
    Set mExplorer = Nothing
    Set mWorkbook = Nothing
    On Error Resume Next
    Set mDeleteMenu = Nothing
    On Error GoTo 0
End Sub
