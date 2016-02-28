Attribute VB_Name = "XLibEeeNaviConstructor"
'概要:
'   Eee-JOBワークシートナビゲーションクラスのコンストラクタモジュール
'
'目的:
'   ブック起動時にナビゲーションのスタートアップカスタムメニューを生成する
'   ①エクスプローラーのデータツリーを構成するための各ツール名やその配給元の初期化
'   　→エクスプローラーのためのリーダーオブジェクトを生成する[2009/02/20変更]
'   ②ナビゲーションを実行するためのオブジェクトを生成し初期化を行う
'   ③ナビゲーションを開始するためのユーザーインターフェースをワークシートメニューバーへ登録する
'   ④ツリービュー出力のためのライターオブジェクトを生成
'
'   Revision History:
'   Data        Description
'   2008/12/11　作成
'   2008/12/15　■機能追加
'               　ツリービュー出力機能及びメニューへの追加登録
'   2009/01/16　■不具合修正
'               　IG-XLのイニシャライズによりエクセル終了時のアプリケーション操作が不能になる不具合回避
'               　→イニシャライズ時のセットアップメニュー再構築の際にナビゲーションオブジェクトを破棄する
'   2009/02/20　■仕様変更
'               ①ツリー構成の定義をワークシートへ移動・リーダーの初期化のみを行う
'               ②スタートアップカスタムメニューをチャートメニューバーにも表示させる
'   2009/04/07　Ver1.00リリース [EeeNavigationVer1_0.xla]
'   2009/04/21　■仕様変更
'               ①アドイン開発ガイドラインに従いファイル名変更 [EeeNavigationAddIn.xla]
'               ②アドイン開発ガイドラインに従いプロジェクト名変更 [EeeNavigationAddIn]
'               ③バージョン情報をファイルのカスタムプロパティに設定・バージョン番号をこのプロパティから取得
'               ④プロジェクト名変更によるコード修正
'   2009/04/22　■不具合修正
'               　チャートメニューバーへコピーしたツールバーがテンポラリコントロールにならない不具合解消
'               　→コピーではなくテンポラリで新規作成し、サブメニューのみをコピーする方法に変更
'   2009/05/11　■不具合修正
'               　ナビゲーション起動時のJOBアンロードでエクセルエラーが発生する不具合回避
'               　→CNavigationCommanderクラスの仕様変更とそれに伴う呼び出し側の変更
'               ■仕様変更
'               　セットアップメニューのナビゲーション終了ボタンのステータス設定を追加
'   2009/05/12  Ver1.01リリース
'             　■仕様変更
'               ①セットアップメニューのツリービュー出力ボタンのステータス設定を変更
'                 →アクティブなシートがグラフチャートの場合は無効に設定（IG-XLエラー回避のため）
'               ②ツリービュー出力ボタンのキャプションとアイコンの変更（プリンター出力を連想させるため）
'   2009/06/15  Ver1.01アドイン展開に伴う変更
'             　■仕様変更
'               ①ツリービュー定義シートをJOB側に展開し、シート名で読込先ワークシートを特定する
'               ②ツリービュー出力用のライターをテキスト用に変更
'               ③バージョンインフォメーションをプロパティ取得から固定に変更
'
'作成者:
'   0145206097
'
Option Explicit

Private Const WORKSHEET_MENU_ID = "Worksheet Menu Bar"
Private Const CHART_MENU_ID = "Chart Menu Bar"
Private Const EEENAVI_SETUP_MENU_CAPTION = "EeeNavi SetUp(&N)"
Private Const EEENAVI_SETUP_MENU_ACTION = "SetMenuButtonStatus"
Private Const EEENAVI_SETUP_MENU_STARTUP_BUTTON_CAPTION = "Start Navigation(&S)"
Private Const EEENAVI_SETUP_MENU_STARTUP_BUTTON_POPUP = "Start EeeNavi Tool Bar"
Private Const EEENAVI_SETUP_MENU_STARTUP_BUTTON_ICON = 140
Private Const EEENAVI_SETUP_MENU_STARTUP_BUTTON_ACTION = "ConstructEeeNavigation"
Private Const EEENAVI_SETUP_MENU_PRINT_BUTTON_CAPTION = "Make TreeView(&M)"
Private Const EEENAVI_SETUP_MENU_PRINT_BUTTON_POPUP = "Make Tree View On Worksheet"
Private Const EEENAVI_SETUP_MENU_PRINT_BUTTON_ICON = 512
Private Const EEENAVI_SETUP_MENU_PRINT_BUTTON_ACTION = "CreateTreeView"
Private Const EEENAVI_SETUP_MENU_CLOSE_BUTTON_CAPTION = "End Navigation(&E)"
Private Const EEENAVI_SETUP_MENU_CLOSE_BUTTON_POPUP = "Terminate EeeNavi Tool Bar"
Private Const EEENAVI_SETUP_MENU_CLOSE_BUTTON_ICON = 358
Private Const EEENAVI_SETUP_MENU_CLOSE_BUTTON_ACTION = "TerminateEeeNavigation"
Private Const EEENAVI_SETUP_MENU_INFO_BUTTON_CAPTION = "Infomation(&I)"
Private Const EEENAVI_SETUP_MENU_INFO_BUTTON_POPUP = "Show EeeNavi Infomation"
Private Const EEENAVI_SETUP_MENU_INFO_BUTTON_ICON = 984
Private Const EEENAVI_SETUP_MENU_INFO_BUTTON_ACTION = "LoadVersionInfomation"

Private Const PROPERTY_NAME = "File Version"
Private Const APPLICATION_NAME = "EeeNavigation"

Public mEeeNaviBar As CNavigationCommander
Public mDataFolder As CBookEventsAcceptor

Private Const MAX_HISTORY = 15

Public Sub CreateEeeNaviSetUpMenu()
'内容:
'   ナビゲーションを開始するためのカスタムメニューバーをエクセルメニューバーへ登録する
'   ブックを開いたときに実行する必要がある
'
'注意事項:
'   IG-XL環境下でシステム初期化を行うとワークブックオブジェクトなどの参照が切れるため、
'   スタートアップは自動で行わずこのメニューバーからユーザーが手動で起動する
'   メニューバーはテンポラリに設定し、ブックを閉じると自動的に削除されるようにする
'
    '### ナビゲーションオブジェクトの破棄 ###########################
    TerminateEeeNavigation
    '### ワークシートメニューバーオブジェクトの取得 #################
    Dim wkShtMenuBar As Office.CommandBar
    Set wkShtMenuBar = Application.CommandBars(WORKSHEET_MENU_ID)
    '### チャートメニューバーオブジェクトの取得 #####################
    Dim chartMenuBar As Office.CommandBar
    Set chartMenuBar = Application.CommandBars(CHART_MENU_ID)
    '### 既にメニューバーが存在する場合は削除する ###################
    On Error Resume Next
    wkShtMenuBar.Controls(EEENAVI_SETUP_MENU_CAPTION).Delete
    chartMenuBar.Controls(EEENAVI_SETUP_MENU_CAPTION).Delete
    On Error GoTo 0
    '### ワークシートメニューバーへEeeNaviツールバーの追加 ##########
    Dim eeeNaviMenu As Office.CommandBarPopup
    Set eeeNaviMenu = wkShtMenuBar.Controls.Add(Type:=msoControlPopup, temporary:=True)
    '### チャートメニューバーへEeeNaviツールバーの追加 ##############
    Dim eeeNaviCMenu As Office.CommandBarPopup
    Set eeeNaviCMenu = chartMenuBar.Controls.Add(Type:=msoControlPopup, temporary:=True)
    '### 追加したツールバーのキャプション設定とサブメニュー追加 #####
    Dim startNaviBtn As Office.CommandBarButton
    Dim printTreeBtn As Office.CommandBarButton
    Dim endNaviBtn As Office.CommandBarButton
    Dim helpMenu As Office.CommandBarButton
    With eeeNaviMenu
        .Caption = EEENAVI_SETUP_MENU_CAPTION
        .OnAction = EEENAVI_SETUP_MENU_ACTION
        With .Controls
            Set startNaviBtn = .Add(Type:=msoControlButton)
            Set printTreeBtn = .Add(Type:=msoControlButton)
            Set endNaviBtn = .Add(Type:=msoControlButton)
            Set helpMenu = .Add(Type:=msoControlButton)
        End With
    End With
    With eeeNaviCMenu
        .Caption = eeeNaviMenu.Caption
        .OnAction = eeeNaviMenu.OnAction
    End With
    '### EeeNaviスタートアップボタンの設定と実行マクロ追加 ##########
    With startNaviBtn
        .Caption = EEENAVI_SETUP_MENU_STARTUP_BUTTON_CAPTION
        .TooltipText = EEENAVI_SETUP_MENU_STARTUP_BUTTON_POPUP
        .FaceId = EEENAVI_SETUP_MENU_STARTUP_BUTTON_ICON
        .OnAction = EEENAVI_SETUP_MENU_STARTUP_BUTTON_ACTION
        '### チャートメニューバーへコピー ###########################
        .Copy eeeNaviCMenu.CommandBar
    End With
    '### EeeNaviプリントボタンの設定と実行マクロ追加 ################
    With printTreeBtn
        .Caption = EEENAVI_SETUP_MENU_PRINT_BUTTON_CAPTION
        .TooltipText = EEENAVI_SETUP_MENU_PRINT_BUTTON_POPUP
        .FaceId = EEENAVI_SETUP_MENU_PRINT_BUTTON_ICON
        .OnAction = EEENAVI_SETUP_MENU_PRINT_BUTTON_ACTION
        '### チャートメニューバーへコピー ###########################
        .Copy eeeNaviCMenu.CommandBar
    End With
    '### EeeNavi終了ボタンの設定と実行マクロ追加 ####################
    With endNaviBtn
        .Caption = EEENAVI_SETUP_MENU_CLOSE_BUTTON_CAPTION
        .TooltipText = EEENAVI_SETUP_MENU_CLOSE_BUTTON_POPUP
        .FaceId = EEENAVI_SETUP_MENU_CLOSE_BUTTON_ICON
        .OnAction = EEENAVI_SETUP_MENU_CLOSE_BUTTON_ACTION
        .BeginGroup = True
        '### チャートメニューバーへコピー ###########################
        .Copy eeeNaviCMenu.CommandBar
    End With
    '### EeeNaviヘルプボタンの設定と実行マクロ追加 ##################
    With helpMenu
        .Caption = EEENAVI_SETUP_MENU_INFO_BUTTON_CAPTION
        .TooltipText = EEENAVI_SETUP_MENU_INFO_BUTTON_POPUP
        .FaceId = EEENAVI_SETUP_MENU_INFO_BUTTON_ICON
        .OnAction = EEENAVI_SETUP_MENU_INFO_BUTTON_ACTION
        .BeginGroup = True
        '### チャートメニューバーへコピー ###########################
        .Copy eeeNaviCMenu.CommandBar
    End With
End Sub

Public Sub SetMenuButtonStatus()
'内容:
'   EeeNaviメニュークリックで実行されるマクロ関数
'   プリントボタンと終了ボタンのステータスを設定する
'
'注意事項:
'   ワークブックオブジェクトがNothingの場合かつアクティブなシートが
'   ワークシート以外である場合はプリントボタンは無効になる
'   （IG-XLエラー回避のため）
'
    '### EeeNaviメニューバーオブジェクトの取得 ######################
    Dim eeeNaviMenu As Office.CommandBarPopup
    Set eeeNaviMenu = Application.CommandBars.ActionControl
    '### EeeNaviプリントボタンオブジェクトの取得 ####################
    Dim printTreeBtn As Office.CommandBarButton
    Set printTreeBtn = eeeNaviMenu.Controls(EEENAVI_SETUP_MENU_PRINT_BUTTON_CAPTION)
    '### EeeNaviプリントボタンイネーブルの設定 ######################
    Dim shType As Excel.XlSheetType
    shType = Application.ActiveWorkbook.ActiveSheet.Type
    printTreeBtn.enabled = ((Not mDataFolder Is Nothing) And (shType = Excel.xlWorksheet))
    '### EeeNavi終了ボタンオブジェクトの取得 ########################
    Dim endNaviBtn As Office.CommandBarButton
    Set endNaviBtn = eeeNaviMenu.Controls(EEENAVI_SETUP_MENU_CLOSE_BUTTON_CAPTION)
    '### EeeNavi終了ボタンイネーブルの設定 ##########################
    endNaviBtn.enabled = (Not mEeeNaviBar Is Nothing)
End Sub

Public Sub ConstructEeeNavigation()
'内容:
'   EeeNaviスタートアップボタンで実行されるマクロ関数
'   ナビゲーションのオブジェクトを生成し初期化を行う
'
'注意事項:
'
    On Error GoTo ErrHandler
    '### ツリービュー定義ワークシートの取得 #########################
    Dim wsSheet As Excel.Worksheet
    Set wsSheet = getWsSheet("TreeViewDefinition")
    '### ナビゲーションオブジェクトの生成 ###########################
    Dim eeeNaviCore As CDataHistoryController
    Set eeeNaviCore = New CDataHistoryController
    eeeNaviCore.Initialize MAX_HISTORY
    '### ツリー定義データのリーダー準備 #############################
    Dim treeReader As CWsTreeDataReader
    Set treeReader = New CWsTreeDataReader
    treeReader.Initialize wsSheet
    '### エクスプローラーオブジェクトの生成 #########################
    Dim eeeExplCore As CDataTreeComposer
    Set eeeExplCore = New CDataTreeComposer
    eeeExplCore.Initialize treeReader
    '### ワークブックオブジェクトの生成 #############################
    Set mDataFolder = New CBookEventsAcceptor
    mDataFolder.Initialize Application, eeeNaviCore, eeeExplCore
    '### ナビゲーションGUIオブジェクトの生成 ########################
    Set mEeeNaviBar = New CNavigationCommander
    With mEeeNaviBar
        .Initialize Application, eeeNaviCore, eeeExplCore
        .Create
    End With
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Private Function getWsSheet(ByVal shName As String) As Excel.Worksheet
    '### ワークシートオブジェクト取得用プロシージャ #################
    On Error GoTo NotExist
    Set getWsSheet = ActiveWorkbook.Worksheets(shName)
    Exit Function
NotExist:
    Err.Raise 9999, "Start EeeNavigation", shName & " Worksheet Can Not Find !"
End Function

Public Sub CreateTreeView()
'内容:
'   EeeNaviプリントボタンで実行されるマクロ関数
'   データツリーの一覧をテキストにプリントアウトする
'
'注意事項:
'
    '### ツリービューライターオブジェクトの生成 #####################
    Dim treeViewWriter As CTextTreeViewWriter
    Set treeViewWriter = New CTextTreeViewWriter
    On Error GoTo ErrHandler
    treeViewWriter.OpenFile ActiveWorkbook.Path
    '### ワークブックオブジェクトに対しツリービュー作成を実行 #######
    mDataFolder.WriteTreeView treeViewWriter
    treeViewWriter.CloseFile
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Exit Sub
End Sub

Public Sub TerminateEeeNavigation()
'内容:
'   EeeNavi終了ボタンで実行されるマクロ関数
'   ナビゲーションのデコンストラクタを行う
'
'注意事項:
'
    '### ワークブックオブジェクトの破棄 #############################
    Set mDataFolder = Nothing
    '### ナビゲーションGUI用メニューバーの掃除 ######################
    'オブジェクトが不定状態の場合にエラーにならないように回避（不要？）
    On Error Resume Next
    If Not mEeeNaviBar Is Nothing Then mEeeNaviBar.Destroy
    On Error GoTo 0
    '### ナビゲーションオブジェクトの破棄 ###########################
    Set mEeeNaviBar = Nothing
End Sub

Public Sub LoadVersionInfomation()
'内容:
'   EeeNaviインフォメーションボタンで実行されるマクロ関数
'   ナビゲーション情報をフォーム表示する
'
'注意事項:
'
    Dim revNum As String
'    revNum = ThisWorkbook.CustomDocumentProperties.Item(PROPERTY_NAME).Value
    revNum = "1.01"
    With EeeNaviVerFrm
        .VersionLabel = APPLICATION_NAME & " Ver." & revNum
        .Show
    End With
End Sub
