VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNavigationCommander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ナビゲーションGUIのラッパークラス
'   ナビゲーションメニューの機能拡張とメニュークリック時のイベントを受け取り内部処理を行う
'
'目的:
'   ①ナビゲーションメニューの構築とマクロ関数の登録
'   ②データツリーの表示及びデータヒストリの表示
'   ③メニューからのユーザー要求を受け付け外部へ要求を出す
'
'   Revision History:
'   Data        Description
'   2008/12/11  作成
'   2009/02/20　■仕様変更
'               ■ツリー表示のロジックを変更
'               ①データシートを持たないツールカテゴリはツリービューに非表示
'               ②データシートが一つしかなくツール名から一意に決まる場合はツールカテゴリレイヤーにツール名+シート名でボタン表示
'               ③ツールカテゴリが非表示設定の場合はツリービューに表示しない
'               ④ツールカテゴリのカテゴライズ設定がオフの場合はツールカテゴリレイヤーにシート名でボタン表示
'               ■ヒストリ表示のロジックを変更
'               ①データシートが削除されている場合の表示を追加
'               ②アクティブプロパティ取得とナビゲータのインデックス取得によりカレントデータシートを判別するロジックに変更
'               ③表示するデータシートにワークシート以外のシートオブジェクトを含める
'               ■エクセルオブジェクトの取得
'               ①内部でショートカットキーにマクロ登録を行う（進むボタン、戻るボタン）
'               ②①に伴い、マクロ登録用に進むボタン、戻るボタン動作用メソッドを外部公開
'   2009/05/11　■仕様変更
'               ①Createメソッドの追加：Initializeで行っていたカスタムバー作成処理を移管
'               ②Destroyメソッドの追加：Createメソッドに対するメニューバー削除機能として追加
'                                        ショートカットキーの削除はここでのみ行われる
'
'作成者:
'   0145206097
'
Option Explicit

Private oXL As Excel.Application

Private mNavigator As IDataNavigator
Private mExplorer As IDataExplorer

Private mNaviCommandBar As Office.CommandBar
Private mDataTreeMenu As Office.CommandBarPopup
Private mHistoryMenu As Office.CommandBarPopup
Private WithEvents mForeButton As Office.CommandBarButton
Attribute mForeButton.VB_VarHelpID = -1
Private WithEvents mBackButton As Office.CommandBarButton
Attribute mBackButton.VB_VarHelpID = -1

Private Const NAVI_BAR_NAME = "EeeNavi"

Private Const DATATREE_BAR_MENU_CAPTION = "EeeNavi(&X)"
Private Const DATATREE_BAR_MENU_POPUP = "Data Sheet Browser"
Private Const DATATREE_BAR_MENU_ACTION = "Commander_DataTreeMenu_Events"

Private Const DATATREE_BAR_MENU_LIST_ACTION = "OpenDataFolder"
Private Const DATATREE_BAR_MENU_LIST_ICON = 32
Private Const DATATREE_BAR_MENU_BUTTON_ACTION = "Commander_DataTreeMenuButton_Events"
Private Const DATATREE_BAR_MENU_BUTTON_ICON = 142

Private Const HISTORY_BAR_FORE_BUTTON_POPUP = "Forward"
Private Const HISTORY_BAR_FORE_BUTTON_ICON = 133
Private Const HISTORY_BAR_BACK_BUTTON_POPUP = "Back"
Private Const HISTORY_BAR_BACK_BUTTON_ICON = 132

Private Const HISTORY_BAR_MENU_CAPTION = " "
Private Const HISTORY_BAR_MENU_POPUP = "History Browser"
Private Const HISTORY_BAR_MENU_ACTION = "Commander_HistoryMenu_Events"

Private Const HISTORY_BAR_MENU_BUTTON_CURRENT_CAPTION = "Current Browsing"
Private Const HISTORY_BAR_MENU_BUTTON_ACTION = "Commander_HistoryMenuButton_Events"
Private Const HISTORY_BAR_MENU_BUTTON_CURRENT_ICON = 184
Private Const HISTORY_BAR_MENU_BUTTON_FORE_ICON = 156
Private Const HISTORY_BAR_MENU_BUTTON_BACK_ICON = 155
Private Const HISTORY_BAR_MENU_BUTTON_DELETE_ICON = 358

Private Const FORE_BUTTON_SHORTCUT_KEY = "%{RIGHT}"
Private Const BACK_BUTTON_SHORTCUT_KEY = "%{LEFT}"
Private Const HISTORY_BAR_FORE_BUTTON_MACRO = "ShortCut_HistoryForeButton_Events"
Private Const HISTORY_BAR_BACK_BUTTON_MACRO = "ShortCut_HistoryBackButton_Events"

Private Sub Class_Initialize()
    Set mNavigator = New IDataNavigator
    Set mExplorer = New IDataExplorer
End Sub

Public Sub Initialize(ByVal excelApp As Excel.Application, ByVal navigator As IDataNavigator, ByVal explorer As IDataExplorer)
'内容:
'   ナビゲーションGUIオブジェクトの初期化
'
'パラメータ:
'[excelApp]   In  エクセルオブジェクト
'[navigator]  In  ナビゲーションオブジェクト
'[explorer]   In  エクスプローラーオブジェクト
'
'注意事項:
'
    Set oXL = excelApp
    Set mNavigator = navigator
    Set mExplorer = explorer
End Sub

Public Sub Create()
'内容:
'   ナビゲーションGUIの作成
'
'注意事項:
'
    '### 既にメニューバーが存在する場合は削除する ###################
    Me.Destroy
    oXL.ScreenUpdating = False
    '### ナビゲーション用メニューバーを新規で作成 ###################
    Set mNaviCommandBar = oXL.CommandBars.Add(Name:=NAVI_BAR_NAME, Position:=msoBarTop, temporary:=True)
    mNaviCommandBar.Visible = True
    '### データツリー用、ヒストリ用各メニューを追加 #################
    With mNaviCommandBar.Controls
        Set mDataTreeMenu = .Add(Type:=msoControlPopup)
        Set mBackButton = .Add(Type:=msoControlButton)
        Set mForeButton = .Add(Type:=msoControlButton)
        Set mHistoryMenu = .Add(Type:=msoControlPopup)
    End With
    '### データツリーメニュの設定と実行マクロ追加 ###################
    With mDataTreeMenu
         .Caption = DATATREE_BAR_MENU_CAPTION
        .TooltipText = DATATREE_BAR_MENU_POPUP
        .OnAction = DATATREE_BAR_MENU_ACTION
    End With
    '### ヒストリボタン（進む）の設定 ###############################
    With mForeButton
        .FaceId = HISTORY_BAR_FORE_BUTTON_ICON
        .Style = msoButtonIcon
        .TooltipText = HISTORY_BAR_FORE_BUTTON_POPUP
        .enabled = False
    End With
    '### ヒストリボタン（戻る）の設定 ###############################
    With mBackButton
        .BeginGroup = True
        .FaceId = HISTORY_BAR_BACK_BUTTON_ICON
        .Style = msoButtonIcon
        .TooltipText = HISTORY_BAR_BACK_BUTTON_POPUP
        .enabled = False
    End With
    '### ヒストリメニューの設定と実行マクロ追加 #####################
    With mHistoryMenu
        .Caption = HISTORY_BAR_MENU_CAPTION
        .TooltipText = HISTORY_BAR_MENU_POPUP
        .OnAction = HISTORY_BAR_MENU_ACTION
    End With
    '### データツリー表示用のメニュー初期設定 #######################
    createDataTreeMenu
    '### ショートカットキーと実行マクロの登録 #######################
    With oXL
        .OnKey BACK_BUTTON_SHORTCUT_KEY, HISTORY_BAR_BACK_BUTTON_MACRO
        .OnKey FORE_BUTTON_SHORTCUT_KEY, HISTORY_BAR_FORE_BUTTON_MACRO
    End With
    oXL.ScreenUpdating = True
End Sub

Public Sub Destroy()
'内容:
'   ナビゲーションGUIの破棄
'
'注意事項:
'
    With oXL
        On Error Resume Next
        .CommandBars(NAVI_BAR_NAME).Delete
        .OnKey BACK_BUTTON_SHORTCUT_KEY
        .OnKey FORE_BUTTON_SHORTCUT_KEY
        On Error GoTo 0
    End With
End Sub

Public Sub DisplayDataTreeMenu()
'内容:
'   データツリーのメニュー表示
'
'注意事項:
'
    '### データツリーのデータシートボタンを最初に削除 ###############
    clearDataTreeMenu
    '### エクスプローラーからデータツリーの取得 #####################
    Dim treeData As Collection
    Set treeData = mExplorer.DataTree
    '### データツリーのトップメニューの設定 #########################
    Dim topLayer As CToolsupplier
    For Each topLayer In treeData
        Dim topMenu As Office.CommandBarPopup
        Set topMenu = mDataTreeMenu.Controls(topLayer.Name)
        '### データがない場合はトップメニューを表示しない ###########
        topMenu.Visible = topLayer.IsContain
        If topMenu.Visible Then
            '### データツリーのサブメニューの設定 ###################
            Dim subLayer As CToolCategory
            For Each subLayer In topLayer.DataTools
                Dim subMenu As Office.CommandBarPopup
                Set subMenu = topMenu.Controls(subLayer.Name)
                '### データがない場合、または表示設定が無効に #######
                '### なっている場合はサブメニューを表示しない #######
                With subLayer
                    subMenu.Visible = (.IsContain And .Visible)
                End With
                If subMenu.Visible Then
                    '### メニュー以下にボタンアイテムを作成 #########
                    Dim dataItem As CDataSheet
                    For Each dataItem In subLayer.DataSheets
                        Dim menuItem As Office.CommandBarButton
                        Set menuItem = subMenu.Controls.Add(Type:=msoControlButton, Before:=1)
                        '### ボタンの設定と実行マクロの追加 #########
                        With menuItem
                            .Caption = dataItem.Name
                            .FaceId = DATATREE_BAR_MENU_BUTTON_ICON
                            .OnAction = "'Commander_DataTreeMenuButton_Events " & Chr(34) & dataItem.Name & Chr(34) & "'"
                            If dataItem.IsActive Then
                                .state = msoButtonDown
                            Else
                                .state = msoButtonUp
                            End If
                        End With
                    Next dataItem
                    '### データシートがカテゴリ名でユニークに決まる場合、
                    '### またはカテゴライズを行わない場合は階層表示しない
                    With subLayer
                        If .IsUnique Or Not .Categorize Then
                            For Each menuItem In subMenu.Controls
                                If .IsUnique And .Categorize Then
                                    menuItem.Caption = .Name
                                    menuItem.ShortcutText = "[" & .DataSheets(1).Name & "]"
                                End If
                                menuItem.Move topMenu.CommandBar, subMenu.index
                            Next menuItem
                            subMenu.Visible = False
                        End If
                    End With
                End If
            Next subLayer
        End If
    Next topLayer
End Sub

Public Sub DisplayHistoryMenu()
'内容:
'   データヒストリのメニュー表示
'
'注意事項:
'
    '### データヒストリのデータシートボタンを最初に削除 #############
    clearHistoryMenu
    '### ナビゲーターからデータヒストリの取得 #######################
    Dim dataList As Collection
    Dim histIdx As Long
    With mNavigator
        Set dataList = .History
        histIdx = .HistoryIndex
    End With
    '### ヒストリがない場合はボタンを作成しない #####################
    If dataList Is Nothing Then Exit Sub
    '### ヒストリのデータシートボタンの作成 #########################
    Dim itemIdx As Long
    itemIdx = 1
    Dim dataItem As CDataSheet
    For Each dataItem In dataList
        Dim menuItem As Office.CommandBarButton
        Set menuItem = mHistoryMenu.Controls.Add(Type:=msoControlButton, Before:=1)
        '### データシートボタンの設定 ###############################
        With menuItem
            '### アクティブの場合はボタンを無効に設定 ###############
            If (dataItem.IsActive And histIdx = itemIdx) Then
                .Caption = HISTORY_BAR_MENU_BUTTON_CURRENT_CAPTION
                .FaceId = HISTORY_BAR_MENU_BUTTON_CURRENT_ICON
                .state = msoButtonDown
                .enabled = False
            '### それ以外は実行マクロの追加 #########################
            Else
                .Caption = dataItem.Name
                .OnAction = "'Commander_HistoryMenuButton_Events " & Chr(34) & (itemIdx) & Chr(34) & "'"
                .state = msoButtonUp
                If itemIdx > histIdx Then
                    .FaceId = HISTORY_BAR_MENU_BUTTON_FORE_ICON
                Else
                    .FaceId = HISTORY_BAR_MENU_BUTTON_BACK_ICON
                End If
                '### データシートが削除されている場合は無効 #########
                If dataItem.IsDeleted Then
                    .FaceId = HISTORY_BAR_MENU_BUTTON_DELETE_ICON
                    .enabled = False
                Else
                    .enabled = True
                End If
            End If
        End With
        itemIdx = itemIdx + 1
    Next dataItem
End Sub

Private Sub createDataTreeMenu()
'####################################################################
'ナビゲーションツリービューポップアップメニューの初期設定を行う
'####################################################################
    '### エクスプローラーからデータツリーの取得 #####################
    Dim treeData As Collection
    Set treeData = mExplorer.DataTree
    Dim topLayer As CToolsupplier
    '### データツリーのトップメニューの作成 #########################
    For Each topLayer In treeData
        Dim topMenu As Office.CommandBarPopup
        Set topMenu = mDataTreeMenu.Controls.Add(Type:=msoControlPopup)
        topMenu.Caption = topLayer.Name
        '### データツリーのサブメニューの作成 #######################
        Dim subLayer As CToolCategory
        For Each subLayer In topLayer.DataTools
            Dim subMenu As Office.CommandBarPopup
            Set subMenu = topMenu.Controls.Add(Type:=msoControlPopup)
            subMenu.Caption = subLayer.Name
        Next subLayer
    Next topLayer
End Sub

Private Sub clearDataTreeMenu()
'####################################################################
'ナビゲーションツリービューメニューボタンを削除する
'####################################################################
    '### データツリートップメニューの取得 ###########################
    Dim topMenu As Office.CommandBarPopup
    For Each topMenu In mDataTreeMenu.Controls
        With topMenu
            Dim menuCount As Long
            menuCount = .Controls.Count
            Dim subIdx As Long
            For subIdx = menuCount To 1 Step -1
                Dim itemType As Long
                '### サブメニューのタイプ取得 #######################
                itemType = .Controls(subIdx).Type
                If itemType = Office.msoControlPopup Then
                    '### さらにサブメニューがあれば取得 #############
                    Dim subMenu As Office.CommandBarPopup
                    Set subMenu = .Controls(subIdx)
                    Dim itemCount As Long
                    itemCount = subMenu.Controls.Count
                    '### サブメニュー以下のボタン削除 ###############
                    Dim itemIdx As Long
                    For itemIdx = itemCount To 1 Step -1
                        subMenu.Controls(itemIdx).Delete
                    Next itemIdx
                '### サブメニューがボタンだったら削除 ###############
                ElseIf itemType = Office.msoControlButton Then
                    .Controls(subIdx).Delete
                End If
            Next subIdx
        End With
    Next topMenu
End Sub

Private Sub clearHistoryMenu()
'####################################################################
'ナビゲーションヒストリメニューボタンを削除する
'####################################################################
    Dim menuItem As Office.CommandBarButton
    For Each menuItem In mHistoryMenu.Controls
        menuItem.Delete
    Next menuItem
End Sub

Private Sub mForeButton_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'####################################################################
'ナビゲーションメニューバーの進むボタンが押された時に発生するイベント
'####################################################################
    '### ナビゲータから一つ前のデータ取得 ###########################
    Dim dataItem As CDataSheet
    Set dataItem = mNavigator.ForwardDataSheet
    '### 取得に成功したらイベントマクロ呼び出し #####################
    If Not dataItem Is Nothing Then
        Commander_HistoryButton_Events dataItem.Name
    End If
    '### 自分自身のヒストリメニューステータスの設定 #################
    Me.SetHistoryButtonEnable
End Sub

Private Sub mBackButton_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'####################################################################
'ナビゲーションメニューバーの戻るボタンが押された時に発生するイベント
'####################################################################
    '### ナビゲータから一つ後のデータ取得 ###########################
    Dim dataItem As CDataSheet
    Set dataItem = mNavigator.BackwardDataSheet
    '### 取得に成功したらイベントマクロ呼び出し #####################
    If Not dataItem Is Nothing Then
        Commander_HistoryButton_Events dataItem.Name
    End If
    '### 自分自身のヒストリメニューステータスの設定 #################
    Me.SetHistoryButtonEnable
End Sub

Public Sub HistoryMenuButton_Click(ByVal hIndex As Long)
'内容:
'   コマンドラインからヒストリメニューのボタン操作を行うメソッド
'
'パラメータ:
'[hIndex]  In  コントロールボタンのインデックス番号
'
'注意事項:
'
    '### ナビゲータから指定インデックスのデータ取得 #################
    Dim dataItem As CDataSheet
    Set dataItem = mNavigator.DataSheet(hIndex)
    '### 取得に成功したらイベントマクロ呼び出し #####################
    If Not dataItem Is Nothing Then
        Commander_HistoryButton_Events dataItem.Name
    End If
    '### 自分自身のヒストリメニューステータスの設定 #################
    Me.SetHistoryButtonEnable
End Sub

Public Sub HistoryForeButton_Click()
'内容:
'   コマンドラインから「進む」ボタン操作を行うメソッド
'
'注意事項:
'
    mForeButton.Execute
End Sub

Public Sub HistoryBackButton_Click()
'内容:
'   コマンドラインから「戻る」ボタン操作を行うメソッド
'
'注意事項:
'
    mBackButton.Execute
End Sub

Public Sub SetHistoryButtonEnable()
'内容:
'   「進む」「戻る」コントロールボタンのイネーブルステータスを設定するメソッド
'
'注意事項:
'
    With mNavigator
        '### 「進む」ボタンのイネーブル設定 #########################
        mForeButton.enabled = .IsEnableForward
        '### 「戻る」ボタンのイネーブル設定 #########################
        mBackButton.enabled = .IsEnableBackward
    End With
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    oXL.CommandBars(NAVI_BAR_NAME).Delete
    On Error GoTo 0
    Set mNavigator = Nothing
    Set mExplorer = Nothing
End Sub
