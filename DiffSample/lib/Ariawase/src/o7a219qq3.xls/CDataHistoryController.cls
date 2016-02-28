VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataHistoryController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   データシート閲覧履歴管理用クラス
'
'目的:
'   IDataNavigatorインターフェイスを実装しナビゲーションの実体として存在する
'
'   Revision History:
'   Data        Description
'   2008/11/10  作成
'   2008/11/11　○全コード実行・動作OK
'   2008/12/11　■仕様変更
'             　①前後ヒストリ管理の廃止・インデックス管理への変更
'             　②ヒストリ有効時・リセット時のイベント発生の追加
'             　③インターフェイス機能追加による変更
'   2009/02/20　■仕様変更
'               ①ヒストリ有効時・リセット時のイベントマクロ呼び出しは削除（CBookEventsAcceptorへ移動）
'               ②履歴インデックス前後及び指定のデータシートオブジェクトを返すメソッドのロジックを変更
'               　（データシートが削除されている場合を考慮して自動的に有効なデータシートまで前後にスキップする）
'               ③履歴インデックスの有効判定メソッドのロジックを変更
'               　（インデックス位置だけでなく、インデックス前後に有効なデータシートが含まれているかどうかも考慮する）
'               ④インターフェイス機能追加による変更
'               ⑤データシートのアクティブステータス設定を廃止
'作成者:
'   0145206097
'
Option Explicit

Implements IDataNavigator

Private mHistory As Collection
Private mHistIndex As Long
Private mMaxHistory As Long

Private Sub Class_Initialize()
    Set mHistory = New Collection
    mHistIndex = 0
End Sub

Public Sub Initialize(ByVal maxHistory As Long)
'内容:
'   ナビゲーションオブジェクトの初期化
'
'パラメータ:
'[maxHistory]   In  最大ヒストリ数
'
'注意事項:
'
    If maxHistory < 1 Then GoTo WrongData
    mMaxHistory = maxHistory
    Exit Sub
    '### 最大ヒストリ数が無効な場合はエラー処理 #####################
WrongData:
    Err.Raise 999, "CDataHistoryController.Initialize", "Please Input Max History Number That Are Larger Than 1 !"
End Sub

Private Sub IDataNavigator_Add(ByVal dataSht As CDataSheet)
    '### データシートオブジェクトのヒストリへの追加 #################
    FILO mHistory, dataSht
End Sub

Private Property Get IDataNavigator_History() As Collection
    '### ヒストリがない場合はNothingを返す ##########################
    If mHistory.Count = 0 Then
        Set IDataNavigator_History = Nothing
    '### ヒストリコレクションを返す #################################
    Else
        Set IDataNavigator_History = mHistory
    End If
End Property

Private Function IDataNavigator_DataSheet(ByVal hIndex As Long) As CDataSheet
    '### 初期設定 ###################################################
    Set IDataNavigator_DataSheet = Nothing
    '### インデックスが有効な場合 ###################################
    If hIndex >= 1 And hIndex <= mHistory.Count Then
        '### 指定したデータシートが削除されていなければ #############
        '### データシートオブジェクトを返す             #############
        If Not mHistory(hIndex).IsDeleted Then
            mHistIndex = hIndex
            Set IDataNavigator_DataSheet = mHistory(mHistIndex)
        End If
    End If
End Function

Private Property Get IDataNavigator_ForwardDataSheet() As CDataSheet
    '### 初期設定 ###################################################
    Set IDataNavigator_ForwardDataSheet = Nothing
    '### インデックスをインクリメントしデータオブジェクトを返す #####
    '### 指定したデータシートが削除されている場合はスキップし   #####
    '### 「進む」ステータスが無効になるまで繰り返す             #####
    Do While Me.AsIDataNavigator.IsEnableForward
        mHistIndex = mHistIndex + 1
        If Not mHistory(mHistIndex).IsDeleted Then
            Set IDataNavigator_ForwardDataSheet = mHistory(mHistIndex)
            Exit Do
        End If
    Loop
End Property

Private Property Get IDataNavigator_BackwardDataSheet() As CDataSheet
    '### 初期設定 ###################################################
    Set IDataNavigator_BackwardDataSheet = Nothing
    '### インデックスをデクリメントしデータオブジェクトを返す #######
    '### 指定したデータシートが削除されている場合はスキップし #######
    '### 「戻る」ステータスが無効になるまで繰り返す           #######
    Do While Me.AsIDataNavigator.IsEnableBackward
        mHistIndex = mHistIndex - 1
        If Not mHistory(mHistIndex).IsDeleted Then
            Set IDataNavigator_BackwardDataSheet = mHistory(mHistIndex)
            Exit Do
        End If
    Loop
End Property

Private Property Get IDataNavigator_IsEnableForward() As Boolean
    '### 初期設定 ###################################################
    IDataNavigator_IsEnableForward = False
    '### インデックスがヒストリ数を超えていない場合 ##################
    If mHistIndex < mHistory.Count Then
        '### 現在のインデックス以降のデータが全て ####################
        '### 削除されていたらFalse                ####################
        Dim dataIdx As Long
        For dataIdx = mHistIndex + 1 To mHistory.Count
            If Not mHistory(dataIdx).IsDeleted Then
                IDataNavigator_IsEnableForward = True
                Exit For
            End If
        Next dataIdx
    End If
End Property

Private Property Get IDataNavigator_IsEnableBackward() As Boolean
    '### 初期設定 ###################################################
    IDataNavigator_IsEnableBackward = False
    '### インデックスが1以下でない場合 ###############################
    If mHistIndex - 1 > 0 Then
        '### 現在のインデックス以前のデータが全て ####################
        '### 削除されていたらFalse                ####################
        Dim dataIdx As Long
        For dataIdx = mHistIndex - 1 To 1 Step -1
            If Not mHistory(dataIdx).IsDeleted Then
                IDataNavigator_IsEnableBackward = True
                Exit For
            End If
        Next dataIdx
    End If
End Property

Private Property Get IDataNavigator_HistoryIndex() As Long
    '### カレントインデックスを返す #################################
    IDataNavigator_HistoryIndex = mHistIndex
End Property

Private Sub IDataNavigator_Clear()
    '### ヒストリとインデックスのリセット ###########################
    Set mHistory = New Collection
    mHistIndex = 0
End Sub

Public Function AsIDataNavigator() As IDataNavigator
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IDataNavigator型オブジェクト
'
'注意事項:
'
    Set AsIDataNavigator = Me
End Function

Private Function IsMaxCount(ByVal History As Collection) As Boolean
'####################################################################
'   登録した最大ヒストリ数を超えているかどうかの判定
'####################################################################
    IsMaxCount = (History.Count >= mMaxHistory)
End Function

Private Sub FILO(ByVal History As Collection, ByVal dataSht As CDataSheet)
'####################################################################
'   データシートオブジェクトのヒストリへの追加（先入れ後出し）
'####################################################################
    With History
        '### ヒストリが最大であったら古いデータから削除 #############
        If IsMaxCount(History) Then
            .Remove 1
            mHistIndex = mHistIndex - 1
        End If
        '### 追加するデータが一つ前と同じであれば無視する ###########
        If .Count > 0 Then
            If .Item(.Count).Name = dataSht.Name Then
                mHistIndex = .Count
                Exit Sub
            End If
        End If
        '### データの追加とインデックスの更新 #######################
        mHistIndex = .Count + 1
        .Add dataSht
    End With
End Sub

Private Sub Class_Terminate()
    Set mHistory = Nothing
End Sub

