VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ナビゲーションがユーザーに提供するべき機能を定義するインターフェイス
'
'目的:
'   ナビゲーションの実態がこのインターフェイスを実装し提供機能を保障する
'
'   Revision History:
'   Data        Description
'   2008/11/10  作成
'   2008/11/11　全コード実行・動作OK
'   2008/12/11　インデックス指定でのデータシート取得機能追加
'   2009/02/20　■仕様変更
'               　メソッド・プロパティ名の変更
'               ■機能追加
'               　HistoryIndexプロパティ追加（ヒストリのカレントインデックス取得機能）
'
'作成者:
'   0145206097
'
Option Explicit

Public Sub Add(ByVal dataSht As CDataSheet)
'内容:
'   ナビゲーションヒストリへの追加
'
'パラメータ:
'[dataSht]   In  CDataSheet型オブジェクト
'
'注意事項:
'   ヒストリ数が最大ヒストリ数を超える場合は古いヒストリから自動的に削除される
'
End Sub

Public Property Get History() As Collection
'内容:
'   ナビゲーションヒストリの取得
'
'戻り値：
'   CDataSheet型オブジェクトのコレクション
'
'注意事項:
'   ヒストリ数が0の場合はNothingを返す
'
End Property

Public Function DataSheet(ByVal hIndex As Long) As CDataSheet
'内容:
'   インデックス指定によるヒストリ中のデータシート取得
'
'パラメータ:
'[hIndex]   In  データシートのインデックス番号
'
'戻り値：
'   CDataSheet型オブジェクト
'
'注意事項:
'   インデックスが最大ヒストリ数を超える、または1未満の場合はNothingを返す
'   指定したインデックスのデータシートが削除されている場合はNothingを返す
'
End Function

Public Property Get ForwardDataSheet() As CDataSheet
'内容:
'   ヒストリ中の現在のインデックスの一つ前のデータシート取得
'
'戻り値：
'   CDataSheet型オブジェクト
'
'注意事項:
'   インデックスが最大ヒストリ数を超えた場合はNothingを返す
'   カレントのインデックスのデータシートが削除されている場合、
'   ヒストリ前方の削除されていないデータシートまで
'   インデックスをインクリメントしてオブジェクトを返す
'
End Property

Public Property Get BackwardDataSheet() As CDataSheet
'内容:
'   ヒストリ中の現在のインデックスの一つ後ろのデータシート取得
'
'戻り値：
'   CDataSheet型オブジェクト
'
'注意事項:
'   インデックスが1以下の場合はNothingを返す
'   カレントのインデックスのデータシートが削除されている場合、
'   ヒストリ後方の削除されていないデータシートまで
'   インデックスをデクリメントしてオブジェクトを返す
'
End Property

Public Property Get IsEnableForward() As Boolean
'内容:
'   ヒストリ中のカレントインデックスより前方のデータシート取得が可能かどうかの判別
'
'戻り値：
'   前方のデータ取得が可能であればTrue、不可であればFalse
'
'注意事項:
'   インデックスの範囲が有効であっても、ヒストリ中のカレントより前方の
'   データシートが全て削除されている場合はFalseを返す
'
End Property

Public Property Get IsEnableBackward() As Boolean
'内容:
'   ヒストリ中のカレントインデックスより後方のデータシート取得が可能かどうかの判別
'
'戻り値：
'   後方のデータ取得が可能であればTrue、不可であればFalse
'
'注意事項:
'   インデックスの範囲が有効であっても、ヒストリ中のカレントより後方の
'   データシートが全て削除されている場合はFalseを返す
'
End Property

Public Property Get HistoryIndex() As Long
'内容:
'   ヒストリのカレントインデックス取得
'
'戻り値：
'   カレントインデックス値
'
'注意事項:
'
End Property

Public Sub Clear()
'内容:
'   ヒストリの初期化
'
'注意事項:
'
End Sub

