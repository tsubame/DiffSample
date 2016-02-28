VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   データエクスプローラーがユーザーに提供するべき機能を定義するインターフェイス
'
'目的:
'   エクスプローラーの実態がこのインターフェイスを実装し提供機能を保障する
'
'   Revision History:
'   Data        Description
'   2008/11/10  作成
'   2008/11/11　全コード実行・動作OK
'   2008/12/11　データツリーの分類から漏れたツールカテゴリの取得機能追加
'   2008/12/15　データツリーの外部出力機能追加
'   2008/12/17　12/11追加の機能を削除
'
'作成者:
'   0145206097
'
Option Explicit

Public Sub ConstructDataTree(ByVal dataList As Collection)
'内容:
'   データツリー構築の実行
'
'パラメータ:
'[dataList]   In  CDataSheet型オブジェクトのコレクション
'
'注意事項:
'   予め登録されたデータツリーに従いデータリストを分類する
'
End Sub

Public Property Get DataTree() As Collection
'内容:
'   データツリーの取得
'
'戻り値：
'   CToolSupplier型オブジェクトのコレクション
'
'注意事項:
'
End Property

Public Sub WriteDataTree(ByVal writer As IParameterWriter)
'内容:
'   ツリーデータの出力を実行
'
'パラメータ:
'[writer]   In  ツリーデータの出力先となるライター
'
'注意事項:
'
'
End Sub
