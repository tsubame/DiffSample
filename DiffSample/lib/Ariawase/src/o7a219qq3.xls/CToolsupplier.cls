VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CToolsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   このサプライヤーに属するツールカテゴリオブジェクトを保持するクラス
'
'目的:
'   データツリー構築のための最上位階層として存在する
'
'   Revision History:
'   Data        Description
'   2008/12/11  作成
'   2008/12/15  ■機能追加
'               パラメータの外部出力機能のインターフェイスを実装
'   2008/12/17  ■仕様変更
'               　データシートオブジェクトのプロパティ追加
'               　（サプライヤーオブジェクトがカテゴリに分類しないデータシートを直接管理する）
'   2009/02/06  ■機能削除
'               　Initializeメソッドを廃止・Nameプロパティ設定に変更
'   2009/02/20  ■機能追加
'               ①Addメソッドの追加（ツールカテゴリをコレクションに追加する）
'               ②IsContainプロパティの追加（ツールカテゴリを含んでいるかどうかを取得）
'               ■仕様変更
'               　カスタムツールとして別管理していたツールカテゴリは廃止
'                 ワイルドカードのカテゴリとして他のツールカテゴリと一元管理とする
'   2009/06/15  Ver1.01アドイン展開に伴う変更
'             　■仕様変更
'               実装インターフェースをJOB側の既存インターフェースと共有するためメソッド名を変更
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWritable

Private mDataTools As Collection
Private mName As String

Private Sub Class_Initialize()
    Set mDataTools = New Collection
End Sub

Public Sub Add(ByVal DataTool As CToolCategory)
'内容:
'   ツールカテゴリのコレクション追加
'
'パラメータ:
'[dataTool]   In  ツールカテゴリオブジェクト
'
'注意事項:
'
    On Error GoTo AssignErr
    mDataTools.Add DataTool, DataTool.Name
    Exit Sub
AssignErr:
    Err.Raise 999, "CToolSupplier.Add", "This Is Duplicate Name [ " & DataTool.Name & " ] In Data Table"
End Sub

Public Property Get DataTools() As Collection
'内容:
'   ツールカテゴリコレクションの取得
'
'戻り値：
'   CDataCategory型オブジェクトのコレクション
'
'注意事項:
'
    Set DataTools = mDataTools
End Property

Public Property Let Name(ByVal suppName As String)
'内容:
'   サプライヤー名の設定
'
'パラメータ:
'[suppName]   In  サプライヤー名
'
'注意事項:
'
    mName = suppName
End Property

Public Property Get Name() As String
'内容:
'   サプライヤー名の取得
'
'戻り値：
'   サプライヤー名
'
'注意事項:
'
    Name = mName
End Property

Public Property Get IsContain() As Boolean
'内容:
'   ツールカテゴリを保持しているかどうかのプロパティ取得
'
'戻り値：
'   True/False
'
'注意事項:
'
    IsContain = mDataTools.Count > 0
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    '### パラメータ出力の実行 #######################################
    With writer
        .WriteAsString SUPPLIER_NAME, mName
        .WriteAsBoolean IS_TOOL_CONTAIN, Me.IsContain
    End With
End Sub

Private Property Get IParameterWritable_ParameterList() As Collection
    Err.Raise 9999, "CToolsupplier.IParameterWritable.ParameterList", "Does Not Support This Method !"
End Property

Public Function AsIParameterWritable() As IParameterWritable
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IParameterWritable型オブジェクト
'
'注意事項:
'
    Set AsIParameterWritable = Me
End Function

Private Sub Class_Terminate()
    Set mDataTools = Nothing
End Sub

