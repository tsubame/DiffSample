VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CToolCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   このカテゴリに属するデータシートオブジェクトを保持するクラス
'
'目的:
'   データツリー構築のための中間階層として存在する
'
'   Revision History:
'   Data        Description
'   2008/11/10  作成・全コード実行・動作OK
'   2008/12/15  ■機能追加
'               パラメータの外部出力機能のインターフェイスを実装
'   2009/02/06  ■機能追加
'               　Visibleプロパティの追加（データツリーに表示するかしないかを指定）
'               ■機能削除
'               　Initializeメソッドを廃止・Nameプロパティ設定に変更
'               ○全コード実行・動作OK
'   2009/02/20  ■機能追加
'               ①Categorizeプロパティの追加 (データツリーにカテゴライズして表示するかどうかを指定)
'               ②NameMapプロパティの追加（データシート上のカテゴリ名の定義場所を指定する）
'               ③Addメソッドの追加（データシートをコレクションに追加する）
'               ④Sortメソッドの追加（データリストのソート作業を行う・CDataTreeComposerクラスからの移管機能）
'               ⑤IsContainプロパティの追加（データシートを含んでいるかどうかを取得）
'               ⑥IsUniqueプロパティの追加（データシートがユニークかどうかを取得）
'   2009/06/15  Ver1.01アドイン展開に伴う変更
'             　■仕様変更
'               実装インターフェースをJOB側の既存インターフェースと共有するためメソッド名を変更
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWritable

Private mDataSheets As Collection
Private mName As String
Private mNameMap As String
Private mIsVisible As Boolean
Private mCategorize As Boolean

Private Sub Class_Initialize()
    Set mDataSheets = New Collection
End Sub

Public Sub Clear()
'内容:
'   データシートコレクションの初期化
'
'注意事項:
'
    Set mDataSheets = New Collection
End Sub

Public Sub Add(ByVal dataSht As CDataSheet)
'内容:
'   データシートのコレクション追加
'
'パラメータ:
'[dataSht]   In  データシートオブジェクト
'
'注意事項:
'
    On Error GoTo IsDuplicate
    mDataSheets.Add dataSht, dataSht.Name
    Exit Sub
IsDuplicate:
    Err.Raise 999, "CToolCategory.Add", "This Is Duplicate Name [ " & dataSht.Name & " ] In Tree View Definition"
End Sub

Public Sub Sort(ByVal dataList As Collection)
'内容:
'   データリストから自分のツールに属するデータシートを抜き取る
'
'パラメータ:
'[dataList]   In/Out  データシートオブジェクトのコレクション
'
'注意事項:
'   抜き取ったデータシートはコレクションから削除される
'   自分自身がワイルドカードの場合はこの作業は行われない
'
    '### ツールカテゴリがワイルドカードの場合はソートを行わない #####
    If mName = WILD_CARD Then Exit Sub
    '### ここからデータリストのソート作業 ###########################
    Dim shtIdx As Long
    Dim dataSht As CDataSheet
    For shtIdx = dataList.Count To 1 Step -1
        Set dataSht = dataList.Item(shtIdx)
        '### ツール名が一致するデータシートを抜き取る ###############
        If mName = dataSht.ParentName(mNameMap) Then
            Me.Add dataSht
            dataList.Remove (shtIdx)
        End If
    Next shtIdx
End Sub

Public Property Get DataSheets() As Collection
'内容:
'   データシートコレクションの取得
'
'戻り値：
'   CDataSheet型オブジェクトのコレクション
'
'注意事項:
'
    Set DataSheets = mDataSheets
End Property

Public Property Let Name(ByVal toolName As String)
'内容:
'   ツール名の設定
'
'パラメータ:
'[toolName]   In  ツール名
'
'注意事項:
'
    mName = toolName
End Property

Public Property Get Name() As String
'内容:
'   ツール名の取得
'
'戻り値：
'   ツール名
'
'注意事項:
'
    Name = mName
End Property

Public Property Let Visible(ByVal isVisible As Boolean)
'内容:
'   データツリー表示プロパティの設定
'
'パラメータ:
'[isVisible]   In  True/False
'
'注意事項:
'
    mIsVisible = isVisible
End Property

Public Property Get Visible() As Boolean
'内容:
'   データツリー表示プロパティの取得
'
'戻り値：
'   True/False
'
'注意事項:
'
    Visible = mIsVisible
End Property

Public Property Let Categorize(ByVal isCategorize As Boolean)
'内容:
'   データツリーのカテゴライズ表示プロパティの設定
'
'パラメータ:
'[isCategorize]   In  True/False
'
'注意事項:
'
    mCategorize = isCategorize
End Property

Public Property Get Categorize() As Boolean
'内容:
'   データツリーのカテゴライズ表示プロパティの取得
'
'戻り値：
'   True/False
'
'注意事項:
'
    Categorize = mCategorize
End Property

Public Property Let NameMap(ByVal location As String)
'内容:
'   カテゴリ名取得するセル位置プロパティの設定
'
'パラメータ:
'[location]   In  セル位置を表す文字列（A1,B1形式）
'
'注意事項:
'
    mNameMap = location
End Property

Public Property Get IsContain() As Boolean
'内容:
'   データシートを保持しているかどうかのプロパティ取得
'
'戻り値：
'   True/False
'
'注意事項:
'
    IsContain = (mDataSheets.Count > 0)
End Property

Public Property Get IsUnique() As Boolean
'内容:
'   データシートがツールカテゴリ名で一意に決まるかのプロパティ取得
'
'戻り値：
'   True/False
'
'注意事項:
'
    IsUnique = (mDataSheets.Count = 1)
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    '### パラメータ出力の実行 #######################################
    With writer
        .WriteAsString TOOL_NAME, mName
        .WriteAsString NAME_LOCATION, mNameMap
        .WriteAsBoolean IS_CATEGORIZE, mCategorize
        .WriteAsBoolean IS_VISIBLE_TOOL, mIsVisible
        .WriteAsBoolean IS_SHT_CONTAIN, Me.IsContain
        .WriteAsBoolean IS_SHT_UNIQUE, Me.IsUnique
    End With
End Sub

Private Property Get IParameterWritable_ParameterList() As Collection
    Err.Raise 9999, "CToolCategory.IParameterWritable.ParameterList", "Does Not Support This Method !"
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
    Set mDataSheets = Nothing
End Sub
