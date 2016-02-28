VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataTreeComposer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   データシートのリストからデータツリーを構成し保持するクラス
'
'目的:
'   IDataExplorerインターフェイスを実装しエクスプローラーオブジェクトの実態として存在する
'
'   Revision History:
'   Data        Description
'   2008/11/10  作成
'   2008/11/11　○全コード実行・動作OK
'   2008/12/11　■仕様変更
'             　①ツールカテゴリの上位階層を追加管理へ変更
'             　②データツリー分類から漏れたデータシートをカスタムツールとして別管理へ変更
'             　③インターフェイス機能追加による変更
'   2008/12/15　■機能追加
'               　インターフェイス機能追加によりデータツリーの外部出力機能を実装
'   2008/12/17  ■仕様変更
'               12/11②の機能を修正・分類から漏れるデータシートの管理機能はサプライヤークラスへ移動
'   2009/02/20  ■仕様変更
'               ①ツリー構造の定義データをリーダーから取得する仕様に変更
'               ②データリストのソート機能はCToolCategoryクラスへ移管
'               　ワイルドカードのツールカテゴリのみこのクラスがデータシートの追加作業を直接行う
'
'作成者:
'   0145206097
'
Option Explicit

Implements IDataExplorer

Private mDataTree As Collection
Private mWildCard As CToolCategory

Private Sub Class_Initialize()
    Set mDataTree = New Collection
    Set mWildCard = Nothing
End Sub

Public Sub Initialize(ByVal reader As IParameterReader)
'内容:
'   エクスプローラーオブジェクトの初期化
'
'パラメータ:
'[reader]   In  データツリー構造が定義してあるファイルのリーダー
'
'注意事項:
'
    '### リーダーがNothingの場合はエラー処理 ########################
    If reader Is Nothing Then GoTo ReaderErr
    '### ファイルストリームの準備 ###################################
    Dim fileStream As IFileStream
    Set fileStream = reader
    '### 名前重複チェック用コレクション #############################
    Dim errChecker As New Collection
    Dim errName As String
    '### リーダーからパラメータの読み込み ###########################
    Dim tempSupp As CToolsupplier
    Do While Not fileStream.IsEOR
        Set tempSupp = New CToolsupplier
        '### リーダーからトップレイヤー名の読み込み #################
        tempSupp.Name = reader.ReadAsString(SUPPLIER_NAME)
        '### レイヤー名が空欄の場合はエラー処理 #####################
        If tempSupp.Name = NOT_DEFINE Then GoTo BlankErr
        '### リーダーからサブレイヤーパラメータの読み込み ###########
        Dim tempTool As CToolCategory
        Do While (reader.ReadAsString(TOOL_NAME) <> END_OF_DATA)
            Set tempTool = New CToolCategory
            With tempTool
                '### リーダーからサブレイヤー名の読み込み ###########
                .Name = reader.ReadAsString(TOOL_NAME)
                '### レイヤー名が空欄の場合はエラー処理 #############
                If .Name = NOT_DEFINE Then GoTo BlankErr
                '### その他サブレイヤーのプロパティ設定 #############
                .NameMap = reader.ReadAsString(NAME_LOCATION)
                .Visible = reader.ReadAsBoolean(IS_VISIBLE_TOOL)
                .Categorize = reader.ReadAsBoolean(IS_CATEGORIZE)
            End With
            '### サブレイヤーの登録 #################################
            errName = tempTool.Name
            On Error GoTo AssignErr
            tempSupp.Add tempTool
            errChecker.Add tempTool, tempTool.Name
            On Error GoTo 0
            '### ワイルドカードのカテゴリは保持しておく #############
            If tempTool.Name = WILD_CARD Then
                Set mWildCard = tempTool
            End If
            fileStream.MoveNext
        Loop
        '### トップレイヤーの登録 ###################################
        errName = tempSupp.Name
        On Error GoTo AssignErr
        mDataTree.Add tempSupp, tempSupp.Name
        On Error GoTo 0
        fileStream.MoveNext
    Loop
    Exit Sub
    '### エラー処理 #################################################
ReaderErr:
    Err.Raise 999, "CDataTreeComposer.Initialize", "Reader Object Is Necessary For Initialization"
AssignErr:
    Err.Raise 999, "CDataTreeComposer.Initialize", "This Is Duplicate Name [ " & errName & " ] In Data Table"
BlankErr:
    Err.Raise 999, "CDataTreeComposer.Initialize", "Wrong Blank Is Include In Data Table"
End Sub

Private Sub IDataExplorer_ConstructDataTree(ByVal dataList As Collection)
    '### 与えられたデータシートがNothingの場合は無視する ############
    If dataList Is Nothing Then Exit Sub
    '### ツールカテゴリクラスにリストのソート作業を行わせる #########
    Dim toolSupp As CToolsupplier
    For Each toolSupp In mDataTree
        Dim DataTool As CToolCategory
        For Each DataTool In toolSupp.DataTools
            With DataTool
                .Clear
                .Sort dataList
            End With
        Next DataTool
    Next toolSupp
    '### ワイルドカードのカテゴリに残りのデータシートを入れる #######
    If Not mWildCard Is Nothing Then
        Dim shtIdx As Long
        Dim dataSht As CDataSheet
        For shtIdx = dataList.Count To 1 Step -1
            Set dataSht = dataList.Item(shtIdx)
            mWildCard.Add dataSht
        Next shtIdx
    End If
End Sub

Private Property Get IDataExplorer_DataTree() As Collection
    '### ツールサプライヤーオブジェクトのコレクションを返す #########
    Set IDataExplorer_DataTree = mDataTree
End Property

Private Sub IDataExplorer_WriteDataTree(ByVal writer As IParameterWriter)
    '### 各データ階層にデータ出力を要求する #########################
    '### このクラスの役割はこれらの実行命令とライターを制御すること #
    Dim fileStream As IFileStream
    Set fileStream = writer
    With writer
        Dim toolSupp As CToolsupplier
        For Each toolSupp In mDataTree
            '### ライターへトップレイヤーパラメータの書き出し #######
            toolSupp.AsIParameterWritable.WriteOut writer
            fileStream.MoveNext
            Dim DataTool As CToolCategory
            For Each DataTool In toolSupp.DataTools
                '### ライターへサブレイヤーパラメータの書き出し #####
                DataTool.AsIParameterWritable.WriteOut writer
                fileStream.MoveNext
                Dim dataSht As CDataSheet
                For Each dataSht In DataTool.DataSheets
                    '### ライターへデータシートパラメータの書き出し #
                    dataSht.AsIParameterWritable.WriteOut writer
                    fileStream.MoveNext
                Next dataSht
            Next DataTool
        Next toolSupp
        fileStream.IsEOR
    End With
End Sub

Public Function AsIDataExplorer() As IDataExplorer
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IDataExplorer型オブジェクト
'
'注意事項:
'
    Set AsIDataExplorer = Me
End Function

Private Sub Class_Terminate()
    Set mDataTree = Nothing
    Set mWildCard = Nothing
End Sub
