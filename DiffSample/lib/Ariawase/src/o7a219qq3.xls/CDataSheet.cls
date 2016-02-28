VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   対象となるデータシートのオブジェクトを保持するクラス
'
'目的:
'   ワークシートナビゲーションとユーザーインターフェースは
'   このオブジェクトを利用して情報を交換する
'
'   Revision History:
'   Data        Description
'   2008/11/10  作成・全コード実行・動作OK
'   2008/12/11　■仕様変更
'             　　シート名プロパティをワークシートオブジェクトプロパティへ変更（シート名変更に対応するため）
'   2008/12/15  ■機能追加
'               　パラメータの外部出力機能のインターフェイスを実装
'   2009/02/06  ■仕様変更
'               ①ツール名のプロパティをRead Onlyに変更
'               ②ワークシートが破棄されていた場合のシート名とツール名に削除されている事を示すラベルを返すよう変更
'               ○全コード実行・動作OK
'   2009/02/20　■仕様変更
'               ①ツール名取得セル位置をオプション指定出来るよう修正
'               ②アクティブステータスはこのクラスが保持しているワークシートオブジェクトから取得するよう変更
'               ③このクラスが保持するデータの実体をワークシートからシートオブジェクトに変更
'   2009/06/15  Ver1.01アドイン展開に伴う変更
'             　■仕様変更
'               実装インターフェースをJOB側の既存インターフェースと共有するためメソッド名を変更
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWritable

Private mDataSheet As Object
Private mIsActive As Boolean

Public Sub Initialize(ByVal excelApp As Object)
'内容:
'   データシートオブジェクトの初期化
'
'パラメータ:
'[excelApp]   In  データシートの実体となるシートオブジェクト
'
'注意事項:
'
    Set mDataSheet = excelApp
End Sub

Public Function IsDeleted() As Boolean
'内容:
'   シートオブジェクトの有無を判別
'
'戻り値：
'   シートオブジェクトが破棄されていたらTrueを返す
'
'注意事項:
'   ユーザーのシート操作によりこのクラスが保持しているシートオブジェクトが
'   破棄された場合、このプロパティを利用してシートの存在の有無をチェック出来る
'
    On Error GoTo NotFound
    Dim MyName As String
    MyName = mDataSheet.Name
    IsDeleted = False
    Exit Function
NotFound:
    IsDeleted = True
End Function

Public Property Get Name() As String
'内容:
'   シート名の取得
'
'戻り値：
'   シート名
'
'注意事項:
'   シートが破棄されている場合は"<Deleted>"を返す
'
    If Not Me.IsDeleted Then
        Name = mDataSheet.Name
    Else
        Name = SHEET_MISSING
    End If
End Property

Public Function ParentName(Optional ByVal location As String = DEF_NAME_MAP) As String
'内容:
'   ワークシートが属するツール名の取得
'
'パラメータ:
'[location]   In  ツール名が定義されているワークシートのセル位置
'
'戻り値：
'   ツール名
'
'注意事項:
'   このプロパティが使えるのは対象がワークシートオブジェクトであるときだけ
'   それ以外のシートオブジェクトの場合は"NotWorkSheet"を返す
'   ワークシートのツール名セル位置はデフォルト"B1"
'   ワークシートが破棄されている場合は"<Deleted>"を返す
'
    If Not Me.IsDeleted Then
        If mDataSheet.Type = Excel.xlWorksheet Then
            ParentName = mDataSheet.Range(location).Value
        Else
            ParentName = NOT_WORKSHEET
        End If
    Else
        ParentName = SHEET_MISSING
    End If
End Function

Public Property Get IsActive() As Boolean
'内容:
'   シートのアクティブステータスの取得
'
'戻り値：
'   True/False
'
'注意事項:
'   シートが破棄されている場合はFalseを返す
'
    If Not Me.IsDeleted Then
        With mDataSheet
            IsActive = (.Name = .parent.ActiveSheet.Name)
        End With
    Else
        IsActive = False
    End If
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    '### パラメータ出力の実行 #######################################
    With writer
        .WriteAsString sheet_name, Me.Name
        .WriteAsBoolean IS_SHT_ACTIVE, Me.IsActive
        .WriteAsBoolean IS_SHT_DELETED, Me.IsDeleted
    End With
End Sub

Private Property Get IParameterWritable_ParameterList() As Collection
    Err.Raise 9999, "CDataSheet.IParameterWritable.ParameterList", "Does Not Support This Method !"
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
    Set mDataSheet = Nothing
End Sub
