VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWsTreeDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ツリービューの階層構成定義ワークシートのリーダークラス
'
'目的:
'   ツリービューの階層構成定義をワークシートから読み込む
'
'   Revision History:
'   Data        Description
'   2009/02/03  作成
'   2009/06/15  Ver1.01アドイン展開に伴う変更
'             　■仕様変更
'               実装インターフェースをJOB側の既存インターフェースと共有するためメソッド名を変更
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterReader
Implements IFileStream

Private mTreeDataSh As Excel.Worksheet
Private mDataCell As Excel.Range
Private Const START_CELL = "A4"
Private Const SUBLAYER_COL = 1
Private Const LOCATION_COL = 2
Private Const CATEGORIZE_COL = 3
Private Const VISIBLE_COL = 4

Public Sub Initialize(ByVal wksht As Excel.Worksheet)
'内容:
'   リーダーオブジェクトの初期化
'
'パラメータ:
'[wkSht]  In  読み込み先のワークシートオブジェクト
'
'注意事項:
'
    '### ワークシートオブジェクトの取得 #############################
    Set mTreeDataSh = wksht
    '### ファイルストリームのリセット ###############################
    Me.AsIFileStream.Rewind
    '### 先頭のデータが定義されていない場合はエラー処理 #############
    With mTreeDataSh
        If IsEmpty(mDataCell) Then
            Err.Raise 999, "CWsTreeDataReader.Initialize", "Start Of Data Is Not Defined in  [" & .Name & "] !"
        End If
    '### EOFが定義されていない場合はエラー処理 ######################
        Dim endCell As Excel.Range
        Set endCell = .Columns(mDataCell.Column).Find(END_OF_FILE)
        If endCell Is Nothing Then
            Err.Raise 999, "CWsTreeDataReader.Initialize", "End Of File Is Not Defined in  [" & .Name & "] !"
        End If
    End With
End Sub

Private Function IParameterReader_ReadAsBoolean(ByVal dataName As String) As Boolean
    '### セルデータの読込み #########################################
    Select Case dataName:
        Case IS_CATEGORIZE:
            IParameterReader_ReadAsBoolean = mDataCell.offset(0, CATEGORIZE_COL).Value
        Case IS_VISIBLE_TOOL:
            IParameterReader_ReadAsBoolean = mDataCell.offset(0, VISIBLE_COL).Value
        Case Else
            Err.Raise 999, "CWsTreeDataReader.AsIParameterReader.ReadAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Function IParameterReader_ReadAsDouble(ByVal dataName As String) As Double
    '### このリーダーはこのインターフェイス機能を保障しない #########
    Err.Raise 999, "CWsTreeDataReader.AsIParameterReader.ReadAsDouble", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsInteger(ByVal dataName As String) As Integer
    '### このリーダーはこのインターフェイス機能を保障しない #########
    Err.Raise 999, "CWsTreeDataReader.AsIParameterReader.ReadAsInteger", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsLong(ByVal dataName As String) As Long
    '### このリーダーはこのインターフェイス機能を保障しない #########
    Err.Raise 999, "CWsTreeDataReader.AsIParameterReader.ReadAsLong", "Does Not Support This Method !"
End Function

Private Function IParameterReader_ReadAsString(ByVal dataName As String) As String
    '### セルデータの読込み #########################################
    Select Case dataName:
        Case SUPPLIER_NAME:
            If IsEmpty(mDataCell) Then
                IParameterReader_ReadAsString = NOT_DEFINE
            Else
                IParameterReader_ReadAsString = mDataCell.Value
            End If
        Case TOOL_NAME:
            If IsEmpty(mDataCell.offset(0, SUBLAYER_COL)) Then
                IParameterReader_ReadAsString = NOT_DEFINE
            Else
                IParameterReader_ReadAsString = mDataCell.offset(0, SUBLAYER_COL).Value
            End If
        Case NAME_LOCATION:
            '### 空白の場合はデフォルト値を返す #####################
            If IsEmpty(mDataCell.offset(0, LOCATION_COL)) Then
                IParameterReader_ReadAsString = DEF_NAME_MAP
            Else
                IParameterReader_ReadAsString = mDataCell.offset(0, LOCATION_COL).Value
            End If
        Case Else
            Err.Raise 999, "CWsTreeDataReader.AsIParameterReader.ReadAsString", "Can Not Find The Parameter [" & dataName & "] !"
    End Select
End Function

Private Sub IFileStream_MoveNext()
    '### データセルインデックスのインクリメント #####################
    Set mDataCell = mDataCell.offset(1, 0)
End Sub

Private Function IFileStream_IsEOR() As Boolean
    '### データセルの終了判定 #######################################
    IFileStream_IsEOR = (mDataCell.Value = END_OF_FILE)
End Function

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    '### このリーダーはこのインターフェイス機能を保障しない #########
    Err.Raise 999, "CWsTreeDataReader.AsIFileStream.SetLocation", "Does Not Support This Method !"
End Sub

Private Sub IFileStream_Rewind()
    '### スタートデータセルオブジェクトの取得 #######################
    Set mDataCell = mTreeDataSh.Range(START_CELL)
End Sub

Public Function AsIParameterReader() As IParameterReader
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IParameterReader型オブジェクト
'
'注意事項:
'
    Set AsIParameterReader = Me
End Function

Public Function AsIFileStream() As IFileStream
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IFileStreamr型オブジェクト
'
'注意事項:
'
    Set AsIFileStream = Me
End Function
