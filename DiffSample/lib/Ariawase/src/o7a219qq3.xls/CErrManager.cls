VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CErrManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   エラー情報管理用クラス
'
'目的:
'   エラー情報を集める
'   エラー情報の履歴を収集する
'
'作成者:
'   SLSI今手
'
'注意点:
'   エラー情報の履歴は以下のような場合消去されます
'   ・履歴初期化命令 ClearHistoryを実行
'   ・VBのリセットボタンを押す
'   ・エラーマネージャObjectを廃棄する
'   エラー情報クラス(CErrInfo)が必要となります
'
Option Explicit

'エラー情報蓄積コレクション
Private m_ErrInfoStorage As Collection
'エラー履歴が存在しないときに報告要求されたときに返すメッセージ
Private Const NO_HISTORY_MSG = "There is no error history"

'#Pass-OK
Private Sub Class_Initialize()
'生成時処理（初期化）
    Me.ClearHistory
End Sub

'#Pass-OK
Private Sub Class_Terminate()
'廃棄時処理
    Set m_ErrInfoStorage = Nothing
End Sub

'#Pass-OK
Public Property Get ErrorHistory(ByVal HistoryNumber As Long) As CErrInfo
'内容:
'   指定履歴番号のエラー情報Objectを提供する
'
'パラメータ:
'   [HistoryNumber]  In  履歴番号
'
'戻り値:
'   指定番号のエラー情報Object
'
'注意事項:
'   指定番号のエラー履歴が無いとき戻り値は、Nothingとなります。
'
    With m_ErrInfoStorage
        If (HistoryNumber > 0) And (HistoryNumber <= .Count) Then
            Set ErrorHistory = .Item(HistoryNumber)
        Else
            Set ErrorHistory = Nothing
        End If
    End With

End Property

'#Pass-OK
Public Property Get LastError() As CErrInfo
'内容:
'   最終（最新）エラー情報Objectの公開
'
'パラメータ:
'
'戻り値:
'   最終（最新）エラー情報object
'
'注意事項:
'   エラー履歴が存在しない時戻り値は、Nothingになります。
'
    With m_ErrInfoStorage
        If .Count > 0 Then
            Set LastError = .Item(.Count) '履歴が存在するとき最後のエラー情報を公開
        Else
            Set LastError = Nothing       '履歴が存在しないときはNothingを返す
        End If
    End With

End Property

'#Pass-OK
Public Property Get ErrorCount() As Long
'内容:
'   保存されているエラー情報の数を公開する
'
'パラメータ:
'
'戻り値:
'   保存されているエラー情報数
'
'注意事項:
'   エラー情報が保存されていないときの値は0となります
'
    ErrorCount = m_ErrInfoStorage.Count

End Property

'#Pass-OK
Public Sub ClearHistory()
'内容:
'   エラー履歴情報の初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Set m_ErrInfoStorage = New Collection

End Sub

'#Pass-OK
Public Sub Raise(ByVal ErrNumber As Long, ByVal ErrSource As String, ByVal errMsg As String, _
Optional ByVal HelpFile As String = "", Optional ByVal HelpContext As Long = 0)
'内容:
'   エラーのレイズ
'
'パラメータ:
'   [ErrNumber]  In  エラー番号
'   [ErrSource]  In  発生ソース名称
'   [ErrMsg]     In  メッセージ文字列
'   [HelpFile]   In  エラー時に表示するヘルプファイル名（オプション）
'   [HelpContext]  In エラー時に表示するヘルプファイルのヘルプID番号（オプション）
'
'戻り値:
'
'注意事項:
'   本命令を実行するとエラー情報の記録後、
'   入力されたパラメータを使用してVBAのErr.Raiseが実行されます。
'   エラー時に表示するヘルプファイル名（オプション）はFull Pathで指定して下さい
'
    Dim tmpErr As CErrInfo
    Set tmpErr = New CErrInfo

    '情報をエラーObjectに設定
    With tmpErr
        .SerialNumber = m_ErrInfoStorage.Count + 1
        .Number = ErrNumber
        .Source = ErrSource
        .Message = errMsg
        .HelpFile = HelpFile
        .HelpContext = HelpContext
        .ErrDate = Date
        .ErrTime = TIME
    End With

    'エラー履歴の登録
    Call m_ErrInfoStorage.Add(tmpErr)  'エラー履歴に今回のエラー情報を追加
        
    'VBAのErrObjectにエラーを挙げる
    With tmpErr
        Call Err.Raise(.Number, .Source, .Message, .HelpFile, .HelpContext)
    End With

End Sub

'#Pass-OK
Public Sub ReportDebugPrint()
'内容:
'   エラー履歴情報をイミディエイトWindowへ出力する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Dim retErrInfo As CErrInfo

    If m_ErrInfoStorage.Count > 0# Then
        '履歴Recordがあるときのメッセージ作成
        For Each retErrInfo In m_ErrInfoStorage
            With retErrInfo
                .DebugPrint
            End With
        Next retErrInfo
    Else
        '履歴Recordが無いときのメッセージ
        Debug.Print NO_HISTORY_MSG
    End If

End Sub

'#Pass-OK
Public Sub ReportMsgBox()
'内容:
'   エラー履歴情報をメッセージBOXへ出力する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   エラー履歴の数が多い場合は、すべてのエラーが表示されない場合があります
'
    Call MsgBox(MakeErrHistoryMsg, vbInformation, "JOB error history")

End Sub

'#Pass-OK
Public Sub ReportTxtFile(ByVal fileName As String)
'内容:
'   エラー履歴情報を指定TXTファイルへ追記出力する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   出力先のファイルは、FullPathで指定して下さい
'   指定ファイルがすでに存在している場合には、同ファイルに追記出力されます。
'   情報は、CSV形式で出力されます。
'
    Call AppendTxtFile(fileName, MakeCsvErrHistoryMsg)

End Sub

'#Pass-OK
Private Function MakeErrHistoryMsg() As String
'情報メッセージ作成用関数(MsgBox用)

    Dim retErrInfo As CErrInfo
    Dim PrintMsg As String
    
    If m_ErrInfoStorage.Count > 0# Then
        '履歴Recordがあるときのメッセージ作成
        PrintMsg = ""
        For Each retErrInfo In m_ErrInfoStorage
            With retErrInfo
                PrintMsg = PrintMsg & .Information & vbCrLf
            End With
        Next retErrInfo
        MakeErrHistoryMsg = PrintMsg
    Else
        '履歴Recordが無いときのメッセージ
        MakeErrHistoryMsg = NO_HISTORY_MSG
    End If

End Function

'#Pass-OK
Private Function MakeCsvErrHistoryMsg() As String
'CSV形式のエラーメッセージ作成用関数

    Dim retErrInfo As CErrInfo
    Dim PrintMsg As String
        
    If m_ErrInfoStorage.Count > 0# Then
        '履歴Recordがあるときのメッセージ作成
        PrintMsg = ""
        For Each retErrInfo In m_ErrInfoStorage
            With retErrInfo
                PrintMsg = PrintMsg & .CsvInformation
            End With
        Next retErrInfo
        MakeCsvErrHistoryMsg = PrintMsg
    Else
        '履歴Recordが無いときのメッセージ
        MakeCsvErrHistoryMsg = NO_HISTORY_MSG & vbCrLf
    End If

End Function

'#Pass-OK
Private Function AppendTxtFile(ByVal appendFileName As String, outPutMsg As String) As Boolean
'指定TXTファイルへの文字列、追記出力用の関数
    
    Dim fileNum As Integer
    Const FUNCTION_NAME = "AppendTxtFile"
    
    On Error GoTo OUT_PUT_LOG_ERR
    fileNum = FreeFile
    Open appendFileName For Append As fileNum
    Print #fileNum, outPutMsg
    Close fileNum
    AppendTxtFile = True
    On Error GoTo 0
    Exit Function

OUT_PUT_LOG_ERR:
    AppendTxtFile = False
    Call Err.Raise(9999, FUNCTION_NAME, appendFileName & " Output File Error")

End Function

