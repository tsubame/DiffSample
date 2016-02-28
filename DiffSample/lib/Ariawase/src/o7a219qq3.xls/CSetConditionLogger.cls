VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetConditionLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   測定条件設定ログ提供機能クラス
'
'目的:
'   測定条件設定マクロの動作ログを登録保存する
'

Option Explicit

Private m_ExecuteLogs As Collection           '設定履歴保存用コレクション
Private m_MaxArgNumber As Long                'Argパラメータの最大番号
Private Const ERROR_NUMBER As Long = 9999     'エラー発生時に返すエラー番号

Public Property Let MaxArgNumber(ByVal pArgNumber As Long)
'内容:
'   Arg番号の最大値
'
'パラメータ:
'   [pArgNumber]  In  Long型: Arg番号最大値
'
'戻り値:
'
'注意事項:
'
    
    m_MaxArgNumber = pArgNumber

End Property

Public Property Get MaxArgNumber() As Long
'内容:
'   Arg番号の最大値
'
'パラメータ:
'
'戻り値:
'   Arg番号最大値
'
'注意事項:
'
    
    MaxArgNumber = m_MaxArgNumber

End Property

Public Sub AddRecord(ByVal pExecTime As Double, ByRef pMacroInfo As CSetFunctionInfo)
'内容:
'   ログ情報の追加登録
'
'パラメータ:
'   [pExecTime]     In  Double型：マクロの実行時間[sec]
'   [pMacroInfo]    In  CSetFunctionInfo型：マクロ情報Object
'
'戻り値:
'
'注意事項:
'

    Dim SetHistoryObj As CSetConditionHistory
    Set SetHistoryObj = New CSetConditionHistory
       
    With SetHistoryObj
        .RecordDate = CStr(Date)
        .RecordTime = CStr(TIME)
        .ExecTestInstanceName = TheExec.DataManager.InstanceName   'セット関数実行時のテストインスタンス名記録
        .ExecTime = pExecTime                                      'セット関数の実行時間記録
        Set .SetFunctionInfo = pMacroInfo.Clone                    'セット関数の情報のコピーを記録
    End With

    Call m_ExecuteLogs.Add(SetHistoryObj)

End Sub

Public Sub Initialize()
'内容:
'   機能初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    
    Set m_ExecuteLogs = New Collection

End Sub

Public Function Record(ByVal pRecordNumber As Long) As CSetConditionHistory
'内容:
'   履歴番号指定によるログ情報の公開
'
'パラメータ:
'   [pRecordNumber]     In  Long型：履歴番号
'
'戻り値:
'   条件設定履歴情報(CSetConditionHistory)
'
'注意事項:
'   指定番号の履歴情報が存在しない場合は、戻り値は Nothing になります。
'
    
    If m_ExecuteLogs.Count > 0 Then
        Set Record = m_ExecuteLogs.Item(pRecordNumber)
    Else
        Set Record = Nothing
    End If

End Function

Public Property Get RecordCount() As Long
'内容:
'   記録済情報要素数公開
'
'パラメータ:
'
'戻り値:
'   記録済情報要素数
'
'注意事項:
'
    
    RecordCount = m_ExecuteLogs.Count

End Property

Public Sub ClearRecord()
'内容:
'   記録されている全ての情報の消去
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    
    Set m_ExecuteLogs = New Collection

End Sub

Public Sub SaveRecord(ByVal pOutputFileName As String)
'内容:
'   記録されているログをTXTファイルへ追記する
'
'パラメータ:
'   [pOutputFileName]     In  String型：出力ファイル名(FullPath)
'
'戻り値:
'
'注意事項:
'
    
    If m_ExecuteLogs.Count = 0 Then
        Call AppendTxtFile(pOutputFileName, "There is no SetCondition history")
        Exit Sub
    End If
    
    'データ出力LOOP
    Dim RecordCnt As Long
    For RecordCnt = 1 To m_ExecuteLogs.Count
        On Error GoTo SAVE_HISTORY_LOG_ERR
        Call AppendTxtFile(pOutputFileName, MakeCsvHistoryInfo(RecordCnt))
        On Error GoTo 0
    Next RecordCnt
    
    Exit Sub

SAVE_HISTORY_LOG_ERR:
'設定履歴ログのファイル出力時のエラー処理
    Call RaiseError(pOutputFileName & " write error" & vbCrLf & _
    "The output to the file of the history log was not done")

End Sub

Private Function MakeCsvHistoryInfo(ByVal pHistoryNumber As Long) As String
'内容:
'   '条件設定履歴情報のCSV Format作成関数
'
'パラメータ:
'   [pHistoryNumber]　In Long型：  取得する履歴番号の指定
'
'戻り値:
'   出力する実行ログ
'注意事項:
'

    Dim makeMsg As String
    Dim MsgHeader As String
    Dim ArrayNumber As Long
    Dim HistoryInfo As CSetConditionHistory
    
    On Error GoTo GET_HISTORY_ERR
    Set HistoryInfo = Me.Record(pHistoryNumber)
    On Error GoTo 0
    
    Const MAX_ARGCOUNT As Long = 10 'Argの最大個数
    
    'CSV形式情報作成処理
    With HistoryInfo
        makeMsg = ""

        makeMsg = makeMsg & _
        pHistoryNumber & "," & _
        "" & .RecordDate & "" & "," & _
        "" & .RecordTime & "" & "," & _
        "" & .ExecTestInstanceName & "" & "," & _
        "" & MakeStringMsg(.SetFunctionInfo.ConditionName) & "" & "," & _
        "" & MakeStringMsg(.SetFunctionInfo.FunctionName) & "" & "," & _
        "" & MacroSetting2String(.SetFunctionInfo.IsEnableMacro) & "" & "," & _
            .ExecTime
        'パラメータ部作成
        Dim ArgCount As Long
        ArgCount = 0
        For ArrayNumber = 0 To .SetFunctionInfo.ArgParameterCount - 1 Step 1
            makeMsg = makeMsg & "," & MakeStringMsg(.SetFunctionInfo.Arg(ArrayNumber))
            ArgCount = ArgCount + 1
        Next ArrayNumber
        
        'マクログループラベルまで
        Dim ToExecSettingName As Long
        For ToExecSettingName = 1 To (MAX_ARGCOUNT - ArgCount) Step 1
            makeMsg = makeMsg & ","
        Next ToExecSettingName
        makeMsg = makeMsg & "," & .SetFunctionInfo.ExecuteSettingName
    End With
    
    If pHistoryNumber = 1 Then
        'ヘッダ情報あり
        Dim RetFileName   As String
        Dim RetJobName   As String
        RetFileName = TheExec.ExcelHandle.ActiveWorkbook.Path & "\" & TheExec.ExcelHandle.ActiveWorkbook.Name
        RetJobName = TheExec.CurrentJob                             'JobList の JobName より。無ければファイル名
        MsgHeader = "[Test Program],[" & RetFileName & "]" & vbCrLf & _
                    "[Job],[" & RetJobName & "]" & vbCrLf & _
                    "[RecordNumber],[RecordDate],[RecordTime],[TestInstanceName],[ConditionName],[ProcedureName],[Execute],[ExecTime(s)]"
        Dim ArgNumber As Long
        For ArgNumber = 0 To m_MaxArgNumber Step 1
            MsgHeader = MsgHeader & "," & "[Arg" & ArgNumber & "]"
        Next ArgNumber
        MsgHeader = MsgHeader & "," & "[ExecuteSetting]"
        MakeCsvHistoryInfo = MsgHeader & vbCrLf & makeMsg
    Else
        'ヘッダ情報なし
        MakeCsvHistoryInfo = makeMsg
    End If
    
    Exit Function

GET_HISTORY_ERR:
'履歴データ取り出しエラー処理
    Call RaiseError("HistoryNo=" & pHistoryNumber & " data is nothing" & vbCrLf & _
    "As for the number, there is the need that is a value from 1" & " to " & Me.RecordCount)

End Function

'テキストファイルに追記出力する為の関数
Private Function AppendTxtFile(ByVal pTgtFileName As String, ByVal pOutPutMsg As String) As Boolean
    
    Dim fileNum As Integer
    Dim errFunctionName As String
    
    On Error GoTo OUT_PUT_TXT_ERR
    
    fileNum = FreeFile
    Open pTgtFileName For Append As fileNum
    Print #fileNum, pOutPutMsg
    Close fileNum
    
    AppendTxtFile = True
    
    Exit Function

OUT_PUT_TXT_ERR:
    Close fileNum
    Call RaiseError(pTgtFileName & " file write error")
    AppendTxtFile = False

End Function

Private Function MakeStringMsg(ByVal pMessage As String) As String
'CSVの文字列要素を作成する関数
    Dim makeMsg As String
        
    If Len(pMessage) > 0 Then
        '1:CR、LF、CR+LF等の改行文字を置き換える
        '2:文字列のカンマを \, に置き換える
        '3:文字列のダブルコーテーションを ""に置き換える
        makeMsg = ReplaceDoubleQuotation(ReplaceComma(ReplaceCrLf(pMessage)))
        
        '文字列パラメータの長さが1以上の場合は、文字列をダブルコーテーションでくくる
        makeMsg = Chr(34) & makeMsg & Chr(34)
        MakeStringMsg = makeMsg
    
    Else
        MakeStringMsg = pMessage
    End If
    
End Function

Private Function ReplaceCrLf(ByVal pMessage As String) As String
'文字列に含まれている改行文字を置き換えて返す関数
'vbCrLf-->\r\n ,vbCr --> \r, vbLf --> \n
    Dim makeMsg As String
    makeMsg = Replace(Replace(Replace(pMessage, vbCrLf, "\r\n"), vbCr, "\r"), vbLf, "\n")
    ReplaceCrLf = makeMsg
End Function

Private Function ReplaceDoubleQuotation(ByVal pMessage As String) As String
'文字列に含まれているDoubleQuotationを置き換えて返す関数
'" --> ""
    Dim makeMsg As String
    makeMsg = Replace(pMessage, """", """" & """")
    ReplaceDoubleQuotation = makeMsg
End Function

Private Function ReplaceComma(ByVal pMessage As String) As String
'文字列に含まれているCommaを置き換えて返す関数
', --> \,
    Dim makeMsg As String
    makeMsg = Replace(pMessage, ",", "\,")
    ReplaceComma = makeMsg
End Function

'エラーレイズ用関数
Private Sub RaiseError(ByVal pErrMsg As String)
    If Not (TheError Is Nothing) Then
        'エラーマネージャへ
        Call TheError.Raise(ERROR_NUMBER, TypeName(Me), pErrMsg)
    Else
        'VBのErrへ
        Call Err.Raise(ERROR_NUMBER, TypeName(Me), pErrMsg)
    End If
End Sub

'マクロの有効無効設定から文字列を返す
Private Function MacroSetting2String(ByVal pIsEnable As Boolean) As String
    If pIsEnable = True Then
        MacroSetting2String = "enable"
    Else
        MacroSetting2String = "disable"
    End If
End Function


