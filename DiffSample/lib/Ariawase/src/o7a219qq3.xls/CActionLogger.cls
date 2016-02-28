VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CActionLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   イメージ関連のAction実行履歴を管理するUtility
'
'目的:
'   イメージ関連のAction実行履歴を格納する
'   全実行履歴をCSV形式で一括出力する
'
'作成者:
'   a_oshima
'
'前提条件:
'   CActionInfoクラス、CActionExecutionHistoryクラス、CEeeTimerクラスを利用
'   グローバルオブジェクト TheError:CErrManager が存在すること
'   最初にInitializeメソッドをコールすること（違反時は例外9999発生）
'
'利用手順:
'   ①NewしたらすぐInitializeをコール
'   ②Action.Execute()実行ごとに、
'   　　②-1 Action.Execute()のユーザマクロ実行前にLogBeforeExecutionをコールしてもらう
'   　　②-2 ユーザマクロ実行
'   　　②-3 Action.Execute()のユーザマクロ実行後にLogAfterExecutionをコールしてもらう
'   ③CSV一括出力
'
'Action.Execute()の書式例:
'   Private Sub DummyExec(ByRef pInfo As CActionInfo)
'       Call gActionLogger.LogBeforeExecution(pInfo)
'       Dim ret As Long
'       ret = Application.Run("DummyUserMacro_f")
'       Call gActionLogger.LogAfterExecution(pInfo, ret)
'   End Sub
'
'CSV構成:
'

Option Explicit

Private mActionExecutionHistories As Collection '履歴情報：メンバはCActionExecutionHistory型
Private mEnableLogging As Boolean
Private mTimer As CEeeTimer
Private mCurrentHistoryID As Long
Private mMaxAcquireArgCount As Long   'AcquireTableのArg数がばらばらのときのCSV列数調整用

Private Const mSUCCESS As Long = 0&

Property Get GetHistory(ByRef pActionInfo As CActionInfo, ByRef pTrigger As String) As CActionExecutionHistory
'内容:
'　　アクション情報, Triggerに合致する履歴情報を返す
'
'パラメータ:
'   [pActionInfo]   IN  CActionInfo型:　　アクション情報
'   [pTrigger]   IN  String型:　　トリガー種別(Start/End)
'
'戻り値:
'   CActionExecutionHistory型  該当する履歴1件
'
'例外:
'   指定したActionInfo, Triggerに合致する履歴がない場合は例外発生

    Dim flgFound As Boolean
    flgFound = False
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        If historyObj.Trigger = pTrigger Then
            If pActionInfo.IsSameName(historyObj.ActionInfo) Then
                flgFound = True
                Exit For
            End If
        End If
    Next historyObj

    If flgFound = True Then
        Set GetHistory = historyObj
    Else
        Call TheError.Raise(9999, TypeName(Me) & ".GetHistory", "Not found History:" & pActionInfo.ToString & " Trigger:" & pTrigger)
    End If

End Property
Property Get HistoryCount() As Long
    HistoryCount = mActionExecutionHistories.Count
End Property

Property Get EnableLogging() As Boolean
    EnableLogging = mEnableLogging
End Property

Private Sub Class_Terminate()
    Set mActionExecutionHistories = Nothing
    Set mTimer = Nothing
End Sub

Private Sub Class_Initialize()
'内容:
'　　履歴情報を初期化し、タイマーをスタート
    mEnableLogging = False
    Set mTimer = New CEeeTimer
    Call Initialize
End Sub

Public Sub Initialize()
'内容:
'　　履歴情報を初期化し、タイマーをスタート
    Call ClearHistories
    mTimer.Initialize
    mTimer.StartTimer

End Sub

Public Sub ClearHistories()
'内容:
'   履歴情報を全て破棄
    Set mActionExecutionHistories = New Collection
    mCurrentHistoryID = 0&
    mMaxAcquireArgCount = 0&
End Sub

Public Sub LogBeforeExecution(ByRef pActionInfo As CActionInfo)
'内容:
'   Action.Execの実行前の履歴書き込み
'   (EnableLoggingがFalseのときは何もしない)
'
'パラメータ:
'   [pActionInfo]  IN CActionInfo型:　格納するActionInfo（格納するのはクローン）
'
'前提条件:
'   Action.Execでユーザマクロ実行直前にコールされる
'
    Call CheckInitializeCalled
'Interceptor型に修正してこのメソッドが実行＝ログ出力に変更
'    If mEnableLogging = False Then
'        Exit Sub
'    End If

    mCurrentHistoryID = mCurrentHistoryID + 1

    Dim newHistory As CActionExecutionHistory
    Set newHistory = New CActionExecutionHistory
    newHistory.RecordNumber = CStr(mCurrentHistoryID)
    newHistory.Trigger = "Start"
    newHistory.RecordDate = Date$
    newHistory.RecordTime = TIME$
    newHistory.LapTime = mTimer.LapTimer
    Set newHistory.ActionInfo = pActionInfo.Clone()

    If mMaxAcquireArgCount < pActionInfo.Parameters.ArgParameterCount Then
        mMaxAcquireArgCount = pActionInfo.Parameters.ArgParameterCount
    End If

    On Error GoTo catch
    Call mActionExecutionHistories.Add(newHistory, CStr(mCurrentHistoryID))
    On Error GoTo 0
    Exit Sub
catch:
    If Err.Number = 457 Then
        Call TheError.Raise(9999, TypeName(Me) & ".LogBeforeExecution", "HistoryID appeared towice:" & mCurrentHistoryID)
    Else
        Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

Public Sub LogAfterExecution(ByRef pActionInfo As CActionInfo, Optional ByRef pErrorNo As Long = mSUCCESS)
'内容:
'   Action.Execの実行後の履歴書き込み
'   (EnableLoggingがFalseのときは何もしない)
'
'パラメータ:
'   [pActionInfo]  IN CActionInfo型:　格納するActionInfo（格納するのはクローン）
'   [pErrorNo]   IN  Long型:   Application.Runで返った返値をそのまま渡す
'
'前提条件:
'   Action.Execでユーザマクロ実行直後にコールされる。

    Call CheckInitializeCalled
'Interceptor型に修正してこのメソッドが実行＝ログ出力に変更
'    If mEnableLogging = False Then
'        Exit Sub
'    End If

    mCurrentHistoryID = mCurrentHistoryID + 1

    Dim newHistory As CActionExecutionHistory
    Set newHistory = New CActionExecutionHistory
    newHistory.RecordNumber = CStr(mCurrentHistoryID)
    newHistory.Trigger = "End"
    newHistory.RecordDate = Date$
    newHistory.RecordTime = TIME$
    newHistory.LapTime = mTimer.LapTimer
    newHistory.ErrorNo = pErrorNo
    Set newHistory.ActionInfo = pActionInfo.Clone()

    If mMaxAcquireArgCount < pActionInfo.Parameters.ArgParameterCount Then
        mMaxAcquireArgCount = pActionInfo.Parameters.ArgParameterCount
    End If

    On Error GoTo catch
    Call mActionExecutionHistories.Add(newHistory, CStr(mCurrentHistoryID))
    On Error GoTo 0
    Exit Sub
catch:
    If Err.Number = 457 Then
        Call TheError.Raise(9999, TypeName(Me) & ".LogAfterExecution", "HistoryID appeared towice:" & mCurrentHistoryID)
    Else
        Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub


Public Function IsStarted(ByRef pActionInfo As CActionInfo) As Boolean
'内容:
'   Actionが実行開始したかどうかの問い合わせ
'
'[pActionInfo] IN CActionInfo型:  LogBeforeExecutionで指定したActionと同じテスト名・マクロ名をもつ
'
'戻り値:
'   Boolean型　　True:実行開始済、または実行完了、またはエラー　　False:未実行
'
    Dim histObj As CActionExecutionHistory
    On Error GoTo catch
    Set histObj = GetHistory(pActionInfo, "Start")
    On Error GoTo 0
    IsStarted = True
    Exit Function
catch:
    IsStarted = False
End Function


Public Function IsCompleted(ByRef pActionInfo As CActionInfo) As Boolean
'内容:
'   Actionが実行完了しているかどうかの問い合わせ
'
'[pActionInfo] IN CActionInfo型:  LogBeforeExecutionで指定したActionと同じテスト名・マクロ名をもつ
'
'戻り値:
'   Boolean型　　True:実行完了　　False:実行未完了、エラー発生時
'
'前提条件:
'   LogBeforeExecution未実行の場合False
'
    If IsStarted(pActionInfo) Then
        Dim historyObj As CActionExecutionHistory
        On Error GoTo catch
        Set historyObj = GetHistory(pActionInfo, "End")
        On Error GoTo 0

        If historyObj.ErrorNo = mSUCCESS Then
            IsCompleted = True
        Else
            IsCompleted = False
        End If
        Set historyObj = Nothing
    Else
        IsCompleted = False
    End If
    Exit Function
catch:
    IsCompleted = False
End Function

Public Function IsError(ByRef pActionInfo As CActionInfo) As Boolean
'内容:
'   Actionがエラーかどうかの問い合わせ
'
'[pActionInfo] IN CActionInfo型:  LogBeforeExecutionで指定したActionと同じテスト名・マクロ名をもつ
'
'戻り値:
'   Boolean型　　True:エラー　　False:エラー発生せず、または実行中
'
'前提条件:
'   LogBeforeExecution未実行の場合False
'
    If IsStarted(pActionInfo) Then
        Dim historyObj As CActionExecutionHistory
        On Error GoTo catch
        Set historyObj = GetHistory(pActionInfo, "End")
        On Error GoTo 0

        If historyObj.ErrorNo <> mSUCCESS Then
            IsError = True
        Else
            IsError = False
        End If
        Set historyObj = Nothing
    Else
        IsError = False
    End If
    Exit Function
catch:
    IsError = False
End Function

Private Function FileAndJobIndex() As String
    Dim RetFileName As String
    Dim RetJobName As String

    RetFileName = TheExec.ExcelHandle.ActiveWorkbook.Path & "\" & TheExec.ExcelHandle.ActiveWorkbook.Name
    RetJobName = TheExec.CurrentJob                             'JobList の JobName より。無ければファイル名

    'ファイル名を挿入
    FileAndJobIndex = """[Test Program]"",""[" & RetFileName & "]""" & vbCrLf & _
         """[Job]"",""[" & RetJobName & "]""" & vbCrLf
End Function

Public Function CSVIndex() As String
'内容:
'   CSVの表題を返す
'
'戻り値:
'   String型   CSVの表題
'
'前提条件:
'   履歴が1件以上記録されていること
'
    Dim str As String
    str = ""
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        str = historyObj.CSVIndex(mMaxAcquireArgCount) _
                & vbCrLf
        Exit For
    Next historyObj
    CSVIndex = str
End Function

Public Function ToCSVString() As String
'内容:
'   CSVの内容を返す
'
'戻り値:
'   String型   CSVの内容（複数行：1行1件）
'
'前提条件:
'   履歴が1件以上記録されていること
'
    Dim str As String
    str = ""
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        str = str & historyObj.ToCSVString(mMaxAcquireArgCount) _
                & vbCrLf
    Next historyObj
    ToCSVString = str
End Function

Public Function ToString() As String
'内容:
'   履歴の内容を返す
'
'戻り値:
'   String型   履歴の内容
'
'前提条件:
'   履歴が1件以上記録されていること
'
    Dim str As String
    str = ""
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        str = str & "History[" & vbCrLf
        str = str & historyObj.ToString(4)
        str = str & "]" & vbCrLf
    Next historyObj
    ToString = str
End Function

Public Function CsvInformation() As String
    CsvInformation = ToCSVString()
End Function
Public Function Information() As String
    Information = ToString()
End Function

Public Sub ReportHistoriesWithMsgbox()
'内容:
'   履歴情報を1件ずつメッセージBOXへ出力する
'
'注意事項:
'   内容の文字数が多い場合は、すべて表示されない場合があります
'
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        MsgBox "Report:" & vbCrLf & historyObj.ToString()
    Next historyObj
End Sub
Public Sub ReportHistoriesWithDebugPrint()
'内容:
'   履歴情報を1件ずつイミディエイトウインドウへ出力する
'
'注意事項:
'   内容の文字数が多い場合は、すべて表示されない場合があります
'
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        Debug.Print "Report:" & vbCrLf & historyObj.ToString()
    Next historyObj
End Sub


Private Sub CheckInitializeCalled()
    If mTimer Is Nothing Then
        TheError.Raise 9999, TypeName(Me) & ".CheckInitializeCalled", "Call Initialize method first!"
    End If
End Sub

Public Function SaveHistoryLog(ByVal pFullFilePath As String) As Boolean
'内容:
'   履歴の内容をCSVファイルに出力（追記モード）
'   ファイルが存在する場合は新規作成
'
'[pFullFilePath] IN String型:  出力ファイルの名前（フルパス）
'
'戻り値:
'   Boolean型   True:記録成功    False:記録失敗
'
'前提条件:
'   履歴が1件以上記録されていること
'   ファイルパスが256文字以内（Dir関数の制約）
'
    If mActionExecutionHistories.Count > 0 Then
        SaveHistoryLog = WriteTxtFile(pFullFilePath, FileAndJobIndex & CSVIndex & ToCSVString)
    Else
        SaveHistoryLog = False
    End If
End Function

Private Function WriteTxtFile(ByVal pOutFilePath As String, ByRef pStr As String) As Boolean

    If Len(pOutFilePath) > 256 Then
        Call TheError.Raise(9999, TypeName(Me) & ".WriteTxtFile", "FilePath is too long (max 256):" & pOutFilePath)
    End If
    If Dir(pOutFilePath) = vbNullString Then
         If CreateTxtFile(pOutFilePath) = False Then
            Call TheError.Raise(9999, TypeName(Me) & ".WriteTxtFile", "Failed to create file:" & pOutFilePath)
         End If
    End If
    WriteTxtFile = AppendTxtFile(pOutFilePath, pStr)
End Function

Private Function CreateTxtFile(ByVal pOutFilePath As String) As Boolean

    Dim FileNo As Integer
    On Error GoTo catch
    FileNo = FreeFile
    Open pOutFilePath For Output As #FileNo
    Close #FileNo
    On Error GoTo 0

    CreateTxtFile = True
    Exit Function

catch:
    CreateTxtFile = False
    Call TheError.Raise(9999, TypeName(Me) & ".CreateTxtFile", "Output File Error:" & pOutFilePath)

End Function


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
    Call TheError.Raise(9999, FUNCTION_NAME, appendFileName & " Output File Error")

End Function






