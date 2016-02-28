Attribute VB_Name = "XLibJobUtility"
'概要:
'   JOBやライブラリで共通に使用するユーティリティ郡
'
'目的:
'
'
'作成者:
'   SLSI今手
'
'

Option Explicit

'#Pass
Public Sub OutputErrMsg(ByVal Msg As String)
'内容:
'   エラーをメッセージと共に出力
'
'パラメータ:
'    [Msg]      In   エラー内容メッセージ
'
'戻り値:
'
'注意事項:
'
    Const NEW_JOB_ERR_NUMBER = 999

    Msg = "Error Message: " & vbCrLf & "    " & Msg & vbCrLf
    Msg = Msg & "Test Instance Name: " & vbCrLf & "    " & TheExec.DataManager.InstanceName

'    Call MsgBox(Msg, vbExclamation Or vbOKOnly, "Error")
    Call Err.Raise(NEW_JOB_ERR_NUMBER, "OutputErrMsg", Msg)

End Sub

Public Function CompareDblData(ByVal dblValue1 As Double, ByVal dblValue2 As Double, ByVal DIGIT As Long) As Boolean
'内容:
'   2つのDouble型データを指定有効桁数以降を切り捨てて比較判定する
'
'パラメータ:
'    [dblValue1]    In   比較対象データ1
'    [dblValue2]    In   比較対象データ2
'    [digit]        In   有効桁数
'
'戻り値:
'   比較判定結果
'
'注意事項:
'
    With WorksheetFunction
        CompareDblData = (.RoundDown(dblValue1, DIGIT) = .RoundDown(dblValue2, DIGIT))
    End With
End Function

Public Function RoundDownDblData(ByVal dblValue As Double, ByVal DIGIT As Long) As Double
'内容:
'   Double型データの指定有効桁数以降を切り捨てるワークシート関数のラッパー
'
'パラメータ:
'    [dblValue]     In   対象データ
'    [digit]        In   有効桁数
'
'戻り値:
'   指定された桁数のDouble型データ
'
'注意事項:
'
    RoundDownDblData = WorksheetFunction.RoundDown(dblValue, DIGIT)
End Function

Public Function GetJobRootPath() As String
'内容:
'   TheExec.Rootpathのラッパー関数
'
'戻り値:
'   IG-XLの現在ロードされているバージョンのインストールフォルダの絶対パス
'
'注意事項:
'
    GetJobRootPath = TheExec.Rootpath
End Function

Public Function GetCurrentJobName() As String
'内容:
'   TheExec.CurrentJobのラッパー関数
'
'戻り値:
'   アクティブなJOB名の取得
'
'注意事項:
'
    GetCurrentJobName = TheExec.CurrentJob
End Function

Public Function GetCurrentChanMap() As String
'内容:
'   TheExec.CurrentChanMapのラッパー関数
'
'戻り値:
'   アクティブなチャンネルマップ名の取得
'
'注意事項:
'
    GetCurrentChanMap = TheExec.CurrentChanMap
End Function

Public Function IsJobValid() As Boolean
'内容:
'   TheExec.JobIsValidのラッパー関数
'
'戻り値:
'   バリデーションが正しく実行されたかどうか
'
'注意事項:
'
    IsJobValid = TheExec.JobIsValid
End Function

Public Sub CreateListBox(ByVal selCell As Range, ByRef dataList As Collection)
'内容:
'   エクセルワークシートの任意のセルにリストボックスフォームを設定する
'
'パラメータ:
'   [wsSheet]      In   対象ワークシートオブジェクト
'   [selCell]      In   対象セルオブジェクト
'   [listBoxData]  In   リストに設定するデータコレクション
'
'注意事項:
'   リストボックスの初期選択パラメータは以下の通り
'   ①対象セルに既にパラメータが入力されている場合
'    ・リストに存在するパラメータの場合はそのパラメータを初期選択パラメータとして表示する
'    ・リストに存在しない場合は空白（ListIndex=-1）を初期選択する
'   ②対象セルにパラメータが入力されていない場合は空白（ListIndex=-1）を初期選択する
'
    '### 古いリストボックスの削除 #########################
    Const myDropName = "DropDownList"
    On Error Resume Next
    selCell.parent.DropDowns(myDropName).Delete
    On Error GoTo 0
    '### データリストが存在しない場合はEXIT ###############
    Dim listBoxData As Collection
    Set listBoxData = dataList
    If listBoxData Is Nothing Then Exit Sub
    '### リストの初期表示データの準備 #####################
    Dim dataIndex As Long
    Dim currData As Variant
    Dim IsContain As Boolean
    If IsEmpty(selCell) Then
        dataIndex = 0
    Else
        For Each currData In listBoxData
            If selCell.Value = currData Then
                dataIndex = dataIndex + 1
                IsContain = True
                Exit For
            End If
            dataIndex = dataIndex + 1
        Next currData
        If Not IsContain Then
            dataIndex = 0
        End If
    End If
    '### ウィンドウサイズの調整 ###########################
    Dim currZoom As Double
    currZoom = ActiveWindow.Zoom
    Application.ScreenUpdating = False
    ActiveWindow.Zoom = 100
    '### リストボックスの設定 #############################
    With selCell.parent.DropDowns
        'リストボックス横幅の12のオフセットはドロップボタンの横幅分
        .Add(selCell.Left, selCell.Top, selCell.width + 12, selCell.height).Name = myDropName
        For Each currData In listBoxData
            .AddItem (currData)
        Next currData
        .OnAction = "'selectData " & Chr(34) & myDropName & Chr(34) & "'"
        .ListIndex = dataIndex
    End With
    '### ウィンドウサイズの再調整 #########################
    ActiveWindow.Zoom = currZoom
    Application.ScreenUpdating = True
End Sub

Public Sub SelectData(ByVal dropName As String)
    ActiveCell.Value = ActiveSheet.DropDowns(dropName).List(ActiveSheet.DropDowns(dropName).ListIndex)
End Sub

Public Sub RunAtValidationStart()
'内容:
'   Validationスタート時に実行する関数
'
'注意事項:
'   IG-XLバージョンが3.40.10JDXXの場合、
'   IG-XLのOnVaridationStartそのものが呼ばれない
'   Validation開始時に必ず実行したいものは
'   他の手段を考える必要がある。
'

    '### Job実行中にValidationを実行した場合 ##############
    If TheExec.Flow.IsRunning Then
        MsgBox "CAUTION:" & vbCrLf & _
                "Validation starts when job is running." & vbCrLf & _
                "Please stop job running, and validate the job again.", vbExclamation, _
                "Eee-Job : IG-XL Event Handler"
    End If
    '######################################################

End Sub
