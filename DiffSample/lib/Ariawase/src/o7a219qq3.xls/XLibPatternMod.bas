Attribute VB_Name = "XLibPatternMod"
'概要:
'   パターン制御ライブラリ群
'
'目的:
'   Ⅰ：パターンシートからデータを取得しロードする
'   Ⅱ：テラダインパターン制御APIをいくつかの機能ブロックとして切り出し
'
'   Revision History:
'   Data        Description
'   2010/07/27　■TOPTフレームワーク導入による仕様変更
'               　StartPattern/RunPattern/StartStopPattern関数の引数へタイミング情報を追加
'   2010/07/29　■他モジュールのサブルーチンBreak(), DebugMsg()の利用停止
'                 ファンクションLoadPatternFile()から返値を返せるように変更
'   2012/12/20  H.Arikawa
'               StopPatternを編集。テスタータイプに応じてHaltの動作を変更する。
'               LoadPatternFileを編集。PatGrpsシート読み込み部変更。
'
'作成者:
'   0145206097

Option Explicit

Public Function LoadPatternFile() As Long
'内容:
'   PatGrpsシートからパタングループを読み込みロードする
'
'パラメータ:
'
'返値:
'   成功:0(TL_SUCCESS)、失敗:1(TL_ERROR)
'
'注意事項:
'   ワークシート"PatGrps"が見つからないとき、
'   またはワークシートからのデータ取得時にエラーが発生した場合は
'   メッセージボックスで警告を表示し、TL_ERRORを返します。

    Const DATA_SHEET_NAME = "PatGrps"
    Const PATTERN_GROUP = "GroupName"

    Dim targetWorkSheet As Worksheet
    Dim patGroupName As Range
    Dim tsbName As Range
    Dim tsbSheetName As String
    Dim rowIndex As Long

    Call StopPattern
    TheHdw.Digital.Patterns.UnloadAll
    TheHdw.Digital.Patgen.TimeoutEnable = False

    On Error GoTo ErrHandler
    Set targetWorkSheet = getWorkSheet(DATA_SHEET_NAME)

    With targetWorkSheet
        
        Set patGroupName = .Range(PATTERN_GROUP)
        Set tsbName = .Range("E3")
        
        rowIndex = patGroupName.Row + 1

        Do While .Cells(rowIndex, patGroupName.Column) <> ""
            tsbSheetName = .Cells(rowIndex, tsbName.Column)
            TheHdw.Digital.Timing.Load (tsbSheetName)
            TheHdw.Digital.Patterns.pat(.Cells(rowIndex, patGroupName.Column)).Load
            rowIndex = rowIndex + 1
        Loop

    End With
    LoadPatternFile = TL_SUCCESS
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation Or vbOKOnly, "Error"
    TheHdw.Digital.Patterns.UnloadAll
    TheHdw.Digital.Patgen.TimeoutEnable = False
    LoadPatternFile = TL_ERROR
End Function

Public Sub StartPattern( _
    ByVal patGroupName As String, _
    Optional ByVal startLabel As String = "START", _
    Optional timeSetName As String = "", _
    Optional categoryName As String = "", _
    Optional selectorName As String = "" _
)
'内容:
'   パターンをバーストする
'
'パラメータ:
'[patGroupName] In  パターングループ名
'[startLabel]   In  スタートラベル
'[timeSetName]  In  タイミングセット名
'[categoryName] In　カテゴリ名
'[selectorName] In  セレクタ名
'
'注意事項:
'   バースト終了を待たずに制御がプログラムに戻る
'

    With TheHdw.Digital
        .Timing.Load timeSetName, categoryName, selectorName
        .Patterns.pat(patGroupName).Start startLabel
    End With

End Sub

Public Sub RunPattern( _
    ByVal patGroupName As String, _
    Optional ByVal startLabel As String = "START", _
    Optional timeSetName As String = "", _
    Optional categoryName As String = "", _
    Optional selectorName As String = "" _
)
'内容:
'   パターンをバーストする
'
'パラメータ:
'[patGroupName] In  パターングループ名
'[startLabel]   In  スタートラベル
'[timeSetName]  In  タイミングセット名
'[categoryName] In　カテゴリ名
'[selectorName] In  セレクタ名
'
'注意事項:
'   バースト終了を待ってから制御がプログラムに戻る
'

    With TheHdw.Digital
        .Timing.Load timeSetName, categoryName, selectorName
        .Patterns.pat(patGroupName).Run startLabel
    End With

End Sub

Public Sub StartStopPattern( _
    ByVal patGroupName As String, _
    ByVal startLabel As String, _
    ByVal stopLabel As String, _
    Optional timeSetName As String = "", _
    Optional categoryName As String = "", _
    Optional selectorName As String = "" _
)
'内容:
'   パターンをバーストする
'
'パラメータ:
'[patGroupName] In  パターングループ名
'[startLabel]   In  スタートラベル
'[stopLabel]    In  ストップラベル
'[timeSetName]  In  タイミングセット名
'[categoryName] In　カテゴリ名
'[selectorMame] In  セレクタ名
'
'注意事項:
'   指定のストップラベルにHALTを挿入後、指定のスタートラベルからバーストする
'

    With TheHdw.Digital
        .Timing.Load timeSetName, categoryName, selectorName
        .Patterns.pat(patGroupName).StartStop startLabel, stopLabel
    End With

End Sub

Public Sub StopPattern()
'内容:
'   パターンバーストを終了する
'
'パラメータ:
'
'注意事項:
'        デコーダで使用するパターンのパターンバースト終了を行う。
'
    With TheHdw.Digital.Patgen
    
        If .IsRunningAnySite = True Then
            .Ccall = True
            .HaltWait
            .Ccall = False
        End If

    End With

End Sub

Public Sub StopPattern_Halt()
'内容:
'   パターンバーストを終了する
'
'パラメータ:
'
'注意事項:
'        テスターがIP750の時は、Halt止めしないようにIf文で分岐。
'

    With TheHdw.Digital.Patgen
    
        If TesterType = "IP750" Then
            If .IsRunningAnySite = True Then
                .Ccall = True
                .HaltWait
                .Ccall = False
            End If
        Else
            If .IsRunningAnySite = True Then
                   .Halt
            End If
        End If

    End With

End Sub

Public Sub SetTimeOut( _
    Optional ByVal runOutStatus As Boolean = False, _
    Optional ByVal runOutTime As Long = 5 _
)
'内容:
'   パターンのタイムアウト条件を設定する
'
'パラメータ:
'[runOutStatus] In  タイムアウトステータス
'[runOutTime]   In  タイムアウト時間
'
'注意事項:
'

    With TheHdw.Digital.Patgen
        .TimeoutEnable = runOutStatus
        .TIMEOUT = runOutTime
    End With

End Sub

'{ XLibDcModに同様の関数あり
Private Function getWorkSheet( _
    ByRef SheetName As String _
) As Worksheet

    On Error GoTo errMsg

    Set getWorkSheet = Worksheets(SheetName)

    Exit Function

errMsg:

    Call Err.Raise(Err.Number, Err.Source, "Not " & SheetName & " Sheet Exist !")
    '日本語エラーメッセージ出力
'    Call Err.Raise(Err.Number, Err.Source, sheetName & " シートがありません")

    Set getWorkSheet = Nothing

End Function
'}

Private Function TimingLoad_f() As Long
    Dim myInstance As String
    Dim myTSBName As String
    
    If First_Exec = 0 Then
        myInstance = TheExec.DataManager.InstanceName
        myTSBName = Mid(myInstance, InStr(myInstance, "_") + 1, Len(myInstance))
        TheHdw.Digital.Timing.Load myTSBName
    End If
    
End Function
Public Sub PatGrpsColorMake()

    With Worksheets("PatGrps").Range("E3")
        .Interior.color = RGB(0, 0, 255)
        .Font.color = vbWhite
    End With
    
End Sub
