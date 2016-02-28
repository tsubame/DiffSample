Attribute VB_Name = "XLibOutputInfoMod"
Option Explicit

'{
'情報の出力を行う機能を有する関数類は、このモジュールに集めたい
'}

Private dbgInfHeaderLength As Long  'デバック情報出力ヘッダの総文字数格納用

Public Sub WriteDefectInfoHeader(ByVal testName As String, ByVal SiteNumber As Integer)
'内容:
'   点欠陥情報のヘッダをデータログへ出力する
'
'[testName]   In  表示させるテスト名称
'[siteNumber] In  表示させるサイト番号
'
    Dim infMsg As String

    infMsg = "***** " & testName & " DEFECT ADDRESS & DATA (SITE:" & SiteNumber & ") *****"
    Call WriteComment(infMsg)

End Sub

Public Sub WriteDebugInfoHeader(ByVal infoMessage As String, ByVal SiteNumber As Integer)
'内容:
'   デバック情報のヘッダをデータログへ出力する
'
'[infoMessage] In  表示させるデバック情報名称
'[siteNumber]  In  表示させるサイト番号
'
'注意事項
'dbgInfHeaderLength変数に出力文字数を格納する
'
    Dim infoMsg As String

    infoMsg = "***** " & infoMessage & " (SITE:" & SiteNumber & ") *****"
    
    dbgInfHeaderLength = Len(infoMsg)
    
    Call WriteComment(infoMsg)

End Sub

Public Sub WriteDebugInfoFooter()
'内容:
'   デバック情報のフッタをデータログへ出力する
'
'注意事項:
'  WriteDebugInfoHeaderサブルーチンとペアで使用する
'  dbgInfHeaderLength変数の値を使用する
'
    Dim msgCounter As Long
    Dim outputFooter As String
        
    For msgCounter = 1 To dbgInfHeaderLength Step 1
        outputFooter = outputFooter & "*"
    Next msgCounter
        
    Call WriteComment(outputFooter)

End Sub

Public Sub WriteComment(ByVal outPutMsg As String, Optional ByVal outPutFileName As String = "")
'内容:
'   コメント情報をデータログへ出力する
'
'[outPutMsg]       In  データログへ出力するメッセージ
'[outPutFileName]  In  ファイルへ出力する時のファイル名
'
'注意事項:
'   outPutFileNameはオプション。
'   指定がない場合ファイルへの情報出力は行われない
'
    Call mf_OutPutComment(outPutMsg)
    
    If outPutFileName <> "" Then
        Call mf_AppendTxtFile(outPutFileName, outPutMsg)
    End If

End Sub

Private Sub mf_OutPutComment(ByVal outPutMsg As String)
'内容:
'   情報をデータログWindowへ出力する
'
'[OutPutMsg] In  表示させるメッセージ
'
    TheExec.Datalog.WriteComment outPutMsg

End Sub

Private Function mf_AppendTxtFile(ByVal appendFileName As String, outPutMsg As String) As Boolean
'内容:
'   情報を指定されたテキストファイルへ追記する
'
'[appendFileName] In  情報を追記するファイル名
'[outPutMsg]      In  追記するメッセージ
'
'戻り値：
'   実行結果ステータス
'    エラーなし：True
'    エラーあり：False
'
    Dim fileNum As Integer
    Dim errFunctionName As String
    
    On Error GoTo OUT_PUT_LOG_ERR
    errFunctionName = mf_AppendTxtFile
    
    fileNum = FreeFile
    Open appendFileName For Append As fileNum
    Print #fileNum, outPutMsg
    Close fileNum
    
    mf_AppendTxtFile = True
    
    Exit Function

OUT_PUT_LOG_ERR:
    Call MsgBox(appendFileName & " Output File Error", vbFalse Or vbCritical, "@" & errFunctionName)
    mf_AppendTxtFile = False
'   Stop

End Function
