VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImageEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   撮像テストインスタンスを生成するクラス
'
'   Revision History:
'       Data        Description
'       2010/04/28  撮像テストインスタンスを生成＆実行する機能を実装した
'       2010/05/12  プログラムコードを整理した
'       2010/05/19  SheetReader参照キーワードを固定ラベルに変更した
'       2010/05/25  SheetReaderを外部から受け取るように変更した
'       2010/05/31  Error処理を変更した
'       2010/06/11  プログラムコードを整理した
'
'作成者:
'   0145184346
'

Option Explicit

Private Const CLASS_NAME  As String = "CImageEngine" ' Class名称を保持する

Private m_FlagInterceptor As Boolean

Private mImageTestInstance As Collection       ' 全ての撮像テストインスタンスを保持する
Private mImageTestAction As CImageTestInstance ' 特定の撮像テストインスタンスを保持する
Private mActionLogger As CActionLogger         ' 実行LogデータをダンプするActionLoggerを保持する
Private mSheetTestInstance As CWorkSheetReader ' TestInstanceシートを保持する

Public Function CreateScenario() As Integer
'内容:
'   撮像テストインスタンスを生成する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  撮像テストインスタンスを生成する  #####
    CreateScenario = loadTestInstance


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    CreateScenario = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Strings.Chr$(13) & Strings.Chr$(13) & Err.Description
    Exit Function


End Function

Private Function loadTestInstance() As Integer
'内容:
'   TestInstanceシートを読み込み、撮像テストインスタンスを生成する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  Sheet名称＆Template名称を設定する  #####
    Dim strToptTemplateTool As String ' Template名称を保持する
    strToptTemplateTool = "XLibToptFrameWork_T"


    '#####  TestInstanceシートのデータを取得する  #####
    Dim IFileStream As IFileStream
    Dim IParamReader As IParameterReader
    
    With mSheetTestInstance
        Set IFileStream = .AsIFileStream
        Set IParamReader = .AsIParameterReader
    End With


    '#####  撮像テストインスタンスに、設定しているデータを取得＆保持する  #####
    Dim colAcqIns As Collection            ' AcquireInstance名称を保持する
    Dim strAcquireInstanceName() As String ' AcquireInstance名称を保持する
    Dim strAcqInsName As Variant           ' AcquireInstance名称を保持する
    Dim strTestInstanceName As String      ' TestInstance名称を保持する
    Dim clsUsrMcrAct As IUserMacroAction
    Set mImageTestInstance = New Collection

    Do While Not IFileStream.IsEOR
        
        If IParamReader.ReadAsString("Name@Test Procedure") = strToptTemplateTool Then

            '#####  TestInstance名称を取得する  #####
            Set mImageTestAction = New CImageTestInstance

            strTestInstanceName = IParamReader.ReadAsString("Test Name")
            With mImageTestAction
                .Initialize
                .SetTestInstanceName = strTestInstanceName
            End With

            '#####  AcquireInstance名称を取得する  #####
            Set colAcqIns = New Collection
            
            If IParamReader.ReadAsString("Arg0@Other Parameters") <> "" Then
                strAcquireInstanceName = Split(Replace(IParamReader.ReadAsString("Arg0@Other Parameters"), " ", ""), ",")
                For Each strAcqInsName In strAcquireInstanceName
                    colAcqIns.Add strAcqInsName
                Next strAcqInsName
                mImageTestAction.SetAcquireInstance = colAcqIns
            End If
            
            Set colAcqIns = Nothing
            
            '#####  UserMacro名称を取得する  #####
            Set clsUsrMcrAct = New CUserMacroAction
            
            If IParamReader.ReadAsString("Arg1@Other Parameters") <> "" Then
                clsUsrMcrAct.SetUserMacroName = IParamReader.ReadAsString("Arg1@Other Parameters")
                mImageTestAction.SetUserMacro = clsUsrMcrAct
            End If
            
            Set clsUsrMcrAct = Nothing

            '#####  撮像テストインスタンス情報を保持する  #####
            mImageTestInstance.Add mImageTestAction, mImageTestAction.GetTestInstanceName
            
            Set mImageTestAction = Nothing
        
        End If

        '#####  次の行へ移動する  #####
        IFileStream.MoveNext

    Loop

    Set IFileStream = Nothing
    Set IParamReader = Nothing


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    loadTestInstance = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Strings.Chr$(13) & Strings.Chr$(13) & Err.Description
    Exit Function


End Function

Public Function SetScenario() As Integer
'内容:
'   撮像テストを実行するための準備をする
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  撮像テスト情報が無ければ、撮像テスト情報を生成する  #####
    If mImageTestInstance Is Nothing Then
        SetScenario = Me.CreateScenario
        If SetScenario = TL_ERROR Then
            Set mImageTestInstance = Nothing
            Exit Function
        End If
    End If


    '#####  撮像テスト実行を準備する  #####
    Set mImageTestAction = mImageTestInstance.Item(TheExec.DataManager.InstanceName)


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:

    '#####  テンプレート未使用のため、撮像テストインスタンスを生成する  #####
    SetScenario = createTestInstance
    Exit Function


End Function

Private Function createTestInstance() As Integer
'内容:
'   テンプレート未使用時の撮像テストインスタンスを生成する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  パラメータを取得する  #####
    Dim strArg() As String ' Parameterを保持する
    Dim lngArgCnt As Long  ' Parameter数を保持する

    Call TheExec.DataManager.GetArgumentList(strArg, lngArgCnt)
    
    If lngArgCnt = 0 Then
        Call TheError.Raise(9999, CLASS_NAME & ".createTestInstance", "Can Not Create Image Test Instance ! " & vbCrLf _
                            & "You Have To Check Test Instance Sheet. ( " & TheExec.DataManager.InstanceName & " )")
    ElseIf lngArgCnt > 2 Then
        If strArg(0) = "" And strArg(1) = "" Then
            Call TheError.Raise(9999, CLASS_NAME & ".createTestInstance", "Can Not Create Image Test Instance ! " & vbCrLf _
                                & "You Have To Check Test Instance Sheet. ( " & TheExec.DataManager.InstanceName & " )")
        End If
    End If


    '#####  テストインスタンス名称を取得する  #####
    Dim clsImgTstIns As CImageTestInstance
    Set clsImgTstIns = New CImageTestInstance
    
    With clsImgTstIns
        .Initialize
        .SetTestInstanceName = TheExec.DataManager.InstanceName
    End With


    '#####  Acquireインスタンスを生成する  #####
    Dim colAcqIns As Collection       ' AcquireInstance名称を保持する
    Dim strAcqireInstance() As String ' AcquireInstance名称を保持する
    Dim strTmp As Variant             ' AcquireInstance名称を保持する
    Set colAcqIns = New Collection
    
    If strArg(0) <> "" Then
        strAcqireInstance = Split(Replace(strArg(0), " ", ""), ",")
        For Each strTmp In strAcqireInstance
            colAcqIns.Add strTmp
        Next strTmp
        clsImgTstIns.SetAcquireInstance = colAcqIns
    End If
    
    Set colAcqIns = Nothing
    
    
    '#####  UserMacroインスタンスを生成する  #####
    Dim clsUsrMcr As IUserMacroAction
    Set clsUsrMcr = New CUserMacroAction

    If lngArgCnt >= 2 Then
        If strArg(1) <> "" Then
            clsUsrMcr.SetUserMacroName = strArg(1)
            clsImgTstIns.SetUserMacro = clsUsrMcr
        End If
    End If

    Set clsUsrMcr = Nothing


    '#####  撮像テストインスタンスを格納する  #####
    Set mImageTestAction = clsImgTstIns
    mImageTestInstance.Add clsImgTstIns, clsImgTstIns.GetTestInstanceName

    Set clsImgTstIns = Nothing


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    createTestInstance = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function Execute() As Integer
'内容:
'   撮像テストを実行する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  撮像テストインスタンスの存在を確認する  #####
    If mImageTestAction Is Nothing Then
        Execute = Me.SetScenario
        If Execute = TL_ERROR Then
            Set mImageTestAction = Nothing
            Exit Function
        End If
    End If


    '#####  撮像テストのAcquireFrameWorkを実行する  #####
    Dim i As Integer

    With mImageTestAction
        If .GetAcquireInstanceCount <> 0 Then
            For i = 1 To .GetAcquireInstanceCount Step 1
                Execute = ExecuteAcquireAction(.GetAcquireInstanceName(i))
                If Execute = TL_ERROR Then
                    Exit Function
                End If
            Next i
        End If
    End With


    '#####  撮像テストのUserMacroを実行する  #####
    With mImageTestAction
        If .GetUserMacroCount <> 0 Then
            Execute = .Execute
            If Execute = TL_ERROR Then
                Exit Function
            End If
        End If
    End With

    Set mImageTestAction = Nothing


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    Execute = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function RetryAcquire(ByRef strAcquireInstanceName As String, Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'内容:
'   実行情報を削除して、AcquireFrameWorkを再実行する
'
'パラメータ:
'   [strAcquireInstanceName]   In  AcquireInstance名称を保持する
'   [strAcquireFrameWorkName]  In  AcquireFrameWork名称を保持する
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWorkを再実行する  #####
    If strAcquireInstanceName <> "" Then
        RetryAcquire = RetryAcquireAction(strAcquireInstanceName, strAcquireFrameWorkName)
        If RetryAcquire = TL_ERROR Then
            Call TheError.Raise(9999, CLASS_NAME & ".RetryAcquire", "Can Not Execute RetryAcquire ! " & vbCrLf _
                                & "You Have To Check Parameters. ( " & strAcquireInstanceName & " / " & strAcquireFrameWorkName & " )")
        End If

    Else
        Call TheError.Raise(9999, CLASS_NAME & ".RetryAcquire", "Can Not Execute RetryAcquire ! " & vbCrLf _
                            & "You Have To Set Acquire Instance Name.")
    End If


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    RetryAcquire = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Sub EnableInterceptor(ByVal pFlag As Boolean, ByRef pLogger As CActionLogger)
'まだ
    
    Dim pObj As CImageTestInstance
    For Each pObj In mImageTestInstance
        Call pObj.EnableInterceptor(pFlag, pLogger)
    Next pObj

End Sub

Public Sub Initialize(ByRef clsActionLogger As CActionLogger, ByRef clsWorkSheetReader As CWorkSheetReader)
    Set mImageTestInstance = Nothing
    Set mImageTestAction = Nothing
    Set mActionLogger = clsActionLogger
    Set mSheetTestInstance = clsWorkSheetReader
End Sub

Private Sub Class_Initialize()
    Set mImageTestInstance = Nothing
    Set mImageTestAction = Nothing
    Set mActionLogger = Nothing
    Set mSheetTestInstance = Nothing
End Sub

Private Sub Class_Terminate()
    Set mImageTestInstance = Nothing
    Set mImageTestAction = Nothing
    Set mActionLogger = Nothing
    Set mSheetTestInstance = Nothing
End Sub

