Attribute VB_Name = "XLibTOPT_FW"
'概要:
'   TOPT対応したFrameWork
'   Image ACQTBL シートのからの呼び出しモジュール、および Image ACQTBL シートの情報に基づきオブジェクト群の生成
'
'   Revision History:
'       Data        Description
'       2010/04/28  AcquireFrameWorkを実行する機能を実装した
'                   （FWSetCondition / FWImageAcquire / FWPostImageAcquire）
'       2010/06/11  データ構造見直しのため、プログラムコードを変更した
'       2010/06/22  実行情報リセット機能を実装した
'                   CheckTermination機能を実装した
'                   Arg0,Arg1のエラー処理機能を実装した
'                   FWのエラー処理機能を実装した
'       2010/07/02  AcquireInstanceの作成時の不具合を修正した（Count<>0で登録）
'
'作成者:
'   tomoyoshi.takase
'   0145184346
'

Option Explicit

Private Const ERR_NUMBER = 9999                  'エラー時に渡すエラー番号
Private Const CLASS_NAME = "XLibTOPT_FW"         'このクラスの名前
Private Const FW_KEY = "ImageACQTBL Sheet(ARG0)" 'フレームワークの動作決定する情報の場所。エラー出力用

Private mActionLogger As CActionLogger ' ActionLoggerを保持する
Private mAcquireInstance As Collection ' AcquireFrameWorkを保持する
Private mImageCheckErrorMsg As String
Private mImageCheckCounter As Long

Public Sub AcquireInitialize(ByRef clsActionLogger As CActionLogger, ByRef reader As CWorkSheetReader)
'内容:
'　ImageACQTBL シートから情報収集して、各インスタンスの生成を行います。
'
'パラメータ:
'   [clsActionLogger]  In  ActionLoggerを保持する
'   [reader]           In  ImageAcquireTableシート情報を保持する
'
'戻り値:
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  ActionLoggerインスタンス生成  #####
    Set mActionLogger = clsActionLogger


    '#####  ImageAcquireTable情報を取得する  #####
    Dim strmReader As IFileStream
    Dim paramReader As IParameterReader
    Set strmReader = reader
    Set paramReader = reader


    '#####  AcquireInstanceを生成する  #####
    Dim strFrameworkName As String
    Dim strInstanceName As String
    Dim strAutoAcquire As String
    Dim strLastInsName As String
    Dim strArg0 As String
    Dim strArg1 As String
    Dim strArg2 As String
    Dim strArg3 As String
    Dim clsAcqAct As IAcquireAction
    Dim clsAcqIns As CAcquireInstance
    Set clsAcqIns = New CAcquireInstance
    Set mAcquireInstance = New Collection

    mImageCheckCounter = 0
    mImageCheckErrorMsg = ""
    
    Do While strmReader.IsEOR <> True
        strLastInsName = strInstanceName
        strFrameworkName = paramReader.ReadAsString("Macro Name")
        strInstanceName = paramReader.ReadAsString("Instance Name")
        strAutoAcquire = paramReader.ReadAsString("Auto Acquire")

        '#####  Instance名称の確認  #####
        If strInstanceName = "" Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                                "Can Not Found Instance Name ! => " & strFrameworkName & "(ImageACQTBL Sheet)")
        ElseIf strLastInsName <> "" And strInstanceName <> strLastInsName Then
            On Error GoTo VBAErrHandler
            If clsAcqIns.Count <> 0 Then
                mAcquireInstance.Add clsAcqIns, strLastInsName
            End If
            Set clsAcqIns = New CAcquireInstance
            On Error GoTo ErrHandler
        End If

        '#####  Arg0,Arg1の確認  #####
        strArg0 = paramReader.ReadAsString("Arg0@Parameters")
        If strArg0 = "" Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                                "Can Not Found Instance Name !(Arg0) => " & strFrameworkName & _
                                "(" & strInstanceName & ")" & " (ImageACQTBL Sheet)")
        End If
        If strArg0 <> strInstanceName Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                                "Acquire Instance Name and Arg0 is not same! => " & strFrameworkName & _
                                "(Inst:" & strInstanceName & " <=> Arg0:" & strArg0 & ")" & _
                                " (ImageACQTBL Sheet)")
        End If
        If strFrameworkName = "FWImageAcquire" Or strFrameworkName = "FWPostImageAcquire" Then
            strArg1 = paramReader.ReadAsString("Arg1@Parameters")
            If strArg1 = "" Then
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                                    "Can Not Found Acquire Macro Name !(Arg1) => " & strFrameworkName & _
                                    "(" & strInstanceName & ")" & " (ImageACQTBL Sheet)")
            End If
        End If

        '#####  FrameWorkごとの処理  #####
        If strAutoAcquire <> "nop" Then
            If strFrameworkName = "FWSetCondition" Then
                Set clsAcqAct = New CSetAction
'                Call clsAcqAct.Initialize(paramReader, mActionLogger)
                Call clsAcqAct.Initialize(paramReader)
                clsAcqIns.SetAction = clsAcqAct
            ElseIf strFrameworkName = "FWImageAcquire" Then
                Set clsAcqAct = New CAcquireAction
'                Call clsAcqAct.Initialize(paramReader, mActionLogger)
                Call clsAcqAct.Initialize(paramReader)
                clsAcqIns.SetAction = clsAcqAct
            ElseIf strFrameworkName = "FWPostImageAcquire" Then
                Set clsAcqAct = New CPostAcquireAction
'                Call clsAcqAct.Initialize(paramReader, mActionLogger)
                Call clsAcqAct.Initialize(paramReader)
                clsAcqIns.SetAction = clsAcqAct
            Else
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                                    "Illegal Frame Work Name ! => " & strFrameworkName & "(ImageACQTBL Sheet)")
            End If
        End If

        '#####  Arg2,3の画像パラメータのチェック  #####
        If strFrameworkName = "FWImageAcquire" Then
            strArg2 = paramReader.ReadAsString("Arg2@Parameters")
            strArg3 = paramReader.ReadAsString("Arg3@Parameters")
            If ChkImageParamter(strArg2, strArg3) = False Then
                Call StockErr(strFrameworkName, strInstanceName, strArg2, strArg3)
            End If
        End If
        
        '#####  次の行に移動  #####
        Call strmReader.MoveNext

        '#####  最終行の処理  #####
        If strmReader.IsEOR = True Then
            If clsAcqIns.Count <> 0 Then
                mAcquireInstance.Add clsAcqIns, strInstanceName
            End If
            Set clsAcqAct = Nothing
            Set clsAcqIns = Nothing
        End If
    
    Loop

    If mImageCheckCounter > 0 Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME, _
                            "ImageParameterCheck Error!" & vbCrLf & "ImageAcquireSheet Arg2-3" & vbCrLf & vbCrLf & mImageCheckErrorMsg)
    End If
    
    Set strmReader = Nothing
    Set paramReader = Nothing

    Exit Sub

'#####  エラーメッセージ処理＆終了  #####
VBAErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                        "This Instance Name is already setting." & "(" & strLastInsName & ")")
    Exit Sub
ErrHandler:
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub

End Sub

Private Sub StockErr(ByVal pMacroName As String, ByVal pInstanceName As String, ByVal pBaseName As String, ByVal pStrBitDepth As String)
    mImageCheckErrorMsg = mImageCheckErrorMsg & _
                            "FrameWorkname : " & pMacroName & vbCrLf & _
                            "InstanceName  : " & pInstanceName & vbCrLf & _
                            "ImageBaseName : " & pBaseName & vbCrLf & _
                            "ImageBitDepth : " & pStrBitDepth & vbCrLf
    mImageCheckCounter = mImageCheckCounter + 1
End Sub

Private Function ChkImageParamter(ByVal pBaseName As String, ByVal pStrBitDepth As String) As Boolean
'パラメータチェックしてだめならFalse
    Dim pBitDepth As CIdpBitDepth
    Dim PNum As Long
    On Error GoTo NOTHING_IMAGE
    Set pBitDepth = New CIdpBitDepth
    Call pBitDepth.SetValue(pStrBitDepth)
    PNum = TheIDP.PlaneManager(pBaseName).Count(pBitDepth.GetValue)
    If PNum > 0 Then
        ChkImageParamter = True
    Else
        ChkImageParamter = False
    End If
    Exit Function
NOTHING_IMAGE:
    ChkImageParamter = False
End Function

Private Function GetAcqIns(pFuncName As String, pArg0 As String) As CAcquireInstance
    On Error GoTo VBAErrHandler
    Set GetAcqIns = mAcquireInstance.Item(pArg0)
    Exit Function
VBAErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & "." & pFuncName, _
                        "Illegal Instance Name." & "(" & pArg0 & ")" & "(ImageACQTBL Sheet)")
End Function

Public Function FWSetCondition() As Long
'内容:
'　TOPT.GetArgumentListで得た情報でSetConditionを実行します。
'　TOPT.Start または TOPT Auto Acquire で実行されます。
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
    
    '#####  ImageAcquireTableのパラメータを取得  #####
    Dim ArgImageAcqtbl() As String

    Call TheHdw.TOPT.GetArgumentList(ArgImageAcqtbl)
    If ArgImageAcqtbl(0) = "" Then
        Call TheError.Raise(ERR_NUMBER, "FWSetCondition", FW_KEY & " is Nothing!")
    End If
    
'    TheExec.Datalog.WriteComment ArgImageAcqtbl(0)
'    Call StartTime
    '#####  AcquireFrameWorkのインスタンスを取得する  #####
    Dim clsAcqIns As CAcquireInstance
    
    Set clsAcqIns = GetAcqIns("FWSetCondition", ArgImageAcqtbl(0))

    '#####  動いているTOPT情報の提示用  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = ArgImageAcqtbl(0)
    TheIDP.PlaneList.ACQTBL_Arg1 = ArgImageAcqtbl(1)

    '#####  AcquireFrameWorkのインスタンスを実行する  #####
    FWSetCondition = clsAcqIns.Execute("FWSetCondition")

    Set clsAcqIns = Nothing

    '#####  TOPT FW の終了情報  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = "-"
    TheIDP.PlaneList.ACQTBL_Arg1 = "-"
    
'    Call StopTime
    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    FWSetCondition = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function FWImageAcquire() As Long
'内容:
'　TOPT.GetArgumentListで得た情報でImageAcquireを実行します。
'　TOPT.Start または TOPT Auto Acquire で実行されます。
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

    Dim Acqsite As Long
    
    For Acqsite = 0 To nSite
        If TheExec.sites.site(Acqsite).Active = False Then
            TheExec.sites.site(Acqsite).Active = True
            If Flg_FailSiteImage(Acqsite) = False Then
                '@@@ DUT情報毎に中身を替えないといけない。@@@
                Call DisconnectAllDevicePins(Acqsite)                 'FailSite All OPEN   '2012/11/16 175JobMakeDebug
                Call GND_DisConnect(Acqsite)                          '2012/11/16 175JobMakeDebug
                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                Flg_FailSiteImage(Acqsite) = True
            End If
            TheExec.sites.site(Acqsite).Active = False
        End If
    Next

    '#####  ImageAcquireTableのパラメータを取得  #####
    Dim ArgImageAcqtbl() As String

    Call TheHdw.TOPT.GetArgumentList(ArgImageAcqtbl)
    If ArgImageAcqtbl(0) = "" Then
        Call TheError.Raise(ERR_NUMBER, "FWImageAcquire", FW_KEY & " is Nothing!")
    End If


    '#####  AcquireFrameWorkのインスタンスを取得する  #####
    Dim clsAcqIns As CAcquireInstance

    Set clsAcqIns = GetAcqIns("FWImageAcquire", ArgImageAcqtbl(0))

    '#####  動いているTOPT情報の提示用  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = ArgImageAcqtbl(0)
    TheIDP.PlaneList.ACQTBL_Arg1 = ArgImageAcqtbl(1)

    '#####  AcquireFrameWorkのインスタンスを実行する  #####
    FWImageAcquire = clsAcqIns.Execute("FWImageAcquire")

    Set clsAcqIns = Nothing

    '#####  TOPT FW の終了情報  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = "-"
    TheIDP.PlaneList.ACQTBL_Arg1 = "-"

    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    FWImageAcquire = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function FWPostImageAcquire() As Long
'内容:
'　TOPT.GetArgumentListで得た情報でPostImageAcquireを実行します。
'　TOPT.Start または TOPT Auto Acquire で実行されます。
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


    '#####  ImageAcquireTableのパラメータを取得する  #####
    Dim ArgImageAcqtbl() As String

    Call TheHdw.TOPT.GetArgumentList(ArgImageAcqtbl)
    If ArgImageAcqtbl(0) = "" Then
        Call TheError.Raise(ERR_NUMBER, "FWPostImageAcquire", FW_KEY & " is Nothing!")
    End If


    '#####  AcquireFrameWorkのインスタンスを取得する  #####
    Dim clsAcqIns As CAcquireInstance
    
    Set clsAcqIns = GetAcqIns("FWPostImageAcquire", ArgImageAcqtbl(0))

    '#####  動いているTOPT情報の提示用  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = ArgImageAcqtbl(0)
    TheIDP.PlaneList.ACQTBL_Arg1 = ArgImageAcqtbl(1)

    '#####  AcquireFrameWorkのインスタンスを実行する  #####
    FWPostImageAcquire = clsAcqIns.Execute("FWPostImageAcquire")

    Set clsAcqIns = Nothing

    '#####  TOPT FW の終了情報  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = "-"
    TheIDP.PlaneList.ACQTBL_Arg1 = "-"

    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    FWPostImageAcquire = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function ExecuteAcquireAction(ByRef strAcquireInstanceName As String, Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'内容:
'   AcquireFrameWorkインスタンスを実行する
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

    '#####  AcquireFrameWorkインスタンスを取得する  #####
    Dim clsAcqIns As CAcquireInstance
    Set clsAcqIns = GetAcqIns("ExecuteAcquireAction", strAcquireInstanceName)

    '#####  AcquireFrameWorkを実行する  #####
    ExecuteAcquireAction = clsAcqIns.ToptStart(strAcquireFrameWorkName)
    If ExecuteAcquireAction = TL_ERROR Then
        Exit Function
    End If

    Set clsAcqIns = Nothing


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    ExecuteAcquireAction = TL_ERROR
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function


End Function

Public Function RetryAcquireAction(ByRef strAcquireInstanceName As String, ByRef strAcquireFrameWorkName As String) As Integer
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


    '#####  AcquireFrameWorkの実行履歴を削除する  #####
    RetryAcquireAction = StartClearStatus(strAcquireInstanceName, strAcquireFrameWorkName)
    If RetryAcquireAction = TL_ERROR Then
        Exit Function
    End If
    

    '#####  AcquireFrameWorkを再実行する  #####
    TheIDP.PlaneBank.IsOverwriteMode = True
    RetryAcquireAction = ExecuteAcquireAction(strAcquireInstanceName, strAcquireFrameWorkName)
'    TheIDP.PlaneBank.IsOverWriteMode = False


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    RetryAcquireAction = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function StartClearStatus(ByRef strAcquireInstanceName As String, Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'内容:
'   AcquireFrameWorkインスタンスの実行履歴を削除する
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


    '#####  AcquireFrameWorkを取得する  #####
    Dim clsAcqIns As CAcquireInstance
    Set clsAcqIns = GetAcqIns("StartClearStatus", strAcquireInstanceName)


    '#####  AcquireFrameWorkを実行する  #####
    StartClearStatus = clsAcqIns.ClearStatus(strAcquireFrameWorkName)
    If StartClearStatus = TL_ERROR Then
        Exit Function
    End If

    Set clsAcqIns = Nothing


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    StartClearStatus = TL_ERROR
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function


End Function

Public Sub ResetStatus()
'内容:
'   AcquireFrameWorkインスタンスの実行履歴をクリアする
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'


    '#####  AcquireFrameWorkを取得する  #####
    Dim i As Integer
    Dim clsAcqIns As CAcquireInstance
    
    For i = 1 To mAcquireInstance.Count
        Set clsAcqIns = mAcquireInstance.Item(i)
        clsAcqIns.Reset
    Next i

    Set clsAcqIns = Nothing


    '#####  終了  #####
    Exit Sub


End Sub

Public Function CheckTermination(ByRef strAcquireInstanceName As String, Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'内容:
'   AcquireFrameWorkインスタンスの実行を確認する
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


    '#####  AcquireFrameWorkインスタンスを取得する  #####
    Dim clsAcqIns As CAcquireInstance
    Set clsAcqIns = GetAcqIns("CheckTermination", strAcquireInstanceName)


    '#####  AcquireFrameWorkインスタンスのStatusによって、TOPTを実行する  #####
    CheckTermination = clsAcqIns.CheckTermination(strAcquireFrameWorkName)

    Set clsAcqIns = Nothing


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    CheckTermination = TL_ERROR
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function


End Function

Public Sub DestroyTOPTFW()
    Set mActionLogger = Nothing
    Set mAcquireInstance = Nothing
End Sub

Public Sub CreateTOPTFWIfNothing()
'内容:
'   As Newの代替として初回に呼ばせるインスタンス生成処理
'
'パラメータ:
'   なし
'
'注意事項:
'
    
    On Error GoTo ErrHandler
    If mAcquireInstance Is Nothing Then
        Call AcquireInitialize(GetActionLoggerInstance, GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeAcquire))
    End If
    Exit Sub
ErrHandler:
    Call DestroyTOPTFW
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub EnableReadImage(ByVal pFlag As Boolean, ByVal pPathName As String)
    Dim pAcqObj As CAcquireInstance
    For Each pAcqObj In mAcquireInstance
        Call pAcqObj.EnableReadImage(pFlag, pPathName)
    Next pAcqObj
End Sub

Public Sub EnableSaveImage(ByVal pFlag As Boolean, ByVal pPathName As String)
    Dim pAcqObj As CAcquireInstance
    For Each pAcqObj In mAcquireInstance
        Call pAcqObj.EnableSaveImage(pFlag, pPathName)
    Next pAcqObj
End Sub

Public Sub EnableShowImage(ByVal pFlag As Boolean)
    Dim pAcqObj As CAcquireInstance
    For Each pAcqObj In mAcquireInstance
        Call pAcqObj.EnableShowImage(pFlag)
    Next pAcqObj
End Sub

Public Sub EnableInterceptor(ByVal pFlag As Boolean, ByRef pLogger As CActionLogger)
    Dim pAcqObj As CAcquireInstance
    For Each pAcqObj In mAcquireInstance
        Call pAcqObj.EnableInterceptor(pFlag, pLogger)
    Next pAcqObj
End Sub
