VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcquireInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   AcquireInstanceの情報を保持＆実行するクラス
'
'   Revision History:
'       Data        Description
'       2010/06/11  AcquireInstanceを実行する機能を実装した
'       2010/06/22  AcquireInstanceのCheckTermination機能を実装した
'                   実行情報リセット機能を実装した
'       2010/07/02  ActionのCountを参照する機能を実装した
'       2011/02/10  Interceptor、ReadImage、SaveImageの実装。
'                   制御はAcqシート、Enable～メソッドシリーズの２通り。設定はEnable～シリーズが優先されます。
'
'作成者:
'   0145184346
'

Option Explicit

Private Const CLASS_NAME = "CAcquireInstance" ' Class名称を保持する

Private mAcquireInstanceAction As Collection ' AcquireInstance情報を保持する

Private m_ActionLogger As CActionLogger

'Enable Flag が有効かどうか
Private m_ValidEnableReadImage As Boolean
Private m_ValidEnableSaveImage As Boolean
Private m_ValidEnableShowImage As Boolean
'Private m_ValidEnableInterceptor As Boolean

'Enable Flag の設定
Private m_FlagEnableReadImage As Boolean
Private m_FlagEnableSaveImage As Boolean
Private m_FlagEnableShowImage As Boolean
Private m_FlagEnableInterceptor As Boolean

Public Function Execute(ByRef strAcquireFrameWorkName As String) As Integer
'内容:
'   AcquireFrameWorkを実行する
'
'パラメータ:
'   [strAcquireFrameWorkName]  In  AcquireFrameWork名称を保持する
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    '#####  AcquireFrameWorkを取得する  #####
    On Error GoTo VBAErrHandler
    Dim clsAcqAct As IAcquireAction

    Set clsAcqAct = mAcquireInstanceAction.Item(strAcquireFrameWorkName)
    
    
    '#####  AcquireFrameWorkを実行する  #####
    On Error GoTo ErrHandler
'    Execute = clsAcqAct.Execute
    'ユーザーマクロからの返り値はIG-XLに返さない。
    clsAcqAct.Execute

    Set clsAcqAct = Nothing


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
VBAErrHandler:
    Execute = TL_ERROR
    Call TheError.Raise(9999, CLASS_NAME & ".Execute", "Can Not Found Frame Work Instance : " & strAcquireFrameWorkName)
    Exit Function
ErrHandler:
    Execute = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function ToptStart(Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'内容:
'   AcquireFrameWorkインスタンスの実行を確認する
'
'パラメータ:
'   [strAcquireFrameWorkName]  In  AcquireFrameWork名称を保持する
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWorkインスタンスの実行を確認する  #####
    Dim lngExeCnt As Integer ' AcquireFrameWork実行数を保持する
    Dim clsAcqAct As IAcquireAction

    For Each clsAcqAct In mAcquireInstanceAction
        With clsAcqAct
            If strAcquireFrameWorkName = "" Or strAcquireFrameWorkName = .GetFrameWorkName Then
                lngExeCnt = lngExeCnt + 1
                Call .ToptStart
                ToptStart = .GetStatus
                If ToptStart = TL_ERROR Then
                    Exit Function
                End If
            End If
        End With
    Next clsAcqAct

    Set clsAcqAct = Nothing


    '#####  AcquireFrameWorkインスタンスを実行有無の確認  #####
    If strAcquireFrameWorkName <> "" And lngExeCnt = 0 Then
        ToptStart = TL_ERROR
        Call TheError.Raise(9999, CLASS_NAME & ".ToptStart", "Can Not Execute Acquire Frame Work ! " _
                            & "You Have To Check AcquireFrameWork Name. ( " & strAcquireFrameWorkName & " )")
    End If


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    ToptStart = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function ClearStatus(Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'内容:
'   AcquireFrameWorkインスタンスの実行履歴を削除する
'
'パラメータ:
'   [strAcquireFrameWorkName]  In  AcquireFrameWork名称を保持する
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWorkインスタンスの実行履歴を削除する  #####
    Dim lngExeCnt As Integer ' AcquireFrameWork実行数を保持する
    Dim clsAcqAct As IAcquireAction

    For Each clsAcqAct In mAcquireInstanceAction
        With clsAcqAct
            If strAcquireFrameWorkName = "" Or strAcquireFrameWorkName = .GetFrameWorkName Then
                lngExeCnt = lngExeCnt + 1
                ClearStatus = .ClearStatus
                If ClearStatus = TL_ERROR Then
                    Exit Function
                End If
            End If
        End With
    Next clsAcqAct

    Set clsAcqAct = Nothing


    '#####  AcquireFrameWorkインスタンスを実行有無の確認  #####
    If strAcquireFrameWorkName <> "" And lngExeCnt = 0 Then
        ClearStatus = TL_ERROR
        Call TheError.Raise(9999, CLASS_NAME & ".ClearStatus", "Can Not Execute Clear Status ! " _
                            & "You Have To Check AcquireFrameWork Name. ( " & strAcquireFrameWorkName & " )")
    End If


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    ClearStatus = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function CheckTermination(Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'内容:
'   AcquireFrameWorkインスタンスの実行履歴を確認する
'
'パラメータ:
'   [strAcquireFrameWorkName]  In  AcquireFrameWork名称を保持する
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWorkインスタンスの実行履歴を確認する  #####
    Dim lngExeCnt As Integer ' AcquireFrameWork実行数を保持する
    Dim clsAcqAct As IAcquireAction

    For Each clsAcqAct In mAcquireInstanceAction
        With clsAcqAct
            If strAcquireFrameWorkName = "" Or strAcquireFrameWorkName = .GetFrameWorkName Then
                lngExeCnt = lngExeCnt + 1
                If .GetStatus = -1 Then
                    CheckTermination = .ToptStart
                    If CheckTermination = TL_ERROR Then
                        Exit Function
                    End If
                End If
            End If
        End With
    Next clsAcqAct

    Set clsAcqAct = Nothing


    '#####  AcquireFrameWorkインスタンスを実行有無の確認  #####
    If strAcquireFrameWorkName <> "" And lngExeCnt = 0 Then
        CheckTermination = TL_ERROR
        Call TheError.Raise(9999, CLASS_NAME & ".CheckTermination", "Can Not Execute CheckTermination ! " _
                            & "You Have To Check AcquireFrameWork Name. ( " & strAcquireFrameWorkName & " )")
    End If


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    CheckTermination = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Sub Reset()
'内容:
'   AcquireFrameWorkインスタンスの実行履歴をクリアする
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'


    '#####  AcquireFrameWorkインスタンスの実行履歴を削除する  #####
    Dim clsAcqAct As IAcquireAction

    For Each clsAcqAct In mAcquireInstanceAction
        clsAcqAct.Reset
    Next clsAcqAct

    Set clsAcqAct = Nothing


    '#####  終了  #####
    Exit Sub


End Sub

Public Property Let SetAction(ByRef IAcquireAction As IAcquireAction)
    If mAcquireInstanceAction Is Nothing Then
        Set mAcquireInstanceAction = New Collection
    End If
    Set IAcquireAction = InsertWrap(IAcquireAction, IAcquireAction.AcquireFunctionInfo)
    mAcquireInstanceAction.Add IAcquireAction, IAcquireAction.GetFrameWorkName
End Property

Public Property Get Count() As Long
    If Not mAcquireInstanceAction Is Nothing Then
        Count = mAcquireInstanceAction.Count
    Else
        Count = 0
    End If
End Property

Public Sub Initialize()
    Set mAcquireInstanceAction = Nothing
End Sub

Private Function ChangeConstructionObj()
'要件
'・参照を差し替え（AcquireActionInstance）
'・インスタンスの順番をきっちり並べることが可能。順番が合ってないと無意味な動作となる。
'　参照の持ち方
'　Acquire→Interceptor→Dump→Fake→実体
'・参照を変更することで実行フローが変更される
'
'留意事項
'Interceptorは任意のタイミングでオブジェクト構成を変更できる必要がある。
'Dump、Fakeクラスはとりあえず、Init時に構築を想定。任意のタイミングで変更可能にするかどうか・・・
'インスタンスのリスト構造はCAcquireInstanceがコントロールする。
'動くか否か＝インスタンスがあるか否か
'　Fake、Interceptor
'動くか否か＝インスタンスがあるか否か＋フラグ制御(Show、dump)
'　dump
'
'実行内容
'・リスト構造になってるインスタンスから、実体を取り出す。
'　現状がどういう構成になっているかは考慮しない。
'　実体クラスの型によって、割り込むクラスに制限が必要。
'　　　　　　　　　　　Interceptor　　　Dump　　　　　Fake
'CSetAction　　　　　　　　○　　　　　　×　　　　　　×
'CAcquireAction　　　　　　○　　　　　　○　　　　　　△(CAcquireFakeAction限定)
'CPostAcquireAction　　　　○　　　　　　○　　　　　　△(CPostAcquireFakeAction限定)

    Dim pFlagRead As Boolean
    Dim pFlagInterceptor As Boolean
    Dim pFlagSave As Boolean
    Dim pFlagShow As Boolean
    
    Dim FWmacro As IAcquireAction

    Dim pCollection As Collection
    Set pCollection = New Collection
    
    'CAcquireFunctionInfo、pFlagInterceptorの情報に合わせて再構成
    Dim pSubstance As IAcquireAction
    Dim pInfo As CAcquireFunctionInfo
    
    For Each FWmacro In mAcquireInstanceAction
        
        Set pSubstance = SearchSubstance(FWmacro)
        Set pInfo = pSubstance.AcquireFunctionInfo
        Set pSubstance = InsertWrap(pSubstance, pInfo)
        
        'Refer Collectionが持っている参照を修正。
        Call pCollection.Add(pSubstance, pSubstance.GetFrameWorkName)
    
    Next FWmacro
    
    Set mAcquireInstanceAction = Nothing
    Set mAcquireInstanceAction = pCollection

End Function

Private Function InsertWrap(ByRef pSubstance As IAcquireAction, ByRef pInfo As CAcquireFunctionInfo) As IAcquireAction

    '割り込むか否か
    Dim pReadImage As Boolean
    Dim pSaveImage As Boolean
    Dim pShowImage As Boolean
    Dim pInterceptor As Boolean

    Dim pFakeObj As CAcquireFakeAction
    Dim pPostFakeObj As CPostAcquireFakeAction
    Dim pInterceptorObj As CAcquireActionInterceptor
    Dim pDumpObj As CAcquireDumpAction
    
    Dim pSetObj As IAcquireActionInstance       'インスタンスのセッター、ゲッター用

    Call FixFlag(pInfo, pReadImage, pSaveImage, pShowImage, pInterceptor)
    
    '##### 各割り込みインスタンス生成 #####
    '### Fake ###
    'Fake(ACQ Sheet Arg5)
    If Not pInfo.macroName = "FWSetCondition" Then      'SetConditionの時は記述に関係なく無視
        If pReadImage = True Then
            If pInfo.macroName = "FWImageAcquire" Then
                Set pFakeObj = New CAcquireFakeAction
            ElseIf pInfo.macroName = "FWPostImageAcquire" Then
                Set pPostFakeObj = New CPostAcquireFakeAction
            Else
                Call TheError.Raise(9999, "CAcquireInstance", "FrameWorkName is different.")
            End If
        Else
            Set pFakeObj = Nothing
            Set pPostFakeObj = Nothing
        End If
    End If
    '### Dump ###
    'Dump(ACQ Sheet Arg4 or Arg7)
    If Not pInfo.macroName = "FWSetCondition" Then      'SetConditionの時は記述に関係なく無視
        If pSaveImage = True Or pShowImage = True Then
            Set pDumpObj = New CAcquireDumpAction
            pDumpObj.FlagSaveImage = pSaveImage
            pDumpObj.FlagShow = pShowImage
        Else
            Set pDumpObj = Nothing
        End If
    End If
    
    '### Interceptor ###
    If pInterceptor = True Then
        Set pInterceptorObj = New CAcquireActionInterceptor
        Set pInterceptorObj.actionLogger = m_ActionLogger
    Else
        Set pInterceptorObj = Nothing
    End If
        
        
    '##### リスト構造を生成 #####
    'Fake(ACQ Sheet Arg5)
    If Not pFakeObj Is Nothing Then
        Set pSetObj = pFakeObj                              'キャスト
        Set pSetObj.AcquireActionInstance = pSubstance      '実体参照をセット
        Set pSubstance = pSetObj                            '
    End If
    'PostFake(ACQ Sheet Arg5)
    If Not pPostFakeObj Is Nothing Then
        Set pSetObj = pPostFakeObj                              'キャスト
        Set pSetObj.AcquireActionInstance = pSubstance      '実体参照をセット
        Set pSubstance = pSetObj                            '
    End If
    'Dump(ACQ Sheet Arg4 or Arg7)
    If Not pDumpObj Is Nothing Then
        Set pSetObj = pDumpObj                              'キャスト
        Set pSetObj.AcquireActionInstance = pSubstance      '実体参照をセット
        Set pSubstance = pSetObj                            '
    End If
    'Interceptor(pFlagInterceptor)
    If Not pInterceptorObj Is Nothing Then
        Set pSetObj = pInterceptorObj                       'キャスト
        Set pSetObj.AcquireActionInstance = pSubstance      '実体参照をセット
        Set pSubstance = pSetObj                            '
    End If
    
    Set InsertWrap = pSubstance

End Function

Private Sub FixFlag(pInfo As CAcquireFunctionInfo, ByRef pReadImage As Boolean, ByRef pSaveImage As Boolean, ByRef pShowImage As Boolean, ByRef pInterceptor As Boolean)
'クラスが割り込むかどうかを決定する
        
    Dim pAfiArg4 As String
    Dim pAfiArg5 As String
    Dim pAfiArg6 As String
    Dim pAfiArg7 As String
    Dim pAfiArg8 As String

    'pInfo.Arg(4):Show
    'pInfo.Arg(5):Read
    'pInfo.Arg(6):FilePath
    'pInfo.Arg(7):Write
    'pInfo.Arg(8):FilePath
    
    '大小文字区別しないためUCase
    pAfiArg4 = UCase(pInfo.Arg(4))
    pAfiArg5 = UCase(pInfo.Arg(5))
    pAfiArg6 = pInfo.Arg(6)
    pAfiArg7 = UCase(pInfo.Arg(7))
    pAfiArg8 = pInfo.Arg(8)
    
    'シート情報、モジュールフラグを含めて最終的に割り込むクラスを決定する。
    '### Fake ###
    If m_ValidEnableReadImage = True Then
        pReadImage = m_FlagEnableReadImage
    Else
        If pAfiArg5 = "READ" Then
            pReadImage = True
        Else
            pReadImage = False
        End If
    End If
    
    '### Dump ###
    If m_ValidEnableShowImage = True Then
        pShowImage = m_FlagEnableShowImage
    Else
        If pAfiArg4 = "SHOW" Then
            pShowImage = True
        Else
            pShowImage = False
        End If
    End If
        
    If m_ValidEnableSaveImage = True Then
        pSaveImage = m_FlagEnableSaveImage
    Else
        If pAfiArg7 = "WRITE" Then
            pSaveImage = True
        Else
            pSaveImage = False
        End If
    End If
    
    '### Interceptor ###
    If m_FlagEnableInterceptor = True Then
        pInterceptor = True
    Else
        pInterceptor = False
    End If

End Sub

Public Sub EnableOffImageDebugger()
    m_ValidEnableReadImage = False
    m_ValidEnableSaveImage = False
    m_ValidEnableShowImage = False
    Call ChangeConstructionObj
End Sub

Public Sub EnableReadImage(ByVal pEnableFlag As Boolean, ByVal pPathName As String)
    m_ValidEnableReadImage = True
    m_FlagEnableReadImage = pEnableFlag
    Call ChangeConstructionObj
End Sub

Public Sub EnableSaveImage(ByVal pEnableFlag As Boolean, ByVal pPathName As String)
    m_ValidEnableSaveImage = True
    m_FlagEnableSaveImage = pEnableFlag
    Call ChangeConstructionObj
End Sub

Public Sub EnableShowImage(ByVal pEnableFlag As Boolean)
    m_ValidEnableShowImage = True
    m_FlagEnableShowImage = pEnableFlag
    Call ChangeConstructionObj
End Sub

Public Sub EnableInterceptor(ByVal pEnableFlag As Boolean, ByRef pLogger As CActionLogger)
    m_FlagEnableInterceptor = pEnableFlag
    Set m_ActionLogger = pLogger
    Call ChangeConstructionObj
End Sub

Private Function SearchSubstance(ByRef pObj As IAcquireAction) As IAcquireAction
'Set pSetObj = pIAAが失敗すると実体なので、pIAAを返す。

    Dim pSetObj As IAcquireActionInstance       'インスタンスのセッター、ゲッター用
    Dim pIAA As IAcquireAction
    
    Set pIAA = pObj                             'とりあえず、ローカルのオブジェクト変数に渡す
    
    On Error GoTo NON_WRAP
    
    Do
        '持っているIAcquireActionインスタンスが無くなるまで繰り返しSet。
        Set pSetObj = pIAA                              'このSetで失敗したらpIAAが実体
        Set pIAA = pSetObj.AcquireActionInstance        'このSetは成功する
    Loop
    
NON_WRAP:
    Set SearchSubstance = pIAA

End Function

Private Sub Class_Initialize()
    Set mAcquireInstanceAction = Nothing
End Sub

Private Sub Class_Terminate()
    Set mAcquireInstanceAction = Nothing
End Sub
