Attribute VB_Name = "XLibTheVarBankUtility"
'概要:
'   TheVarBankのユーティリティ
'
'目的:
'   TheVarBank:CVarBankの初期化/破棄のユーテリティを定義する
'
'作成者:
'   a_oshima

Option Explicit

Public TheVarBank As IVarBank

Private mSaveMode As Boolean
Private mSaveFileName As String
Private mTheVarBank As CVarBank
Private mTheVarBankInterceptor As CVarBankInterceptor


Public Sub CreateTheVarBankIfNothing()
'内容:
'   As Newの代替として初回に呼ばせるインスタンス生成処理
'
'パラメータ:
'   なし
'
'注意事項:
'
    
    On Error GoTo ErrHandler
    If TheVarBank Is Nothing Then
        Set mTheVarBank = New CVarBank
        Set TheVarBank = mTheVarBank
        mSaveMode = False
    End If
    Exit Sub
ErrHandler:
    Set TheVarBank = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub SaveModeTheVarBank(ByVal pEnableLoggingTheVar As Boolean, Optional saveFileName As String)
    mSaveFileName = saveFileName
    If mSaveMode <> pEnableLoggingTheVar Then
        If pEnableLoggingTheVar = True Then
            '割り込みいれる
            Set mTheVarBankInterceptor = New CVarBankInterceptor
            Call mTheVarBankInterceptor.Initialize(mTheVarBank)
            Set TheVarBank = mTheVarBankInterceptor.AsIVarBank
            mSaveMode = True
            TheExec.Datalog.WriteComment "Eee JOB Output Log! :TheCondition"
        Else
            '割り込み外す
            Call mTheVarBankInterceptor.Initialize(Nothing)
            Set mTheVarBankInterceptor = Nothing
            Set TheVarBank = mTheVarBank
            mSaveMode = False
        End If
    End If
End Sub

Public Sub DestroyTheVarBank()
    Set TheVarBank = Nothing
    Set mTheVarBank = Nothing
    Set mTheVarBankInterceptor = Nothing
End Sub

Public Function RunAtJobEnd() As Long
    
    If Not (TheVarBank Is Nothing) Then
        TheVarBank.Clear
        If Not (mTheVarBankInterceptor Is Nothing) Then
            Call mTheVarBankInterceptor.SaveLogFile(mSaveFileName)
        End If
        Call SaveModeTheVarBank(False)            'テスト終了でFalse
    End If

End Function
