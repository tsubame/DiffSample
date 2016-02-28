Attribute VB_Name = "XEeeAuto_TestInstanceInterface"
Option Explicit

'概要:
'   TestInstanceからのArg取得に関する関数群
'
'目的:
'
'
'作成者:
'   2011/12/21 Ver0.1 D.Maruyama
'   2012/04/09 Ver0.2 D.Maruyama　　dcsetup関数の引数変更に伴い、「EEE_AUTO_DCSETUP_PARAM」の定義を修正
'   2012/04/09 Ver0.2 D.Maruyama　　dcsetup関数の引数変更に伴い、「EEE_AUTO_DCSETUP_PARAM」の定義を修正


Private Const EEE_AUTO_TEST_INSTANCE_ARG_START As Long = 20

Public Const EEE_AUTO_VARIABLE_PARAM As Long = -1

Public Const EEE_AUTO_DCSETUP_PARAM As Long = 3
Public Const EEE_AUTO_ENDOFTEST_PARAM As Long = 1

'内容:
'   TestInstanceからのArgの読み取りをラップする。
'   中途半端なところをArgの開始位置としたため、ラップ関数で可読性をよくする。
'
'
'パラメータ:
'[arystrParam]         Out  結果配列
'[lNumOfParam]         In　 Argの数
'
'注意事項:
'   arystrParamは確保してない動的配列を渡すこと
'
Public Function EeeAutoGetArgument(ByRef arystrParam() As String, ByVal lNumOfParam As Long) As Boolean

    EeeAutoGetArgument = False

    Dim ArgArr() As String
    Dim Argnum As Long
    
    Call TheExec.DataManager.GetArgumentList(ArgArr, Argnum)

    '期待するパラメータが異なる場合はエラーとする。可変引数の場合は無視
    If lNumOfParam <> (Argnum - EEE_AUTO_TEST_INSTANCE_ARG_START) And _
            lNumOfParam <> EEE_AUTO_VARIABLE_PARAM Then
        EeeAutoGetArgument = False
        Exit Function
    End If
    
    '必要な数だけ引数の配列を確保する
    Dim lUsedNum As Long
    lUsedNum = Argnum - EEE_AUTO_TEST_INSTANCE_ARG_START
    ReDim arystrParam(lUsedNum)
    
    'コピーする
    Dim i As Long
    For i = 0 To lUsedNum - 1
        arystrParam(i) = ArgArr(EEE_AUTO_TEST_INSTANCE_ARG_START + i)
    Next i
    
    EeeAutoGetArgument = True
    
End Function


