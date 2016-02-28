VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDynamicLsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   デジタルコード換算用LSB値のデータを取得＆保持するクラス
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfigurationsシートのDynamicLSB情報を保持する機能を実装した
'
'作成者:
'   0145184346
'

Option Explicit

Implements ILSBSnapper ' LSB取得用I/Fを宣言する

Private Const ERR_NUMBER As Long = 9999             ' Error番号を保持する
Private Const CLASS_NAME As String = "CDynamicLsb"  ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty ' Default値"Empty"を保持する

Private mMacroName As String ' LSB値取得用のMacro名称を保持する

Private Property Get ILSBSnapper_LSB() As CParamDoubleArray

    '#####  LSB取得用のUserMacroを実行して、LSB値を取得する  #####
    If mMacroName <> "" Then

        '#####  UserMacroを実行＆データを格納する  #####
        If IsAlphabet(mMacroName) = True Then

            '#####  UserMacroを実行する  #####
            Dim dblArray() As Double ' Macroからの戻り値を保持する
            On Error GoTo ErrHandler
            dblArray = Application.Run(mMacroName)
            On Error GoTo 0
    
            '#####  データを格納する  #####
            Dim clsPrmDblAry As CParamDoubleArray ' パラメータ情報を保持する
            Set clsPrmDblAry = New CParamDoubleArray
            clsPrmDblAry.SetArrayAsDouble dblArray
            Set ILSBSnapper_LSB = clsPrmDblAry
            Set clsPrmDblAry = Nothing

        End If

    Else
        Set ILSBSnapper_LSB = Nothing
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".LSB", "Nothing Macro Name.")
    End If

    '#####  終了  #####
    Exit Property

ErrHandler:
    Set ILSBSnapper_LSB = Nothing
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".LSB", "Mismatch Macro Name.[" & mMacroName & "]")
    Exit Property
End Property

Public Sub Initialize(ByRef strMacroName As String)
    mMacroName = strMacroName
End Sub

Private Sub Class_Initialize()
    mMacroName = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    mMacroName = INITIAL_EMPTY_VALUE
End Sub
