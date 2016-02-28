VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFixedLsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   デジタルコード換算用LSB値のデータを取得＆保持するクラス
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfigurationsシートのFixedLSB情報を保持する機能を実装した
'
'作成者:
'   0145184346
'

Option Explicit

Implements ILSBSnapper ' LSB取得用I/Fを宣言する

Private Const ERR_NUMBER As Long = 9999             ' Error番号を保持する
Private Const CLASS_NAME As String = "CFixedLsb"    ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty ' Default値"Empty"を保持する

Private mLSBValue As String ' LSB値を保持する

Private Property Get ILSBSnapper_LSB() As CParamDoubleArray

    '#####  LSB値やLSB設定変数名称からデータを取得する  #####
    If mLSBValue <> "" Then
        Dim clsPrmDblAry As CParamDoubleArray ' パラメータ情報を保持する
        Set clsPrmDblAry = New CParamDoubleArray
        If IsNumeric(mLSBValue) = True Then
            clsPrmDblAry.SetAsDouble CDbl(mLSBValue) ' 数値の場合
        Else
            clsPrmDblAry.SetArrayAsDouble XLibTheDeviceProfilerUtility.GetLSBParam(mLSBValue) ' 変数名称の場合
        End If
    Else
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".LSB", "Nothing LSB Value.")
    End If

    '#####  終了  #####
    Set ILSBSnapper_LSB = clsPrmDblAry
    Set clsPrmDblAry = Nothing

End Property

Public Sub Initialize(ByRef strLSBData As String)
    mLSBValue = strLSBData
End Sub

Private Sub Class_Initialize()
    mLSBValue = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    mLSBValue = INITIAL_EMPTY_VALUE
End Sub
