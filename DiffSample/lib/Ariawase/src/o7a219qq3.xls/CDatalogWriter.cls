VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDatalogWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   入出力ファイル情報ダンプクラス
'
'目的:
'   画像のファイル読込/ファイル出力時に
'   ファイル名をDatalogに出力する
'
'作成者:
'   0145184934
'
Option Explicit

Implements IProcParamWriter

Private mMarker As String

Public Sub IProcParamWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
End Sub

Public Sub IProcParamWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
End Sub

Public Sub IProcParamWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
End Sub

Public Sub IProcParamWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
End Sub

Public Sub IProcParamWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
'内容:
'   テキスト情報をDatalogにダンプする。
'
'[dataName]    IN String型:     データ名
'[strData]     IN String型:     データ
'
'備考:
'
    If dataName <> vbNullString Then
        TheExec.Datalog.WriteComment mMarker & dataName & ":" & strData
    Else
        TheExec.Datalog.WriteComment mMarker & strData
    End If
End Sub

Public Sub IProcParamWriter_WriteAsBlockEnd(ByVal pParamName As String)
'内容:
'   マーカーを初期化する。
'
'[pParamName]    IN String型:     文字列
'
'備考:
'
    mMarker = ""
End Sub

Public Sub IProcParamWriter_WriteAsBlockStart(ByVal pParamName As String)
'内容:
'   マーカーを設定する。
'
'[pParamName]    IN String型:     文字列
'
'備考:
'
    mMarker = "#"
End Sub

Public Sub IProcParamWriter_ReturnToRootBlock()
End Sub

Public Property Get IProcParamWritermyName() As String
End Property

Public Sub IProcParamWriter_WriteAsPlane(ByVal pParamName As String, ByRef pImgPlane As CImgPlane)
End Sub

Private Property Get IProcParamWriter_MyName() As String
End Property

Public Function AsIProcParamWriter() As IProcParamWriter
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IProcParamWriter型オブジェクト
'
'備考:
'
    Set AsIProcParamWriter = Me
End Function
