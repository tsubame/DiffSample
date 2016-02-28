VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMeasure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   電流・電圧測定クラスが実装するインターフェイス
'
'目的:
'   電流・電圧測定クラスが持つパラメータの初期化を行う
'   測定結果をResultクラスオブジェクトへ書き出す
'
'作成者:
'   0145206097
'
Option Explicit

Public Sub Initialize(ByVal measSite As Long, ByVal MeasurePins As Collection, ByVal setContainer As CContainer, ByVal reader As IParameterReader)
End Sub

Public Sub WriteResult(ByVal resultKeeper As CResultManager)
End Sub

Public Sub GetBoardInfo()
End Sub

Public Property Get MeasurePins() As Collection
End Property
