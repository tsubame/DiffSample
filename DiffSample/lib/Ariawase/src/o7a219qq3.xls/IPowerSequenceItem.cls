VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPowerSequenceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   IPowerSequenceItemクラス
'
'目的:
'   PowerSequenceのアイテムを抽象化する
'
'作成者:
'   2011/12/05 Ver0.1 D.Maruyama
'   2012/11/02 Ver0.2 H.Arikawa
'                       ・ExcecuteForUS追加。
'
Option Explicit

Public Sub Excecute(ByVal strPowerConditionName As String)

End Sub
Public Sub ExcecuteForUS(ByVal strPowerConditionName As String)

End Sub

Public Function SetParam(ByVal Param As Variant) As Boolean


End Function
