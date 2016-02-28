VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDebugPointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   任意のUIにブレイクポイント設定UIとしての側面を持たせるためのインターフェース
'   利用する側はブレイクポイント設定がどこにされるかを気にしなくて良い
'
'作成者:
'   0145206097
'
Option Explicit

Public Sub SetBreak()

End Sub

Public Sub ClearAllDebugPoint()

End Sub

Public Property Let IsEnable(ByVal RHS As Boolean)

End Property

Public Property Get IsEnable() As Boolean

End Property

