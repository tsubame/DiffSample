VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAcquireActionInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   クラスが持つインスタンスを返します。
'目的:
'   AcquireActionクラスが持つインスタンスのSetter、Getter。
'   IAcquireActionをI/Fに持つクラス専用で使ってください。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月10日

Option Explicit

Public Property Set AcquireActionInstance(pInstance As IAcquireAction)
End Property

Public Property Get AcquireActionInstance() As IAcquireAction
End Property


