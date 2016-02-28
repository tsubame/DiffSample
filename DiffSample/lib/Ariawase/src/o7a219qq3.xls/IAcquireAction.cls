VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAcquireAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   AcquireFrameWorkのインターフェースとなるクラス
'
'   Revision History:
'       Data        Description
'       2010/06/11  インターフェース機能を実装した
'       2010/06/22  実行情報リセット機能を実装した
'
'作成者:
'   0145184346
'

Option Explicit

Private Const CLASS_NAME = "IAcquireAction"

Public Function Execute() As Integer
End Function

Public Function ToptStart() As Integer
End Function

Public Function ClearStatus() As Integer
End Function

Public Property Get GetFrameWorkName() As String
End Property

Public Property Get GetStatus() As Integer
End Property

Public Sub Reset()
End Sub

Public Sub Initialize(reader As IParameterReader)
End Sub

Public Property Get AcquireFunctionInfo() As CAcquireFunctionInfo
End Property


