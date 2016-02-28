VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVarBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   CVarBankクラスのインターフェイス
'目的:
'   割り込みの有無で、インスタンスの型を切り替える
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月22日

Option Explicit

Public Property Get Value(VariableName As String) As Variant
End Property

Public Property Let Value(VariableName As String, pValue As Variant)
End Property

Public Property Get ReadOnly(VariableName As String) As Boolean
End Property

Public Property Let ReadOnly(VariableName As String, pReadOnly As Boolean)
End Property

Public Property Get Fixed(VariableName As String) As Boolean
End Property

Public Property Let Fixed(VariableName As String, pFixed As Boolean)
End Property

Public Property Get Comment(VariableName As String) As String
End Property

Public Property Let Comment(VariableName As String, pComment As String)
End Property

Public Property Get Count() As Long
End Property

Public Function Add(VariableName As String, FirstValue As Variant, Optional FixedStatus As Boolean, Optional Comment As String) As Boolean
End Function

Public Function IsExist(VariableName As String) As Boolean
End Function

Public Function Search(ByVal SearchPattern As String, Optional NotFlag As Boolean = False) As String
End Function

Public Sub Clear()
End Sub

Public Sub AllClear()
End Sub

