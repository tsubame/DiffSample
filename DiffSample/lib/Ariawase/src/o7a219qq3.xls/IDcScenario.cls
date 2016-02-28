VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDcScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Apichoice()
End Sub

Public Sub ClearContainer()
End Sub

Public Sub Execute()
End Sub

Public Sub GetTempResult(ByVal pName As String, ByRef pResult() As Double)
End Sub

Public Sub Initialize(ByVal dcPerformer As IDcTest, ByVal scenarioReader As IParameterReader, ByVal instanceReader As IParameterReader, ByVal scenarioWriter As IParameterWriter, ByVal pinInfoWriter As IParameterWriter)
End Sub

Public Sub Load()
End Sub

Public Property Let ResultManager(ByVal resultKeeper As CResultManager)
End Property

Public Sub SetScenario(ByVal currTest As String)
End Sub
