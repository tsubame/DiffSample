VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPowerSquenceWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   CPowerSquenceWaitクラス
'
'目的:
'   IPowerSequenceItemを具現化する
'   Waitを行う、PowerSequenceItem。
'
'作成者:
'   2011/12/05 Ver0.1 D.Maruyama
'   2012/11/02 Ver0.2 H.Arikawa
'                       ・IPowerSequenceItem_WaitSet追加。Excecuteの中身をPrivate関数化
'                       ・IPowerSequenceItem_ExcecuteForUS追加。
'
Option Explicit

Implements IPowerSequenceItem

Public mdblWait As Double

Private Sub Class_Initialize()
    mdblWait = 0
End Sub

Public Sub IPowerSequenceItem_Excecute(ByVal strPowerConditionName As String)

    Call IPowerSequenceItem_WaitSet
End Sub
' 『Power_Off_V』対応の為、追加。
Public Sub IPowerSequenceItem_ExcecuteForUS(ByVal strPowerConditionName As String)

    Call IPowerSequenceItem_WaitSet
End Sub
Public Function IPowerSequenceItem_SetParam(ByVal Param As Variant) As Boolean

    If Not IsNumeric(Param) Then
        mdblWait = Param
    End If
    
    mdblWait = Param
    
End Function
' 関数名が異なるが同じ内容を実行する為、中身を関数化し、共用。
Private Sub IPowerSequenceItem_WaitSet()
    
#If OFFLINE_DEBUG = 1 Then
    Debug.Print "Wait " & mdblWait & " " & strPowerConditionName
#Else
    TheHdw.WAIT mdblWait * mS
#End If

End Sub

