VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPowerSquenceWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   CPowerSquenceWait�N���X
'
'�ړI:
'   IPowerSequenceItem���������
'   Wait���s���APowerSequenceItem�B
'
'�쐬��:
'   2011/12/05 Ver0.1 D.Maruyama
'   2012/11/02 Ver0.2 H.Arikawa
'                       �EIPowerSequenceItem_WaitSet�ǉ��BExcecute�̒��g��Private�֐���
'                       �EIPowerSequenceItem_ExcecuteForUS�ǉ��B
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
' �wPower_Off_V�x�Ή��ׁ̈A�ǉ��B
Public Sub IPowerSequenceItem_ExcecuteForUS(ByVal strPowerConditionName As String)

    Call IPowerSequenceItem_WaitSet
End Sub
Public Function IPowerSequenceItem_SetParam(ByVal Param As Variant) As Boolean

    If Not IsNumeric(Param) Then
        mdblWait = Param
    End If
    
    mdblWait = Param
    
End Function
' �֐������قȂ邪�������e�����s����ׁA���g���֐������A���p�B
Private Sub IPowerSequenceItem_WaitSet()
    
#If OFFLINE_DEBUG = 1 Then
    Debug.Print "Wait " & mdblWait & " " & strPowerConditionName
#Else
    TheHdw.WAIT mdblWait * mS
#End If

End Sub
