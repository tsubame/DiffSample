VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPowerSequencePin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   CPowerSquenceWait�N���X
'
'�ړI:
'   IPowerSequenceItem���������
'   �s���ւ̃t�H�[�X���s���APowerSequenceItem�B
'
'�쐬��:
'   2011/12/05 Ver0.1 D.Maruyama
'   2012/11/02 Ver0.2 H.Arikawa
'                   �EIPowerSequenceItem_ExcecuteForUS�ǉ��B(APMU UnderShoot�΍�)
'                   �ECP_GetChanType�ǉ��B
'   2013/02/12 Ver0.3 H.Arikawa
'                   �EIPowerSequenceItem_ExcecuteForUS�C���B(ZERO_V,ZERO�Ή�)
'   2013/03/04 Ver0.4 H.Arikawa
'                   �EIPowerSequenceItem_ExcecuteForUS�C���B(Exit sub�ǉ�)

Option Explicit

Implements IPowerSequenceItem

Public mstrPins As String

Private Sub Class_Initialize()
    mstrPins = ""
End Sub

Public Sub IPowerSequenceItem_Excecute(ByVal strPowerConditionName As String)

    Dim chanType As chtype
    Dim ClampI As Double
    Dim MIRange As ApmuIRange
    Dim ForceV As Variant

#If OFFLINE_DEBUG = 1 Then
    Debug.Print "Force " & mstrPins & " " & strPowerConditionName
#Else
    Call ShtPowerV.GetPowerInfo(strPowerConditionName, mstrPins).Force(mstrPins)
#End If

End Sub

Public Sub IPowerSequenceItem_ExcecuteForUS(ByVal strPowerConditionName As String)
' 11/2 H.Arikawa For APMU UnderShoot.

    Dim chanType As chtype
    Dim ClampI As Double
    Dim MIRange As ApmuIRange
    Dim ForceV As Variant

#If OFFLINE_DEBUG = 1 Then
    Debug.Print "Force " & mstrPins & " " & strPowerConditionName
#Else
    chanType = CP_GetChanType(mstrPins)
    If chanType = chAPMU Then
        '### For Power Down ###
            ForceV = 0 * V
            MIRange = apmu50mA
            ClampI = 5 * mA
        '######################
        Call SetFVMI_APMU(mstrPins, ForceV, ClampI, , , MIRange)
    Else
        On Error GoTo ERROR_DETECTION1
        Call ShtPowerV.GetPowerInfo(strPowerConditionName, mstrPins).Force(mstrPins)
    End If
    Exit Sub
ERROR_DETECTION1:
        strPowerConditionName = strPowerConditionName & "_V"
        Call ShtPowerV.GetPowerInfo(strPowerConditionName, mstrPins).Force(mstrPins)
#End If

End Sub

Public Function IPowerSequenceItem_SetParam(ByVal Param As Variant) As Boolean

    If StrComp(TypeName(Param), "String", vbTextCompare) <> 0 Then
        IPowerSequenceItem_SetParam = False
    End If
    
    mstrPins = Param

End Function
Public Function CP_GetChanType(ByVal PinList As String) As chtype
'[GetChanType]��Copy�BEee-Job��Private�֐��ׁ̈ACopy���쐬�B
'Eee-Job�̊֐��ύX���������^�C�~���O�ō폜���ē���ւ���B

    Dim chanType As chtype
    Dim pinArr() As String
    Dim pinNum As Long
    Dim i As Long

    With TheExec.DataManager
        chanType = .chanType(PinList)

        If chanType = chUnk Then
            On Error GoTo INVALID_PINLIST
            Call .DecomposePinList(PinList, pinArr, pinNum)
            On Error GoTo -1

            chanType = .chanType(pinArr(0))
            For i = 0 To pinNum - 1
                If chanType <> .chanType(pinArr(i)) Then
                    chanType = chUnk
                    Exit For
                End If
            Next i
        End If
    End With

    CP_GetChanType = chanType
    Exit Function

INVALID_PINLIST:
    CP_GetChanType = chUnk

End Function