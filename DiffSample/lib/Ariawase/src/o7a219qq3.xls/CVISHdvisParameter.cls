VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVISHdvisParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'HDVIS���W���[���̃p�����[�^�񋟗p �@�\�Q�N���X
'�n�[�h�E�F�A�̕s��C�����s����܂ł̎b��΍��p

'Tool�Ή���ɃR�����g�O���Ď��������ɂ���B�@2013/03/07 H.Arikawa
#Const HDVIS_USE = 0      'HVDIS���\�[�X�̎g�p�@ 0�F���g�p�A0�ȊO�F�g�p  <PALS��EeeAuto�����Ŏg�p>

#If HDVIS_USE <> 0 Then
Private Const CS_NAME = "CVISHdvisParameter"
Private Const PARAM_WSHT_NAME = "HDVIS_Param" 'HDVIS Measure�p�����[�^���[�N�V�[�g��(�����V�[�g��؂�ւ��邱�Ƃ͍l�����Ă��Ȃ�)
Private m_ParameterSht As Object
'Private m_ParameterSht As ShtHdvisParam '�R�[�f�B���O���͓��͎x�����Ȃ��Ƃ炢�̂ł�������g�p����

'#Pass-Ok
Private Sub Class_Initialize()
    Call Me.Initialize
End Sub

'#Pass-Ok
Public Sub Initialize()
'�������iParameter���[�N�V�[�g�̐ݒ�A����Parameter Cache�̃t���b�V���AParameter Cache�̍쐬�j

    On Error GoTo PARAM_WSHT_ERROR
    Set m_ParameterSht = Worksheets(PARAM_WSHT_NAME) 'HDVIS�p�����[�^�V�[�g�̎w��
    On Error GoTo 0
    
    With m_ParameterSht
        .Initialize
        .LoadParameter
    End With

    Exit Sub

PARAM_WSHT_ERROR:
    Set m_ParameterSht = Nothing
    Call OutputErrMsg(PARAM_WSHT_NAME & " worksheet is nothing! ")
    
End Sub

'#Pass-Ok
Public Sub LoadParameter()
'���[�N�V�[�g�̃p�����[�^�̓ǂݍ���
'(����Parameter Cache�̃t���b�V���AParameter Cache�̍쐬)
    
    If m_ParameterSht Is Nothing Then
        Call Me.Initialize
    End If
    
    m_ParameterSht.LoadParameter

End Sub


'------------------------------------------------------------------------------
'�p�����[�^���J�p�@�\�Q
'------------------------------------------------------------------------------
'#Pass-Ok
Public Property Get ParameterSheetName() As String
'HDVIS�p�����[�^���[�N�V�[�g���̌��J
    
    If m_ParameterSht Is Nothing Then
        Call Me.Initialize
    End If
    
    ParameterSheetName = m_ParameterSht.Name

End Property

'#Pass-Ok
Public Property Get PinCapacitor(ByVal PinName As String) As Double
'�w��s���̗e�ʒl�����J
    
    If m_ParameterSht Is Nothing Then
        Call Me.Initialize
    End If
    
    PinCapacitor = m_ParameterSht.PinCapacitor(PinName)

End Property

'#Pass-Ok
Public Property Get Samples(ByVal PinName As String, ByVal Irange As HdvisIRange) As Long
'�w��s���A�w�背���W��Samples�l�����J
            
    If m_ParameterSht Is Nothing Then
        Call Me.Initialize
    End If
    
    Dim tmpParameter() As Double
    
    With m_ParameterSht
        Select Case Irange
            Case hdvis5uA
                tmpParameter = .Hdvis5uAParameter(PinName)
                Samples = CLng(tmpParameter(0))
            Case hdvis50uA
                tmpParameter = .Hdvis50uAParameter(PinName)
                Samples = CLng(tmpParameter(0))
            Case Else
                Call OutputErrMsg(PinName & ", IRange=" & Irange & " Samples parameter is not found" & " at " & CS_NAME)
        End Select
    End With

End Property

'#Pass-Ok
Public Property Get Period(ByVal PinName As String, ByVal Irange As HdvisIRange) As Double
'�w��s���A�w�背���W��Period�l�����J
        
    If m_ParameterSht Is Nothing Then
        Call Me.Initialize
    End If
    
    Dim tmpParameter() As Double
    
    With m_ParameterSht
        Select Case Irange
            Case hdvis5uA
                tmpParameter = .Hdvis5uAParameter(PinName)
                Period = tmpParameter(1)
            Case hdvis50uA
                tmpParameter = .Hdvis50uAParameter(PinName)
                Period = tmpParameter(1)
            Case Else
                Call OutputErrMsg(PinName & ", IRange=" & Irange & " Period parameter is not found" & " at " & CS_NAME)
        End Select
    End With

End Property

'#Pass-Ok
Public Property Get SlewRate() As HdvisSlewRate
'���[�N�V�[�g��SlewRate�ݒ��HdvisSlewRate�^�ɕϊ����Č��J
    
    If m_ParameterSht Is Nothing Then
        Call Me.Initialize
    End If
    
    Dim tmpParam As String
    
    tmpParam = ""
    tmpParam = m_ParameterSht.SlewRate
    
    Select Case LCase(tmpParam)
        Case "slow"
            SlewRate = hdvisSlewRateSlow
        Case "normal"
            SlewRate = hdvisSlewRateNormal
        Case Else
            Call OutputErrMsg(tmpParam & " Slew Rate parameter is unknown" & " at " & CS_NAME) '�L���b�V���쐬���`�F�b�N������̂Œʏ�R�R�̃G���[�͏o�Ȃ�
    End Select

End Property


'------------------------------------------------------------------------------
'���W���[���̃p�����[�^�`�F�b�N�@�\�Q
'------------------------------------------------------------------------------
'#Pass-Ok
Public Function ChkSameIrange(ByRef HdvisIRange() As HdvisIRange) As Boolean
'�z��Ɋi�[����Ă��鑪�背���W�l�����ׂē���l�ł��邱�Ƃ��m�F����
    
    Dim chCnt As Long

    For chCnt = 1 To UBound(HdvisIRange) Step 1
        If HdvisIRange(0) <> HdvisIRange(chCnt) Then 'HdvisIrange�^(Long)�̒l��r
            ChkSameIrange = False
            Exit Function
        End If
    Next chCnt

    ChkSameIrange = True
    
End Function

'#Pass-Ok
Public Function ChkSameCapacitor(ByRef PinList As String) As Boolean
'�s�����X�g�̃s���e�ʂ̒�`�l�����ׂē���ł��邱�Ƃ��m�F
    
    Dim pinNameArr() As String
    Dim PinCnt As Long
    Dim refCapacitor As Double
        
    '�J���}��؂�̃s�����X�g��z��`���ɕϊ�
    Call SeparatePinList(PinList, pinNameArr) '�����PinMap�ɒ�`�̖���Pin��n���Ǝ��s���G���[�ƂȂ�

    '�s�����X�g�̐擪�s���̗e�ʒl����Ƃ���
    With Me
        refCapacitor = .PinCapacitor(pinNameArr(0))
        For PinCnt = 1 To UBound(pinNameArr) Step 1 '�v�f0��Ref�Ȃ̂ŗv�f1����
            If CompareDblData(refCapacitor, .PinCapacitor(pinNameArr(PinCnt)), 15) = False Then '�ŏ��P��fF�Ő؎̂Ă������e�ʒl���𔻒�
                ChkSameCapacitor = False
                Exit Function
            End If
        Next PinCnt
    End With

    ChkSameCapacitor = True
    
End Function

'#Pass-Ok
Public Function ChkIRangeIncluded(ByVal targetIrange As HdvisIRange, ByRef chkIrange() As HdvisIRange) As Boolean
'�z��Ɋi�[����Ă��鑪�背���W�l�ɁA�w�背���W�̒l���ЂƂł��܂܂�Ă��邩���m�F
    
    Dim chCnt As Long

    For chCnt = 0 To UBound(chkIrange) Step 1
        If targetIrange = chkIrange(chCnt) Then
            ChkIRangeIncluded = True
            Exit Function
        End If
    Next chCnt

    ChkIRangeIncluded = False
    
End Function

'#Pass-Ok
Public Function ChkSameMeasureParameter(ByVal PinList As String, Irange As HdvisIRange) As Boolean
'�s�����X�g�Ɏw�肳�ꂽ�s���̃��W���[�p�����[�^(Samples�APeriod)�����ׂē���ł��邩���m�F
    
    Dim pinNameArr() As String
    Dim PinCnt As Long
    Dim refSamples As Double
    Dim refPeriod As Double
        
    '�J���}��؂�̃s�����X�g��z��`���ɕϊ�
    Call SeparatePinList(PinList, pinNameArr) '�����PinMap�ɒ�`�̖���Pin��n���Ǝ��s���G���[�ƂȂ�

    '�s�����X�g�̐擪�s����Samples�APeriod����i��r���j�Ƃ���
    With Me
        refSamples = .Samples(pinNameArr(0), Irange)
        refPeriod = .Period(pinNameArr(0), Irange)
        For PinCnt = 1 To UBound(pinNameArr) Step 1 '�v�f0��Ref�Ȃ̂ŗv�f1����
            'Samples�̒l��r
            If refSamples <> .Samples(pinNameArr(PinCnt), Irange) Then
                ChkSameMeasureParameter = False
                Exit Function
            End If
            'Period�̒l��r
            If CompareDblData(refPeriod, .Period(pinNameArr(PinCnt), Irange), 15) = False Then '�ŏ��P��fSec�Ő؎̂Ă��������Ԃ��𔻒�
                ChkSameMeasureParameter = False
                Exit Function
            End If
        Next PinCnt
    End With

    ChkSameMeasureParameter = True

End Function

Public Sub DispMeasureCondition(ByVal PinList As String)
'���W���[���Ƀp�����[�^�����f����Ă��邩�m�F���邽�߂̎����@�\

    Dim tmpMode() As HdvisMode
    Dim tmpVrange() As HdvisVRange
    Dim tmpIrange() As HdvisIRange
    Dim tmpSlewRate() As HdvisSlewRate

    Dim MeaIrange As HdvisIRange
    Dim MeaPriod As Double
    Dim MeaPinNames() As String
    Dim MeaSlewRate As HdvisSlewRate
    Dim statMsg As String
    
    '�e���_�C��API���{�[�h�̐ݒ�p�����[�^���擾
    With TheHdw.HDVIS.Pins(PinList)
        Call .ReadRangesAndMode(tmpMode, tmpVrange, tmpIrange)
        Call .ReadSlewRates(tmpSlewRate)
    End With
    
    Call SeparatePinList(PinList, MeaPinNames)
    
    MeaPriod = TheHdw.HDVIS.Period
    MeaSlewRate = tmpSlewRate(0)

    statMsg = ("Pin=" & PinList & vbCrLf & _
            "Irange=" & tmpIrange(0) & vbCrLf & _
            "Period=" & MeaPriod & vbCrLf & _
            "SlewRate=" & MeaSlewRate & vbCrLf & _
            "Ave=" & Me.Samples(MeaPinNames(0), tmpIrange(0)))

    Call MsgBox(statMsg, vbInformation, "HDVIS Mesure Infomation")

End Sub


#End If
