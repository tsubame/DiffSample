VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIcul1gChanInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@@@IP750EX Support
'�T�v:
'   ICUL1G�̃`�����l�����Ǘ��p�N���X
'
'�ړI:
'   �֘A�����Z�b�g�ŊǗ�����
'

Option Explicit

'Tool�Ή���ɃR�����g�O���Ď��������ɂ���B�@2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1               '1CUL1G�{�[�h�̎g�p�@  0�F���g�p�A0�ȊO�F�g�p  <Tester��IP750EX�Ȃ�Default:1�ɂ��Ă���>

#If ICUL1G_USE <> 0 Then

Private m_ChanNumber As Long
Private m_ChanType As chtype
Private m_Icul1gChanType As Icul1gChannelType
Private m_PinName As String
Private m_SiteNumber As Long

'�R���X�g���N�^
Private Sub Class_Initialize()
    m_ChanNumber = -1
    m_ChanType = chUnk
    m_Icul1gChanType = icul1gChannelTypeUnknown
    m_PinName = Empty
    m_SiteNumber = -1
End Sub

'�s����
Public Property Get PinName() As String
    PinName = m_PinName
End Property
Public Property Let PinName(ByVal pName As String)
    '�{���͎w�肳�ꂽ�s����ICUL1G-PPMU�s�����𒲂ׂ�`�F�b�N������ق����悢
    m_PinName = pName
End Property

'�T�C�g�ԍ�
Public Property Get SiteNumber() As Long
    SiteNumber = m_SiteNumber
End Property
Public Property Let SiteNumber(ByVal pValue As Long)
    m_SiteNumber = pValue
End Property

'�`�����l���ԍ�(icul1gXX�̐���)
Public Property Get ChanNumber() As Long
    ChanNumber = m_ChanNumber
End Property
Public Property Let ChanNumber(ByVal pNumber As Long)
    m_ChanNumber = pNumber
End Property

'�`�����l���^�C�v
Public Property Get chanType() As chtype
    chanType = m_ChanType
End Property
Public Property Let chanType(ByVal pType As chtype)
    m_ChanType = pType
    m_Icul1gChanType = ChType2Icul1gChannelType(pType)
End Property

'ICUL1G�̃`�����l���^�C�v(�`�����l���^�C�v�l���玩���ݒ�Ȃ̂�Set�͂Ȃ�)
Public Property Get Icul1gChanType() As Icul1gChannelType
    Icul1gChanType = m_Icul1gChanType
End Property


'--- �ȉ�Private ---

'�`�����l���^�C�v��ICUL1G�`�����l���^�C�v�ɕϊ�����֐�
Private Function ChType2Icul1gChannelType(ByVal pType As chtype) As Icul1gChannelType
    Select Case pType
        Case chICUL1Gclk_high
            ChType2Icul1gChannelType = icul1gChannelTypeClkHigh
        Case chICUL1Gclk_low
            ChType2Icul1gChannelType = icul1gChannelTypeClkLow
        Case chICUL1Gdata_high
            ChType2Icul1gChannelType = icul1gChannelTypeDataHigh
        Case chICUL1Gdata_low
            ChType2Icul1gChannelType = icul1gChannelTypeDataLow
        Case Else
            Call Err.Raise(9999, TypeName(Me), "icul1gChannelTypeUnknown")
'            ChType2Icul1gChannelType = icul1gChannelTypeUnknown
    End Select
End Function

#End If