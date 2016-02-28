VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIcul1gPpmuInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@@@IP750EX Support
'概要:
'   ICUL1GのPPMU情報格納用クラス
'
'目的:
'   PPMUの設定パラメータをセットで管理する
'

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1               '1CUL1Gボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

#If ICUL1G_USE <> 0 Then

Private m_PinName As String
Private m_ChannelNumber As Long
Private m_ChannelType As chtype
Private m_Icul1gChannelType As Icul1gChannelType
Private m_SiteNumber As Long
Private m_ForceVoltage As Double
Private m_ForceCurrent As Double
Private m_CurrentRange As Icul1gPpmuCurrentRange
Private m_TestLimitHigh As Double
Private m_TestLimitLow As Double
Private m_ClampVHi As Double
Private m_ClampVLo As Double
Private m_RelayState As Icul1gRelayState
Private m_FIMode As Boolean
Private m_FVMode As Boolean
Private m_TestLimitValid As PmuLimits

Private Const BOARD_NAME As String = "ICUL1GPPMU"

'コンストラクタ
Private Sub Class_Initialize()
    m_PinName = Empty
    m_ChannelNumber = -1
    m_ChannelType = chUnk
    m_Icul1gChannelType = icul1gChannelTypeUnknown
    m_SiteNumber = -1
    m_ForceVoltage = 9999#
    m_ForceCurrent = 9999#
    m_CurrentRange = icul1gPpmuCurrentRangeUnknown
    m_TestLimitHigh = 9999#
    m_TestLimitLow = 9999#
    m_ClampVHi = 9999#
    m_ClampVLo = 9999#
    m_RelayState = icul1gRelayStateUnknown
    m_FIMode = False
    m_FVMode = False
    m_TestLimitValid = pmuNeitherLimitValid
End Sub

'上限/下限値の有効・無効設定値
Public Property Get TestLimitValid() As PmuLimits
    TestLimitValid = m_TestLimitValid
End Property
Public Property Let TestLimitValid(ByVal pLimits As PmuLimits)
    m_TestLimitValid = pLimits
End Property

'ピンの名前
Public Property Get PinName() As String
    PinName = m_PinName
End Property
Public Property Let PinName(ByVal pName As String)
    m_PinName = pName
End Property

'チャンネル番号
Public Property Get ChannelNumber() As Long
    ChannelNumber = m_ChannelNumber
End Property
Public Property Let ChannelNumber(ByVal pNumber As Long)
    m_ChannelNumber = pNumber
End Property

'チャンネルタイプ
Public Property Get ChannelType() As chtype
    ChannelType = m_ChannelType
End Property
Public Property Let ChannelType(ByVal pType As chtype)
    m_ChannelType = pType
End Property

'ICUL1G チャンネルタイプ
Public Property Get Icul1gChannelType() As Icul1gChannelType
    Icul1gChannelType = m_Icul1gChannelType
End Property
Public Property Let Icul1gChannelType(ByVal pType As Icul1gChannelType)
    m_Icul1gChannelType = pType
End Property

'サイト番号
Public Property Get SiteNumber() As Long
    SiteNumber = m_SiteNumber
End Property
Public Property Let SiteNumber(ByVal pNumber As Long)
    m_SiteNumber = pNumber
End Property

'印加電圧
Public Property Get ForceVoltage() As Double
    ForceVoltage = m_ForceVoltage
End Property
Public Property Let ForceVoltage(ByVal pValue As Double)
    m_ForceVoltage = pValue
End Property

'印加電流
Public Property Get ForceCurrent() As Double
    ForceCurrent = m_ForceCurrent
End Property
Public Property Let ForceCurrent(ByVal pValue As Double)
    m_ForceCurrent = pValue
End Property

'電流レンジ
Public Property Get CurrentRange() As Icul1gPpmuCurrentRange
    CurrentRange = m_CurrentRange
End Property
Public Property Let CurrentRange(ByVal pRange As Icul1gPpmuCurrentRange)
    m_CurrentRange = pRange
End Property

'テスト上限値
Public Property Get TestLimitHigh() As Double
    TestLimitHigh = m_TestLimitHigh
End Property
Public Property Let TestLimitHigh(ByVal pValue As Double)
    m_TestLimitHigh = pValue
End Property

'テスト下限値
Public Property Get TestLimitLow() As Double
    TestLimitLow = m_TestLimitLow
End Property
Public Property Let TestLimitLow(ByVal pValue As Double)
    m_TestLimitLow = pValue
End Property

'電圧クランプHi
Public Property Get ClampVoltageHigh() As Double
    ClampVoltageHigh = m_ClampVHi
End Property
Public Property Let ClampVoltageHigh(ByVal pValue As Double)
    m_ClampVHi = pValue
End Property

'電圧クランプLo
Public Property Get ClampVoltageLow() As Double
    ClampVoltageLow = m_ClampVLo
End Property
Public Property Let ClampVoltageLow(ByVal pValue As Double)
    m_ClampVLo = pValue
End Property

'接続状態
Public Property Get relayState() As Icul1gRelayState
    relayState = m_RelayState
End Property
Public Property Let relayState(ByVal pValue As Icul1gRelayState)
    m_RelayState = pValue
End Property

'電流印加モード
Public Property Get FIMode() As Boolean
    FIMode = m_FIMode
End Property
Public Property Let FIMode(ByVal pValue As Boolean)
    m_FIMode = pValue
End Property

'電圧印加モード
Public Property Get FVMode() As Boolean
    FVMode = m_FVMode
End Property
Public Property Let FVMode(ByVal pValue As Boolean)
    m_FVMode = pValue
End Property

'ボード名
Public Property Get BoardName() As String
    BoardName = BOARD_NAME
End Property

'チャンネルタイプ文字列
Public Property Get StrChannelType() As String
    StrChannelType = ChannelType2String(m_ChannelType)
End Property

'ICUL1Gチャンネルタイプ文字列
Public Property Get StrIcul1gChannelType() As String
    StrIcul1gChannelType = Icul1gChannelType2String(m_Icul1gChannelType)
End Property

'CurrentRange文字列
Public Property Get StrCurrentRange() As String
    StrCurrentRange = Icul1gPpmuCurrentRange2String(m_CurrentRange)
End Property

'TestLimitValid文字列
Public Property Get StrTestLimitValid() As String
    StrTestLimitValid = PmuLimits2String(m_TestLimitValid)
End Property

'RelayState文字列
Public Property Get StrRelayState() As String
    StrRelayState = Icul1gRelayState2String(m_RelayState)
End Property


'--- 以下 Private ---

'Relay State設定を文字列に変換する関数
Private Function Icul1gRelayState2String(ByVal pState As Icul1gRelayState) As String
    Select Case pState
        Case icul1gRelayStateClosed
            Icul1gRelayState2String = "icul1gRelayStateClosed"
        Case icul1gRelayStateOpen
            Icul1gRelayState2String = "icul1gRelayStateOpen"
        Case icul1gRelayStateUnknown
            Icul1gRelayState2String = "icul1gRelayStateUnknown"
        Case Else
            Icul1gRelayState2String = "UnKnown"
    End Select
End Function

'PMUのリミット設定を文字列情報に変換する関数
Private Function PmuLimits2String(ByVal pLimits As PmuLimits) As String
    Select Case pLimits
        Case pmuBothLimitsValid
            PmuLimits2String = "pmuBothLimitsValid"
        Case pmuHighLimitValid
            PmuLimits2String = "pmuHighLimitValid"
        Case pmuLowLimitValid
            PmuLimits2String = "pmuLowLimitValid"
        Case pmuNeitherLimitValid
            PmuLimits2String = "pmuNeitherLimitValid"
        Case Else
            PmuLimits2String = "UnKnown"
    End Select
End Function

'ICUL1G電流レンジを文字列情報に変換する関数
Private Function Icul1gPpmuCurrentRange2String(ByVal pRange As Icul1gPpmuCurrentRange) As String
    Select Case pRange
        Case icul1gPpmuCurrentRange10mA
            Icul1gPpmuCurrentRange2String = "icul1gPpmuCurrentRange10mA"
        Case icul1gPpmuCurrentRange200uA
            Icul1gPpmuCurrentRange2String = "icul1gPpmuCurrentRange200uA"
        Case icul1gPpmuCurrentRange20uA
            Icul1gPpmuCurrentRange2String = "icul1gPpmuCurrentRange20uA"
        Case icul1gPpmuCurrentRange2mA
            Icul1gPpmuCurrentRange2String = "icul1gPpmuCurrentRange2mA"
        Case icul1gPpmuCurrentRangeAutoRange
            Icul1gPpmuCurrentRange2String = "icul1gPpmuCurrentRangeAutoRange"
        Case icul1gPpmuCurrentRangeSmartRange
            Icul1gPpmuCurrentRange2String = "icul1gPpmuCurrentRangeSmartRange"
        Case icul1gPpmuCurrentRangeUnknown
            Icul1gPpmuCurrentRange2String = "icul1gPpmuCurrentRangeUnknown"
        Case Else
            Icul1gPpmuCurrentRange2String = "UnKnown"
    End Select
End Function

'ICUL1Gチャンネルタイプを文字列情報に変換する関数
Private Function Icul1gChannelType2String(ByVal pType As Icul1gChannelType) As String
    Select Case pType
        Case icul1gChannelTypeClkHigh
            Icul1gChannelType2String = "icul1gChannelTypeClkHigh"
        Case icul1gChannelTypeClkLow
            Icul1gChannelType2String = "icul1gChannelTypeClkLow"
        Case icul1gChannelTypeDataHigh
            Icul1gChannelType2String = "icul1gChannelTypeDataHigh"
        Case icul1gChannelTypeDataLow
            Icul1gChannelType2String = "icul1gChannelTypeDataLow"
        Case icul1gChannelTypeUnknown
            Icul1gChannelType2String = "icul1gChannelTypeUnknown"
        Case Else
            Icul1gChannelType2String = "UnKnown"
    End Select
End Function

'チャンネルタイプを文字列情報に変換する関数
Private Function ChannelType2String(ByVal pType As chtype) As String
    Select Case pType
        Case chICUL1Gclk_high
            ChannelType2String = "chICUL1Gclk_high"
        Case chICUL1Gclk_low
            ChannelType2String = "chICUL1Gclk_low"
        Case chICUL1Gdata_high
            ChannelType2String = "chICUL1Gdata_high"
        Case chICUL1Gdata_low
            ChannelType2String = "chICUL1Gdata_low"
        Case Else
            ChannelType2String = "UnKnown"
    End Select
End Function

#End If
