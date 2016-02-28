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
'概要:
'   ICUL1Gのチャンネル情報管理用クラス
'
'目的:
'   関連情報をセットで管理する
'

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1               '1CUL1Gボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

#If ICUL1G_USE <> 0 Then

Private m_ChanNumber As Long
Private m_ChanType As chtype
Private m_Icul1gChanType As Icul1gChannelType
Private m_PinName As String
Private m_SiteNumber As Long

'コンストラクタ
Private Sub Class_Initialize()
    m_ChanNumber = -1
    m_ChanType = chUnk
    m_Icul1gChanType = icul1gChannelTypeUnknown
    m_PinName = Empty
    m_SiteNumber = -1
End Sub

'ピン名
Public Property Get PinName() As String
    PinName = m_PinName
End Property
Public Property Let PinName(ByVal pName As String)
    '本当は指定されたピンがICUL1G-PPMUピンかを調べるチェックを入れるほうがよい
    m_PinName = pName
End Property

'サイト番号
Public Property Get SiteNumber() As Long
    SiteNumber = m_SiteNumber
End Property
Public Property Let SiteNumber(ByVal pValue As Long)
    m_SiteNumber = pValue
End Property

'チャンネル番号(icul1gXXの数字)
Public Property Get ChanNumber() As Long
    ChanNumber = m_ChanNumber
End Property
Public Property Let ChanNumber(ByVal pNumber As Long)
    m_ChanNumber = pNumber
End Property

'チャンネルタイプ
Public Property Get chanType() As chtype
    chanType = m_ChanType
End Property
Public Property Let chanType(ByVal pType As chtype)
    m_ChanType = pType
    m_Icul1gChanType = ChType2Icul1gChannelType(pType)
End Property

'ICUL1Gのチャンネルタイプ(チャンネルタイプ値から自動設定なのでSetはなし)
Public Property Get Icul1gChanType() As Icul1gChannelType
    Icul1gChanType = m_Icul1gChanType
End Property


'--- 以下Private ---

'チャンネルタイプをICUL1Gチャンネルタイプに変換する関数
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
