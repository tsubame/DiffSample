VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIcul1gSnapshotProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@@@IP750EX Support
'概要:
'   ICUL1Gスナップショット提供クラス
'
'目的:
'   PPMU情報の提供
'

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1               '1CUL1Gボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

#If ICUL1G_USE <> 0 Then

Private m_Icul1gPPMUList As Collection   'PPMUリスト
Private m_ChannelList As CIcul1gChanList 'チャンネルリスト

'コンストラクタ
Private Sub Class_Initialize()
    Set m_ChannelList = New CIcul1gChanList
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Set m_ChannelList = Nothing
    Set m_Icul1gPPMUList = Nothing
End Sub

'情報初期化
Public Sub Initialize()
    Set m_Icul1gPPMUList = Nothing
    Set m_ChannelList = New CIcul1gChanList
End Sub

'すべてのPPMU情報リストの公開
Public Property Get PpmuList() As Collection
    Call MakePPMUList
    Set PpmuList = m_Icul1gPPMUList
End Property

'指定ピンのPPMU情報リストの公開
Public Property Get PpmuListAtPinList(ByVal pCsvPinList As String) As Collection
    If m_ChannelList.AllInfo Is Nothing = False And m_ChannelList.IsICUL1GPins(pCsvPinList) = True Then
        Dim PinList As Collection
        Set PinList = m_ChannelList.AtPinList(pCsvPinList)
        Dim RetChanInfo As CIcul1gChanInfo
        Set PpmuListAtPinList = New Collection
        For Each RetChanInfo In PinList
            Call PpmuListAtPinList.Add(MakePpmuInfo(RetChanInfo))
        Next
    Else
        Set PpmuListAtPinList = Nothing
    End If
End Property

'外部機能公開I/F(従来機能の互換用)
Public Sub CreatePpmuInfo(Optional ByVal pIdLabel As String = "*", _
Optional ByVal pDataWindowUse As Integer = 1, _
Optional ByVal pFileName As String = Empty)

    'リソース存在判定
    Dim retList As Collection
    Set retList = m_ChannelList.AllInfo
    If Not retList Is Nothing Then
        
        'PPMUリストを作る
        Call MakePPMUList
    
        'CSV PPMUリストを作る
        Dim CsvPPMUList As Collection
        Set CsvPPMUList = MakeCSVPPMUList(pIdLabel)
    
        '情報出力処理
        Dim Retinfo As Variant
        Dim itemCount As Long
        itemCount = 1
        For Each Retinfo In CsvPPMUList
            'IG-XLのログに出す
            If pDataWindowUse = 1 Then
                Call TheExec.Datalog.WriteComment(CStr(Retinfo))
                If itemCount = CsvPPMUList.Count Then
                    Call TheExec.Datalog.WriteComment(Empty)
                End If
            End If
            'TXTファイルに出す
            If pFileName <> Empty Then
                Call AppendTxtWriter(pFileName, CStr(Retinfo))
                If itemCount = CsvPPMUList.Count Then
                    Call AppendTxtWriter(pFileName, Empty)
                End If
            End If
            itemCount = itemCount + 1
        Next
    
    Else
    '定義がない（ボードリソース未使用）
        Dim outPutMsg As String
        outPutMsg = "@@@" & "," & "[SnapShot]" & "," & "ICUL1GPPMU" & ".Type_doesn't_exist_in_the_ChannelMap"
        
        'IG-XLのログに出す
        If pDataWindowUse = 1 Then
            Call TheExec.Datalog.WriteComment(outPutMsg)
            Call TheExec.Datalog.WriteComment(Empty)
        End If
        
        'TXTファイルに出す
        If pFileName <> Empty Then
            Call AppendTxtWriter(pFileName, outPutMsg)
            Call AppendTxtWriter(pFileName, Empty)
        End If
    
    End If

End Sub

'PPMUリストの作成関数
Private Sub MakePPMUList()
    '情報取得処理
    If m_ChannelList.AllInfo Is Nothing = False Then
        Set m_Icul1gPPMUList = New Collection
        'チャンネルリストLOOP
        Dim RetChanInfo As CIcul1gChanInfo
        For Each RetChanInfo In m_ChannelList.AllInfo
            Call m_Icul1gPPMUList.Add(MakePpmuInfo(RetChanInfo))
        Next
    Else
        Set m_Icul1gPPMUList = Nothing
    End If
End Sub

'PPMU状態取得用関数
Private Function MakePpmuInfo(ByRef pTarget As CIcul1gChanInfo) As CIcul1gPpmuInfo
    
    'PPMUオブジェクトの取得
    Dim TargetChan As DriverICUL1G_Chan
    Set TargetChan = TheHdw.ICUL1G.Chans(pTarget.ChanNumber, pTarget.Icul1gChanType)
    Dim TargetPpmu As DriverICUL1G_PPMU_chan
    Set TargetPpmu = TargetChan.PPMU
                    
    'PPMU情報取得
    Set MakePpmuInfo = New CIcul1gPpmuInfo
    Dim RetState As Icul1gRelayState
    With MakePpmuInfo
        .ChannelNumber = pTarget.ChanNumber
        .ChannelType = pTarget.chanType
        .Icul1gChannelType = pTarget.Icul1gChanType
        .ClampVoltageHigh = TargetPpmu.ClampVHi
        .ClampVoltageLow = TargetPpmu.ClampVLo
        Call TargetChan.IsConnected(RetState)
        .relayState = RetState
        .CurrentRange = TargetPpmu.CurrentRange
        .FIMode = TargetPpmu.isForcingCurrent
        .ForceCurrent = TargetPpmu.ForceCurrent(.CurrentRange)
        .ForceVoltage = TargetPpmu.ForceVoltage(.CurrentRange)
        .FVMode = TargetPpmu.IsForcingVoltage
        .PinName = pTarget.PinName
        .SiteNumber = pTarget.SiteNumber
        .TestLimitHigh = TargetPpmu.TestLimitHigh
        .TestLimitLow = TargetPpmu.TestLimitLow
        .TestLimitValid = TargetPpmu.TestLimitValid
    End With

End Function

'TXTファイルに文字列を出力する関数
Private Sub AppendTxtWriter(ByVal pFileName As String, pMessage As String)
    Dim fp As Integer
    On Error GoTo APPEND_ERR
    fp = FreeFile
    Open pFileName For Append As fp
    Print #fp, pMessage
    Close fp
    Exit Sub
APPEND_ERR:
    Err.Raise "9999", TypeName(Me), "Append text error"
End Sub

'PPMUリストからCSV形式のPPMUリストを作成する関数
Private Function MakeCSVPPMUList(ByVal pLabel As String) As Collection
    If m_Icul1gPPMUList.Count <> 0 Then
        Set MakeCSVPPMUList = New Collection
        Dim RetPpmuinfo As CIcul1gPpmuInfo
        For Each RetPpmuinfo In m_Icul1gPPMUList
            Call MakeCSVPPMUList.Add(MakeCSVPPMUInfo(pLabel, RetPpmuinfo))
        Next
    Else
        Set MakeCSVPPMUList = Nothing
    End If
End Function

'PPMU情報からCSV形式のPPMU情報を作成する関数
Private Function MakeCSVPPMUInfo(ByVal pIdLable As String, ByRef pPpmuinfo As CIcul1gPpmuInfo) As String
    With pPpmuinfo
        MakeCSVPPMUInfo = pIdLable & "," _
            & "[" & .BoardName & "]" & "," _
            & "PIN=" & .PinName & "," _
            & "SITE=" & .SiteNumber & "," _
            & "CH_NUM=" & .ChannelNumber & "," _
            & "STR_CH_TYPE=" & .StrChannelType & "," _
            & "CH_TYPE=" & .ChannelType & "," _
            & "STR_ICUL1G_CH_TYPE=" & .StrIcul1gChannelType & "," _
            & "ICUL1G_CH_TYPE=" & .Icul1gChannelType & "," _
            & "FVMODE=" & .FVMode & "," _
            & "FV=" & .ForceVoltage & "," _
            & "FIMODE=" & .FIMode & "," _
            & "FI=" & .ForceCurrent & "," _
            & "STR_IRANGE=" & .StrCurrentRange & "," _
            & "IRANGE=" & .CurrentRange & "," _
            & "STR_RELAY=" & .StrRelayState & "," _
            & "RELAY=" & .relayState & "," _
            & "CLAMP_VH=" & .ClampVoltageHigh & "," _
            & "CLAMP_VL=" & .ClampVoltageLow & "," _
            & "LIMIT_HIGH=" & .TestLimitHigh & "," _
            & "LIMIT_LOW=" & .TestLimitLow & "," _
            & "STR_TEST_LIMIT_VALID=" & .StrTestLimitValid & "," _
            & "LIMIT_VALID=" & .TestLimitValid
    End With
End Function


#End If
