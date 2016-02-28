VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIcul1gChanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@@@IP750EX Support
'概要:
'   ICUL1Gのチャンネルリスト管理用クラス
'
'目的:
'   ICUL1Gリソースのみのチャンネルマップを提供
'

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1               '1CUL1Gボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

#If ICUL1G_USE <> 0 Then

Private m_List As Collection 'チャンネルリスト

'コンストラクタ
Private Sub Class_Initialize()
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Set m_List = Nothing
End Sub

'指定のピンが、ICUL1Gボードに所属するピンかを確認する関数
Public Function IsICUL1GPins(ByVal pPinList As String, Optional ByVal pType As Icul1gPinType = icul1gPinAny) As Boolean
    'TERADYNE-APIを利用
    IsICUL1GPins = TheHdw.ICUL1G.IsICUL1GPin(pPinList, pType)
End Function

'すべてのチャンネルリストの公開
Public Property Get AllInfo() As Collection
    Call MakeList
    If m_List.Count <> 0 Then
        Set AllInfo = m_List
    Else
        Set AllInfo = Nothing
    End If
End Property

'指定タイプのチャンネルリスト公開
Public Property Get AtChanType(ByVal pTargetType As Icul1gChannelType) As Collection
    Call MakeList
    If m_List.Count <> 0 Then
        Set AtChanType = MakeTypeOfList(pTargetType)
    Else
        Set AtChanType = Nothing
    End If
End Property

'指定ピンのチャンネルリスト公開
Public Property Get AtPin(ByVal pPinName As String) As Collection
    If CheckIsSinglePin(pPinName) = True And Me.IsICUL1GPins(pPinName) = True Then
        Call MakeList
        Set AtPin = New Collection
        Dim Retinfo As CIcul1gChanInfo
        For Each Retinfo In m_List
            If StrComp(Retinfo.PinName, pPinName, vbTextCompare) = 0 Then
                Call AtPin.Add(Retinfo)
            End If
        Next
    Else
        Set AtPin = Nothing
    End If
End Property

'指定ピンリストのチャンネルリスト公開
Public Property Get AtPinList(ByVal pCsvPinList As String) As Collection
    If Me.IsICUL1GPins(pCsvPinList) = True Then
        Dim PinList() As String
        Dim pinCount As Long
        Call TheExec.DataManager.DecomposePinList(pCsvPinList, PinList, pinCount)
        Set AtPinList = New Collection
        Dim itemCount As Long
        Dim Retinfo As CIcul1gChanInfo
        Dim ChanList As Collection
        For itemCount = 0 To UBound(PinList) Step 1
            Set ChanList = Me.AtPin(PinList(itemCount))
            If Not ChanList Is Nothing Then
                For Each Retinfo In ChanList
                    Call AtPinList.Add(Retinfo)
                Next
            End If
        Next itemCount
    Else
        Set AtPinList = Nothing
    End If
End Property

'ひとつのピンのチャンネル情報を返す関数
Public Property Get AtSiteAndPin(ByVal pPinName As String, ByVal pSiteNumber As Long) As CIcul1gChanInfo
    If CheckIsSinglePin(pPinName) = True And Me.IsICUL1GPins(pPinName) = True And CheckSiteNumber(pSiteNumber) = True Then
        Dim Retinfo As Collection
        Set Retinfo = Me.AtSiteAndPinList(pPinName, pSiteNumber)
        Set AtSiteAndPin = Retinfo.Item(1)
    Else
        Set AtSiteAndPin = Nothing
    End If
End Property

'指定サイトのチャンネルリストを返す関数
Public Property Get AtSiteAndPinList(ByVal pPinList As String, ByVal pSiteNumber As Long) As Collection
    If Me.IsICUL1GPins(pPinList) = True And CheckSiteNumber(pSiteNumber) = True Then
        Set AtSiteAndPinList = New Collection
        Dim PinList() As String
        Dim pinCount As Long
        Call TheExec.DataManager.DecomposePinList(pPinList, PinList, pinCount)
        Dim AllSiteList As Collection
        Dim itemCount As Long
        Dim RetChanInfo As CIcul1gChanInfo
        For itemCount = 0 To UBound(PinList) Step 1
            Set AllSiteList = Me.AtPin(PinList(itemCount))
            For Each RetChanInfo In AllSiteList
                If RetChanInfo.SiteNumber = pSiteNumber Then
                    Call AtSiteAndPinList.Add(RetChanInfo)
                End If
            Next
        Next itemCount
    Else
        Set AtSiteAndPinList = Nothing
    End If
End Property


'============== 以下 Private 機能群 ==============

'サイト番号の確認(TRUE:OK)
Private Function CheckSiteNumber(ByVal pNumber As Long) As Boolean
    If pNumber < 0 Or TheExec.sites.ExistingCount - 1 < pNumber Then
        CheckSiteNumber = False
    Else
        CheckSiteNumber = True
    End If
End Function

'ピンリストに複数ピンが定義されているか否かを確認する関数(TRUE:OK)
Private Function CheckIsSinglePin(ByVal pPinName As String) As Boolean
    Dim RetPins() As String
    Dim pinCount As Long
    Call TheExec.DataManager.DecomposePinList(pPinName, RetPins, pinCount)
    If pinCount <> 1 Then
        CheckIsSinglePin = False
    Else
        CheckIsSinglePin = True
    End If
End Function

'指定タイプのチャンネルリストを作成して返す関数
Private Function MakeTypeOfList(ByVal pType As Icul1gChannelType) As Collection
    Set MakeTypeOfList = New Collection
    Dim Retinfo As CIcul1gChanInfo
    For Each Retinfo In m_List
        If Retinfo.Icul1gChanType = pType Then
            Call MakeTypeOfList.Add(Retinfo)
        End If
    Next
    If MakeTypeOfList.Count = 0 Then
        Set MakeTypeOfList = Nothing
    End If
End Function

'ピンとチャンネルのリスト作成関数
Private Sub MakeList()

    Dim RetChannelList() As Long
    Dim RetPinList() As String
    Dim RetSiteCount As Long

    '調査対象のチャンネルタイプリストを作成
    Dim TargetBoardType As Collection
    Set TargetBoardType = New Collection
    With TargetBoardType
        Call .Add(chICUL1Gclk_high)
        Call .Add(chICUL1Gclk_low)
        Call .Add(chICUL1Gdata_high)
        Call .Add(chICUL1Gdata_low)
    End With
    
    'チャンネル情報作成処理
    Set m_List = New Collection
    Dim TargetType As Variant
    Dim RetStatus As Boolean
    For Each TargetType In TargetBoardType
        RetStatus = GetChannelMapInformations(TargetType, RetChannelList, RetPinList, RetSiteCount)
        If RetStatus = True Then
            Call MakeChannelIInformations(TargetType, RetChannelList, RetPinList, m_List, RetSiteCount)
        End If
    Next

End Sub

'チャンネル情報を作る関数
Private Sub MakeChannelIInformations(ByVal pType As Icul1gChannelType, _
ByRef pTargetChanList() As Long, ByRef pTargetPinList() As String, _
ByRef pChannelList As Collection, ByVal pSiteCount As Long)
    
    Dim pinCount As Long
    Dim ChanCount As Long
    Dim StartCount As Long
    Dim CreateCount As Long
    Dim SiteNumber As Long
    Dim ChannelInfo As CIcul1gChanInfo
                
    StartCount = 0
    CreateCount = 0
        
    For pinCount = 0 To UBound(pTargetPinList) Step 1
        SiteNumber = 0
        For ChanCount = StartCount To StartCount + pSiteCount - 1 Step 1
            Set ChannelInfo = New CIcul1gChanInfo
            With ChannelInfo
                .chanType = pType
                .ChanNumber = pTargetChanList(ChanCount)
                .PinName = pTargetPinList(pinCount)
                .SiteNumber = SiteNumber
            End With
            Call pChannelList.Add(ChannelInfo)
            CreateCount = CreateCount + 1
            SiteNumber = SiteNumber + 1
        Next ChanCount
        StartCount = CreateCount
    Next pinCount

End Sub

'指定リソースのチャンネルマップ情報を取得する関数
Private Function GetChannelMapInformations(ByVal pChannelType As chtype, _
ByRef pChannelList() As Long, _
ByRef pPinList() As String, _
ByRef pSiteCount As Long) As Boolean
    
    Const ALL_SITE As Long = -1
    
    '指定リソースを使用しているPIN情報を取得
    Dim RetPinCnt As Long
    Call TheExec.DataManager.GetPinNames(pPinList, pChannelType, RetPinCnt)
                                                   
    '指定されたリソースが存在しないときはFalseを返して終了
    If RetPinCnt = 0 Then
        GetChannelMapInformations = False
        Exit Function
    End If
                                                
    'CSVピンリストを作る
    Dim CsvPinList As String
    CsvPinList = MakeCsvPinList(pPinList)
                    
    '指定リソースのすべてのPINのチャンネル番号を取得
    Dim RetChanCnt As Long
    Dim RetMsg As String
    Call TheExec.DataManager.GetChanList(CsvPinList, ALL_SITE, pChannelType, pChannelList, RetChanCnt, pSiteCount, RetMsg)
    GetChannelMapInformations = True

End Function

'配列形式のピンリストをCSV形式文字列に変換する関数
Private Function MakeCsvPinList(ByRef pPinList() As String) As String
    MakeCsvPinList = pPinList(0)
    Dim itemCount As Long
    For itemCount = 1 To UBound(pPinList) Step 1
        MakeCsvPinList = MakeCsvPinList & "," & pPinList(itemCount)
    Next itemCount
End Function


#End If

