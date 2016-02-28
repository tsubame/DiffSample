VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIcul1gPpmuResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@@@IP750EX Support
'概要:
'   PinListData型のクラスをICUL1のMeasure結果用に拡張したクラス
'
'目的:
'   不足機能の追加
'

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1               '1CUL1Gボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

#If ICUL1G_USE <> 0 Then

Private m_PinListData As PinListData 'PINLISTデータObejct

'コンストラクタ
Private Sub Class_Initialize()
    Set m_PinListData = New PinListData                    'PINLISTデータObejct生成
    m_PinListData.ResultType = tlResultTypeParametricValue '測定値格納Mode設定
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Set m_PinListData = Nothing
End Sub

'測定対象ピンリストの公開
Public Property Get PinList() As String
    If m_PinListData.Pins.Count = 0 Then
        PinList = Empty
    Else
        PinList = PinListData2CsvPinList
    End If
End Property

'PinListDataオブジェクトの公開
Public Property Get PinListData() As PinListData
    Set PinListData = m_PinListData
End Property

'PerPinResultの公開
Public Property Get PerPinResult() As Collection
    If m_PinListData.Pins.Count = 0 Then
        '測定データ未登録
        Set PerPinResult = Nothing
    Else
        Set PerPinResult = PinListData2PerPinResults(m_PinListData)
    End If
End Property

'PinResultの公開
Public Property Get PinResult(ByVal pPinName As String) As Double()
    If m_PinListData.Pins.Count = 0 Then
        '測定データ未登録
        Erase PinResult
    Else
        Dim PerPinResult As Collection
        Set PerPinResult = PinListData2PerPinResults(m_PinListData)
        PinResult = PerPinResult.Item(pPinName)
    End If
End Property


'--------- 以下 Private --------------------------------------------------

'ピンリストデータからPerPinResultデータを作成（アベレージの割戻し、無効サイトの結果値は0を返す）
Private Function PinListData2PerPinResults(ByRef pResultData As PinListData) As Collection

    Set PinListData2PerPinResults = New Collection
    
    'CSVピンリストを配列ピンリストにする
    Dim RetPinList() As String
    Dim RetPinCount As Long '未使用
    With TheExec.DataManager
        Call .DecomposePinList(PinListData2CsvPinList, RetPinList, RetPinCount)
    End With
        
    Dim pinCount As Long
    Dim tempResult As Double
    Dim SiteCount As Long
    Dim retResult As Variant
    Dim TargetPin As PinData
    Dim SiteResult() As Double
    Dim SampleCount As Long
    Dim targetSite As Long
    
    'ピンリストLOOP
    For pinCount = 0 To UBound(RetPinList) Step 1
        
        '処理対象ピンの選択
        Set TargetPin = pResultData.Pins(RetPinList(pinCount))
        
        'サイト毎結果の入れ物用意
        ReDim SiteResult(TheExec.sites.ExistingCount - 1)
    
        '測定時の有効サイト番号LOOP
        For SiteCount = 0 To UBound(pResultData.sites) Step 1
                
                'ターゲットサイトのデータ取り出し（測定時の有効サイトのみ）
                targetSite = pResultData.sites(SiteCount)
                retResult = TargetPin.Value(targetSite)
                
                'アベレージあり測定データ
                If IsArray(retResult) = True Then
                    'アベレージ割戻し処理
                    tempResult = 0#
                    For SampleCount = 0 To UBound(retResult) Step 1
                        'SAMPLE合計データ算出
                        tempResult = tempResult + retResult(SampleCount)
                    Next SampleCount
                    '割戻し処理
                    SiteResult(targetSite) = tempResult / (UBound(retResult) + 1)
                'アベレージなし測定データ
                Else
                    SiteResult(targetSite) = CDbl(retResult)
                End If
                
        Next SiteCount
    
        'ピンのデータとしてサイト配列を登録
        Call PinListData2PerPinResults.Add(SiteResult, RetPinList(pinCount))
    
    Next pinCount

End Function

'PinListのデータからCSVPinListを作成する関数
Private Function PinListData2CsvPinList() As String
    Dim pinCount As Long
    PinListData2CsvPinList = Empty
    For pinCount = 0 To m_PinListData.Pins.Count - 1 Step 1
        If pinCount = 0 Then
            PinListData2CsvPinList = m_PinListData.Pins(pinCount).Name
        Else
            PinListData2CsvPinList = PinListData2CsvPinList & "," & m_PinListData.Pins(pinCount).Name
        End If
    Next pinCount
End Function

#End If

