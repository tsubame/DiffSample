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

'HDVISメジャー時のパラメータ提供用 機能群クラス
'ハードウェアの不具合修正が行われるまでの暫定対策用

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const HDVIS_USE = 0      'HVDISリソースの使用　 0：未使用、0以外：使用  <PALSとEeeAuto両方で使用>

#If HDVIS_USE <> 0 Then
Private Const CS_NAME = "CVISHdvisParameter"
Private Const PARAM_WSHT_NAME = "HDVIS_Param" 'HDVIS Measureパラメータワークシート名(複数シートを切り替えることは考慮していない)
Private m_ParameterSht As Object
'Private m_ParameterSht As ShtHdvisParam 'コーディング時は入力支援がないとつらいのでこちらを使用した

'#Pass-Ok
Private Sub Class_Initialize()
    Call Me.Initialize
End Sub

'#Pass-Ok
Public Sub Initialize()
'初期化（Parameterワークシートの設定、現存Parameter Cacheのフラッシュ、Parameter Cacheの作成）

    On Error GoTo PARAM_WSHT_ERROR
    Set m_ParameterSht = Worksheets(PARAM_WSHT_NAME) 'HDVISパラメータシートの指定
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
'ワークシートのパラメータの読み込み
'(現存Parameter Cacheのフラッシュ、Parameter Cacheの作成)
    
    If m_ParameterSht Is Nothing Then
        Call Me.Initialize
    End If
    
    m_ParameterSht.LoadParameter

End Sub


'------------------------------------------------------------------------------
'パラメータ公開用機能群
'------------------------------------------------------------------------------
'#Pass-Ok
Public Property Get ParameterSheetName() As String
'HDVISパラメータワークシート名の公開
    
    If m_ParameterSht Is Nothing Then
        Call Me.Initialize
    End If
    
    ParameterSheetName = m_ParameterSht.Name

End Property

'#Pass-Ok
Public Property Get PinCapacitor(ByVal PinName As String) As Double
'指定ピンの容量値を公開
    
    If m_ParameterSht Is Nothing Then
        Call Me.Initialize
    End If
    
    PinCapacitor = m_ParameterSht.PinCapacitor(PinName)

End Property

'#Pass-Ok
Public Property Get Samples(ByVal PinName As String, ByVal Irange As HdvisIRange) As Long
'指定ピン、指定レンジのSamples値を公開
            
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
'指定ピン、指定レンジのPeriod値を公開
        
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
'ワークシートのSlewRate設定をHdvisSlewRate型に変換して公開
    
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
            Call OutputErrMsg(tmpParam & " Slew Rate parameter is unknown" & " at " & CS_NAME) 'キャッシュ作成時チェックがあるので通常ココのエラーは出ない
    End Select

End Property


'------------------------------------------------------------------------------
'メジャー時のパラメータチェック機能群
'------------------------------------------------------------------------------
'#Pass-Ok
Public Function ChkSameIrange(ByRef HdvisIRange() As HdvisIRange) As Boolean
'配列に格納されている測定レンジ値がすべて同一値であることを確認する
    
    Dim chCnt As Long

    For chCnt = 1 To UBound(HdvisIRange) Step 1
        If HdvisIRange(0) <> HdvisIRange(chCnt) Then 'HdvisIrange型(Long)の値比較
            ChkSameIrange = False
            Exit Function
        End If
    Next chCnt

    ChkSameIrange = True
    
End Function

'#Pass-Ok
Public Function ChkSameCapacitor(ByRef PinList As String) As Boolean
'ピンリストのピン容量の定義値がすべて同一であることを確認
    
    Dim pinNameArr() As String
    Dim PinCnt As Long
    Dim refCapacitor As Double
        
    'カンマ区切りのピンリストを配列形式に変換
    Call SeparatePinList(PinList, pinNameArr) 'これにPinMapに定義の無いPinを渡すと実行時エラーとなる

    'ピンリストの先頭ピンの容量値を基準とする
    With Me
        refCapacitor = .PinCapacitor(pinNameArr(0))
        For PinCnt = 1 To UBound(pinNameArr) Step 1 '要素0はRefなので要素1から
            If CompareDblData(refCapacitor, .PinCapacitor(pinNameArr(PinCnt)), 15) = False Then '最小単位fFで切捨てし同じ容量値かを判定
                ChkSameCapacitor = False
                Exit Function
            End If
        Next PinCnt
    End With

    ChkSameCapacitor = True
    
End Function

'#Pass-Ok
Public Function ChkIRangeIncluded(ByVal targetIrange As HdvisIRange, ByRef chkIrange() As HdvisIRange) As Boolean
'配列に格納されている測定レンジ値に、指定レンジの値がひとつでも含まれているかを確認
    
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
'ピンリストに指定されたピンのメジャーパラメータ(Samples、Period)がすべて同一であるかを確認
    
    Dim pinNameArr() As String
    Dim PinCnt As Long
    Dim refSamples As Double
    Dim refPeriod As Double
        
    'カンマ区切りのピンリストを配列形式に変換
    Call SeparatePinList(PinList, pinNameArr) 'これにPinMapに定義の無いPinを渡すと実行時エラーとなる

    'ピンリストの先頭ピンのSamples、Periodを基準（比較元）とする
    With Me
        refSamples = .Samples(pinNameArr(0), Irange)
        refPeriod = .Period(pinNameArr(0), Irange)
        For PinCnt = 1 To UBound(pinNameArr) Step 1 '要素0はRefなので要素1から
            'Samplesの値比較
            If refSamples <> .Samples(pinNameArr(PinCnt), Irange) Then
                ChkSameMeasureParameter = False
                Exit Function
            End If
            'Periodの値比較
            If CompareDblData(refPeriod, .Period(pinNameArr(PinCnt), Irange), 15) = False Then '最小単位fSecで切捨てし同じ時間かを判定
                ChkSameMeasureParameter = False
                Exit Function
            End If
        Next PinCnt
    End With

    ChkSameMeasureParameter = True

End Function

Public Sub DispMeasureCondition(ByVal PinList As String)
'メジャー時にパラメータが反映されているか確認するための実験機能

    Dim tmpMode() As HdvisMode
    Dim tmpVrange() As HdvisVRange
    Dim tmpIrange() As HdvisIRange
    Dim tmpSlewRate() As HdvisSlewRate

    Dim MeaIrange As HdvisIRange
    Dim MeaPriod As Double
    Dim MeaPinNames() As String
    Dim MeaSlewRate As HdvisSlewRate
    Dim statMsg As String
    
    'テラダインAPIよりボードの設定パラメータを取得
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

