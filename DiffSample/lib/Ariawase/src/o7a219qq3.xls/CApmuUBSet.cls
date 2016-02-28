VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CApmuUBSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   APMUのUtilityBit設定用クラス
'
'目的:
'   Utility Bit条件表に定義された情報を使用して
'   Utility Bitを設定する
'
'注意事項:
'   APMU Utility Bit条件表のワークシートと組み合わせて使用します。
'
'作成者:
'   SLSI今手
'
'Code Checked
'

Option Explicit

Private m_UtyBitSetWs As Object        ' APMU UtilityBit条件表シートOBJ
Private m_LastConditionName As String  ' 最終設定条件名格納用変数
Private m_UtyBitStatus() As Long       ' UB設定状態記録用
Private Const MAX_BIT_CH = 255         ' UBCHの最大値

'#Pass-Ok
Public Property Get BoardNumber() As Long
'内容:
'   APMUボードNOの設定値を取得する
'
'パラメータ:
'
'戻り値:
'   シートに定義されているAPMUボード番号
'
'注意事項:
'
    
    'ボード番号の定義情報が存在しない場合は
    'ワークシートのデータを読み込みます
    If m_UtyBitSetWs.ApmuBoardNum = -1 Then
        Call Me.LoadCondition
    End If

    BoardNumber = m_UtyBitSetWs.ApmuBoardNum

End Property

#If UB_SNAPSHOT_USE <> 0 Then
'#Pass-Ok UBの設定不整合問題が解決するまでは封印
Public Sub ClearUBInformation()
'内容:
'   Bit設定情報を消去する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    ReDim m_UtyBitStatus(MAX_BIT_CH)

End Sub

'#Pass-Ok UBの設定不整合問題が解決するまでは封印
Public Property Get UBSnapshot(ByVal ChannelNumber As Long) As Long
'内容:
'   指定CH番号の現在の設定値を取得する
'
'パラメータ:
'   [ChannelNumber]  In  Long型:    状態を取得するUBのCH番号
'
'戻り値:
'   Bit設定値
'
'注意事項:
'   Utility Bits Display、TERADYNE-APIによる設定値は
'   取得結果に反映されません
'

    UBSnapshot = m_UtyBitStatus(ChannelNumber)

End Property

'#Pass-Ok UBの設定不整合問題が解決するまでは封印
Public Property Get UBSnapshotInHex() As String
'内容:
'   0-255CH 各CHの現在の設定値をHex形式で取得する
'
'パラメータ:
'
'戻り値:
'   HEX形式の0-255chのBit設定値
'
'注意事項:
'   Utility Bits Display、TERADYNE-APIを使用して行った設定値は
'   取得結果に反映されません
'
    
    UBSnapshotInHex = MakeHexMsg(m_UtyBitStatus)

End Property
#End If

'#Pass-Ok
Public Property Set UBSetSht(ByVal UtyBitConditionSht As Worksheet)
'内容:
'   UtilityBit設定条件データのあるワークシートを指定する
'
'パラメータ:
'   [UtyBitConditionSht]    In   UtilityBit条件設定ワークシート
'
'戻り値:
'
'注意事項:
'
    
    Set m_UtyBitSetWs = UtyBitConditionSht

End Property

'#Pass-Ok
Public Property Get UBSetShtName() As String
'内容:
'  UtilityBit設定条件ワークシート名
'
'パラメータ:
'
'戻り値:
'  UtilityBit設定条件ワークシート名
'
'注意事項:
'
    
    UBSetShtName = m_UtyBitSetWs.Name

End Property

'#Pass-Ok
Public Sub LoadCondition()
'内容:
'   リレー設定条件データの初期化を行います。
'  （ワークシートに定義されたデータを読み込みます。）
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   ワークシートイベントが発行できない状況(例:ブレークポイントによる中断中)
'   にてリレー条件表の値を変更した場合は変更した内容が反映されないので､
'   リレー設定機能の使用前に必ずこの命令を実行し､条件表の値の再読み込みを
'   行ってください｡
'
'   ＊リレー条件表ワークシート中のデータの読み込み機能を実行しています。
'
    
    'APMU Utility Bit settingのワークシートがプロパティ[UBSetSht]
    'に設定されていないときにはエラー
    If m_UtyBitSetWs Is Nothing Then
        Call OutputErrMsg("APMU Utility Bit setting worksheet is nothing" & vbCrLf _
        & "    Please set UBSetSht")
        Exit Sub
    End If
    
    Call m_UtyBitSetWs.Initialize
    
End Sub

'#Pass-Ok
Public Sub SetUB(ByVal AssignName As String, ByVal SetValue As Long)
'内容:
'   AssignNameに対応するUBを指定値に設定にする
'
'パラメータ:
'   [AssignName]    In     AssignName
'   [Status]        In     設定値
'
'戻り値:
'
'注意事項:
'   UB条件表に定義されていない名前を指定した場合はエラーとなります。
'   複数Bitを一度に設定する際は、[AssignName]をカンマ区切りで指定します。
'       例:  UbObj.SetUB("Name1,Name2,Name3",1)
'

    'APMU Utility Bit settingのワークシートがプロパティ[UBSetSht]
    'に設定されていないときにはエラー
    If m_UtyBitSetWs Is Nothing Then
        Call OutputErrMsg("APMU Utility Bit setting worksheet is nothing" & vbCrLf _
        & "Please set UBSetSht")
        Exit Sub
    End If

    'アサイン名とUB_CHの定義情報が存在しない場合は
    'ワークシートのデータを読み込みます
    If m_UtyBitSetWs.UBChanMap Is Nothing Then
        Call Me.LoadCondition
    End If
        
    Dim assignNameArr() As String
    Dim tmpAssignName As Variant
        
    'カンマ区切り文字列を配列形式に変換
    Call StrToArray(AssignName, assignNameArr)
    
    'UBをシリアルに設定していく
    For Each tmpAssignName In assignNameArr
        With TheHdw.APMU.board(m_UtyBitSetWs.ApmuBoardNum)
            On Error GoTo UB_SET_ERR
            .UtilityBit(m_UtyBitSetWs.UBChanMap(CStr(tmpAssignName))) = SetValue
            m_UtyBitStatus(m_UtyBitSetWs.UBChanMap(CStr(tmpAssignName))) = SetValue 'UtyBitの設定状態取得用
            On Error GoTo 0
        End With
    Next tmpAssignName
    
    Exit Sub
    
UB_SET_ERR:
    
    Call OutputErrMsg(tmpAssignName & " Assign UB_CH# is Not Found")
    
End Sub

Public Sub SetUBCondition(ByVal pCsvLabelList As String)
'内容:
'   指定したコンディション名で定義された値にUBを一括設定する
'
'パラメータ:
'   [pCsvLabelList]  In  String型:    条件名称
'
'戻り値:
'
'注意事項:
'   リレー条件表に定義されていないコンディション名称を
'   指定した場合はエラーとなります。
'
'   条件表で値が空欄（未設定）のCHは、現状の値のままとなります。
'
'   複数の条件を一括指定する際には、カンマ区切りで指定します。
'   例:  UbObj.SetUBCondition("Name1,Name2,Name3")
                       
                       
    'パラメータシートの存在確認
    If m_UtyBitSetWs Is Nothing Then
        Call OutputErrMsg("APMU Utility Bit setting worksheet is nothing" & vbCrLf _
        & "Please set UBSetSht")
        Exit Sub
    End If
                       
    'パラメータキャッシュ存在確認
    If m_UtyBitSetWs.UBCondition Is Nothing Then
        Call Me.LoadCondition
    End If
                
    'パラメータキャッシュ参照取り出し
    Dim UBParamsCache As Collection
    Set UBParamsCache = m_UtyBitSetWs.UBCondition
                
    'ラベルリストを配列に展開
    If Len(pCsvLabelList) = 0 Then
        'リスト空文字列チェック
        Call OutputErrMsg("Label name is empty")
    End If
    Dim LabelList() As String
    LabelList = Split(pCsvLabelList, ",", , vbTextCompare)
    Dim LabelCount As Long
    For LabelCount = 0 To UBound(LabelList) Step 1
        If Len(LabelList(LabelCount)) = 0 Then
            'リスト内の空文字列チェック
            Call OutputErrMsg("The empty label is included")
        End If
    Next
                                    
    'マージラベルを作る(配列内の文字列を全結合)
    Dim Mergelabel As String
    Mergelabel = Join(LabelList, Empty)
    
    'マージラベルで設定データを取得
    Dim MergeParams As Collection
    On Error Resume Next
    Set MergeParams = UBParamsCache(Mergelabel)
    On Error GoTo 0
    
    'マージパラメータキャッシュ存在判定
    If MergeParams Is Nothing Then
        'マージパラメータ作成処理
        Set MergeParams = MakeMergeParams(LabelList, UBParamsCache)
        'キャッシュにマージパラメータ登録
        Call UBParamsCache.Add(MergeParams, Mergelabel)
    End If
    
    'UB設定動作
    Dim UBParam As Variant
    With TheHdw.APMU.board(m_UtyBitSetWs.ApmuBoardNum)
        On Error GoTo UB_SET_ERR
        For Each UBParam In MergeParams
            .UtilityBit(CLng(UBParam(0))) = CLng(UBParam(1))
            m_UtyBitStatus(CLng(UBParam(0))) = CLng(UBParam(1))
        Next
        On Error GoTo 0
    End With
    
    '最終設定名称格納(ただし取り出し口は無い)
    m_LastConditionName = Mergelabel
    
    Exit Sub
    
'CH設定処理エラーハンドル
UB_SET_ERR:
    Call OutputErrMsg("Could not be set CH=" & UBParam(0))

End Sub

'#Pass-Ok
Public Sub ResetUtilityBits()
'内容:
'   ビットの設定値のリセット(全ビットをOFFにする)
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    
    Dim apmuCh As Long
        
    'ボード番号の定義情報が存在しない場合は
    'ワークシートのデータを読み込みます
    If m_UtyBitSetWs.ApmuBoardNum = -1 Then
        Call Me.LoadCondition
    End If
        
    'すべてのAPMU CHのBit設定値を0に設定する
    For apmuCh = 0 To MAX_BIT_CH Step 1
        With TheHdw.APMU.board(m_UtyBitSetWs.ApmuBoardNum)
            .UtilityBit(apmuCh) = 0
        End With
        m_UtyBitStatus(apmuCh) = 0
    Next apmuCh

End Sub

'#Pass-Ok
Private Sub Class_Initialize()
    
    Set m_UtyBitSetWs = Nothing
    m_LastConditionName = ""
    ReDim m_UtyBitStatus(MAX_BIT_CH)

End Sub

'#Pass-Ok
Private Sub StrToArray(ByVal astrInputStr As String, ByRef alngOutputArray() As String)
'カンマ区切り文字列を配列形式に変換する  CUBUtilBitCtrlよりコピー
    
    Dim ret As Long
    Dim i As Long

    Erase alngOutputArray()

    Do
        ret = InStr(1, astrInputStr, ",")

        If ret = 0 Then
            ReDim Preserve alngOutputArray(i)
            alngOutputArray(i) = astrInputStr
            Exit Do
        End If
        ReDim Preserve alngOutputArray(i)
        alngOutputArray(i) = Left(astrInputStr, ret - 1)
        astrInputStr = Right(astrInputStr, Len(astrInputStr) - ret)
        i = i + 1
    Loop

End Sub

'#Pass-Ok
Private Function MakeHexMsg(bitArr() As Long) As String
'配列に格納されているBit値よりHex出力形式のメッセージを作成

    Dim lpCnt As Long
    Dim strBin As String
    Dim chkCnt As Long
    Dim dispMsg As String

    chkCnt = UBound(bitArr) - 3
    strBin = ""
    dispMsg = ""
        
    For lpCnt = UBound(bitArr) To 0 Step -1
        strBin = strBin & bitArr(lpCnt)
        If lpCnt = chkCnt Then
            dispMsg = dispMsg & Hex$(BinToDec(strBin))
            strBin = ""
            chkCnt = chkCnt - 4
        End If
    Next lpCnt

    MakeHexMsg = dispMsg

End Function

'#Pass-Ok
Private Function BinToDec(BinValue As String) As Variant
'Bit値文字列を10進数に変換する
    
    Dim lpCnt As Long
    Dim StrVal As String
    Dim Cursor As Long
  
    '数値でないものが入ってきたら変換しない
    If IsNumeric(BinValue) = False Then
        BinToDec = Empty
        Exit Function
    End If
  
    Cursor = 0
    
    For lpCnt = Len(BinValue) To 1 Step -1
        Select Case Mid$(BinValue, lpCnt, 1)
            '0のとき
            Case "0"
            '1のとき
            Case "1"
                BinToDec = BinToDec + (2 ^ Cursor)
            '0/1以外の数値が指定された場合
            Case Else
                BinToDec = Empty
                Exit Function
        End Select
        Cursor = Cursor + 1
    Next lpCnt

End Function

'マージパラメータ作成関数
Private Function MakeMergeParams(ByRef pLabelList() As String, ByRef pParams As Collection) As Collection
    Dim LabelCount As Long
    Dim UBParams As Collection
    Dim UBParam As Variant
    Set MakeMergeParams = New Collection
    For LabelCount = 0 To UBound(pLabelList) Step 1
        On Error GoTo PARAM_READ_ERROR
        Set UBParams = pParams.Item(pLabelList(LabelCount))
        On Error GoTo 0
        For Each UBParam In UBParams
            On Error Resume Next
            Call MakeMergeParams.Remove(CStr(UBParam(0)))
            On Error GoTo 0
            Call MakeMergeParams.Add(UBParam, CStr(UBParam(0)))
        Next UBParam
    Next LabelCount
    Exit Function

'パラメータ未定義エラー
PARAM_READ_ERROR:
    Set MakeMergeParams = Nothing
    Call OutputErrMsg(pLabelList(LabelCount) & " Condition is Not Found")
End Function
