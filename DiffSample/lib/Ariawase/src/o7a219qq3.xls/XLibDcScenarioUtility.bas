Attribute VB_Name = "XLibDcScenarioUtility"
'概要:
'   DCテストシナリオシートのためのライブラリ群
'
'目的:
'
'作成者:
'   SLSI大谷
'   2013/10/16 H.Arikawa Ver:1.1 Eee-Job V2.14より変更点を入れ込み

Option Explicit

#Const HSD200_USE = 1               'HSD200ボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

Public Enum CONTROL_STATUS
    INITIAL
    TEST_STEP
    TEST_CONTINUE
    TEST_REPEAT
    TEST_RETURN
    TEST_END
End Enum

Public Enum MEASURE_STATUS
    MEAS_INITIAL
    MEAS_STOP
    MEAS_RESTART
    MEAS_EXIT
End Enum

Public Enum status
    TEST_START
    FIRST_ACT
    RUNNING
    END_OF_ACT
    END_OF_TEST
End Enum

Private Const UPPER_VCLUMP_ICUL1G As Double = 6.5
Private Const LOWER_VCLUMP_ICUL1G As Double = -1.5
Private Const UPPER_VLIMIT_ICUL1G As Double = 6#
Private Const LOWER_VLIMIT_ICUL1G As Double = -1#

#If HSD200_USE = 0 Then
Private Const UPPER_VLIMIT_PPMU As Double = 7#
Private Const LOWER_VLIMIT_PPMU As Double = -2#
#Else
Private Const UPPER_VLIMIT_PPMU As Double = 6.5
Private Const LOWER_VLIMIT_PPMU As Double = -1.5
#End If
    
'#V21-Release
'############# 以下ユーティリティ群 ###############################################################
Public Sub CalculateTempValue(ByRef retValue() As Double, ByRef refValue As Variant, ByVal operateKey As String, ByVal container As CContainer, Optional ByVal site As Long = ALL_SITE)
'内容:
'   指定されたテンポラリ値との計算結果を返す
'
'[retValue()]   OUT Double型:       計算後のデータ配列
'[refValue]     OUT Variant型:      計算対象のデータ配列
'[operateKey]   IN String型:        計算式とテンポラリ変数名を表す文字列
'[Container]    IN CContainer型:    テンポラリ結果が格納されたコンテナ
'[Site]         In Long型:          サイト指定機能用　(Default:-1)
'注意事項:
'
'
    Dim dataIndex As Long
    Dim keyName As String
    Dim operator As String
    ReDim retValue(UBound(refValue))
    Dim TempValue() As Double
    operator = Left(operateKey, 1)
    keyName = Replace(operateKey, operator, "")
    On Error GoTo ErrHandler
    Select Case site
    Case ALL_SITE:
        Select Case operator
            Case "=":
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex)
                Next dataIndex
                container.AddTempResult keyName, retValue
            Case "+":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex) + TempValue(dataIndex)
                Next dataIndex
            Case "-":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex) - TempValue(dataIndex)
                Next dataIndex
            Case "*":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex) * TempValue(dataIndex)
                Next dataIndex
            Case "/":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex) / TempValue(dataIndex)
                Next dataIndex
            Case Else:
                For dataIndex = 0 To UBound(refValue)
                    retValue(dataIndex) = refValue(dataIndex)
                Next dataIndex
        End Select
    Case Else:  'Site指定
        dataIndex = site
        Select Case operator
            Case "=":
                retValue(dataIndex) = refValue(dataIndex)
                container.AddTempResult keyName, retValue
            Case "+":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                retValue(dataIndex) = refValue(dataIndex) + TempValue(dataIndex)
            Case "-":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                retValue(dataIndex) = refValue(dataIndex) - TempValue(dataIndex)
            Case "*":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                retValue(dataIndex) = refValue(dataIndex) * TempValue(dataIndex)
            Case "/":
                ReDim TempValue(UBound(refValue))
                container.TempResults.GetResult keyName, TempValue
                retValue(dataIndex) = refValue(dataIndex) / TempValue(dataIndex)
            Case Else:
                retValue(dataIndex) = refValue(dataIndex)
        End Select
    End Select
    Exit Sub
ErrHandler:
    For dataIndex = 0 To UBound(refValue)
        retValue(dataIndex) = refValue(dataIndex)
    Next dataIndex
    Err.Raise 9999, "CalculateTempValue", "Can Not Calculate The Operate [" & operateKey & "] !"
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

'############# 電源設定クラスの初期化 #############################################################
Public Function CreateVISConnector() As IDcTest
'内容：
'   DCテストシナリオを電源設定Objectを接続するための
'   接続用Objectの生成と初期設定
'
'パラメータ：
'
'戻り値：
'   IDcTestを実装した電源クラス接続用Object
'
'注意事項：
'

    '電源クラス接続用オブジェクト生成
    Dim VisLibObj As CVISConnectDcScenario
    Set VisLibObj = New CVISConnectDcScenario

    '電源設定OBJを設定
    Set VisLibObj.VISrcSelector = TheDC
    
    '##### スナップショット取得機能ON時の設定 #####
    If (IsSnapshotOn = True) And (TheExec.RunMode = runModeDebug) Then 'PMC実行時には動作しないように判定追加
                                                                                                                    
        '設定済みのスナップショットOBJをコレクションに登録
        '複数機能のスナップショットを取得することを考慮した設計のため
        'コレクションに取得したいスナップショット機能を追加する形をとる
        Dim snapFncList As Collection
        Set snapFncList = New Collection
        
        'Snapshot機能を追加
        Call snapFncList.Add(TheSnapshot)
        
        'スナップショットOBJリストに設定
        Set VisLibObj.SnapshotObjList = snapFncList
                
        'スナップショット機能実行フラグをTrue（実行する）に設定
        VisLibObj.CanUseSnapshot = True
    
        'スナップショット取得モードに設定されたことを表示
        Call MsgBox("Snapshot Save Mode !!", vbInformation, "XLibDcScenarioUtility")

    '##### スナップショット取得機能OFF時の設定 #####
    Else
        'スナップショット機能実行フラグをFalse（実行しない）に設定
        VisLibObj.CanUseSnapshot = False
    End If
    
    'すべての設定が終了した電源クラス接続用オブジェクトを返す
    Set CreateVISConnector = VisLibObj

End Function

'############# テスト環境系モジュール群 ###########################################################
Public Function GetInstanceName() As String
'内容:
'   カレントのテストインスタンス名を返す
'
'パラメータ:
'
'戻り値：
'   インスタンス名
'
'注意事項:
'   テストカテゴリ取得用

    GetInstanceName = TheExec.DataManager.InstanceName

End Function

Public Function GetInstansNameAsUCase() As String
'内容:
'   カレントのテストインスタンス名を大文字で返す
'
'パラメータ:
'
'戻り値：
'   インスタンス名（大文字）
'
'注意事項:
'   テストラベル取得用

    GetInstansNameAsUCase = UCase(TheExec.DataManager.InstanceName)

End Function

Public Function GetSiteCount() As Long
'内容:
'   サイト数の取得
'
'戻り値：
'   サイト数からマイナス1をした数値
'
'注意事項:
'

    GetSiteCount = TheExec.sites.ExistingCount - 1
End Function

Public Function GetTesterNum() As Long
'内容:
'   Sw_Nodeパブリック変数のラッピング
'
'戻り値：
'   テスタ番号
'
'注意事項:
'

    GetTesterNum = Sw_Node
End Function

Public Sub CreateSiteArray(ByRef retArray() As Double)
'内容:
'   配列をサイト数分確保する
'
'戻り値：
'   サイト数分確保された配列変数
'
'注意事項:
'

    ReDim retArray(GetSiteCount)
End Sub

Public Function IsGangPins(ByVal PinList As String) As Boolean
'内容:
'   指定されたピンリストにギャングピンが含まれているか確認
'   IsGangPinListのラッパー
'
'パラメータ:
'    [PinList]       In   確認を行うPinList
'
'戻り値:
'   確認結果（ギャングピンが含まれている=True）
'
'注意事項:
'
    If TheDC.Pins(PinList).BoardName = "dcICUL1G" Then
        IsGangPins = False
    Else
        IsGangPins = IsGangPinlist(PinList, GetChanType(PinList))
    End If
End Function

Public Function GetGangPinCount(ByVal PinList As String) As Long
'内容:
'   ギャングピン数を返す
'
'パラメータ:
'    [PinList]       In   確認を行うPinList
'
'戻り値:
'   ギャングのピン数
'
'注意事項:
'
    Dim pinArr() As String
    TheExec.DataManager.DecomposePinList PinList, pinArr, GetGangPinCount
End Function

Public Function ValidateMeasureRange(ByVal mPin As CMeasurePin) As Long
'内容:
'   測定レンジが最適かどうかの判定を行う
'
'パラメータ:
'    [mPin]       In   判定対象のPinオブジェクト
'
'戻り値:
'   判定結果（Constant型）
'
'注意事項:
'
    With mPin
        '### ピンのレンジ判定が不可能な場合 ###############
        If .TestLabel = NOT_DEFINE Or .BoardRange = INVALIDATION_VALUE Then
            ValidateMeasureRange = DISABEL_TO_VALIDATION
            Exit Function
        End If
        '### 判定パラメータが0または3の場合 ###############
        If .JudgeNumber = 0 Or .JudgeNumber = 3 Then
            '### PPMUのMVモードのための判定ロジック #######
            'レンジは HSD100:-2～7V, HSD200:-1.5V～6.5V のみなので最適値の判定はしない
            If .BoardName = "dcPPMU" And GetUnit(.Unit) = "V" Then
                If Not isCorrectLimitForPPMU(.UpperLimit, .LowerLimit) Then
                    If .JudgeNumber = 3 Then
                        ValidateMeasureRange = VALIDATE_NG
                    ElseIf .JudgeNumber = 0 Then
                        ValidateMeasureRange = VALIDATE_NG_NO_JUDGE
                    End If
                Else
                    If .JudgeNumber = 3 Then
                        ValidateMeasureRange = VALIDATE_OK
                    ElseIf .JudgeNumber = 0 Then
                        ValidateMeasureRange = VALIDATE_OK_NO_JUDGE
                    End If
                End If
            '### ICLU1GのMVモードのための判定ロジック #######
            'レンジは -1V～6V のみなので最適値の判定はしない
            ElseIf .BoardName = "dcICUL1G" And GetUnit(.Unit) = "V" Then
                If Not isCorrectLimitForICUL1G(.UpperLimit, .LowerLimit) Then
                    If .JudgeNumber = 3 Then
                        ValidateMeasureRange = VALIDATE_NG
                    ElseIf .JudgeNumber = 0 Then
                        ValidateMeasureRange = VALIDATE_NG_NO_JUDGE
                    End If
                Else
                    If .JudgeNumber = 3 Then
                        ValidateMeasureRange = VALIDATE_OK
                    ElseIf .JudgeNumber = 0 Then
                        ValidateMeasureRange = VALIDATE_OK_NO_JUDGE
                    End If
                End If
            '### その他の判定ロジック #####################
            'レンジの最適値まで含めた判定を行う
            Else
                Dim maxLimit As Double
                maxLimit = absMaxLimit(.UpperLimit, .LowerLimit)
                If maxLimit >= RoundDownDblData(.BoardRange, DIGIT_NUMBER) Then
                    If .JudgeNumber = 3 Then
                        ValidateMeasureRange = VALIDATE_NG
                    ElseIf .JudgeNumber = 0 Then
                        ValidateMeasureRange = VALIDATE_NG_NO_JUDGE
                    End If
                Else
                    '### 最適レンジ値かどうかの判定 #######
                    Dim isOptimalRange As Boolean
                    Select Case .BoardName
                        Case "dcAPMU":
                            isOptimalRange = isOptimalRangeAPMU(.Name, maxLimit, .BoardRange, .Unit)
                        Case "dcPPMU":
                            isOptimalRange = isOptimalRangePPMU(maxLimit, .BoardRange, .Unit)
                        Case "dcBPMU":
                            isOptimalRange = isOptimalRangeBPMU(maxLimit, .BoardRange, .Unit)
                        Case "dcDPS":
                            isOptimalRange = isOptimalRangeDPS(maxLimit, .BoardRange, .Unit)
                        Case "dcHDVIS":
                            isOptimalRange = isOptimalRangeHDVIS(.Name, maxLimit, .BoardRange, .Unit)
                        Case "dcICUL1G":
                            isOptimalRange = isOptimalRangeICUL1G(maxLimit, .BoardRange, .Unit)
                    End Select
                    If isOptimalRange Then
                        If .JudgeNumber = 3 Then
                            ValidateMeasureRange = VALIDATE_OK
                        ElseIf .JudgeNumber = 0 Then
                            ValidateMeasureRange = VALIDATE_OK_NO_JUDGE
                        End If
                    Else
                        If .JudgeNumber = 3 Then
                            ValidateMeasureRange = VALIDATE_WARNING
                        ElseIf .JudgeNumber = 0 Then
                            ValidateMeasureRange = VALIDATE_WARNING_NO_JUDGE
                        End If
                    End If
                End If
            End If
        '### 判定パラメータが0と3以外はノーチェック #######
        Else
            ValidateMeasureRange = NO_JUDGE
        End If

    End With
End Function

Private Function isCorrectLimitForPPMU(ByVal HiLimit As Double, ByVal LoLimit As Double) As Boolean
    Dim roundHLimit As Double
    Dim roundLLimit As Double
    roundHLimit = RoundDownDblData(HiLimit, DIGIT_NUMBER)
    roundLLimit = RoundDownDblData(LoLimit, DIGIT_NUMBER)
    If UPPER_VLIMIT_PPMU > roundHLimit And LOWER_VLIMIT_PPMU < roundLLimit Then
        isCorrectLimitForPPMU = True
    Else
        isCorrectLimitForPPMU = False
    End If
End Function

Private Function isCorrectLimitForICUL1G(ByVal HiLimit As Double, ByVal LoLimit As Double) As Boolean
    Dim roundHLimit As Double
    Dim roundLLimit As Double
    roundHLimit = RoundDownDblData(HiLimit, DIGIT_NUMBER)
    roundLLimit = RoundDownDblData(LoLimit, DIGIT_NUMBER)
    If UPPER_VLIMIT_ICUL1G > roundHLimit And LOWER_VLIMIT_ICUL1G < roundLLimit Then
        isCorrectLimitForICUL1G = True
    Else
        isCorrectLimitForICUL1G = False
    End If
End Function

Private Function absMaxLimit(ByVal HiLimit As Double, ByVal LoLimit As Double) As Double
    Dim roundHLimit As Double
    Dim roundLLimit As Double
    roundHLimit = RoundDownDblData(HiLimit, DIGIT_NUMBER)
    roundLLimit = RoundDownDblData(LoLimit, DIGIT_NUMBER)
    If Abs(roundHLimit) > Abs(roundLLimit) Then
        absMaxLimit = Abs(roundHLimit)
    Else
        absMaxLimit = Abs(roundLLimit)
    End If
End Function

Private Function isOptimalRangeAPMU(ByVal pName As String, ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            If maxLimit < 2 Then
                isOptimalRangeAPMU = CompareDblData(rangeVal, 2, DIGIT_NUMBER)
            ElseIf maxLimit < 5 Then
                isOptimalRangeAPMU = CompareDblData(rangeVal, 5, DIGIT_NUMBER)
            ElseIf maxLimit < 10 Then
                isOptimalRangeAPMU = CompareDblData(rangeVal, 10, DIGIT_NUMBER)
            ElseIf maxLimit < 35 Then
                isOptimalRangeAPMU = CompareDblData(rangeVal, 35, DIGIT_NUMBER)
            End If
        Case "A":
            If IsGangPins(pName) Then
                'ギャングピンは常にWarningを返す仕様に変更 '08/04/04
'                isOptimalRangeAPMU = CompareDblData(rangeVal, 0.05 * GetGangPinCount(pName), DIGIT_NUMBER)
                isOptimalRangeAPMU = False
            Else
                If maxLimit < 0.0000002 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.0000002, DIGIT_NUMBER)
                ElseIf maxLimit < 0.000002 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.000002, DIGIT_NUMBER)
                ElseIf maxLimit < 0.00001 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.00001, DIGIT_NUMBER)
                ElseIf maxLimit < 0.00004 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.00004, DIGIT_NUMBER)
                ElseIf maxLimit < 0.0002 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.0002, DIGIT_NUMBER)
                ElseIf maxLimit < 0.001 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.001, DIGIT_NUMBER)
                ElseIf maxLimit < 0.005 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.005, DIGIT_NUMBER)
                ElseIf maxLimit < 0.05 Then
                    isOptimalRangeAPMU = CompareDblData(rangeVal, 0.05, DIGIT_NUMBER)
                End If
            End If
    End Select
End Function

Private Function isOptimalRangePPMU(ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            'MVモード時は常にTrueを返す仕様に変更 '08/04/24
'            isOptimalRangePPMU = CompareDblData(rangeVal, 2, DIGIT_NUMBER)
            isOptimalRangePPMU = True
        Case "A":
            If maxLimit < 0.0000002 Then
                #If HSD200_USE = 0 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.0000002, DIGIT_NUMBER)
                #Else
                'HSD200にはレンジ200nAはない
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.000002, DIGIT_NUMBER)
                #End If
            ElseIf maxLimit < 0.000002 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.000002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.00002 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.00002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.0002 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.0002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.002 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.002, DIGIT_NUMBER)
            #If HSD200_USE <> 0 Then
            ElseIf maxLimit < 0.05 Then
                isOptimalRangePPMU = CompareDblData(rangeVal, 0.05, DIGIT_NUMBER)
            #End If
            End If
    End Select
End Function

Private Function isOptimalRangeICUL1G(ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            'MVモード時は常にTrueを返す
            isOptimalRangeICUL1G = True
        Case "A":
            If maxLimit < 0.00002 Then
                isOptimalRangeICUL1G = CompareDblData(rangeVal, 0.00002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.0002 Then
                isOptimalRangeICUL1G = CompareDblData(rangeVal, 0.0002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.002 Then
                isOptimalRangeICUL1G = CompareDblData(rangeVal, 0.002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.01 Then
                isOptimalRangeICUL1G = CompareDblData(rangeVal, 0.01, DIGIT_NUMBER)
            End If
    End Select
End Function

Private Function isOptimalRangeBPMU(ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            If maxLimit < 2 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 2, DIGIT_NUMBER)
            ElseIf maxLimit < 5 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 5, DIGIT_NUMBER)
            ElseIf maxLimit < 10 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 10, DIGIT_NUMBER)
            ElseIf maxLimit < 24 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 24, DIGIT_NUMBER)
            End If
        Case "A":
            If maxLimit < 0.000002 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.000002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.00002 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.00002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.0002 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.0002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.002 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.002, DIGIT_NUMBER)
            ElseIf maxLimit < 0.02 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.02, DIGIT_NUMBER)
            ElseIf maxLimit < 0.2 Then
                isOptimalRangeBPMU = CompareDblData(rangeVal, 0.2, DIGIT_NUMBER)
            End If
    End Select
End Function

Private Function isOptimalRangeDPS(ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            Call OutputErrMsg("DPS is not support FI mode")
        Case "A":
            If maxLimit < 0.00005 Then
                isOptimalRangeDPS = CompareDblData(rangeVal, 0.00005, DIGIT_NUMBER)
            ElseIf maxLimit < 0.0005 Then
                isOptimalRangeDPS = CompareDblData(rangeVal, 0.0005, DIGIT_NUMBER)
            ElseIf maxLimit < 0.01 Then
                isOptimalRangeDPS = CompareDblData(rangeVal, 0.01, DIGIT_NUMBER)
            ElseIf maxLimit < 0.1 Then
                isOptimalRangeDPS = CompareDblData(rangeVal, 0.1, DIGIT_NUMBER)
            ElseIf maxLimit < 1 Then
                isOptimalRangeDPS = CompareDblData(rangeVal, 1, DIGIT_NUMBER)
            End If
    End Select
End Function

Private Function isOptimalRangeHDVIS(ByVal pName As String, ByVal maxLimit As Double, ByVal rangeVal As Double, ByVal dataUnit As String) As Boolean
    Select Case GetUnit(dataUnit)
        Case "V":
            isOptimalRangeHDVIS = CompareDblData(rangeVal, 10, DIGIT_NUMBER)
        Case "A":
            Dim pinCount As Long
            pinCount = GetGangPinCount(pName)
            If maxLimit < RoundDownDblData(0.000005 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.000005 * pinCount, DIGIT_NUMBER)
            ElseIf maxLimit < RoundDownDblData(0.00005 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.00005 * pinCount, DIGIT_NUMBER)
            ElseIf maxLimit < RoundDownDblData(0.0005 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.0005 * pinCount, DIGIT_NUMBER)
            ElseIf maxLimit < RoundDownDblData(0.005 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.005 * pinCount, DIGIT_NUMBER)
            ElseIf maxLimit < RoundDownDblData(0.05 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.05 * pinCount, DIGIT_NUMBER)
            ElseIf maxLimit < RoundDownDblData(0.2 * pinCount, DIGIT_NUMBER) Then
                isOptimalRangeHDVIS = CompareDblData(rangeVal, 0.2 * pinCount, DIGIT_NUMBER)
            End If
    End Select
End Function

'############# シートパラメータ入力支援用ユーティリティ関数群 #####################################
Public Sub CreateActionParameterList(ByVal selCell As Range)
'内容:
'   DCシナリオワークシートにおける入力支援マクロ関数①
'   ユーザーのためのデータリスト作成
'
'パラメータ:
'   [selCell]      In   対象セルオブジェクト
'
'注意事項:
'
    If Not IsJobValid Then Exit Sub
    '### データラベルに応じたデータリストの作成 ###########
    Dim dataList As Collection
    Set dataList = Nothing
    On Error GoTo LIST_ERROR
    With selCell.parent
    '### カテゴリパラメータのリスト作成 ###################
    If IsCategoryHeader(selCell) Then
        '### 実験フラグリスト作成 #########################
        If selCell.Column = .Range(EXAMIN_FLAG).Column Then
            Set dataList = examinFlagList
        '### 実験モードリスト作成 #########################
        ElseIf selCell.Column = .Range(EXAMIN_MODE).Column Then
            If targetCell(selCell, EXAMIN_FLAG).Value Then
                Set dataList = examinModeList
            End If
        End If
    End If
    '### ピングループパラメータのリスト作成 ###############
    If IsGroupHeader(selCell) Then
        '### ピンリストの取得 #############################
        Dim PinList As String
        PinList = createPinList(selCell, False)
        '### ピンリストが無効の場合はリスト作成はしない ###
        If PinList = NOT_DEFINE Then GoTo LIST_ERROR
        Dim boardType As String
        boardType = targetCell(selCell, TEST_PINTYPE).Value
        '### アクションリスト作成 #########################
        If selCell.Column = .Range(TEST_ACTION).Column Then
            Set dataList = actionList
        '### ピンタイプ（I/Oピンのみ）リスト作成 ##########
        ElseIf selCell.Column = .Range(TEST_PINTYPE).Column Then
            Set dataList = ioPinTypeList(PinList)
        End If
        '### 以下のパラメータは有効であればリスト作成する #
        Dim ParamList As Collection
        Set ParamList = actionParamList(selCell)
        '### 測定モードリスト作成 #########################
        If selCell.Column = .Range(SET_MODE).Column And _
               IsEnableParameter(ParamList, SET_MODE) Then
            Set dataList = modeList(PinList)
        '### 測定レンジリスト作成 #########################
        ElseIf selCell.Column = .Range(SET_RANGE).Column And _
               IsEnableParameter(ParamList, SET_RANGE) Then
            Dim measureMode As String
            measureMode = targetCell(selCell, SET_MODE).Value
            Set dataList = enableBoardRange(PinList, boardType, measureMode)
        '### サイト測定モードリスト作成 ###################
        ElseIf selCell.Column = .Range(MEASURE_SITE).Column And _
               IsEnableParameter(ParamList, MEASURE_SITE) Then
            Set dataList = measureSiteList(PinList, boardType)
        End If
    '### ポストアクション用パラメータのリスト作成 #########
    ElseIf IsGroupFooter(selCell) Then
        If selCell.Column = .Range(TEST_ACTION).Column Then
            Set dataList = postActionList
        End If
    End If
    End With
LIST_ERROR:
    '### データリストをリストボックスに設定 ###############
    CreateListBox selCell, dataList
End Sub

Public Sub ValidateActionParameter(ByVal chCell As Range)
'内容:
'   DCシナリオワークシートにおける入力支援マクロ関数②
'   有効なパラメータラベルの書式チェック及びデータチェック
'
'パラメータ:
'   [chCell]      In   対象セルオブジェクト
'
'注意事項:
'
    If Not IsJobValid Then Exit Sub
    On Error GoTo DATA_ERROR
    If IsEmpty(chCell) Then enableCell chCell
    '### カテゴリパラメータのパラメータチェック ###########
    If IsCategoryHeader(chCell) Then
        '現在はカテゴリパラメータのチェックは特になし
    End If
    '### ピングループパラメータのチェック #################
    If IsGroupHeader(chCell) Then
        '### ピングループのチェック #######################
        Dim PinList As String
        PinList = createPinList(chCell, True)
        '### ピングループ先頭ピンのボード名更新 ###########
        updateBoardName chCell
        'ピンリストが無効の場合はチェックはしない
        If PinList = NOT_DEFINE Then Exit Sub
        '### アクションパラメータリストの取得 #############
        Dim ParamList As Collection
        Set ParamList = actionParamList(chCell)
        '### アクショングループフォーマットの整形 #########
        actionParamFormatter targetCell(chCell, TEST_ACTION), ParamList
        '無効なアクションの場合は以下のパラメータチェックはしない
        If ParamList.Count = 0 Then Exit Sub
        '### 測定モードパラメータチェック #################
        Dim dataList As Collection
        If IsEnableParameter(ParamList, SET_MODE) Then
            Set dataList = modeList(PinList)
            verifyListParamData targetCell(chCell, SET_MODE), dataList
        End If
        '### サイトモードパラメータチェック ###############
        If IsEnableParameter(ParamList, MEASURE_SITE) Then
            Dim boardType As String
            boardType = targetCell(chCell, TEST_PINTYPE).Value
            Set dataList = measureSiteList(PinList, boardType)
            verifyListParamData targetCell(chCell, MEASURE_SITE), dataList
        End If
        '### 測定レンジパラメータチェック #################
        If IsEnableParameter(ParamList, SET_RANGE) Then
            verifyRangeData targetCell(chCell, SET_RANGE), PinList
        End If
        '### 印加パラメータチェック #######################
        If IsEnableParameter(ParamList, SET_FORCE) Then
            verifyForceValue targetCell(chCell, SET_FORCE), PinList
        End If
        '### ウェイトパラメータチェック ###################
        If IsEnableParameter(ParamList, MEASURE_WAIT) Then
            verifyWaitValue targetCell(chCell, MEASURE_WAIT)
        End If
        '### 平均回数パラメータチェック ###################
        If IsEnableParameter(ParamList, MEASURE_AVG) Then
            verifyAverageValue targetCell(chCell, MEASURE_AVG)
        End If
    Else
        '### ピングループのチェック #######################
        If chCell.Column = chCell.parent.Range(TEST_PINS).Column Then
            createPinList chCell, True
        End If
    End If
    '### ポストアクション用パラメータチェック #############
    If IsGroupFooter(chCell) Then
        enableCell targetCell(chCell, TEST_ACTION)
        verifyListParamData targetCell(chCell, TEST_ACTION), postActionList
    End If
    Exit Sub
DATA_ERROR:
End Sub

Private Sub updateBoardName(ByVal chCell As Range)
    '### グループ先頭ピンのボード名更新 ###################
    '前提条件：対象セルがピングループの先頭であること
    '08/05/12 OK
    '### ボード名の取得 ###################################
    Dim boardType As String
    boardType = queryBoardName(targetCell(chCell, TEST_PINS).Value)
    '### 現在の設定値取得 #################################
    Dim typeCell As Range
    Set typeCell = targetCell(chCell, TEST_PINTYPE)
    '### ボード名の設定更新 ###############################
    If boardType <> typeCell.Value Then
        If boardType <> "PPMU" Or typeCell.Value <> "BPMU" Then
            Application.EnableEvents = False
            typeCell.Value = boardType
            Application.EnableEvents = True
        End If
    End If
End Sub

Private Sub verifyRangeData(ByVal chCell As Range, ByVal PinList As String)
    '### 入力測定レンジがリストに存在するかどうかの判定 ###
    '前提条件：対象セルがピングループの先頭であること
    'パラメータが書式・スペックを満たさない場合、パラメータクラスでエラーが起こる
    '08/05/09 OK
    On Error GoTo DATA_ERROR
    enableCell chCell
    Dim boardType As String
    boardType = targetCell(chCell, TEST_PINTYPE).Value
    Dim measureMode As String
    measureMode = targetCell(chCell, SET_MODE).Value
    '### レンジリストの取得 ###############################
    Dim rangeList As Collection
    Set rangeList = enableBoardRange(PinList, boardType, measureMode)
    '### リストにNoneが含まれている場合は他の入力値はNG ###
    If IsEnableParameter(rangeList, "None") And Not chCell.Value = "None" Then
        If boardType = "PPMU" Then
            GoTo DATA_ERROR
        ElseIf boardType = "ICUL1G" Then
            'Continue
        Else
            On Error GoTo 0
            Call Err.Raise(9999, "XLibDcScenarioUtility.verifyRangeData", "Internal error!")
            On Error GoTo DATA_ERROR
        End If
    End If
    '### 入力値がリストに含まれていなければワーニング書式設定
    'ピンリストがN.Dの時はIsEnableParameterでFlaseとなる
    If Not IsEnableParameter(rangeList, chCell.Value) Then
        '### 入力値の書式チェック #########################
        Dim paramRange As CParamStringWithUnit
        Set paramRange = CreateCParamStringWithUnit
        Select Case measureMode
            Case "MI":
                paramRange.Initialize "A"
            Case "MV":
                paramRange.Initialize "V"
            Case Else:
                GoTo DATA_ERROR
        End Select
        With paramRange.AsIParameter
            If boardType = "ICUL1G" And measureMode = "MV" Then
                If chCell.Value = Empty Then
                    chCell.Value = UPPER_VCLUMP_ICUL1G & "V"
                Else
                    'ICUL1Gに限り上側クランプ値判定を行う
                    '（-1.5V<value<=6.5V）
                    .UpperLimit = UPPER_VCLUMP_ICUL1G
                    .LowerLimit = LOWER_VCLUMP_ICUL1G
                    .AsString = chCell.Value
                    If CompareDblData(.AsDouble, LOWER_VCLUMP_ICUL1G, DIGIT_NUMBER) Then
                        Err.Raise 9999, "XLibDcScenarioUtility.verifyRangeData", "ICUL1G Upper Clamp must be > " & LOWER_VCLUMP_ICUL1G
                    End If
                End If
            Else
                .LowerLimit = 0
                .AsString = chCell.Value
            End If
        End With
        warningCell chCell
    End If
    Exit Sub
DATA_ERROR:
    errorCell chCell
End Sub

Private Sub verifyForceValue(ByVal chCell As Range, ByVal PinList As String)
    '### 入力印加パラメータの書式・スペックの判定 #########
    '前提条件：対象セルがピングループの先頭であること
    'ピンリストがN.Dの時はcompareForceLimitでFlaseとなる
    'パラメータが書式・スペックを満たさない場合、パラメータクラスでエラーが起こる
    '08/05/12 OK
    On Error GoTo DATA_ERROR
    enableCell chCell
    If IsEmpty(chCell) Then Exit Sub
    Dim boardType As String
    boardType = targetCell(chCell, TEST_PINTYPE).Value
    Dim measureMode As String
    measureMode = targetCell(chCell, SET_MODE).Value
    '### 入力値の書式チェック #############################
    Dim paramForce As CParamStringWithUnit
    Set paramForce = CreateCParamStringWithUnit
    '### 単位付文字列でない場合 ###########################
    If Not IsEmpty(targetCell(chCell, OPERATE_FORCE)) Then
        With paramForce
            .Initialize ""
            .AsIParameter.AsDouble = chCell.Value
        End With
    '### 単位付文字列の場合 ###############################
    Else
        Select Case measureMode
            Case "MI":
                paramForce.Initialize "V"
            Case "MV":
                paramForce.Initialize "A"
            Case Else:
                GoTo DATA_ERROR
        End Select
        paramForce.AsIParameter.AsString = chCell.Value
        '### 入力値がスペック内でなければエラー書式設定 ###
        If Not compareForceLimit(PinList, boardType, measureMode, paramForce.AsIParameter.AsDouble) Then
            errorCell chCell
        End If
    End If
    Exit Sub
DATA_ERROR:
    errorCell chCell
End Sub

Private Sub verifyWaitValue(ByVal chCell As Range)
    '### ウェイト入力値の書式チェック #####################
    '前提条件：対象セルがピングループの先頭であること
    'パラメータが書式・スペックを満たさない場合、パラメータクラスでエラーが起こる
    '08/05/12 OK
    On Error GoTo DATA_ERROR
    enableCell chCell
    Dim paramWait As CParamStringWithUnit
    Set paramWait = CreateCParamStringWithUnit
    With paramWait
        .Initialize "S"
        With .AsIParameter
            .LowerLimit = 0
            .AsString = chCell.Value
        End With
    End With
    Exit Sub
DATA_ERROR:
    errorCell chCell
End Sub

Private Sub verifyAverageValue(ByVal chCell As Range)
    '### アベレージ入力値の書式チェック ###################
    '前提条件：対象セルがピングループの先頭であること
    'パラメータが書式・スペックを満たさない場合、パラメータクラスでエラーが起こる
    '08/05/12 OK
    On Error GoTo DATA_ERROR
    enableCell chCell
    Dim paramAvg As CParamLong
    Set paramAvg = CreateCParamLong
    With paramAvg.AsIParameter
        .LowerLimit = 1
        .AsLong = chCell.Value
    End With
    Exit Sub
DATA_ERROR:
    errorCell chCell
End Sub

Private Sub verifyListParamData(ByVal chCell As Range, ByVal ParamList As Collection)
    '### その他リスト入力値のチェック #####################
    'リストに存在しないパラメータはNGとする
    'リストが存在しない場合はチェックしない
    '08/05/12 OK
    If ParamList Is Nothing Then Exit Sub
    If Not IsEnableParameter(ParamList, chCell.Value) Then
        errorCell chCell
    End If
End Sub

Private Sub verifyTestNameOrLabel(ByVal chCell As Range)
    '### テストカテゴリ・テストラベルの定義の有無を検証 ###
    '未検証・未使用
    On Error GoTo DATA_ERROR
    enableCell chCell
    '### アクティブデータシートの取得 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(JOB_LIST_TOOL, TEST_INSTANCES_TOOL)
    If DataSheet Is Nothing Then GoTo DATA_ERROR
    '### インスタンスシートリーダーの作成 #################
    Dim instanceReader As CInstanceSheetReader
    Set instanceReader = CreateCInstanceSheetReader
    With instanceReader
        .Initialize DataSheet.Name
        .AsIFileStream.SetLocation chCell.Value
    End With
    Exit Sub
DATA_ERROR:
    warningCell chCell
End Sub

Private Function actionList() As Collection
    '### アクションパラメータリスト #######################
    '現状は存在するアクションの問い合わせ先がないので直にリスト作成
    Set actionList = New Collection
    With actionList
        .Add ""
        .Add "SET"
        .Add "MEASURE"
        .Add "OPEN"
    End With
End Function

Private Function postActionList() As Collection
    '### ポストアクションパラメータリスト #################
    '現状は存在するアクションの問い合わせ先がないので直にリスト作成
    Set postActionList = New Collection
    With postActionList
        .Add ""
        .Add "OPEN"
    End With
End Function

Private Function ioPinTypeList(ByVal PinList As String) As Collection
    '### ボード名リスト ###################################
    '現状はPPMU/BPMUの問い合わせ先がないのでリスト作成
    Select Case queryBoardName(PinList)
        Case "PPMU":
            Set ioPinTypeList = New Collection
            With ioPinTypeList
                .Add "PPMU"
                .Add "BPMU"
            End With
        Case Else
            Set ioPinTypeList = Nothing
    End Select
End Function
'V21-Release
Private Function measureSiteList(ByVal PinList As String, ByVal boardType As String) As Collection
    '### メージャーサイトモードリスト #####################
    '現状はメージャーサイトモードの問い合わせ先がないので直にリスト作成
    'Measure ActionのBPMUはここでConcurrentにしてもアクションのセット時に強制的にSerialなので問題なし

    Set measureSiteList = New Collection
    With measureSiteList
        If IsGangPins(PinList) Then
            .Add "Serial"
        Else
            .Add "Concurrent"
            .Add "Serial"
            'site指定用
            Dim Num As Integer
            Num = GetSiteCount
            Dim i As Integer
            For i = 0 To Num
                .Add i
            Next i
        End If
    End With
    
End Function

Private Function modeList(ByVal PinList As String) As Collection
    '### メージャーモードリスト ###########################
    '現状はメージャーモードの問い合わせ先がないので直にリスト作成
    Set modeList = New Collection
    With modeList
        Select Case queryBoardName(PinList)
            Case "DPS":
                .Add "MI"
            Case Else
                .Add "MV"
                .Add "MI"
        End Select
    End With
End Function

Private Function examinFlagList() As Collection
    '### 実験フラグリスト #################################
    Set examinFlagList = New Collection
    With examinFlagList
        .Add "FALSE"
        .Add "TRUE"
    End With
End Function

Private Function examinModeList() As Collection
    '### 実験モードリスト #################################
    Set examinModeList = New Collection
    With examinModeList
        .Add ""
        .Add "BREAK"
        .Add "END"
    End With
End Function

Private Sub actionParamFormatter(ByVal chCell As Range, ByVal ParamList As Collection)
    '### カレントアクションのパラメータ収集 ###############
    '前提条件：対象セルがピングループの先頭であること

    '=== Add Eee-Job V2.14 ===
    Dim ErrorCells As Range
    Dim EnableCells As Range
    Dim DisableCells As Range
    '=== Add Eee-Job V2.14 ===

    Dim currCell As Range
    Set currCell = targetCell(chCell, TEST_ACTION)
    'EnableCell currCell                        'Add Eee-Job V2.14
    MakeUnionRange EnableCells, currCell        'Add Eee-Job V2.14

    '### パラメータが存在しないアクションはNG #############
    If ParamList Is Nothing Or ParamList.Count = 0 Then
        'errorCell currCell                     'Add Eee-Job V2.14
        MakeUnionRange ErrorCells, currCell     'Add Eee-Job V2.14
    End If
    '### シート上データラベルの準備 #######################
    Dim shtDataList As New Collection
    With shtDataList
        .Add SET_MODE
        .Add SET_RANGE
        .Add SET_FORCE
        .Add MEASURE_WAIT
        .Add MEASURE_AVG
        .Add MEASURE_SITE
        .Add MEASURE_LABEL
        .Add OPERATE_FORCE
        .Add OPERATE_RESULT
    End With
    '### ピングループセルの収集 ###########################
    Dim pinCells As Collection
    Set pinCells = groupCellList(currCell, False)
    '### パラメータフォーマットの設定 #####################
    Dim currParam As Variant
    Dim currData As Collection
    Dim currPin As Range
    For Each currParam In shtDataList
        '### フォーマット対象セルの設定 ###################
        If currParam = MEASURE_LABEL Or currParam = OPERATE_RESULT Then
            Set currData = pinCells
        Else
            Set currData = New Collection
            currData.Add pinCells.Item(1)
        End If
        '### フォーマットの順次設定 #######################
        For Each currPin In currData
            If IsEnableParameter(ParamList, currParam) Then
                'EnableCell targetCell(currPin, currParam)                      'Add Eee-Job V2.14
                MakeUnionRange EnableCells, targetCell(currPin, currParam)      'Add Eee-Job V2.14
            Else
                'disableCell targetCell(currPin, currParam)                     'Add Eee-Job V2.14
                MakeUnionRange DisableCells, targetCell(currPin, currParam)     'Add Eee-Job V2.14
            End If
        Next currPin
    Next currParam
    '=== Add Eee-Job V2.14 ===
    '書式一括設定
    If Not EnableCells Is Nothing Then
        enableCell EnableCells
    End If
    If Not ErrorCells Is Nothing Then
        errorCell ErrorCells
    End If
    If Not DisableCells Is Nothing Then
        disableCell DisableCells
    End If
     '=== Add Eee-Job V2.14 ===
End Sub
'#V21-Release
Private Function actionParamList(ByVal currCell As Range) As Collection
    '### アクションに必要なデータラベルのリストを作成する
    '前提条件：対象セルがピングループの先頭であること
    '08/05/09 OK
    '10/08/20 BPMU制限解除+エラーセル　バグ修正
    
    Set actionParamList = New Collection
    Dim currAction As Range
    Set currAction = targetCell(currCell, TEST_ACTION)
    Dim actions As New Collection
    If currAction.Value = "SET" Or IsEmpty(currAction) Then
        Dim setAct As New CSetFI
        actions.Add setAct
        actionParamList.Add SET_MODE
    End If
    If currAction.Value = "MEASURE" Or IsEmpty(currAction) Then
        Dim measAct As New CMeasureI
        Dim measPin As New CMeasurePin
        With actions
            .Add measAct
            .Add measPin
        End With
        actionParamList.Add SET_MODE
    End If
    If currAction.Value = "OPEN" Then
        Dim disconAct As New CDisconnect
        actions.Add disconAct
    End If
    actionParamList.Add MEASURE_SITE
    addActionParamList actionParamList, actions
    
    
End Function

Private Function IsEnableParameter(ByVal ParamList As Collection, ByVal dataLabel As String) As Boolean
    '### パラメータリストに任意のデータラベルが存在するかどうかを判定する
    '08/05/09 OK
    If ParamList Is Nothing Then Exit Function
    Dim currParam As Variant
    For Each currParam In ParamList
        If currParam = dataLabel Then
            IsEnableParameter = True
            Exit Function
        End If
    Next currParam
End Function

Private Function addActionParamList(ByVal ParamList As Collection, ByVal actions As Collection)
    '### 登録されたアクションからパラメータを読み込む #####
    '08/05/09 OK
    Dim currAct As IParameterWritable
    Dim currParam As Variant
    For Each currAct In actions
        For Each currParam In currAct.ParameterList
            ParamList.Add currParam
        Next currParam
    Next currAct
End Function

Private Function compareForceLimit(ByVal PinList As String, ByVal boardType As String, ByVal measureMode As String, ByVal inputVal As Double) As Boolean
    '### 入力データと印加限界値の比較を行う '08/04/30 OK ##
    'queryBoardSpecから無効な値が返された場合はFalse
    Dim limitVal() As Double
    compareForceLimit = False
    Select Case boardType
        Case "BPMU":
            limitVal = queryBoardSpecForBPMU(PinList, measureMode)
        Case Else:
            limitVal = queryBoardSpec(PinList, measureMode)
    End Select
    If limitVal(0) = INVALIDATION_VALUE And limitVal(1) = INVALIDATION_VALUE Then GoTo IS_INVALID
    Dim roundHLimit As Double
    Dim roundLLimit As Double
    Dim roundFVal As Double
    roundHLimit = RoundDownDblData(limitVal(1), DIGIT_NUMBER)
    roundLLimit = RoundDownDblData(limitVal(0), DIGIT_NUMBER)
    roundFVal = RoundDownDblData(inputVal, DIGIT_NUMBER)
    If roundLLimit <= roundFVal And roundFVal <= roundHLimit Then
        compareForceLimit = True
    End If
    Exit Function
IS_INVALID:
End Function

Private Function enableBoardRange(ByVal PinList As String, ByVal boardType As String, ByVal measureMode As String) As Collection
    '### ピンリストから有効なレンジリストを取得する #######
    '①BPMU以外はピン名からリソースを自動判定するのでboardTypeの引数は無効になる
    '②BPMU指定はリソースの固有値を返すのでPinListの引数は無効になる
    '③ギャングピンのリストは最大レンジ*ギャング数のみ
    '08/05/09 OK
    Select Case boardType
        Case "BPMU":
            Set enableBoardRange = queryBoardRangeForBPMU(PinList, measureMode)
        Case Else
            Dim tempRange As Collection
            Set tempRange = queryBoardRange(PinList, measureMode)
            If tempRange Is Nothing Then
                Set enableBoardRange = Nothing
                Exit Function
            End If
            If IsGangPins(PinList) And measureMode = "MI" Then
                Set enableBoardRange = New Collection
                Dim maxRange As Variant
                'HDVISは200mA決め打ち：次フェーズで修正が必要
                Select Case boardType
                    Case "HDVIS":
                        maxRange = "200mA"
                    Case "APMU"
                        maxRange = tempRange.Item(1)
                End Select
                Dim MainUnit As String
                Dim SubUnit As String
                Dim SubValue As Double
                SplitUnitValue maxRange, MainUnit, SubUnit, SubValue
                Dim gangRange As Double
                gangRange = RoundDownDblData(SubValue * GetGangPinCount(PinList), DIGIT_NUMBER)
                enableBoardRange.Add gangRange & SubUnit & MainUnit
            Else
                Set enableBoardRange = tempRange
            End If
    End Select
End Function

Private Function queryBoardName(ByVal PinList As String) As String
    '### ボード名の取得 '08/04/28 OK ######################
    '①単数ピン/カンマ区切りのピンリストはOK
    '②一部にギャング/マージを含むピンリストでもOK
    '③DPSのギャングピンまたはそれを含むピンリストはNG
    '④異なるボードのピンリストはNG
    '⑤定義されていないピンまたはそれを含むピンリストはNG
    Dim ip750board As New CVISVISrcSelector
    ip750board.Initialize
    On Error GoTo IS_INVALID
    queryBoardName = Replace(ip750board.Pins(PinList).BoardName, "dc", "")
    Exit Function
IS_INVALID:
    queryBoardName = NOT_DEFINE
End Function

Private Function queryBoardSpec(ByVal PinList As String, ByVal measureMode As String) As Double()
    '### ボード印加限界値の取得 '08/04/28 OK ##############
    '①単数ピン/カンマ区切りのピンリストはOK
    '②上記の場合、DPSのGetForceILimitはNG
    '③一部にギャング/マージを含むピンリストのGetForceVLimitはOK
    '④一部にギャング/マージを含むピンリストのGetForceILimitはNG
    '⑤DPSのギャングピンまたはそれを含むピンリストはNG
    '⑥異なるボードのピンリストはNG
    '⑦定義されていないピンまたはそれを含むピンリストはNG
    Dim ip750board As New CVISVISrcSelector
    ip750board.Initialize
    Dim errLimit(1) As Double
    errLimit(0) = INVALIDATION_VALUE
    errLimit(1) = INVALIDATION_VALUE
    On Error GoTo IS_INVALID
    Select Case measureMode
        Case "MI":
            queryBoardSpec = ip750board.Pins(PinList).GetForceVLimit
        Case "MV":
            queryBoardSpec = ip750board.Pins(PinList).GetForceILimit
        Case Else
            queryBoardSpec = errLimit
    End Select
    Exit Function
IS_INVALID:
    queryBoardSpec = errLimit
End Function

Private Function queryBoardSpecForBPMU(ByVal PinList As String, ByVal measureMode As String) As Double()
    '### BPMUボード印加限界値の取得 '08/04/28 OK ##########
    '①単数ピン/カンマ区切りのピンリストはOK
    '②異なるボードピンまたはそれを含むピンリストでもOKとなってしまう
    '③定義されていないピンまたはそれを含むピンリストでもOKとなってしまう
    '④NOT_DEFINEのピンはNG
    Dim ip750board As New CVISVISrcSelector
    ip750board.Initialize
    Dim errLimit(1) As Double
    errLimit(0) = INVALIDATION_VALUE
    errLimit(1) = INVALIDATION_VALUE
    If PinList = NOT_DEFINE Then GoTo IS_INVALID
    On Error GoTo IS_INVALID
    Select Case measureMode
        Case "MI":
            queryBoardSpecForBPMU = ip750board.Pins(PinList, dcBPMU).GetForceVLimit
        Case "MV":
            queryBoardSpecForBPMU = ip750board.Pins(PinList, dcBPMU).GetForceILimit
        Case Else
            queryBoardSpecForBPMU = errLimit
    End Select
    Exit Function
IS_INVALID:
    queryBoardSpecForBPMU = errLimit
End Function

Private Function queryBoardRange(ByVal PinList As String, ByVal measureMode As String) As Collection
    '### ボード測定レンジリストの取得 '08/04/28 OK ########
    '①単数ピン/カンマ区切りのピンリストはOK
    '②上記の場合、PPMU/DPSのMeasVRangeListの返り値は"None"のみ
    '③一部にギャング/マージを含むピンリストはギャング/マージピンが優先される
    '④DPSのギャングピンまたはそれを含むピンリストはNG
    '⑤異なるボードのピンリストはNG
    '⑥定義されていないピンまたはそれを含むピンリストはNG
    Dim ip750board As New CVISVISrcSelector
    ip750board.Initialize
    Dim rangeList As Collection
    On Error GoTo IS_INVALID
    Select Case measureMode
        Case "MI":
            Set queryBoardRange = DecomposeStringList(ip750board.Pins(PinList).MeasIRangeList)
        Case "MV":
            If ip750board.Pins(PinList).BoardName = "dcICUL1G" Then
                'ICUL1GではRangeのリストはなし
                Set queryBoardRange = Nothing
            Else
                Set queryBoardRange = DecomposeStringList(ip750board.Pins(PinList).MeasVRangeList)
            End If
        Case Else
            Set queryBoardRange = Nothing
    End Select
    Exit Function
IS_INVALID:
    Set queryBoardRange = Nothing
End Function

Private Function queryBoardRangeForBPMU(ByVal PinList As String, ByVal measureMode As String) As Collection
    '### BPMUボード測定レンジリストの取得 '08/04/28 OK ####
    '①単数ピン/カンマ区切りのピンリストはOK
    '②異なるボードピンまたはそれを含むピンリストでもOKとなってしまう
    '③定義されていないピンまたはそれを含むピンリストでもOKとなってしまう
    '④NOT_DEFINEのピンはNG
    Dim ip750board As New CVISVISrcSelector
    ip750board.Initialize
    Dim rangeList As Collection
    If PinList = NOT_DEFINE Then GoTo IS_INVALID
    On Error GoTo IS_INVALID
    Select Case measureMode
        Case "MI":
            Set queryBoardRangeForBPMU = DecomposeStringList(ip750board.Pins(PinList, dcBPMU).MeasIRangeList)
        Case "MV":
            Set queryBoardRangeForBPMU = DecomposeStringList(ip750board.Pins(PinList, dcBPMU).MeasVRangeList)
        Case Else
            Set queryBoardRangeForBPMU = Nothing
    End Select
    Exit Function
IS_INVALID:
    Set queryBoardRangeForBPMU = Nothing
End Function

Private Function createPinList(ByVal currCell As Range, ByVal cellFrmt As Boolean) As String
    '### ピンリストの作成 #################################
    '08/05/09 OK
    On Error GoTo IS_INVALID
    Dim pinCells As Collection
    Set pinCells = groupCellList(currCell, cellFrmt)
    If pinCells Is Nothing Then GoTo IS_INVALID
    '### カンマ区切りの文字列に展開 #######################
    createPinList = ComposeStringList(pinCells)
    Exit Function
IS_INVALID:
    createPinList = NOT_DEFINE
End Function

Private Function groupCellList(ByVal currCell As Range, ByVal cellFrmt As Boolean) As Collection
    '### グループセルオブジェクトの収集 ###################
    '08/05/09 OK
    On Error GoTo IS_INVALID
    Dim dataCell As Range
    Set dataCell = targetCell(currCell, TEST_PINS)
    If IsEmpty(dataCell) Or Not isEnableArea(currCell) Then GoTo IS_INVALID
    '### グループの頭出し #################################
    Dim dataIndex As Long
    Dim cellIndex As Range
    Do While IsGroupHeader(dataCell.offset(dataIndex, 0)) = False
        dataIndex = dataIndex - 1
    Loop
    Set cellIndex = dataCell.offset(dataIndex, 0)
    '### グループのセルオブジェクトを収集 #################
    dataIndex = 0
    Set groupCellList = New Collection
    Do While IsGroupFooter(cellIndex.offset(dataIndex, 0)) = False
        groupCellList.Add cellIndex.offset(dataIndex, 0)
        dataIndex = dataIndex + 1
    Loop
    '### 有効なピン名のみリストに格納 #####################
    Dim enablePinList As Collection
    Set enablePinList = collectEnablePins(groupCellList, cellFrmt)
    '### 無効なピンがあった場合は無効リストを返す #########
    If groupCellList.Count <> enablePinList.Count Then GoTo IS_INVALID
    Exit Function
IS_INVALID:
    Set groupCellList = Nothing
End Function

Private Function collectEnablePins(ByVal PinList As Collection, ByVal cellFrmt As Boolean) As Collection
    '### ピンリストから有効なピン名を収集する #############
    '08/05/09 OK
    Set collectEnablePins = New Collection

    '=== Add Eee-Job V2.14 ===
    '書式設定セル用
    Dim EnableCells As Range
    Dim ErrorCells As Range
    '=== Add Eee-Job V2.14 ===

    Dim currPin As Range
    Dim topPinType As String
    '### 先頭ピンのボード名を取得 #########################
    topPinType = queryBoardName(PinList.Item(1))
    Dim currType As String
    For Each currPin In PinList
        '### 先頭ピンが無効の場合は全てNG #################
        If topPinType = NOT_DEFINE Then
            If cellFrmt Then
                'errorCell currPin                      '=== Add Eee-Job V2.14 ===
                MakeUnionRange ErrorCells, currPin      '=== Add Eee-Job V2.14 ===
            End If
        Else
        '=== Add Eee-Job V2.14 ===
            currType = queryBoardName(currPin)
            '### 無効及び先頭ピンと異なるリソースの場合はNG
            If currType = NOT_DEFINE Or currType <> topPinType Then
                If cellFrmt Then
                    'errorCell currPin
                    MakeUnionRange ErrorCells, currPin
                End If
            '### 複数ピンかつギャングピンの場合はNG #######
            ElseIf IsGangPins(currPin) And PinList.Count > 1 Then
                If cellFrmt Then
                    'errorCell currPin
                    MakeUnionRange ErrorCells, currPin
                End If
            '### 有効ピンのみピン名を収集 #################
            Else
                collectEnablePins.Add currPin.Value
                If cellFrmt Then
                    'enableCell currPin
                    MakeUnionRange EnableCells, currPin
                End If
            End If
        End If
    Next currPin

    '一括書式設定
    If Not EnableCells Is Nothing Then
        enableCell EnableCells
    End If
    If Not ErrorCells Is Nothing Then
        errorCell ErrorCells
    End If
        '=== Add Eee-Job V2.14 ===
End Function
'=== Add Eee-Job V2.14 ===
'レンジ結合用の関数
Private Sub MakeUnionRange(ByRef pUnionRange As Range, ByRef pCurrentRange As Range)
    If pUnionRange Is Nothing Then
        Set pUnionRange = pCurrentRange
    Else
        Set pUnionRange = Union(pUnionRange, pCurrentRange)
    End If
End Sub

Private Sub enableCell(ByVal currCell As Range)
    '### セルの書式設定：有効である事を表示する ###########
    With currCell.Font
        If .Name <> "Tahoma" Or IsNull(.Name) Then .Name = "Tahoma"
        If .FontStyle <> "標準" Or IsNull(.FontStyle) Then .FontStyle = "標準"
        If .Size <> 10 Or IsNull(.Size) Then .Size = 10
        If .Strikethrough <> False Or IsNull(.Strikethrough) Then .Strikethrough = False
        If .Superscript <> False Or IsNull(.Superscript) Then .Superscript = False
        If .Subscript <> False Or IsNull(.Subscript) Then .Subscript = False
        If .OutlineFont <> False Or IsNull(.OutlineFont) Then .OutlineFont = False
        If .Shadow <> False Or IsNull(.Shadow) Then .Shadow = False
        If .Underline <> xlUnderlineStyleNone Or IsNull(.Underline) Then .Underline = xlUnderlineStyleNone
        If .ColorIndex <> xlAutomatic Or IsNull(.ColorIndex) Then .ColorIndex = xlAutomatic
    End With
    With currCell.Interior
        If .ColorIndex <> xlNone Or IsNull(.ColorIndex) Then .ColorIndex = xlNone
    End With
End Sub

Private Sub disableCell(ByVal currCell As Range)
    '### セルの書式設定：無効である事を表示する ###########
    With currCell.Font
        If .Name <> "Tahoma" Or IsNull(.Name) Then .Name = "Tahoma"
        If .FontStyle <> "標準" Or IsNull(.FontStyle) Then .FontStyle = "標準"
        If .Size <> 10 Or IsNull(.Size) Then .Size = 10
        If .Strikethrough <> False Or IsNull(.Strikethrough) Then .Strikethrough = False
        If .Superscript <> False Or IsNull(.Superscript) Then .Superscript = False
        If .Subscript <> False Or IsNull(.Subscript) Then .Subscript = False
        If .OutlineFont <> False Or IsNull(.OutlineFont) Then .OutlineFont = False
        If .Shadow <> False Or IsNull(.Shadow) Then .Shadow = False
        If .Underline <> xlUnderlineStyleNone Or IsNull(.Underline) Then .Underline = xlUnderlineStyleNone
        If .ColorIndex <> 16 Or IsNull(.ColorIndex) Then .ColorIndex = 16
    End With
    With currCell.Interior
        If .ColorIndex <> 0 Or IsNull(.ColorIndex) Then .ColorIndex = 0
        If .Pattern <> xlLightUp Or IsNull(.Pattern) Then .Pattern = xlLightUp
        If .PatternColorIndex <> 48 Or IsNull(.PatternColorIndex) Then .PatternColorIndex = 48
    End With
End Sub

Private Sub errorCell(ByVal currCell As Range)
    '### セルの書式設定：データ入力ミスを表示する #########
    With currCell.Font
        If .Name <> "Tahoma" Or IsNull(.Name) Then .Name = "Tahoma"
        If .FontStyle <> "標準" Or IsNull(.FontStyle) Then .FontStyle = "標準"
        If .Size <> 10 Or IsNull(.Size) Then .Size = 10
        If .Strikethrough <> False Or IsNull(.Strikethrough) Then .Strikethrough = False
        If .Superscript <> False Or IsNull(.Superscript) Then .Superscript = False
        If .Subscript <> False Or IsNull(.Subscript) Then .Subscript = False
        If .OutlineFont <> False Or IsNull(.OutlineFont) Then .OutlineFont = False
        If .Shadow <> False Or IsNull(.Shadow) Then .Shadow = False
        If .Underline <> xlUnderlineStyleNone Or IsNull(.Underline) Then .Underline = xlUnderlineStyleNone
        If .ColorIndex <> xlAutomatic Or IsNull(.ColorIndex) Then .ColorIndex = xlAutomatic
    End With
    With currCell.Interior
        If .ColorIndex <> 38 Or IsNull(.ColorIndex) Then .ColorIndex = 38
    End With
End Sub

Private Sub warningCell(ByVal currCell As Range)
    '### セルの書式設定：データが無効である事を表示する ###
    With currCell.Font
        If .Name <> "Tahoma" Or IsNull(.Name) Then .Name = "Tahoma"
        If .FontStyle <> "標準" Or IsNull(.FontStyle) Then .FontStyle = "標準"
        If .Size <> 10 Or IsNull(.Size) Then .Size = 10
        If .Strikethrough <> False Or IsNull(.Strikethrough) Then .Strikethrough = False
        If .Superscript <> False Or IsNull(.Superscript) Then .Superscript = False
        If .Subscript <> False Or IsNull(.Subscript) Then .Subscript = False
        If .OutlineFont <> False Or IsNull(.OutlineFont) Then .OutlineFont = False
        If .Shadow <> False Or IsNull(.Shadow) Then .Shadow = False
        If .Underline <> xlUnderlineStyleNone Or IsNull(.Underline) Then .Underline = xlUnderlineStyleNone
        If .ColorIndex <> xlAutomatic Or IsNull(.ColorIndex) Then .ColorIndex = xlAutomatic
    End With
    With currCell.Interior
        If .ColorIndex <> 36 Or IsNull(.ColorIndex) Then .ColorIndex = 36
    End With
End Sub
'=== Add Eee-Job V2.14 ===
Private Function IsCategoryHeader(ByVal currCell As Range) As Boolean
    '### カテゴリのパラメータ有効行判定 ###################
    '08/05/01 OK
    '指定セルのカテゴリカラムにデータ入力があると有効
    If isEnableArea(currCell) Then
        Dim categoryCell As Range
        Set categoryCell = targetCell(currCell, TEST_CATEGORY)
        If Not IsEmpty(categoryCell) Then
            IsCategoryHeader = True
        End If
    End If
End Function

Private Function IsGroupHeader(ByVal currCell As Range) As Boolean
    '### ピングループのパラメータ有効行判定 ###############
    '08/05/01 OK
    '①指定セルのメジャーピンカラムに入力がある
    '②そのカラムの直前のカラムが空白である又はパラメータ有効領域の先頭セルである
    '上記2点を満たすと有効となる
    If isEnableArea(currCell) Then
        Dim pinNameCell As Range
        Set pinNameCell = targetCell(currCell, TEST_PINS)
        If Not IsEmpty(pinNameCell) Then
            Dim topCell As Range
            Set topCell = currCell.parent.Range(TEST_PINS)
            If IsEmpty(pinNameCell.offset(-1, 0)) Or _
               topCell.Row = pinNameCell.offset(-1, 0).Row Then
                IsGroupHeader = True
            End If
        End If
    End If
End Function

Private Function IsGroupFooter(ByVal currCell As Range) As Boolean
    '### ピングループのフッターパラメータ有効行判定 #######
    '08/05/01 OK
    '①指定セルのメジャーピンカラムが空白である
    '②そのカラムの直前のカラムに入力がある又はパラメータ有効領域の先頭セルでない
    '上記2点を満たすと有効となる
    If isEnableArea(currCell) Then
        Dim pinNameCell As Range
        Set pinNameCell = targetCell(currCell, TEST_PINS)
        If IsEmpty(pinNameCell) Then
            Dim topCell As Range
            Set topCell = currCell.parent.Range(TEST_PINS)
            If Not IsEmpty(pinNameCell.offset(-1, 0)) Or _
               topCell.Row <> pinNameCell.offset(-1, 0).Row Then
                IsGroupFooter = True
            End If
        End If
    End If
End Function

Private Function isEnableArea(ByVal currCell As Range) As Boolean
    '### 指定セルのパラメータ有効領域判定 #################
    '08/05/01 OK
    'ENDキーワードが存在しない時は最終行まで有効
    Dim topCell As Range
    Dim endCell As Range
    With currCell.parent
        Set topCell = .Range(TEST_CATEGORY)
        Set endCell = .Columns(topCell.Column).Find("END")
    End With
    If currCell.Row > topCell.Row Then
        If endCell Is Nothing Then
            isEnableArea = True
        ElseIf endCell.Row > currCell.Row Then
            isEnableArea = True
        End If
    End If
End Function

Private Function targetCell(ByVal refCell As Range, ByVal targetLabel As String) As Range
    With refCell.parent
        Set targetCell = .Cells(refCell.Row, .Range(targetLabel).Column)
    End With
End Function

'############# エクセルシートマクロ関数群 #########################################################
Public Sub ShowSpecInfo()
'内容:
'   DCシナリオシートスペック表示マクロ関数
'
'パラメータ:
'
'注意事項:
'   アクティブなDCテストシナリオシート上でテストラベルセルを右クリックした時のメニューから呼び出される
'
    On Error GoTo ErrHandler
    '### アクティブデータシートの取得 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(JOB_LIST_TOOL, TEST_INSTANCES_TOOL)
    If DataSheet Is Nothing Then
        Err.Raise 9999, "showSpecInfo", "Can Not Find The Active Instance Sheet !"
    End If
    '### アクティブセルの取得 #############################
    Dim currCell As Range
    Set currCell = ActiveSheet.Application.ActiveCell
    '### インスタンスシートリーダーの作成 #################
    Dim instanceReader As CInstanceSheetReader
    Set instanceReader = CreateCInstanceSheetReader
    With instanceReader
        .Initialize DataSheet.Name
        .AsIFileStream.SetLocation currCell.Value
    End With
    '### 各パラメータオブジェクト作成と読み込み ###########
    Dim paramLoLimit As CParamDouble
    Set paramLoLimit = CreateCParamDouble
    With paramLoLimit.AsIParameter
        .Name = USERMACRO_LOLIMIT
        .Read instanceReader
    End With
    Dim paramHiLimit As CParamDouble
    Set paramHiLimit = CreateCParamDouble
    With paramHiLimit.AsIParameter
        .Name = USERMACRO_HILIMIT
        .Read instanceReader
    End With
    Dim paramJudge As CParamLong
    Set paramJudge = CreateCParamLong
    Dim judgeChar As String
    With paramJudge.AsIParameter
        .Name = USERMACRO_JUDGE
        .Read instanceReader
        If .AsDouble = 0 Then
            judgeChar = "NONE"
        ElseIf .AsDouble = 1 Then
            judgeChar = "D <"
        ElseIf .AsDouble = 2 Then
            judgeChar = "D >"
        ElseIf .AsDouble = 3 Then
            judgeChar = "< D >"
        Else
            judgeChar = "ERROR"
        End If
    End With
    Dim paramUnit As CParamString
    Set paramUnit = CreateCParamString
    Dim MainUnit As String
    Dim SubUnit As String
    Dim SubValue As Double
    With paramUnit.AsIParameter
        .Name = USERMACRO_UNIT
        .Read instanceReader
        SplitUnitValue "999" & .AsString, MainUnit, SubUnit, SubValue
    End With
    '### スペック情報の表示 ###############################
    MsgBox "INSTANCE SHEET" & Chr(9) & " : [" & DataSheet.Name & "]" & Chr(13) & _
           "TEST LABEL" & Chr(9) & " : [" & currCell.Value & "]" & Chr(13) & _
           "LOWER LIMIT" & Chr(9) & " : [" & paramLoLimit.AsIParameter.AsDouble / SubUnitToValue(SubUnit) & paramUnit.AsIParameter.AsString & "]" & Chr(13) & _
           "UPPER LIMIT" & Chr(9) & " : [" & paramHiLimit.AsIParameter.AsDouble / SubUnitToValue(SubUnit) & paramUnit.AsIParameter.AsString & "]" & Chr(13) & _
           "JUDEGE" & Chr(9) & Chr(9) & " : LOWER [ " & judgeChar & " ] UPPER", vbOKOnly + vbInformation, "SPEC INFOMATION"
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub CheckExaminationMode()
'内容:
'   DCシナリオシート実験フラグのチェックマクロ関数
'
'パラメータ:
'
'注意事項:
'   ワークブックが閉じられる直前に呼び出される
'
    '### アクティブデータシートの取得 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    '### DCシナリオシートリーダーの作成 ###################
    Dim scenarioReader As CDcScenarioSheetReader
    Set scenarioReader = CreateCDcScenarioSheetReader
    scenarioReader.Initialize DataSheet.Name
    '### パラメータオブジェクトの作成 #####################
    Dim paramExamin As CParamBoolean
    Set paramExamin = CreateCParamBoolean
    With paramExamin.AsIParameter
        .Name = EXAMIN_FLAG
    End With
    '### 実験フラグがON設定のカウント #####################
    Dim trueMode As Long
    Do While Not scenarioReader.AsIActionStream.IsEndOfCategory
        paramExamin.AsIParameter.Read scenarioReader
        If paramExamin.AsIParameter.AsBoolean Then
            trueMode = trueMode + 1
        End If
        scenarioReader.AsIActionStream.MoveNextCategory
    Loop
    '### 実験フラグがONの場合注意を促す ###################
    Dim myAns As Integer
    If trueMode > 0 Then
        myAns = MsgBox("[DC Test Scenario]" & vbCrLf & trueMode & _
                       " Cells With 'TRUE' Found In [Examination - Flag] Field!" & vbCrLf & _
                       " Do You Want To Replace Them With 'FALSE' ?", _
                         vbYesNo + vbExclamation, "Examination Mode Alert")
        Select Case myAns:
            Case vbYes:
                clearExaminMode DataSheet
                clearResultData DataSheet
                reverseExaminFlag DataSheet, False
            Case vbNo:
        End Select
    End If
    '### パラメータオブジェクトの作成 #####################
    Dim paramMode As CParamBoolean
    Set paramMode = CreateCParamBoolean
    With paramMode.AsIParameter
        .Name = IS_VALIDATE
        .Read scenarioReader
    End With
    '### レンジバリデーションモードがONの場合注意を促す ###
    If paramMode.AsIParameter.AsBoolean Then
        myAns = MsgBox("[DC Test Scenario]" & vbCrLf & _
                       " 'TRUE' Found In [Range Validation Check Box] Field!" & vbCrLf & _
                       " Do You Want To Replace It With 'FALSE' ?", _
                         vbYesNo + vbExclamation, "Examination Mode Alert")
        Select Case myAns:
            Case vbYes:
                clearValidationMode DataSheet
            Case vbNo:
        End Select
    End If
End Sub

Public Sub SetChangeStatus()
'内容:
'   DCシナリオシートパラメータ変更プロパティ操作マクロ関数
'
'パラメータ:
'
'注意事項:
'
    '### アクティブデータシートの取得 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    '### パラメータプロパティ操作 #########################
    Dim shtObject As Object
    Set shtObject = DataSheet
    shtObject.IsChanged = True
End Sub

Public Sub SwitchExamFlag()
'内容:
'   DCシナリオシート実験フラグの切り替えマクロ関数
'
'パラメータ:
'
'注意事項:
'
    On Error GoTo ErrHandler
    '### アクティブデータシートの取得 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    '### DCシナリオシートリーダーの作成 ###################
    Dim scenarioReader As CDcScenarioSheetReader
    Set scenarioReader = CreateCDcScenarioSheetReader
    scenarioReader.Initialize DataSheet.Name
    '### パラメータオブジェクトの作成 #####################
    Dim paramExamin As CParamBoolean
    Set paramExamin = CreateCParamBoolean
    With paramExamin.AsIParameter
        .Name = EXAMIN_FLAG
        .Read scenarioReader
    End With
    '### 先頭の実験フラグの読み込み #######################
    Dim WriteFlag As Boolean
    If paramExamin.AsIParameter.AsBoolean Then
        WriteFlag = False
    Else
        WriteFlag = True
    End If
    '### 実験フラグの切り替え実行 #########################
    reverseExaminFlag DataSheet, WriteFlag
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub SetExamMode()
'内容:
'   DCシナリオシート実験モードクリアマクロ関数
'
'パラメータ:
'
'注意事項:
'   マクロ登録されたボタンのクリックで呼び出される
'
    On Error GoTo ErrHandler
    '### アクティブデータシートの取得 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    '### 実験モード設定のクリア実行 #######################
    clearExaminMode DataSheet
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub ClearExamResult()
'内容:
'   DCシナリオシート実験結果クリアマクロ関数
'
'パラメータ:
'
'注意事項:
'   マクロ登録されたボタンのクリックで呼び出される
'
    On Error GoTo ErrHandler
    '### アクティブデータシートの取得 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    '### 測定結果データのクリア実行 #######################
    clearResultData DataSheet
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub MakePlaybackTable()
'内容:
'   DC再生リファレンスデータ作成マクロ関数
'
'パラメータ:
'
'注意事項:
'   マクロ登録されたボタンのクリックで呼び出される
'
    On Error GoTo ErrHandler
    '### DCシナリオシートリーダーの作成 ###################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then
        Err.Raise 9999, "MakePlaybackTable", "Can Not Find The Active Dc Scenario Sheet !"
    End If
    Dim scenarioReader As CDcScenarioSheetReader
    Set scenarioReader = CreateCDcScenarioSheetReader
    scenarioReader.Initialize DataSheet.Name
    '### DC再生データシートライターの作成 #################
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_PLAYBACK_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    Dim playbackWriter As CDcPlaybackSheetWriter
    Set playbackWriter = CreateCDcPlaybackSheetWriter
    With playbackWriter
        .Initialize DataSheet.Name
        .ClearCells
    End With
    '### インスタンスシートリーダーの作成 #################
    Set DataSheet = GetUsableDataSht(JOB_LIST_TOOL, TEST_INSTANCES_TOOL)
    If DataSheet Is Nothing Then
        Err.Raise 9999, "MakePlaybackTable", "Can Not Find The Active Instance Sheet !"
    End If
    Dim instanceReader As CInstanceSheetReader
    Set instanceReader = CreateCInstanceSheetReader
    instanceReader.Initialize DataSheet.Name
    '### パラメータオブジェクトの作成 #####################
    Dim paramTName As CParamString
    Set paramTName = CreateCParamString
    Dim ParamLabel As CParamName
    Set ParamLabel = CreateCParamName
    Dim paramLLow As CParamDouble
    Set paramLLow = CreateCParamDouble
    paramLLow.AsIParameter.Name = USERMACRO_LOLIMIT
    Dim paramLHi As CParamDouble
    Set paramLHi = CreateCParamDouble
    paramLHi.AsIParameter.Name = USERMACRO_HILIMIT
    Dim paramUnit As CParamString
    Set paramUnit = CreateCParamString
    paramUnit.AsIParameter.Name = USERMACRO_UNIT
    Dim paramSLLow As CParamString
    Set paramSLLow = CreateCParamString
    paramSLLow.AsIParameter.Name = PB_LIMIT_LO
    Dim paramSLHi As CParamString
    Set paramSLHi = CreateCParamString
    paramSLHi.AsIParameter.Name = PB_LIMIT_HI
    '### データテーブル作成実行 ###########################
    Do While Not scenarioReader.AsIActionStream.IsEndOfCategory
        With paramTName.AsIParameter
            .Name = TEST_CATEGORY
            .Read scenarioReader
            .Name = PB_CATEGORY
            .WriteOut playbackWriter
        End With
        Do While Not scenarioReader.AsIActionStream.IsEndOfGroup
            Do While Not scenarioReader.AsIActionStream.IsEndOfData
                With ParamLabel.AsIParameter
                    .Name = MEASURE_LABEL
                    .Read scenarioReader
                End With
                If ParamLabel.AsIParameter.AsString <> NOT_DEFINE Then
                    With ParamLabel.AsIParameter
                        .Name = PB_LABEL
                        .WriteOut playbackWriter
                    End With
                    instanceReader.AsIFileStream.SetLocation ParamLabel.AsIParameter.AsString
                    Dim MainUnit As String
                    Dim SubUnit As String
                    Dim SubValue As Double
                    With paramUnit.AsIParameter
                        .Read instanceReader
                        SplitUnitValue "999" & .AsString, MainUnit, SubUnit, SubValue
                    End With
                    paramLHi.AsIParameter.Read instanceReader
                    paramLLow.AsIParameter.Read instanceReader
                    With paramSLHi.AsIParameter
                        .AsString = paramLHi.AsIParameter.AsDouble / SubUnitToValue(SubUnit) & paramUnit.AsIParameter.AsString
                        .WriteOut playbackWriter
                    End With
                    With paramSLLow.AsIParameter
                        .AsString = paramLLow.AsIParameter.AsDouble / SubUnitToValue(SubUnit) & paramUnit.AsIParameter.AsString
                        .WriteOut playbackWriter
                    End With
                    playbackWriter.AsIFileStream.MoveNext
                End If
                scenarioReader.AsIActionStream.MoveNextData
            Loop
            scenarioReader.AsIActionStream.MoveNextGroup
        Loop
        scenarioReader.AsIActionStream.MoveNextCategory
    Loop
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub RangeValidationCheckBox_Click()
'内容:
'   DCシナリオログレポートチェックボックスのイベントマクロ
'
'パラメータ:
'
'注意事項:
'   シナリオシート上のチェックボックスのオン/オフで呼び出される
'
    On Error GoTo ErrHandler
    '### アクティブデータシートの取得 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    If DataSheet.Name <> ActiveSheet.Name Then Exit Sub
    '### シートマネージャの初期化 #########################
    InitControlShtReader
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub HoldSheetInfo(ByVal chCell As Range, ByVal toolName As String)
'内容:
'   データシート上の名前、バージョンの管理を行う
'
'パラメータ:
'    [changedCell]  In   変更されたセル
'    [toolName]     In   保持するデータシート名
'
'注意事項:
'
    If chCell.Address = TOOL_NAME_CELL Then
        Application.EnableEvents = False
        chCell.Value = toolName
        Application.EnableEvents = True
    ElseIf chCell.Address = VERSION_CELL Then
        Application.EnableEvents = False
        chCell.Value = CURR_VERSION
        Application.EnableEvents = True
    End If
End Sub

Private Sub reverseExaminFlag(ByVal ActiveSheet As Worksheet, ByVal examinFlag As Boolean)
    '### DCシナリオシートライターの作成 ###################
    Dim scenarioWriter As CDcScenarioSheetWriter
    Set scenarioWriter = CreateCDcScenarioSheetWriter
    scenarioWriter.Initialize ActiveSheet.Name
    scenarioWriter.AsIActionStream.Rewind
    On Error GoTo SHEET_ERROR
    '### パラメータオブジェクトの作成 #####################
    Dim paramFlag As CParamBoolean
    Set paramFlag = CreateCParamBoolean
    With paramFlag.AsIParameter
        .Name = EXAMIN_FLAG
        .AsBoolean = examinFlag
    End With
    '### 実験フラグの切り替え実行 #########################
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    With scenarioWriter
        Do While Not .AsIActionStream.IsEndOfCategory
            paramFlag.AsIParameter.WriteOut scenarioWriter
            .AsIActionStream.MoveNextCategory
        Loop
        .AsIParameterWriter.WriteAsBoolean DATA_CHANGED, True
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    Exit Sub
SHEET_ERROR:
    scenarioWriter.AsIParameterWriter.WriteAsBoolean DATA_CHANGED, True
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

Private Sub clearExaminMode(ByVal ActiveSheet As Worksheet)
    '### DCシナリオシートライターの作成 ###################
    Dim scenarioWriter As CDcScenarioSheetWriter
    Set scenarioWriter = CreateCDcScenarioSheetWriter
    scenarioWriter.Initialize ActiveSheet.Name
    scenarioWriter.AsIActionStream.Rewind
    On Error GoTo SHEET_ERROR
    '### パラメータオブジェクトの作成 #####################
    Dim paramMode As CParamString
    Set paramMode = CreateCParamString
    With paramMode.AsIParameter
        .Name = EXAMIN_MODE
        .AsString = ""
    End With
    '### 実験モード設定のクリア実行 #######################
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    With scenarioWriter
        Do While Not .AsIActionStream.IsEndOfCategory
            paramMode.AsIParameter.WriteOut scenarioWriter
            .AsIActionStream.MoveNextCategory
        Loop
        .AsIParameterWriter.WriteAsBoolean DATA_CHANGED, True
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    Exit Sub
SHEET_ERROR:
    scenarioWriter.AsIParameterWriter.WriteAsBoolean DATA_CHANGED, True
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

Private Sub clearResultData(ByVal ActiveSheet As Worksheet)
    '### DCシナリオシートライターの作成 ###################
    Dim scenarioWriter As CDcScenarioSheetWriter
    Set scenarioWriter = CreateCDcScenarioSheetWriter
    scenarioWriter.Initialize ActiveSheet.Name
    scenarioWriter.AsIActionStream.Rewind
    '### パラメータオブジェクトの作成 #####################
    Dim paramTime As CParamString
    Set paramTime = CreateCParamString
    With paramTime.AsIParameter
        .Name = EXAMIN_EXECTIME
        .AsString = ""
    End With
    Dim paramUnit As CParamString
    Set paramUnit = CreateCParamString
    With paramUnit.AsIParameter
        .Name = EXAMIN_RESULTUNIT
        .AsString = ""
    End With
    Dim paramResult As New Collection
    Dim dataIndex As Long
    For dataIndex = 0 To GetSiteCount
        paramResult.Add ""
    Next dataIndex
    '### 測定結果データのクリア実行 #######################
    Application.ScreenUpdating = False
    With scenarioWriter
        Do While Not .AsIActionStream.IsEndOfCategory
            Do While Not .AsIActionStream.IsEndOfGroup
                paramTime.AsIParameter.WriteOut scenarioWriter
                Do While Not .AsIActionStream.IsEndOfData
                    .AsIParameterWriter.WriteAsString EXAMIN_RESULT, ComposeStringList(paramResult)
                    paramUnit.AsIParameter.WriteOut scenarioWriter
                    .AsIActionStream.MoveNextData
                Loop
                .AsIActionStream.MoveNextGroup
            Loop
            .AsIActionStream.MoveNextCategory
        Loop
    End With
    Application.ScreenUpdating = True
End Sub

Private Sub clearValidationMode(ByVal ActiveSheet As Worksheet)
    '### DCシナリオシートライターの作成 ###################
    Dim scenarioWriter As CDcScenarioSheetWriter
    Set scenarioWriter = CreateCDcScenarioSheetWriter
    scenarioWriter.Initialize ActiveSheet.Name
    '### パラメータオブジェクトの作成 #####################
    Dim paramMode As CParamBoolean
    Set paramMode = CreateCParamBoolean
    '### レンジバリデーションモード設定のクリア実行 #######
    With paramMode.AsIParameter
        .Name = IS_VALIDATE
        .AsBoolean = False
        .WriteOut scenarioWriter
    End With
End Sub

Public Sub ValidateDCTestSenario()
'内容:
'   DCシナリオシートフォーマットを整理するマクロ関数
'
'パラメータ:
'
'注意事項:
'   ①セルのグルーピング処理
'   ②アクショングループフォーマット整形
'   ③各パラメータのチェック
'   を実行する
'
    On Error GoTo ErrHandler
    '### アクティブデータシートの取得 #####################
    Dim DataSheet As Worksheet
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, DC_SCENARIO_TOOL)
    If DataSheet Is Nothing Then Exit Sub
    '### ワークシートライターの準備 #######################
    Dim scenarioWriter As CDcScenarioSheetWriter
    Set scenarioWriter = CreateCDcScenarioSheetWriter
    Application.ScreenUpdating = False
    '### フォーマット整形実行 #############################
    With scenarioWriter
        .Initialize DataSheet.Name
        .SetGrouping
        .Validate
    End With
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
End Sub

Public Sub SetSheetBackground(ByVal activeSht As Object)
    '### アクティブデータシートの取得 #####################
    Dim DataSheet As Worksheet
    Dim toolName As String
    toolName = activeSht.Range("B1").Value
    On Error Resume Next
    Set DataSheet = GetUsableDataSht(SHEET_MANAGER_TOOL, toolName)
    If Err.Number = 9999 Then
        activeSht.SetBackgroundPicture fileName:=""
        Exit Sub
    End If
    If DataSheet Is Nothing Then
        activeSht.SetBackgroundPicture fileName:= _
            GetJobRootPath & "\bin\DT_NotInJob.gif"
        Exit Sub
    End If
    If DataSheet.Name = activeSht.Name Then
        activeSht.SetBackgroundPicture fileName:=""
    Else
        activeSht.SetBackgroundPicture fileName:= _
            GetJobRootPath & "\bin\DT_NotInJob.gif"
    End If
End Sub

Public Function GetUsableDataSht(ByVal ctrlShName As String, ByVal toolName As String) As Worksheet
    '### アクティブデータシートオブジェクトの作成 #########
    Dim ctrlSheet As CDataSheetManager
    Set ctrlSheet = CreateCDataSheetManager
    ctrlSheet.Initialize ctrlShName
    Set GetUsableDataSht = ctrlSheet.GetActiveDataSht(toolName)
End Function


