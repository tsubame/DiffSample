Attribute VB_Name = "PALS_LoopAdj_Mod"
Option Explicit

'==========================================================================
' モジュール名：  PALS_LoopAdj_mod.bas
' 概要        ：  LOOP調整で使用する関数群
' 備考        ：  なし
' 更新履歴    ：  Rev1.0      2010/09/30　新規作成        K.Sumiyashiki
'==========================================================================

Public Const LOOPTOOLNAME As String = "Auto Loop Parameter Adjust"
Public Const LOOPTOOLVER As String = "1.41"

Public g_blnLoopStop As Boolean

Enum enum_DataTrendType
    em_trend_None       '傾向無し
    em_trend_Shift      'シフト
    em_trend_Slope      '上昇・下降
    em_trend_Sudden     '飛び値
    em_trend_Uneven     'バラツキ
End Enum


Public Const CLM_NO     As Integer = 1                '測定Noを記入する列
Public Const CLM_TEST   As Integer = 2                '項目名を記入する列
Public Const CLM_UNIT   As Integer = 3                '単位を記入する列
Public Const CLM_CNT    As Integer = 4                'ループ回数を記入する列
Public Const CLM_MIN    As Integer = 5                '最小値を記入する列
Public Const CLM_AVG    As Integer = 6                '平均値を記入する列
Public Const CLM_MAX    As Integer = 7                '最大値を記入する列
Public Const CLM_SIGMA  As Integer = 8                'σを記入する列
Public Const CLM_3SIGMA As Integer = 9                '3σを記入する列
Public Const CLM_1PAR10 As Integer = 10               '規格幅/10を記入する列
Public Const CLM_LOW    As Integer = 11               '下限規格を記入する列
Public Const CLM_HIGH   As Integer = 12               '上限規格を記入する列
Public Const CLM_3SIGMAPARSPEC As Integer = 13        '3σ/規格幅を記入する列
'>>>2010/12/13 K.SUMIYASHIKI ADD
Public Const CLM_JUDGELIMIT As Integer = 14           'ループバラツキの判断基準レベルを記入する列(Test InstancesのLoopJudgeLimitを記入する列)
'<<<2010/12/13 K.SUMIYASHIKI ADD

Public Const ROW_NAME   As Integer = 1                '測定Lot名を記入する行
Public Const ROW_WAFER  As Integer = 2                'ウェーハNoを記入する行
Public Const ROW_MACHINEJOB As Integer = 3            '測定装置、JOB名を記入する行
Public Const ROW_LOOPCOUNT As Integer = 4             'ループ回数(全site合計)を記入する行
Public Const ROW_DATE   As Integer = 5                '測定日を記入する行
Public Const ROW_LABEL  As Integer = 6                'ループ結果のラベルを記入する行
Public Const ROW_DATASTART As Integer = ROW_LABEL + 1 'ループ結果のデータを記入する先頭行

Public Const MODE_AUTO As String = "AUTO"


Public Type ChangeParamsInformation
    MinWait As Double                        '設定可能な最小ウェイト
    MaxWait As Double                        '設定可能な最大ウェイト
    WaitTrialCnt As Integer                  'Wait変更回数
    AveTrialCnt As Integer                   'Average変更回数
    Pre_Average As Integer                   '前回の取り込み回数
    Pre_Wait As Double                       '前回の取り込み前ウェイト
    Pre_VariationTrend As enum_DataTrendType '前測定時のバラツキ傾向
    Flg_WaitFinish As Boolean                'Wait調整完了フラグ
    Flg_AverageFinish As Boolean             '取り込み回数調整完了フラグ
End Type


'単位換算用係数
'Private Const TERA   As Double = 1000000000000#         'テラ
'Private Const GIGA   As Long = 1000000000               'ギガ
Private Const MEGA    As Long = 1000000                  'メガ
Private Const KIRO    As Long = 1000                     'キロ
Private Const MILLI   As Double = 0.001                  'ミリ
Private Const MAICRO  As Double = 0.000001               'マイクロ
Private Const NANO    As Double = 0.000000001            'ナノ
Private Const PIKO    As Double = 0.000000000001         'ピコ
Private Const FEMTO   As Double = 0.000000000000001      'フェムト
Private Const percent As Double = 0.01

Private Const LABEL_GRADE As String = "grade"

Public g_MaxPalsCount As Long     'フォームに入力した最大測定回数(デフォルト:100回)

Public Const FIRST_VARIATION_CHECK_CNT As Integer = 25        '最初に傾向分析を行う回数(デフォルト:30回)
Public Const VARIATION_CHECK_STEP As Integer = 1            '傾向分析を行うステップ(デフォルト:30回以降、毎測定)

Public ChangeParamsInfo() As ChangeParamsInformation        '各カテゴリのパラメータ推移を保存する変数

'測定データの情報を保存する構造体
Public Type DatalogInfo
    MeasureDate As String
    JobName     As String
    SwNode      As String
End Type

Public g_AnalyzeIgnoreCnt As Integer

Public Sub sub_LoopFrmShow()
    frm_PALS_LoopAdj_Main.Show
End Sub

'********************************************************************************************
' 名前 : sub_SetLoopData
' 内容 : ダイアログからLOOP帳票をするデータログを選択し、ファイルパスを取得
' 引数 : なし
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_SetLoopData()

On Error GoTo errPALSsub_SetLoopData

    g_strOutputDataText = ""

    g_strOutputDataText = Application.GetOpenFilename( _
        title:="!!!!!!!!!!!!!!!!!!!!   Select Target LoopData   !!!!!!!!!!!!!!!!!!!!", _
        fileFilter:="IP750 LoopDataFile (*.txt), *.txt ")

Exit Sub

errPALSsub_SetLoopData:
    Call sub_errPALS("Set datalog name error at 'sub_SetLoopData'", "2-2-01-0-04")

End Sub


'********************************************************************************************
' 名前 : sub_CheckLoopData
' 内容 : ダイアログからLOOP帳票をするデータログを選択し、ファイルパスを取得
' 引数 : lngNowLoopCnt:現状の測定回数
' 戻値： True  :バラツキなし
'        False :バラツキあり
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Function sub_CheckLoopData(ByVal lngNowLoopCnt As Long) As Boolean
    
    '返り値をTrueで初期化
    'バラツキに問題がなければTrueが返る
    sub_CheckLoopData = True
    
    If g_ErrorFlg_PALS Then
        Exit Function
    End If
    
On Error GoTo errPALSsub_CheckLoopData
    
    Dim TestNo As Long          'ループカウンタ(テスト項目を示す)
    Dim sitez As Long           'ループカウンタ(Site番号を示す)
    
    'ユーザーフォームのステータスを変更
    Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Checking...")

    '全項目を繰り返し
    For TestNo = 0 To PALS.CommonInfo.TestCount
        
        'DC以外のデータ傾向判断を行う
        If PALS.CommonInfo.TestInfo(TestNo).CapCategory1 <> "DC" _
            Or Len(PALS.CommonInfo.TestInfo(TestNo).CapCategory1) > 0 Then                       'DCを定数化
            
            '全Site繰り返し
            For sitez = 0 To nSite
                '3σ/規格幅が1項目でも規定値を超えていた場合、返り値をFalseに変更
                If Not sub_JudgeLoopData(TestNo, sitez, lngNowLoopCnt) Then
                    sub_CheckLoopData = False
                End If
            Next sitez
        End If
    Next TestNo

    If sub_CheckLoopData = False And g_AnalyzeIgnoreCnt > 0 Then
        g_AnalyzeIgnoreCnt = g_AnalyzeIgnoreCnt - 1
        sub_CheckLoopData = True
    End If

Exit Function

errPALSsub_CheckLoopData:
    Call sub_errPALS("Check LoopData error at 'sub_CheckLoopData'", "2-2-02-0-05")

End Function


'********************************************************************************************
' 名前: sub_JudgeLoopData
' 内容: 引数"lngTestNo"と"sitez"で渡された項目・サイトの特性値・規格幅から、3σ/規格幅を計算し、
'       その結果が許容範囲(基本は0.1)を超えた場合、バラツキ傾向判断を行う。
'       バラツキがあった場合、各カテゴリのバラツキ情報(最大バラツキ項目等)を更新する。
' 引数: lngTestNo      : 項目を示す番号
'       sitez          : サイト番号
'       lngNowLoopCnt  : 測定済み回数
' 戻値: True  : バラツキ問題なし
'       False : バラツキあり
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_JudgeLoopData(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As Boolean

    '返り値をTrueで初期化
    'バラツキに問題がなければTrueが返る
    sub_JudgeLoopData = True

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_JudgeLoopData

    Dim dblStandardWidth As Double          '規格幅
    Dim enumJudge As enum_DataTrendType     '傾向分類を示す列挙体
    
    '初期化
    enumJudge = em_trend_None
    
    With PALS.CommonInfo.TestInfo(lngTestNo)
    
        Select Case .arg2
            '規格無し
            Case 0
                sub_JudgeLoopData = True
                Exit Function
            
            '下限規格のみ
            Case 1
                dblStandardWidth = Abs(.LowLimit)
            
            '上限規格のみ
            Case 2
                dblStandardWidth = Abs(.HighLimit)
            
            '上下限規格両方
            Case 3
                dblStandardWidth = .HighLimit - .LowLimit
    
            Case Else
                Call sub_errPALS("Get standard width error at 'sub_JudgeLoopData'", "2-2-03-2-06")
        End Select
        
        '規格幅が0以外の時のみ実行（0割り防止）
        If dblStandardWidth <> 0 Then
        
            '3σ/規格幅が規定値以上の場合、傾向確認を行う
            If ((.site(sitez).Sigma * 3# / dblStandardWidth)) >= .LoopJudgeLimit And .LoopJudgeLimit <> 0 Then
    
                'バラツキがある場合、Falseを返り値に設定
                sub_JudgeLoopData = False
    
                '傾向を確認する為に、指定回数は傾向判断へ行かない
                If g_AnalyzeIgnoreCnt > 0 Then
                    Exit Function
                End If
    
                '傾向を判断し、傾向分類(列挙体:enum_DataTrendTypeで定義)を返り値として返す
                enumJudge = sub_AnalyzeLoopData(lngTestNo, sitez, lngNowLoopCnt)
    
                '次測定で改善を行う為に必要な各カテゴリーの情報を更新
                If Not sub_UpdateVariationLoopData(lngTestNo, sitez, dblStandardWidth, enumJudge) Then
                    '途中でエラーになった場合、エラーを返す
                    Call sub_errPALS("Update variation Loopdata error at 'sub_UpdateVariationLoopData'", "2-2-03-0-07")
                End If
            End If
        End If
    
    End With
        
Exit Function

errPALSsub_JudgeLoopData:
    Call sub_errPALS("Judge LoopData error at 'sub_JudgeLoopData'", "2-2-03-0-08")

End Function
        

'********************************************************************************************
' 名前: sub_AnalyzeLoopData
' 内容: バラツキの傾向を確認する関数
' 引数: lngTestNo         : 項目を示す番号
'       sitez             : サイト番号
'       lngNowLoopCnt     : 測定済み回数
' 戻値: バラツキ傾向を示す列挙体
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
' 　　　　　 Rev2.0      2011/06/20　処理変更   K.Sumiyashiki
'                                    ⇒F検定を使用しての判断アルゴリズムへ変更
'                                      (関数全体のフローを変更)
'********************************************************************************************
Private Function sub_AnalyzeLoopData(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As enum_DataTrendType

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_AnalyzeLoopData

    '初期化
    sub_AnalyzeLoopData = em_trend_None
    
    With PALS.CommonInfo.TestInfo(lngTestNo).site(sitez)
        Dim dbl_F_Data As Double
        '通常の特性値ばらつきと微分値のループばらつきより、F値を算出
        dbl_F_Data = ((.Sigma ^ 2) * 2) / (.Differential_Sigma(lngNowLoopCnt) ^ 2)
    End With
        
    'F検定より、ランダム性有りと判断
    If sub_Get_F_Value(lngNowLoopCnt, 2, "bottom") < dbl_F_Data And dbl_F_Data < sub_Get_F_Value(lngNowLoopCnt, 2, "top") Then

'>>>2010/12/13 K.SUMIYASHIKI ADD
        '飛び値と判断されるような大きなバラツキを判断
        sub_AnalyzeLoopData = sub_JudgeBaratuki(lngTestNo, sitez, lngNowLoopCnt)
'<<<2010/12/13 K.SUMIYASHIKI ADD


'>>>2010/12/13 K.SUMIYASHIKI UPDATE
        '大きなバラツキで無い場合のみ処理
        If sub_AnalyzeLoopData = em_trend_None Then
            '飛び値の判断
            '->メディアン処理後、元の平均+2σ以上のものがあれば飛び値と判断
            sub_AnalyzeLoopData = sub_JudgeTobiti(lngTestNo, sitez, lngNowLoopCnt)
        End If
'<<<2010/12/13 K.SUMIYASHIKI UPDATE

    'F検定より、ランダム性無しと判断
    Else
        '飛び値で無い場合のみ処理
        If sub_AnalyzeLoopData = em_trend_None Then
            'シフトor上昇or下降の判断
            sub_AnalyzeLoopData = sub_JudgeShift(lngTestNo, sitez, lngNowLoopCnt)
        End If

    End If
    
    '飛び値・シフト・上昇・下降で無い場合、バラツキと判断
    If sub_AnalyzeLoopData = em_trend_None Then
        'バラツキを示す値を返り値に設定
        sub_AnalyzeLoopData = em_trend_Uneven
        Debug.Print ("Baratuki")
        Debug.Print ("TestName : " & PALS.CommonInfo.TestInfo(lngTestNo).tname)
        Debug.Print ("Site     : " & sitez) & vbCrLf
    End If
    
Exit Function

errPALSsub_AnalyzeLoopData:
    Call sub_errPALS("Analyze Loopdata error at 'sub_AnalyzeLoopData'", "2-2-04-0-09")
    
End Function


'********************************************************************************************
' 名前: sub_UpdateVariationLoopData
' 内容: バラツキがあった場合、バラツキがあったカテゴリのバラツキデータを確認し、
'       他項目のバラツキデータと比較する。その際、以前のデータよりバラツキが大きく、且つ、データの傾向が悪ければ、
'       今回の項目のバラツキデータで更新を行う。
'       更新内容:傾向・バラツキ値・項目名・サイト情報。
'       データのバラツキは、バラツキ⇒飛び値⇒上昇・下降⇒シフトの順に対応を行う。
' 引数: lngTestNo         : 項目を示す番号
'       sitez             : サイト番号
'       dblStandardWidth  : 規格幅
'       enumJudge         : バラツキ傾向
' 戻値: True  : バラツキ問題なし
'       False : バラツキあり
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_UpdateVariationLoopData(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal dblStandardWidth As Double, _
                                        ByVal enumJudge As enum_DataTrendType) As Boolean
                
    If g_ErrorFlg_PALS Then
        sub_UpdateVariationLoopData = True
        Exit Function
    End If
                
On Error GoTo errPALSsub_UpdateVariationLoopData
                
    Dim colTargetCategory As New Collection     'カテゴリーを格納するコレクション
    
    With PALS.CommonInfo.TestInfo(lngTestNo)
        'コレクションにCapCategory1の値(ex:OF,ML)を追加
        colTargetCategory.Add Item:=PALS.LoopParams.CategoryInfoList(.CapCategory1)
        If Len(.CapCategory2) Then
            'CapCategory2に値(ex:OF,ML)が記述されていれば、コレクションに追加
            colTargetCategory.Add PALS.LoopParams.CategoryInfoList(.CapCategory2)
        End If
    End With
        
    Dim valTargetCategory As Variant                'コレクションに格納されているカテゴリー名が入る
    Dim enumCategoryTrend As enum_DataTrendType     '選択された項目のデータ傾向を格納
    
    'データ分コレクションを繰り返し
    For Each valTargetCategory In colTargetCategory
    
        '現状の指定カテゴリーのデータ傾向(最悪値)をenumCategoryTrendに一時格納
        enumCategoryTrend = PALS.LoopParams.LoopCategory(valTargetCategory).VariationTrend
                
        With PALS.LoopParams.LoopCategory(valTargetCategory)
            '現状の指定カテゴリーのデータ傾向(最悪値)より、今回のデータ傾向が悪い場合は、各データを上書き
            'データ傾向が同じ場合は、3σ/規格幅を比較し、今回の方が悪い場合は、各データを上書き
            'データ傾向の比較条件->ばらつき⇒飛び値⇒上昇・下降⇒シフト
            If enumCategoryTrend = enumJudge Then
                '以前の3σ/規格幅のデータと比較し、今回が悪ければ各データを上書き
                If .VariationLevel < PALS.CommonInfo.TestInfo(lngTestNo).site(sitez).Sigma * 3# / dblStandardWidth Then
                    .VariationLevel = PALS.CommonInfo.TestInfo(lngTestNo).site(sitez).Sigma * 3# / dblStandardWidth
                    .TargetTestName = PALS.CommonInfo.TestInfo(lngTestNo).tname
                    .VariationSite = sitez
                End If
            ElseIf enumCategoryTrend < enumJudge Then
                '今回のデータ傾向の方が悪い場合は、各データを上書き
                .VariationLevel = PALS.CommonInfo.TestInfo(lngTestNo).site(sitez).Sigma * 3# / dblStandardWidth
                .TargetTestName = PALS.CommonInfo.TestInfo(lngTestNo).tname
                .VariationSite = sitez
                .VariationTrend = enumJudge
            End If
        End With
    Next valTargetCategory

    '途中でエラーが無ければ、Trueを返す
    sub_UpdateVariationLoopData = True

Exit Function

errPALSsub_UpdateVariationLoopData:
    Call sub_errPALS("Update variation Loopdata error at 'sub_UpdateVariationLoopData'", "2-2-05-0-10")

End Function


'********************************************************************************************
' 名前: sub_JudgeTobiti
' 内容: バラツキ傾向が飛び値かどうか確認する関数
'       測定データに3タップのメディアンフィルタを掛け、メディアン処理後の値と元の値の減算を行う。
'       その際の差分が2σ以上あれば、飛び値とする。
' 引数: lngTestNo         : 項目を示す番号
'       sitez             : サイト番号
'       lngNowLoopCnt     : 測定済み回数
' 戻値: バラツキ傾向を示す列挙体
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_JudgeTobiti(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As enum_DataTrendType

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_JudgeTobiti

'飛び値の判断
'->メディアン処理後、元の平均+2σ以上のものがあれば飛び値と判断

    'メディアンタップ数の設定(とりあえずはローカルの静的変数で定義)
    Const Median_Tap As Integer = 3
    
    'メディアンタップ数が偶数 or 1 の場合エラーメッセージを表示
    If Median_Tap Mod 2 = 0 Or Median_Tap = 1 Then
        Call sub_errPALS("Median tap number is even number or 1." & vbCrLf & "         Please check median tap number !" & vbCrLf & "         at 'sub_JudgeTobiti'", "2-2-06-5-11")
        Exit Function
    End If

    'メディアン時の除外範囲設定
    Dim intRemoveArea As Integer
    intRemoveArea = Int(Median_Tap / 2)

    'メディアン後のデータ格納配列
    '除外範囲分、配列数を削除
    Dim dblConvertData() As Double
    ReDim dblConvertData(lngNowLoopCnt - intRemoveArea)

    Dim data_cnt As Long            'ループカウンタ(データインデックスを示す)
    Dim tap_cnt As Long             'ループカウンタ(メディアン時のタップ番号を示す)
    Dim dblTmpData() As Double      'タップ数分のデータを一時格納する配列
    
    '除外範囲を除いた箇所を繰り返し
    For data_cnt = 1 + intRemoveArea To lngNowLoopCnt - intRemoveArea
        
        'メディアンタップ数に応じて再定義及び初期化
        ReDim dblTmpData(Median_Tap - 1)
        
        'タップ数分繰り返し
        For tap_cnt = 0 To Median_Tap - 1
            'data_cntの前後タップ数分のデータをdblTmpDataに一時格納
            dblTmpData(tap_cnt) = PALS.CommonInfo.TestInfo(lngTestNo).site(sitez).Data(data_cnt - intRemoveArea + tap_cnt)
        Next tap_cnt

        'dblTmpDataを降順でバブルソート
        Call sub_BubbleSort(dblTmpData)

        '中央値をdblConvertDataに代入
        dblConvertData(data_cnt) = dblTmpData(UBound(dblTmpData) - intRemoveArea)

    Next data_cnt

    'メディアン処理による、測定開始直後の除外エリアを補完
    For data_cnt = 1 To intRemoveArea
        dblConvertData(data_cnt) = dblConvertData(intRemoveArea + 1)
    Next data_cnt

    With PALS.CommonInfo.TestInfo(lngTestNo).site(sitez)
        '現在の測定回数分、データを繰り返す
        For data_cnt = 1 To UBound(dblConvertData)
            'メディアン処理後のデータから元の平均値を減算し、その値が2σ以上の場合飛び値と判断
            If (Abs((.Data(data_cnt) - dblConvertData(data_cnt))) - (.Sigma * 2)) > 0 Then
                Debug.Print ("Tobiti error")
                Debug.Print ("TestName : " & PALS.CommonInfo.TestInfo(lngTestNo).tname)
                Debug.Print ("Site     : " & sitez) & vbCrLf
'                Debug.Print ("Tobiti error")
                sub_JudgeTobiti = em_trend_Sudden
                
            End If
        Next data_cnt
    End With

Exit Function

errPALSsub_JudgeTobiti:
    Call sub_errPALS("Check LoopData error at 'sub_JudgeTobiti'", "2-2-06-0-12")

End Function


'********************************************************************************************
' 名前: sub_BubbleSort
' 内容: バブルソート
' 引数: dblVal         : 並び替えを行う配列
'       blnSortAsc     : 昇順or降順を指定するオプション
'                        (デフォルトはFalse:昇順)
' 戻値: なし
' 備考：ソート後の結果例↓↓
'       False ⇒ dblVal(1):10, dblVal(2):5, dblVal(3):1
'       True  ⇒ dblVal(1):1 , dblVal(2):5, dblVal(3):10
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_BubbleSort(ByRef dblVal() As Double, Optional ByVal blnSortAsc As Boolean = False)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_BubbleSort

    Dim i As Long       'ループカウンタ
    Dim j As Long       'ループカウンタ

    For i = LBound(dblVal) To UBound(dblVal) - 1
        For j = LBound(dblVal) To LBound(dblVal) + UBound(dblVal) - i - 1
            If dblVal(IIf(blnSortAsc, j, j + 1)) > dblVal(IIf(blnSortAsc, j + 1, j)) Then
                Call sub_Swap(dblVal(j), dblVal(j + 1))
            End If
        Next j
    Next i

Exit Sub

errPALSsub_BubbleSort:
    Call sub_errPALS("BubbleSort error at 'sub_BubbleSort'", "2-2-07-0-13")

End Sub


'********************************************************************************************
' 名前: sub_Swap
' 内容: 引数で渡された2つの値を入れ替える関数
' 引数: dblVal1 : 入れ替える変数1
'       dblVal2 : 入れ替える変数2
' 戻値: なし
' 備考：なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_Swap(ByRef dblVal1 As Double, ByRef dblVal2 As Double)

    Dim dblBuf As Double    '一時格納変数
    
    dblBuf = dblVal1
    dblVal1 = dblVal2
    dblVal2 = dblBuf

End Sub


'********************************************************************************************
' 名前: sub_JudgeShift
' 内容: バラツキ傾向が上昇or下降orシフトかどうか確認する関数
'       測定データの開始付近のデータと、現在の測定回数付近のデータの平均値を比較し、
'   　　その差分が1σ以上あった場合、上昇or下降orシフトと判断する。
'       平均値を取る際のデータ幅は、現在の測定回数の1/10としている。
' 引数: lngTestNo         : 項目を示す番号
'       sitez             : サイト番号
'       lngNowLoopCnt     : 測定済み回数
' 戻値: バラツキ傾向を示す列挙体
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_JudgeShift(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As enum_DataTrendType

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_JudgeShift

    Dim dblBeforeDataSum As Double    '測定開始付近のデータ合計(データ数はintCalcWidthで設定)
    Dim dblAfterDataSum As Double     '現在の測定回数付近のデータ合計(データ数はintCalcWidthで設定)
    Dim dblBeforeDataAve As Double    '測定開始付近のデータ平均(データ数はintCalcWidthで設定)
    Dim dblAfterDataAve As Double     '現在の測定回数付近のデータ平均(データ数はintCalcWidthで設定)

    Dim data_cnt As Long                'ループカウンタ(データインデックスを示す)
    Dim intCalcWidth As Integer         '判断を行う際に使用するデータ幅

    'データ幅を、現在の測定回数の1/10に設定
    intCalcWidth = Int(lngNowLoopCnt / 10)

    With PALS.CommonInfo.TestInfo(lngTestNo).site(sitez)

        '測定開始付近のデータの合計を取得(データ数はintCalcWidthで設定)
        For data_cnt = 1 To intCalcWidth
            dblBeforeDataSum = dblBeforeDataSum + .Data(data_cnt)
        Next data_cnt
        '平均値を計算
        dblBeforeDataAve = (dblBeforeDataSum / intCalcWidth)
    
        '現在の測定回数付近のデータの合計を取得(データ数はintCalcWidthで設定)
        For data_cnt = lngNowLoopCnt - intCalcWidth + 1 To lngNowLoopCnt
            dblAfterDataSum = dblAfterDataSum + .Data(data_cnt)
        Next data_cnt
        '平均値を計算
        dblAfterDataAve = (dblAfterDataSum / intCalcWidth)
    
        '測定開始時のデータ平均と現在の測定回数付近のデータ平均を比較
        '差分が1σ以上あった場合、シフトor上昇or下降と判断
        '1σ以下の場合、関数を抜ける
        If (Abs(dblBeforeDataAve - dblAfterDataAve) < .Sigma * 2) Then
            sub_JudgeShift = em_trend_None
            Exit Function
        End If
    End With

    'シフトor上昇or下降の場合、データがどのパターンか判断する
    sub_JudgeShift = sub_CheckShiftType(lngTestNo, sitez, lngNowLoopCnt)

Exit Function

errPALSsub_JudgeShift:
    Call sub_errPALS("Data Judge error at 'sub_JudgeShift'", "2-2-08-0-14")

End Function


'********************************************************************************************
' 名前: sub_CheckShiftType
' 内容: バラツキ傾向が上昇or下降かシフトかどうか切り分けを行う関数
'       測定データに対し、一つ飛びのデータとの差分を取る処理を行い、そのデータ内に、
'       1σ以上の差分が存在した場合、シフトと判断する。
'       1σ以上の差分がない場合、徐々に変動したとし、上昇or下降と判断する。
' 引数: lngTestNo         : 項目を示す番号
'       sitez             : サイト番号
'       lngNowLoopCnt     : 測定済み回数
' 戻値: バラツキ傾向を示す列挙体
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_CheckShiftType(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As enum_DataTrendType

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_CheckShiftType

    '次のデータとの差分値を格納する配列
    Dim dblConvertData() As Double
    '最後のデータは差分を求められない為、現在の測定回数から-1をし配列を再定義
    ReDim dblConvertData(lngNowLoopCnt - 1)

    Dim data_cnt As Long        'ループカウンタ(データインデックスを示す)
    
    With PALS.CommonInfo.TestInfo(lngTestNo).site(sitez)
    
        'データ数マイナス1だけ繰り返し
        For data_cnt = 2 To lngNowLoopCnt - 1
            
            '一つ飛びのデータとの差分を取得
            dblConvertData(data_cnt) = .Data(data_cnt + 1) - .Data(data_cnt - 1)
            
            'データの差分がσ以上の場合シフトとする
            If Abs(dblConvertData(data_cnt)) - (.Sigma) > 0 Then
                'シフトを示す値を返し終了
                sub_CheckShiftType = em_trend_Shift
                Call sub_errPALS("This data trend is Shift." & vbCrLf & "Please check data!", "2-2-09-7-15")
                Debug.Print ("Shift error")
                Debug.Print ("TestName : " & PALS.CommonInfo.TestInfo(lngTestNo).tname)
                Debug.Print ("Site     : " & sitez) & vbCrLf
                Exit Function
            End If
        Next data_cnt
    End With

    'シフトではない場合、上昇or下降を示す値を返す
    sub_CheckShiftType = em_trend_Slope
        Debug.Print ("Rise or Fall!")
        Debug.Print ("TestName : " & PALS.CommonInfo.TestInfo(lngTestNo).tname)
        Debug.Print ("Site     : " & sitez) & vbCrLf
    Call sub_errPALS("This data trend is Rise or Fall." & vbCrLf & "Please check data!")

Exit Function

errPALSsub_CheckShiftType:
    Call sub_errPALS("Data Judge error at 'sub_CheckShiftType'", "2-2-09-7-16")

End Function


'********************************************************************************************
' 名前 : sub_UpdataLoopParams
' 内容 : 傾向判断の結果から、カテゴリ毎にWait、Averageの変更を行う
' 引数 : なし
' 戻値： True  :エラーなし
'        False :エラーあり
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
'バラツキの傾向確認結果をパラメータへ反映
Public Function sub_UpdataLoopParams() As Boolean

    '返り値の初期化。エラーがなければTrueが返る
    sub_UpdataLoopParams = True

    If g_ErrorFlg_PALS Then
        sub_UpdataLoopParams = True
        Exit Function
    End If
    
On Error GoTo errPALSsub_UpdataLoopParams

    Dim cnt As Long             'ループカウンタ。カテゴリインデックスを示す。
    Dim dblAdjValue As Double   'パラメータを変更する際に、一時的に値を保存する変数
    Dim dblStep As Double       'Waitを変更する際のステップ幅。最小Waitと最大Waitから計算。

    With PALS.LoopParams
        
        'カテゴリの数だけ繰り返し
        For cnt = 1 To .CategoryCount
    
            '初期化
            dblAdjValue = -1
            dblStep = -1
            dblAdjValue = 0
    
            Select Case .LoopCategory(cnt).VariationTrend
                
                '**********今回の傾向が飛び値の場合**********
                Case enum_DataTrendType.em_trend_Sudden
                    
                    'Waitのステップ幅を求める
''                    dblStep = Format((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 6, "#.000")
                    dblStep = Int((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 6)
                    
                    'Waitの初期値が設定最大Waitの場合、調整完了とする
                    If dblStep = 0 Then
                        ChangeParamsInfo(cnt).Flg_WaitFinish = True
                    End If
                    
                    '以前にWaitの変更を行っていない場合
                    If ChangeParamsInfo(cnt).WaitTrialCnt = 0 Then
''                        dblAdjValue = .LoopCategory(cnt).Wait + Format((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 2, "#.000")
                        dblAdjValue = .LoopCategory(cnt).WAIT + Int((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 2)
    
                        'TestConditionシート内のパラメータを変更
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                        
                        'Wait変更を行った回数のインクリメント
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1

                    '以前にWaitの変更を1回行っている場合
                    ElseIf ChangeParamsInfo(cnt).WaitTrialCnt = 1 Then
                        dblAdjValue = .LoopCategory(cnt).WAIT + (dblStep * 2)
                                            
                        'TestConditionシート内のパラメータを変更
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                    
                        'Wait変更を行った回数のインクリメント
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                    
                    '以前にWaitの変更を2回行っている場合
                    ElseIf ChangeParamsInfo(cnt).WaitTrialCnt = 2 Then
                        dblAdjValue = .LoopCategory(cnt).WAIT + dblStep
                                            
                        'TestConditionシート内のパラメータを変更
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                    
                        'Wait変更を行った回数のインクリメント
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                    
                    '以前にWaitの変更を3回以上行っている場合
                    ElseIf ChangeParamsInfo(cnt).WaitTrialCnt >= 3 Then
                        
                        If Not ChangeParamsInfo(cnt).Flg_WaitFinish Then
                        
''                            dblAdjValue = .LoopCategory(cnt).Wait + Format(dblAdjValue + (dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2))), "#.000")
''                            dblAdjValue = .LoopCategory(cnt).Wait + Int(dblAdjValue + (dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2))))
                            dblAdjValue = .LoopCategory(cnt).WAIT + Int(dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2)))
                            
                            'TestConditionシート内のパラメータを変更
                            If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                                sub_UpdataLoopParams = False
                                Exit Function
                            End If
                        
                            'Wait変更を行った回数のインクリメント
                            ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                                            
                            If dblAdjValue > (ChangeParamsInfo(cnt).MaxWait * 0.99) Then
                                If dblAdjValue > ChangeParamsInfo(cnt).MaxWait Then
                                    dblAdjValue = ChangeParamsInfo(cnt).MaxWait
                                End If
                                ChangeParamsInfo(cnt).Flg_WaitFinish = True
                            End If
                                            
                        End If
                    Else
                    
                    
                    End If
    

                '**********今回の傾向がバラツキの場合**********
                Case enum_DataTrendType.em_trend_Uneven
                    If Not ChangeParamsInfo(cnt).Flg_AverageFinish Then
                        
                        Dim intItemNum As Integer       'カテゴリインデックスを一時保存する為の変数
                                                
                        '取り込み回数の初期値が511の場合、調整完了とする
                        If PALS.LoopParams.LoopCategory(cnt).Average = 511 Then
                            
                            ChangeParamsInfo(cnt).Flg_AverageFinish = True
                        
                        Else
                            
                            '変更を行うカテゴリのインデックスを取得
                            intItemNum = PALS.CommonInfo.TestnameInfoList(.LoopCategory(cnt).TargetTestName)
                            
                            '次の取り込み回数を計算
                            dblAdjValue = Int(.LoopCategory(cnt).Average * (.LoopCategory(cnt).VariationLevel _
                                            / PALS.CommonInfo.TestInfo(intItemNum).LoopJudgeLimit) ^ 2)
                            
                            '取り込み回数の倍数指定(TestConditionシートのMode)があった場合、取りこみ回数の調整を行う
                            With .LoopCategory(cnt)
                                '計算後の取りこみ回数が512以上の場合
                                If dblAdjValue > 511 Then
                                    'ModeがAutoに設定されていれば、511に変更
                                    If .mode = MODE_AUTO Then
                                        dblAdjValue = 511
                                    '倍数指定があれば、511に最も近い公倍数に変更
                                    Else
                                        dblAdjValue = 511 - (511 Mod val(.mode))
                                    End If
                                    
                                    'アベレージ調整完了を示すフラグを立てる
                                    ChangeParamsInfo(cnt).Flg_AverageFinish = True
                                
                                Else
                                    'ModeがAutoの場合はそのまま
                                    '倍数指定があれば、現状値以上で最も小さい公倍数に変更
                                    If .mode <> MODE_AUTO Then
                                        dblAdjValue = dblAdjValue + (val(.mode) - (dblAdjValue Mod val(.mode)))
                                    End If
                                
                                    '計算後の取りこみ回数が512以上の場合
                                    If dblAdjValue > 511 Then
                                        '511に最も近い公倍数に変更
                                        dblAdjValue = 511 - (511 Mod val(.mode))
                                        'アベレージ調整完了を示すフラグを立てる
                                        ChangeParamsInfo(cnt).Flg_WaitFinish = True
                                    End If
                                End If
                            End With
                            
                            'TestConditionシート内のパラメータを変更
                            If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Average", dblAdjValue) Then
                                sub_UpdataLoopParams = False
                                Exit Function
                            End If
                            
                            'Average回数変更を行った回数のインクリメント
                            ChangeParamsInfo(cnt).AveTrialCnt = ChangeParamsInfo(cnt).AveTrialCnt + 1
                        End If
                    End If
    
                '今回の傾向が上昇or下降の場合
                Case enum_DataTrendType.em_trend_Slope
    
                '今回の傾向がシフトの場合
                Case enum_DataTrendType.em_trend_Shift
    
                Case Else
                    
            End Select
    
    
    
    
    
            '**********今回の傾向が飛び値以外で、以前に飛び値の対応(Wait変更)を行っている場合**********
            If ChangeParamsInfo(cnt).WaitTrialCnt > 0 _
                And .LoopCategory(cnt).VariationTrend <> enum_DataTrendType.em_trend_Sudden Then
                
                '既にWaitの調整が完了している場合は、調整を行わない
                If Not ChangeParamsInfo(cnt).Flg_WaitFinish Then
                
                    'Waitのステップ幅を求める
''                    dblStep = Format((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 6, "#.000")
                    dblStep = Int((ChangeParamsInfo(cnt).MaxWait - ChangeParamsInfo(cnt).MinWait) / 6)
    
                    '以前にWaitの変更を1回行っている場合
                    If ChangeParamsInfo(cnt).WaitTrialCnt = 1 Then
                        dblAdjValue = .LoopCategory(cnt).WAIT - (dblStep * 2)
                                            
                        'TestConditionシート内のパラメータを変更
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                    
                        'Wait変更を行った回数のインクリメント
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                    
                    '以前にWaitの変更を2回行っている場合
                    ElseIf ChangeParamsInfo(cnt).WaitTrialCnt = 2 Then
                        dblAdjValue = .LoopCategory(cnt).WAIT - dblStep
                                            
                        'TestConditionシート内のパラメータを変更
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                    
                        'Wait変更を行った回数のインクリメント
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                    
                    '以前にWaitの変更を3回以上行っている場合
                    ElseIf ChangeParamsInfo(cnt).WaitTrialCnt >= 3 Then
                    
                    
''                        dblAdjValue = .LoopCategory(cnt).Wait - Format(dblAdjValue - (dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2))), "#.000")
''                        dblAdjValue = .LoopCategory(cnt).Wait - Int(dblAdjValue - (dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2))))
                        dblAdjValue = .LoopCategory(cnt).WAIT - Int(dblStep / (2 ^ (ChangeParamsInfo(cnt).WaitTrialCnt - 2)))
                                            
                        If dblAdjValue < ChangeParamsInfo(cnt).MinWait Then
                            dblAdjValue = ChangeParamsInfo(cnt).MinWait
                        End If
                                            
                        'TestConditionシート内のパラメータを変更
                        If Not .ChangeLoopParams(.LoopCategory(cnt).category, "Wait", dblAdjValue) Then
                            sub_UpdataLoopParams = False
                            Exit Function
                        End If
                    
                        'Wait変更を行った回数のインクリメント
                        ChangeParamsInfo(cnt).WaitTrialCnt = ChangeParamsInfo(cnt).WaitTrialCnt + 1
                
                        'ウェイトが最大Waitの99%以上になった場合、Wait調整を終了
                        '(切捨て誤差がある為、99%以上としている)
                        If dblAdjValue > (ChangeParamsInfo(cnt).MaxWait * 0.99) Then
                            If dblAdjValue > ChangeParamsInfo(cnt).MaxWait Then
                                dblAdjValue = ChangeParamsInfo(cnt).MaxWait
                            End If
                            ChangeParamsInfo(cnt).Flg_WaitFinish = True
                        End If
                    
                    Else
                
                         
                    End If
                End If
            End If
        Next cnt
    End With

Exit Function

errPALSsub_UpdataLoopParams:
    Call sub_errPALS("Updata LoopParameter error at 'sub_UpdataLoopParams'", "2-2-10-0-17")

End Function


'********************************************************************************************
' 名前: sub_Init_ChangeLoopParamsInfo
' 内容: 各カテゴリのパラメータ変更状況や傾向を格納する構造体のデータ初期化
'       最小・最大ウェイトは、同時に設定を行っている
' 引数: なし
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/20　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_Init_ChangeLoopParamsInfo()
    
    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALSsub_Init_ChangeLoopParamsInfo

    Dim i As Long       'ループカウンタ
    
    'カテゴリの数だけ繰り返し
    For i = 1 To UBound(ChangeParamsInfo)
        With ChangeParamsInfo(i)
            .MinWait = PALS.LoopParams.LoopCategory(i).WAIT                 '設定可能な最小ウェイト
            .MaxWait = val(frm_PALS_LoopAdj_Main.txt_maxwait)               '設定可能な最大ウェイト
            .WaitTrialCnt = 0                                               'Wait変更回数
            .AveTrialCnt = 0                                                'Average変更回数
            .Pre_Average = -1                                               '前回の取り込み回数
            .Pre_Wait = -1                                                  '前回の取り込み前ウェイト
            .Pre_VariationTrend = enum_DataTrendType.em_trend_None          '前測定時のバラツキ傾向
            .Flg_WaitFinish = False                                         'Wait調整完了フラグ
            .Flg_AverageFinish = False                                      '取り込み回数調整完了フラグ
        End With
    Next i
    
Exit Sub

errPALSsub_Init_ChangeLoopParamsInfo:
    Call sub_errPALS("Init ChangeLoopParamsInfo error at 'sub_Init_ChangeLoopParamsInfo'", "2-2-11-0-18")
    
End Sub


'********************************************************************************************
' 名前: sub_Update_ChangeLoopParamsInfo
' 内容: バラツキがあり、再測定を実施する際に、今回の各カテゴリの
'       Average、Wait、バラツキ傾向を保存する処理
' 引数: なし
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/20　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Function sub_Update_ChangeLoopParamsInfo() As Boolean

    '返り値の初期化
    'WaitかAverageの変更が一つでもあれば、Trueに変更される
    sub_Update_ChangeLoopParamsInfo = False

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_Update_ChangeLoopParamsInfo

    Dim i As Long       'ループカウンタ
    
    'カテゴリの数だけ繰り返し
    For i = 1 To UBound(ChangeParamsInfo)
        With PALS.LoopParams.LoopCategory(i)
            
            'Averageの変更があれば、データのアップデートを行い、返り値をTrueに変更
            If ChangeParamsInfo(i).Pre_Average <> .Average Then
                ChangeParamsInfo(i).Pre_Average = .Average
                sub_Update_ChangeLoopParamsInfo = True
            End If
            
            'Waitの変更があれば、データのアップデートを行い、返り値をTrueに変更
            If ChangeParamsInfo(i).Pre_Wait <> .WAIT Then
                ChangeParamsInfo(i).Pre_Wait = .WAIT
                sub_Update_ChangeLoopParamsInfo = True
            End If
            
            'バラツキ傾向のアップデート
            ChangeParamsInfo(i).Pre_VariationTrend = .VariationTrend
        End With
    Next i
    
Exit Function

errPALSsub_Update_ChangeLoopParamsInfo:
    Call sub_errPALS("Update ChangeLoopParameterInfo error at 'sub_Update_ChangeLoopParamsInfo'", "2-2-12-0-19")
    
End Function


'********************************************************************************************
' 名前 : sub_OutPutLoopParam
' 内容 : TestConditionシートのパラメータを、測定データログの末尾にテキストで追加
'        下記のようなデータが追加される
'        ########### Parameter ###########
'        Category  Wait      Average
'        ML        0.1       10
'        OF        0.2       20
'        LL        0.4       40
'        SMR       0.7       70
'        DK        1         100
'        #################################
' 引数 : なし
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   M.Imamura
'********************************************************************************************
Public Sub sub_OutPutLoopParam(ByRef MeasureDatalogInfo As DatalogInfo)

'    If g_ErrorFlg_PALS Then
'        Exit Sub
'    End If

On Error GoTo errPALSsub_OutPutLoopParam

    Dim intFileNo As Integer                'ファイル番号
    Dim intCategoryNum As Long              'カテゴリ名を回すループカウンタ
    
    intFileNo = FreeFile                    'ファイル番号の取得
    
    With MeasureDatalogInfo
        .MeasureDate = Year(Date) & "/" & Month(Date) & "/" & Day(Date)
        .JobName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 4)
        .SwNode = Sw_Node
    End With
    
    'TestConditionシートのパラメータを、データログに追記
    'Append(追記)モードで測定データログを開き、各パラメータを追記
    Open g_strOutputDataText For Append As #intFileNo

        With MeasureDatalogInfo
            Print #intFileNo, ""
            Print #intFileNo, "MEASURE DATE : " & .MeasureDate
            Print #intFileNo, "JOB NAME     : " & .JobName
            Print #intFileNo, "SW_NODE      : " & .SwNode
        End With
        
        Print #intFileNo, "########### Parameter ###########"
        Print #intFileNo, "Category" & Space(10 - Len("Category")) & "Wait" & Space(10 - Len("Wait")) & "Average"

        'カテゴリ数繰り返す
        For intCategoryNum = 1 To PALS.LoopParams.CategoryCount
            With PALS.LoopParams.LoopCategory(intCategoryNum)
                Print #intFileNo, .category & Space(10 - Len(.category)) & .WAIT & Space(10 - Len(CStr(.WAIT))) & .Average
            End With
        Next
        
        Print #intFileNo, "#################################"
    
    'データログを閉じる
    Close #intFileNo

Exit Sub

errPALSsub_OutPutLoopParam:
    Call sub_errPALS("OutPut LoopParameter error at 'sub_OutPutLoopParam'", "2-2-13-0-20")

End Sub


'#######################################################################################################################
'########           　帳票作成【データログ、結果】        　     #######################################################
'#######################################################################################################################

'********************************************************************************************
' 名前: sub_MakeLoopResultSheet
' 内容: LOOP帳票作成関数
' 引数: なし
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   H.Ishibashi
'********************************************************************************************
Public Sub sub_MakeLoopResultSheet(ByVal lngMaxCnt As Long, ByRef MeasureDatalogInfo As DatalogInfo)

    ' Excel ApplicationBookSheetオブジェクト定義
    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlWS As Excel.Worksheet

On Error GoTo errPALSsub_MakeLoopResultSheet

    Const DATASTEP As Integer = 200                 'データログを記入する横のMAX数(列256制限のため)
    
    Dim DataRowStep As Integer     'データログを記入する縦の飛ばし数
    Dim Data_Num As Integer     'データ数(ループ回数)
    Dim lngsite As Long
    Dim lngRowJump As Long
    Dim i As Long
    Dim j As Long
    Dim retsu As Long
    Dim gyou As Long

    Dim intSheetCheck As Integer
    Dim SheetName As Worksheet

    'ハッチング用変数　----------
    Const ROW_HAIFUN As Integer = 3
    Const ROW_0 As Integer = 4
    Const ROW_01 As Integer = 5
    Dim lngCount_haifun As Long
    Dim lngCount0 As Long
    Dim lngCount01 As Long
    Dim lngColorAqua As Long
    Dim lngColorYellow As Long
    Dim lngColorOrange As Long
    
    '色のインデックス入力
    '>>>2011/06/14 M.IMAMURA RGB->VBA.RGB Mod.
    '>>>2011/10/03 M.IMAMURA ColorCode Mod.
    lngColorAqua = VBA.RGB(200, 255, 255)
    lngColorYellow = VBA.RGB(250, 250, 200)
    lngColorOrange = VBA.RGB(255, 200, 150)
'    lngColorAqua = VBA.RGB(204, 255, 255)
'    lngColorYellow = VBA.RGB(250, 250, 204)
'    lngColorOrange = VBA.RGB(255, 204, 153)
    '----------------------------
    '<<<2011/10/03 M.IMAMURA ColorCode Mod.
    '<<<2011/06/14 M.IMAMURA RGB->VBA.RGB Mod.

    
    Set xlApp = CreateObject("Excel.Application")       ' Excel Application Object 生成。本プログラムの親Excelではなく、
                                                        ' 新規Excelが起動する事に注意
    xlApp.DisplayAlerts = False                         ' 警告メッセージをFalse設定。上書きしますかなど聞いてこない
                                                        ' 自動で上書き保存する場合等に便利
    xlApp.Visible = False                               ' Bookを表示にする。Falseで非表示(裏で動く)。自動書き込みなどする場合、
                                                        ' 非表示にする事でUserの誤操作を防ぐことも出来る。今回はSampleなので表示。
    Set xlWB = xlApp.Workbooks.Add                      ' Excelに新規Bookを追加。.Open(FileName)メソッドで既存のExcelBookを開くことも可能。
    xlApp.ScreenUpdating = False

    Data_Num = lngMaxCnt
    
    '>>>2011/10/3 M.IMAMURA Add. Dartsとデータの間隔を合わせる
    DataRowStep = Data_Num + 3 + 3
    '<<<2011/10/3 M.IMAMURA Add. Dartsとデータの間隔を合わせる

    For lngsite = nSite To 0 Step -1 '------------------------------------------------------------------ Site_Loop

        '====================================================================================
        '========================= データログシート書き込み処理開始 =========================
        '====================================================================================
        '～ シート追加 (シート名:Data Log_Site0、Data Log_Site1、･･･)　～

'>>>2011/06/02 K.SUMIYASHIKI UPDATE
'シート追加処理を関数化
'''        intSheetCheck = 0
'''        '同じシート名が存在すればフラグを立てる
'''        For Each SheetName In Worksheets
'''            If SheetName.Name = "Data Log_Site" & lngsite Then
'''                intSheetCheck = 1
'''            End If
'''        Next
'''
'''        '既にシートがある場合連番を振ったシート名を付けるためフラグをインクリメント
'''        For Each SheetName In Worksheets
'''            If SheetName.Name = "Data Log_Site" & lngsite & "(" & intSheetCheck & ")" Then
'''                intSheetCheck = intSheetCheck + 1
'''            End If
'''        Next
'''        'シート名変更 (標準：Data Log_Site0、同名シート存在時：Data Log_Site0(1)、Data Log_Site0(2)、、、等)
'''        If intSheetCheck = 0 Then
'''           Sheets.Add.Name = "Data Log_Site" & lngsite
'''        Else
'''           Sheets.Add.Name = "Data Log_Site" & lngsite & "(" & intSheetCheck & ")"
'''        End If
        If lngsite = nSite - 1 Then
            xlWB.Worksheets("Sheet1").Delete
            xlWB.Worksheets("Sheet2").Delete
            xlWB.Worksheets("Sheet3").Delete
        End If
        Set xlWS = xlWB.Worksheets.Add       ' 新規ブックのSheet1にxlWSオブジェクトをセット。
        xlWS.Name = "Data Log_Site" & CStr(lngsite)

'        Call sub_AddSheet("Data Log_Site", lngsite)
'<<<2011/06/02 K.SUMIYASHIKI UPDATE
        
        'IG-XLが入っていないPCで実行する際は、バリデーションを行わない
        If Not frm_PALS_LoopAdj_Main.chk_IGXL_Check Then
            Call sub_Validate
        End If
        
        '～ データは200項目毎に行を変更して書き出す　～
        lngRowJump = 0
        retsu = 0
        For i = 0 To PALS.CommonInfo.TestCount - 1  'koumoku_num=項目数
        
            If PALS.CommonInfo.TestInfo(i).Label = LABEL_GRADE Then
                Exit For
            End If
        
        
            If i > 0 Then
                If i Mod DATASTEP = 0 Then
                    lngRowJump = lngRowJump + 1 '200項目に達したら行を変えるカウントアップ
                    retsu = 0                   '列は0に戻す
                End If
            End If
            retsu = retsu + 1
            gyou = 1 + (DataRowStep * lngRowJump) 'DataRowStep 103

            '>>>2011/10/3 M.IMAMURA Add. Dartsとデータの間隔を合わせる
            If retsu = 1 And gyou > 1 Then
                xlWS.Cells(gyou - 1, retsu).Value = "----------" '項目名出力
            End If
            '<<<2011/10/3 M.IMAMURA Add. Dartsとデータの間隔を合わせる
            
            xlWS.Cells(gyou, retsu).Value = PALS.CommonInfo.TestInfo(i).tname  '項目名出力


'>>>2011/04/20 K.SUMIYASHIKI UPDATE
            With PALS.CommonInfo.TestInfo(i)
                'データログシートに単位を入力
                xlWS.Cells(gyou + 1, retsu).Value = "[" & .Unit & "]"
                
                '下限規格を入力
                If .arg2 = 0 Or .arg2 = 2 Then
                    xlWS.Cells(gyou + 2, retsu).Value = "No_Limit"
                Else
                    xlWS.Cells(gyou + 2, retsu).Value = sub_ReverseConvertUnit(.LowLimit, i)
                End If
                
                '上限規格を入力
                If .arg2 = 0 Or .arg2 = 1 Then
                    xlWS.Cells(gyou + 3, retsu).Value = "No_Limit"
                Else
                    xlWS.Cells(gyou + 3, retsu).Value = sub_ReverseConvertUnit(.HighLimit, i)
                End If
                gyou = gyou + 4
            End With
'<<<2011/04/20 K.SUMIYASHIKI UPDATE


            For j = 1 To lngMaxCnt '[測定回数分ループ]
'>>>2010/12/13 K.SUMIYASHIKI UPDATE
'>>>2011/05/13 K.SUMIYASHIKI UPDATE
'old 101213               Cells(gyou + 1 + j, retsu).value = PALS.CommonInfo.TestInfo(i).Site(lngSite).Data(j)
'old 110513               Cells(gyou + j, retsu).value = PALS.CommonInfo.TestInfo(i).Site(lngSite).Data(j)
                '>>>2011/10/3 M.IMAMURA Add.
                If j = 1 Then xlWS.Cells(gyou, retsu).Value = "0"
                '<<<2011/10/3 M.IMAMURA Add.
                If PALS.CommonInfo.TestInfo(i).site(lngsite).Enable(j) = True Then
                    xlWS.Cells(gyou + j, retsu).Value = sub_ReverseConvertUnit(PALS.CommonInfo.TestInfo(i).site(lngsite).Data(j), i)
                End If
'<<<2011/05/13 K.SUMIYASHIKI UPDATE
'<<<2010/12/13 K.SUMIYASHIKI UPDATE
            Next j
        Next i
        
        '>>>2011/10/3 M.IMAMURA Add.
        xlWS.Cells(gyou + lngMaxCnt * 2, 1).Value = "LoopTimes"
        xlWS.Cells(gyou + lngMaxCnt * 2, 2).Value = lngMaxCnt
        '<<<2011/10/3 M.IMAMURA Add.
        
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Making LoopResultSheet..." & Int((nSite - lngsite + 1) / (nSite + 1) * 50) & "%")
        
    Next lngsite
        
        
        '====================================================================================
        '========================= データログシート書き込み処理終了 =========================
        '====================================================================================

    For lngsite = nSite To 0 Step -1 '------------------------------------------------------------------ Site_Loop
        '>>>2011/10/3 M.Imamura Add.
        '色付セルの個数カウント変数を初期化
        lngCount0 = 0
        lngCount_haifun = 0
        lngCount01 = 0
        '<<<2011/10/3 M.Imamura Add.

        '//////////////////////////////////////////////////////////////////////////////
        '////////////////////// ループ結果シート書き込み処理開始 //////////////////////
        '//////////////////////////////////////////////////////////////////////////////
        '～ シート追加 (基本→TestResult_Site0、TestResult_Site1、････)　～
'>>>2011/06/02 K.SUMIYASHIKI UPDATE
'シート追加処理を関数化
'''        intSheetCheck = 0
'''        '''同じシート名が存在すればフラグを立てる
'''        For Each SheetName In Worksheets
'''            If SheetName.Name = "TestResult_Site" & lngsite Then  '同じシート名が存在すればフラグを立てる
'''                intSheetCheck = 1
'''            End If
'''        Next
'''        '''すでにシートがある場合→ (TestResult_Site0(1)、TestResult_Site0(2)、等連番を振る)
'''        For Each SheetName In Worksheets
'''            If SheetName.Name = "TestResult_Site" & lngsite & "(" & intSheetCheck & ")" Then
'''                intSheetCheck = intSheetCheck + 1
'''            End If
'''        Next
'''        '''シート名変更
'''        If intSheetCheck = 0 Then
'''           Sheets.Add.Name = "TestResult_Site" & lngsite
'''        Else
'''           Sheets.Add.Name = "TestResult_Site" & lngsite & "(" & intSheetCheck & ")"
'''        End If

        Set xlWS = xlWB.Worksheets.Add       ' 新規ブックのSheet1にxlWSオブジェクトをセット。
        xlWS.Name = "TestResult_Site" & CStr(lngsite)
'        Call sub_AddSheet("TestResult_Site", lngsite)
'<<<2011/06/02 K.SUMIYASHIKI UPDATE

        'IG-XLが入っていないPCで実行する際は、バリデーションを行わない
        If Not frm_PALS_LoopAdj_Main.chk_IGXL_Check Then
            Call sub_Validate
        End If
        
        '～ ラベルを書き出す　～
        xlWS.Cells(ROW_LABEL, CLM_NO).Value = "No"
        xlWS.Cells(ROW_LABEL, CLM_TEST).Value = "Test"
        xlWS.Cells(ROW_LABEL, CLM_UNIT).Value = "Unit"
        xlWS.Cells(ROW_LABEL, CLM_CNT).Value = "Cnt"
        xlWS.Cells(ROW_LABEL, CLM_MIN).Value = "MIN"
        xlWS.Cells(ROW_LABEL, CLM_AVG).Value = "AVG"
        xlWS.Cells(ROW_LABEL, CLM_MAX).Value = "MAX"
        xlWS.Cells(ROW_LABEL, CLM_SIGMA).Value = "sigma"
        xlWS.Cells(ROW_LABEL, CLM_3SIGMA).Value = "3sigma"
        xlWS.Cells(ROW_LABEL, CLM_1PAR10).Value = "'l/10"
        xlWS.Cells(ROW_LABEL, CLM_LOW).Value = "Low"
        xlWS.Cells(ROW_LABEL, CLM_HIGH).Value = "High"
        xlWS.Cells(ROW_LABEL, CLM_3SIGMAPARSPEC).Value = "3sigma/spec width"
'>>>2010/12/13 K.SUMIYASHIKI ADD
        xlWS.Cells(ROW_LABEL, CLM_JUDGELIMIT).Value = "LoopJudgeLimit"
'<<<2010/12/13 K.SUMIYASHIKI ADD

        '～ 各項目毎のデータを書き出す　～
        For i = 0 To PALS.CommonInfo.TestCount - 1

            If PALS.CommonInfo.TestInfo(i).Label = LABEL_GRADE Then
                Exit For
            End If


            With PALS.CommonInfo.TestInfo(i)
                xlWS.Cells(ROW_DATASTART + i, CLM_NO).Value = i + 1
                xlWS.Cells(ROW_DATASTART + i, CLM_TEST).Value = .tname
                xlWS.Cells(ROW_DATASTART + i, CLM_UNIT).Value = "[" & .Unit & "]"
                
'>>>2010/12/13 K.SUMIYASHIKI ADD
                If (.LoopJudgeLimit <> 0.1 And .LoopJudgeLimit <> 0) Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_JUDGELIMIT).Value = .LoopJudgeLimit
                End If
'<<<2010/12/13 K.SUMIYASHIKI ADD
                
                With .site(lngsite)
'>>>2011/05/13 K.SUMIYASHIKI ADD
'old101213                    Cells(ROW_DATASTART + i, CLM_CNT).value = lngMaxCnt
                    xlWS.Cells(ROW_DATASTART + i, CLM_CNT).Value = .ActiveValueCnt
'<<<2011/05/13 K.SUMIYASHIKI ADD
                    xlWS.Cells(ROW_DATASTART + i, CLM_MIN).Value = sub_ReverseConvertUnit(.Min, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_AVG).Value = sub_ReverseConvertUnit(.ave, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_MAX).Value = sub_ReverseConvertUnit(.max, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_SIGMA).Value = sub_ReverseConvertUnit(.Sigma, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMA).Value = xlWS.Cells(ROW_DATASTART + i, CLM_SIGMA).Value * 3
                End With
            


                '上限規格なし、下限規格ありの場合　→　規格幅/10＝下限規格×1/10、3σ/規格幅＝3σ/下限規格
                If .arg2 = 1 Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value = sub_ReverseConvertUnit(.LowLimit, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).Value = Abs(xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value / 10)
                    
                    
                    If sub_ReverseConvertUnit(.LowLimit, i) = 0 Then
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = 0
                    Else
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = Abs(xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMA).Value / xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value)
                    End If
                    xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value = "-"
    
    
                '下限規格なし、上限規格ありの場合　→　規格幅/10＝上限規格×1/10、3σ/規格幅＝3σ/上限規格
                ElseIf .arg2 = 2 Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value = sub_ReverseConvertUnit(.HighLimit, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).Value = Abs(xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value / 10)
                    
                    If sub_ReverseConvertUnit(.HighLimit, i) = 0 Then
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = 0
                    Else
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = Abs(xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMA).Value / xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value)
                    End If
                    xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value = "-"
        
    
                '上下規格なしの場合　→　規格幅/10＝"-"、3σ/規格幅＝"-"
                ElseIf .arg2 = 0 Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).Value = "-"
                    xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).HorizontalAlignment = xlCenter
                    xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = "-"
                    xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).HorizontalAlignment = xlCenter
                    xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value = "-"
                    xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value = "-"
    
    
                '上下規格ありの場合　→　規格幅/10＝(上限-下限)/10、3σ/規格幅＝3σ/(上限-下限)
                Else
                    xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value = sub_ReverseConvertUnit(.LowLimit, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value = sub_ReverseConvertUnit(.HighLimit, i)
                    xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).Value = (xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value - xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value) / 10
                    
                    If sub_ReverseConvertUnit((.HighLimit - .LowLimit), i) = 0 Then
'>>>2011/12/12 M.IMAMURA MOD
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = "-"
                        xlWS.Cells(ROW_DATASTART + i, CLM_1PAR10).Value = "-"
'                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).value = 0
'<<<2011/12/12 M.IMAMURA MOD
                    Else
                        xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = Abs(xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMA).Value / _
                                                                            (xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value - xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value))
                    End If
    

                End If
    
'-----------------------------------　3σ÷規格幅=0、0.1の時のハッチング設定　------------------------------------
                '3σ÷規格幅=0 なら　No,Testを薄黄色でハッチングする
                If xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = 0 Then
                    lngCount0 = lngCount0 + 1
                    xlWS.Range(xlWS.Cells(ROW_DATASTART + i, CLM_NO), xlWS.Cells(ROW_DATASTART + i, CLM_TEST)).Interior.color = lngColorYellow
                    
                '3σ÷規格幅="-" なら　No,Testを水色でハッチングする
                ElseIf xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value = "-" Then
                    lngCount_haifun = lngCount_haifun + 1
                    xlWS.Range(xlWS.Cells(ROW_DATASTART + i, CLM_NO), xlWS.Cells(ROW_DATASTART + i, CLM_TEST)).Interior.color = lngColorAqua
                    'ただし、σ=0なら　No,Testを薄黄色でハッチングする
                    If xlWS.Cells(ROW_DATASTART + i, CLM_SIGMA).Value = 0 Then
                        lngCount_haifun = lngCount_haifun - 1
                        lngCount0 = lngCount0 + 1
                        xlWS.Range(xlWS.Cells(ROW_DATASTART + i, CLM_NO), xlWS.Cells(ROW_DATASTART + i, CLM_TEST)).Interior.color = lngColorYellow
                    End If
                '3σ÷規格幅>=0.1 なら　その行を薄オレンジ色でハッチングする
                ElseIf val(xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC).Value) >= 0.1 Then
                    lngCount01 = lngCount01 + 1
                    xlWS.Range(xlWS.Cells(ROW_DATASTART + i, CLM_NO), xlWS.Cells(ROW_DATASTART + i, CLM_3SIGMAPARSPEC)).Interior.color = lngColorOrange
                End If
                '----------------------------------------------------------------------------------------------------------------------
    
    
                If xlWS.Cells(ROW_DATASTART + i, CLM_LOW).Value = "-" Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_LOW).HorizontalAlignment = xlCenter
                End If
                If xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).Value = "-" Then
                    xlWS.Cells(ROW_DATASTART + i, CLM_HIGH).HorizontalAlignment = xlCenter
                End If

            End With

        Next


        '～ 測定情報出力　～
        Dim strLotName As String
        strLotName = Mid$(g_strOutputDataText, InStrRev(g_strOutputDataText, "\") + 1)
        strLotName = Left$(strLotName, Len(strLotName) - 4)
'>>>2011/12/07 M.IMAMURA Mod
        If Left$(strLotName, 12) = "LoopAdjData_" Then
            strLotName = Mid$(strLotName, 13)
        End If
        
        xlWS.Cells(ROW_NAME, CLM_NO).Value = "TestResult_Site" & CStr(lngsite) & "[" & strLotName & "] (no exclusion)"
'        Cells(ROW_NAME, CLM_NO).value = "TestResult[" & strLotName & "] (no exclusion)"
'<<<2011/12/07 M.IMAMURA Mod
        xlWS.Cells(ROW_WAFER, CLM_NO).Value = "Wafer : 1"

'>>>2011/12/07 M.IMAMURA Mod
        xlWS.Cells(ROW_MACHINEJOB, CLM_NO).Value = "Device : " & PALS.CommonInfo.g_strTesterName & "  Program : " & MeasureDatalogInfo.JobName & "  (Re-calculation : " & Left$(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 4) & " )"
'        xlWS.Cells(ROW_MACHINEJOB, CLM_NO).value = "Device : " & PALS.CommonInfo.g_strTesterName & "  Program : " & MeasureDatalogInfo.JobName & "  (Re-calculation : " & MeasureDatalogInfo.JobName & " )"
'<<<2011/12/07 M.IMAMURA Mod
'        xlWS.Cells(ROW_MACHINEJOB, CLM_NO).value = "Device : SKCCDS" & MeasureDatalogInfo.SwNode & "  Program : " & MeasureDatalogInfo.JobName & "  (Re-calculation : " & MeasureDatalogInfo.JobName & " )"
        xlWS.Cells(ROW_LOOPCOUNT, CLM_NO).Value = "Measurement count : " & lngMaxCnt
               
        xlWS.Cells(ROW_DATE, CLM_NO).Value = "Measurement date : " & MeasureDatalogInfo.MeasureDate

        '～　3σ÷規格幅=0、"-"、0.1以上の個数をそれぞれ出力　～
        xlWS.Cells(ROW_HAIFUN, CLM_1PAR10).Value = lngCount_haifun
        xlWS.Cells(ROW_0, CLM_1PAR10).Value = lngCount0
        xlWS.Cells(ROW_01, CLM_1PAR10).Value = lngCount01

        '～　色づけ　～
        xlWS.Cells(ROW_HAIFUN, CLM_3SIGMA).Interior.color = lngColorAqua
        xlWS.Cells(ROW_0, CLM_3SIGMA).Interior.color = lngColorYellow
        xlWS.Cells(ROW_01, CLM_3SIGMA).Interior.color = lngColorOrange

        '～ 掛線、書式設定　～
        Call sub_SetLoopFormat(xlApp, xlWB, xlWS)

        '～　印刷範囲設定　～
        Call sub_PrintSetting(xlApp, xlWB, xlWS)
        
        Call sub_TestingStatusOutPals(frm_PALS_LoopAdj_Main, "Now Making LoopResultSheet..." & 50 + Int((nSite - lngsite + 1) / (nSite + 1) * 50) & "%")

        '//////////////////////////////////////////////////////////////////////////////
        '////////////////////// ループ結果シート書き込み処理終了 //////////////////////
        '//////////////////////////////////////////////////////////////////////////////
        
    Next lngsite '------------------------------------------------------------------------------------- Site_Loop

    xlApp.ScreenUpdating = True
    
    Dim xlFileName As String
    xlFileName = Left$(g_strOutputDataText, Len(g_strOutputDataText) - 4) & ".xls"
    xlWB.SaveAs xlFileName              ' 新規ブックを別名保存
'    xlApp.Visible = True
    xlWB.Close              ' 新規ブック閉じる
    xlApp.Quit              ' Excelを落とす
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    ThisWorkbook.Activate

Exit Sub
    Dim xlFileName2 As String
    xlFileName2 = Left$(g_strOutputDataText, Len(g_strOutputDataText) - 4) & ".xls"
    xlWB.SaveAs xlFileName2              ' 新規ブックを別名保存
'    xlApp.Visible = True
    xlWB.Close              ' 新規ブック閉じる
    xlApp.Quit              ' Excelを落とす
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing

errPALSsub_MakeLoopResultSheet:
    Call sub_errPALS("Make Loop sheet error at 'sub_MakeLoopResultSheet'", "2-2-14-0-21")

End Sub


'#######################################################################################################################
'########           　掛線、書式【結果シート】        　     ###########################################################
'#######################################################################################################################
Private Sub sub_SetLoopFormat(xlApp As Excel.Application, xlWB As Excel.Workbook, xlWS As Excel.Worksheet)

On Error GoTo errPALSsub_SetLoopFormat

    'フォントサイズ基本9、一番上だけ16
    xlWS.Cells.Select
    xlApp.Selection.Font.Size = 9
    xlWS.Range(xlWS.Cells(ROW_NAME, CLM_NO), xlWS.Cells(ROW_NAME, CLM_NO)).Select
    xlApp.Selection.Font.Size = 16

    'フォント＝MS ゴシック
    xlWS.Cells.Font.Name = "ＭＳ ゴシック"

    '書式処理→MIN～3σ、3σ/規格幅の列を小数点第4位まで表示
    xlWS.Range(xlWS.Cells(ROW_DATASTART, CLM_MIN), xlWS.Cells(xlWS.Rows.Count, CLM_3SIGMA)).NumberFormatLocal = "0.00000;-0.00000;0;@"
    xlWS.Range(xlWS.Cells(ROW_DATASTART, CLM_3SIGMAPARSPEC), xlWS.Cells(xlWS.Rows.Count, CLM_3SIGMAPARSPEC)).NumberFormatLocal = "0.00000;-0.00000;0;@"

    '掛線、幅を最適化
    xlWS.Cells(ROW_LABEL, CLM_NO).Select
    xlWS.Range(xlApp.Selection, xlApp.Selection.End(xlToRight)).Select
    xlWS.Range(xlApp.Selection, xlApp.Selection.End(xlDown)).Select
    With xlApp.Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Columns.AutoFit
    End With

    Dim i As Long
    'データ最小、平均、最大セル幅をオートから0.5広げる(印刷時#####になる可能性ありのため)
    For i = CLM_MIN To CLM_MAX
        xlWS.Cells(ROW_LABEL, i).ColumnWidth = xlWS.Cells(ROW_LABEL, i).ColumnWidth + 0.5
    Next

    'セル高さを12に統一
    xlWS.Cells.RowHeight = 12

    '先頭行のみセル高さを24にする
    xlWS.Cells(ROW_NAME, CLM_NO).RowHeight = 24

    'ラベル行を中央揃えに設定
    xlWS.Range(xlWS.Cells(ROW_LABEL, CLM_NO), xlWS.Cells(ROW_LABEL, CLM_3SIGMAPARSPEC)).HorizontalAlignment = xlCenter

Exit Sub

errPALSsub_SetLoopFormat:
    Call sub_errPALS("Set Line error at 'sub_SetLoopFormat'", "2-2-15-0-22")

End Sub
'#######################################################################################################################
'########           　プリント設定【結果シート】        　     #########################################################
'#######################################################################################################################
Private Sub sub_PrintSetting(xlApp As Excel.Application, xlWB As Excel.Workbook, xlWS As Excel.Worksheet)

On Error GoTo errPALSsub_PrintSetting

    Const NEWPAGE = 61

    Dim i As Long
    Dim lngRowlast As Long 'ループ結果シートの最終行の行番号

    'データの部分を印刷範囲に指定
    xlWS.Range(xlWS.Cells(ROW_NAME, CLM_NO), xlWS.Cells(ROW_NAME, CLM_3SIGMAPARSPEC)).Select
    xlWS.Range(xlApp.Selection, xlApp.Selection.End(xlDown)).Select
    lngRowlast = xlApp.Selection.Rows.Count + xlApp.Selection.Row - 1 '最終行の行番号取得
''    ActiveSheet.PageSetup.PrintArea = ActiveCell.CurrentRegion.Address
'↓修正後
''    ActiveSheet.PageSetup.PrintArea = Range(Cells(ROW_NAME, CLM_NO), Cells(lngRowlast, CLM_3SIGMAPARSPEC)).Address
    
    '改行設定
'''''    For i = Page1 + 1 To lngRowlast
    For i = 1 To lngRowlast
        If i Mod NEWPAGE = 0 Then xlWS.Rows(i).PageBreak = xlPageBreakManual
    Next

    'ラベルの行を全ページに印刷されるよう設定
    xlWS.Range(xlWS.Cells(ROW_LABEL, CLM_NO), xlWS.Cells(ROW_LABEL, CLM_NO)).Select
''    ActiveSheet.PageSetup.PrintTitleRows = "$6:$6"
''
''    'ヘッダー、フッター設定
''    With ActiveSheet.PageSetup
''        .CenterHeader = ActiveSheet.name        '中央ヘッダー：シート名
''        .CenterFooter = "&P / &N" & " ページ"   '中央フッター：ページ番号
''        .RightFooter = "ループツール帳票"       '右側フッター：ツールの名前など
''    End With
''
''    '横1ページ分で収める
''    ActiveSheet.PageSetup.FitToPagesWide = 1

    'A1セルを選択して終了
    xlWS.Range("A1").Select

Exit Sub

errPALSsub_PrintSetting:
    Call sub_errPALS("Set PrintArea error at 'sub_PrintSetting'", "2-2-16-0-23")

End Sub


'********************************************************************************************
' 名前: sub_ReverseConvertUnit
' 内容: LOOP帳票に出力する値を、テストインスタンスから取得した単位によって単位変換する関数
' 引数: dblValue   : 変換前の特性値
'       lngTestCnt : 項目番号を示す値
' 戻値: 単位変換後の値
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_ReverseConvertUnit(ByVal dblValue As Double, ByVal lngTestCnt As Long) As Double

On Error GoTo errPALSsub_ReverseConvertUnit

    '単位換算
    Select Case PALS.CommonInfo.TestInfo(lngTestCnt).Unit
        Case ""
            sub_ReverseConvertUnit = dblValue
        
        Case "MA"
            sub_ReverseConvertUnit = dblValue / MEGA

        Case "MV"
            sub_ReverseConvertUnit = dblValue / MEGA

        Case "KV"
            sub_ReverseConvertUnit = dblValue / KIRO

        Case "KA"
            sub_ReverseConvertUnit = dblValue / KIRO
        
        Case "V"
            sub_ReverseConvertUnit = dblValue
        
        Case "v"
            sub_ReverseConvertUnit = dblValue
        
        Case "A"
            sub_ReverseConvertUnit = dblValue
        
        Case "a"
            sub_ReverseConvertUnit = dblValue
        
        Case "mV"
            sub_ReverseConvertUnit = dblValue / MILLI
        
        Case "mv"
            sub_ReverseConvertUnit = dblValue / MILLI
                        
        Case "mA"
            sub_ReverseConvertUnit = dblValue / MILLI
                
        Case "uV"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
        Case "uv"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
        Case "uA"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
        Case "nV"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "nv"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "nA"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "pV"
            sub_ReverseConvertUnit = dblValue / PIKO
        
        Case "pv"
            sub_ReverseConvertUnit = dblValue / PIKO
        
        Case "pA"
            sub_ReverseConvertUnit = dblValue / PIKO
        
        Case "fV"
            sub_ReverseConvertUnit = dblValue / FEMTO
        
        Case "fv"
            sub_ReverseConvertUnit = dblValue / FEMTO
        
        Case "fA"
            sub_ReverseConvertUnit = dblValue / FEMTO
        
        Case "ms"
            sub_ReverseConvertUnit = dblValue / MILLI
        
        Case "us"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
'>>>2010/12/13 K.SUMIYASHIKI ADD
        Case "ns"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "ps"
            sub_ReverseConvertUnit = dblValue / PIKO
'<<<2010/12/13 K.SUMIYASHIKI ADD
        Case "S"
            sub_ReverseConvertUnit = dblValue
                        
        Case "mS"
            sub_ReverseConvertUnit = dblValue / MILLI
        
        Case "uS"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
        Case "nS"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "pS"
            sub_ReverseConvertUnit = dblValue / PIKO
        
        Case "%"
            sub_ReverseConvertUnit = dblValue / percent
                                                
        Case "Kr"
            sub_ReverseConvertUnit = dblValue
'>>>2013/12/03 T.Morimoto ADD
        Case "W"
            sub_ReverseConvertUnit = dblValue
                        
        Case "mW"
            sub_ReverseConvertUnit = dblValue / MILLI
        
        Case "uW"
            sub_ReverseConvertUnit = dblValue / MAICRO
        
        Case "nW"
            sub_ReverseConvertUnit = dblValue / NANO
        
        Case "pW"
            sub_ReverseConvertUnit = dblValue / PIKO
'<<<2013/12/04

        Case Else

'>>>2010/12/13 K.SUMIYASHIKI MESSEGE CHANGE
            Call MsgBox("Error! Not Entry Unit" & "->" & PALS.CommonInfo.TestInfo(lngTestCnt).Unit & vbCrLf & "ErrCode.2-2-17-4-24", vbExclamation)
'<<<2010/12/13 K.SUMIYASHIKI MESSEGE CHANGE
        
    End Select

Exit Function

errPALSsub_ReverseConvertUnit:
    Call sub_errPALS("Convert Unit error at 'sub_ReverseConvertUnit'", "2-2-17-0-25")

End Function


'********************************************************************************************
' 名前: sub_CheckTestConditionWaitData
' 内容: TestConditionに設定してある各カテゴリのWaitが、フォームで指定した最大Wait以上になっていた場合エラーを返す
' 引数: dblMaxWait : フォームで指定した最大Wait
' 戻値: True  :異常値なし
'       False :異常値あり
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Function sub_CheckTestConditionWaitData(ByVal dblMaxWait As Double) As Boolean

    '初期化
    '問題なければFalseが返る
    sub_CheckTestConditionWaitData = True

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_CheckTestConditionWaitData

    With PALS.LoopParams
        Dim cnt As Long
        '全カテゴリを繰り返し
        For cnt = 1 To .CategoryCount
            'フォームで指定した値以上のWaitが設定されていた場合、エラーを返す
            If .LoopCategory(cnt).WAIT > (dblMaxWait) Then
                sub_CheckTestConditionWaitData = True
                '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
                If g_RunAutoFlg_PALS = False Then
                    Call MsgBox("Error" & vbCrLf & "TestCondition no Wait ga saidaiti wo koeteimasu!" & vbCrLf & .LoopCategory(cnt).category & vbCrLf & "ErCode.2-2-18-5-26", vbExclamation)
                End If
                '<<<2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
            End If
        Next cnt
    End With

Exit Function

errPALSsub_CheckTestConditionWaitData:
    Call sub_errPALS("Check TestCondition Wait Data error at 'sub_CheckTestConditionWaitData'", "2-2-18-0-27")

End Function


'********************************************************************************************
' 名前: sub_Validate
' 内容: IG-XLのバリデーションを行う。
'       IG-XLでシートを追加する際に、バリデーションを行わないとエラーが発生する為使用している。
' 引数: なし
' 戻値: なし
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_Validate()
On Error GoTo errPALSsub_Validate
    
    TheExec.Validate

Exit Sub

errPALSsub_Validate:
    Call sub_errPALS("IG-XL Validate error at 'sub_Validate'", "2-2-19-0-28")

End Sub


'********************************************************************************************
' 名前: sub_GetMeasureData
' 内容: LOOPツールによって作成されたデータログの下部にあるデータを、引数で渡された1行分のデータログから
'       読み取る為の関数｡
' 引数: strBuf     :データが入っている1行分のデータログ
'     : strGetType :検索するデータの種類
' 戻値: 1行分のデータログから抜き出した各種類のデータ
' 備考： なし
' 更新履歴： Rev1.0      2010/09/30　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Function sub_GetMeasureData(ByVal strbuf As String, ByVal strGetType As String) As String

On Error GoTo errPALSsub_GetMeasureData

    Select Case strGetType

        Case "Date"
            sub_GetMeasureData = Mid$(strbuf, 16)
        
        Case "JobName"
            sub_GetMeasureData = Mid$(strbuf, 16)

        Case "Node"
            sub_GetMeasureData = Mid$(strbuf, 16)

    End Select

Exit Function

errPALSsub_GetMeasureData:
    Call sub_errPALS("Get Measure Data Error at 'sub_GetMeasureData'", "2-2-20-0-29")

End Function

'>>>2010/12/13 K.SUMIYASHIKI ADD FUNCTION
'********************************************************************************************
' 名前: sub_JudgeBaratuki
' 内容: 傾向が飛び値と判断されるような大きなバラツキか判断する関数
'       次の測定データとの差分(絶対値)を積算し、差分の平均値が1σ以上あれば、
'       大きなバラツキがあると判断する。
' 引数: lngTestNo         : 項目を示す番号
'       sitez             : サイト番号
'       lngNowLoopCnt     : 測定済み回数
' 戻値: バラツキ傾向を示す列挙体
' 備考： なし
' 更新履歴： Rev1.0      2010/12/13　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_JudgeBaratuki(ByVal lngTestNo As Long, ByVal sitez As Long, ByVal lngNowLoopCnt As Long) As enum_DataTrendType

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSsub_JudgeBaratuki

'飛び値と判断される可能性のある、大きなバラツキがないかの判断
'->次測定の特性値との差分の平均を求め、値が1σ以上の場合、大きなバラツキがあると判断する

    Dim dblSumDelta As Double   '次のデータとの差分を積算する為の変数
    Dim data_cnt As Long        'ループカウンタ(データインデックスを示す)
    
    dblSumDelta = 0
    
    With PALS.CommonInfo.TestInfo(lngTestNo).site(sitez)
    
        'データ数マイナス1だけ繰り返し
        For data_cnt = 1 To lngNowLoopCnt - 1
            '次のデータとの差分の絶対値を取得
            dblSumDelta = dblSumDelta + Abs(.Data(data_cnt + 1) - .Data(data_cnt))
        Next data_cnt
    
        '次のデータとの差分平均が0.9σ以上の場合バラツキとする
        If (dblSumDelta / (lngNowLoopCnt - 1)) > (.Sigma * 0.9) Then
            Debug.Print ("Big Baratuki!")
            Debug.Print ("TestName : " & PALS.CommonInfo.TestInfo(lngTestNo).tname)
            Debug.Print ("Site     : " & sitez) & vbCrLf            'バラツキを示す値を返し終了
            sub_JudgeBaratuki = em_trend_Uneven
        End If
    End With
Exit Function

errPALSsub_JudgeBaratuki:
    Call sub_errPALS("Check LoopData error at 'sub_JudgeBaratuki'", "2-2-21-0-30")

End Function



Public Function sub_CheckTestInstancesParams() As Boolean

    sub_CheckTestInstancesParams = True

    Dim CategoryCnt As Long
    Dim TestItemCnt As Long
    Dim strTmpCategoryName1 As String
    Dim strTmpCategoryName2 As String

    Dim blnCategory1_OK As Boolean
    Dim blnCategory2_OK As Boolean

    With PALS
        For TestItemCnt = 0 To .CommonInfo.TestCount
            strTmpCategoryName1 = .CommonInfo.TestInfo(TestItemCnt).CapCategory1
            strTmpCategoryName2 = .CommonInfo.TestInfo(TestItemCnt).CapCategory2
            
            If strTmpCategoryName1 = "" Or strTmpCategoryName1 = "DC" Then
                blnCategory1_OK = True
            Else
                blnCategory1_OK = False
            End If
            
            If strTmpCategoryName2 = "" Or strTmpCategoryName1 = "DC" Then
                blnCategory2_OK = True
            Else
                blnCategory2_OK = False
            End If
            
            If blnCategory1_OK = False Or blnCategory2_OK = False Then
                
                For CategoryCnt = 1 To .LoopParams.CategoryCount
                    If blnCategory1_OK = False And strTmpCategoryName1 = .LoopParams.LoopCategory(CategoryCnt).category Then
                        blnCategory1_OK = True
                    ElseIf blnCategory2_OK = False And strTmpCategoryName2 = .LoopParams.LoopCategory(CategoryCnt).category Then
                        blnCategory2_OK = True
                    End If
                Next CategoryCnt
            
                If blnCategory1_OK = False And blnCategory2_OK = False Then
                    '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
                    If g_RunAutoFlg_PALS = False Then
                        Call MsgBox(.CommonInfo.TestInfo(TestItemCnt).tname & " CapCategory1 and 2 is wrong!!" & vbCrLf & "Please Chack CapCategory." & vbCrLf & "ErrCode.2-2-22-4-31", vbExclamation)
                    End If
                    sub_CheckTestInstancesParams = False
                    Exit Function
                ElseIf blnCategory1_OK = False Then
                    If g_RunAutoFlg_PALS = False Then
                        Call MsgBox(.CommonInfo.TestInfo(TestItemCnt).tname & " CapCategory1 is wrong!!" & vbCrLf & "Please Chack CapCategory." & vbCrLf & "ErrCode.2-2-22-4-32", vbExclamation)
                    End If
                    sub_CheckTestInstancesParams = False
                    Exit Function
                ElseIf blnCategory2_OK = False Then
                    If g_RunAutoFlg_PALS = False Then
                        Call MsgBox(.CommonInfo.TestInfo(TestItemCnt).tname & " CapCategory2 is wrong!!" & vbCrLf & "Please Chack CapCategory." & vbCrLf & "ErrCode.2-2-22-4-33", vbExclamation)
                    End If
                    sub_CheckTestInstancesParams = False
                    Exit Function
                End If
            End If
        Next TestItemCnt
    End With
    
End Function


Public Sub sub_LoopParamsCheck()

    Dim cnt As Long
    
    With PALS.LoopParams
        For cnt = 1 To .CategoryCount
            If .LoopCategory(cnt).Average = 511 Then
                '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
                If g_RunAutoFlg_PALS = False Then
                    Call MsgBox(.LoopCategory(cnt).category & "Average count is 511..." & vbCrLf & "Please check average!" & vbCrLf & "ErrCode.2-2-23-5-34", vbExclamation)
                End If
'            ElseIf .LoopCategory(cnt).Wait > ChangeParamsInfo(cnt).MaxWait * 0.99 Then
            ElseIf .LoopCategory(cnt).WAIT > val(frm_PALS_LoopAdj_Main.txt_maxwait) * 0.99 Then
                If g_RunAutoFlg_PALS = False Then
                    Call MsgBox(.LoopCategory(cnt).category & "Wait is max..." & vbCrLf & "Please check average!" & vbCrLf & "ErrCode.2-2-23-5-35", vbExclamation)
                End If
            End If
        Next cnt
    End With

End Sub


'>>>2011/06/20 K.SUMIYASHIKI ADD FUNCTION
'********************************************************************************************
' 名前: sub_AddSheet
' 内容: 指定したワークシート名のシートを追加する。同じシートがあれば、末尾をインクリメントし追加する
' 引数: strSheetName   : ワークシート名
'       sitez          : サイト
' 戻値: なし
' 備考：なし
' 更新履歴： Rev1.0      2011/06/20　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_AddSheet(ByVal strSheetName As String, ByVal sitez As Integer)

    Sheets.Add.Name = "TempAddSheet"
    
    Dim intSheetCheck As Long
    intSheetCheck = 0 '枝番初期値
    On Error Resume Next
    Do
        Err.Clear
        If intSheetCheck = 0 Then
            ActiveSheet.Name = strSheetName & sitez
        Else
            ActiveSheet.Name = strSheetName & sitez & "(" & intSheetCheck & ")"
        End If
        intSheetCheck = intSheetCheck + 1
    Loop Until Err.Number = 0
    On Error GoTo 0

End Sub


'>>>2011/06/20 K.SUMIYASHIKI ADD FUNCTION
'********************************************************************************************
' 名前: sub_Get_F_Value
' 内容: 指定した測定回数、有意水準、上限/下限でのF値データをテーブルから取得する
' 引数: MeasureCnt      : 測定回数
'       SigmaNum        : 有意水準
'       TopOrBottom     : 上限or下限("top"or"bottom"で指定)
' 戻値: F値データ
' 備考： なし
' 更新履歴： Rev1.0      2011/06/20　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_Get_F_Value(ByVal MeasureCnt As Integer, ByVal SigmaNum As String, ByVal TopOrBottom As String) As Double

    Dim F_Table() As Double     'F値テーブルから取得したデータ配列を格納
    
    'F値テーブルより配列データを取得
    F_Table = sub_Get_F_Table(SigmaNum, TopOrBottom)
    
    '測定回数が100回以上の場合は、100回時のF値で近似
    If MeasureCnt > 100 Then
        sub_Get_F_Value = F_Table(100)
    Else
        '指定測定回数時のF値を返す
        sub_Get_F_Value = F_Table(MeasureCnt)
    End If


End Function


'>>>2011/06/20 K.SUMIYASHIKI ADD FUNCTION
'********************************************************************************************
' 名前: sub_Get_F_Table
' 内容: 指定した測定回数、有意水準、上限/下限でのF値データをテーブルから取得する
' 引数: SigmaNum        : 有意水準
'       TopOrBottom     : 上限or下限("top"or"bottom"で指定)
' 戻値: F値データ配列
' 備考： なし
' 更新履歴： Rev1.0      2011/06/20　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_Get_F_Table(ByVal SigmaNum As String, ByVal TopOrBottom As String) As Double()

    Dim DataTable(100) As Double        'F値データを格納する配列

    '有意水準3σの場合
    If SigmaNum = 3 Then
    
        '下限データ
        If TopOrBottom = "bottom" Then
            DataTable(3) = 0.001503
            DataTable(4) = 0.008852
            DataTable(5) = 0.021938
            DataTable(6) = 0.038185
            DataTable(7) = 0.055678
            DataTable(8) = 0.07335
            DataTable(9) = 0.090656
            DataTable(10) = 0.107334
            DataTable(11) = 0.123277
            DataTable(12) = 0.138451
            DataTable(13) = 0.152869
            DataTable(14) = 0.166561
            DataTable(15) = 0.179568
            DataTable(16) = 0.191932
            DataTable(17) = 0.203697
            DataTable(18) = 0.214907
            DataTable(19) = 0.225599
            DataTable(20) = 0.235812
            DataTable(21) = 0.245578
            DataTable(22) = 0.25493
            DataTable(23) = 0.263895
            DataTable(24) = 0.2725
            DataTable(25) = 0.280768
            DataTable(26) = 0.288721
            DataTable(27) = 0.296379
            DataTable(28) = 0.30376
            DataTable(29) = 0.310882
            DataTable(30) = 0.317758
            DataTable(31) = 0.324405
            DataTable(32) = 0.330833
            DataTable(33) = 0.337057
            DataTable(34) = 0.343085
            DataTable(35) = 0.34893
            DataTable(36) = 0.354601
            DataTable(37) = 0.360105
            DataTable(38) = 0.365453
            DataTable(39) = 0.37065
            DataTable(40) = 0.375705
            DataTable(41) = 0.380624
            DataTable(42) = 0.385413
            DataTable(43) = 0.390079
            DataTable(44) = 0.394626
            DataTable(45) = 0.399059
            DataTable(46) = 0.403385
            DataTable(47) = 0.407606
            DataTable(48) = 0.411728
            DataTable(49) = 0.415754
            DataTable(50) = 0.419688
            DataTable(51) = 0.423534
            DataTable(52) = 0.427295
            DataTable(53) = 0.430975
            DataTable(54) = 0.434576
            DataTable(55) = 0.438101
            DataTable(56) = 0.441553
            DataTable(57) = 0.444935
            DataTable(58) = 0.448249
            DataTable(59) = 0.451498
            DataTable(60) = 0.454682
            DataTable(61) = 0.457806
            DataTable(62) = 0.460871
            DataTable(63) = 0.463878
            DataTable(64) = 0.466829
            DataTable(65) = 0.469727
            DataTable(66) = 0.472573
            DataTable(67) = 0.475369
            DataTable(68) = 0.478115
            DataTable(69) = 0.480814
            DataTable(70) = 0.483467
            DataTable(71) = 0.486076
            DataTable(72) = 0.488641
            DataTable(73) = 0.491163
            DataTable(74) = 0.493645
            DataTable(75) = 0.496087
            DataTable(76) = 0.498491
            DataTable(77) = 0.500856
            DataTable(78) = 0.503185
            DataTable(79) = 0.505478
            DataTable(80) = 0.507737
            DataTable(81) = 0.509961
            DataTable(82) = 0.512153
            DataTable(83) = 0.514313
            DataTable(84) = 0.516441
            DataTable(85) = 0.518538
            DataTable(86) = 0.520606
            DataTable(87) = 0.522645
            DataTable(88) = 0.524655
            DataTable(89) = 0.526638
            DataTable(90) = 0.528593
            DataTable(91) = 0.530522
            DataTable(92) = 0.532426
            DataTable(93) = 0.534304
            DataTable(94) = 0.536157
            DataTable(95) = 0.537987
            DataTable(96) = 0.539793
            DataTable(97) = 0.541575
            DataTable(98) = 0.543336
            DataTable(99) = 0.545074
            DataTable(100) = 0.546791


        '上限データ
        ElseIf TopOrBottom = "top" Then
            DataTable(3) = 222221.722183
            DataTable(4) = 665.833264
            DataTable(5) = 104.378009
            DataTable(6) = 42.000687
            DataTable(7) = 24.318542
            DataTable(8) = 16.822538
            DataTable(9) = 12.868018
            DataTable(10) = 10.479585
            DataTable(11) = 8.899609
            DataTable(12) = 7.784354
            DataTable(13) = 6.958128
            DataTable(14) = 6.322781
            DataTable(15) = 5.819582
            DataTable(16) = 5.411412
            DataTable(17) = 5.073739
            DataTable(18) = 4.789744
            DataTable(19) = 4.547531
            DataTable(20) = 4.338458
            DataTable(21) = 4.156104
            DataTable(22) = 3.995602
            DataTable(23) = 3.853195
            DataTable(24) = 3.725944
            DataTable(25) = 3.611509
            DataTable(26) = 3.508013
            DataTable(27) = 3.413927
            DataTable(28) = 3.327995
            DataTable(29) = 3.249176
            DataTable(30) = 3.1766
            DataTable(31) = 3.109535
            DataTable(32) = 3.047358
            DataTable(33) = 2.989539
            DataTable(34) = 2.93562
            DataTable(35) = 2.885207
            DataTable(36) = 2.837959
            DataTable(37) = 2.793575
            DataTable(38) = 2.751795
            DataTable(39) = 2.712388
            DataTable(40) = 2.675149
            DataTable(41) = 2.639898
            DataTable(42) = 2.606473
            DataTable(43) = 2.574731
            DataTable(44) = 2.544542
            DataTable(45) = 2.51579
            DataTable(46) = 2.488372
            DataTable(47) = 2.462192
            DataTable(48) = 2.437164
            DataTable(49) = 2.413212
            DataTable(50) = 2.390263
            DataTable(51) = 2.368254
            DataTable(52) = 2.347124
            DataTable(53) = 2.326821
            DataTable(54) = 2.307293
            DataTable(55) = 2.288496
            DataTable(56) = 2.270387
            DataTable(57) = 2.252927
            DataTable(58) = 2.23608
            DataTable(59) = 2.219812
            DataTable(60) = 2.204094
            DataTable(61) = 2.188895
            DataTable(62) = 2.17419
            DataTable(63) = 2.159953
            DataTable(64) = 2.146162
            DataTable(65) = 2.132794
            DataTable(66) = 2.119829
            DataTable(67) = 2.107249
            DataTable(68) = 2.095035
            DataTable(69) = 2.083171
            DataTable(70) = 2.071642
            DataTable(71) = 2.060431
            DataTable(72) = 2.049527
            DataTable(73) = 2.038915
            DataTable(74) = 2.028583
            DataTable(75) = 2.01852
            DataTable(76) = 2.008715
            DataTable(77) = 1.999157
            DataTable(78) = 1.989837
            DataTable(79) = 1.980746
            DataTable(80) = 1.971873
            DataTable(81) = 1.963212
            DataTable(82) = 1.954755
            DataTable(83) = 1.946493
            DataTable(84) = 1.93842
            DataTable(85) = 1.930528
            DataTable(86) = 1.922813
            DataTable(87) = 1.915266
            DataTable(88) = 1.907883
            DataTable(89) = 1.900658
            DataTable(90) = 1.893585
            DataTable(91) = 1.88666
            DataTable(92) = 1.879877
            DataTable(93) = 1.873232
            DataTable(94) = 1.866721
            DataTable(95) = 1.860339
            DataTable(96) = 1.854083
            DataTable(97) = 1.847947
            DataTable(98) = 1.84193
            DataTable(99) = 1.836026
            DataTable(100) = 1.830233
        
        Else
            MsgBox ("Program Argument Error!!")
                
        End If

    '有意水準2σの場合
    ElseIf SigmaNum = 2 Then
        
        '下限データ
        If TopOrBottom = "bottom" Then
            DataTable(3) = 0.02597
            DataTable(4) = 0.062328
            DataTable(5) = 0.100208
            DataTable(6) = 0.135357
            DataTable(7) = 0.167013
            DataTable(8) = 0.195366
            DataTable(9) = 0.220821
            DataTable(10) = 0.243786
            DataTable(11) = 0.264623
            DataTable(12) = 0.283634
            DataTable(13) = 0.30107
            DataTable(14) = 0.317141
            DataTable(15) = 0.332017
            DataTable(16) = 0.345844
            DataTable(17) = 0.358742
            DataTable(18) = 0.370814
            DataTable(19) = 0.382148
            DataTable(20) = 0.392818
            DataTable(21) = 0.402889
            DataTable(22) = 0.412416
            DataTable(23) = 0.421449
            DataTable(24) = 0.430031
            DataTable(25) = 0.438199
            DataTable(26) = 0.445986
            DataTable(27) = 0.453423
            DataTable(28) = 0.460536
            DataTable(29) = 0.467348
            DataTable(30) = 0.473882
            DataTable(31) = 0.480155
            DataTable(32) = 0.486186
            DataTable(33) = 0.491991
            DataTable(34) = 0.497583
            DataTable(35) = 0.502976
            DataTable(36) = 0.508182
            DataTable(37) = 0.513211
            DataTable(38) = 0.518074
            DataTable(39) = 0.522781
            DataTable(40) = 0.527339
            DataTable(41) = 0.531757
            DataTable(42) = 0.536041
            DataTable(43) = 0.540199
            DataTable(44) = 0.544237
            DataTable(45) = 0.548161
            DataTable(46) = 0.551977
            DataTable(47) = 0.555688
            DataTable(48) = 0.559301
            DataTable(49) = 0.562819
            DataTable(50) = 0.566247
            DataTable(51) = 0.569589
            DataTable(52) = 0.572848
            DataTable(53) = 0.576028
            DataTable(54) = 0.579132
            DataTable(55) = 0.582163
            DataTable(56) = 0.585124
            DataTable(57) = 0.588018
            DataTable(58) = 0.590847
            DataTable(59) = 0.593614
            DataTable(60) = 0.596321
            DataTable(61) = 0.598971
            DataTable(62) = 0.601564
            DataTable(63) = 0.604105
            DataTable(64) = 0.606593
            DataTable(65) = 0.609032
            DataTable(66) = 0.611422
            DataTable(67) = 0.613765
            DataTable(68) = 0.616064
            DataTable(69) = 0.618319
            DataTable(70) = 0.620531
            DataTable(71) = 0.622703
            DataTable(72) = 0.624835
            DataTable(73) = 0.626929
            DataTable(74) = 0.628985
            DataTable(75) = 0.631006
            DataTable(76) = 0.632991
            DataTable(77) = 0.634942
            DataTable(78) = 0.636861
            DataTable(79) = 0.638747
            DataTable(80) = 0.640602
            DataTable(81) = 0.642427
            DataTable(82) = 0.644222
            DataTable(83) = 0.645989
            DataTable(84) = 0.647728
            DataTable(85) = 0.649439
            DataTable(86) = 0.651125
            DataTable(87) = 0.652784
            DataTable(88) = 0.654419
            DataTable(89) = 0.656029
            DataTable(90) = 0.657615
            DataTable(91) = 0.659178
            DataTable(92) = 0.660719
            DataTable(93) = 0.662237
            DataTable(94) = 0.663734
            DataTable(95) = 0.66521
            DataTable(96) = 0.666665
            DataTable(97) = 0.668101
            DataTable(98) = 0.669517
            DataTable(99) = 0.670913
            DataTable(100) = 0.672291


        '上限データ
        ElseIf TopOrBottom = "top" Then
            DataTable(3) = 799.5
            DataTable(4) = 39.165495
            DataTable(5) = 15.100979
            DataTable(6) = 9.364471
            DataTable(7) = 6.977702
            DataTable(8) = 5.69547
            DataTable(9) = 4.899341
            DataTable(10) = 4.357233
            DataTable(11) = 3.963865
            DataTable(12) = 3.664914
            DataTable(13) = 3.429613
            DataTable(14) = 3.239263
            DataTable(15) = 3.081854
            DataTable(16) = 2.949321
            DataTable(17) = 2.836047
            DataTable(18) = 2.737998
            DataTable(19) = 2.652204
            DataTable(20) = 2.576425
            DataTable(21) = 2.508943
            DataTable(22) = 2.448414
            DataTable(23) = 2.393775
            DataTable(24) = 2.344171
            DataTable(25) = 2.298907
            DataTable(26) = 2.257412
            DataTable(27) = 2.219213
            DataTable(28) = 2.183913
            DataTable(29) = 2.15118
            DataTable(30) = 2.120728
            DataTable(31) = 2.092317
            DataTable(32) = 2.065736
            DataTable(33) = 2.040804
            DataTable(34) = 2.017366
            DataTable(35) = 1.995283
            DataTable(36) = 1.974435
            DataTable(37) = 1.954715
            DataTable(38) = 1.936029
            DataTable(39) = 1.918292
            DataTable(40) = 1.901431
            DataTable(41) = 1.885377
            DataTable(42) = 1.870071
            DataTable(43) = 1.855459
            DataTable(44) = 1.841492
            DataTable(45) = 1.828124
            DataTable(46) = 1.815317
            DataTable(47) = 1.803033
            DataTable(48) = 1.791239
            DataTable(49) = 1.779903
            DataTable(50) = 1.769
            DataTable(51) = 1.758501
            DataTable(52) = 1.748384
            DataTable(53) = 1.738628
            DataTable(54) = 1.729211
            DataTable(55) = 1.720115
            DataTable(56) = 1.711323
            DataTable(57) = 1.702819
            DataTable(58) = 1.694588
            DataTable(59) = 1.686616
            DataTable(60) = 1.678891
            DataTable(61) = 1.671399
            DataTable(62) = 1.664131
            DataTable(63) = 1.657075
            DataTable(64) = 1.650222
            DataTable(65) = 1.643562
            DataTable(66) = 1.637087
            DataTable(67) = 1.630789
            DataTable(68) = 1.62466
            DataTable(69) = 1.618692
            DataTable(70) = 1.61288
            DataTable(71) = 1.607216
            DataTable(72) = 1.601695
            DataTable(73) = 1.59631
            DataTable(74) = 1.591057
            DataTable(75) = 1.58593
            DataTable(76) = 1.580925
            DataTable(77) = 1.576037
            DataTable(78) = 1.571261
            DataTable(79) = 1.566594
            DataTable(80) = 1.562031
            DataTable(81) = 1.557569
            DataTable(82) = 1.553204
            DataTable(83) = 1.548933
            DataTable(84) = 1.544753
            DataTable(85) = 1.54066
            DataTable(86) = 1.536652
            DataTable(87) = 1.532725
            DataTable(88) = 1.528878
            DataTable(89) = 1.525107
            DataTable(90) = 1.521411
            DataTable(91) = 1.517786
            DataTable(92) = 1.514231
            DataTable(93) = 1.510743
            DataTable(94) = 1.50732
            DataTable(95) = 1.503961
            DataTable(96) = 1.500664
            DataTable(97) = 1.497426
            DataTable(98) = 1.494246
            DataTable(99) = 1.491123
            DataTable(100) = 1.488054
        Else
            
            MsgBox ("Program Argument Error!!")
        End If
    
    Else
        MsgBox ("Program Argument Error!!")
    
    End If

    '配列データを引数で返す
    sub_Get_F_Table = DataTable

End Function

'********************************************************************************************
' 名前: sub_RunLoopAuto
' 内容: 指定した測定回数・ノード情報で、ループ測定を自動で実施する
' 引数: lngLoopCnt    : ループ回数
'       intSwNode     : テスタノード
' 戻値: 終了フラグ
' 備考： なし
' 更新履歴： Rev1.0      2011/08/10　新規作成   K.Sumiyashiki
' 更新履歴： Rev2.0      2012/03/08　光量との連携機能追加   M.Imamura
'********************************************************************************************

Public Function sub_RunLoopAuto(ByVal lngLoopCnt As Long, ByVal intSwNode As Integer, Optional ByVal blnRunMode As Boolean = False, Optional blnDataMode As Boolean = True, Optional intMaxWait As Integer = 500, Optional intMaxTrialCount As Integer = 1) As Long

    sub_RunLoopAuto = 1
    Sw_Node = intSwNode

On Error GoTo errPALSsub_RunLoopAuto
    
    ThisWorkbook.Activate

    PALS_ParamFolder = ThisWorkbook.Path & "\" & PALS_PARAMFOLDERNAME
    Call sub_PalsFileCheck

    Set PALS = Nothing
    Set PALS = New csPALS

    '連携時フラグ 2012/6/19
    g_RunAutoFlg_PALS = True

    'TestConditionシートデータの再読込
    Call ReadCategoryData

    With frm_PALS_LoopAdj_Main
        .Show vbModeless
        .txt_loop_num.Value = lngLoopCnt
        
        If blnRunMode = False Then
            .op_NotAdjust.Value = True
            .op_AutoAdjust.Value = False
        Else
            .op_NotAdjust.Value = False
            .op_AutoAdjust.Value = True
        End If
        
        .Btn_ContinueOnFail.Value = blnDataMode
        .txt_maxwait = intMaxWait
        .txt_maxtrial_num = intMaxTrialCount
        Call .cmd_start_Click
    End With

    Unload frm_PALS_LoopAdj_Main

    If g_ErrorFlg_PALS = True Then
        GoTo errPALSsub_RunLoopAuto
    End If
    
    g_RunAutoFlg_PALS = False
    sub_RunLoopAuto = 0
    Exit Function

errPALSsub_RunLoopAuto:
    g_RunAutoFlg_PALS = False
    g_ErrorFlg_PALS = False

End Function





