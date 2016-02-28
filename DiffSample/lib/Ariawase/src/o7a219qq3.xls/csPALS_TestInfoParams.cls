VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_TestInfoParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' クラス名  ：  csLoopPTestInfo.cls
' 概要      ：  測定データの格納、処理を行うクラス
'               =>data,Min,Ave,Max,σ
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/10　新規作成        K.Sumiyashiki
'           ：  Rev2.0      2011/05/16　メンバ追加      K.Sumiyashiki
'                                       ⇒Enable, ActiveValueCnt追加
'           ：  Rev3.0      2011/06/17　関数追加        K.Sumiyashiki
'                                       ⇒sub_CalcDifferentialSigma追加
'==========================================================================

'特性値データ
Private m_Min   As Double               '最小値
Private m_Ave   As Double               '平均値
Private m_Max   As Double               '最大値
Private m_Sigma As Double               'シグマ
Private m_testdata() As Double          '特性値データ
'>>>2011/05/12 K.SUMIYASHIKI ADD
Private m_PassOrFail() As Boolean       'Result(PASS/FAIL)情報
Private m_ActiveValueCnt As Integer     'このテスト項目をPASSした回数
'<<<2011/05/12 K.SUMIYASHIKI ADD


'********************************************************************************************
' 名前: init_csPALS_TestInfoParams
' 内容: 各変数の初期化
' 引数: なし
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub init_csPALS_TestInfoParams()
'クラス初期化
    m_Min = 0
    m_Ave = 0
    m_Max = 0
    m_Sigma = 0
'>>>2011/06/20 K.SUMIYASHIKI ADD
    m_ActiveValueCnt = 0
'>>>2011/06/20 K.SUMIYASHIKI ADD
End Sub


'*********************************************************************************
' 関数名    ：  Class_Initialize
' 概要      ：  インスタンス生成時の初期化
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()
    '初期化関数の呼び出し
    Call init_csPALS_TestInfoParams
End Sub


'*********************************************************************************
' 関数名    ：  Class_Terminate
' 概要      ：  インスタンス開放時の処理
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Terminate()

    Erase m_testdata

End Sub


'*********************************************************************************
' プロパティ名  ：  Data
' 概要          ：  各Siteの特性値情報を返すプロパティ
' 引数          ：  index :測定回数
' 備考          ：  なし
' 更新履歴      ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Public Property Get Data(index As Long) As Double
    
'    If g_ErrorFlg_PALS Then
'        Exit Sub
'    End If
    
On Error GoTo errPALSDataGet

    'ここにdataプロパティの値を取得する処理を書く
    If index > g_MaxPalsCount Then
        Call sub_errPALS("Data count can't set over " & g_MaxPalsCount, "0-7-01-5-26")
        Exit Property
    End If
    
    Data = m_testdata(index)

Exit Property

errPALSDataGet:
    Call sub_errPALS("Get measure data error at 'csPALS_TestInfoParams -> Get Data'", "0-7-01-0-27")

End Property

Public Property Let Data(index As Long, ByVal testdata As Double)
    
'    If g_ErrorFlg_PALS Then
'        Exit Property
'    End If
    
On Error GoTo errPALSDataLet
    
    'ここにdataプロパティに値をセットする処理を書く
    If index > g_MaxPalsCount Then
        Call sub_errPALS("Data count can't set over  " & g_MaxPalsCount, "0-7-02-5-28")
        Exit Property
    End If
    
    m_testdata(index) = testdata
    
'>>>2011/05/12 K.SUMIYASHIKI ADD
    m_ActiveValueCnt = m_ActiveValueCnt + 1
'<<<2011/05/12 K.SUMIYASHIKI ADD

    '平均、最大、最小、σを計算
    Call calc_ave(index, testdata)
    Call calc_max(index, testdata)
    Call calc_min(index, testdata)
    Call calc_sigma(index)
    
Exit Property

errPALSDataLet:
    Call sub_errPALS("Input measure data error at 'csPALS_TestInfoParams -> Let Data'", "0-7-02-0-29")
    
End Property


'********************************************************************************************
' 名前: DataCount
' 内容: 最大測定回数を返す関数
' 引数: なし
' 戻値: 最大測定回数
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Function DataCount() As Integer

    DataCount = UBound(m_testdata)

End Function


'*********************************************************************************
' プロパティ名  ：  Ave
' 概要          ：  指定項目の平均値を返すプロパティ
' 引数          ：  なし
' 備考          ：  読み取り専用
' 更新履歴      ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Public Property Get ave() As Double

    ave = m_Ave

End Property


'*********************************************************************************
' プロパティ名  ：  Max
' 概要          ：  指定項目の最大値を返すプロパティ
' 引数          ：  なし
' 備考          ：  読み取り専用
' 更新履歴      ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Public Property Get max() As Double

    max = m_Max

End Property


'*********************************************************************************
' プロパティ名  ：  Min
' 概要          ：  指定項目の最小値を返すプロパティ
' 引数          ：  なし
' 備考          ：  読み取り専用
' 更新履歴      ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Public Property Get Min() As Double

    Min = m_Min

End Property


'*********************************************************************************
' プロパティ名  ：  Sigma
' 概要          ：  指定項目のσを返すプロパティ
' 引数          ：  なし
' 備考          ：  読み取り専用
' 更新履歴      ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Public Property Get Sigma() As Double

    Sigma = m_Sigma

End Property


'********************************************************************************************
' 名前: calc_ave
' 内容: 指定項目の平均値を計算する関数
' 引数: mes_cnt   : 測定回数
'       testdata  : 最新の測定データ
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub calc_ave(ByVal mes_cnt As Long, ByVal testdata As Double)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALScalc_ave

'>>>2011/06/20 K.SUMIYASHIKI UPDATE
'    If mes_cnt = 1 Then
    If m_ActiveValueCnt = 1 Then
        m_Ave = testdata
'<<<2011/06/20 K.SUMIYASHIKI UPDATE
    
    Else
'>>>2011/05/12 K.SUMIYASHIKI CHANGE
'old        m_Ave = (m_Ave * (mes_cnt - 1) + testdata) / mes_cnt
        m_Ave = (m_Ave * (m_ActiveValueCnt - 1) + testdata) / m_ActiveValueCnt
'<<<2011/05/12 K.SUMIYASHIKI CHANGE
    End If

Exit Sub

errPALScalc_ave:
    Call sub_errPALS("Calculate average error at 'csPALS_TestInfoParams -> calc_ave'", "0-7-03-0-30")

End Sub


'********************************************************************************************
' 名前: calc_max
' 内容: 指定項目の最大値を計算する関数
' 引数: mes_cnt   : 測定回数
'       testdata  : 最新の測定データ
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub calc_max(ByVal mes_cnt As Long, ByVal testdata As Double)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALScalc_max

'>>>2011/06/20 K.SUMIYASHIKI UPDATE
'    If mes_cnt = 1 Then
    If m_ActiveValueCnt = 1 Then
        m_Max = testdata
'<<<2011/06/20 K.SUMIYASHIKI UPDATE
    ElseIf testdata > m_Max Then
        m_Max = testdata
    Else
    
    End If

Exit Sub

errPALScalc_max:
    Call sub_errPALS("Calculate max data error at 'csPALS_TestInfoParams -> calc_max'", "0-7-04-0-31")

End Sub


'********************************************************************************************
' 名前: calc_min
' 内容: 指定項目の最小値を計算する関数
' 引数: mes_cnt   : 測定回数
'       testdata  : 最新の測定データ
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub calc_min(ByVal mes_cnt As Long, ByVal testdata As Double)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALScalc_min

'>>>2011/06/20 K.SUMIYASHIKI UPDATE
'    If mes_cnt = 1 Then
    If m_ActiveValueCnt = 1 Then
        m_Min = testdata
'<<<2011/06/20 K.SUMIYASHIKI UPDATE
    ElseIf m_Min > testdata Then
        m_Min = testdata
    Else
    
    End If

Exit Sub

errPALScalc_min:
    Call sub_errPALS("Calculate minimum data error at 'csPALS_TestInfoParams -> calc_min'", "0-7-05-0-32")

End Sub


'********************************************************************************************
' 名前: calc_sigma
' 内容: 指定項目のσを計算する関数
' 引数: mes_cnt   : 測定回数
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function calc_sigma(ByVal mes_cnt As Long) As Double

    If mes_cnt = 1 Or g_ErrorFlg_PALS Then
        m_Sigma = 0
        Exit Function
    End If
    
On Error GoTo errPALScalc_sigma

    m_Sigma = 0

    Dim sigma_i As Long
'>>>2011/05/12 K.SUMIYASHIKI CHANGE
    For sigma_i = 1 To mes_cnt
        If m_PassOrFail(sigma_i) = True Then
            m_Sigma = m_Sigma + ((m_Ave - m_testdata(sigma_i)) ^ 2)
        End If
    Next sigma_i
    
'old    m_Sigma = Sqr(m_Sigma / (mes_cnt))
    m_Sigma = Sqr(m_Sigma / (m_ActiveValueCnt))
'<<<2011/05/12 K.SUMIYASHIKI CHANGE

Exit Function

errPALScalc_sigma:
    Call sub_errPALS("Calculate sigma data error at 'csPALS_TestInfoParams -> calc_sigma'", "0-7-06-0-33")

End Function


'********************************************************************************************
' 名前: calc_sigma
' 内容: フォームで指定したループ回数に応じて、配列数を変更する関数
' 引数: lngCnt   : 最大測定回数
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_ChangeDataDivision(ByVal lngCnt As Long)

    ReDim m_testdata(lngCnt)

'>>>2011/05/12 K.SUMIYASHIKI ADD
    ReDim m_PassOrFail(lngCnt)
'<<<2011/05/12 K.SUMIYASHIKI ADD

End Sub


'*********************************************************************************
' プロパティ名  ：  ActiveValueCnt
' 概要          ：  PASSした項目数を返すプロパティ
' 引数          ：  なし
' 備考          ：  読み取り専用
' 更新履歴      ：  Rev1.0      2011/05/13　新規作成        K.Sumiyashiki
'*********************************************************************************
Public Property Get ActiveValueCnt() As Double

    ActiveValueCnt = m_ActiveValueCnt

End Property


'*********************************************************************************
' プロパティ名  ：  Enable
' 概要          ：  特性値のPASS or FAILを返すプロパティ
' 引数          ：  index :測定回数
' 備考          ：  なし
' 更新履歴      ：  Rev1.0      2011/05/16　新規作成        K.Sumiyashiki
'*********************************************************************************
Public Property Get Enable(index As Long) As Boolean

On Error GoTo errPALSEnableGet

    'ここにEnableプロパティの値を取得する処理を書く
    Enable = m_PassOrFail(index)

Exit Property

errPALSEnableGet:
    Call sub_errPALS("Get measure data error at 'csPALS_TestInfoParams -> Get Enable'", "0-7-07-0-34")

End Property

Public Property Let Enable(index As Long, ByVal ChangeData As Boolean)
    
On Error GoTo errPALSEnableLet
    
    'ここにEnableプロパティに値をセットする処理を書く
    m_PassOrFail(index) = ChangeData
    
Exit Property

errPALSEnableLet:
    Call sub_errPALS("Input measure data error at 'csPALS_TestInfoParams -> Let Enable'", "0-7-08-0-35")
    
End Property



'*********************************************************************************
' プロパティ名  ：  Differential_Sigma
' 概要          ：  指定項目の微分値のσを返すプロパティ
' 引数          ：  なし
' 備考          ：  読み取り専用
' 更新履歴      ：  Rev1.0      2011/06/17　新規作成   K.Sumiyashiki
'*********************************************************************************
Public Property Get Differential_Sigma(ByVal mes_cnt As Long) As Double

    Differential_Sigma = sub_CalcDifferentialSigma(mes_cnt)

End Property


'********************************************************************************************
' 名前: sub_CalcDifferentialSigma
' 内容: 指定項目の微分値のσを計算する関数
' 引数: mes_cnt   : 測定回数
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2011/06/17　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function sub_CalcDifferentialSigma(ByVal mes_cnt As Long) As Double

    If mes_cnt = 1 Then
        sub_CalcDifferentialSigma = 0
        Exit Function
    End If
    
On Error GoTo errPALSsub_CalcDifferentialSigma

    Dim dblDifferentialData() As Double     '微分値を格納する配列
    Dim sigma_i As Long                     'ループ変数
    Dim lngActiveDataCnt As Long            '有効な微分データ数
    Dim tmpStandardData As Double           '微分値取得時に前のデータを一時格納しておく変数

    '有効微分データ数の初期化(Redimの処理がやりやすくなるので、初期値が-1)
    lngActiveDataCnt = -1

    'データ数分繰り返し
    For sigma_i = 1 To mes_cnt
        '有効なデータの場合のみ処理を実施
        If m_PassOrFail(sigma_i) = True Then
        
            '有効データ1回目の処理
            If lngActiveDataCnt = -1 Then
                '微分値を取得する為に、データを一時格納
                tmpStandardData = m_testdata(sigma_i)
                '有効な微分データ数をインクリメント
                lngActiveDataCnt = lngActiveDataCnt + 1
            Else
                '微分データを格納する配列の再定義
                ReDim Preserve dblDifferentialData(lngActiveDataCnt)
                '微分データ取得
                dblDifferentialData(lngActiveDataCnt) = m_testdata(sigma_i) - tmpStandardData
                '次の測定データとの微分値を取得する為、データを一時格納
                tmpStandardData = m_testdata(sigma_i)
                '有効な微分データ数をインクリメント
                lngActiveDataCnt = lngActiveDataCnt + 1
            End If
        End If
    Next sigma_i
    
    '有効微分データが0の場合のエラー処理
    If lngActiveDataCnt < 1 Then
        sub_CalcDifferentialSigma = 0
        Exit Function
    End If
    
    
    Dim dblDifferentialAve As Double
    '変数初期化
    dblDifferentialAve = 0
    
    '有効データ数分繰り返し
    For sigma_i = 0 To lngActiveDataCnt - 1
        '合計値取得
        dblDifferentialAve = dblDifferentialAve + dblDifferentialData(sigma_i)
    Next sigma_i
    
    '平均値取得
    dblDifferentialAve = dblDifferentialAve / lngActiveDataCnt
        
        
    Dim dblSumData As Double
    '変数初期化
    dblSumData = 0
    
    'σの計算
    For sigma_i = 0 To lngActiveDataCnt - 1
        dblSumData = dblSumData + ((dblDifferentialAve - dblDifferentialData(sigma_i)) ^ 2)
    Next sigma_i
    
    sub_CalcDifferentialSigma = Sqr(dblSumData / lngActiveDataCnt)

Exit Function

errPALSsub_CalcDifferentialSigma:
    Call sub_errPALS("Calculate differential sigma data error at 'csPALS_TestInfoParams -> sub_CalcDifferentialSigma'", "0-7-06-0-36")
End Function
