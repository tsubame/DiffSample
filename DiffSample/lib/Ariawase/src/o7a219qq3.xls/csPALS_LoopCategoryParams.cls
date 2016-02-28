VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_LoopCategoryParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' クラス名  ：  csPALS_LoopParamsMenber.cls
' 概要      ：  測定条件の格納・処理を行うクラス
'               =>Category、Average回数、Wait
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/10　新規作成        K.Sumiyashiki
'==========================================================================

Public category As String                       'カテゴリ名(ex:ML,OF,etc..)
Public Average  As Integer                      '取り込み回数
Public WAIT     As Double                       '取り込み前ウェイト
Public mode     As String                       '取り込み回数の倍数指定設定(自動の場合Auto)

Public TargetTestName As String                 '最大バラツキ項目名
Public VariationLevel As Double                 '最大バラツキ量(3σ/規格幅)
Public VariationSite  As Integer                '最大バラツキSite
Public VariationTrend As enum_DataTrendType     'バラツキ傾向


'*********************************************************************************
' 関数名    ：  Class_Initialize
' 概要      ：  インスタンス生成時の初期化
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub init_csPALS_LoopCategoryParams()

On Error GoTo errPALSinit_csPALS_LoopCategoryParams

    category = ""
    Average = -1
    WAIT = -1               '0以上の値が有り得るので、-1で初期化。他の変数も統一して-1で初期化。
    mode = ""

    TargetTestName = ""
    VariationLevel = -1
    VariationSite = -1
    VariationTrend = em_trend_None

Exit Sub

errPALSinit_csPALS_LoopCategoryParams:
    Call sub_errPALS("csPALS_LoopCategoryParams not initialize at 'init_csPALS_LoopCategoryParams'", "2-3-01-0-36")

End Sub


'*********************************************************************************
' 関数名    ：  Class_Terminate
' 概要      ：  インスタンス開放時の処理
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()
    'クラスのインスタンスが生成された際に実行される関数
    '初期化関数の呼び出し
    Call init_csPALS_LoopCategoryParams
End Sub



