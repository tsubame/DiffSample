VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEeeTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   PerformanceCounterを利用したタイマー機能の提供
'
'目的:
'   Privateタイマーの実装
'
'作成者:
'   SLSI今手
'
'注意点:
'

Option Explicit

'機能の準備
Private Declare Function QueryPerformanceCounter Lib "kernel32" (pPerformanceCount As Currency) As Boolean
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (pFrequency As Currency) As Boolean

Private Const ERR_NUMBER = 9999                'エラー時に渡すエラー番号
Private Const CLASS_NAME = "CEeeTimer"         'このクラスの名前

'時間記録用
Private mStartCount As Currency

'内部情報初期化処理
Public Sub Initialize()
    mStartCount = 0#
End Sub

'開始時間の取得
Public Sub StartTimer()
    Call QueryPerformanceCounter(mStartCount)
End Sub

'途中時間の取得
Public Function LapTimer() As Double
    LapTimer = CalculationTime(mStartCount)
End Function

'終了時間の算出
Public Function StopTimer() As Double
    StopTimer = CalculationTime(mStartCount)
    mStartCount = 0#
End Function

'時間算出処理
Private Function CalculationTime(ByVal pStartCount As Currency) As Double
    Dim curFrequency As Currency
    Dim EndCount As Currency
    
    If pStartCount <= 0# Then
        Call RaiseError("StartTimer is unexecution")
    Else
        Call QueryPerformanceCounter(EndCount)
        Call QueryPerformanceFrequency(curFrequency)
        If mStartCount > EndCount Then
            Call RaiseError("Time calculation error")
        Else
'            Debug.Print "END=" & EndCount & " Start=" & mStartCount
            CalculationTime = (EndCount - mStartCount) / curFrequency
        End If
    End If

End Function

'エラーレイズ用関数
Private Sub RaiseError(ByVal pErrMsg As String)
    If Not (TheError Is Nothing) Then
        'エラーマネージャへ
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME, pErrMsg)
    Else
        'VBのErrへ
        Call Err.Raise(ERR_NUMBER, CLASS_NAME, pErrMsg)
    End If
End Sub

