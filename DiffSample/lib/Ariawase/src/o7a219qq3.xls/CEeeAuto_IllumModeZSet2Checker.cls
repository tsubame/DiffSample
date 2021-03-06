VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEeeAuto_IllumModeZSet2Checker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   コンディションチェッカ　光源退避の省略可否を判断する
'
'目的:
'   そのときの条件設定を光源から取得する(OptModOrModZ2用)
'作成者:
'   丸山 Create 2013/02/05
'   有川 編集 Ver1.1 2013/02/25
'   丸山/有川 Add 2013/02/08 2013/03/11追加
'
Option Explicit

Private Const THIS_CLASS_NAME As String = "CEeeAuto_IllumModeZSet2Checker"

Private m_IsEnable As Boolean 'FW_OptModOrModZ2を実行すべきかどうか

Private m_eEndDownPos As DownPosi '最終的な移動先の瞳距離
Private m_iEndFNumberTurret As Integer '最終的な移動先のターレット位置


Implements IEeeAuto_TestConditionChecker

'クラスの初期化
Private Sub Class_Initialize()
    m_IsEnable = True
End Sub

'OptModOrModZ2設定前の条件を取得する
'が本来の意味であるが、OptModOrModZ1を行うかどうかは、現在位置と最終的な
'行き先だけで決まってしまうため、この関数で判断まで行う。
'OptModOrModZ2が呼ばれるタイミングでは、「退避」-「退避に必要な設定」は完了しており
'単純に最終な行き先とあっているかどうかで判断してよい。
'最終的な行き先は[IEeeAuto_TestConditionChecker->SetEndPosition]で設定しておくこと
Private Sub IEeeAuto_TestConditionChecker_CheckBeforeCondition()

    '初期値は絶対やる
    m_IsEnable = True
    
    Dim eCurrentPos As DownPosi
    Dim iFNumberTurret As Integer
    
    If OptCond.IllumMaker = NIKON Then
    
        '現在のF値、瞳距離取得
        eCurrentPos = GetCurrentDownPosition
        iFNumberTurret = GetCurrentFNumberTurret
        
        Select Case OptCond.IllumModel
            Case "N-SIS3", "N-SIS5", "N-SIS5KAI"
                'F値、瞳距離ともに移動先と同じな場合はガチ無視で省略とする
                If m_eEndDownPos = eCurrentPos And m_iEndFNumberTurret = iFNumberTurret Then
                    m_IsEnable = False
                    Exit Sub
                End If
                
                'この時点でユニット方向は動作済み
                '高さが一緒の場合は上の「移動先と同じ場合」でFALSEになる。
                m_IsEnable = True
            
            Case "N-SIS3KAI"
                'F値、瞳距離ともに移動先と同じな場合はガチ無視で省略とする
                If m_eEndDownPos = eCurrentPos And m_iEndFNumberTurret = iFNumberTurret Then
                    m_IsEnable = False
                    Exit Sub
                End If
                
                'この時点で高さ方向は動作済み
                'F値が一緒の場合は上の「移動先と同じ場合」でFALSEになる。
                m_IsEnable = True
                
            Case Else
                'Nikon光源だけど "N-SIS3", "N-SIS5", "N-SIS5KAI","N-SIS3KAI"でない場合
                '未実装
                
        End Select
        
    Else
        'NIKON光源でない場合
        '未実装
    End If
    
    
End Sub

'CheckBeforeConditionの結果を返すのみ
Private Function IEeeAuto_TestConditionChecker_CheckAfterCondition() As Boolean
    
    'このクラスはBeforeだけで行うべきかどうかが確定する。
    '値を返すのみ
    IEeeAuto_TestConditionChecker_CheckAfterCondition = m_IsEnable

End Function

Private Sub IEeeAuto_TestConditionChecker_SetOperationMode(ByVal strMode As String)
    
    'モードはこのクラスにはない
    
End Sub

'光源設定の最終的な行き先をこのクラスに教える
'CheckBeforeConditionを呼ぶ前に教えてあげること
Private Sub IEeeAuto_TestConditionChecker_SetEndPosition(ByVal strPos As String)
    
    'PALSのGlobalObjectから行き先の、瞳距離、F値をとる
    With OptCond.CondInfo(strPos)
        m_eEndDownPos = .OptModDownPosition
        m_iEndFNumberTurret = .OptModFnumberTurret
    End With
    
End Sub

'瞳距離位置取得のラッパ関数
'のちのち装置からとるように変えるべき
Private Function GetCurrentDownPosition() As DownPosi

    If OptCond.IllumMaker = "NIKON" Then
        'Nikon光源の場合、装置によらずPALSからとる。
        'が若干脆弱。本当は装置からとるべき。
        GetCurrentDownPosition = PALS_IlluminatorMod.Flg_DownPosi
    End If
    
End Function

'ターレット位置取得のラッパ関数
'のちのち装置からとるように変えるべき
Private Function GetCurrentFNumberTurret() As Integer

    If OptCond.IllumMaker = "NIKON" Then
        'Nikon光源の場合、装置によらずPALSからとる。
        'が若干脆弱。本当は装置からとるべき。
        GetCurrentFNumberTurret = PALS_IlluminatorMod.Flg_FUnit
    End If
    
End Function
