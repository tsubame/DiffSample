Attribute VB_Name = "XEeeAuto_PowerCondition"
'概要:
'   PowerConditionの概念を提供する
'
'目的:
'   PowerSequenceシートと、PowerConditionシートを読み込み
'   PowerSupplyVoltageの電圧をSequenceにしたがって印加する。
'
'作成者:
'   2011/12/05 Ver0.1 D.Maruyama
'   2011/12/07 Ver0.2 D.Maruyama
'       ・PowerCondition名からPowerSequence名、PowerSupplyVoltage名を返せるようにした
'       ・コメントを一部修正
'   2012/04/06 Ver0.3 D.Maruyma  ApplyPowerConditionを条件名とシーケンス名の2つを取るように変更
'                                 それにあわせて、不要な関数、変数を削除した

Option Explicit

'固定値
Private Const POWER_SEQUENCE_SHEET_COND_INDEX_CELL = "B4"
Private Const POWER_SEQUENCE_SHEET_NAME = "PowerSequence"

'モジュール内変数
Private m_colPowerSequence As Collection


'内容:
'   モジュールを初期化
'
'備考:
'   モジュール内変数をいったん空にして、情報をシートから読み直す。
'
Public Sub InitializePowerCondition()

    'いったん空にする
    Set m_colPowerSequence = Nothing

    'コレクションの生成
    Set m_colPowerSequence = New Collection
    
    'シートの読み込み
    Call ReadPowerSequenceSheet
    
End Sub

Public Sub UninitializePowerCondition()

    'いったん空にする
    Set m_colPowerSequence = Nothing

End Sub

'内容:
'   PowerConditionを設定
'
'パラメータ:
'[strPowerConditionName]    IN   String:    設定するPowerCondition名
'
'備考:
'   指定したPowerConditionを実行する
'
Public Sub ApplyPowerCondition(ByVal strPowerConditionName As String, ByVal strSequenceName As String)
   
    Dim pPowerSequence As CPowerSequence
    
On Error GoTo SEQUENCE_NOT_FOUND
    Set pPowerSequence = m_colPowerSequence.Item(strSequenceName)
On Error GoTo 0

On Error GoTo ErrHandler
    Call pPowerSequence.Execute(strPowerConditionName)
On Error GoTo 0
    
    Exit Sub
        
ErrHandler:
    Call MsgBox("ApplyPowerCondition Fucntion Error Detect! @ " & strPowerConditionName & " @ " & strSequenceName)
    Exit Sub

SEQUENCE_NOT_FOUND:
    Call MsgBox("ApplyPowerCondition Fucntion Sequence not found! @ " & strPowerConditionName & " @ " & strSequenceName)
    Exit Sub

End Sub

'内容:
'   PowerConditionを設定。For APMU UnderShoot
'
'パラメータ:
'[strPowerConditionName]    IN   String:    設定するPowerCondition名
'
'備考:
'   指定したPowerConditionを実行する
'
Public Sub ApplyPowerConditionForUS(ByVal strPowerConditionName As String, ByVal strSequenceName As String)
   
    Dim pPowerSequence As CPowerSequence
    
On Error GoTo SEQUENCE_NOT_FOUND
    Set pPowerSequence = m_colPowerSequence.Item(strSequenceName)
On Error GoTo 0

On Error GoTo ErrHandler
    Call pPowerSequence.ExcecuteForUS(strPowerConditionName)
On Error GoTo 0
    
    Exit Sub
        
ErrHandler:
    Call MsgBox("ApplyPowerCondition Fucntion Error Detect! @ " & strPowerConditionName & " @ " & strSequenceName)
    Exit Sub

SEQUENCE_NOT_FOUND:
    Call MsgBox("ApplyPowerCondition Fucntion Sequence not found! @ " & strPowerConditionName & " @ " & strSequenceName)
    Exit Sub

End Sub


'内容:
'   PowerSequenceSheetから読み込みを行う
'
'備考:
'
'
Private Sub ReadPowerSequenceSheet()

    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(POWER_SEQUENCE_SHEET_NAME)
    
    Dim NumOfItem As Long
    NumOfItem = sht.Range(POWER_SEQUENCE_SHEET_COND_INDEX_CELL).End(xlToRight).Column _
        - sht.Range(POWER_SEQUENCE_SHEET_COND_INDEX_CELL).Column - 1
    
    Dim i As Long
    Dim j As Long

    With sht.Range(POWER_SEQUENCE_SHEET_COND_INDEX_CELL)
        For i = 0 To NumOfItem
            Dim tempPowerSeq As CPowerSequence
            Set tempPowerSeq = New CPowerSequence
            j = 0
            Call tempPowerSeq.InitializeThisClass(.offset(j, i + 1))
            While .offset(j + 1, i + 1) <> ""
                Dim tempPowerSequenceItem As IPowerSequenceItem
                If IsNumeric(.offset(j + 1, i + 1)) Then
                    Set tempPowerSequenceItem = New CPowerSquenceWait
                    Call tempPowerSequenceItem.SetParam(.offset(j + 1, i + 1))
                Else
                    Set tempPowerSequenceItem = New CPowerSequencePin
                    Call tempPowerSequenceItem.SetParam(.offset(j + 1, i + 1).Text)
                End If
                Call tempPowerSeq.Add(tempPowerSequenceItem)
                j = j + 1
            Wend
            Call m_colPowerSequence.Add(tempPowerSeq, tempPowerSeq.Name)
        Next i
    End With

End Sub

'内容:SetVoltage(US対策版)
Public Sub PowerDown4ApmuUnderShoot() '2012/11/16 175Debug Arikawa

        'パタン停止
    Call StopPattern 'EeeJob関数
   
    Dim pPowerSequence As CPowerSequence
    If getPowerDownSequence(pPowerSequence) = False Then Exit Sub
    
    On Error GoTo ERROR_DETECTION1
    Call pPowerSequence.ExcecuteForUS("ZERO")
    Exit Sub
ERROR_DETECTION1:
    Call pPowerSequence.ExcecuteForUS("ZERO_V")
End Sub

'内容
'   電源Off時のPower Sequence名を取得する。
'   Gangの有無によって、予め"PowerSequence"シートに生成される、レジスタ通信I/F
'   に依存しない電源Offシーケンス名が異なるため、
'       1. APMU Gangの可能性のある端子が含まれる場合
'       2. APMU Gangの可能性のある端子が含まれない場合
'   の順に、シーケンス情報を取得する。
'Description
'   To return power sequence object for power down.
'   In order to support both pin assigns with APMU Gang and without APMU Gang,
'   it is done in the following order.
'       1. Get power sequence with name "ANY_SeqOff_GangOff"
Private Function getPowerDownSequence(ByRef pPowerSequence As CPowerSequence) As Boolean
    Const POWER_DOWN_SEQUENCE_GANG As String = "ANY_SeqOff_GangOff"
    Const POWER_DOWN_SEQUENCE_NOGANG As String = "ANY_SeqOff"

    '1. APMU Gang OFF時のシーケンス。Gangの主ピンをOFFすれば、全部OFFになるはず。
    On Error GoTo ErrorGangNotFound
    Set pPowerSequence = m_colPowerSequence.Item(POWER_DOWN_SEQUENCE_GANG)
    getPowerDownSequence = True
    Exit Function
    
ErrorGangNotFound:
    '2. APMU Gangの可能性のある端子が含まれない場合のOFFシーケンス。
    On Error GoTo ErrorSeqNotFound
    Set pPowerSequence = m_colPowerSequence.Item(POWER_DOWN_SEQUENCE_NOGANG)
    getPowerDownSequence = True
    Exit Function

ErrorSeqNotFound:
    Err.Raise 9999, "getPowerDownSequence", "Power down sequence not found [" & GetInstanceName & "] !"
    Call DisableAllTest
End Function

'内容:SetVoltage(US対策版)
Public Sub Set_Voltage(ByVal strPowerConditionName As String, ByVal strSequenceName As String) '2012/11/16 175Debug Arikawa

        'パタン停止
    Call StopPattern 'EeeJob関数
   
    Dim pPowerSequence As CPowerSequence
    Set pPowerSequence = m_colPowerSequence.Item(strSequenceName)

    Call pPowerSequence.ExcecuteForUS(strPowerConditionName)

End Sub


