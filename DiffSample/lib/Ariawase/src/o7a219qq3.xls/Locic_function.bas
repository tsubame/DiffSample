Attribute VB_Name = "Locic_function"
Option Explicit


Public PatCheckCounter_Logic As Double

Public Function FW_LogicJudgeTOPT(ByVal Parameter As CSetFunctionInfo) As Long
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Dim lngReturnVal(nSite) As Double
    
    Dim strTestLabel As String
        
    With Parameter
        strTestLabel = .Arg(0)
    End With
    
    
    If TheExec.RunOptions.AutoAcquire = True Then
        Dim iStatus As Long
        If TheHdw.Digital.Patgen.IsRunningAnySite = True Then      'True:Still Running
            iStatus = 0
        ElseIf TheHdw.Digital.Patgen.IsRunningAnySite = False Then 'False:haltend or keepalive
            iStatus = 1
        End If
        
        While (iStatus <> 1)
            If PatCheckCounter_Logic < 999 Then
                TheHdw.TOPT.Recall
                PatCheckCounter_Logic = PatCheckCounter_Logic + 1
                Call WaitSet(10 * mS)
                Exit Function
            Else
                Call StopPattern
                iStatus = 1
            End If
        Wend
        
    Else
        Call StopPattern
    End If


    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            lngReturnVal(site) = TheHdw.Digital.FailedPinsCount(site)
            'Debug.Print "Failed pin count : " & lngReturnVal(site)
            If lngReturnVal(site) = 0 Then
                retResult(site) = 1
            ElseIf lngReturnVal(site) >= 1 Then
                retResult(site) = 0
            Else
                MsgBox ("FW_LogicJudgeTOPT Error")
            End If
         End If
    Next site

    'Call test(retResult)

    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(strTestLabel), retResult)
    
    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function

'内容:
'   パタンの開始をおこなう(終了をまたない)
'
'パラメータ:
'    [Arg0]      In   パタン名
'    [Arg1]      In   TSBシート名
'    [Arg2]      In   実行後のウェイトタイム(省略可能 省略時Waitなし)
'
'戻り値:
'
'注意事項:IP750 or Decoder Patは、専用で設定する。
'
Public Sub FW_PatSetTOPT(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
'    Const PAT_START_LABEL As String = "pat_start"
    Const PAT_START_LABEL As String = "START"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
    
    Call StopPattern_Halt 'EeeJob関数
    Call SetTimeOut 'EeeJob関数
    
    With TheHdw.Digital
        Call .Timing.Load(strTsbName)
        Call .Patterns.pat(strPatGroupName).Start(PAT_START_LABEL)
    End With
    
   '待ち時間の指定がある場合、待つ
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        If TheExec.RunOptions.AutoAcquire = True Then
            Call TheHdw.TOPT.WAIT(toptTimer, dblWaitTime * 1000)
        Else
            Call TheHdw.WAIT(dblWaitTime)
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub
