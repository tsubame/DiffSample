Attribute VB_Name = "XEeeAuto_DC_Std_Iana"

Option Explicit

'概要:
'
'
'目的:
'   Binning用モジュール
'
'作成者:
'   2013/02/14 Ver1.0 K.Hamada
'   2013/02/22 Ver1.1 K.Hamada

Public Binning_Judge_Flg(nSite) As Double


'内容:
'   TestInstanceに書かれたキーからLimitで判定して、Binning_Judge_Flgを処理する
'
'注意事項:
'

Public Function ReturnResultBinningPreJudge_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long

    Dim retResult(nSite) As Double
    Erase retResult
    
    Dim ArgArr() As String
    Dim dblCalcValid As Double
    If Not EeeAutoGetArgument(ArgArr, 1) Then
        Err.Raise 9999, "ReturnResultBinningPreJudge_f", "Argument type is Mismatch """ & ArgArr(0) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalcValid = CDbl(ArgArr(0))
    
    Dim TempValue() As Double
    Call mf_GetResult(UCase(GetInstanceName), TempValue)
  
    'Limit_Get
    Dim LoLimit As Double
    Dim HiLimit As Double
    Call m_GetLimit(LoLimit, HiLimit)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
                retResult(site) = TempValue(site)
                
                Select Case dblCalcValid
                    Case 0
                            Binning_Judge_Flg(site) = 0
                    Case 1
                            If retResult(site) < LoLimit Then
                                Binning_Judge_Flg(site) = Binning_Judge_Flg(site) + 1
                            End If
                    Case 2
                            If retResult(site) > HiLimit Then
                                Binning_Judge_Flg(site) = Binning_Judge_Flg(site) + 1
                            End If
                    Case 3
                            If retResult(site) < LoLimit Or retResult(site) > HiLimit Then
                                Binning_Judge_Flg(site) = Binning_Judge_Flg(site) + 1
                            End If
                    Case Else
                End Select
                
        End If
    Next site

    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)
    
    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

Public Function ReturnResultBinningSumPreJudge_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long

    Dim retResult(nSite) As Double
    Erase retResult
    
    Dim TempValue() As Double
    
    '本関数に対するパラメータ数は不定。; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM

    Dim i As Long
    'パラメータの取得; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ReturnResultBinningSumPreJudge_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 + 1 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ReturnResultBinningSumPreJudge_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    Dim dblCalcValid As Double
    dblCalcValid = CDbl(ArgArr(0))
    
    '足し合わせ
    Dim tmpValue() As Double
    For i = 0 + 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    
    'Limit_Get
    Dim LoLimit As Double
    Dim HiLimit As Double
    Call m_GetLimit(LoLimit, HiLimit)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
        
            Select Case dblCalcValid
                Case 0
                        Binning_Judge_Flg(site) = 0
                Case 1
                        If retResult(site) < LoLimit Then
                            Binning_Judge_Flg(site) = Binning_Judge_Flg(site) + 1
                        End If
                Case 2
                        If retResult(site) > HiLimit Then
                            Binning_Judge_Flg(site) = Binning_Judge_Flg(site) + 1
                        End If
                Case 3
                        If retResult(site) < LoLimit Or retResult(site) > HiLimit Then
                            Binning_Judge_Flg(site) = Binning_Judge_Flg(site) + 1
                        End If
                Case Else
            End Select
            
        End If
    Next site

    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)
    
    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function


Public Function ReturnResultBinningPostJudge_f() As Double

    On Error GoTo ErrorExit

    Call SiteCheck

    Dim site As Long

    Dim retResult(nSite) As Double
    Erase retResult

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = Binning_Judge_Flg(site)
        End If
    Next site

    For site = 0 To nSite
        Binning_Judge_Flg(site) = 0
    Next site

    'ジャッジ
    Call test(retResult)

    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function

ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

