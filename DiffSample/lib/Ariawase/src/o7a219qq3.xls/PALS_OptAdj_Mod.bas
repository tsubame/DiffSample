Attribute VB_Name = "PALS_OptAdj_Mod"
Option Explicit

Public Const OPTTOOLNAME As String = "PALS - Auto Opt Adjust"
Public Const OPTTOOLVER As String = "1.40"

Public Const intOptTryNum As Integer = 10
Public Const intOptAveNum As Integer = 1

Public Const g_blnOptDebOffline As Boolean = False

Public Const g_strOptDataTextDeb As String = "C:\Documents and Settings\0020205267\デスクトップ\OptAdjData_p7n678akb_opt_#71_20101018_180048.txt"
'Public Const g_strOptDataTextDeb As String = "\\43.24.100.12\simulator\2Section\imamura\ILX155K_tenken_100312.txt"

Public g_blnOptStop As Boolean
Public dblDataPrev() As Double
Public intWedgePrev() As Integer
Public g_dblMaxLux As Double
Public Const WEDGEPERMITTEDPER As Double = 0
Public Const LUXPERMITTEDPER As Double = 0.1
Public g_blnOptCondAdjusted(500) As Boolean

Public Sub sub_OptFrmShow()
    frm_PALS_OptAdj_Main.Show
End Sub

'********************************************************************************************
' 名前 : sub_OutPutOptParam
' 内容 : Opt(NSIS)シートのパラメータを、測定データログの末尾にテキストで追加
'        下記のようなデータが追加される
'        ########### Parameter ###########
'        Identifier  LUX    WEGDE
'        LL          3.76      -1
'        HL            -1     400
'        #################################
' 引数 : なし
' 備考： なし
' 更新履歴： Rev1.0      2010/10/18　新規作成   M.Imamura
'********************************************************************************************
Public Sub sub_OutPutOptParam()

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_OutPutOptParam

    Dim intFileNo As Integer                'ファイル番号
    Dim intCategoryNum As Long              'カテゴリ名を回すループカウンタ
    
    intFileNo = FreeFile                    'ファイル番号の取得
        
    'Opt(NSIS)シートのパラメータを、データログに追記
    'Append(追記)モードで測定データログを開き、各パラメータを追記

    If g_blnOptDebOffline = False Then Open g_strOutputDataText For Append As #intFileNo
    If g_blnOptDebOffline = True Then Open g_strOptDataTextDeb For Append As #intFileNo


    Print #intFileNo, ""
    Print #intFileNo, "MEASURE DATE : " & Year(Date) & "/" & Month(Date) & "/" & Day(Date)
    Print #intFileNo, "JOB NAME     : " & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 4)
    Print #intFileNo, "SW_NODE      : " & Sw_Node

        
    Print #intFileNo, "########### Parameter ###########"
    Print #intFileNo, "Identifier" & Space(15 - Len("Identifier")) & "LUX" & Space(10 - Len("LUX")) & "Wedge" & Space(10 - Len("Wedge")) & "ND"

    'カテゴリ数繰り返す
    For intCategoryNum = 0 To OptCond.OptCondNum
        With OptCond.CondInfoI(intCategoryNum)
            Print #intFileNo, .OptIdentifier & Space(15 - Len(.OptIdentifier)) & CStr(.AxisLevel) & Space(10 - Len(CStr(.AxisLevel))) & .WedgeFilter & Space(10 - Len(CStr(.WedgeFilter))) & CStr(.NDFilter)
        End With
    Next
        
    Print #intFileNo, "#################################"
    
    'データログを閉じる
    Close #intFileNo

Exit Sub

errPALSsub_OutPutOptParam:
    Call sub_errPALS("OutPut OptParameter error at 'sub_OutPutOptParam'", "4-2-01-6-01")

End Sub

'********************************************************************************************
' 名前: sub_CheckOptTarget
' 内容: OptResultSheetNameへのデータ排出を行い、
'       OptTarget±OptJudgeLimitに入っていない場合、光量更新の計算を行なう
' 引数: blnOptUpdate   : 光量調整Go/No
'       lngNowLoopCnt  : 測定済み回数
' 戻値: True  : 更新無
'       False : 更新有
' 備考： なし
' 更新履歴： Rev1.0      2010/10/21　新規作成   M.Imamura
'********************************************************************************************
'######### Write Log & Check Data & UpdateOpt
Public Function sub_CheckOptTarget(ByVal lngNowLoopCnt As Long, blnOptUpdate As Boolean, Optional strOptIdenShTgt As String = "") As Boolean
    Dim myrange As Range
    Dim myrow As Integer
    Dim dblNextLux As Double
    '>>>2011/4/20 M.IMAMURA UPDATE Integer ->Long
    Dim lngNextWedge As Long
    '<<<2011/4/20 M.IMAMURA UPDATE
    '>>>2011/6/06 M.IMAMURA Add.
    Dim blnNDUp As Boolean
    Dim blnNDDwn As Boolean
    '>>>2011/6/06 M.IMAMURA Add.

    On Error GoTo errPALSsub_CheckOptTarget
    
    For Each myrange In Worksheets(OptResultSheetName).Range("B5:B65535")
        If myrange.Value = vbNullString And Worksheets(OptResultSheetName).Cells(myrange.Row + 1, 2).Value = vbNullString Then
            myrow = myrange.Row + 1
        Exit For
        End If
    Next myrange

    sub_CheckOptTarget = True
    Call sub_TestingStatusOutPals(frm_PALS_OptAdj_Main, "Now Checking...")
    
    Dim intParamsLoop As Integer
    Dim lngAveLoop As Long
    '>>>2011/9/5 M.IMAMURA Add.
    Dim dblDataAve As Double
    Dim dblDataMax As Double
    Dim dblDataMin As Double
    Dim dblDataSigma As Double
    Dim dblDataAve1time As Double
    Dim lngMySite As Double
    '>>>2011/9/5 M.IMAMURA Add.
    
    For intParamsLoop = 0 To PALS.CommonInfo.TestCount
        '>>>2011/6/06 M.IMAMURA Add.
        blnNDUp = False
        blnNDDwn = False
        '<<<2011/6/06 M.IMAMURA Add.
        If PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier <> vbNullString And (strOptIdenShTgt = "" Or strOptIdenShTgt = PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier) Then
            g_blnOptCondAdjusted(OptCond.CondInfoNo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier)) = True
            
            '>>>2011/9/5 M.IMAMURA Add.
            If frm_PALS_OptAdj_Main.ob_AveAllSite.Value = True Then
                dblDataAve = 0
                dblDataMax = -9999
                dblDataMin = 9999
                For lngMySite = 0 To nSite
                    '>>>2011/10/3 M.IMAMURA Add.
                    If dblDataMax < PALS.CommonInfo.TestInfo(intParamsLoop).site(lngMySite).max Then
                        dblDataMax = PALS.CommonInfo.TestInfo(intParamsLoop).site(lngMySite).max
                    End If
                    If dblDataMin > PALS.CommonInfo.TestInfo(intParamsLoop).site(lngMySite).Min Then
                        dblDataMin = PALS.CommonInfo.TestInfo(intParamsLoop).site(lngMySite).Min
                    End If
                    '<<<2011/10/3 M.IMAMURA Add.
                    dblDataAve = dblDataAve + PALS.CommonInfo.TestInfo(intParamsLoop).site(lngMySite).ave
                Next lngMySite
                dblDataAve = dblDataAve / (nSite + 1)
                dblDataSigma = 0
            Else
                dblDataAve = PALS.CommonInfo.TestInfo(intParamsLoop).site(0).ave
                dblDataMax = PALS.CommonInfo.TestInfo(intParamsLoop).site(0).max
                dblDataMin = PALS.CommonInfo.TestInfo(intParamsLoop).site(0).Min
                dblDataSigma = PALS.CommonInfo.TestInfo(intParamsLoop).site(0).Sigma
            End If
            '<<<2011/9/5 M.IMAMURA Add.
            
            '######### Write Log
            With Worksheets(OptResultSheetName)
            .Cells(myrow, 2).Value = Sw_Node
            .Cells(myrow, 3).Value = lngNowLoopCnt
            .Cells(myrow, 4).Value = Format(Date, "yyyy/mm/dd") & " " & Format(TIME, "hh:mm:ss")
            .Cells(myrow, 5).Value = PALS.CommonInfo.TestInfo(intParamsLoop).tname
            .Cells(myrow, 6).Value = PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier
            If OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).AxisLevel > 0 Then
                .Cells(myrow, 7).Value = OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).AxisLevel  'Lux
            Else
                .Cells(myrow, 8).Value = OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter 'WegdeFilter
            End If
            .Cells(myrow, 9).Value = PALS.CommonInfo.TestInfo(intParamsLoop).OptTarget
            .Cells(myrow, 10).Value = PALS.CommonInfo.TestInfo(intParamsLoop).OptJudgeLimit * 100

            .Cells(myrow, 11).Value = dblDataAve
            .Cells(myrow, 13).Value = dblDataMax
            .Cells(myrow, 14).Value = dblDataMin
            .Cells(myrow, 15).Value = dblDataSigma
            For lngAveLoop = 1 To val(frm_PALS_OptAdj_Main.cbo_AveNum.Text)
                '>>>2011/9/5 M.IMAMURA Mod.
                If frm_PALS_OptAdj_Main.ob_AveAllSite.Value = True Then
                    dblDataAve1time = 0
                    For lngMySite = 0 To nSite
                        dblDataAve1time = dblDataAve1time + PALS.CommonInfo.TestInfo(intParamsLoop).site(lngAveLoop).Data(lngAveLoop)
                    Next lngMySite
                    dblDataAve1time = dblDataAve1time / (nSite + 1)
                Else
                    dblDataAve1time = PALS.CommonInfo.TestInfo(intParamsLoop).site(0).Data(lngAveLoop)
                End If
                .Cells(myrow, 15 + lngAveLoop).Value = dblDataAve1time
                '<<<2011/9/5 M.IMAMURA Mod.
            Next
            
            '######### Check Data
            '######### Check NG
            If Abs(dblDataAve - PALS.CommonInfo.TestInfo(intParamsLoop).OptTarget) > PALS.CommonInfo.TestInfo(intParamsLoop).OptTarget * PALS.CommonInfo.TestInfo(intParamsLoop).OptJudgeLimit Then
                sub_CheckOptTarget = False
                .Cells(myrow, 12).Value = "NG"
                .Cells(myrow, 12).Interior.color = vbYellow
                '######### UpdateOpt
                If blnOptUpdate = True Then
                    '######### Update LUX
                    If OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).AxisLevel > 0 Then
                        '>>>2011/4/20 M.IMAMURA ADD
                        If dblDataAve = 0 Then
                            Call sub_errPALS("OptAdjust Error Occured(" & PALS.CommonInfo.TestInfo(intParamsLoop).tname & ".Site(0).Ave = 0) at 'sub_CheckOptTarget'", "4-2-02-5-02")
                            Exit Function
                        End If
                        '<<<2011/4/20 M.IMAMURA ADD
                        dblNextLux = OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).AxisLevel * PALS.CommonInfo.TestInfo(intParamsLoop).OptTarget / dblDataAve
                        'LuxValue Check < 0.01
                        '>>> 2011/8/1 M.Imamura Changed LuxValue Check < 0.1
                        If dblNextLux < 0.1 Then
                            Call sub_errPALS("OptAdjust Error Occured(" & PALS.CommonInfo.TestInfo(intParamsLoop).tname & ".dblNextLux<0.01) at 'sub_CheckOptTarget'", "4-2-01-3-03")
                            Exit Function
                        End If
                        
                        'LuxValue Check > Max
                        If OptCond.IllumMaker = NIKON And OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).AxisLevel = g_dblMaxLux * (1 - LUXPERMITTEDPER) Then
                            Call sub_errPALS("OptAdjust Error Occured(" & PALS.CommonInfo.TestInfo(intParamsLoop).tname & ".AxisLevel > MaxLux * (1 - LUXPERMITTEDPER)) at 'sub_CheckOptTarget'", "4-2-01-3-04")
                            Exit Function
                        ElseIf OptCond.IllumMaker = NIKON And dblNextLux > g_dblMaxLux * (1 - LUXPERMITTEDPER) Then
                            dblNextLux = g_dblMaxLux * (1 - LUXPERMITTEDPER)
                        End If
                        
                        
                        If sub_UpdateOpt(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier, dblNextLux, "Lux") = False Then
                            Exit Function
                        End If
                    '######### Update WEDGE
                    Else
NDupdate:
                        If lngNowLoopCnt = 1 Then
                            'if target>now
                            If PALS.CommonInfo.TestInfo(intParamsLoop).OptTarget > dblDataAve Then
                                lngNextWedge = OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter - 100
                             'if target<now
                            Else
                                lngNextWedge = OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter + 100
                            End If
                        Else
                            'Newton Method
                            If g_blnOptDebOffline = True Then dblDataPrev(intParamsLoop) = dblDataPrev(intParamsLoop) * 0.99
                            If dblDataAve - dblDataPrev(intParamsLoop) = 0 Then
                                Call sub_errPALS("OptAdjust Error Occured(" & PALS.CommonInfo.TestInfo(intParamsLoop).tname & ".Ave=dblDataPrev) at 'sub_CheckOptTarget'", "4-2-01-5-05")
                                Exit Function
                            Else
                                '>>>2011/6/06 M.IMAMURA Mod.
                                If blnNDUp = True Then
                                    lngNextWedge = Int(0.5 + OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter _
                                            + (OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter - intWedgePrev(intParamsLoop)) _
                                            / (dblDataAve / 4 - dblDataPrev(intParamsLoop) / 3) _
                                            * (PALS.CommonInfo.TestInfo(intParamsLoop).OptTarget - dblDataAve / 4))
                                ElseIf blnNDDwn = True Then
                                    lngNextWedge = Int(0.5 + OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter _
                                            + (OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter - intWedgePrev(intParamsLoop)) _
                                            / (dblDataAve * 3 - dblDataPrev(intParamsLoop) * 3) _
                                            * (PALS.CommonInfo.TestInfo(intParamsLoop).OptTarget - dblDataAve * 3))
                                Else
                                    lngNextWedge = Int(0.5 + OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter _
                                            + (OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter - intWedgePrev(intParamsLoop)) _
                                            / (dblDataAve - dblDataPrev(intParamsLoop)) _
                                            * (PALS.CommonInfo.TestInfo(intParamsLoop).OptTarget - dblDataAve))
                                End If
                                '<<<2011/6/06 M.IMAMURA Mod.
                            End If
                        End If
                        
                        'WedgeValue Check < Min
                        If blnNDUp = False And blnNDDwn = False And OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter = OptCond.IllumWedgeMin + OptCond.IllumWedgeMax * WEDGEPERMITTEDPER Then
                        '>>>2011/6/6 M.IMAMURA Mod.
                            If OptCond.IllumMaker = NIKON Or OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).NDFilter <= OptCond.IllumNdMin Then
                                Call sub_errPALS("OptAdjust Error Occured(" & PALS.CommonInfo.TestInfo(intParamsLoop).tname & ".lngNextWedge < OptCond.IllumWedgeMin) at 'sub_CheckOptTarget'", "4-2-01-3-06")
                                Exit Function
                            Else
                                OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).NDFilter = OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).NDFilter - 1
                                blnNDUp = False
                                blnNDDwn = True
                                GoTo NDupdate
                            End If
                        '<<<2011/6/6 M.IMAMURA Mod.
                        ElseIf lngNextWedge < OptCond.IllumWedgeMin + OptCond.IllumWedgeMax * WEDGEPERMITTEDPER Then
                            lngNextWedge = OptCond.IllumWedgeMin + OptCond.IllumWedgeMax * WEDGEPERMITTEDPER
                        End If

                        'WedgeValue Check > Max
                        If blnNDUp = False And blnNDDwn = False And OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter = OptCond.IllumWedgeMax * (1 - WEDGEPERMITTEDPER) Then
                        '>>>2011/6/6 M.IMAMURA Mod.
                            If OptCond.IllumMaker = NIKON Or OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).NDFilter >= OptCond.IllumNdMax Then
                                Call sub_errPALS("OptAdjust Error Occured(" & PALS.CommonInfo.TestInfo(intParamsLoop).tname & ".lngNextWedge > OptCond.IllumWedgeMax) at 'sub_CheckOptTarget'", "4-2-01-3-07")
                                Exit Function
                            Else
                                OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).NDFilter = OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).NDFilter + 1
                                blnNDUp = True
                                blnNDDwn = False
                                GoTo NDupdate
                            End If
                        '<<<2011/6/6 M.IMAMURA Mod.
                        ElseIf lngNextWedge > OptCond.IllumWedgeMax * (1 - WEDGEPERMITTEDPER) Then
                             lngNextWedge = OptCond.IllumWedgeMax * (1 - WEDGEPERMITTEDPER)
                        End If
                        
                        
                        If sub_UpdateOpt(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier, lngNextWedge, "Wedge") = False Then
                            Exit Function
                        End If
                        '>>>2011/6/6 M.IMAMURA Add.
                        If OptCond.IllumMaker = KESILLUM Or OptCond.IllumMaker = INTERACTION Then
                            If sub_UpdateOpt(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier, OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).NDFilter, "ND") = False Then
                                Exit Function
                            End If
                        End If
                        '<<<2011/6/6 M.IMAMURA Add.
                    End If
                End If
            '######### Check OK
            Else
                .Cells(myrow, 12).Value = "OK"
                .Cells(myrow, 12).Interior.color = vbCyan
                If OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).AxisLevel > 0 Then
                    If sub_UpdateOpt(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier, OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).AxisLevel, "Lux", True) = False Then
                        Exit Function
                    End If
                Else
                    If sub_UpdateOpt(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier, OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter, "Wedge", True) = False Then
                        Exit Function
                    End If
                End If
            End If
                        
            '>>>2011/6/15 M.IMAMURA Mod.
            If blnNDUp = True Then
                dblDataPrev(intParamsLoop) = dblDataAve / 4
            ElseIf blnNDDwn = True Then
                dblDataPrev(intParamsLoop) = dblDataAve * 3
            Else
                dblDataPrev(intParamsLoop) = dblDataAve
            End If
            '<<<2011/6/15 M.IMAMURA Mod.
            intWedgePrev(intParamsLoop) = OptCond.CondInfo(PALS.CommonInfo.TestInfo(intParamsLoop).OptIdentifier).WedgeFilter

            '######### GoNext Line
            myrow = myrow + 1
            End With
        End If
    Next

    Exit Function

errPALSsub_CheckOptTarget:
    Call sub_errPALS("OptAdjust Tool Run error at 'sub_CheckOptTarget'", "4-2-01-0-08")

End Function

'********************************************************************************************
' 名前: sub_UpdateOpt
' 内容: strSearchSheetへのLUX/WEDGEの更新を行なう
' 引数: strIllumMode  : 光量識別子
'       OptTarget  　 : 更新値
'       TargetMode  　: Lux or Wedge
' 戻値: True  : 更新OK
'       False : 更新エラー
' 備考： なし
' 更新履歴： Rev1.0      2010/10/21　新規作成   M.Imamura
'********************************************************************************************
Public Function sub_UpdateOpt(strIllumMode As String, ByRef OptTarget, TargetMode As String, Optional blnIsFinished As Boolean = False, Optional blnIsInit As Boolean = False) As Boolean
    Dim myrowloop As Long
    
    Dim strSearchSheet As String
    Dim strSearchStr As String
    
    Dim nodePoint As Variant
    Dim IdenPoint As Variant
    Dim axisPoint As Variant

    sub_UpdateOpt = False
    
    On Error GoTo errPALSsub_UpdateOpt
    
    If OptCond.IllumMaker = NIKON Then
        strSearchSheet = NIKON_WRKSHT_NAME
        If TargetMode = "Lux" Then
            strSearchStr = "Axis.Level"
        ElseIf TargetMode = "Wedge" Then
            strSearchStr = "WedgeFilter"
        '>>>2011/6/6 M.IMAMURA Mod.
        ElseIf TargetMode = "ND" Then
            strSearchStr = "NDFilter"
        '<<<2011/6/6 M.IMAMURA Mod.
        Else
            Call sub_errPALS("OptAdjust Error Occured(Unknown TargetMode) at 'sub_UpdateOpt'", "4-2-02-4-09")
            Exit Function
        End If
    ElseIf OptCond.IllumMaker = INTERACTION Or OptCond.IllumMaker = KESILLUM Then
        strSearchSheet = IA_WRKSHT_NAME
        If TargetMode = "Lux" Then
            strSearchStr = "L"
        ElseIf TargetMode = "Wedge" Then
            strSearchStr = "A"
        '>>>2011/6/6 M.IMAMURA Mod.
        ElseIf TargetMode = "ND" Then
            strSearchStr = "N"
        '<<<2011/6/6 M.IMAMURA Mod.
        Else
            Call sub_errPALS("OptAdjust Error Occured(Unknown TargetMode) at 'sub_UpdateOpt'", "4-2-02-4-10")
            Exit Function
        End If
    End If

    '======= Base Point Find ========
    Set nodePoint = Worksheets(strSearchSheet).Range("A1:R10").Find("Sw_Node")
    If nodePoint Is Nothing Then
        GoTo errPALSsub_UpdateOpt
    End If
    Set IdenPoint = Worksheets(strSearchSheet).Range("A1:R100").Find("Identifier")
    If IdenPoint Is Nothing Then
        GoTo errPALSsub_UpdateOpt
    End If
    Set axisPoint = Worksheets(strSearchSheet).Range("A1:R100").Find(strSearchStr)
    If axisPoint Is Nothing Then
        GoTo errPALSsub_UpdateOpt
    End If

    For myrowloop = nodePoint.Row + 2 To 65535
        If Worksheets(strSearchSheet).Cells(myrowloop, nodePoint.Column).Value = Sw_Node And Worksheets(strSearchSheet).Cells(myrowloop, IdenPoint.Column).Value = strIllumMode Then
            Worksheets(strSearchSheet).Cells(myrowloop, axisPoint.Column).Value = Format(OptTarget, "0.00")
            If blnIsFinished = True Then
                Worksheets(strSearchSheet).Cells(myrowloop, axisPoint.Column).Interior.color = vbCyan
            Else
                Worksheets(strSearchSheet).Cells(myrowloop, axisPoint.Column).Interior.color = vbYellow
            End If
            If blnIsInit = True Then
                Worksheets(strSearchSheet).Cells(myrowloop, axisPoint.Column).Interior.color = vbWhite
            End If
            
            sub_UpdateOpt = True
            Exit Function
        End If
    Next myrowloop

errPALSsub_UpdateOpt:
    Call sub_errPALS("OptAdjust Tool Run error at 'sub_UpdateOpt'", "4-2-02-0-11")
        
End Function


Function sub_XOptCalculate(strTargetIden As String) As Boolean

Dim intOptCalcLoop As Integer               '光源シート検索用Loop
Dim intOptModePosiRef As Integer            '基本Modeの位置特定用
Dim intOptModePosi As Integer               'Modeの位置特定用
Dim strOptMode As String                    '基本Modeの名前
Dim intOptTimesPosi As Integer              'Modeの倍数位置用
Dim dblOptTimes As Double                   'Modeの倍数
Dim dblOptCalcLux As Double                 '計算した光量

      
    'PALS項目名から"_"の位置を検索してModeの最終位置を格納
    If InStr(strTargetIden, "_") > 0 Then
        intOptModePosiRef = InStr(strTargetIden, "_") - 1
    Else
        intOptModePosiRef = Len(strTargetIden)
    End If
    
    'PALS項目名から基本Modeのみを抽出
    strOptMode = Left(strTargetIden, intOptModePosiRef)
    
    '光源シートに基本Modeから計算する倍数項目があるかどうかの検索のためのループ
    For intOptCalcLoop = 0 To OptCond.OptCondNum
    
        'Modeの中に"基本Mode&X"が入っているか判定
        If strOptMode & "X" = Left(OptCond.CondInfoI(intOptCalcLoop).OptIdentifier, Len(strOptMode & "X")) Then
            '倍数抽出のために"X"の位置を検索して格納
            intOptTimesPosi = InStr(2, OptCond.CondInfoI(intOptCalcLoop).OptIdentifier, "X") + 1
            '倍数抽出のために"_"の位置を検索して格納
            If InStr(OptCond.CondInfoI(intOptCalcLoop).OptIdentifier, "_") > 0 Then
                intOptModePosi = InStr(OptCond.CondInfoI(intOptCalcLoop).OptIdentifier, "_") - 1
            Else
                intOptModePosi = Len(OptCond.CondInfoI(intOptCalcLoop).OptIdentifier)
            End If
            If Mid(OptCond.CondInfoI(intOptCalcLoop).OptIdentifier, intOptModePosi + 2) <> Mid(strTargetIden, intOptModePosiRef + 2) Then
                GoTo NextOptCond
            End If
            
            '倍数格納
            dblOptTimes = CDbl(Mid(OptCond.CondInfoI(intOptCalcLoop).OptIdentifier, intOptTimesPosi, intOptModePosi - intOptTimesPosi + 1))
            
            '基本Modeの光量から演算
            dblOptCalcLux = dblOptTimes * OptCond.CondInfo(strTargetIden).AxisLevel
                
            'シートに書き込み
            If sub_UpdateOpt(OptCond.CondInfoI(intOptCalcLoop).OptIdentifier, dblOptCalcLux, "Lux", True) = False Then
                '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
                If g_RunAutoFlg_PALS = False Then
                    MsgBox "Error [OptCalculate] Don't Write Lux!!", vbCritical, PALS_ERRORTITLE
                End If
                '<<<2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
                Exit Function
            End If
            
        End If
NextOptCond:
    Next intOptCalcLoop

End Function

'********************************************************************************************
' 名前: sub_RunLoopAuto
' 内容: 指定した測定回数・ノード情報で、ループ測定を自動で実施する
' 引数: lngAdjCnt    : 調整回数
'       lngAveCnt    : 平均回数
'       intSwNode    : テスタノード
' 戻値: 終了フラグ
' 備考： なし
' 更新履歴： Rev1.0      2011/09/15　新規作成   M.Imamura
'********************************************************************************************

Public Function sub_RunOptAuto(ByVal lngAdjCnt As Long, ByVal lngAveCnt As Long, ByVal intSwNode As Integer) As Long

    g_RunAutoFlg_PALS = True
    sub_RunOptAuto = 1
    Sw_Node = intSwNode

On Error GoTo errPALSsub_RunOptAuto
    ThisWorkbook.Activate

    PALS_ParamFolder = ThisWorkbook.Path & "\" & PALS_PARAMFOLDERNAME
    Call sub_PalsFileCheck

    Set PALS = Nothing
    Set PALS = New csPALS

    'TestConditionシートデータの再読込
    Call ReadCategoryData

    With frm_PALS_OptAdj_Main
        .Show vbModeless
        .cbo_AdjNum.Value = lngAdjCnt
        .cbo_AveNum.Value = lngAveCnt
        Call .cmd_start_Click
    End With

    If OptCond.IllumMaker = NIKON And g_blnUseCSV = True Then
        Call sub_OutPutCsv(NIKON_WRKSHT_NAME, OptFileName, False)
    ElseIf OptCond.IllumMaker = INTERACTION And g_blnUseCSV = True Then
        Call sub_OutPutCsv(IA_WRKSHT_NAME, OptFileName, False)
    End If

    Unload frm_PALS_OptAdj_Main

    If g_ErrorFlg_PALS = True Then
        GoTo errPALSsub_RunOptAuto
    End If
    
    g_RunAutoFlg_PALS = False
    sub_RunOptAuto = 0
    Exit Function

errPALSsub_RunOptAuto:
    g_RunAutoFlg_PALS = False
    g_ErrorFlg_PALS = False

End Function

