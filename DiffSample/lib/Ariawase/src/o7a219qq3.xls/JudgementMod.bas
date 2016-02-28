Attribute VB_Name = "JudgementMod"
'****************************************************************************************
' TEST PROGRAM FOR IP750 @ TERADYNE
'
'Global Variables Module - This module is called "JudgementMod" is written Setup
'
'Revision History:
'Jun 18/2001 S.KOMEO @CCD TEST : Created this Module
'July 17/2001 S.KOMEO @CCD TEST : Modified ICX282AQ
'
'****************************************************************************************

Option Explicit

Public DisableSiteCount(nSite) As Long

Public parmFlag As Long
Public glngUnits As Long ' fuku edit


'NAME VBT/Flow  : test
'SYNOPSIS       : test(result() As Double,lngTestStatus() As Long)
'DESCRIPTION    : Judge test result
'ARGUMENTS      : result()          - Test Result
'RETURN VALUE   : none
'SIDE EFFECTS   : none
'SEE ALSO       : Please see ResultReport sub-routing
'AUTHOR         : Yonemura@TERADYNE
'REVISION       : Copied from ILX135 JOB

Public Sub test(result() As Double)

    '========== Default Variables ==========
    Dim lngTestStatus As Long
    Dim lngChannelNumber As Long
    Dim dblLoLimit As Double
    Dim dblHiLimit As Double
    Dim lngParmFlag As Long
    Dim lngUnits As Long
    Dim lngForceUnits As Long
    Dim strPinNameInput As String
    Dim loc As Long
    Dim lngHiLoLimValid As Long
    Dim dblTestResult As Double
    Dim dblForceValue As Double
    Dim site As Long
    '========================================
    
    If IsRankEnable Then Call get_testresult(result)
    
    Call GetLimData(dblLoLimit, dblHiLimit, lngHiLoLimValid)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            dblTestResult = result(site)
            lngTestStatus = PassFail(dblTestResult, dblLoLimit, dblHiLimit, lngHiLoLimValid)
            Call ResultReport(site, lngChannelNumber, dblTestResult, dblLoLimit, dblHiLimit, lngParmFlag, lngTestStatus, lngUnits, dblForceValue, lngForceUnits, loc, strPinNameInput)
            If lngTestStatus = logTestFail Then 'JIDOUKA
                ng_data(site) = result(site)
            End If
        End If
    Next site

End Sub

'NAME VBT/Flow  : PassFail
'SYNOPSIS       : PassFail(ByVal dblTestResult As Double, ByVal dblLowLimit As Double, ByVal dblHighLimit As Double, ByVal strHiLoLimValid As String, lngTestStatus As Long)
'DESCRIPTION    : Judge test result
'ARGUMENTS      : dblTestResult     - Test Result
'               : dblLoLimit        - Lo Limit
'               : dblHiLimit        - Hi Limit
'               : lngHiLoLimValid   - Hi Limit & Lo Limit valid status
'RETURN VALUE   : Return logTestPass or logTestFail
'SIDE EFFECTS   : none
'SEE ALSO       : Please see ResultReport sub-routing
'AUTHOR         : N.SHIN@TERADYNE
'REVISION       : Feb 29/2000

Public Function PassFail(ByVal dblTestResult As Double, ByVal dblLoLimit As Double, ByVal dblHiLimit As Double, ByVal lngHiLoLimValid As Long) As Long

    If ((dblTestResult < dblLoLimit) And ((lngHiLoLimValid = TL_C_HILIM1LOLIM1) Or (lngHiLoLimValid = TL_C_HILIM0LOLIM1))) Then
        PassFail = logTestFail
        parmFlag = parmLow
    ElseIf ((dblTestResult > dblHiLimit) And ((lngHiLoLimValid = TL_C_HILIM1LOLIM1) Or (lngHiLoLimValid = TL_C_HILIM1LOLIM0))) Then
        PassFail = logTestFail
        parmFlag = parmHigh
    ElseIf ((dblTestResult >= dblLoLimit) Or (dblTestResult <= dblHiLimit)) Then
        PassFail = logTestPass
        parmFlag = parmPass
    End If
    
End Function

Public Sub ResultReport(ByVal site As Long, ByVal lngChannelNumber As Long, ByVal dblTestResult As Double, _
                        ByVal dblLoLimit As Double, ByVal dblHiLimit As Double, _
                        ByVal lngParmFlag As Long, ByVal lngTestStatus As Long, _
                        ByVal lngUnits As Long, dblForceValue As Double, lngForceUnits As Long, _
                        loc As Long, Optional strPinNameInput As String)

    Dim strPinName As String
    Dim lngTestNumber As Long

'06/09/15 Add Start
    Dim intHitCound As Integer      '条件ヒット数
'    Const SITE_MAX  As Integer = 1  'サイト数　-★代わりがあるならば変更必須
'06/09/15 Add End

    lngParmFlag = parmFlag
    
    If strPinNameInput <> TL_C_EMPTYSTR Then
        strPinName = strPinNameInput
        lngChannelNumber = -1
    Else
        strPinName = "Empty"
    End If

    lngTestNumber = TheExec.sites.site(site).TestNumber
    
    If lngTestStatus <> logTestPass Then
        If True <> TheExec.RunOptions.DoAll Then
            If True = TheHdw.Digital.Patgen.IsRunning Then
'                thehdw.Digital.Patgen.Halt
            End If
        End If
        
        TheExec.sites.site(site).TestResult = siteFail
'        /*** 17/Mar/02 takayama append ***/
'''        If TheExec.RunOptions.DoAll = False Then     'ORG
        If TheExec.CurrentJob = NormalJobName Then      '31/Mar/02 takayama modified
            DisableSiteCount(site) = 1
        End If
        
        If TheExec.CurrentJob = "TENKEN" Then
            If Ng_test(site) < 1000 Then
                TheExec.sites.site(site).Active = False
                DisableSiteCount(site) = 1
                Call RouteSetup(lngTestStatus, site)
            End If
        End If
                  
'06/09/15 Delete End
'06/09/15 Add Start
        '必須条件なので別に明記
        If DisableSiteCount(0) = 1 And _
           InGrade = False Then
           
'★---------------------------Start
            Dim i As Long  'ループ変数
            intHitCound = 0
            For i = 1 To nSite Step 1
                If DisableSiteCount(i) = 1 Or TheExec.sites.site(i).Active = False Then
                    intHitCound = intHitCound + 1
                End If
            Next
'★---------------------------End
        
            '全てのサイト（１～７）が条件を満たしていた場合
            If intHitCound = SITE_MAX - 1 Then
                '処理を実行する。
                Call RouteSetup(lngTestStatus, site)
            End If
        End If
'06/09/15 Change End
    Else
        TheExec.sites.site(site).TestResult = sitePass
    End If
    '*** fuku edit ***
    '--- add unit into datalog ---
    lngUnits = glngUnits
    '-----------------------------

    If Flg_Print = 1 Then
        Call printMyResult(site, lngTestNumber, lngTestStatus, lngParmFlag, _
                                                   strPinName, lngChannelNumber, dblLoLimit, dblTestResult, dblHiLimit, _
                                                   lngUnits, dblForceValue, lngForceUnits, loc)
    ElseIf Flg_Print = 0 Then
        Call TheExec.Datalog.WriteParametricResult(site, lngTestNumber, lngTestStatus, lngParmFlag, _
                                                   strPinName, lngChannelNumber, dblLoLimit, dblTestResult, dblHiLimit, _
                                                   lngUnits, dblForceValue, lngForceUnits, loc)
    End If

End Sub

'NAME VBT/Flow  : GetLimData
'SYNOPSIS       : GetLimData(ByRef dblLoLimit As Double, ByRef dblHiLimit As Double, ByRef lngHiLoLimValid As String)
'DESCRIPTION    : Get Argument from Test Instance Sheet.
'                 Argument datas are read are Test Lo/Hi Limit & Limit Valid.
'ARGUMENTS      : dblLoLimit        - Test Lo Limit     - Arg0(Test Instance Sheet)
'               : dblHiLimit        - Test Hi Limit     - Arg1(Test Instance Sheet)
'               : lngHiLoLimValid   - Test Limit Valid  - Arg2(Test Instance Sheet)
'RETURN VALUE   : These Arguments are ByRef Args.
'SIDE EFFECTS   : none
'SEE ALSO       : Please see PassFail Function
'AUTHOR         : S.KOMEO @CCD TEST
'REVISION       : Feb 29/2000

Public Sub GetLimData(ByRef dblLoLimit As Double, ByRef dblHiLimit As Double, ByRef lngHiLoLimValid As Long)
    
    '********** Private Variables **********
    Dim strArgList() As String
    Dim lngArgCnt As Long
    
    Call TheExec.DataManager.GetArgumentList(strArgList, lngArgCnt)
    dblLoLimit = val(strArgList(5 * LimitSetIndex + 0))
    dblHiLimit = val(strArgList(5 * LimitSetIndex + 1))
    If strArgList(5 * LimitSetIndex + 2) = "3" Then
        lngHiLoLimValid = TL_C_HILIM1LOLIM1
    ElseIf strArgList(5 * LimitSetIndex + 2) = "1" Then
        lngHiLoLimValid = TL_C_HILIM0LOLIM1
    ElseIf strArgList(5 * LimitSetIndex + 2) = "2" Then
        lngHiLoLimValid = TL_C_HILIM1LOLIM0
    End If
    
    '*** fuku edit ***
    '--- add unit into datalog ---
    Select Case strArgList(5 * LimitSetIndex + 3)
        Case "A":
            glngUnits = unitAmp
        Case "V":
            glngUnits = unitVolt
        Case "Hz":
            glngUnits = unitHz
        Case "S"
            glngUnits = unitTime
        Case Else:
            glngUnits = unitNone
    End Select
    
    
End Sub

Public Sub TestEndSetup()

    If True = TheHdw.Digital.Patgen.IsRunning Then
        TheHdw.Digital.Patgen.Halt
    End If
    
End Sub


