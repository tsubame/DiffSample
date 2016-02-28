Attribute VB_Name = "XEeeAuto_Logic"
'�T�v:
'   IGXL�e���v���[�g�FFunctional_T�ɂ������֐��Q
'
'�ړI:
'
'
'�쐬��:
'   2011/12/21 Ver0.1 D.Maruyama
'   2012/03/02 Ver0.2 D.Maruyama �W���b�W��ǉ�

Option Explicit

Public Logic_judge(nSite) As Double '�_���p

'���e:
'   IGXL�e���v���[�g�FFunctional_T�ւ�StartOfBody�֐�
'   �s���̂�TestCondition�̌Ăяo���̂���
'
'�߂�l�F
'   �e���v���[�g���s����
'   �����FTL_SUCCESS
'   ���s�FTL_ERROR
'
'���ӎ���:
'
Public Function logic_setup(argc As Long, argv() As String) As Long

    Call SiteCheck
    
    If argc <> 1 Then
        logic_setup = TL_ERROR
        MsgBox "The Number of logic_setup's arguments are invalid!"
        Exit Function
    End If
    
    Call TheCondition.SetCondition(argv(0))
    
    logic_setup = TL_SUCCESS
    
End Function

'���e:
'   IGXL�e���v���[�g�FFunctional_T�ւ̃W���b�W�֐�
'�߂�l�F
'
'���ӎ���:
'�@�@�@���̂܂܃R�s�[��������

Public Function judge_LogicTest(argc As Long, argv() As String) As Long
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
    Dim dblReturnVal(20) As Double
    Dim lngReturnVal(20) As Long
    Dim dblTestResult As Double
    Dim lngTestResult As Long
    Dim blnReturnCode As Boolean
    Dim dblForceValue As Double
    
    Dim dblRankResult(nSite) As Double

    Dim site As Long
    
    Call SiteCheck
    dblLoLimit = 1
    dblHiLimit = 1
    lngHiLoLimValid = 3

    If argc > 0 Then
        If argv(0) = "NOBIN" Then
            lngHiLoLimValid = 0
        End If
    End If

    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            lngReturnVal(site) = TheHdw.Digital.FailedPinsCount(site)
            If lngReturnVal(site) = 0 Then
                dblTestResult = 1
                dblRankResult(site) = 1
            ElseIf lngReturnVal(site) >= 1 Then
                dblTestResult = 0
                dblRankResult(site) = 0
                Logic_judge(site) = Logic_judge(site) + 1
            Else
                MsgBox ("Error @judge_LogicTest()")
            End If

            lngTestStatus = PassFail(dblTestResult, dblLoLimit, dblHiLimit, lngHiLoLimValid)
            Call ResultReport_Logic(site, lngChannelNumber, dblTestResult, dblLoLimit, dblHiLimit, lngParmFlag, lngTestStatus, lngUnits, dblForceValue, lngForceUnits, loc, strPinNameInput)
        End If
    Next site

    'For Rank
    If IsRankEnable Then Call get_testresult(dblRankResult)
    
End Function

'���e:
'   ���W�b�N�W���b�W�֐��̌��ʃ��|�[�g
'�߂�l�F
'
'���ӎ���:
'�@�@�@���̂܂܃R�s�[��������

Private Sub ResultReport_Logic(ByVal site As Long, ByVal lngChannelNumber As Long, ByVal dblTestResult As Double, _
                        ByVal dblLoLimit As Double, ByVal dblHiLimit As Double, _
                        ByVal lngParmFlag As Long, ByVal lngTestStatus As Long, _
                        ByVal lngUnits As Long, dblForceValue As Double, lngForceUnits As Long, _
                        loc As Long, Optional strPinNameInput As String)

    Dim strPinName As String
    Dim lngTestNumber As Long
    Dim intHitCound As Integer

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
            End If
        End If
        
        TheExec.sites.site(site).TestResult = siteFail
        If TheExec.CurrentJob = NormalJobName Then
            DisableSiteCount(site) = 1
        End If
        
        If TheExec.CurrentJob = "TENKEN" Then
            If Ng_test(site) < 1000 Then
                TheExec.sites.site(site).Active = False
                DisableSiteCount(site) = 1
                Call RouteSetup(lngTestStatus, site)
            End If
        End If

        If DisableSiteCount(0) = 1 And InGrade = False Then
            '--------- Start -------------
            Dim i As Long  '���[�v�ϐ�
            intHitCound = 0
            For i = 1 To nSite Step 1
                If DisableSiteCount(i) = 1 Or TheExec.sites.site(i).Active = False Then
                    intHitCound = intHitCound + 1
                End If
            Next
            '--------- END ---------------
        
            If intHitCound = nSite Then Call RouteSetup(lngTestStatus, site)
        
        End If
    Else
        TheExec.sites.site(site).TestResult = sitePass
    End If
    
    '--- add unit into datalog ---
    lngUnits = 0
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




