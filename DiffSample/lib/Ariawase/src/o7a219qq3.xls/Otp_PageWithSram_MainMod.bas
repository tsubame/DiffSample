Attribute VB_Name = "Otp_PageWithSram_MainMod"
Option Explicit

Public Function otp_test_f() As Double

'+++ Test Infomation +++
'OTPのLOGICチェック
'+++++++++++++++++++++++

'*** Result Infomation ***
'Judge_value = 0  Fail
'Judge_value = 1  Pass
'*************************

    '----- 標準変数(タイプ共通) -----
    Dim site As Long
    Dim ArgArr() As String
    Dim Judge_value(nSite) As Double
    
    Call SiteCheck
    
    '========== Condition Set =================================================
    Call Set_OtpCondition(ArgArr)

    '========== PatRun ========================================================
    Call PatRun(ArgArr(2))

    '========== PASS/FAIL CHECK ===============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If TheHdw.Digital.FailedPinsCount(site) = 0 Then
                Judge_value(site) = 1
            Else
                Judge_value(site) = 0
            End If
        End If
    Next site


    Call test(Judge_value)
    
End Function

Public Function otp_bla_f() As Double

'+++ Test Infomation +++
'OTPのBlank(ALL "L")チェック
'初期測定か再測定かのチェック
'+++++++++++++++++++++++

'*** Result Infomation ***
'Judge_value = 0  Fail(No Blank or ECC Fail)
'Judge_value = 1  Pass(Blank)
'Judge_value = 2  Pass(ReTest 1stTestPass)
'*************************

    '----- 標準変数(タイプ共通) -----
    Dim site As Long
    Dim NowPage As Integer
    Dim BlankCheck(nSite) As Integer
    Dim ReTestCheck(nSite) As Integer
    Dim Flg_ReTest As Integer
    Dim ArgArr() As String
    Dim Judge_value(nSite) As Double
    
    Call SiteCheck
    
    
    '========== Condition Set =================================================
    Call Set_OtpCondition(ArgArr)

    '========== PatRun ========================================================  'まずはBlankかどうかのチェック
    For NowPage = 0 To OtpPageEnd                                                'Pageまわし
        Call PatRun("BlankCheckPage" & CStr(NowPage) & "_Pat")                   'Blankチェックパターン（ALL "L"期待値のRollCall）
        
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If TheHdw.Digital.FailedPinsCount(site) > 0 Then BlankCheck(site) = BlankCheck(site) + 1    'Failチェック
            End If
        Next site
    Next NowPage
    

    '========== PASS/FAIL CHECK ===============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If BlankCheck(site) = 0 Then                                         '全PageともBlankなら最終結果をPass。
                Judge_value(site) = 1
            Else                                                                 'NoBlankのPageがあったら、一旦、最終結果をFail。さらに再測定チェックフラグを立てる。
                Judge_value(site) = 0
                Flg_ReTest = Flg_ReTest + 1
            End If
        End If
    Next site
    
    
    '========== BLANK CHECK NG -> RE TEST CHECK ===============================  'Blankじゃなければ次は、再測定かのチェック
    If Flg_ReTest > 0 Then
            
        '========== PatRun ====================================================
        For NowPage = 0 To OtpPageEnd                                            'Pageまわし
            Call PatRun("OtpFixedValueCheckPage" & CStr(NowPage) & "_Pat")       '固定値チェックパターン
            
            For site = 0 To nSite
                If TheExec.sites.site(site).Active = True Then
                    If TheHdw.Digital.FailedPinsCount(site) > 0 Then ReTestCheck(site) = ReTestCheck(site) + 1
                End If
            Next site
        Next NowPage
            
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If Judge_value(site) = 0 And ReTestCheck(site) = 0 Then                '固定値チェックが全PageともOKなら、再測定ということで、最終結果をPassにする
                    Judge_value(site) = 2
                End If
            End If
        Next site
    
    End If

    TheResult.Add "OTP_BLA", Judge_value
    Call test(Judge_value)

End Function

Public Function otp_blw_f() As Double

'+++ Test Infomation +++++
'OTPのBlow
'+++++++++++++++++++++++++

'*** Result Infomation ***
'Otpblw = -4 Fail(Defect NG ZONE) & (Defect Max Repair Over Flow)
'Otpblw = -3 Fail(Defect NG ZONE)
'Otpblw = -2 Fail(Defect Max Repair Over Flow)
'Otpblw = -1 Fail(No Blow)
'Otpblw = 0  Fail(Blow NG)
'Otpblw = 1  Pass(Blow OK)
'Otpblw = 2  Pass(ReTest Skip)
'*************************

    Dim site As Long
    Dim BlowExec_Site As Integer
    Dim Flg_OtpBreak As Integer
    Dim NowPage As Integer
    Dim NowBit As Integer
    Dim ModifyDataDeb As String
    Dim LenLen As Long
    Dim iiii As Long
    Dim ByteCount  As Long
    Dim BlowCheck(nSite) As Double
    Dim Judge_value(nSite) As Double
    Dim OTP_BLA() As Double
    Dim DefectOverFlow(nSite) As Double      '欠陥補正の上限チェックフラグ
    Dim DefectNgAddress(nSite) As Double     'NG-ZONEに存在する欠陥個数
    Dim ArgArr() As String
    Const ByteBit = 8
    Dim ModifyDataAutoBlow(ByteBit - 1) As String
    Dim AutoBlowBitCnt As Long

    Call SiteCheck
    TheResult.GetResult "OTP_BLA", OTP_BLA

    '========== Condition Set =================================================
    Call Set_OtpCondition(ArgArr)


If Flg_OTP_BLOW = 1 Then    'NoBin対策（欠陥個数が配列上限を超えて来る可能性があるため）

'Blow Infomation Choice　↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

        Call MakeBlowData_Lot(BitWidthAll_Lot1, Page_Lot1, Bit_Lot1, 1)
        Call MakeBlowData_Lot(BitWidthAll_Lot2, Page_Lot2, Bit_Lot2, 2)
        Call MakeBlowData_Lot(BitWidthAll_Lot7, Page_Lot7, Bit_Lot7, 7)
        Call MakeBlowData_Lot(BitWidthAll_Lot8, Page_Lot8, Bit_Lot8, 8)
        Call MakeBlowData_Lot(BitWidthAll_Lot9, Page_Lot9, Bit_Lot9, 9)
        Call MakeBlowData_Wafer(BitWidthAll_Wafer, Page_Wafer, Bit_Wafer)
        Call MakeBlowData_Chip(BitWidthAll_Chip, Page_Chip, Bit_Chip)
        Call MakeBlowData_Temp(BitWidth_O_TEMP, BitWidth_S_TEMP, Page_TEMP, Bit_TEMP, "TMP_OFS", "TMP_SLP")
        Call MakeBlowData_Defect_SinCpFd(MaxRepair_Single_CP_FD, BitWidthN_Single_CP_FD, BitWidthX_Single_CP_FD, BitWidthY_Single_CP_FD, BitWidthS_Single_CP_FD, BitWidthD_Single_CP_FD, DefRep_SrcType_Single_CP_FD, NgAddress_LeftS_Single_CP_FD, NgAddress_LeftE_Single_CP_FD, NgAddress_RightS_Single_CP_FD, NgAddress_RightE_Single_CP_FD, Page_Single_CP_FD, Bit_Single_CP_FD, DefectOverFlow, DefectNgAddress, "DKH_FDL_Z2D", "HLD_33FSC", "OF_FDL_Z2D", "OF_ZL1")


'Blow Infomation Choice　↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑


    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■Pattern Modify
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    For site = 0 To nSite                                                           'Siteまわし
        If TheExec.sites.site(site).Active = True Then                              'ActiveSiteまわし
            For NowPage = 0 To OtpPageEnd                                           'Pageまわし
                If Flg_ModifyPage(NowPage) = True Then                              'BlowPageのみModify
                    
                    ByteCount = 0
                    AutoBlowBitCnt = 0
                    
                    '===== Make Blow Modify Data ==================================
                    For NowBit = 0 To BitParPage(NowPage) - 1                       'Bitまわし
                        ByteCount = ByteCount + 1                                   'Byteカウンター　インクリメント

                        If ByteCount < ByteBit Then                                 'Bit1～7はそのまま
                            ModifyDataDeb = ModifyDataDeb & BlowDataAllBin(site, NowPage, NowBit)
                        ElseIf ByteCount = ByteBit Then                             'Bit8の時は、ACK情報として後ろに"X"を付けるよ
                            ModifyDataDeb = ModifyDataDeb & BlowDataAllBin(site, NowPage, NowBit) & "X"
                        End If
                        
                        If ByteCount = 8 Then ByteCount = 0                         'Byteカウンター　インクリメント
                    Next NowBit
                    
                    LenLen = Len(ModifyDataDeb)                                     'ModifyするBit数をGet
                    ReDim ModifyData(LenLen - 1) As String
                    For iiii = 1 To LenLen                                          'Modify情報を配列形式に置き換え
                        ModifyData(iiii - 1) = Mid(ModifyDataDeb, iiii, 1)
                    Next iiii

                    '===== Let's Modify ===========================================
                    TheHdw.Digital.Patterns.pat("OtpBlowPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockDataSITE Label_OtpBlow & CStr(NowPage), 0, RejiIn, ModifyData, site

                    
                    '===== Make Blow Modify Data (AutoBlow Register) ==============
                    For iiii = (LenLen - 2) To ((LenLen - 1) - ByteBit) Step -1     'AutoBlow用のModify情報を配列形式に置き換え
                        AutoBlowBitCnt = AutoBlowBitCnt + 1
                        ModifyDataAutoBlow(ByteBit - AutoBlowBitCnt) = ModifyData(iiii)
                    Next iiii
                    
                    '===== Let's Modify (AutoBlow Register) =======================
                    TheHdw.Digital.Patterns.pat("OtpBlowPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockDataSITE Label_OtpBlowAuto & CStr(NowPage), 0, RejiIn, ModifyDataAutoBlow, site

                    ModifyDataDeb = ""
                    Erase ModifyData()
                    Erase ModifyDataAutoBlow()
                    
                End If
            Next NowPage
        End If
    Next site


End If

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■OTPBLOW実行
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    '========== 未Blowの特性値として-1を設定 ==================================
    For site = 0 To nSite
        Judge_value(site) = -1
    Next site


    If Flg_OTP_BLOW = 1 Then

        '========== OTPBLOWが必要でないSiteは、一時的にDisableにする ==========
        Call ActiveSite_Check_OTP
        For site = 0 To nSite
            If OTP_BLA(site) = 1 And Flg_ActiveSite_OTP(site) = 1 And DefectOverFlow(site) = 0 And DefectNgAddress(site) = 0 Then   '初期測定かつ、SiteActiveかつ、欠陥補正上限を超えていないかつ、アドレスもOK
                BlowExec_Site = BlowExec_Site + 1
            Else
                TheExec.sites.site(site).Active = False
            End If
        Next site

        '========== OTPBLOW ===================================================
        If BlowExec_Site >= 1 Then
        
            For NowPage = 0 To OtpPageSize - 1                                      'Pageまわし(Blowパターンがページごとに異なるから)
                If Flg_OtpBlowPage(NowPage) = True Then                             'OTPBlowの必要があるPageはTrueで先に進む
                    If NowPage <> Flg_OtpBlowFixValPage Then                        '固定値が存在するPageは後回し。瞬低対策。
                        Call PatRun("OtpBlowPage" & CStr(NowPage) & "_Pat")         'Blowパターン実行
                        For site = 0 To nSite                                       'Site毎にFailPinCount(さらに外側でページもまわってる)
                            If TheExec.sites.site(site).Active = True Then
                                If TheHdw.Digital.FailedPinsCount(site) > 0 Then
                                    BlowCheck(site) = BlowCheck(site) + 1
                                End If
                            End If
                        Next site
                    End If
                End If
            Next NowPage
            
            '========== OTPBLOW(瞬低対策で固定値Pageを後回し) =================
            Call PatRun("OtpBlowPage" & CStr(Flg_OtpBlowFixValPage) & "_Pat")      'Blowパターン実行
            For site = 0 To nSite                                                  'Site毎にFailPinCount(さらに外側でページもまわってる)
                If TheExec.sites.site(site).Active = True Then
                    If TheHdw.Digital.FailedPinsCount(site) > 0 Then
                        BlowCheck(site) = BlowCheck(site) + 1
                    End If
                End If
            Next site

        End If

        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If BlowCheck(site) = 0 Then
                    Judge_value(site) = 1
                Else
                    Judge_value(site) = 0
                End If
            End If
        Next site

        '========== 補正上限アウトや再測定の時 ================================
        For site = 0 To nSite
            If DefectOverFlow(site) = 1 Then Judge_value(site) = -2                                 '補正上限アウトはFail
            If DefectNgAddress(site) > 0 Then Judge_value(site) = -3                                'NG-ZONEの欠陥があればFail
            If DefectOverFlow(site) = 1 And DefectNgAddress(site) > 0 Then Judge_value(site) = -4   '補正上限アウト かつ NG-ZONEの欠陥あり
            If OTP_BLA(site) = 2 Then Judge_value(site) = 2                                         'それでも再測定なら必ずPass
        Next site

        '========== OTPBLOW FAIL -> OTP BREAK =================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If Judge_value(site) = 0 Then
                    Flg_OtpBreak = Flg_OtpBreak + 1
                Else
                    TheExec.sites.site(site).Active = False
                End If
            End If
        Next site

        If Flg_OtpBreak > 0 Then
            Call PatRun("OtpBlow_Break_Pat")                                        'FFF Blow用パターン
        End If

        '========== 一時的にDisableにしていたSiteのActive戻し =================
        Call ActiveSite_Return_OTP

    End If

    If Flg_Debug = 1 Then Call Output_OtpBlowData
    
    Call test(Judge_value)
End Function

Public Function otp_vr_f() As Double

'+++ Test Infomation +++++
'OTPのVerify
'+++++++++++++++++++++++++

'*** Result Infomation ***
'Otpvr = 0  Fail(Verify NG)
'Otpvr = 1  Pass(Verify OK)
'Otpvr = 2  Pass(ReTest Skip)
'*************************
    
    Dim site As Long
    Dim NowPage As Integer
    Dim NowBit As Integer
    Dim ReadErr(nSite) As Double
    Dim Judge_value(nSite) As Double
    Dim OTP_BLA() As Double
    Dim Flg_OtpBreak As Integer
    Dim ModifyDataVerifyDeb As String
    Dim ByteCount As Long
    Dim LenLen As Long
    Dim iiii As Long
    Dim ArgArr() As String
    Const ByteBit = 8

    Erase ReadDataAllBin
    
    Call SiteCheck
    TheResult.GetResult "OTP_BLA", OTP_BLA

    '========== Condition Set =================================================
    Call Set_OtpCondition(ArgArr)

    
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■Pattern Modify
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '===== Site Loop ======================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            '===== Page Loop ======================================
            For NowPage = 0 To OtpPageEnd
                If Flg_ModifyPage(NowPage) = True Then  'No Modify Page -> Skip
                    ByteCount = 0
                    '===== Bit Loop =======================================
                    For NowBit = 0 To BitParPage(NowPage) - 1
                        
                        ByteCount = ByteCount + 1
                        If ByteCount < ByteBit Then
                            If SramBlowDataAllBin(site, NowPage, NowBit) = "1" Then 'SRAM REPAIR
                                ModifyDataVerifyDeb = ModifyDataVerifyDeb & "H"
                            Else
                                If BlowDataAllBin(site, NowPage, NowBit) = "0" Then
                                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "L"
                                ElseIf BlowDataAllBin(site, NowPage, NowBit) = "1" Then
                                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "H"
                                End If
                            End If
                        ElseIf ByteCount = ByteBit Then
                            If SramBlowDataAllBin(site, NowPage, NowBit) = "1" Then 'SRAM REPAIR
                                ModifyDataVerifyDeb = ModifyDataVerifyDeb & "HX"
                            Else
                                If BlowDataAllBin(site, NowPage, NowBit) = "0" Then
                                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "LX"
                                ElseIf BlowDataAllBin(site, NowPage, NowBit) = "1" Then
                                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "HX"
                                End If
                            End If
                            ByteCount = 0
                        End If
                        
                    Next NowBit
                    
                    '===== Make Modify Data ===============================
                    LenLen = Len(ModifyDataVerifyDeb)
                    ReDim ModifyData(LenLen - 1) As String
                    For iiii = 1 To LenLen
                        ModifyData(iiii - 1) = Mid(ModifyDataVerifyDeb, iiii, 1)
                    Next iiii

                    '===== Pattern Modify =================================
                    TheHdw.Digital.Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockDataSITE Label_OtpVerify & CStr(NowPage), 0, RejiOut, ModifyData, site

                    '===== Valiable Clear =================================
                    Erase ModifyData()
                    ModifyDataVerifyDeb = ""
                    
                End If
            Next NowPage
        End If
    Next site


    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■FailPinsCount
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    For NowPage = 0 To OtpPageEnd
        '========== PATTERN RUN ===========================================
        TheHdw.Digital.Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").Run ("Start")
    
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If TheHdw.Digital.FailedPinsCount(site) > 0 Then
                    ReadErr(site) = ReadErr(site) + 1
                End If
            End If
        Next site
    Next NowPage
        



    '========== Result ========================================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If OTP_BLA(site) = 2 Then             'ReTest Skip = Pass
                Judge_value(site) = 2
            Else
                If ReadErr(site) = 0 Then
                    Judge_value(site) = 1                     'Verify OK = Pass
                Else
                    Judge_value(site) = 0                     'Verify NG = Fail
                    Flg_OtpBreak = Flg_OtpBreak + 1     'FF Blow Chip Break Flag Set
                End If
            End If
        End If
    Next site

    '========== OTP VERIFY FAIL -> OTP BREAK ==================================
    If Flg_OTP_BLOW = 1 Then
        If Flg_OtpBreak >= 1 Then
            Call ActiveSite_Check_OTP
            For site = 0 To nSite
                If Judge_value(site) >= 1 Then TheExec.sites.site(site).Active = False
            Next site
            Call PatRun("OtpBlow_Break_Pat")                                     'FFF Blow用パターン
            Call ActiveSite_Return_OTP
        End If
    End If

    If Flg_Debug = 1 Then Call Output_OtpReadData

    Call test(Judge_value)
End Function

Public Function otp_bwc_f() As Double

'+++ Test Infomation +++++
'OTPの固定値Blow確認
'+++++++++++++++++++++++++

'*** Result Infomation ***
'Otpbwc = 0  Fail(固定値Blow NG)
'Otpbwc = 1  Pass(固定値Blow OK)
'*************************

    Dim site As Long
    Dim ActiveSiteNow As Integer
    Dim BlowCheck(nSite) As Integer
    Dim Judge_value(nSite) As Double
    Dim NowPage As Integer
    Dim ArgArr() As String
    
    Call SiteCheck

    '========== Condition Set =============================================
    Call Set_OtpCondition(ArgArr)

    '========== PatRun ====================================================
    For NowPage = 0 To OtpPageEnd
        Call PatRun("OtpFixedValueCheckPage" & CStr(NowPage) & "_Pat")       '固定値チェックパターン
        
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If TheHdw.Digital.FailedPinsCount(site) > 0 Then BlowCheck(site) = BlowCheck(site) + 1
            End If
        Next site
    Next NowPage
        
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If BlowCheck(site) = 0 Then
                Judge_value(site) = 1    'Pass
            Else
                Judge_value(site) = 0    'Fail
            End If
        End If
    Next site

            
    '===== STOP PMC ===========================================================
    OTPBWC_ERR = 0
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            ActiveSiteNow = ActiveSiteNow + 1               '今のActiveSite数をカウント
            If Judge_value(site) = 0 Then
                OTPBWC_ERR = OTPBWC_ERR + 1                 'OTPBWCのFailSite数をカウント
            End If
        End If
    Next site

    If ActiveSiteNow = OTPBWC_ERR And Flg_Tenken = 0 Then       'TENKEN測定だったらStopPMCは行わない
        OTPBWC_ERR = 1
        TheExec.Datalog.WriteComment "ERROR!! OTPBWC!!"
    Else
        OTPBWC_ERR = 0
    End If

    Call test(Judge_value)
End Function

