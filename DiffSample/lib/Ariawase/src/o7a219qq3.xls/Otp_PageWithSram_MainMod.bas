Attribute VB_Name = "Otp_PageWithSram_MainMod"
Option Explicit

Public Function otp_test_f() As Double

'+++ Test Infomation +++
'OTP��LOGIC�`�F�b�N
'+++++++++++++++++++++++

'*** Result Infomation ***
'Judge_value = 0  Fail
'Judge_value = 1  Pass
'*************************

    '----- �W���ϐ�(�^�C�v����) -----
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
'OTP��Blank(ALL "L")�`�F�b�N
'�������肩�đ��肩�̃`�F�b�N
'+++++++++++++++++++++++

'*** Result Infomation ***
'Judge_value = 0  Fail(No Blank or ECC Fail)
'Judge_value = 1  Pass(Blank)
'Judge_value = 2  Pass(ReTest 1stTestPass)
'*************************

    '----- �W���ϐ�(�^�C�v����) -----
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

    '========== PatRun ========================================================  '�܂���Blank���ǂ����̃`�F�b�N
    For NowPage = 0 To OtpPageEnd                                                'Page�܂킵
        Call PatRun("BlankCheckPage" & CStr(NowPage) & "_Pat")                   'Blank�`�F�b�N�p�^�[���iALL "L"���Ғl��RollCall�j
        
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If TheHdw.Digital.FailedPinsCount(site) > 0 Then BlankCheck(site) = BlankCheck(site) + 1    'Fail�`�F�b�N
            End If
        Next site
    Next NowPage
    

    '========== PASS/FAIL CHECK ===============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If BlankCheck(site) = 0 Then                                         '�SPage�Ƃ�Blank�Ȃ�ŏI���ʂ�Pass�B
                Judge_value(site) = 1
            Else                                                                 'NoBlank��Page����������A��U�A�ŏI���ʂ�Fail�B����ɍđ���`�F�b�N�t���O�𗧂Ă�B
                Judge_value(site) = 0
                Flg_ReTest = Flg_ReTest + 1
            End If
        End If
    Next site
    
    
    '========== BLANK CHECK NG -> RE TEST CHECK ===============================  'Blank����Ȃ���Ύ��́A�đ��肩�̃`�F�b�N
    If Flg_ReTest > 0 Then
            
        '========== PatRun ====================================================
        For NowPage = 0 To OtpPageEnd                                            'Page�܂킵
            Call PatRun("OtpFixedValueCheckPage" & CStr(NowPage) & "_Pat")       '�Œ�l�`�F�b�N�p�^�[��
            
            For site = 0 To nSite
                If TheExec.sites.site(site).Active = True Then
                    If TheHdw.Digital.FailedPinsCount(site) > 0 Then ReTestCheck(site) = ReTestCheck(site) + 1
                End If
            Next site
        Next NowPage
            
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If Judge_value(site) = 0 And ReTestCheck(site) = 0 Then                '�Œ�l�`�F�b�N���SPage�Ƃ�OK�Ȃ�A�đ���Ƃ������ƂŁA�ŏI���ʂ�Pass�ɂ���
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
'OTP��Blow
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
    Dim DefectOverFlow(nSite) As Double      '���ו␳�̏���`�F�b�N�t���O
    Dim DefectNgAddress(nSite) As Double     'NG-ZONE�ɑ��݂��錇�׌�
    Dim ArgArr() As String
    Const ByteBit = 8
    Dim ModifyDataAutoBlow(ByteBit - 1) As String
    Dim AutoBlowBitCnt As Long

    Call SiteCheck
    TheResult.GetResult "OTP_BLA", OTP_BLA

    '========== Condition Set =================================================
    Call Set_OtpCondition(ArgArr)


If Flg_OTP_BLOW = 1 Then    'NoBin�΍�i���׌����z�����𒴂��ė���\�������邽�߁j

'Blow Infomation Choice�@����������������������������������������������������������������������

        Call MakeBlowData_Lot(BitWidthAll_Lot1, Page_Lot1, Bit_Lot1, 1)
        Call MakeBlowData_Lot(BitWidthAll_Lot2, Page_Lot2, Bit_Lot2, 2)
        Call MakeBlowData_Lot(BitWidthAll_Lot7, Page_Lot7, Bit_Lot7, 7)
        Call MakeBlowData_Lot(BitWidthAll_Lot8, Page_Lot8, Bit_Lot8, 8)
        Call MakeBlowData_Lot(BitWidthAll_Lot9, Page_Lot9, Bit_Lot9, 9)
        Call MakeBlowData_Wafer(BitWidthAll_Wafer, Page_Wafer, Bit_Wafer)
        Call MakeBlowData_Chip(BitWidthAll_Chip, Page_Chip, Bit_Chip)
        Call MakeBlowData_Temp(BitWidth_O_TEMP, BitWidth_S_TEMP, Page_TEMP, Bit_TEMP, "TMP_OFS", "TMP_SLP")
        Call MakeBlowData_Defect_SinCpFd(MaxRepair_Single_CP_FD, BitWidthN_Single_CP_FD, BitWidthX_Single_CP_FD, BitWidthY_Single_CP_FD, BitWidthS_Single_CP_FD, BitWidthD_Single_CP_FD, DefRep_SrcType_Single_CP_FD, NgAddress_LeftS_Single_CP_FD, NgAddress_LeftE_Single_CP_FD, NgAddress_RightS_Single_CP_FD, NgAddress_RightE_Single_CP_FD, Page_Single_CP_FD, Bit_Single_CP_FD, DefectOverFlow, DefectNgAddress, "DKH_FDL_Z2D", "HLD_33FSC", "OF_FDL_Z2D", "OF_ZL1")


'Blow Infomation Choice�@����������������������������������������������������������������������


    '��������������������������������������������������������������������������
    '��Pattern Modify
    '��������������������������������������������������������������������������
    For site = 0 To nSite                                                           'Site�܂킵
        If TheExec.sites.site(site).Active = True Then                              'ActiveSite�܂킵
            For NowPage = 0 To OtpPageEnd                                           'Page�܂킵
                If Flg_ModifyPage(NowPage) = True Then                              'BlowPage�̂�Modify
                    
                    ByteCount = 0
                    AutoBlowBitCnt = 0
                    
                    '===== Make Blow Modify Data ==================================
                    For NowBit = 0 To BitParPage(NowPage) - 1                       'Bit�܂킵
                        ByteCount = ByteCount + 1                                   'Byte�J�E���^�[�@�C���N�������g

                        If ByteCount < ByteBit Then                                 'Bit1�`7�͂��̂܂�
                            ModifyDataDeb = ModifyDataDeb & BlowDataAllBin(site, NowPage, NowBit)
                        ElseIf ByteCount = ByteBit Then                             'Bit8�̎��́AACK���Ƃ��Č���"X"��t�����
                            ModifyDataDeb = ModifyDataDeb & BlowDataAllBin(site, NowPage, NowBit) & "X"
                        End If
                        
                        If ByteCount = 8 Then ByteCount = 0                         'Byte�J�E���^�[�@�C���N�������g
                    Next NowBit
                    
                    LenLen = Len(ModifyDataDeb)                                     'Modify����Bit����Get
                    ReDim ModifyData(LenLen - 1) As String
                    For iiii = 1 To LenLen                                          'Modify����z��`���ɒu������
                        ModifyData(iiii - 1) = Mid(ModifyDataDeb, iiii, 1)
                    Next iiii

                    '===== Let's Modify ===========================================
                    TheHdw.Digital.Patterns.pat("OtpBlowPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockDataSITE Label_OtpBlow & CStr(NowPage), 0, RejiIn, ModifyData, site

                    
                    '===== Make Blow Modify Data (AutoBlow Register) ==============
                    For iiii = (LenLen - 2) To ((LenLen - 1) - ByteBit) Step -1     'AutoBlow�p��Modify����z��`���ɒu������
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

    '��������������������������������������������������������������������������
    '��OTPBLOW���s
    '��������������������������������������������������������������������������

    '========== ��Blow�̓����l�Ƃ���-1��ݒ� ==================================
    For site = 0 To nSite
        Judge_value(site) = -1
    Next site


    If Flg_OTP_BLOW = 1 Then

        '========== OTPBLOW���K�v�łȂ�Site�́A�ꎞ�I��Disable�ɂ��� ==========
        Call ActiveSite_Check_OTP
        For site = 0 To nSite
            If OTP_BLA(site) = 1 And Flg_ActiveSite_OTP(site) = 1 And DefectOverFlow(site) = 0 And DefectNgAddress(site) = 0 Then   '�������肩�ASiteActive���A���ו␳����𒴂��Ă��Ȃ����A�A�h���X��OK
                BlowExec_Site = BlowExec_Site + 1
            Else
                TheExec.sites.site(site).Active = False
            End If
        Next site

        '========== OTPBLOW ===================================================
        If BlowExec_Site >= 1 Then
        
            For NowPage = 0 To OtpPageSize - 1                                      'Page�܂킵(Blow�p�^�[�����y�[�W���ƂɈقȂ邩��)
                If Flg_OtpBlowPage(NowPage) = True Then                             'OTPBlow�̕K�v������Page��True�Ő�ɐi��
                    If NowPage <> Flg_OtpBlowFixValPage Then                        '�Œ�l�����݂���Page�͌�񂵁B�u��΍�B
                        Call PatRun("OtpBlowPage" & CStr(NowPage) & "_Pat")         'Blow�p�^�[�����s
                        For site = 0 To nSite                                       'Site����FailPinCount(����ɊO���Ńy�[�W���܂���Ă�)
                            If TheExec.sites.site(site).Active = True Then
                                If TheHdw.Digital.FailedPinsCount(site) > 0 Then
                                    BlowCheck(site) = BlowCheck(site) + 1
                                End If
                            End If
                        Next site
                    End If
                End If
            Next NowPage
            
            '========== OTPBLOW(�u��΍�ŌŒ�lPage�����) =================
            Call PatRun("OtpBlowPage" & CStr(Flg_OtpBlowFixValPage) & "_Pat")      'Blow�p�^�[�����s
            For site = 0 To nSite                                                  'Site����FailPinCount(����ɊO���Ńy�[�W���܂���Ă�)
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

        '========== �␳����A�E�g��đ���̎� ================================
        For site = 0 To nSite
            If DefectOverFlow(site) = 1 Then Judge_value(site) = -2                                 '�␳����A�E�g��Fail
            If DefectNgAddress(site) > 0 Then Judge_value(site) = -3                                'NG-ZONE�̌��ׂ������Fail
            If DefectOverFlow(site) = 1 And DefectNgAddress(site) > 0 Then Judge_value(site) = -4   '�␳����A�E�g ���� NG-ZONE�̌��ׂ���
            If OTP_BLA(site) = 2 Then Judge_value(site) = 2                                         '����ł��đ���Ȃ�K��Pass
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
            Call PatRun("OtpBlow_Break_Pat")                                        'FFF Blow�p�p�^�[��
        End If

        '========== �ꎞ�I��Disable�ɂ��Ă���Site��Active�߂� =================
        Call ActiveSite_Return_OTP

    End If

    If Flg_Debug = 1 Then Call Output_OtpBlowData
    
    Call test(Judge_value)
End Function

Public Function otp_vr_f() As Double

'+++ Test Infomation +++++
'OTP��Verify
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

    
    '��������������������������������������������������������������������������
    '��Pattern Modify
    '��������������������������������������������������������������������������
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


    '��������������������������������������������������������������������������
    '��FailPinsCount
    '��������������������������������������������������������������������������
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
            Call PatRun("OtpBlow_Break_Pat")                                     'FFF Blow�p�p�^�[��
            Call ActiveSite_Return_OTP
        End If
    End If

    If Flg_Debug = 1 Then Call Output_OtpReadData

    Call test(Judge_value)
End Function

Public Function otp_bwc_f() As Double

'+++ Test Infomation +++++
'OTP�̌Œ�lBlow�m�F
'+++++++++++++++++++++++++

'*** Result Infomation ***
'Otpbwc = 0  Fail(�Œ�lBlow NG)
'Otpbwc = 1  Pass(�Œ�lBlow OK)
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
        Call PatRun("OtpFixedValueCheckPage" & CStr(NowPage) & "_Pat")       '�Œ�l�`�F�b�N�p�^�[��
        
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
            ActiveSiteNow = ActiveSiteNow + 1               '����ActiveSite�����J�E���g
            If Judge_value(site) = 0 Then
                OTPBWC_ERR = OTPBWC_ERR + 1                 'OTPBWC��FailSite�����J�E���g
            End If
        End If
    Next site

    If ActiveSiteNow = OTPBWC_ERR And Flg_Tenken = 0 Then       'TENKEN���肾������StopPMC�͍s��Ȃ�
        OTPBWC_ERR = 1
        TheExec.Datalog.WriteComment "ERROR!! OTPBWC!!"
    Else
        OTPBWC_ERR = 0
    End If

    Call test(Judge_value)
End Function

