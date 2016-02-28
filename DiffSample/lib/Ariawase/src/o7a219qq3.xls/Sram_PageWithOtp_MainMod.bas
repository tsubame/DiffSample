Attribute VB_Name = "Sram_PageWithOtp_MainMod"
Option Explicit


Public Function rom_min_f() As Double
   Call SramValiableClear

'+++ Test Infomation +++
'ROM��BIST����

'Result = 0  Fail
'Result = 1  Pass
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck
    
    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
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

    '========== Result ========================================================
    Call test(Judge_value)

End Function

Public Function sram1_min_f() As Double

'+++ Test Infomation +++
'RAM��BIST����(Pre)
'BIST��NG�ƂȂ��Ă��A�璷�\�ł����FailStop�͂����Ȃ��B���̏ꍇ�͏璷���BIST�őI�ʁB
'�璷�s�Ƃ������Ƃ����̎��_�ł킩���FailStop������B

'Result = 0  Pass(BIST NG -> Try Repair)
'Result = 1  Pass(BIST OK)
'Result = 2  Fail(BIST NG -> Multi I/O NG)
'Result = 3  Fail(BIST NG -> Nothing Repair I/O)
'Result = 4  Fail(BIST NG -> ALPG NG)
'+++++++++++++++++++++++

    Dim site As Long
    Dim CycleNo() As Long
    Dim MemoryNo() As Integer
    Dim IoNo() As Integer
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
    '========== PatRun & Get BIST Result ======================================
    Call SramPatRun_GetFailInfo(ArgArr(2), Judge_value())

    '========== Result ========================================================
    Call test(Judge_value)

End Function

Public Function sram2_min_f() As Double

'+++ Test Infomation +++
'RAM��BIST����(Pre)
'BIST��NG�ƂȂ��Ă��A�璷�\�ł����FailStop�͂����Ȃ��B���̏ꍇ�͏璷���BIST�őI�ʁB
'�璷�s�Ƃ������Ƃ����̎��_�ł킩���FailStop������B

'Result = 0  Pass(BIST NG -> Try Repair)
'Result = 1  Pass(BIST OK)
'Result = 2  Fail(BIST NG -> Multi I/O NG)
'Result = 3  Fail(BIST NG -> Nothing Repair I/O)
'Result = 4  Fail(BIST NG -> ALPG NG)
'+++++++++++++++++++++++

    Dim site As Long
    Dim CycleNo() As Long
    Dim MemoryNo() As Integer
    Dim IoNo() As Integer
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
    '========== PatRun & Get BIST Result ======================================
    Call SramPatRun_GetFailInfo(ArgArr(2), Judge_value())

    '========== Result ========================================================
    Call test(Judge_value)

End Function


Public Function sram3b_min_f() As Double

'+++ Test Infomation +++
'ROM��BIST����

'Result = 0  Fail
'Result = 1  Pass
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck
    
    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
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

    '========== Result ========================================================
    Call test(Judge_value)

End Function


Public Function rom_hv_f() As Double

'+++ Test Infomation +++
'ROM��BIST����

'Result = 0  Fail
'Result = 1  Pass
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck
    
    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
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

    '========== Result ========================================================
    Call test(Judge_value)

End Function

Public Function sram1_hv_f() As Double

'+++ Test Infomation +++
'RAM��BIST����(Pre)
'BIST��NG�ƂȂ��Ă��A�璷�\�ł����FailStop�͂����Ȃ��B���̏ꍇ�͏璷���BIST�őI�ʁB
'�璷�s�Ƃ������Ƃ����̎��_�ł킩���FailStop������B

'Result = 0  Pass(BIST NG -> Try Repair)
'Result = 1  Pass(BIST OK)
'Result = 2  Fail(BIST NG -> Multi I/O NG)
'Result = 3  Fail(BIST NG -> Nothing Repair I/O)
'Result = 4  Fail(BIST NG -> ALPG NG)
'+++++++++++++++++++++++

    Dim site As Long
    Dim CycleNo() As Long
    Dim MemoryNo() As Integer
    Dim IoNo() As Integer
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
    '========== PatRun & Get BIST Result ======================================
    Call SramPatRun_GetFailInfo(ArgArr(2), Judge_value())

    '========== Result ========================================================
    Call test(Judge_value)

End Function

Public Function sram2_hv_f() As Double

'+++ Test Infomation +++
'RAM��BIST����(Pre)
'BIST��NG�ƂȂ��Ă��A�璷�\�ł����FailStop�͂����Ȃ��B���̏ꍇ�͏璷���BIST�őI�ʁB
'�璷�s�Ƃ������Ƃ����̎��_�ł킩���FailStop������B

'Result = 0  Pass(BIST NG -> Try Repair)
'Result = 1  Pass(BIST OK)
'Result = 2  Fail(BIST NG -> Multi I/O NG)
'Result = 3  Fail(BIST NG -> Nothing Repair I/O)
'Result = 4  Fail(BIST NG -> ALPG NG)
'+++++++++++++++++++++++

    Dim site As Long
    Dim CycleNo() As Long
    Dim MemoryNo() As Integer
    Dim IoNo() As Integer
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
    '========== PatRun & Get BIST Result ======================================
    Call SramPatRun_GetFailInfo(ArgArr(2), Judge_value())

    '========== Result ========================================================
    Call test(Judge_value)

End Function


Public Function sram3b_hv_f() As Double

'+++ Test Infomation +++
'ROM��BIST����

'Result = 0  Fail
'Result = 1  Pass
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck
    
    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
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

    '========== Result ========================================================
    Call test(Judge_value)

End Function


Public Function rom_tck_f() As Double

'+++ Test Infomation +++
'ROM��BIST����

'Result = 0  Fail
'Result = 1  Pass
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck
    
    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
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

    '========== Result ========================================================
    Call test(Judge_value)

End Function

Public Function sram1_tck_f() As Double

'+++ Test Infomation +++
'RAM��BIST����(Pre)
'BIST��NG�ƂȂ��Ă��A�璷�\�ł����FailStop�͂����Ȃ��B���̏ꍇ�͏璷���BIST�őI�ʁB
'�璷�s�Ƃ������Ƃ����̎��_�ł킩���FailStop������B

'Result = 0  Pass(BIST NG -> Try Repair)
'Result = 1  Pass(BIST OK)
'Result = 2  Fail(BIST NG -> Multi I/O NG)
'Result = 3  Fail(BIST NG -> Nothing Repair I/O)
'Result = 4  Fail(BIST NG -> ALPG NG)
'+++++++++++++++++++++++

    Dim site As Long
    Dim CycleNo() As Long
    Dim MemoryNo() As Integer
    Dim IoNo() As Integer
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
    '========== PatRun & Get BIST Result ======================================
    Call SramPatRun_GetFailInfo(ArgArr(2), Judge_value())

    '========== Result ========================================================
    Call test(Judge_value)

End Function

Public Function sram2_tck_f() As Double

'+++ Test Infomation +++
'RAM��BIST����(Pre)
'BIST��NG�ƂȂ��Ă��A�璷�\�ł����FailStop�͂����Ȃ��B���̏ꍇ�͏璷���BIST�őI�ʁB
'�璷�s�Ƃ������Ƃ����̎��_�ł킩���FailStop������B

'Result = 0  Pass(BIST NG -> Try Repair)
'Result = 1  Pass(BIST OK)
'Result = 2  Fail(BIST NG -> Multi I/O NG)
'Result = 3  Fail(BIST NG -> Nothing Repair I/O)
'Result = 4  Fail(BIST NG -> ALPG NG)
'+++++++++++++++++++++++

    Dim site As Long
    Dim CycleNo() As Long
    Dim MemoryNo() As Integer
    Dim IoNo() As Integer
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
    '========== PatRun & Get BIST Result ======================================
    Call SramPatRun_GetFailInfo(ArgArr(2), Judge_value())

    '========== Result ========================================================
    Call test(Judge_value)

End Function


Public Function sram3b_tck_f() As Double

'+++ Test Infomation +++
'ROM��BIST����

'Result = 0  Fail
'Result = 1  Pass
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck
    
    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)
    
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

    '========== Result ========================================================
    Call test(Judge_value)

End Function

Public Function sram4_svdr_f() As Double

'+++ Test Infomation +++
'RAM��BIST����(Pre)
'BIST��NG�ƂȂ��Ă��A�璷�\�ł����FailStop�͂����Ȃ��B���̏ꍇ�͏璷���BIST�őI�ʁB
'�璷�s�Ƃ������Ƃ����̎��_�ł킩���FailStop������B

'Result = 0  Pass(BIST NG -> Try Repair)
'Result = 1  Pass(BIST OK)
'Result = 2  Fail(BIST NG -> Multi I/O NG)
'Result = 3  Fail(BIST NG -> Nothing Repair I/O)
'Result = 4  Fail(BIST NG -> ALPG NG)
'+++++++++++++++++++++++

    Dim site As Long
    Dim CycleNo() As Long
    Dim MemoryNo() As Integer
    Dim IoNo() As Integer
    Dim Judge_value_Write(nSite) As Double
    Dim Judge_value_Read(nSite) As Double
    Dim Judge_value(nSite) As Double
    Dim LowVddPins() As String
    Dim Vdd2nd As String
    Dim Vdd3rd As String
    Dim LowVddWait As Double
    Dim WritePat As String
    Dim ReadPat As String
    Dim ArgArr() As String
    
    Call SiteCheck
    
    '========== Condition Set =================================================
    Call Set_SramCondition_SVDR(ArgArr, LowVddPins, Vdd2nd, Vdd3rd, LowVddWait, WritePat, ReadPat)
    
    '========== PatRun & Get BIST Result ======================================
    Call SramPatRun_GetFailInfo(WritePat, Judge_value_Write())
        
        
    '========== Change Voltage ================================================
    Dim i As Long
    For i = 0 To UBound(LowVddPins)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 400 * mV)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 100 * mV)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 0 * mV)
    Next i

    TheHdw.WAIT LowVddWait * S

    For i = 0 To UBound(LowVddPins)
        ShtPowerV.GetPowerInfo(Vdd3rd, LowVddPins(i)).Force (LowVddPins(i))
    Next i
    
    
    '========== PatRun & Get BIST Result ======================================
    Call SramPatRun_GetFailInfo(ReadPat, Judge_value_Read())
    
    '========== Result ========================================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            If Judge_value_Write(site) = 1 And Judge_value_Read(site) = 1 Then
                Judge_value(site) = 1                           'Pass(BIST OK)
            ElseIf Judge_value_Write(site) >= 2 Or Judge_value_Read(site) >= 2 Then
                If Judge_value_Write(site) > Judge_value_Read(site) Then
                    Judge_value(site) = Judge_value_Write(site) 'Fail(BIST NG)
                Else
                    Judge_value(site) = Judge_value_Read(site)  'Fail(BIST NG)
                End If
            Else
                Judge_value(site) = 0                           'Pass(BIST NG -> Try Repair)
            End If
        End If
    Next site

    Call test(Judge_value)

End Function

Public Function sram5_svdr_f() As Double

'+++ Test Infomation +++
'RAM��BIST����(Pre)
'BIST��NG�ƂȂ��Ă��A�璷�\�ł����FailStop�͂����Ȃ��B���̏ꍇ�͏璷���BIST�őI�ʁB
'�璷�s�Ƃ������Ƃ����̎��_�ł킩���FailStop������B

'Result = 0  Pass(BIST NG -> Try Repair)
'Result = 1  Pass(BIST OK)
'Result = 2  Fail(BIST NG -> Multi I/O NG)
'Result = 3  Fail(BIST NG -> Nothing Repair I/O)
'Result = 4  Fail(BIST NG -> ALPG NG)
'+++++++++++++++++++++++

    Dim site As Long
    Dim CycleNo() As Long
    Dim MemoryNo() As Integer
    Dim IoNo() As Integer
    Dim Judge_value_Write(nSite) As Double
    Dim Judge_value_Read(nSite) As Double
    Dim Judge_value(nSite) As Double
    Dim LowVddPins() As String
    Dim Vdd2nd As String
    Dim Vdd3rd As String
    Dim LowVddWait As Double
    Dim WritePat As String
    Dim ReadPat As String
    Dim ArgArr() As String
    
    Call SiteCheck
    
    '========== Condition Set =================================================
    Call Set_SramCondition_SVDR(ArgArr, LowVddPins, Vdd2nd, Vdd3rd, LowVddWait, WritePat, ReadPat)
    
    '========== PatRun & Get BIST Result ======================================
    Call SramPatRun_GetFailInfo(WritePat, Judge_value_Write())
        
        
    '========== Change Voltage ================================================
    Dim i As Long
    For i = 0 To UBound(LowVddPins)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 400 * mV)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 100 * mV)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 0 * mV)
    Next i

    TheHdw.WAIT LowVddWait * S

    For i = 0 To UBound(LowVddPins)
        ShtPowerV.GetPowerInfo(Vdd3rd, LowVddPins(i)).Force (LowVddPins(i))
    Next i
    
    
    '========== PatRun & Get BIST Result ======================================
    Call SramPatRun_GetFailInfo(ReadPat, Judge_value_Read())
    
    '========== Result ========================================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            If Judge_value_Write(site) = 1 And Judge_value_Read(site) = 1 Then
                Judge_value(site) = 1                           'Pass(BIST OK)
            ElseIf Judge_value_Write(site) >= 2 Or Judge_value_Read(site) >= 2 Then
                If Judge_value_Write(site) > Judge_value_Read(site) Then
                    Judge_value(site) = Judge_value_Write(site) 'Fail(BIST NG)
                Else
                    Judge_value(site) = Judge_value_Read(site)  'Fail(BIST NG)
                End If
            Else
                Judge_value(site) = 0                           'Pass(BIST NG -> Try Repair)
            End If
        End If
    Next site

    Call test(Judge_value)

End Function

Public Function sram_rep_f() As Double

'+++ Test Infomation ++++++
'�ŏI�I�ȏ璷�۔�����s��
'�璷�\�ł����RCON�`���̏璷�f�[�^���쐬����
'++++++++++++++++++++++++++

'*** Result Infomation ***
'Result = 0  Pass(No Repair = PreSRAM ALL Pass)
'Result = 1  Pass(Repairable = Try Blow)
'Result = 10 Pass(No Repairable = Multi I/O Fail)
'Result = 20 Pass(No Repairable = Max Repairable Memory Over)
'*************************


    Dim site As Long
    Dim Cat_P As String
    Dim Judge_value(nSite) As Double
    
    Call SiteCheck
    
    '========== Repair Judge & Make RCON Rrepair Data =========================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Call Judge_Repair_ThisChip(site, Judge_value())     'OUTPUT -> EF_BIST_REPAIR_DATA�ɏ璷�f�[�^�����������
        End If
    Next site


    '========== Result ========================================================
    Call test(Judge_value)

End Function


Public Function sram_blw_f() As Double

'+++ Test Infomation ++++++
'�璷����Blow����B�iBlow�f�[�^�쐬�`�p�^�[��Modeify���܂ށj

'Judge_value = 0  Fail(Blow -> Blow Function Test NG)
'Judge_value = 1  Pass(Blow -> Try Repair)
'Judge_value = 2  Pass(ReTest Pass)
'Judge_value = 3  Pass(No Blow = PreSRAM ALL Pass)
'++++++++++++++++++++++++++


    Dim site As Long
    Dim Data As Long
    Dim BlowData_SramRep(nSite) As String
    Dim BlowExec_Site As Integer
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    Dim OTP_BLA() As Double
    Dim NowPage As Integer
    Dim NowBit As Integer
    Dim ByteCount  As Long
    Const ByteBit = 8
    Dim ModifyDataDeb As String
    Dim LenLen As Long
    Dim iiii As Long
    Dim AutoBlowBitCnt As Long
    Dim ModifyDataAutoBlow(ByteBit - 1) As String
    Dim BlowCheck(nSite) As Double
    Dim Flg_OtpBreak As Integer

    Call SiteCheck
    TheResult.GetResult "OTP_BLA", OTP_BLA


If Flg_PostSramRun = True Then  'PreSRAM���SSite�SPass�Ȃ�ABlow�̓X�L�b�v����


    '========== Condition Set =================================================
    Call Set_SramCondition(ArgArr)


    '��������������������������������������������������������������������������
    '��Make Blow Data
    '��������������������������������������������������������������������������
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If Bist_Repairable_Flag(site) = 1 Then                                  '�璷���K�v��Site�����s��
                
                '===== Repair Data Compression (Case RCON ChanType) ===========
                If RCON_ChainType = "Descending" Then
                    Call Comp_RconData_ChainType_Descending(site)
                ElseIf RCON_ChainType = "Ascending" Then
                    Call Comp_RconData_ChainType_Ascending(site)
                End If
                
                '===== Make Blow Data =========================================
                Call MakeBlowData_Sram(site, Ef_Bist_Rd_En_Width, Ef_Bist_Rd_Addr_Width, Ef_Bist_Rd_Data_Width, Page_SRAM, Bit_SRAM)
                
            End If
        End If
    Next site


    '��������������������������������������������������������������������������
    '��Pattern Modify
    '��������������������������������������������������������������������������
    '===== Site LOOP ==================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If Bist_Repairable_Flag(site) = 1 Then                     '�璷���K�v��Site�����s��
                '===== Page LOOP ==================================
                For NowPage = 0 To OtpPageEnd
                    If Flg_ModifyPageSRAM(NowPage) = True Then         'BlowPage�̂�Modify
                        ByteCount = 0
                        AutoBlowBitCnt = 0
                        '===== Bit LOOP ===================================
                        For NowBit = 0 To BitParPage(NowPage) - 1                       'Bit�܂킵
                            ByteCount = ByteCount + 1                                   'Byte�J�E���^�[�@�C���N�������g
                            
                            '===== Make Blow Modify Data ======================
                            If ByteCount < ByteBit Then                                 'Bit1�`7�͂��̂܂�
                                ModifyDataDeb = ModifyDataDeb & SramBlowDataAllBin(site, NowPage, NowBit)
                            ElseIf ByteCount = ByteBit Then                             'Bit8�̎��́AACK���Ƃ��Č���"X"��t�����
                                ModifyDataDeb = ModifyDataDeb & SramBlowDataAllBin(site, NowPage, NowBit) & "X"
                            End If
                            
                            If ByteCount = 8 Then ByteCount = 0                         'Byte�J�E���^�[�@�C���N�������g
                        Next NowBit
                        
                        '===== Valiable Type Change ===================================
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
            Else
                Judge_value(site) = 3   '�璷���K�v�Ȃ�Site�͓����l3 = Pass
            End If
        End If
    Next site


    '��������������������������������������������������������������������������
    '��OTPBLOW���s
    '��������������������������������������������������������������������������

    If Flg_OTP_BLOW = 1 Then

        '========== OTPBLOW���K�v�łȂ�Site�́A�ꎞ�I��Disable�ɂ��� ==========
        Call ActiveSite_Check_OTP
        For site = 0 To nSite
            If OTP_BLA(site) = 1 And Flg_ActiveSite_OTP(site) = 1 And Bist_Repairable_Flag(site) = 1 Then   '�������肩�ASiteActive���ASRAM�璷���K�v�ł���ꍇ
                BlowExec_Site = BlowExec_Site + 1
            Else
                TheExec.sites.site(site).Active = False
            End If
        Next site

        '========== OTPBLOW ===================================================
        If BlowExec_Site >= 1 Then
            For NowPage = 0 To OtpPageSize - 1                                      'Page�܂킵(Blow�p�^�[�����y�[�W���ƂɈقȂ邩��)
                If Flg_ModifyPageSRAM(NowPage) = True Then
                    Call PatRun("OtpBlowPage" & CStr(NowPage) & "_Pat")             'Blow�p�^�[�����s
                    For site = 0 To nSite                                           'Site����FailPinCount(����ɊO���Ńy�[�W���܂���Ă�)
                        If TheExec.sites.site(site).Active = True Then
                            If TheHdw.Digital.FailedPinsCount(site) > 0 Then
                                BlowCheck(site) = BlowCheck(site) + 1
                            End If
                        End If
                    Next site
                End If
            Next NowPage
        End If
        
        '========== PatRun Result =============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If BlowCheck(site) = 0 Then
                    Judge_value(site) = 1       'Blow Function OK
                Else
                    Judge_value(site) = 0       'Blow Function NG
                End If
            End If
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


    If Flg_Debug = 1 Then Call Output_OtpBlowData_Sram


Else  'PreSRAM���SSite�SPass�Ȃ�A�����l��3����

    '========== Result ====================================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Judge_value(site) = 3
        End If
    Next site

End If

      '�ł��đ��肾������A�����l��2����
    '========== ReTest Result =============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If OTP_BLA(site) = 2 Then
                Judge_value(site) = 2       'RETEST PASS
            End If
        End If
    Next site


    '========== Result ========================================================
    Call test(Judge_value)
    
End Function


Public Function sram_vr_f() As Double

'+++ Test Infomation ++++++
'SRAM�璷�ł�OTP��Verify
'++++++++++++++++++++++++++

'*** Result Infomation ***
'Judge_value = 0  Fail(Verify NG)
'Judge_value = 1  Pass(Verify OK)
'Judge_value = 2  Pass(ReTest Pass)
'Judge_value = 3  Pass(No Verify = PreSRAM ALL Pass)
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
    
    Call SiteCheck
    TheResult.GetResult "OTP_BLA", OTP_BLA


If Flg_PostSramRun = True Then    'PreSRAM���SSite�SPass�Ȃ�AVerify�̓X�L�b�v����


    '========== Condition Set =================================================
    Call Set_OtpCondition(ArgArr)

    
    '��������������������������������������������������������������������������
    '��Pattern Modify
    '��������������������������������������������������������������������������
    '===== Site Loop ======================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If Bist_Repairable_Flag(site) = 1 Then                                  '�璷���K�v��Site�����s��
                '===== Page Loop ======================================
                For NowPage = 0 To OtpPageEnd
                    If Flg_ModifyPageSRAM(NowPage) = True Then  'No Modify Page -> Skip
                        ByteCount = 0
                        '===== Bit Loop =======================================
                        For NowBit = 0 To BitParPage(NowPage) - 1
                            
                            ByteCount = ByteCount + 1
                            If ByteCount < ByteBit Then
                                If SramBlowDataAllBin(site, NowPage, NowBit) = "0" Then
                                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "L"
                                ElseIf SramBlowDataAllBin(site, NowPage, NowBit) = "1" Then
                                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "H"
                                End If
                            ElseIf ByteCount = ByteBit Then
                                If SramBlowDataAllBin(site, NowPage, NowBit) = "0" Then
                                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "LX"
                                ElseIf SramBlowDataAllBin(site, NowPage, NowBit) = "1" Then
                                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "HX"
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
            Else
                Judge_value(site) = 3   '�璷���K�v�Ȃ�Site�͓����l3 = Pass
            End If
        End If
    Next site


    '��������������������������������������������������������������������������
    '��PatRun & FailPinsCount
    '��������������������������������������������������������������������������
    For NowPage = 0 To OtpPageEnd
        If Flg_ModifyPageSRAM(NowPage) = True Then  'No Modify Page -> Skip
            '========== PATTERN RUN ===========================================
            TheHdw.Digital.Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").Run ("Start")
        
            For site = 0 To nSite
                If TheExec.sites.site(site).Active = True Then
                    If TheHdw.Digital.FailedPinsCount(site) > 0 Then
                        ReadErr(site) = ReadErr(site) + 1
                    End If
                End If
            Next site
        End If
    Next NowPage
        


    '========== Result ========================================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If OTP_BLA(site) = 2 Then           'ReTest Skip = Pass
                Judge_value(site) = 2
            Else
                If Judge_value(site) < 2 Then               '2�͍đ��肾����Pass�B3�͏璷���K�v�Ȃ�����Pass�B
                    If ReadErr(site) = 0 Then
                        Judge_value(site) = 1               'Verify OK = Pass
                    Else
                        Judge_value(site) = 0               'Verify NG = Fail
                        Flg_OtpBreak = Flg_OtpBreak + 1     'FF Blow Chip Break Flag Set
                    End If
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


Else  'PreSRAM���SSite�SPass�Ȃ�A�����l��3����

    '========== Result ====================================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Judge_value(site) = 3
        End If
    Next site

End If

      '�ł��đ��肾������A�����l��2����
    '========== ReTest Result =============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If OTP_BLA(site) = 2 Then
                Judge_value(site) = 2       'RETEST PASS
            End If
        End If
    Next site


    '========== Result ========================================================
    Call test(Judge_value)
    
End Function

Public Function sramp1_min_f() As Double

'+++ Test Infomation +++
'Result = 0  Pass(BIST NG)
'Result = 1  Pass(BIST OK)
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    If Flg_PostSramRun = True Then
    
        '========== Condition Set =================================================
        Call Set_SramCondition(ArgArr)
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

    Else
    
        '========== PASS/FAIL CHECK ===============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Judge_value(site) = 1
            End If
        Next site
        
    End If
    
    
    '========== Result ========================================================
    Call test(Judge_value)


End Function

Public Function sramp2_min_f() As Double

'+++ Test Infomation +++
'Result = 0  Pass(BIST NG)
'Result = 1  Pass(BIST OK)
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    If Flg_PostSramRun = True Then
    
        '========== Condition Set =================================================
        Call Set_SramCondition(ArgArr)
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

    Else
    
        '========== PASS/FAIL CHECK ===============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Judge_value(site) = 1
            End If
        Next site
        
    End If
    
    
    '========== Result ========================================================
    Call test(Judge_value)


End Function

Public Function sramp1_hv_f() As Double

'+++ Test Infomation +++
'Result = 0  Pass(BIST NG)
'Result = 1  Pass(BIST OK)
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    If Flg_PostSramRun = True Then
    
        '========== Condition Set =================================================
        Call Set_SramCondition(ArgArr)
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

    Else
    
        '========== PASS/FAIL CHECK ===============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Judge_value(site) = 1
            End If
        Next site
        
    End If
    
    
    '========== Result ========================================================
    Call test(Judge_value)


End Function

Public Function sramp2_hv_f() As Double

'+++ Test Infomation +++
'Result = 0  Pass(BIST NG)
'Result = 1  Pass(BIST OK)
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    If Flg_PostSramRun = True Then
    
        '========== Condition Set =================================================
        Call Set_SramCondition(ArgArr)
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

    Else
    
        '========== PASS/FAIL CHECK ===============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Judge_value(site) = 1
            End If
        Next site
        
    End If
    
    
    '========== Result ========================================================
    Call test(Judge_value)


End Function

Public Function sramp1_tck_f() As Double

'+++ Test Infomation +++
'Result = 0  Pass(BIST NG)
'Result = 1  Pass(BIST OK)
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    If Flg_PostSramRun = True Then
    
        '========== Condition Set =================================================
        Call Set_SramCondition(ArgArr)
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

    Else
    
        '========== PASS/FAIL CHECK ===============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Judge_value(site) = 1
            End If
        Next site
        
    End If
    
    
    '========== Result ========================================================
    Call test(Judge_value)


End Function

Public Function sramp2_tck_f() As Double

'+++ Test Infomation +++
'Result = 0  Pass(BIST NG)
'Result = 1  Pass(BIST OK)
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim ArgArr() As String
    
    Call SiteCheck

    If Flg_PostSramRun = True Then
    
        '========== Condition Set =================================================
        Call Set_SramCondition(ArgArr)
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

    Else
    
        '========== PASS/FAIL CHECK ===============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Judge_value(site) = 1
            End If
        Next site
        
    End If
    
    
    '========== Result ========================================================
    Call test(Judge_value)


End Function


Public Function sramp4_svdr_f() As Double

'+++ Test Infomation +++
'Result = 0  Pass(BIST NG)
'Result = 1  Pass(BIST OK)
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim LowVddPins() As String
    Dim Vdd2nd As String
    Dim Vdd3rd As String
    Dim LowVddWait As Double
    Dim WritePat As String
    Dim ReadPat As String
    Dim ArgArr() As String

    Call SiteCheck

    If Flg_PostSramRun = True Then
    
        '========== Condition Set =================================================
        Call Set_SramCondition_SVDR(ArgArr, LowVddPins, Vdd2nd, Vdd3rd, LowVddWait, WritePat, ReadPat)
    
        '========== PatRun(Write) =================================================
        Call PatRun(WritePat)
        
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

        '========== Change Voltage ================================================
        Dim i As Long
        For i = 0 To UBound(LowVddPins)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 400 * mV)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 100 * mV)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 0 * mV)
        Next i

        TheHdw.WAIT LowVddWait * S

        For i = 0 To UBound(LowVddPins)
            ShtPowerV.GetPowerInfo(Vdd3rd, LowVddPins(i)).Force (LowVddPins(i))
        Next i
        
        '========== PatRun ========================================================
        Call PatRun(ReadPat)
    
        '========== PASS/FAIL CHECK ===============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If TheHdw.Digital.FailedPinsCount(site) = 0 And Judge_value(site) = 1 Then
                    Judge_value(site) = 1
                Else
                    Judge_value(site) = 0
                End If
            End If
        Next site
    Else
        '========== PASS/FAIL CHECK ===============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Judge_value(site) = 1
            End If
        Next site
        
    End If

    '========== Result ========================================================
    Call test(Judge_value)

End Function


Public Function sramp5_svdr_f() As Double

'+++ Test Infomation +++
'Result = 0  Pass(BIST NG)
'Result = 1  Pass(BIST OK)
'+++++++++++++++++++++++

    Dim site As Long
    Dim Judge_value(nSite) As Double
    Dim LowVddPins() As String
    Dim Vdd2nd As String
    Dim Vdd3rd As String
    Dim LowVddWait As Double
    Dim WritePat As String
    Dim ReadPat As String
    Dim ArgArr() As String

    Call SiteCheck

    If Flg_PostSramRun = True Then
    
        '========== Condition Set =================================================
        Call Set_SramCondition_SVDR(ArgArr, LowVddPins, Vdd2nd, Vdd3rd, LowVddWait, WritePat, ReadPat)
    
        '========== PatRun(Write) =================================================
        Call PatRun(WritePat)
        
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

        '========== Change Voltage ================================================
        Dim i As Long
        For i = 0 To UBound(LowVddPins)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 400 * mV)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 100 * mV)
        Call ShtPowerV.GetPowerInfo(Vdd2nd, LowVddPins(i)).ForceWithOffset(LowVddPins(i), 0 * mV)
        Next i

        TheHdw.WAIT LowVddWait * S

        For i = 0 To UBound(LowVddPins)
            ShtPowerV.GetPowerInfo(Vdd3rd, LowVddPins(i)).Force (LowVddPins(i))
        Next i
        
        '========== PatRun ========================================================
        Call PatRun(ReadPat)
    
        '========== PASS/FAIL CHECK ===============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If TheHdw.Digital.FailedPinsCount(site) = 0 And Judge_value(site) = 1 Then
                    Judge_value(site) = 1
                Else
                    Judge_value(site) = 0
                End If
            End If
        Next site
    Else
        '========== PASS/FAIL CHECK ===============================================
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Judge_value(site) = 1
            End If
        Next site
        
    End If

    '========== Result ========================================================
    Call test(Judge_value)

End Function

