Attribute VB_Name = "Otp_PageWithSram_SubMod"
Option Explicit

Public blnFlg_BlowCheck As Boolean

Public Sub OtpVariableClear()
'OTP Initialize Sub.
'OTP����Ŏg�p���Ă���ϐ��̃N���A

Erase AddrParPage
Erase BitParPage
Erase BlowDataAllBin
Erase BlowDataAllBin2
Erase ReadDataAllBin
Erase FFBlowInfo



Dim NowPage As Integer
For NowPage = 0 To OtpPageEnd
    Flg_ModifyPage(NowPage) = False
Next NowPage

End Sub

Public Sub Set_OtpCondition(ByRef ArgArr() As String)

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h

    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_VARIABLE_PARAM) Then
        Err.Raise 9999, "OTP", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
    End If

    '�I�������񂪌�����Ȃ��̂�����
    Dim IsFound As Boolean
    Dim lCount As Long
    Dim i As Long

    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "OTP", """#EOP"" is not found! [" & GetInstanceName & "] !"
    End If

    Dim testConditionList() As String
    testConditionList = Split(ArgArr(0), ",")
    For i = 0 To UBound(testConditionList)
        If Trim(testConditionList(i) <> "") Then
            Call TheCondition.SetCondition(testConditionList(i))
        End If
    Next i

End Sub

Public Sub OtpInitialize_Get_AddressParPage()
'OTP Initialize Sub.
'OTPMAP�V�[�g����ePage��Address����Bit����GET�����B

'���ӎ����i���̊֐��̐��񎖍��B���L�̐���Ɉᔽ���Ă���ꍇ�ɂ͓���ۏ؂��Ȃ��B�j
'����1�FPage�ԍ��́A�ŏIPage�̍ŏIAddress�܂�1�Z�����ԍ����L�����Ă����B�O�Z������Page������Ƃ����āA�L�������Ȃ�������A�Z���𓝍����Ă͂����Ȃ��B


    Dim NowPage As Integer
    Dim RowCount As Long

    Worksheets("OTPMAP").Select                                                                         'OTP��񂪋L�ڂ���Ă���Sheet��I��(OTPMAP)

    For NowPage = 0 To OtpPageSize - 1                                                                  'Page�܂킵

        Do While Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value <> ""          'OTPMAP��Page��擪����Ō���܂ł�LOOP

            If NowPage = Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value Then   'Page����Address����GET
                AddrParPage(NowPage) = AddrParPage(NowPage) + 1
            End If

            RowCount = RowCount + BitParHex
        Loop

        BitParPage(NowPage) = BitParHex * AddrParPage(NowPage)                                          'Page����Bit����GET
        RowCount = 0

    Next NowPage

End Sub

Public Sub OtpInitialize_Get_FixedValue()
'OTP Initialize Sub.
'OTPMAP�V�[�g����Œ�l����GET�BBlow��Verify�Ŏg�p�����B

'���ӎ����i���̊֐��̐��񎖍��B���L�̐���Ɉᔽ���Ă���ꍇ�ɂ͓���ۏ؂��Ȃ��B�j
'����1�FValue(Bin)���͋�NG�B�L�ڂł��镶���͂��ꂾ����"0" or "1" or "X"


    Dim site As Long
    Dim NowPage As Integer
    Dim NowBit As Integer
    Dim RowCount As Long
    Dim TotalRowCount As Long
    
    Worksheets("OTPMAP").Select                                                                         'OTP��񂪋L�ڂ���Ă���Sheet��I��(OTPMAP)
    
    '===== �Œ�l����Bin�ϐ��ɓ��ꍞ�� ======================================
    For site = 0 To nSite
'        If TheExec.Sites.site(site).Active = True Then
            
            For NowPage = 0 To OtpPageSize - 1                                                          'Page�܂킵
                For NowBit = 0 To (BitParPage(NowPage)) - 1                                             'Bit�܂킵
                    
                    ReadDataAllBin(site, NowPage, NowBit) = Cells(NowBit + TotalRowCount + OtpInfoSheet_Row_Value, OtpInfoSheet_Column_Value).Value     'Verify�p�ϐ��ɌŒ�l�����i�[
                    BlowDataAllBin(site, NowPage, NowBit) = ReadDataAllBin(site, NowPage, NowBit)                                                       'Blow�p�ϐ��ɌŒ�l�����i�[
                    BlowDataAllBin2(site, NowPage, NowBit) = ReadDataAllBin(site, NowPage, NowBit)                                                       'Blow�p�ϐ��ɌŒ�l�����i�[
                    If Cells(OtpInfoSheet_Row_BlowInfo + RowCount, OtpInfoSheet_Column_BlowInfo).Value = "SRAM" Then        '���݂�Bit��SRAM�p�ł���΁A����Bit��Blow�p�ϐ���0�̌Œ�l�Ƃ���(�璷��ɏ㏑�������Ȃ��悤��)
                        BlowDataAllBin(site, NowPage, NowBit) = "0"
                        BlowDataAllBin2(site, NowPage, NowBit) = "X"
                    End If
                    
                    RowCount = RowCount + 1
                    
                Next NowBit
                TotalRowCount = RowCount
            Next NowPage
            
            RowCount = 0
            TotalRowCount = 0
            
'        End If
    Next site
    
End Sub

Public Sub OtpInitialize_Get_FFBlowInfo()
'OTP Initialize Sub.
'OTPMAP�V�[�g����FFBlow����Page��Bit����GET�����B

'���ӎ����i���̊֐��̐��񎖍��B���L�̐���Ɉᔽ���Ă���ꍇ�ɂ͓���ۏ؂��Ȃ��B�j
'����1�FFF�������݂̗�ŋ󗓂�NG�BFF�������݂����"1"�BFF�������݂��Ȃ���"0"�B1��0�̂ǂ��炩��K���L�ڂ��Ă��邱�ƁB
'����2�FFF�������݂������y�[�W�ɂ܂�����̂�NG�B����1�y�[�W�����ɂ��ĂˁB����1�y�[�W���ł���΁A��Bit�ł�Blow�ł��܂�(FF(11111111)����Ȃ��Ă��C�P��(EE�Ƃ���OK))�B


    Dim site As Long
    Dim NowPage As Integer
    Dim NowBit As Integer
    Dim RowCount As Long
    
    Worksheets("OTPMAP").Select                                                             'OTP��񂪋L�ڂ���Ă���Sheet��I��(OTPMAP)
    
    '===== �Œ�l����Bin�ϐ��ɓ��ꍞ�� ======================================
    Do While Cells(OtpInfoSheet_Row_FF + RowCount, OtpInfoSheet_Column_FF).Value <> ""      'OTPMAP��FF��񂪋L�ڂ���Ă����̐擪����Ō���܂ł�LOOP

        NowPage = Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value   'LOOP���̌��݂�Page�Ƃ���Bit�ԍ���GET
        NowBit = Cells(OtpInfoSheet_Row_Bit + RowCount, OtpInfoSheet_Column_Bit).Value
        
        If Cells(OtpInfoSheet_Row_FF + RowCount, OtpInfoSheet_Column_FF).Value = 1 Then     'FFBlow�p�̕ϐ��쐬�B���ƁAPage����GET�B
            FFBlowInfo(NowPage, NowBit) = 1
            FFBlowPage = NowPage
        Else
            FFBlowInfo(NowPage, NowBit) = 0
        End If
        
        RowCount = RowCount + 1
        
    Loop
                
End Sub

Public Sub OtpInitialize_Select_OtpBlow_Page()
'OTP Initialize Sub.
'OTPMAP�V�[�g����OTPBLOW���K�v�ƂȂ�y�[�W�����擾�B���̃y�[�W�ł�BlowValue���S��"0"�ł���΁A���̃y�[�W��Blow��PatRun���s��Ȃ��B

'���ӎ����i���̊֐��̐��񎖍��B���L�̐���Ɉᔽ���Ă���ꍇ�ɂ͓���ۏ؂��Ȃ��B�j
'����1�FValue(Bin)���͋�NG�B�L�ڂł��镶���͂��ꂾ����"0" or "1" or "X"

    Dim NowPage As Integer
    Dim RowCount As Long
    
    Worksheets("OTPMAP").Select                                                                     'OTP��񂪋L�ڂ���Ă���Sheet��I��(OTPMAP)
    
    For NowPage = 0 To OtpPageSize - 1
        Flg_OtpBlowPage(NowPage) = False                                                            '�܂��͑S�y�[�WFalse�ɂ��Ă����B�S�y�[�WBlow���s�����B
    Next NowPage
        
    Do While Cells(OtpInfoSheet_Row_Value + RowCount, OtpInfoSheet_Column_Value).Value <> ""        '�ePage�̊eBit��LOOP
        If Cells(OtpInfoSheet_Row_Value + RowCount, OtpInfoSheet_Column_Value).Value <> "0" Then
            NowPage = Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value       '���݂�Page��"0"�ȊO��BlowValue�����݂���΁ATrue�ɂ���Blow���s�t���O�𗧂Ă�
            Flg_OtpBlowPage(NowPage) = True
            
            If Cells(OtpInfoSheet_Row_Value + RowCount, OtpInfoSheet_Column_Value).Value = 1 Then   '���݂�Page��"1"��BlowValue(=�Œ�l)�����݂���΂���Page����ێ��B�Ώ�Page����������΁A���̒��̍Ō��Page����ێ��B
                Flg_OtpBlowFixValPage = Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value
            End If
            
        End If
        RowCount = RowCount + 1
    Loop
        
End Sub

Private Function Dec2Bin(myDecvalue As String, OutBit As Integer) As String
'OTP Standard Function.
'10�i����2�i���ɕϊ�����

    Dim lngdecnumber As Long
    Dim strbinnumber As String
    strbinnumber = ""
    lngdecnumber = 0

    lngdecnumber = CLng(myDecvalue)

    Do
        strbinnumber = CStr(lngdecnumber Mod 2) & strbinnumber
        lngdecnumber = Fix(lngdecnumber / 2)
    Loop While lngdecnumber > 0

    Do While Len(strbinnumber) < OutBit
        strbinnumber = "0" & strbinnumber
    Loop

    Dec2Bin = strbinnumber

End Function

Public Sub OtpInitialize_Make_FixedValuePattern()
'OTP Initialize Sub.
'�Œ�l���͍ŏ���Modify���Ă����B


    Dim site As Long
    Dim NowPage As Integer
    Dim NowBit As Integer

    Dim PageBit As Long
    Dim BinData As String
    Const PageInfoBit = 8
    Const ByteBit = 8
    Dim PG_ARY(PageInfoBit - 1) As String
    Dim ByteCount As Long
    Dim LenLen As Long
    Dim iiii As Long
    Dim ModifyDataBlowDeb As String
    Dim ModifyDataVerifyDeb As String
    Dim ModifyDataFFDeb As String
    Dim i As Integer

    If Flg_Simulator = 1 Then Exit Sub
    
    'Active Site PickUp (1site de OK)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Exit For
        End If
    Next site
            
    Call StopPattern
     
    For NowPage = 0 To OtpPageEnd
    
        '��������������������������������������������������������������������������
        '��Page����Modify-> Blow�p�^�[���AVerify�p�^�[���A�Œ�l�p�^�[���ABlank�p�^�[���AFFBlow�p�^�[��
        '��������������������������������������������������������������������������
    
        BinData = Dec2Bin(CStr(NowPage), PageInfoBit)
        
        For PageBit = 0 To PageInfoBit - 1
            PG_ARY(PageBit) = Mid(BinData, PageBit + 1, 1)
        Next PageBit
    
        TheHdw.Digital.Patterns.pat("OtpBlowPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_Page_OtpBlow & CStr(NowPage), 0, RejiIn, PG_ARY
        TheHdw.Digital.Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_Page_OtpVerify & CStr(NowPage), 0, RejiOut, PG_ARY
        TheHdw.Digital.Patterns.pat("OtpFixedValueCheckPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_Page_OtpFixedValueCheck & CStr(NowPage), 0, RejiOut, PG_ARY
        TheHdw.Digital.Patterns.pat("BlankCheckPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_Page_BlankCheck & CStr(NowPage), 0, RejiOut, PG_ARY
        If FFBlowPage = NowPage Then    'FFBlow����y�[�W����Modify(FFBlow����y�[�W��1Page�ł��邱�ƑO��)
            TheHdw.Digital.Patterns.pat("OtpBlow_Break_Pat").ModifyPinVectorBlockData Label_Page_OtpBlow_Break, 0, RejiIn, PG_ARY
        End If
    
    
        '��������������������������������������������������������������������������
        '���Œ�lBlow����Blow�p�^�[����Modify
        '��������������������������������������������������������������������������
        ByteCount = 0
        ModifyDataBlowDeb = ""
        '===== Make Blow Modify Data ==================================
        For NowBit = 0 To BitParPage(NowPage) - 1
            ByteCount = ByteCount + 1
            If ByteCount < ByteBit Then
                ModifyDataBlowDeb = ModifyDataBlowDeb & BlowDataAllBin(site, NowPage, NowBit)
            ElseIf ByteCount = ByteBit Then
                ModifyDataBlowDeb = ModifyDataBlowDeb & BlowDataAllBin(site, NowPage, NowBit) & "X"
            End If
            If ByteCount = 8 Then ByteCount = 0
        Next NowBit
        
        LenLen = Len(ModifyDataBlowDeb)
        ReDim ModifyDataBlow(LenLen - 1) As String
        For iiii = 1 To LenLen
            ModifyDataBlow(iiii - 1) = Mid(ModifyDataBlowDeb, iiii, 1)
        Next iiii
        
        '===== Modify Blow Pattern (Fixed Value) ======================
        TheHdw.Digital.Patterns.pat("OtpBlowPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_OtpBlow & CStr(NowPage), 0, RejiIn, ModifyDataBlow
    
    
        '��������������������������������������������������������������������������
        '��AutoBlow����Blow�p�^�[����Modify
        '��������������������������������������������������������������������������

        Dim ModifyDataBlowAuto(ByteBit - 1) As String
        For iiii = 1 To ByteBit
            ModifyDataBlowAuto(iiii - 1) = "0"
        Next iiii
        
        '===== Modify Blow Pattern (Fixed Value) ======================
        TheHdw.Digital.Patterns.pat("OtpBlowPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_OtpBlowAuto & CStr(NowPage), 0, RejiIn, ModifyDataBlowAuto

    
        '��������������������������������������������������������������������������
        '���Œ�lBlow����Verify�p�^�[����Modify
        '��������������������������������������������������������������������������
        ByteCount = 0
        ModifyDataVerifyDeb = ""
        '===== Make Verify Modify Data ================================
        For NowBit = 0 To BitParPage(NowPage) - 1

            ByteCount = ByteCount + 1

            If ByteCount < ByteBit Then
                If BlowDataAllBin(site, NowPage, NowBit) = "0" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "L"
                ElseIf BlowDataAllBin(site, NowPage, NowBit) = "1" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "H"
                ElseIf BlowDataAllBin(site, NowPage, NowBit) = "X" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "X"
                End If
            ElseIf ByteCount = ByteBit Then
                If BlowDataAllBin(site, NowPage, NowBit) = "0" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "LX"
                ElseIf BlowDataAllBin(site, NowPage, NowBit) = "1" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "HX"
                ElseIf BlowDataAllBin(site, NowPage, NowBit) = "X" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "XX"
                End If

                ByteCount = 0
            End If

        Next NowBit

        LenLen = Len(ModifyDataVerifyDeb)
        ReDim ModifyDataVerifyFix(LenLen - 1) As String
        For iiii = 1 To LenLen
            ModifyDataVerifyFix(iiii - 1) = Mid(ModifyDataVerifyDeb, iiii, 1)
        Next iiii

        '===== Modify Verify Pattern (Fixed Value) ====================
        TheHdw.Digital.Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_OtpVerify & CStr(NowPage), 0, RejiOut, ModifyDataVerifyFix


        '��������������������������������������������������������������������������
        '���Œ�lBlow�����Œ�l�p�^�[����Modify
        '��������������������������������������������������������������������������
        ByteCount = 0
        ModifyDataVerifyDeb = ""
        '===== Make Verify Modify Data ================================
        For NowBit = 0 To BitParPage(NowPage) - 1

            ByteCount = ByteCount + 1

            If ByteCount < ByteBit Then
                If BlowDataAllBin2(site, NowPage, NowBit) = "0" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "L"
                ElseIf BlowDataAllBin2(site, NowPage, NowBit) = "1" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "H"
                ElseIf BlowDataAllBin2(site, NowPage, NowBit) = "X" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "X"
                End If
            ElseIf ByteCount = ByteBit Then
                If BlowDataAllBin2(site, NowPage, NowBit) = "0" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "LX"
                ElseIf BlowDataAllBin2(site, NowPage, NowBit) = "1" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "HX"
                ElseIf BlowDataAllBin2(site, NowPage, NowBit) = "X" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "XX"
                End If

                ByteCount = 0
            End If

        Next NowBit

        LenLen = Len(ModifyDataVerifyDeb)
        ReDim ModifyDataVerifyFix2(LenLen - 1) As String
        For iiii = 1 To LenLen
            ModifyDataVerifyFix2(iiii - 1) = Mid(ModifyDataVerifyDeb, iiii, 1)
        Next iiii

        '===== Modify Verify Pattern (Fixed Value) ====================
        TheHdw.Digital.Patterns.pat("OtpFixedValueCheckPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_OtpFixedValueCheck & CStr(NowPage), 0, RejiOut, ModifyDataVerifyFix2


        '��������������������������������������������������������������������������
        '��"FF"Blow����FF��pBlow�p�^�[����Modify
        '��������������������������������������������������������������������������
        ByteCount = 0
        ModifyDataFFDeb = ""
        If FFBlowPage = NowPage Then    'FFBlow����y�[�W����Modify(FFBlow����y�[�W��1Page�ł��邱�ƑO��)

            '===== Make Blow Modify Data ==================================
            For NowBit = 0 To BitParPage(NowPage) - 1
                
                ByteCount = ByteCount + 1
                
                If ByteCount < ByteBit Then
                    ModifyDataFFDeb = ModifyDataFFDeb & FFBlowInfo(NowPage, NowBit)
                ElseIf ByteCount = ByteBit Then
                    ModifyDataFFDeb = ModifyDataFFDeb & FFBlowInfo(NowPage, NowBit) & "X"
                End If
                If ByteCount = 8 Then ByteCount = 0
            Next NowBit

            LenLen = Len(ModifyDataFFDeb)
            ReDim ModifyDataFF(LenLen - 1) As String
            For iiii = 1 To LenLen
                ModifyDataFF(iiii - 1) = Mid(ModifyDataFFDeb, iiii, 1)
            Next iiii

            '===== Modify Blow Pattern (Fixed Value) ======================
            TheHdw.Digital.Patterns.pat("OtpBlow_Break_Pat").ModifyPinVectorBlockData Label_OtpBlow_Break, 0, RejiIn, ModifyDataFF


            '��������������������������������������������������������������������������
            '��AutoBlow����FF��pBlow�p�^�[����Modify
            '��������������������������������������������������������������������������
    
            Dim ModifyDataFFBlowAuto(ByteBit - 1) As String
            For iiii = 1 To ByteBit
                ModifyDataFFBlowAuto(iiii - 1) = "0"
            Next iiii
            
            '===== Modify Blow Pattern (Fixed Value) ======================
            TheHdw.Digital.Patterns.pat("OtpBlow_Break_Pat").ModifyPinVectorBlockData Label_OtpBlowAuto_Break, 0, RejiIn, ModifyDataFFBlowAuto

        End If


    Next NowPage

End Sub

Private Function MakeModifyData(ByRef InBinData() As String, ByRef OutBinData() As String, ByRef PageSelect() As Long, ByRef BitSelect() As Long, ByVal site As Long, Optional ByVal Case_SramRep As Boolean = False) As String
'OTP�̌X��Blow�f�[�^���AModify�p�̂܂Ƃ߂��ϐ��֓���Ȃ���

    Dim strSteps As Long
    Dim width As Long
    
    width = Len(InBinData(site))

    For strSteps = 0 To width - 1
    
        If Mid(InBinData(site), strSteps + 1, 1) = "0" Then OutBinData(site, PageSelect(strSteps), BitSelect(strSteps)) = "0"
        If Mid(InBinData(site), strSteps + 1, 1) = "1" Then OutBinData(site, PageSelect(strSteps), BitSelect(strSteps)) = "1"
               
        If Case_SramRep = False Then                       'SRAM�璷�̏ꍇ�͂��̃t���O�͗��ĂȂ��BSRAM�璷��p�t���O�͕ʂɕێ��B
            Flg_ModifyPage(PageSelect(strSteps)) = True    '����Page��Blow���K�v�Ƃ������ƂŁABlowPage�t���O��True�ɂ����B�K�v�Œ����Blow���s�ł����悤�ɂˁB
        End If
        
    Next

End Function

Public Sub OtpInitialize_Get_PageBit(ByVal Label As String, ByVal BitWidthAll As Long, ByRef Page() As Long, ByRef Bit() As Long)
'OTP Initialize Sub.
'OTP�̊e�ϓ��l����Page��Bit����GET�����B

'Arg1 Input:  ���x��(�ϓ��l��񖈂Ɏ����Ă������)�B���̃��x����OTPMAP�V�[�g���Ō������āA������������Page��Bit�ԍ���ϐ��֊i�[���ĕۑ����Ă�����B
'Arg2 Input:  Bit���B���x���Ɋ��蓖�Ă��Ă���OTP������Bit���̂��ƁB
'Arg3 Output: Page���B���x���Ɋ��蓖�Ă��Ă���OTP�������̂���Bit���A�ǂ�Page�ɓ�����̂��������ϐ��B
'Arg4 Output: Bit���B���x���Ɋ��蓖�Ă��Ă���OTP�������̂���Bit���A�ǂ�Bit�ɓ�����̂��������ϐ��B

'���ӎ����i���̊֐��̐��񎖍��B���L�̐���Ɉᔽ���Ă���ꍇ�ɂ͓���ۏ؂��Ȃ��B�j


    Dim RowCount As Long
    Dim NowPage As Integer
    Dim ii As Long
    
    Worksheets("OTPMAP").Select                                                                             'OTP��񂪋L�ڂ���Ă���Sheet��I��(OTPMAP)
        
    '========== Get Page&Address Start Infomation =============================
    Do While Cells(OtpInfoSheet_Row_BlowInfo + RowCount, OtpInfoSheet_Column_BlowInfo).Value <> Label       'OTP���̐擪����A���x���܂ł̍s�����J�E���g
        RowCount = RowCount + 1
        If Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value = "" Then Exit Do        '����LOOP�h�~
    Loop
    
    Do While ii <> BitWidthAll
        If Cells(OtpInfoSheet_Row_BlowInfo + RowCount, OtpInfoSheet_Column_BlowInfo).Value = Label Then     '�J�E���g�����s������A���̃��x��������Bit������Page����Bit�ԍ�����GET
            Page(ii) = Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value
            Bit(ii) = Cells(OtpInfoSheet_Row_Bit + RowCount, OtpInfoSheet_Column_Bit).Value
            ii = ii + 1
        End If
        RowCount = RowCount + 1
        If Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value = "" Then Exit Do        '����LOOP�h�~
    Loop
    
End Sub

Public Function ActiveSite_Check_OTP() As Long
'OTP Standard Function.
'Active��Site�Ƀt���O�𗧂Ă�B���̌�AActiveSite���ꎞ�I��Disable�ɂ��鎞�Ȃǂɗp����B
    
    Erase Flg_ActiveSite_OTP
    Dim site As Long

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Flg_ActiveSite_OTP(site) = 1
        End If
    Next site

End Function

Public Function ActiveSite_Return_OTP() As Long
'OTP Standard Function.
'ActiveSite_Check�Ńt���O�𗧂ĂĂ�����Site��Active�ɂ���B����ȑO�Ɉꎞ�I��Disable�ɂ��Ă���Site��Active�ɖ߂����Ȃǂɗp����B

    Dim site As Long

    For site = 0 To nSite
        If Flg_ActiveSite_OTP(site) = 1 Then
            TheExec.sites.site(site).Active = True
        End If
    Next site

End Function

Public Sub Output_OtpBlowData()

'OTPBLOW�̃f�o�b�O���O�f���o���֐�
'�f�[�^���O��ۑ���AExcel�ɓ\��t���āA�w:�x�ŋ�؂�΂��ꂢ�ɕ��т܂��B
'�Œ�l�́AOTPMAP�Ƃ��̂܂ܔ�r����΂�����B

Dim site As Long
Dim NowBit As Long
Dim NowPage As Long
Dim PageA As Long
Dim Data As String
Dim PageData As String
Dim first As Boolean
first = True
    
    TheExec.Datalog.WriteComment "OTP BLOW DATA"
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            PageA = 0
            For NowBit = 0 To BitParPage(PageA) - 1

                For NowPage = 0 To OtpPageSize - 1
                    If first = True Then
                        PageData = PageData + " : " & "P" & NowPage
                    End If
                    Data = Data + " :  " & BlowDataAllBin(site, NowPage, NowBit)
                Next NowPage
                
                If first = True Then
                    TheExec.Datalog.WriteComment "---- - : --- - : ----" & PageData
                End If
                
                first = False
                TheExec.Datalog.WriteComment "Site " & site & " : " & "Bit " & NowBit & " : " & "Data" & Data
                Data = ""
                
                If PageA < OtpPageSize Then
                    PageA = PageA + 1
                End If

            Next NowBit
            
            PageData = ""
            first = True
        End If
    Next site
                    

End Sub

Public Sub Output_OtpReadData()

Dim site As Long
Dim NowBit As Long
Dim NowPage As Long
Dim PageA As Long
Dim Data As String
Dim PageData As String
Dim first As Boolean
first = True


    Dim site_status As Long
    Dim NowCapData As Long
    Dim VectorOffset As Long
    Dim HramLoop As Long
    Dim HramSizeMod As Long
    Dim NowHramLoop As Long
    Dim HramSetSize As Long
    Const HramSize As Integer = 256
    Const BitParByte As Long = 8
    Dim DataOffset As Long
    Dim ReadErr(nSite) As Double
    Dim Deb_ReadDataAllBin(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String
    


    '��������������������������������������������������������������������������
    '��RollCall���s
    '��������������������������������������������������������������������������

    If TheExec.sites.ActiveCount > 0 Then
        
        For NowPage = 0 To OtpPageEnd
            
            '========== PATTERN LOAD ==========================================
            With TheHdw.Digital
                .Timing.Load
                .Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").Load
            End With
        
            DataOffset = 0
            VectorOffset = 0
            HramLoop = Int(BitParPage(NowPage) / HramSize) + 1
            HramSizeMod = BitParPage(NowPage) Mod HramSize
            If HramSizeMod = 0 Then
                HramLoop = HramLoop - 1
                HramSizeMod = HramSize
            End If
            
            For NowHramLoop = 0 To (HramLoop - 1)
                
                '========== HRAM SETUP OFFSET CALCULATE ===========================
                If NowHramLoop <> (HramLoop - 1) Then
                    VectorOffset = NowHramLoop * (HramSize / BitParByte * ByteParVector_VerifyPat)
                    HramSetSize = HramSize
                Else
                    VectorOffset = NowHramLoop * (HramSize / BitParByte * ByteParVector_VerifyPat)
                    HramSetSize = HramSizeMod
                End If
                
                '========== HRAM SETUP ============================================
                Call StopPattern
                With TheHdw.Digital
                    .Patgen.EventCycleEnabled = False
                    .Patgen.NoHaltMode = noHaltAlways                                                           ' noHaltAlways:�p�^�[���̒��ɋL�q���Ă���Halt�A����VBA���Halt�����s�����܂Ńp�^�[���͎~�܂�Ȃ�
                    .Patgen.EventSetVector True, "OtpVerifyPage" & CStr(NowPage) & "_Pat", Label_OtpVerify & CStr(NowPage), VectorOffset
                    .HRAM.SetTrigger trigFirst, True, 0, True                                                           ' trigFail:�ŏ���Fail�T�C�N�������荞�݊J�n    True:EventCycleCount��L��    0:��荞�݊J�n�T�C�N���̉��T�C�N���O�����荞�ނ�   True:��荞��ł���T�C�N������HRAM�T�C�Y�ɒB�����ꍇ�ɂ͂����Ŏ�荞�݂���߂�
                    .HRAM.SetCapture captSTV, True                                                               ' captFail:Fail�T�C�N���݂̂���荞��   Ture: ���s�[�g����Vector�ł���ꍇ�͍Ō�̃��s�[�g�T�C�N��������荞��
                    .HRAM.Size = HramSetSize                                                                       ' �Ƃ肠�����ő�l���w��
                End With
                
                '========== PATTERN RUN ===========================================
                TheHdw.Digital.Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").Run ("Start")

                '========== HRAM DATA SITE ============================
                site_status = TheExec.sites.SelectFirst
                While site_status <> loopDone
                    site = TheExec.sites.SelectedSite
                    
                    For NowCapData = 0 To (HramSize - 1)
                        If TheHdw.Digital.HRAM.Pins(RejiOut).PinData(NowCapData) = "L" Then
                            Deb_ReadDataAllBin(site, NowPage, NowCapData + DataOffset) = "0"
                        ElseIf TheHdw.Digital.HRAM.Pins(RejiOut).PinData(NowCapData) = "H" Then
                            Deb_ReadDataAllBin(site, NowPage, NowCapData + DataOffset) = "1"
                        End If
                    Next NowCapData
                    
                    site_status = TheExec.sites.SelectNext(site_status)
                Wend
                
                DataOffset = DataOffset + HramSize

            Next NowHramLoop
            
        Next NowPage
        
        '========== PATTERN SETUP =================================
        With TheHdw.Digital.Patgen
            .EventCycleEnabled = False
            .EventCycleCount = 1
            .MaskTilCycle = False
        End With

    End If



    TheExec.Datalog.WriteComment "OTP READ DATA"
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            PageA = 0
            For NowBit = 0 To BitParPage(PageA) - 1

                For NowPage = 0 To OtpPageSize - 1
                    If first = True Then
                        PageData = PageData + " : " & "P" & NowPage
                    End If
                    Data = Data + " :  " & Deb_ReadDataAllBin(site, NowPage, NowBit)
                Next NowPage
                
                If first = True Then
                    TheExec.Datalog.WriteComment "---- - : --- - : ----" & PageData
                End If
                
                first = False
                TheExec.Datalog.WriteComment "Site " & site & " : " & "Bit " & NowBit & " : " & "Data" & Data
                Data = ""
                
                If PageA < OtpPageSize Then
                    PageA = PageA + 1
                End If

            Next NowBit
            
            PageData = ""
            first = True
        End If
    Next site
                               

End Sub

Private Sub Get_ChipId_Debug()
'OTP Debug Sub. GET LOTNo/WaferNo/ChipNo Dummy Data.

    Dim TypeName_OTP_Debug As String
    TypeName_OTP_Debug = Mid(NormalJobName, 4, 3)
            
    LotName = "98M" & TypeName_OTP_Debug & "76543"
    WaferNo = "4"
    
    Dim site As Long
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            DeviceNumber_site(site) = site + 1
        End If
    Next site

End Sub

Public Sub MakeBlowData_Lot(ByVal BitWidth As Integer, ByRef Page() As Long, ByRef Bit() As Long, ByRef Beam As Long)

    Dim site As Long
    Dim BlowData(nSite) As String

    If Flg_OTP_BLOW = 0 Then Exit Sub
    
    '========== For Otp Blow Debug ============================================
    If Flg_AutoMode = False And Flg_OTP_BLOW = 1 Then Call Get_ChipId_Debug
    
    '========== Dec -> Bin ============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then

            BlowData(site) = "XXXX"

            If Mid(LotName, Beam, 1) = "0" Then BlowData(site) = "0000"
            If Mid(LotName, Beam, 1) = "1" Then BlowData(site) = "0001"
            If Mid(LotName, Beam, 1) = "2" Then BlowData(site) = "0010"
            If Mid(LotName, Beam, 1) = "3" Then BlowData(site) = "0011"
            If Mid(LotName, Beam, 1) = "4" Then BlowData(site) = "0100"
            If Mid(LotName, Beam, 1) = "5" Then BlowData(site) = "0101"
            If Mid(LotName, Beam, 1) = "6" Then BlowData(site) = "0110"
            If Mid(LotName, Beam, 1) = "7" Then BlowData(site) = "0111"
            If Mid(LotName, Beam, 1) = "8" Then BlowData(site) = "1000"
            If Mid(LotName, Beam, 1) = "9" Then BlowData(site) = "1001"
            If Mid(LotName, Beam, 1) = "A" Then BlowData(site) = "1010"
            If Mid(LotName, Beam, 1) = "B" Then BlowData(site) = "1011"
            If Mid(LotName, Beam, 1) = "C" Then BlowData(site) = "1100"
            If Mid(LotName, Beam, 1) = "D" Then BlowData(site) = "1101"
            If Mid(LotName, Beam, 1) = "E" Then BlowData(site) = "1110"
            If Mid(LotName, Beam, 1) = "F" Then BlowData(site) = "1111"

            If Mid(LotName, Beam, 1) = "S" Then BlowData(site) = "1000"
            If Mid(LotName, Beam, 1) = "M" Then BlowData(site) = "0100"

            If BlowData(site) = "XXXX" Then
                blnFlg_BlowCheck = True
            End If

            '========== Blow�p�ϐ��쐬 ========================================
            Call MakeModifyData(BlowData, BlowDataAllBin, Page, Bit, site)

        End If
    Next site
    
    '### DEBUG LOG OUTPUT ###
    If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site**]" & " " & "LotNo" & Beam & " = " & Mid(LotName, Beam, 1)
    
End Sub

Public Sub MakeBlowData_Wafer(ByVal BitWidth As Integer, ByRef Page() As Long, ByRef Bit() As Long)

    Dim site As Long
    Dim BlowData(nSite) As String

    If Flg_OTP_BLOW = 0 Then Exit Sub

    '========== For Otp Blow Debug ============================================
    If Flg_AutoMode = False And Flg_OTP_BLOW = 1 Then Call Get_ChipId_Debug
    
    '========== Dec -> Bin ============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
                        
            BlowData(site) = Dec2Bin(CStr(WaferNo), BitWidth)
            
            '========== Blow�p�ϐ��쐬 ========================================
            Call MakeModifyData(BlowData, BlowDataAllBin, Page, Bit, site)

        End If
    Next site
    
    '### DEBUG LOG OUTPUT ###
    If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site**]" & " " & "WaferNo" & " = " & CStr(WaferNo)

End Sub

Public Sub MakeBlowData_Chip(ByVal BitWidth As Integer, ByRef Page() As Long, ByRef Bit() As Long)

    Dim site As Long
    Dim BlowData(nSite) As String

    If Flg_OTP_BLOW = 0 Then Exit Sub

    '========== For Otp Blow Debug ============================================
    If Flg_AutoMode = False And Flg_OTP_BLOW = 1 Then Call Get_ChipId_Debug
    
    '========== Dec -> Bin ============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
                
            BlowData(site) = Dec2Bin(CStr(DeviceNumber_site(site)), BitWidth)

            '========== Blow�p�ϐ��쐬 ========================================
            Call MakeModifyData(BlowData, BlowDataAllBin, Page, Bit, site)

            '### DEBUG LOG OUTPUT ###
            If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "ChipNo" & " = " & CStr(DeviceNumber_site(site))

        End If
    Next site
    
    
End Sub

Public Sub MakeBlowData_Temp(ByVal BitWidth_Ofs As Integer, ByVal BitWidth_Slp As Integer, ByRef Page() As Long, ByRef Bit() As Long, ByVal TEMP_OFS As String, ByVal TEMP_SLP As String)
                                
'OTP���ו␳Blow�f�[�^�����pSub�iTEMP�n�␳�j

'Arg1 Input:���x�v�I�t�Z�b�g����Bit��
'Arg2 Input:���x�v�X������Bit��
'Arg3 Input:�eBlow���Bit�̑Ώ�Page�ԍ�
'Arg4 Input:�eBlow���Bit�̑Ώ�Bit�ԍ�
'Arg5 Input:���x�v�I�t�Z�b�g���̃��x�����i���ږ��j
'Arg6 Input:���x�v�X�����̃��x�����i���ږ��j

'���ӎ����i���̊֐��̐��񎖍��B���L�̐���Ɉᔽ���Ă���ꍇ�ɂ͓���ۏ؂��Ȃ��B�j
'����1:���x�v�̃I�t�Z�b�g���ƌX�����͘A��Bit�ł��邱�ƁB����ɁA�I�t�Z�b�g���̕����Ⴂ�A�h���X�ł��邱�ƁB


    Dim site As Long
    Dim TempInfo_Ofs() As Double
    Dim TempInfo_Slp() As Double
    Dim AllTempInfo(nSite) As String
    Dim data_ofs(nSite) As String
    Dim data_slp(nSite) As String

    '========== GET TEMP INFOMATION =======================
    TheResult.GetResult TEMP_OFS, TempInfo_Ofs
    TheResult.GetResult TEMP_SLP, TempInfo_Slp
    
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
        
            '========== Dec->Bin & Offset + Slope =========================
            data_ofs(site) = Dec2Bin(CStr(TempInfo_Ofs(site)), BitWidth_Ofs)
            data_slp(site) = Dec2Bin(CStr(TempInfo_Slp(site)), BitWidth_Slp)
            
            AllTempInfo(site) = data_ofs(site) & data_slp(site)
        
            '========== Blow�p�ϐ��쐬 ====================================
            Call MakeModifyData(AllTempInfo, BlowDataAllBin, Page, Bit, site)


            '##### DEBUG LOG OUTPUT #####
            If Flg_Debug = 1 Then
                TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "TEMP:" & "OFS" & " = " & TempInfo_Ofs(site)
                TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "TEMP:" & "SLP" & " = " & TempInfo_Slp(site)
            End If

        End If
    Next site
                        
End Sub

Public Sub MakeBlowData_Sram(tsite As Long, ByVal BitWidth_En As Integer, ByVal BitWidth_Addr As Integer, ByVal BitWidth_Data As Integer, ByRef Page() As Long, ByRef Bit() As Long)
                                
'OTP���ו␳Blow�f�[�^�����pSub�iSRAM�璷�j

'Arg1 Input:SRAM�璷��Bit���BFuse�璷�▢�g�pBit�͊܂߂���ʖڂł�
'Arg2 Input:�eBlow���Bit�̑Ώ�Page�ԍ�
'Arg3 Input:�eBlow���Bit�̑Ώ�Bit�ԍ�

'���ӎ����i���̊֐��̐��񎖍��B���L�̐���Ɉᔽ���Ă���ꍇ�ɂ͓���ۏ؂��Ȃ��B�j
'����1:EN -> Address -> Data �̕��т�1�璷�f�[�^���\������Ă���d�l�ł��邱�ƁB


    Dim RepairNo As Integer
    Dim BlowData_SramRep(nSite) As String
    Dim i As Long
    
    '===== SRAM�璷�f�[�^���Ȃ��� ===========================================
    For RepairNo = 1 To MAX_EF_BIST_RD_BIT
        BlowData_SramRep(tsite) = BlowData_SramRep(tsite) + Dec2Bin(CStr(Ef_Enbl_Addr(tsite, RepairNo - 1)), BitWidth_En)
        BlowData_SramRep(tsite) = BlowData_SramRep(tsite) + Dec2Bin(CStr(Ef_Rcon_Addr(tsite, RepairNo - 1)), BitWidth_Addr)
        BlowData_SramRep(tsite) = BlowData_SramRep(tsite) + Dec2Bin(CStr(Ef_Repr_Data(tsite, RepairNo - 1)), BitWidth_Data)
    Next RepairNo

    '========== Blow�p�ϐ��쐬 ====================================
    Call MakeModifyData(BlowData_SramRep, SramBlowDataAllBin, Page, Bit, tsite, True)
    
    '========== SRAM�璷��Modify���s��Page��I�� ==================
    For i = 0 To Len(BlowData_SramRep(tsite)) - 1
        Flg_ModifyPageSRAM(Page(i)) = True
    Next i
                        
End Sub

Public Sub MakeBlowData_Defect_SinCpFd(ByVal MaxRepairNo As Integer, ByVal BitWidthN As Integer, ByVal BitWidthX As Integer, ByVal BitWidthY As Integer, ByVal BitWidthS As Integer, ByVal BitWidthD As Integer, _
                                       ByVal SourceType As String, ByVal NgAdd_LeftS As Long, ByVal NgAdd_LeftE As Long, ByVal NgAdd_RightS As Long, ByVal NgAdd_RightE As Long, _
                                       ByRef Page() As Long, ByRef Bit() As Long, ByRef OverFlowCheck() As Double, ByRef NgAddressCheck() As Double, ParamArray DefectInfo() As Variant)
                                
'OTP���ו␳Blow�f�[�^�����pSub�iSingle/Couplet/FD�␳�̈悪���ʂ̃^�C�v�j

'Arg1  Input:�ő�\�␳��
'Arg2  Input:������bit��
'Arg3  Input:X�A�h���X����bit��(1���ד������)
'Arg4  Input:Y�A�h���X����bit��(1���ד������)
'Arg5  Input:Sorce����bit��(1���ד������)
'Arg6  Input:Direction����bit��(1���ד������)
'Arg7  Input:Sorce���̃^�C�v
'Arg9  Input:���ׂ����݂��Ă͂����Ȃ�ZONE�̍����X�^�[�g�A�h���X
'Arg10 Input:���ׂ����݂��Ă͂����Ȃ�ZONE�̍����G���h�A�h���X
'Arg11 Input:���ׂ����݂��Ă͂����Ȃ�ZONE�̉E���X�^�[�g�A�h���X
'Arg12 Input:���ׂ����݂��Ă͂����Ȃ�ZONE�̉E���G���h�A�h���X
'Arg13 Input:�eBlow���Bit�̑Ώ�Page�ԍ�
'Arg14 Input:�eBlow���Bit�̑Ώ�Bit�ԍ�
'Arg15 Output:�␳����I�[�o�[�����m�点����t���O
'Arg16 Output:���ׂ����݂��Ă͂����Ȃ�ZONE�̌��חL�������m�点����t���O
'Arg17�ȍ~ Input:�B����������n����錇�׏��BDK��HL�Ƃ��������׎�ސ��ɉ����āAArg18��Arg19�ƈ�����������B


'���ӎ����i���̊֐��̐��񎖍��B���L�̐���Ɉᔽ���Ă���ꍇ�ɂ͓���ۏ؂��Ȃ��B�j
'����1�FBlow�f�[�^�S�̂̂Ȃ���́i���ׂ�3�������ꍇ�̗�j�A���ˌ��ׇ@SourceSorce�ˌ��ׇ@Direction�ˌ��ׇ@X�A�h���X�ˌ��ׇ@Y�A�h���X�ˌ��ׇASourceSorce�ˌ��ׇADirection�ˌ��ׇAX�A�h���X�ˌ��ׇAY�A�h���X�ˌ��ׇBSourceSorce�ˌ��ׇBDirection�ˌ��ׇBX�A�h���X�ˌ��ׇBY�A�h���X�ˁ@�Ƃ����d�l�ł��邱�ƁB
'����2�FOTP���̌��ו␳��Blowbit�A�h���X�������Blow����ŏIDirection���܂őS�ĘA�����Ă��邱�ƁB���ˌ��ׇ@SourceSorce��CP���ׇ@Direction�ˌ��ׇ@X�A�h���X��CP���ׇ@Y�A�h���X�ˁ@LOT���@�ˌ��ׇASourceSorce�ˌ��ׇADirection�ˌ��ׇAX�A�h���X��C�ׇAY�A�h���X�@�݂����Ȏd�l�̓_���B
    
    Dim site As Long
    
    Dim BlowData_All(nSite) As String
    Dim amari(nSite) As Integer
    Dim ParamLoop As Long
    Dim ParamNo As Long
    ParamNo = UBound(DefectInfo)
    Dim ParamLabel As String
    Dim NowRepairNo As Integer
    Dim i(nSite) As Long
    Dim SameAddress As Boolean
    Dim Go_NgAddCheck As Boolean
    Dim NgLeft As Long
    Dim NgRight As Long
    
    Dim DefInfo_Num() As Double
    Dim DefInfo_Hadd() As Double
    Dim DefInfo_Vadd() As Double
    Dim DefInfo_Dire() As Double
    Dim DefInfo_Src1() As Double
    Dim DefInfo_Src2 As Double

    Dim AllDefInfo_Num(nSite) As Double
    ReDim AllDefInfo_Hadd(nSite, MaxRepairNo - 1) As Double
    ReDim AllDefInfo_Vadd(nSite, MaxRepairNo - 1) As Double
    ReDim AllDefInfo_Dire(nSite, MaxRepairNo - 1) As Double
    ReDim AllDefInfo_Src(nSite, MaxRepairNo - 1) As Double
    
    For ParamLoop = 0 To ParamNo                                                                            '���ׂ̎�ސ���LOOP
        ParamLabel = CStr(DefectInfo(ParamLoop))                                                            '���׍��ږ���GET
        
        '========== GET DEFECT INFOMATION ===============================                                   '���݂̌��׍��ڂ�Defect����GET
        TheResult.GetResult ParamLabel & "_Info_Num", DefInfo_Num
        TheResult.GetResult ParamLabel & "_Info_Hadd", DefInfo_Hadd
        TheResult.GetResult ParamLabel & "_Info_Vadd", DefInfo_Vadd
        TheResult.GetResult ParamLabel & "_Info_Dire", DefInfo_Dire
        TheResult.GetResult ParamLabel & "_Info_Sorc", DefInfo_Src1
        
        For site = 0 To nSite                                                                               'site�܂킵
            If TheExec.sites.site(site).Active = True Then
                If OverFlowCheck(site) = 0 Then                                                             '�␳������I�[�o�[���Ă��Ȃ���΁A�␳�p�ϐ��ɕ␳�����i�[

                    For NowRepairNo = 0 To DefInfo_Num(site) - 1                                            '���݂̌��׍��ڂ̌���LOOP
                        SameAddress = False                                                                 '���׃A�h���X���Ԃ�Check�t���O���N���A
                        
                        '===== ADDRESS CHECK =====
                        If AllDefInfo_Num(site) > 0 Then
                            For i(site) = 0 To AllDefInfo_Num(site) - 1
                                If AllDefInfo_Hadd(site, i(site)) = DefInfo_Hadd(site, NowRepairNo) And _
                                   AllDefInfo_Vadd(site, i(site)) = DefInfo_Vadd(site, NowRepairNo) Then
                                    SameAddress = True                                                      '���׃A�h���X���Ԃ肪����������t���O��True
                                    Exit For
                                End If
                            Next i
                        End If
                        
                        '===== DEFECT ADDITION =====
                        If SameAddress = False Then                                                         '���׃A�h���X���Ԃ�łȂ���΁A�␳�p�ϐ��ɕ␳�����i�[
                            AllDefInfo_Num(site) = AllDefInfo_Num(site) + 1                                 'Tatal�̕␳�����C���N�������g
                            
                            If AllDefInfo_Num(site) <= MaxRepairNo Then                                     '�␳����𒴂��Ă��Ȃ������`�F�b�N
                                '----- X address -----
                                AllDefInfo_Hadd(site, AllDefInfo_Num(site) - 1) = DefInfo_Hadd(site, NowRepairNo)
                                '----- Y address -----
                                AllDefInfo_Vadd(site, AllDefInfo_Num(site) - 1) = DefInfo_Vadd(site, NowRepairNo)
                                '----- Direction -----
                                AllDefInfo_Dire(site, AllDefInfo_Num(site) - 1) = DefInfo_Dire(site, NowRepairNo)
                                '----- Sorce -----
                                Call BlowSorceInfoCheck(SourceType, DefInfo_Src1(site, NowRepairNo), DefInfo_Src2, Go_NgAddCheck)    'Single(PD)/CP/FD��Source�R�[�h�̊m�F
                                AllDefInfo_Src(site, AllDefInfo_Num(site) - 1) = DefInfo_Src2
                            
                                '----- NG ADDRESS CHECK -----
                                If Go_NgAddCheck = True Then
                                    '----- ZONE LEFT (Center Pix) -----
                                    If NgAdd_LeftS <= DefInfo_Hadd(site, NowRepairNo) And DefInfo_Hadd(site, NowRepairNo) <= NgAdd_LeftE Then
                                        NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                    End If
                                    '----- ZONE LEFT (Direction Pix) -----
                                    If DefInfo_Dire(site, NowRepairNo) = 0 Then     'Direction=Right
                                        If NgAdd_LeftS <= DefInfo_Hadd(site, NowRepairNo) + 2 And DefInfo_Hadd(site, NowRepairNo) + 2 <= NgAdd_LeftE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 1 Then     'Direction=RightBottom
                                        If NgAdd_LeftS <= DefInfo_Hadd(site, NowRepairNo) + 2 And DefInfo_Hadd(site, NowRepairNo) + 2 <= NgAdd_LeftE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 2 Then     'Direction=Bottom
                                        'Bottom = X-Address is Same Center Address
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 3 Then     'Direction=LeftBottom
                                        If NgAdd_LeftS <= DefInfo_Hadd(site, NowRepairNo) - 2 And DefInfo_Hadd(site, NowRepairNo) - 2 <= NgAdd_LeftE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                    
                                    '----- ZONE RIGHT (Center Pix) -----
                                    If NgAdd_RightS <= DefInfo_Hadd(site, NowRepairNo) And DefInfo_Hadd(site, NowRepairNo) <= NgAdd_RightE Then
                                        NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                    End If
                                    '----- ZONE RIGHT (Direction Pix) -----
                                    If DefInfo_Dire(site, NowRepairNo) = 0 Then     'Direction=Right
                                        If NgAdd_RightS <= DefInfo_Hadd(site, NowRepairNo) + 2 And DefInfo_Hadd(site, NowRepairNo) + 2 <= NgAdd_RightE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 1 Then     'Direction=RightBottom
                                        If NgAdd_RightS <= DefInfo_Hadd(site, NowRepairNo) + 2 And DefInfo_Hadd(site, NowRepairNo) + 2 <= NgAdd_RightE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 2 Then     'Direction=Bottom
                                        'memo : Bottom = X-Address is Same Center Address
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 3 Then     'Direction=LeftBottom
                                        If NgAdd_RightS <= DefInfo_Hadd(site, NowRepairNo) - 2 And DefInfo_Hadd(site, NowRepairNo) - 2 <= NgAdd_RightE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                
                                End If

                            Else
                                OverFlowCheck(site) = 1                                                     '�␳����𒴂��Ă�����t���O�𗧂Ă�
                                Exit For
                            End If
                        End If
                    Next NowRepairNo
    
                End If
                
                '##### DEBUG LOG OUTPUT #####
                If Flg_Debug = 1 Then
                    TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Num" & " = " & DefInfo_Num(site)
                    For NowRepairNo = 0 To DefInfo_Num(site) - 1
                        TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Hadd" & " = " & DefInfo_Hadd(site, NowRepairNo)
                        TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Vadd" & " = " & DefInfo_Vadd(site, NowRepairNo)
                        If DefInfo_Src1(site, NowRepairNo) = 0 Then
                            TheExec.Datalog.WriteComment "SourceCode Nothing"
                        End If
                        If DefInfo_Src1(site, NowRepairNo) = 1 Then
                            TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Sorc" & " = " & "Single"
                        End If
                        If DefInfo_Src1(site, NowRepairNo) = 2 Then
                            TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Sorc" & " = " & "Couplet"
                        End If
                        If DefInfo_Src1(site, NowRepairNo) = 3 Then
                            TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Sorc" & " = " & "FD"
                        End If
                        TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Dire" & " = " & DefInfo_Dire(site, NowRepairNo)
                    Next NowRepairNo
                End If
                
            End If
        Next site
                        
    Next ParamLoop


    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
                        
            If OverFlowCheck(site) = 0 And NgAddressCheck(site) = 0 Then
            
                '========== �␳������Modify�ϐ��֊i�[ ================
                BlowData_All(site) = Dec2Bin(CStr(AllDefInfo_Num(site)), BitWidthN)
                
                '========== CUPLET�␳�A�h���X����Modify�ϐ��֊i�[ ======
                For NowRepairNo = 0 To AllDefInfo_Num(site) - 1
                        BlowData_All(site) = BlowData_All(site) _
                                            + Dec2Bin(CStr(AllDefInfo_Src(site, NowRepairNo)), BitWidthS) _
                                            + Dec2Bin(CStr(AllDefInfo_Dire(site, NowRepairNo)), BitWidthD) _
                                            + Dec2Bin(CStr((AllDefInfo_Hadd(site, NowRepairNo) + OtpPixOffset_X)), BitWidthX) _
                                            + Dec2Bin(CStr((AllDefInfo_Vadd(site, NowRepairNo) + OtpPixOffset_Y)), BitWidthY)
                Next NowRepairNo
                        
                
                '========== �␳����ɑ���Ȃ�����"0"�ł��߂� =============�@�@'�O�`�b�v��Modify��񂪃p�^�[���Ɏc���Ă邩��]��ɂ�0�������K�v
                amari(site) = MaxRepairNo - AllDefInfo_Num(site)
                
                If amari(site) > 0 Then
                    For NowRepairNo = 0 To amari(site) - 1
                            BlowData_All(site) = BlowData_All(site) _
                                                + Dec2Bin("0", BitWidthS) _
                                                + Dec2Bin("0", BitWidthD) _
                                                + Dec2Bin("0", BitWidthX) _
                                                + Dec2Bin("0", BitWidthY)
                    Next NowRepairNo
                End If
                            
                '========== Blow�p�ϐ��쐬 ====================================
                Call MakeModifyData(BlowData_All, BlowDataAllBin, Page, Bit, site)

            End If
            
        End If
    Next site

End Sub

Public Sub BlowSorceInfoCheck(ByVal SourceType As String, ByRef SrcIn As Double, ByRef SrcOut As Double, ByRef Go_NgAddCheck As Boolean)
'2013/8/2�@SrcType1�̂݁BSource�R�[�h�̃i���o�����O���ς�����^�C�v���o�Ă�����A���̓s�x�ǉ����K�v�B

    Go_NgAddCheck = False
    
    Select Case SourceType
        Case "SrcType1"
            If SrcIn = 1 Then SrcOut = 1
            If SrcIn = 2 Then
                SrcOut = 2
                Go_NgAddCheck = True
            End If
            If SrcIn = 3 Then SrcOut = 3
        Case Else
            MsgBox "Error!! Nothing SourceType!!"
            Stop
    End Select
        
End Sub

