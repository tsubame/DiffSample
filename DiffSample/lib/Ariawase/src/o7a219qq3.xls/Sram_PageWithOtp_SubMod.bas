Attribute VB_Name = "Sram_PageWithOtp_SubMod"
Option Explicit

Public Sub Set_SramCondition(ByRef ArgArr() As String)

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h

    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_VARIABLE_PARAM) Then
        Err.Raise 9999, "SRAM", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
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
        Err.Raise 9999, "SRAM", """#EOP"" is not found! [" & GetInstanceName & "] !"
    End If

    Dim testConditionList() As String
    testConditionList = Split(ArgArr(0), ",")
    For i = 0 To UBound(testConditionList)
        If Trim(testConditionList(i) <> "") Then
            Call TheCondition.SetCondition(testConditionList(i))
        End If
    Next i

End Sub

Public Sub Set_SramCondition_SVDR(ByRef ArgArr() As String, ByRef LowVddPins() As String, ByRef Vdd2nd As String, ByRef Vdd3rd As String, ByRef LowVddWait As Double, ByRef WritePat As String, ByRef ReadPat As String)
    
    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h

    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_VARIABLE_PARAM) Then
        Err.Raise 9999, "SRAM", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
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
        Err.Raise 9999, "SRAM", """#EOP"" is not found! [" & GetInstanceName & "] !"
    End If

    Dim testConditionList() As String
    testConditionList = Split(ArgArr(0), ",")
    For i = 0 To UBound(testConditionList)
        If Trim(testConditionList(i) <> "") Then
            Call TheCondition.SetCondition(testConditionList(i))
        End If
    Next i


'SVDR ONLY��

    '�d����������Pin���擾
    LowVddPins = Split(ArgArr(4), ",")

    '������d���l�Ɩ߂��d���l�̎擾
    Dim VddName() As String
    VddName = Split(ArgArr(1), ",")
    Vdd2nd = VddName(1)
    Vdd3rd = VddName(2)

    '�������d�����ێ����鎞�Ԃ̎擾
    LowVddWait = ArgArr(3)

    'Write/Read�p�^�[���̎擾
    Dim pat() As String
    pat = Split(ArgArr(2), ",")
    WritePat = pat(0)
    ReadPat = pat(2)
    
End Sub


'�V�[�g����g�p����p�^�[����TBL�t�@�C����R�t����
Public Sub READ_TBL_LIST()

    
    Worksheets("TBL_LIST").Select
    
    Dim i As Long
    i = 0
    Erase TblInfo
    Do While (Cells(TBL_CELL_LIST_STRow + i, TBL_CELL_PAT) <> "")
        ReDim Preserve TblInfo(i)
        TblInfo(i).PatFileName = Cells(TBL_CELL_LIST_STRow + i, TBL_CELL_PAT)
        TblInfo(i).TblFileName = Cells(TBL_CELL_LIST_STRow + i, TBL_CELL_PAT + 1)
        i = i + 1
    Loop

    dirTblFile = Cells(TBL_DIR_CELL_Row, TBL_DIR_CELL_Col)


End Sub

'TBL�t�@�C���̒��g�����[�h����
Public Sub READ_TBL_FILE()

    Dim intFileNo As Integer
    Dim nTBLFile As Long
    Dim nLine As Long
    Dim i As Long, j As Long, k As Long
    Dim tmp As String, strLineData() As String
    Dim strSplitData() As String
    Dim lngBitLength As Long, lngPosiPrefix As Long, lngPosiPostfix As Long
    
    On Error GoTo ERROR_READ_FILE
    
    nTBLFile = UBound(TblInfo)
    For i = 0 To nTBLFile
        Erase strLineData
        intFileNo = FreeFile
        Open dirTblFile & TblInfo(i).TblFileName For Input As #intFileNo
        Do Until EOF(intFileNo)
            Line Input #intFileNo, tmp
        Loop
        Close intFileNo
        
        strLineData = Split(tmp, vbLf)  'TBL�t�@�C���̉��s�R�[�h��LF
        nLine = UBound(strLineData)
        k = 0
        For j = 0 To nLine
            Erase strSplitData
            If strLineData(j) = "" Then
                Exit For
            End If
            strSplitData = Split(strLineData(j), vbTab)
            If Left(strSplitData(0), 1) <> "#" Then
                With TblInfo(i)
                    ReDim Preserve .FailInfo(k)
                    .FailInfo(k).CycleNo = CLng(strSplitData(TBL_INDEX_CYCLE))
                    .FailInfo(k).MemoryNo = CLng(strSplitData(TBL_INDEX_MACRO))
                    lngBitLength = Len(strSplitData(TBL_INDEX_BIT))
                    lngPosiPrefix = InStr(strSplitData(TBL_INDEX_BIT), "[")
                    lngPosiPostfix = InStr(strSplitData(TBL_INDEX_BIT), "]")
                    .FailInfo(k).IoNo = CLng(Mid(strSplitData(TBL_INDEX_BIT), lngPosiPrefix + 1, lngBitLength - lngPosiPrefix - (lngBitLength - lngPosiPostfix + 1)))
                End With
                k = k + 1
            End If
        Next j
    Next i
    
    
    Exit Sub
    
ERROR_READ_FILE:
    MsgBox "ERROR!!! TBL_FILE READ"
End Sub

Public Sub SramPatRun_GetFailInfo(Pattern As String, ByRef Judge_value() As Double, Optional ByVal Exec_Site As Long = ALL_SITE)

    Dim meas_result As Integer
    Dim meas_loop_cnt  As Integer
    Dim FailNo_End As Integer
    Dim FailNo As Long
    Dim FailCycle As Long
    Dim FailCycle_idx As Long
    Dim FailCount As Long
    Const HramSize As Integer = 255
    Dim FailCycle_start As Variant
    Dim FailSramInfo(nSite, Bist_Num_Mem + 1, Bist_Max_Num_Io + 1) As Byte
    Dim Flag_ALPG_Error(nSite) As Integer
    Dim serch_left As Long
    Dim serch_right As Long
    Dim PinPfData As String
    Dim Flag_SramFail(nSite) As Integer
    Dim xx As Integer
    Dim yy As Integer
    Dim keepVal(nSite) As Double
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim site As Long
    Dim presite As Long
    Dim ALPG_LogOutPut(nSite) As Boolean
    
    For presite = 0 To nSite
        ALPG_LogOutPut(presite) = True
    Next presite
    
    Dim TblNum As Long
    Dim i As Long
    Dim j As Long
    
    '===== Choice USE TBL Number =====================
    i = UBound(TblInfo)
    For j = 0 To i
        If CStr(Pattern) = "PG_" & CStr(TblInfo(j).PatFileName) Then
            TblNum = j
            Exit For
        End If
    Next j
    
    Erase Flag_ALPG_Error
    serch_left = 0
    serch_right = UBound(TblInfo(TblNum).FailInfo)

    meas_result = 0
    meas_loop_cnt = 0
    FailCount = 0
    FailNo_End = 0
    FailCycle = 0
    FailCycle_idx = 0
    FailCycle_start = 0            '0x0
    
    Erase FailSramInfo
    Erase Judge_value

    '��������������������������������������������������������������������������
    '���p�^�[���Z�b�g�@�`�@�����@�`�@Fail���擾
    '��������������������������������������������������������������������������
    If TheExec.sites.ActiveCount > 0 Then
    
        '===== Pattern & Hram Set =============================================
        Call StopPattern
        TheHdw.Digital.Timing.Load
        TheHdw.Digital.Patterns.pat(Pattern).Load
        TheHdw.Digital.HRAM.Size = HramSize
        TheHdw.Digital.Patgen.NoHaltMode = noHaltAlways                 'noHaltAlways: Fail Stop Cancel
        Call TheHdw.Digital.HRAM.SetTrigger(trigFail, True, 0, True)     'trig: fail, cycle�w��: true, before cycle: 0, stopOnFull: true
        Call TheHdw.Digital.HRAM.SetCapture(captFail, False)            'captFail: Fail Caputure, comprepeat: ���s�[�g�����������Ď�荞��


        Do
            meas_loop_cnt = meas_loop_cnt + 1

            TheHdw.Digital.Patgen.EventCycleEnabled = True
            TheHdw.Digital.Patgen.EventCycleCount = FailCycle_start + 1
            TheHdw.Digital.Patgen.MaskTilCycle = True

            '===== Pat Run ====================================================
            Call TheHdw.Digital.Patterns.pat(Pattern).Run("START")
            
            '===== Pat Fail Count =============================================
            FailCount = TheHdw.Digital.Patgen.FailCount
            If FailCount < TheHdw.Digital.HRAM.CapturedCycles Then FailCount = TheHdw.Digital.HRAM.CapturedCycles    ' When Patgen.FailCount is large than 65536, it is counted from 0. FailCount use HRAM.CaputuredCycles when Patgen.FailCount is large than 65536.

            If FailCount > HramSize Then
                FailNo_End = HramSize
            Else
                FailNo_End = FailCount
            End If
            
            
            'SSSSS SKIP START SSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS
            If FailCount > 0 Then
                
                    '===== Check Fail Infomation ==================================
                    For FailNo = 0 To FailNo_End - 1
                        FailCycle = TheHdw.Digital.HRAM.ReadPatGenInfo(FailNo, pgCycle)                         'FailCycle     : Fail���Ă���Cycle�i���o�[
                        FailCycle_idx = Serch_ArrangementNo(FailCycle, TblNum, serch_left, serch_right)       'FailCycle_idx : Fail���Ă���Cycle�̔z��i���o�[
        
                        siteStatus = TheExec.sites.SelectFirst
                        Do While siteStatus <> loopDone
                            curSite = TheExec.sites.SelectedSite
                            If Exec_Site = ALL_SITE Or Exec_Site = curSite Then
                                PinPfData = TheHdw.Digital.HRAM.Pins(SRAMRD_OUTPUT_PIN).PinPF(FailNo)
                                If PinPfData Like "*F*" Then
                                
                                    If FailCycle_idx = -1 Then
                                        '##### ALPG Error #####
                                        Flag_ALPG_Error(curSite) = 1   'Flag
                                        Judge_value(curSite) = 4    'Result
                                        '##### DEBUG LOG OUTPUT #####
                                        If Flg_Debug = 1 And ALPG_LogOutPut(curSite) = True Then
                                            TheExec.Datalog.WriteComment "Pre-SRAM-BIST Fail Infomation" & " " & "[Site" & curSite & "]" & " " & "ALPG ERROR"
                                            ALPG_LogOutPut(curSite) = False
                                        End If
                                    Else
                                        With TblInfo(TblNum).FailInfo(FailCycle_idx)
                                            '##### Fail���Ă��郁�����[��IO�̔z��i���o�[��"1"�𗧂Ă� #####
                                            FailSramInfo(curSite, .MemoryNo, .IoNo) = 1
                                            
                                            '===== Skip Flag ==============================================
                                            Flag_SramFail(curSite) = Flag_SramFail(curSite) + 1

                                            '##### DEBUG LOG OUTPUT #####
                                            If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Pre-SRAM-BIST Fail Infomation" & " " & "[Site" & curSite & "]" & " " & "FailMemory:" & .MemoryNo & " " & "FailIO:" & .IoNo
                                        End With
                                    End If
                                
                                End If
                            End If
                            siteStatus = TheExec.sites.SelectNext(siteStatus)
                        Loop
                    
                    Next FailNo
                
            End If
            'EEEEE SKIP END EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE


            FailCycle_start = FailCycle     ' FailCycle is the last fail pattern count

        Loop While FailCount > HramSize And meas_loop_cnt < serch_right \ HramSize + 10 'mugen loop boushi (multi?) patterncycle/255 ga best

      ' MASK OFF
        TheHdw.Digital.Patgen.EventCycleEnabled = False
        TheHdw.Digital.Patgen.EventCycleCount = 1
        TheHdw.Digital.Patgen.MaskTilCycle = False
    End If



    '��������������������������������������������������������������������������
    '������BIST���������ł̃��y�A����i����BIST�����Ƃ�OR����������y�A����ł͂Ȃ��̂Œ��Ӂj
    '��������������������������������������������������������������������������
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If Exec_Site = ALL_SITE Or Exec_Site = site Then
                
                If (Judge_value(site) <> 4) And Flag_SramFail(site) > 0 Then      'ALPG->PASS  &  SRAM->FAIL
                    
                    '===== Repair Check =======================================
                    Judge_value(site) = Judge_Repair_ThisBist(FailSramInfo(), site)
                    
                    '===== Fail Infomation Merge ==============================
                    For xx = 0 To Bist_Num_Mem + 1
                        For yy = 0 To Bist_Max_Num_Io + 1
                            BIST_FAIL_REG(site, xx, yy) = FailSramInfo(site, xx, yy) Or BIST_FAIL_REG(site, xx, yy)
                        Next yy
                    Next xx
                    
                ElseIf Judge_value(site) = 4 Then
                    Bist_Alpg_Fail_Flag(site) = Flag_ALPG_Error(site) Or Bist_Alpg_Fail_Flag(site)
                End If
                
                '===== Result��0/1����̈Ӗ����t�]������B(LOGIC������1=Pass 0=Fail�Ƃ����Ӗ��ɓ���)
                keepVal(site) = Judge_value(site)
                If keepVal(site) = 0 Then Judge_value(site) = 1
                If keepVal(site) = 1 Then Judge_value(site) = 0

            End If
        End If
    Next site

End Sub

Public Function Serch_ArrangementNo(ByVal FailCycle As Long, ByVal TblNum As Long, ByVal Left As Integer, ByVal Right As Long) As Integer
'Memo:Fail���Ă���Cycle���A�p�^�[�����ϐ��̉��Ԗڂ̔z��ł��邩����������
    
    Dim middle As Integer
    Dim serch_end As Integer

    serch_end = 0

    Do While Left <= Right
        middle = (Left + Right) / 2

        If TblInfo(TblNum).FailInfo(middle).CycleNo = FailCycle Then
            Serch_ArrangementNo = middle
            serch_end = 1
            Exit Do
        End If
        If TblInfo(TblNum).FailInfo(middle).CycleNo > FailCycle Then
            Right = middle - 1
        Else
            Left = middle + 1
        End If
    Loop

    If serch_end <> 1 Then              '�΂��悯
        Serch_ArrangementNo = -1
    End If

End Function

Private Function Judge_Repair_ThisBist(FailSramInfo() As Byte, tsite As Long) As Long
'����1��BIST�������ł́ASRAM�璷�۔�����s��Function

    Dim mem_no As Integer
    Dim io_no As Integer
    Dim fail_io_cnt As Integer
    

    For mem_no = 1 To Bist_Num_Mem                                              'Memory��
        
        '===== Get Fail I/O Total No ==========================================
        fail_io_cnt = 0
        For io_no = BIST_NUM_IO(mem_no) To 1 Step -1                            'I/O��
            If FailSramInfo(tsite, mem_no, io_no) = 1 Then
                fail_io_cnt = fail_io_cnt + 1                                   '���݂�Memory���̕s��IO��
            End If
        Next io_no

        '===== JUDGE!! ======================================================== '�璷�Z�������s�ǂƃ}���`I/O�s�ǂ�������ꍇ�ɂ́A�����ł͏璷�Z�������s�ǂƂ��Č��ʂ�f���o���B
        
        If BIST_RED_TYPE(mem_no) = 0 And fail_io_cnt > 0 Then                   '�璷�Z������Memory�s�ǔ���
            Judge_Repair_ThisBist = 3
            '##### DEBUG LOG OUTPUT #####
            If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Pre-SRAM-BIST Fail Infomation" & " " & "[Site" & tsite & "]" & " " & "DON'T HAVE REPAIR-IO-CELL (This BIST Result) Memory:" & mem_no

        ElseIf fail_io_cnt > 1 Then                                             '�}���`I/O�s�ǔ���
            If Judge_Repair_ThisBist < 2 Then Judge_Repair_ThisBist = 2
            '##### DEBUG LOG OUTPUT #####
            If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Pre-SRAM-BIST Fail Infomation" & " " & "[Site" & tsite & "]" & " " & "MULTI Fail IO-CELL (This BIST Result) Memory:" & mem_no
        
        ElseIf fail_io_cnt > 0 Then                                             '���y�A�\����
            If Judge_Repair_ThisBist < 1 Then Judge_Repair_ThisBist = 1
            '##### DEBUG LOG OUTPUT #####
            If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Pre-SRAM-BIST Fail Infomation" & " " & "[Site" & tsite & "]" & " " & "REPAIR POSSIBLE (This BIST Result) Memory:" & mem_no

        End If
    
    Next mem_no


End Function

Public Sub Judge_Repair_ThisChip(tsite As Long, ByRef JUDGE_FLAG() As Double)
'�S�Ă�PreBIST�������I�����Ă����OR��������ASRAM�璷�۔�����s��Function
'�������璷�\�ł���Ɣ��肳�ꂽ��A�璷�f�[�^���쐬���ɍs��

    Dim mem_no As Integer
    Dim io_no As Integer
    Dim fail_io_cnt As Integer


    For mem_no = 1 To Bist_Num_Mem                                                      'Memory��
        
        '===== Get Fail I/O Infomation ========================================
        fail_io_cnt = 0
        BIST_FAIL_IO_NO(tsite, mem_no) = -1
        
        For io_no = BIST_NUM_IO(mem_no) - 1 To 0 Step -1                                'I/O��
            If BIST_FAIL_REG(tsite, mem_no, io_no) = 1 Then
                fail_io_cnt = fail_io_cnt + 1                                           '���݂�Memory���̕s��I/O���i�SSRAM������OR�����j
                BIST_FAIL_IO_NO(tsite, mem_no) = io_no                                  '���݂�Memory���ŕs�ǂƂȂ��Ă���I/O�i���o�[
            End If
        Next io_no

        '===== JUDGE!! ========================================================
        If fail_io_cnt > 1 Then
            JUDGE_FLAG(tsite) = 10          '�}���`I/O�s�ǔ���
            '##### DEBUG LOG OUTPUT #####
            If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Pre-SRAM-BIST Fail Infomation" & " " & "[Site" & tsite & "]" & " " & "MULTI Fail IO-CELL (All BIST Result)"
        
        ElseIf fail_io_cnt > 0 And JUDGE_FLAG(tsite) < 10 Then
            JUDGE_FLAG(tsite) = 1           '���y�A�\����
            '##### DEBUG LOG OUTPUT #####
            If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Pre-SRAM-BIST Fail Infomation" & " " & "[Site" & tsite & "]" & " " & "REPAIR is POSSIBLE (All BIST Result)"
        
        End If
        
    Next mem_no

    
    '===== Make Repair RCON Data ==============================================
    If Bist_Alpg_Fail_Flag(tsite) = 0 And JUDGE_FLAG(tsite) = 1 Then                    'ALPG->PASS  &  SRAM->FAIL
        Call Set_RepairData_RCON(tsite, RepairMemoryCount(tsite))                       'RCON�֏����ރf�[�^�쐬�i�f�[�^���k�O�j�B�K�vMemory�璷�����A�ő�\Memoy�璷���𖞂����Ă��邩���m�F�B
        Bist_Repairable_Flag(tsite) = 1                                                 '�璷Blow���s�t���O�𗧂Ă�
    End If
    
    
    '===== Chack Repair Memory Count ==========================================
    If RepairMemoryCount(tsite) > MAX_EF_BIST_RD_BIT Then                                      '�K�vMemory�璷�� > �ő�\Memoy�璷���@�ł���΁AFail�̓����l�ɂ��āA�璷Blow�t���O��0�ɂ���
        Bist_Repairable_Flag(tsite) = 0
        JUDGE_FLAG(tsite) = 20
    End If
    
    '===== 'PostSramBist���s�t���O ============================================
    If Bist_Repairable_Flag(tsite) = 1 Then
        Flg_PostSramRun = True
    End If

End Sub

Public Sub Set_RepairData_RCON(tsite As Long, RepairMemoryCount As Long)
'���y�A����RCON�`���f�[�^�ɕϊ�
    
    Dim mem_no As Integer
    Dim bit_width As Integer
    Dim FailInfo As Integer
    Dim failinfo_mod As Integer
    Dim failinfo_no As Integer
    

    For mem_no = 1 To Bist_Num_Mem                                                          'Memory��
        If BIST_FAIL_IO_NO(tsite, mem_no) >= 0 Then                                         '���݂�Memory�ɕs��I/O�������RCON�璷�f�[�^���������֓˓�
        If BIST_RED_TYPE(mem_no) = 1 Then                                                   '���݂�Memory���璷�Z����ێ����Ă����RCON�璷�f�[�^���������֓˓�
            bit_width = Len(Dec2Bin(CStr(BIST_NUM_IO(mem_no)), 1))                              '���݂�Memory�ɕR�t������RCON�̃f�[�^Bit��(EN��BIT�͊܂܂Ȃ�)

            '===== Set EN Bit =================================================             '���݂�Memory�ɕR�t������RCON�̏璷EN��BIT�Ƀt���O�𗧂Ă�
            If RCON_FirstInfoType = "MEMID_1st" Then
                EF_BIST_REPAIR_DATA(tsite, BIST_IO_EN_NO(mem_no)) = 1
            ElseIf RCON_FirstInfoType = "FAILINFO_1st" Then
                If RCON_ChainType = "Descending" Then
                    EF_BIST_REPAIR_DATA(tsite, (BIST_IO_EN_NO(mem_no) - bit_width)) = 1
                ElseIf RCON_ChainType = "Ascending" Then
                    EF_BIST_REPAIR_DATA(tsite, (BIST_IO_EN_NO(mem_no) + bit_width)) = 1
                End If
            End If
                
            '===== Set Data Bit ===============================================
            FailInfo = BIST_FAIL_IO_NO(tsite, mem_no)                                       '���݂�Memory��Fail���Ă���I/O�i���o�[
            
            For failinfo_no = 0 To (bit_width - 1)
                failinfo_mod = FailInfo Mod 2                                               'Dec -> Bin
                
                'RCON�̎d�l�ɂ���āA�Z�b�g����f�[�^���Ԃ��قȂ�B
                '�d�l������Ă��A���ʂƂ��ē����v�Z���ɂȂ�g�ݍ��킹�����邪�A�킩��₷���悤�ɂ����đS�ẴP�[�X���ׂ������L�q���Ă���B
                If RCON_ChainType = "Ascending" And RCON_FirstInfoType = "MEMID_1st" And RCON_FailInfoType = "Ascending" Then           '�`�F�[�����F�����@  1stBit�FMEMID�@�@ �@FAILINFO���F����
                    EF_BIST_REPAIR_DATA(tsite, BIST_IO_EN_NO(mem_no) + 1 + failinfo_no) = failinfo_mod                                  'ex)IMX170
                                
                ElseIf RCON_ChainType = "Ascending" And RCON_FirstInfoType = "MEMID_1st" And RCON_FailInfoType = "Descending" Then      '�`�F�[�����F�����@  1stBit�FMEMID�@�@ �@FAILINFO���F�~��
                    EF_BIST_REPAIR_DATA(tsite, BIST_IO_EN_NO(mem_no) + bit_width - failinfo_no) = failinfo_mod                          'ex)IMX145
                
                ElseIf RCON_ChainType = "Ascending" And RCON_FirstInfoType = "FAILINFO_1st" And RCON_FailInfoType = "Ascending" Then    '�`�F�[�����F�����@�@1stBit�FFAILINFO�@�@FAILINFO���F����
                    EF_BIST_REPAIR_DATA(tsite, BIST_IO_EN_NO(mem_no) + failinfo_no) = failinfo_mod
                
                ElseIf RCON_ChainType = "Ascending" And RCON_FirstInfoType = "FAILINFO_1st" And RCON_FailInfoType = "Descending" Then   '�`�F�[�����F�����@�@1stBit�FFAILINFO�@�@FAILINFO���F�~��
                    EF_BIST_REPAIR_DATA(tsite, BIST_IO_EN_NO(mem_no) + bit_width - 1 - failinfo_no) = failinfo_mod
                
                ElseIf RCON_ChainType = "Descending" And RCON_FirstInfoType = "MEMID_1st" And RCON_FailInfoType = "Ascending" Then      '�`�F�[�����F�~���@  1stBit�FMEMID�@�@ �@FAILINFO���F����
                    EF_BIST_REPAIR_DATA(tsite, BIST_IO_EN_NO(mem_no) - 1 - failinfo_no) = failinfo_mod                                  'ex)ISX014,IMX164
                    
                ElseIf RCON_ChainType = "Descending" And RCON_FirstInfoType = "MEMID_1st" And RCON_FailInfoType = "Descending" Then     '�`�F�[�����F�~���@  1stBit�FMEMID�@�@ �@FAILINFO���F�~��
                    EF_BIST_REPAIR_DATA(tsite, BIST_IO_EN_NO(mem_no) - bit_width + failinfo_no) = failinfo_mod
                
                ElseIf RCON_ChainType = "Descending" And RCON_FirstInfoType = "FAILINFO_1st" And RCON_FailInfoType = "Ascending" Then   '�`�F�[�����F�~���@  1stBit�FFAILINFO�@�@FAILINFO���F����
                    EF_BIST_REPAIR_DATA(tsite, BIST_IO_EN_NO(mem_no) - failinfo_no) = failinfo_mod
                
                ElseIf RCON_ChainType = "Descending" And RCON_FirstInfoType = "FAILINFO_1st" And RCON_FailInfoType = "Descending" Then  '�`�F�[�����F�~���@  1stBit�FFAILINFO�@�@FAILINFO���F�~��
                    EF_BIST_REPAIR_DATA(tsite, BIST_IO_EN_NO(mem_no) - bit_width + 1 + failinfo_no) = failinfo_mod
                
                End If
                
                FailInfo = Int(FailInfo \ 2)                                                'Dec -> Bin
            Next failinfo_no
            
            RepairMemoryCount = RepairMemoryCount + 1                                       '�s�ǂ�Memory�����J�E���g
        End If
        End If
    Next mem_no


End Sub

Private Function Dec2Bin(myDecvalue As String, OutBit As Integer) As String
'10�i����2�i���ɕϊ�����

    Dim lngdecnumber As Long
    Dim strbinnumber As String
    strbinnumber = ""
    lngdecnumber = 0

    lngdecnumber = CLng(myDecvalue)

    Do
        strbinnumber = strbinnumber & CStr(lngdecnumber Mod 2)
        lngdecnumber = Fix(lngdecnumber / 2)
    Loop While lngdecnumber > 0

    Do While Len(strbinnumber) < OutBit
        strbinnumber = strbinnumber & "0"
    Loop

    Dec2Bin = strbinnumber

End Function

Public Sub Comp_RconData_ChainType_Descending(tsite As Long)
'RCON�`���̃��y�A�f�[�^���ABlow���邽�߂̌`���ւƃf�[�^�ϊ����邽�߂̏����Ƃ��Ĉ��k�f�[�^���쐬����B
'Caution!! -> RCON�̃`�F�[�������~����p�̏���

    Dim rcon_cnt As Integer
    Dim RepairNo As Integer
    Dim write_data(nSite, Ef_Bist_Rd_Data_Width - 1) As Integer
    Dim i As Long
    
    rcon_cnt = RCON_END_Addr
    
    
    Do
        If EF_BIST_REPAIR_DATA(tsite, rcon_cnt) = 1 Then
                       
            '===== Set Enable (Dec) ===========================================
            Ef_Enbl_Addr(tsite, RepairNo) = 1
            
            '===== Set Address (Dec) ==========================================
            Ef_Rcon_Addr(tsite, RepairNo) = rcon_cnt
            
            '===== Set Data (Dec) =============================================
            For i = 0 To Ef_Bist_Rd_Data_Width - 1
                If rcon_cnt + i < Rep_Bist_Data_Len - 1 Then
                    write_data(tsite, i) = EF_BIST_REPAIR_DATA(tsite, rcon_cnt + i)
                Else
                    write_data(tsite, i) = 0
                End If
            Next i
            
            For i = 0 To Ef_Bist_Rd_Data_Width - 1
                Ef_Repr_Data(tsite, RepairNo) = Ef_Repr_Data(tsite, RepairNo) + (write_data(tsite, Ef_Bist_Rd_Data_Width - 1 - i) * 2 ^ i)
            Next i
            
          
            If rcon_cnt + Ef_Bist_Rd_Data_Width < Rep_Bist_Data_Len - 1 Then
                rcon_cnt = rcon_cnt + Ef_Bist_Rd_Data_Width
            Else
                rcon_cnt = RCON_START_Addr + 1
            End If
            
            RepairNo = RepairNo + 1
            If RepairNo > MAX_EF_BIST_RD_BIT Then Exit Do
     
        Else
            rcon_cnt = rcon_cnt + 1
     
        End If
        
    Loop While rcon_cnt < Rep_Bist_Data_Len - 1
    
    
End Sub

Public Sub Comp_RconData_ChainType_Ascending(tsite As Long)
'RCON�`���̃��y�A�f�[�^���ABlow���邽�߂̌`���ւƃf�[�^�ϊ����邽�߂̏����Ƃ��Ĉ��k�f�[�^���쐬����B
'Caution!! -> RCON�̃`�F�[������������p�̏���

    Dim rcon_cnt As Integer
    Dim RepairNo As Integer
    Dim write_data(nSite, Ef_Bist_Rd_Data_Width - 1) As Integer
    Dim i As Long
    
    rcon_cnt = RCON_END_Addr
    
    
    Do
        If EF_BIST_REPAIR_DATA(tsite, rcon_cnt) = 1 Then
                       
            '===== Set Enable (Dec) ===========================================
            Ef_Enbl_Addr(tsite, RepairNo) = 1
            
            '===== Set Address (Dec) ==========================================
            Ef_Rcon_Addr(tsite, RepairNo) = rcon_cnt
            
            '===== Set Data (Dec) =============================================
            For i = 0 To Ef_Bist_Rd_Data_Width - 1
                If rcon_cnt - i > 0 Then
                    write_data(tsite, i) = EF_BIST_REPAIR_DATA(tsite, rcon_cnt - i)
                Else
                    write_data(tsite, i) = 0
                End If
            Next i
            
            For i = 0 To Ef_Bist_Rd_Data_Width - 1
                Ef_Repr_Data(tsite, RepairNo) = Ef_Repr_Data(tsite, RepairNo) + (write_data(tsite, Ef_Bist_Rd_Data_Width - 1 - i) * 2 ^ i)
            Next i
                        
          
            If rcon_cnt > Ef_Bist_Rd_Data_Width Then
                rcon_cnt = rcon_cnt - Ef_Bist_Rd_Data_Width
            Else
'                rcon_cnt = RCON_END_Addr - 1
                rcon_cnt = 0 '�Ƃ肠����0�ɂ��Ċ֐��𔲂���Brcon_cnt���}�C�i�X�ɂ͂ł��Ȃ�����B
            End If
            
            RepairNo = RepairNo + 1
            If RepairNo > MAX_EF_BIST_RD_BIT Then Exit Do
     
        Else
            rcon_cnt = rcon_cnt - 1
     
        End If
        
    Loop While rcon_cnt > 0

End Sub

Public Sub Make_VerifyData_SramRepair(tsite As Long, ByRef BlowData() As String, ByRef VerifyData() As String)

    Dim strSteps As Long
    
    VerifyData(tsite) = ""
    For strSteps = 1 To Len(BlowData(tsite))
        If Mid(BlowData(tsite), strSteps, 1) = "0" Then VerifyData(tsite) = VerifyData(tsite) & "L"
        If Mid(BlowData(tsite), strSteps, 1) = "1" Then VerifyData(tsite) = VerifyData(tsite) & "H"
    Next

End Sub

Public Function HRAM_INITAL_SramRep() As Double

    With TheHdw.Digital
        .HRAM.SetTrigger trigFail, False, 0, True
        .HRAM.SetCapture captFail, False
        .HRAM.Size = 256
        .Patgen.EventCycleEnabled = True
        .Patgen.EventCycleCount = 1
        .Patgen.EventSetVector False, "", "", 0
    End With

End Function

Public Function HRAM_SETUP_SramRep(ByVal UsePat As String, ByVal TrigerLabel As String, ByVal VectorOffset As Long, Optional ByVal CapSize As Integer = 256)

    With TheHdw.Digital
        .HRAM.SetTrigger trigFirst, True, 0, True
        .HRAM.SetCapture captAll, False
        .HRAM.Size = CapSize
        .Patgen.EventCycleEnabled = False
        .Patgen.EventSetVector True, UsePat, TrigerLabel, VectorOffset
    End With

End Function

Public Function Get_HramData_SramRep(ByRef MemData() As String, ByVal HRSize As Integer, ByVal RejiOut As String)

    Dim site As Long
    Dim site_status As Long
    Dim CapSize As Long
    
    '========== HRAM DATA SITE ============================
    site_status = TheExec.sites.SelectFirst
    While site_status <> loopDone
        site = TheExec.sites.SelectedSite
        
            For CapSize = 0 To (HRSize - 1)
                MemData(site) = MemData(site) & TheHdw.Digital.HRAM.Pins(RejiOut).PinData(CapSize)
            Next CapSize
        
        site_status = TheExec.sites.SelectNext(site_status)
    Wend

End Function

Public Sub Output_OtpBlowData_Sram()

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
    
    TheExec.Datalog.WriteComment "OTP BLOW DATA (SRAM REPAIR)"
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            PageA = 0
            For NowBit = 0 To BitParPage(PageA) - 1

                For NowPage = 0 To OtpPageSize - 1
                    If first = True Then
                        PageData = PageData + " : " & "P" & NowPage
                    End If
                    Data = Data + " :  " & SramBlowDataAllBin(site, NowPage, NowBit)
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

Public Sub Output_OtpReadData_Sram()

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
                    .Patgen.EventSetVector True, "OtpVerifyPage" & CStr(NowPage) & "_Pat", Label_OtpFixedValueCheck & CStr(NowPage), VectorOffset
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
                    
'                    '========== FUNCTION ERROR CHECK ==========================
'                    If TheHdw.Digital.FailedPinsCount(site) > 0 Then
'                        ReadErr(site) = ReadErr(site) + 1
'                    Else
'                        ReadErr(site) = ReadErr(site) + 0
'                    End If

                    site_status = TheExec.sites.SelectNext(site_status)
                Wend
                
                DataOffset = DataOffset + HramSize

            Next NowHramLoop
            
        Next NowPage
        
        '========== PATTERN SETUP =================================
        With TheHdw.Digital.Patgen
            .EventCycleEnabled = False
            .EventCycleCount = 0
            .MaskTilCycle = False
        End With

    End If



    TheExec.Datalog.WriteComment "OTP READ DATA (SRAM REPAIR)"
    
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
