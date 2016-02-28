Attribute VB_Name = "CreateDummySTDF"
Option Explicit

Type dataRecord
    OpCode      As String
    ParmName    As String
    TestNum     As Long
    testName    As String
    LoLimit     As Double
    HiLimit     As Double
    LimitTyp    As Integer
    ScaleVal    As Integer
    Unit        As String
    Form        As String
    PassBin     As Integer
    PassSortBin As Integer
    FailBin     As Integer
    FailSortBin As Integer
    result      As String
    BinName     As String
    KanjiName   As String
    OptFlag     As String
End Type

Dim JobData() As dataRecord
Dim numOfJobData As Integer

Const OpCodeCol As Integer = 7
Const ParmCol As Integer = 8
Const TNameCol As Integer = 9
Const TnumCol As Integer = 10
Const PassCol As Integer = 11
Const FailCol As Integer = 12
Const PassSortCol As Integer = 13
Const FailSortCol As Integer = 14
Const ResultCol As Integer = 15
Const BinNCol As Integer = 29
Const TnameInstanceCol = 2

Const OtherLimitBaseCol As Integer = 14
Const offsetLoLimit As Integer = 0
Const offsetHiLimit As Integer = 1
Const offsetLimType As Integer = 2
Const offsetUnit As Integer = 3
Const offsetForm As Integer = 4

Const OtherLoLimitCol As Integer = 14
Const OtherHiLimitCol As Integer = 15
Const OtherLimTypeCol As Integer = 16
Const OtherUnitCol As Integer = 17
Const OtherFormCol As Integer = 18

Const PPMUTempLimTypeCol As Integer = 33
Const PPMUTempHiLimCol As Integer = 34
Const PPMUTempLoLimCol As Integer = 35
Const PPMUTempUnitCol As Integer = 92       ' Arg78
Const PPMUTempFormCol As Integer = 93       ' Arg79

Const BPMUTempLimTypeCol As Integer = 37
Const BPMUTempHiLimCol As Integer = 38
Const BPMUTempLoLimCol As Integer = 39
Const BPMUTempUnitCol As Integer = 92       ' Arg78
Const BPMUTempFormCol As Integer = 93       ' Arg79

Const CTOAdcTempLimTypeCol As Integer = 33  ' There is no limit column for this type of test
Const CTOAdcTempHiLimCol As Integer = 34
Const CTOAdcTempLoLimCol As Integer = 35
Const CTOAdcTempUnitCol As Integer = 92     ' Arg78
Const CTOAdcTempFormCol As Integer = 93     ' Arg79

Const CTOPmuTempLimTypeCol As Integer = 33
Const CTOPmuTempHiLimCol As Integer = 34
Const CTOPmuTempLoLimCol As Integer = 35
Const CTOPmuTempUnitCol As Integer = 92       ' Arg78
Const CTOPmuTempFormCol As Integer = 93       ' Arg79

Const FunctionalTempLimTypeCol As Integer = 33  ' There is no limit column in this type of test
Const FunctionalTempHiLimCol As Integer = 34
Const FunctionalTempLoLimCol As Integer = 35
Const FunctionalTempUnitCol As Integer = 92     ' Arg78
Const FunctionalTempFormCol As Integer = 93     ' Arg79

Const PowerTempLimTypeCol As Integer = 32
Const PowerTempHiLimCol As Integer = 33
Const PowerTempLoLimCol As Integer = 34
Const PowerTempUnitCol As Integer = 92       ' Arg78
Const PowerTempFormCol As Integer = 93       ' Arg79

Const CTODacTempLimTypeCol As Integer = 33  ' There is no limit column in this type of test
Const CTODacTempHiLimCol As Integer = 34
Const CTODacTempLoLimCol As Integer = 35
Const CTODacTempUnitCol As Integer = 92     ' Arg78
Const CTODacTempFormCol As Integer = 93     ' Arg79

Const MTOMemTempLimTypeCol As Integer = 33  ' There is no limit column in this type of test
Const MTOMemTempHiLimCol As Integer = 34
Const MTOMemTempLoLimCol As Integer = 35
Const MTOMemTempUnitCol As Integer = 92     ' Arg78
Const MTOMemTempFormCol As Integer = 93     ' Arg79

Const TypeInstance As Integer = 3
Const NameInstance As Integer = 4

Const StartRowInFlow As Integer = 5
Const StartRowInInstance As Integer = 5

Const CouldNotCreateFile As Integer = 1
Const InstanceSheetNotFound As Integer = 2
Const FlowSheetNotFound As Integer = 3
Const MemoryError As Integer = 4
Const TypeMismatchError As Integer = 5
Const FormatCheckError As Integer = 6
Const BinnameSheetNotFound As Integer = 7
Const BnmCheckError As Integer = 8
'
'2012/11/16 175JobMakeDebug Arikawa nPath NG!!
Const SeqDcparName As String = "dcpar"
Const SeqDcparNumber As Integer = 2
Const SeqImageName As String = "image"
Const SeqImageNumber As Integer = 1002
Const SeqGradeName As String = "grade"
Const SeqGradeNumber As Integer = 5002
Const SeqShirotenName As String = "shiroten"
Const SeqShirotenNumber As Integer = 6002
Const SeqMarginName As String = "margin"
Const SeqMarginNumber As Integer = 8002
Const SeqMatchLength As Integer = 32

Const TnameMatchLength As Integer = 32
Const SpaceChar As String = " "

Const BnmSheetName As String = "BinNames"
Const BnmBinNumberCol As Integer = 2
Const BnmBinNameCol As Integer = 3
Const BnmBinNameRow As Integer = 6
Const BnmMasterRow As Integer = 3

Const DefaultBnmMasterFilename As String = "bnm_master_tmp.txt"
Private BnmMasterFilename As String

Const MaxBinSize = 100
Private SpecialBin(MaxBinSize) As String

Public Function GetGradeFirstTestNumber() As Long
    GetGradeFirstTestNumber = SeqGradeNumber
End Function

Public Function GetShirotenFirstTestNumber() As Long
    GetShirotenFirstTestNumber = SeqShirotenNumber
End Function

Public Function GetMarginFirstTestNumber() As Long
    GetMarginFirstTestNumber = SeqMarginNumber
End Function
Public Function tlSearchJobData(ByVal fileName As String, ByVal InstanceName As String, ByVal FlowName As String) As Long
    
    Dim endFlag As Integer
    Dim Row As Integer
    Dim wksht As Worksheet
    Dim i As Integer, j As Integer
    Dim n As Integer
    Dim SeqName As String
    Dim BinName(MaxBinSize) As String
    Dim tnum As Integer
    Dim fp As Integer
    Dim search_tnum As Integer
    
    Dim FlowSheetName As String
    Dim InstanceSheetName As String
    
    On Error GoTo err1SearchJob
    
    tlSearchJobData = 0 ' normal end

    If InstanceName = "" Then
        InstanceSheetName = "Test Instances"
    Else
        InstanceSheetName = InstanceName
    End If
    
    If FlowName = "" Then
        FlowSheetName = "Flow Table"
    Else
        FlowSheetName = FlowName
    End If
    
    endFlag = 0
    
    ' activate flow sheet
    For Each wksht In ActiveWorkbook.Sheets
        If wksht.Name = FlowSheetName Then
            wksht.Select
            Exit For
        End If
    Next
    
    If wksht.Name <> FlowSheetName Then
        outPutMessage "[Error] Flow sheet not found"
        tlSearchJobData = FlowSheetNotFound
        Exit Function
    End If

    ' search instance name (it should be test name)
    Row = StartRowInFlow
    numOfJobData = 0
    SeqName = ""
    Erase BinName
    Do
        ' check if instance name is empty
        If Cells(Row, ParmCol) = "" Then
            Exit Do
        Else
            ' check if each test function belongs to a sequencer or not
            If Cells(Row, OpCodeCol) = "nop" And Cells(Row, ParmCol) = "SEQ" Then
                SeqName = Cells(Row, TNameCol)
                
                Select Case Left$(SeqName, SeqMatchLength)
                  Case Left$(SeqDcparName, SeqMatchLength), Left$(SeqImageName, SeqMatchLength), _
                        Left$(SeqGradeName, SeqMatchLength), Left$(SeqShirotenName, SeqMatchLength), Left$(SeqMarginName, SeqMatchLength)
                    
                    search_tnum = 1
                    Do While IsEmpty(Cells(Row + search_tnum, TnumCol))
                        search_tnum = search_tnum + 1
                    Loop
                    tnum = Cells(Row + search_tnum, TnumCol)

                  Case Else
                    SeqName = ""
                    tnum = 0
                    outPutMessage "[Error] Unknown sequencer name: '" & Cells(Row, TNameCol) & "'", "Dummy STDF Parameter Check"
                    tlSearchJobData = FormatCheckError
                End Select
            End If
            
            If Cells(Row, OpCodeCol) = "Test" Then
                If SeqName = "" Then
                    outPutMessage "[Error] Each test must belong to a sequencer", "Dummy STDF Parameter Check"
                    tlSearchJobData = FormatCheckError
                End If
            
                If tnum <> 0 And (Not IsEmpty(Cells(Row, TnumCol))) Then
                    'test number must be continuous
                    n = Cells(Row, TnumCol)
                    If tnum <> n Then
                        outPutMessage "[Error] Invalid TNum: " & n & " (must be " & tnum & ")", "Dummy STDF Parameter Check"
                        tlSearchJobData = FormatCheckError
                    End If
                    tnum = tnum + 1
                End If
            End If
            
            If Cells(Row, TnumCol) <> "" _
                    Or (Cells(Row, OpCodeCol) = "nop" And Cells(Row, ParmCol) = "SEQ") Then
                ReDim Preserve JobData(numOfJobData)
                With JobData(numOfJobData)
                    .OpCode = Cells(Row, OpCodeCol)
                    .ParmName = Cells(Row, ParmCol)
                    If .ParmName <> "SEQ" Then
                        .testName = UCase$(Trim$(Cells(Row, TNameCol)))
                    Else
                        .testName = Trim$(Cells(Row, TNameCol))
                    End If
                    .TestNum = Cells(Row, TnumCol)
                    .PassBin = Cells(Row, PassCol)
                    .FailBin = Cells(Row, FailCol)
                    .PassSortBin = Cells(Row, PassSortCol)
                    .FailSortBin = Cells(Row, FailSortCol)
                    .result = Cells(Row, ResultCol)
                    .BinName = Cells(Row, BinNCol)
                    
                    'Check reserved bin number
                    If Not IsEmpty(Cells(Row, FailCol)) And (.FailBin = 0 Or .FailBin = 8 Or .FailBin = 31) Then
                        outPutMessage "[Error] Reserved bin number '" & .FailBin & "' found in TNum " & .TestNum, "DummySTDF Parameter Check"
                        tlSearchJobData = FormatCheckError
                    End If
                    
                    'Check range of pass bin
                    
                    'Check range of fail bin in sequencer dcpar
                    If .OpCode = "Test" And SeqName = SeqDcparName Then
                        If Not IsEmpty(Cells(Row, FailCol)) And (.FailBin < 50 Or 99 < .FailBin) Then
                            outPutMessage "[Error] Fail Bin Number must be 50-99 in TNum " & .TestNum, "Dummy STDF Parameter Check"
                            tlSearchJobData = FormatCheckError
                        End If
                        If Not IsEmpty(Cells(Row, FailSortCol)) And .FailSortBin < 50 Or 99 < .FailSortBin Then
                            outPutMessage "[Error] Fail Sort Number must be 50-99 in TNum " & .TestNum, "Dummy STDF Parameter Check"
                            tlSearchJobData = FormatCheckError
                        End If
                    End If
                    
                End With
                numOfJobData = numOfJobData + 1
            End If
        End If
        Row = Row + 1
    Loop
    
''    MsgBox "flow table OK!"
    
    ' activate instance sheet
    For Each wksht In ActiveWorkbook.Sheets
        If wksht.Name = InstanceSheetName Then
            wksht.Select
            Exit For
        End If
    Next
    
    If wksht.Name <> InstanceSheetName Then
        outPutMessage "[Error] Instance sheet not found"
        tlSearchJobData = InstanceSheetNotFound
        Exit Function
    End If
    
    endFlag = 0
    
    Row = StartRowInInstance
    
    Do
        If Cells(Row, TnameInstanceCol) = "" Then
            Exit Do
        End If
        
        For i = 0 To numOfJobData - 1
            If JobData(i).ParmName = Cells(Row, TnameInstanceCol) Then
                Select Case Cells(Row, TypeInstance)
                Case "Other"
                    If IsNumeric(Cells(Row, OtherLimitBaseCol + 5 * LimitSetIndex + offsetLoLimit)) Then
                        JobData(i).LoLimit = Cells(Row, OtherLimitBaseCol + 5 * LimitSetIndex + offsetLoLimit)
                    Else
                        outPutMessage "[Error] Type mismatch at R" & Row & "C" & OtherLimitBaseCol + 5 * LimitSetIndex + offsetLoLimit & " in '" & InstanceSheetName & "'", "Dummy STDF Parameter Check"
                        tlSearchJobData = TypeMismatchError
                    End If
                    If IsNumeric(Cells(Row, OtherLimitBaseCol + 5 * LimitSetIndex + offsetHiLimit)) Then
                        JobData(i).HiLimit = Cells(Row, OtherLimitBaseCol + 5 * LimitSetIndex + offsetHiLimit)
                    Else
                        outPutMessage "[Error] Type mismatch at R" & Row & "C" & OtherLimitBaseCol + 5 * LimitSetIndex + offsetHiLimit & " in '" & InstanceSheetName & "'", "Dummy STDF Parameter Check"
                        tlSearchJobData = TypeMismatchError
                    End If
                    JobData(i).LimitTyp = Cells(Row, OtherLimitBaseCol + 5 * LimitSetIndex + offsetLimType)
                    JobData(i).Unit = Cells(Row, OtherLimitBaseCol + 5 * LimitSetIndex + offsetUnit)
                    JobData(i).Form = Cells(Row, OtherLimitBaseCol + 5 * LimitSetIndex + offsetForm)
                    createField (i)
                    Exit For
                
                Case "IG-XL Template"
                    Select Case Cells(Row, NameInstance)
                    Case "PinPmu_T"
                        If IsNumeric(Cells(Row, PPMUTempLoLimCol)) Then
                            JobData(i).LoLimit = Cells(Row, PPMUTempLoLimCol)
                        Else
                            outPutMessage "[Error] Type mismatch at R" & Row & "C" & PPMUTempLoLimCol & " in '" & InstanceSheetName & "'", "Dummy STDF Parameter Check"
                            tlSearchJobData = TypeMismatchError
                        End If
                        If IsNumeric(Cells(Row, PPMUTempHiLimCol)) Then
                            JobData(i).HiLimit = Cells(Row, PPMUTempHiLimCol)
                        Else
                            outPutMessage "[Error] Type mismatch at R" & Row & "C" & PPMUTempHiLimCol & " in '" & InstanceSheetName & "'", "Dummy STDF Parameter Check"
                            tlSearchJobData = TypeMismatchError
                        End If
                        JobData(i).LimitTyp = Cells(Row, PPMUTempLimTypeCol)
                        JobData(i).Unit = Cells(Row, PPMUTempUnitCol)
                        JobData(i).Form = Cells(Row, PPMUTempFormCol)
                        createField (i)
                        Exit For
                
                    Case "BoardPmu_T"
                        If IsNumeric(Cells(Row, BPMUTempLoLimCol)) Then
                            JobData(i).LoLimit = Cells(Row, BPMUTempLoLimCol)
                        Else
                            outPutMessage "[Error] Type mismatch at R" & Row & "C" & BPMUTempLoLimCol & " in '" & InstanceSheetName & "'", "Dummy STDF Parameter Check"
                            tlSearchJobData = TypeMismatchError
                        End If
                        If IsNumeric(Cells(Row, BPMUTempHiLimCol)) Then
                            JobData(i).HiLimit = Cells(Row, BPMUTempHiLimCol)
                        Else
                            outPutMessage "[Error] Type mismatch at R" & Row & "C" & BPMUTempHiLimCol & " in '" & InstanceSheetName & "'", "Dummy STDF Parameter Check"
                            tlSearchJobData = TypeMismatchError
                        End If
                        JobData(i).LimitTyp = Cells(Row, BPMUTempLimTypeCol)
                        JobData(i).Unit = Cells(Row, BPMUTempUnitCol)
                        JobData(i).Form = Cells(Row, BPMUTempFormCol)
                        createField (i)
                        Exit For
                
                    Case "CtoAdc_T"
                        JobData(i).LoLimit = 0 ' Cells(Row, CTOAdcTempLoLimCol)
                        JobData(i).HiLimit = 0 ' Cells(Row, CTOAdcTempHiLimCol)
                        JobData(i).LimitTyp = 0 ' Cells(Row, CTOAdcTempLimTypeCol)
                        JobData(i).Unit = "" ' Cells(Row, CTOAdcTempUnitCol)
                        JobData(i).Form = "" 'Cells(Row, CTOAdcTempFormCol)
                        createField (i)
                        Exit For
                
                    Case "CtoPmu_T"
                        If IsNumeric(Cells(Row, CTOPmuTempLoLimCol)) Then
                            JobData(i).LoLimit = Cells(Row, CTOPmuTempLoLimCol)
                        Else
                            outPutMessage "[Error] Type mismatch at R" & Row & "C" & CTOPmuTempLoLimCol & " in '" & InstanceSheetName & "'", "Dummy STDF Parameter Check"
                            tlSearchJobData = TypeMismatchError
                        End If
                        If IsNumeric(Cells(Row, CTOPmuTempHiLimCol)) Then
                            JobData(i).HiLimit = Cells(Row, CTOPmuTempHiLimCol)
                        Else
                            outPutMessage "[Error] Type mismatch at R" & Row & "C" & CTOPmuTempHiLimCol & " in '" & InstanceSheetName & "'", "Dummy STDF Parameter Check"
                            tlSearchJobData = TypeMismatchError
                        End If
                        JobData(i).LimitTyp = Cells(Row, CTOPmuTempLimTypeCol)
                        JobData(i).Unit = Cells(Row, CTOPmuTempUnitCol)
                        JobData(i).Form = Cells(Row, CTOPmuTempFormCol)
                        createField (i)
                        Exit For
                
                    Case "FunctionalPmu_T"
                        JobData(i).LoLimit = 0 ' Cells(Row, FunctionalTempLoLimCol)
                        JobData(i).HiLimit = 0 ' Cells(Row, FunctionalTempHiLimCol)
                        JobData(i).LimitTyp = 0 ' Cells(Row, FunctionalTempLimTypeCol)
                        JobData(i).Unit = "" ' Cells(Row, FunctionalTempUnitCol)
                        JobData(i).Form = "" ' Cells(Row, FunctionalTempFormCol)
                        createField (i)
                        Exit For
                
                    Case "PowerSupply_T"
                        If IsNumeric(Cells(Row, PowerTempLoLimCol)) Then
                            JobData(i).LoLimit = Cells(Row, PowerTempLoLimCol)
                        Else
                            outPutMessage "[Error] Type mismatch at R" & Row & "C" & PowerTempLoLimCol & " in '" & InstanceSheetName & "'", "Dummy STDF Parameter Check"
                            tlSearchJobData = TypeMismatchError
                        End If
                        If IsNumeric(Cells(Row, PowerTempHiLimCol)) Then
                            JobData(i).HiLimit = Cells(Row, PowerTempHiLimCol)
                        Else
                            outPutMessage "[Error] Type mismatch at R" & Row & "C" & PowerTempHiLimCol & " in '" & InstanceSheetName & "'", "Dummy STDF Parameter Check"
                            tlSearchJobData = TypeMismatchError
                        End If
                        JobData(i).LimitTyp = Cells(Row, PowerTempLimTypeCol)
                        JobData(i).Unit = Cells(Row, PowerTempUnitCol)
                        JobData(i).Form = Cells(Row, PowerTempFormCol)
                        createField (i)
                        Exit For
                
                    Case "CtoDac_T"
                        JobData(i).LoLimit = 0 ' Cells(Row, CTODacTempLoLimCol)
                        JobData(i).HiLimit = 0 ' Cells(Row, CTODacTempHiLimCol)
                        JobData(i).LimitTyp = 0 '  Cells(Row, CTODacTempLimTypeCol)
                        JobData(i).Unit = "" ' Cells(Row, CTODacTempUnitCol)
                        JobData(i).Form = "" ' Cells(Row, CTODacTempFormCol)
                        createField (i)
                        Exit For
                
                    Case "MtoMemory_T"
                        JobData(i).LoLimit = 0 ' Cells(Row, MTOMemTempLoLimCol)
                        JobData(i).HiLimit = 0 ' Cells(Row, MTOMemTempHiLimCol)
                        JobData(i).LimitTyp = 0 ' Cells(Row, MTOMemTempLimTypeCol)
                        JobData(i).Unit = "" ' Cells(Row, MTOMemTempUnitCol)
                        JobData(i).Form = "" ' Cells(Row, MTOMemTempFormCol)
                        createField (i)
                        Exit For
                
                    End Select
                End Select
            End If
            JobData(i).Unit = Trim$(JobData(i).Unit)
            JobData(i).Form = Trim$(JobData(i).Form)
        Next i
            
        Row = Row + 1
    Loop
    
''    MsgBox "test instance ok!"
    
    ' activate bin name sheet
    For Each wksht In ActiveWorkbook.Sheets
        If wksht.Name = BnmSheetName Then
            wksht.Select
            Exit For
        End If
    Next
    
    If wksht.Name <> BnmSheetName Then
        outPutMessage "[Error] Bin name sheet not found"
        tlSearchJobData = BinnameSheetNotFound
        Exit Function
    End If
    
    BnmMasterFilename = Trim$(Cells(BnmMasterRow, BnmBinNameCol))
    If BnmMasterFilename = "" Then BnmMasterFilename = DefaultBnmMasterFilename

''    MsgBox "bnm OK!"
    
    Erase SpecialBin
    For i = 0 To MaxBinSize - 1
        If IsEmpty(Cells(BnmBinNameRow + i, BnmBinNumberCol)) Then Exit For
        n = Cells(BnmBinNameRow + i, BnmBinNumberCol)
        If SpecialBin(n) = "" Then
            SpecialBin(n) = Trim$(Cells(BnmBinNameRow + i, BnmBinNameCol))
        Else
            outPutMessage "[Warning] Duplicate bin found in bin name sheet: " & n, "BNM FILE CHECK"
        End If
    Next i
    
''    MsgBox "Bin number OK!"
    
    If fileName <> "" Then
        ' write job data into the designated file
        
        On Error GoTo err2SearchJob
        
        ' open the file
        fp = FreeFile
        Open fileName For Binary As fp
        
        ' write all data into the file
        For i = 0 To numOfJobData - 1
            Put fp, , JobData(i)
        Next i
        
        Close fp
    End If
    
''    MsgBox "file name OK!"
    
    ' validate job data
    If tlJobValidate() = False Then tlSearchJobData = FormatCheckError
    
    Exit Function

err1SearchJob:
    
    tlSearchJobData = MemoryError
    
    Exit Function

err2SearchJob:
    
    tlSearchJobData = CouldNotCreateFile
    
    Exit Function
    
    
End Function

Private Sub createField(ByVal i As Integer)
    Dim ScaleChr As String
    
    
    ' make scale data from Unit, and modify Unit data
    If 1 < Len(JobData(i).Unit) Then
        ScaleChr = Left$(Right$(JobData(i).Unit, 2), 1)
        Select Case ScaleChr
        Case "p"
            JobData(i).ScaleVal = 12
        
        Case "n"
            JobData(i).ScaleVal = 9

        Case "u"
            JobData(i).ScaleVal = 6
            
        Case "m"
            JobData(i).ScaleVal = 3
            
'        Case "%"
'            JobData(i).ScaleVal = 2
            
        Case "K"
            JobData(i).ScaleVal = -3
            
        Case "M"
            JobData(i).ScaleVal = -6
            
        Case "G"
            JobData(i).ScaleVal = -9
            
        End Select
        
        JobData(i).Unit = Right$(JobData(i).Unit, 1)
    Else
        Select Case JobData(i).Unit
            Case "%"
                JobData(i).ScaleVal = 2
            Case Else
                JobData(i).ScaleVal = 0
        End Select
    End If
    
    ' make optflag data from LimitType data
    ' LimitType : 0=none, 1=low, 2=high, 3=both
    Select Case JobData(i).LimitTyp
    Case 0
        JobData(i).OptFlag = "HL"
        
    Case 1
        JobData(i).OptFlag = "H"
        
    Case 2
        JobData(i).OptFlag = "L"
        
    Case 3
        JobData(i).OptFlag = ""
        
    End Select
    
    
    If JobData(i).Form = "" Then
        JobData(i).Form = "%6.3f"
    End If

End Sub

Function tlMoveJobData(ByRef numOfItems As Long, ByRef OpCode() As String, ByRef ParameterName() As String, ByRef TestNum() As Long, testName() As String, LoLimit() As Double, _
                        HiLimit() As Double, LimitTyp() As Integer, ScaleVal() As Integer, Unit() As String, Form() As String, PassBin() As Integer, _
                        FailBin() As Integer, PassSort() As Integer, FailSort() As Integer, BinName() As String, OptFlag() As String, BnmMaster As String) As Long
    Dim i As Integer
    
    
    If numOfJobData = 0 Then
        tlMoveJobData = 1
        Exit Function
    End If
    
    numOfItems = numOfJobData
    
    ReDim OpCode(numOfJobData - 1)
    ReDim ParameterName(numOfJobData - 1)
    ReDim TestNum(numOfJobData - 1)
    ReDim testName(numOfJobData - 1)
    ReDim LoLimit(numOfJobData - 1)
    ReDim HiLimit(numOfJobData - 1)
    ReDim LimitTyp(numOfJobData - 1)
    ReDim ScaleVal(numOfJobData - 1)
    ReDim Unit(numOfJobData - 1)
    ReDim Form(numOfJobData - 1)
    ReDim PassBin(numOfJobData - 1)
    ReDim FailBin(numOfJobData - 1)
    ReDim PassSort(numOfJobData - 1)
    ReDim FailSort(numOfJobData - 1)
    ReDim OptFlag(numOfJobData - 1)
    
    For i = 0 To numOfJobData - 1
        OpCode(i) = JobData(i).OpCode
        ParameterName(i) = JobData(i).ParmName
        TestNum(i) = JobData(i).TestNum
        testName(i) = JobData(i).testName
        LoLimit(i) = JobData(i).LoLimit
        HiLimit(i) = JobData(i).HiLimit
        LimitTyp(i) = JobData(i).LimitTyp
        ScaleVal(i) = JobData(i).ScaleVal
        Unit(i) = JobData(i).Unit
        Form(i) = JobData(i).Form
        PassBin(i) = JobData(i).PassBin
        FailBin(i) = JobData(i).FailBin
        PassSort(i) = JobData(i).PassSortBin
        FailSort(i) = JobData(i).FailSortBin
        OptFlag(i) = JobData(i).OptFlag
    Next i
    
    Erase BinName
    For i = 0 To MaxBinSize - 1
        BinName(i) = SpecialBin(i)
    Next i
    
    BnmMaster = BnmMasterFilename
    
    ' Normal end
    tlMoveJobData = 0
    
End Function

Function tlJobValidate() As Boolean
    Dim status As Long
    Dim flag As Boolean
    Dim valid As Boolean
    Dim found As Boolean
    Dim tnum As Integer
    Dim i As Integer
    Dim j As Integer
    Dim tname(5000) As String '2000->5000 '08/11/25
    Dim TnameNum As Integer
    Dim SeqName As String
    
    valid = True
    
    'Check usage of "nop"
    For i = 0 To numOfJobData - 1
        'must be used with 'SEQ'
        If JobData(i).OpCode = "nop" And JobData(i).ParmName <> "SEQ" Then
            outPutMessage "[Error] 'nop' can be used only with 'SEQ'", "Dummy STDF Parameter Check"
'            valid = False
        End If
    Next i
    
    'Check the range of "Bin Number"
    For i = 0 To numOfJobData - 1
        With JobData(i)
            If .PassBin < 0 Or .PassBin > 99 Then
                outPutMessage "[Error] Bad 'Pass Bin Number':" & .PassBin & " at TNum " & .TestNum, "Dummy STDF Parameter Check"
                valid = False
            End If
            If .FailBin < 0 Or .FailBin > 99 Then
                outPutMessage "[Error]Bad 'Fail Bin Number':" & .FailBin & " at TNum " & .TestNum, "Dummy STDF Parameter Check"
                valid = False
            End If
        End With
    Next i
    
    'Check usage of "Test" - this has already been performed in tlSearchJobData()
   
    'Check usage of limit columns - this has already been performed in tlSearchJobData()
    
    'Check usage of "Result" column, "Bin Number" and "Sort Number"
    
    For i = 0 To numOfJobData - 1
        With JobData(i)
            '"Bin Number" and "Sort Number" must be null if "Result" is "None"
            If .result = "None" And _
                    (.PassBin <> 0 Or .FailBin <> 0 Or .PassSortBin <> 0 Or .FailSortBin <> 0) Then
                outPutMessage "[Error] Bin and Sort Number must be empty if result is 'None'", "Dummy STDF Parameter Check"
                valid = False
            End If
        End With
    Next i
    
    'Check test names (Tname)
    
    
    Erase tname
    TnameNum = 0
    For i = 0 To numOfJobData - 1
        With JobData(i)
            If .OpCode = "Test" Then
                'Check null test name
                If .testName <> "" Then
                    'Check space character
                    If InStr(.testName, SpaceChar) <> 0 Then
                        outPutMessage "[Error] Space character found in Tname: '" & .testName & "'", "Dummy STDF Parameter Check"
                        valid = False
                    End If
                    
                    'Check duplication
                    found = False
                    For j = 0 To TnameNum
                        If tname(j) = Left$(.testName, TnameMatchLength) Then
                            found = True
                            outPutMessage "[Error] Duplicate Tname: '" & tname(j) & "' and '" & .testName & "'", "Dummy STDF Parameter Check"
                            valid = False
                        End If
                    Next j
                    If Not found Then
                        tname(TnameNum) = Left$(.testName, TnameMatchLength)
                        TnameNum = TnameNum + 1
                    End If
                End If
'
'                'Check null bin name and space character
'                If .BinName <> "" And InStr(.BinName, SpaceChar) <> 0 Then
'                    OutputMessage "[Error] Space character found in bin name: '" & .BinName & "'", "Dummy STDF Parameter Check"
'                    valid = False
'                End If
            End If
        End With
    Next i
    
    'Check unit
'    For i = 0 To numOfJobData - 1
'        With JobData(i)
'            'Check space character
'            If InStr(.Unit, SpaceChar) <> 0 Then
'                OutputMessage "[Error] Space found in Unit: '" & .Unit & "'", "Dummy STDF Parameter Check"
'                valid = False
'            End If
'        End With
'    Next i
    
    'Check format
'    For i = 0 To numOfJobData - 1
'        With JobData(i)
'            'Check space character
'            If InStr(.Form, SpaceChar) <> 0 Then
'                OutputMessage "[Error] Space found in Format: '" & .Form & "'", "Dummy STDF Parameter Check"
'                valid = False
'            End If
'        End With
'    Next i
    
    tlJobValidate = valid
End Function

Private Sub DummyTest()
    Dim status As Long
    
    'for debug
    status = tlSearchJobData("", "", "")

End Sub


