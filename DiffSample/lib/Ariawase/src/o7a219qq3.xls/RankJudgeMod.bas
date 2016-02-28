Attribute VB_Name = "RankJudgeMod"
Option Explicit

Type RankType
    tnum As Integer                 'Test Number
    tname As String                 'Test Name
    BinNo As Integer                'Bin No.
    Specs() As Double               'Rank Spec
    HiLoLim_Flg() As Integer        'High/Low Limit flag
    TestResult(nSite) As Double     'Test Result
    Srank() As Integer              'SRank
End Type

Public RankData() As RankType       'Rank Data
Public SRankData() As Integer       'SRank Data

Public Max_test_num As Integer      'Number of Test
Public Max_rank_num As Integer      'Number of Rank

Public SRnk_flg As Boolean          'SRank Judge Do or Not

Public S_rank(nSite) As Double      'For FC Test

Public Rank_ng(nSite) As Double
Public G2ngbn(nSite) As Double
Public G3ngbn(nSite) As Double
Public G4ngbn(nSite) As Double
Public G5ngbn(nSite) As Double
Public G2_flg(nSite) As Double
Public G3_flg(nSite) As Double
Public G4_flg(nSite) As Double
Public G5_flg(nSite) As Double
Public G2rank(nSite) As Double
Public G3rank(nSite) As Double
Public G4rank(nSite) As Double
Public G5rank(nSite) As Double

Public Rselect2(nSite) As Double
Public Rselect3(nSite) As Double
Public Rselect4(nSite) As Double
Public Rselect5(nSite) As Double

Public gFlg_StopPMC As Boolean

Sub RankInit()

    Dim Shts As Worksheet           'For Work Sheet Check
    
    Dim basePoint As Variant        'For Base Point Serch
    Dim EndPoint As Variant         'For End Point Serch
    
    Dim EPRow As Long               'End Point Row
    Dim EPColumn As Integer         'End Point Column
    
    Dim TEnd_Row As Long            'Test End Row
    
    Dim Para_Name As String         'Parameter Name
    Dim TIns_Unit As String         'Unit in Test Instance
    Dim Unit As Double              'Unit in rank Sheet
    
    Dim Find_Label As String        'For Serch Label
    Dim Cell_Position As String     'For Serch Cell Position
    
    Dim Srnk_Cnt_Flg As Boolean     'Check SRANK Judge do or not
    
    Dim i As Integer
    Dim j As Integer
    
    gFlg_StopPMC = False            'Stop PMC
    
    'Check "rank_sheet" Sheet exist in JOB or not
    On Error GoTo ErrorDetected
    Set Shts = Sheets("rank_sheet")
    Set Shts = Sheets("Flow Table")
    
    'Get Test End Row For Serch TName, tnum and BIN No.
    Worksheets("Flow Table").Select
    
    Cell_Position = "G4:G65536"
    Find_Label = "set-device"
    Set basePoint = Range(Cell_Position).Find(Find_Label, , , xlWhole)
    
    If basePoint Is Nothing Then
        MsgBox "[set-device] is Nothing in Flow Table!!!"
        GoTo ErrorDetected
    End If
    
    TEnd_Row = basePoint.Row
    
    Worksheets("rank_sheet").Select
    
'Serch "RANK"
'If "RANK" was Nothing, ERROR
    Find_Label = "RANK"
    
    If Cells(2, 2) <> Find_Label Then
        MsgBox "Please check [RANK] is in [B2] or not !!!"
        GoTo ErrorDetected
    End If


'Serch "TNAME"
'if "TNAME" was Nothing, ERROR
    Cell_Position = "B8"
    Find_Label = "TNAME"
    
    Set basePoint = Range(Cell_Position).Find(Find_Label, , , xlWhole)

    If basePoint Is Nothing Then
        MsgBox "Please check [TNAME] is in [B8] or not !!!"
        GoTo ErrorDetected
    End If

'test count
    Set EndPoint = Range(Cell_Position).End(xlDown)
    EPRow = EndPoint.Row
        
    If EPRow = 65536 Then
        Max_test_num = 0
    Else
        Max_test_num = EPRow - basePoint.Row
        
        Cell_Position = "B" & EPRow
        Set EndPoint = Range(Cell_Position).End(xlDown)
        
        If EndPoint.Row <> 65536 Then
            MsgBox "[Space] Found in TNAME!!!"
            GoTo ErrorDetected
        End If
    End If
    
'rank count
    j = 0
    While Cells(2, 5 + j * 2) <> ""
        j = j + 1
    Wend
    
    Max_rank_num = j
    
    Set EndPoint = Range(Cells(2, 5 + j * 2), Cells(2, 5 + j * 2)).End(xlToRight)
    
    If EndPoint.Column <> 256 Then
        MsgBox "[Space] Found in Rank No!!!"
        GoTo ErrorDetected
    End If
    
're define
    ReDim SRankData(5, Max_rank_num)
    ReDim RankData(Max_test_num)
    
    For i = 0 To Max_test_num
        ReDim RankData(i).HiLoLim_Flg(Max_rank_num)
        ReDim RankData(i).Specs(Max_rank_num * 2)
    Next i
    
'Start Data reading
    'RANK No. reading
    For j = 0 To Max_rank_num - 1
        If Cells(2, 5 + (j * 2)) > 9 Then
            MsgBox "[Rank No.] is over [9] or Wrong! Please check!!!"
            GoTo ErrorDetected
        Else
            RankData(0).Specs(j * 2) = Cells(2, 5 + (j * 2))
        End If
    Next j
        
    'TName reading
    For i = 1 To Max_test_num
        RankData(i).tname = Cells(8 + i, 2)
        
        'unit check
        If Cells(8 + i, 3) = "" Then
            MsgBox "[Space] Found!!! Please check Unit!!!"
            GoTo ErrorDetected
        End If
        
        Cell_Position = "I4:I" & TEnd_Row
        Set basePoint = Sheets("Flow Table").Range(Cell_Position).Find(RankData(i).tname, , , xlWhole)
        
        If basePoint Is Nothing Then
            MsgBox " Test Name = " & RankData(i).tname & " is Wrong!!!"
            GoTo ErrorDetected
        End If
        
        Para_Name = Sheets("Flow Table").Cells(basePoint.Row, basePoint.Column - 1)
        
        Cell_Position = "B4:B" & TEnd_Row
        Set basePoint = Sheets("Test Instances").Range(Cell_Position).Find(Para_Name, , , xlWhole)
        If basePoint Is Nothing Then
            MsgBox " Test Name = " & RankData(i).tname & " is Wrong!!!"
            GoTo ErrorDetected
        End If
        
        TIns_Unit = Sheets("Test Instances").Cells(basePoint.Row, basePoint.Column + 15)
        
        If TIns_Unit <> "" Then
            If Cells(8 + i, 3) <> TIns_Unit Then
                MsgBox "Unit is Wrong!!! Please check Unit!!!"
                GoTo ErrorDetected
            End If
        Else
            If Cells(8 + i, 3) <> "-" Then
                MsgBox "[-] Not Found!!! Please check Unit!!!"
                GoTo ErrorDetected
            End If
        End If
    
        'Unit reading
        Select Case Cells(8 + i, 3)
            Case "V", "A", "W"
                Unit = 1
            Case "mV", "mA", "mW"
                Unit = 0.001
            Case "uV", "uA", "uW"
                Unit = 0.000001
            Case "nV", "nA", "nW"
                Unit = 0.000000001
            Case "%"
                Unit = 0.01
            Case "Kr"
                Unit = 1000
            Case "r"
                Unit = 1
            Case "db"
                Unit = 1
            Case "S"
                Unit = 1
            Case "-"
                Unit = 1
            Case Else
                'error message
                MsgBox "[" & Cells(8 + i, 3) & "] Not Found!!! Please check Unit!!!"
                GoTo ErrorDetected
        End Select
        
        'BIN No. reading
        If Cells(8 + i, 4) = "" Then
            MsgBox "[Space] Found!!! Please check BIN No.!!!"
            GoTo ErrorDetected
        End If
        
        RankData(i).BinNo = Cells(8 + i, 4)
    
        'Specs reading
        'if "-", then skip
        For j = 0 To Max_rank_num * 2 - 1
            If Cells(8 + i, 5 + j) = "" Then
                MsgBox "[Space] Found!!! Please change [Space] to [-] !!!"
                GoTo ErrorDetected
            ElseIf Cells(8 + i, 5 + j) <> "-" Then
                RankData(i).Specs(j) = val(Cells(8 + i, 5 + j)) * Unit
                
                If Cells(2, 5 + j) <> "" Then
                    RankData(i).HiLoLim_Flg(j / 2) = 1
                ElseIf Cells(2, 5 + j) = "" Then
                    RankData(i).HiLoLim_Flg((j - 1) / 2) = RankData(i).HiLoLim_Flg((j - 1) / 2) + 2
                End If
            End If
        Next j
    Next i

'Serch "SRANK"
'if "SRANK" was Nothing, ERROR
    Find_Label = "SRANK" & Chr$(10) & "(FC only)"
    
    If Cells(3, 2) <> Find_Label Then
        MsgBox "Please check [SRANK] is in [B3] or not !!!"
        GoTo ErrorDetected
    End If
    
    'SRANK TEST  DOFTrue,  DO NOTFFalse
    SRnk_flg = False

    'Srank reading
    For i = 1 To 5
        For j = 0 To Max_rank_num - 1
            'If "-" , then 999
            If Cells(2 + i, 5 + j * 2) = "-" Then
                SRankData(i, j) = 999
            ElseIf Cells(2 + i, 5 + j * 2) >= 0 And Cells(2 + i, 5 + j * 2) <= 5 Then
                SRankData(i, j) = Cells(2 + i, 5 + j * 2)
                SRnk_flg = True
            Else
                MsgBox "[Space] Found!!! Please change [Space] to [-] !!!"
                GoTo ErrorDetected
            End If
        Next j
    Next i

    Worksheets("Flow Table").Select
    
    'Read tnum From FlowTable
    For i = 1 To Max_test_num
        'Serch same TName
        Cell_Position = "I4:I" & TEnd_Row
        Set basePoint = Range(Cell_Position).Find(RankData(i).tname, , , xlWhole)
        'Same TName found, then get tnum
        If basePoint Is Nothing Then
            'if Same TName was Nothing, then ERROR!!
            MsgBox " Test Name = " & RankData(i).tname & " is Wrong!!!"
            GoTo ErrorDetected
        Else
            RankData(i).tnum = Cells(basePoint.Row, basePoint.Column + 1)
        End If
    Next i

'BIN No. check
    For i = 1 To Max_test_num
        'Serch same BIN No.
        Cell_Position = "L4:L" & TEnd_Row
        Set basePoint = Range(Cell_Position).Find(RankData(i).BinNo)
        If basePoint Is Nothing Then
            'if Bin No. was Nothing, ERROR!!
            MsgBox "Bin No. " & RankData(i).BinNo & " is Nothing in FrowTable!!!"
            GoTo ErrorDetected
        End If
    Next i

Exit Sub

ErrorDetected:
    MsgBox "ERROR!!!!!!"
    gFlg_StopPMC = True

End Sub

Public Function Rank_Judge() As Boolean

    Dim Fst_Fail_Bin As Integer         'Keep First Fail Bin No.
    Dim rank(nSite) As String           'Rank

    Dim JudgeTest_Flg As Boolean        'Judge Test or not
    Dim Rank_Flg As Boolean             'Rank Determined or not
    
    Dim site As Long
    Dim i As Integer
    Dim j As Integer
    Dim si As Integer
    
    Rank_Judge = False
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then

            Fst_Fail_Bin = 0
    
            For j = 0 To Max_rank_num - 1
                
                Rank_Flg = True
                JudgeTest_Flg = False
                
                'SRANK JUDGE
                If SRnk_flg = True Then
                    For si = 1 To 5
                        If S_rank(site) = SRankData(si, j) Then
                            JudgeTest_Flg = True
                            Exit For
                        End If
                    Next si
                    
                    If JudgeTest_Flg = False And Fst_Fail_Bin = 0 Then
                        Fst_Fail_Bin = S_rank(site)
                    End If
                Else
                    JudgeTest_Flg = True
                End If
                
                'TEST JUDGE
                If JudgeTest_Flg = True Then
                    For i = 1 To Max_test_num
                        'Low Spec judge
                        If (RankData(i).HiLoLim_Flg(j) = 1) Or (RankData(i).HiLoLim_Flg(j) = 3) Then
                            If RankData(i).TestResult(site) < RankData(i).Specs(j * 2) Then
                                'Get First FAIL BIN
                                If Fst_Fail_Bin = 0 Then
                                    Fst_Fail_Bin = RankData(i).BinNo
                                End If
                                
                                Rank_Flg = False
                                
                                Exit For
                            End If
                        End If
                        
                        'High Spec judge
                        If (RankData(i).HiLoLim_Flg(j) = 2) Or (RankData(i).HiLoLim_Flg(j) = 3) Then
                            If RankData(i).TestResult(site) > RankData(i).Specs(j * 2 + 1) Then
                                'Get First FAIL BIN
                                If Fst_Fail_Bin = 0 Then
                                    Fst_Fail_Bin = RankData(i).BinNo
                                End If
                                
                                Rank_Flg = False
                                
                                Exit For
                            End If
                        End If
                    Next i
                    
                    
                    
                    'If All Test Cleared, Rank Determined
                    If Rank_Flg = True Then
                        rank(site) = RankData(0).Specs(j * 2)
                        Exit For
                    End If
                End If
            Next j
                
            'If All Test Not Cleared, Rank NG
            If Rank_Flg = False Then
                rank(site) = "NG"
            End If
    
            Select Case rank(site)
                'rank 1
                Case "1"
                
                'rank 2
                Case "2"
                    G2ngbn(site) = Fst_Fail_Bin
                    G2_flg(site) = Fst_Fail_Bin
                    G2rank(site) = Fst_Fail_Bin
                    
                    Rselect2(site) = Fst_Fail_Bin
                
                'rank 3
                Case "3"
                    G3ngbn(site) = Fst_Fail_Bin
                    G3_flg(site) = Fst_Fail_Bin
                    G3rank(site) = Fst_Fail_Bin
                    
                    Rselect3(site) = Fst_Fail_Bin
                
                'rank 4
                Case "4"
                    G4ngbn(site) = Fst_Fail_Bin
                    G4_flg(site) = Fst_Fail_Bin
                    G4rank(site) = Fst_Fail_Bin
                    
                    Rselect4(site) = Fst_Fail_Bin
                
                'rank 5
                Case "5"
                    G5ngbn(site) = Fst_Fail_Bin
                    G5_flg(site) = Fst_Fail_Bin
                    G5rank(site) = Fst_Fail_Bin
                    
                    Rselect5(site) = Fst_Fail_Bin
                
                'rank NG
                Case "NG"
                    Rank_ng(site) = Fst_Fail_Bin
                
                Case Else
                    Exit Function
            
            End Select
        End If
    Next site

    Rank_Judge = True

End Function

Sub get_testresult(result As Variant)

    Dim lngTestNumber As Integer
    Dim i As Integer

    Dim site As Long
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            lngTestNumber = TheExec.sites.site(site).TestNumber
        End If
    Next site

    For i = 1 To Max_test_num
        If lngTestNumber = RankData(i).tnum Then
            For site = 0 To nSite
                If TheExec.sites.site(site).Active = True Then
                    RankData(i).TestResult(site) = result(site)
                End If
            Next

            Exit For
        End If
    Next

End Sub

Sub RankDataClear()

    Dim i As Integer
    Dim site As Long

    For i = 1 To Max_test_num
        For site = 0 To nSite
            RankData(i).TestResult(site) = 0
        Next
    Next

End Sub
