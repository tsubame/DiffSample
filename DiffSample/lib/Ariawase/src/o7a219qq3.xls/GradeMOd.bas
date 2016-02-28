Attribute VB_Name = "GradeMOd"
Option Explicit

'========= GRADE VARIABLES ===========
Public Ng_test(nSite) As Double
Public Watchs(nSite) As Double
Public Watcht(nSite) As Double
Public Watchc(nSite) As Double

Public Now_Day(nSite) As Double
Public Now_Time(nSite) As Double
'====================================

Public GradeLastBin As Integer
Public GradeSortBin As Integer
Private watchwaferno As Integer
Private watchcount As Integer

Private Function ng_test_f() As Long

    '===== GO TO GRADE FUNCTION ===
    Call GradeOn
    '==============================
    Call test(Ng_test)

End Function

Private Function watchs_f() As Long  '2012/11/16 175JobMakeDebug Arikawa

    Dim site As Long

    For site = 0 To nSite
        Watchs(site) = site
    Next

    Call test(Watchs)

End Function

Private Function watchc_f() As Long

    Dim site As Long
    
    If WaferNo <> "" Then
        If watchwaferno = 0 Then
            watchwaferno = CInt(WaferNo)
            watchcount = watchcount + 1
        Else
            If watchwaferno = CInt(WaferNo) Then
                watchcount = watchcount + 1
            Else
                watchwaferno = CInt(WaferNo)
                watchcount = 1
            End If
        End If
    End If
    
    For site = 0 To nSite
        Watchc(site) = watchcount
    Next

    Call test(Watchc)

End Function

Public Function watcht_f() As Long
    
    Dim site As Long
    Dim testtime As Double
    
    Call stopTime(testtime)
    
    For site = 0 To nSite
        Watcht(site) = testtime
    Next
    Call test(Watcht)

End Function

Private Function now_day_f() As Long
        
    Dim site As Long
    
    For site = 0 To nSite
        Now_Day(site) = Format$(Now, "yymmdd")
    Next site
                
    Call test(Now_Day)

End Function

Private Function now_time_f() As Long
    
    Dim site As Long
    
    For site = 0 To nSite
        Now_Time(site) = Format$(Now, "hhmmss")
    Next site
                            
    Call test(Now_Time)

End Function

Public Function rankng_f() As Long

    Call SiteCheck
    
'========= For Human Error 09/03/16 ==========
    If Rank_Judge = False Then
        MsgBox "ERROR Detected in RANK Proccess!!!"
        MsgBox "PMC STOP!!!"
        gFlg_StopPMC = True 'EeeAuto
    End If
'========= For Human Error 09/03/16 ==========

    Call test(Rank_ng)

End Function

Public Function g2ngbn_f() As Double

    Dim site As Long

    Call test(G2ngbn)

End Function

Public Function g2_flg_f() As Double

    Dim site As Long

    Call test(G2_flg)

End Function

Private Function g2rank_f() As Double

    Dim site As Long
    GradeLastBin = 2
    GradeSortBin = 2

'    /*** 17/Mar/02 takayama append
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If TheExec.sites.site(site).BinNumber <> -1 Then
                LastBin(site) = TheExec.sites.site(site).FirstBinNumber
                SortBin(site) = TheExec.sites.site(site).FirstSortNumber
            Else
                LastBin(site) = GradeLastBin
                SortBin(site) = GradeSortBin
            End If
        End If
    Next site
'    /*** 17/Mar/02 takayama append

    Call test(G2rank)

End Function

Public Function g3ngbn_f() As Double

    Dim site As Long

    Call test(G3ngbn)

End Function

Public Function g3_flg_f() As Double

    Dim site As Long

    Call test(G3_flg)

End Function

Private Function g3rank_f() As Double

    Dim site As Long
    GradeLastBin = 3
    GradeSortBin = 3

'    /*** 17/Mar/02 takayama append
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If TheExec.sites.site(site).BinNumber <> -1 Then
                LastBin(site) = TheExec.sites.site(site).FirstBinNumber
                SortBin(site) = TheExec.sites.site(site).FirstSortNumber
            Else
                LastBin(site) = GradeLastBin
                SortBin(site) = GradeSortBin
            End If
        End If
    Next site
'    /*** 17/Mar/02 takayama append

    Call test(G3rank)

End Function

Public Function g4ngbn_f() As Double

    Dim site As Long

    Call test(G4ngbn)

End Function

Public Function g4_flg_f() As Double

    Dim site As Long

    Call test(G4_flg)

End Function

Private Function g4rank_f() As Double

    Dim site As Long
    GradeLastBin = 4
    GradeSortBin = 4

'    /*** 17/Mar/02 takayama append
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If TheExec.sites.site(site).BinNumber <> -1 Then
                LastBin(site) = TheExec.sites.site(site).FirstBinNumber
                SortBin(site) = TheExec.sites.site(site).FirstSortNumber
            Else
                LastBin(site) = GradeLastBin
                SortBin(site) = GradeSortBin
            End If
        End If
    Next site
'    /*** 17/Mar/02 takayama append

    Call test(G4rank)

End Function

Public Function g5ngbn_f() As Double

    Dim site As Long

    Call test(G5ngbn)

End Function

Public Function g5_flg_f() As Double

    Dim site As Long

    Call test(G5_flg)

End Function

Private Function g5rank_f() As Double

    Dim site As Long
    GradeLastBin = 5
    GradeSortBin = 5

'    /*** 17/Mar/02 takayama append
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If TheExec.sites.site(site).BinNumber <> -1 Then
                LastBin(site) = TheExec.sites.site(site).FirstBinNumber
                SortBin(site) = TheExec.sites.site(site).FirstSortNumber
            Else
                LastBin(site) = GradeLastBin
                SortBin(site) = GradeSortBin
            End If
        End If
    Next site
'    /*** 17/Mar/02 takayama append

    Call test(G5rank)

End Function
