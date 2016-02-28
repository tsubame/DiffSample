Attribute VB_Name = "Shmoo_Ascii2"
Option Explicit

' Shmoo ASCII module
' Call from Shmoo Interpose function
' Start     : InitShmooAscii
' PostPoint : CheckShmooResult
' End       : OutputShmooAscii

Private DataObj As RtaDataObj
Private Type POINTDAT
    Xval As Double
    Yval As Double
    result As Long
End Type
Private Type SHMOODAT
    xunit As String
    XParameter As String
    yunit As String
    YParameter As String
    Data() As POINTDAT
End Type
Private ShmooData() As SHMOODAT
Private XParameters() As Double
Private YParameters() As Double
Private Flg_conv As Boolean   ' 2009/03/04 Updata Ozawa

'' You need to call this function in the Start Interpose
'Public Function InitShmooAscii(argc As Long, argv() As String) As Long
'
'    Set DataObj = Nothing
'    Erase ShmooData
'    ReDim ShmooData(TheExec.Sites.ExistingCount - 1)
'    Erase XParameters
'    Erase YParameters
'
'End Function

' You need to call this function in the Start Interpose
Public Function InitShmooAscii2(argc As Long, argv() As String) As Long

    Set DataObj = Nothing
    Erase ShmooData
    ReDim ShmooData(TheExec.sites.ExistingCount - 1)
    Erase XParameters
    Erase YParameters
    Flg_conv = False    ' 2009/03/04 Updata Ozawa Conversion Flag Initialize
    
End Function

'' You need to call this function in the PostPoint Interpose
'Public Function CheckShmooResult(argc As Long, argv() As String) As Long
'    Dim site As Integer
'
'    If DataObj Is Nothing Then
'        Set DataObj = TheExec.DevChar.ActiveDataObject
'        For site = 0 To UBound(ShmooData)
'            ReDim ShmooData(site).Data(0)
'            With DataObj
'                ShmooData(site).xunit = .xunit
'                ShmooData(site).yunit = .yunit
'                ShmooData(site).XParameter = .XParameter
'                ShmooData(site).YParameter = .YParameter
'                ReDim XParameters(.XDim)
'                ReDim YParameters(.YDim)
'            End With
'        Next site
'    End If
'
'    For site = 0 To DataObj.SiteDim           ' do each site
'        DataObj.site = site
'        ReDim Preserve ShmooData(site).Data(UBound(ShmooData(site).Data) + 1)
'        With ShmooData(site).Data(UBound(ShmooData(site).Data))
'            .Xval = DataObj.Xval
'            .Yval = DataObj.Yval
'            .Result = DataObj.PtResult
'        End With
'
'        XParameters(DataObj.X) = DataObj.Xval
'        YParameters(DataObj.Y) = DataObj.Yval
'    Next site
'
'End Function

' You need to call this function in the PostPoint Interpose
Public Function CheckShmooResult2(argc As Long, argv() As String) As Long
    Dim site As Integer
    
    If argc <> 0 Then                      ' 2009/03/04 Updata Ozawa
        If LCase(argv(0)) = "conv" Then    ' 2009/03/04 Updata Ozawa
            Flg_conv = True                ' Conversion flag
        End If                             ' 2009/03/04 Updata Ozawa
    End If
    If DataObj Is Nothing Then
        Set DataObj = TheExec.DevChar.ActiveDataObject
        For site = 0 To UBound(ShmooData)
            ReDim ShmooData(site).Data(0)
            If Flg_conv = False Then  ' 2009/03/04 Updata Ozawa
                With DataObj
                    ShmooData(site).xunit = .xunit
                    ShmooData(site).yunit = .yunit
                    ShmooData(site).XParameter = .XParameter
                    ShmooData(site).YParameter = .YParameter
                    ReDim XParameters(.XDim)
                    ReDim YParameters(.YDim)
                End With
            ElseIf Flg_conv = True Then  ' 2009/03/04 Updata Ozawa
                With DataObj
                    ShmooData(site).xunit = .yunit
                    ShmooData(site).yunit = .xunit
                    ShmooData(site).XParameter = .YParameter
                    ShmooData(site).YParameter = .XParameter
                    ReDim XParameters(.YDim)
                    ReDim YParameters(.XDim)
                End With
            End If
        Next site
    End If

    For site = 0 To DataObj.SiteDim           ' do each site
        DataObj.site = site
        ReDim Preserve ShmooData(site).Data(UBound(ShmooData(site).Data) + 1)
        With ShmooData(site).Data(UBound(ShmooData(site).Data))
            .Xval = DataObj.Xval
            .Yval = DataObj.Yval
            .result = DataObj.PtResult
        End With
        
        If Flg_conv = False Then     ' 2009/03/04 Updata Ozawa
            XParameters(DataObj.x) = DataObj.Xval
            YParameters(DataObj.y) = DataObj.Yval
        ElseIf Flg_conv = True Then  ' 2009/03/04 Updata Ozawa
            XParameters(DataObj.y) = DataObj.Yval
            YParameters(DataObj.x) = DataObj.Xval
        End If
    Next site

'If DataObj.Y = 49 Then Stop
End Function


' You need to call this function in the End Interpose
Public Function OutputShmooAscii2(argc As Long, argv() As String) As Long
    Dim LastBurst As String
    Dim Isgroup As Boolean
    Dim LastLabel As String
    Dim i As Integer
    Dim site As Long
    
    Call TheHdw.Raw.patvba.PatternReadLastStart(LastBurst, Isgroup, LastLabel)
    
    OutToLog ""
    OutToLog "****  " & DataObj.XParameter & " vs " & DataObj.YParameter & " Shmoo Plot  ****"

    OutToLog "    Date          : " & Format(DataObj.DateStamp, "mm/dd/yy")
    OutToLog "    Start Time    : " & Format(DataObj.StartTime, "hh:mm:ss")
    OutToLog "    Stop Time     : " & Format(Now, "hh:mm:ss") 'Format(DataObj.StopTime, "hh:mm:ss")
    OutToLog "    Program       : " & DataObj.ProgName
    OutToLog "    Job           : " & DataObj.JobName
    OutToLog "    Channel Map   : " & DataObj.ChanMap
    OutToLog "    Part          : " & DataObj.Part
    OutToLog "    Test Instance : " & DataObj.testName
    OutToLog "    AC Context    : " & DataObj.accontext
    OutToLog "    DC Context    : " & DataObj.dccontext
    OutToLog "    setup         : " & DataObj.Setup
    OutToLog "    Comment       : " & DataObj.Comment
    OutToLog "    Pattern       : " & LastBurst & IIf(LastLabel = "", "", " at " & LastLabel)
    OutToLog ""
    
    Dim x As Long, y As Long
    Dim xunit As String, yunit As String
    xunit = "": yunit = ""
    Dim xmax As Integer, ymax As Integer
    xmax = 0: ymax = 0
    
        ReDim XValues(UBound(XParameters)) As String
        ReDim YValues(UBound(YParameters)) As String
        For x = 0 To UBound(XParameters)
            XValues(x) = FormatValue(XParameters(x))
            If xmax < Len(XValues(x)) Then xmax = Len(XValues(x))
            If XParameters(x) <> 0 And xunit = "" Then xunit = Right(XValues(x), 1)
        Next x
        For y = 0 To UBound(YParameters)
            YValues(y) = FormatValue(YParameters(y))
            If ymax < Len(YValues(y)) Then ymax = Len(YValues(y))
            If YParameters(y) <> 0 And yunit = "" Then yunit = Right(YValues(y), 1)
        Next y
        
        For x = 0 To UBound(XValues)
            If XParameters(x) = 0 Then XValues(x) = Trim(XValues(x)) & " " & xunit
            If Len(XValues(x)) < xmax Then XValues(x) = Space(xmax - Len(XValues(x))) & XValues(x)
        Next x
        For y = 0 To UBound(YValues)
            If YParameters(y) = 0 Then YValues(y) = Trim(YValues(y)) & " " & yunit
            If Len(YValues(y)) < ymax Then YValues(y) = Space(ymax - Len(YValues(y))) & YValues(y)
        Next y
        
        For site = 0 To DataObj.SiteDim
            If Flg_conv = False Then
                If TheExec.sites.site(site).Active Then Call DisplayShmooAscii(XValues, YValues, site)
            ElseIf Flg_conv = True Then
                If TheExec.sites.site(site).Active Then Call DisplayShmooAscii_conv(XValues, YValues, site)
            End If
        Next site

End Function


'' You need to call this function in the End Interpose
'Public Function OutputShmooAscii(argc As Long, argv() As String) As Long
'    Dim LastBurst As String
'    Dim Isgroup As Boolean
'    Dim LastLabel As String
'    Dim i As Integer
'    Dim site As Long
'
'    Call TheHdw.Raw.patvba.PatternReadLastStart(LastBurst, Isgroup, LastLabel)
'
'    OutToLog ""
'    OutToLog "****  " & DataObj.XParameter & " vs " & DataObj.YParameter & " Shmoo Plot  ****"
'
'    OutToLog "    Date          : " & Format(DataObj.DateStamp, "mm/dd/yy")
'    OutToLog "    Start Time    : " & Format(DataObj.StartTime, "hh:mm:ss")
'    OutToLog "    Stop Time     : " & Format(Now, "hh:mm:ss") 'Format(DataObj.StopTime, "hh:mm:ss")
'    OutToLog "    Program       : " & DataObj.ProgName
'    OutToLog "    Job           : " & DataObj.JobName
'    OutToLog "    Channel Map   : " & DataObj.ChanMap
'    OutToLog "    Part          : " & DataObj.Part
'    OutToLog "    Test Instance : " & DataObj.testName
'    OutToLog "    AC Context    : " & DataObj.accontext
'    OutToLog "    DC Context    : " & DataObj.dccontext
'    OutToLog "    setup         : " & DataObj.Setup
'    OutToLog "    Comment       : " & DataObj.Comment
'    OutToLog "    Pattern       : " & LastBurst & IIf(LastLabel = "", "", " at " & LastLabel)
'    OutToLog ""
'
'    Dim X As Long, Y As Long
'    Dim xunit As String, yunit As String
'    xunit = "": yunit = ""
'    Dim xmax As Integer, ymax As Integer
'    xmax = 0: ymax = 0
'    ReDim XValues(UBound(XParameters)) As String
'    ReDim YValues(UBound(YParameters)) As String
'    For X = 0 To UBound(XParameters)
'        XValues(X) = FormatValue(XParameters(X))
'        If xmax < Len(XValues(X)) Then xmax = Len(XValues(X))
'        If XParameters(X) <> 0 And xunit = "" Then xunit = Right(XValues(X), 1)
'    Next X
'    For Y = 0 To UBound(YParameters)
'        YValues(Y) = FormatValue(YParameters(Y))
'        If ymax < Len(YValues(Y)) Then ymax = Len(YValues(Y))
'        If YParameters(Y) <> 0 And yunit = "" Then yunit = Right(YValues(Y), 1)
'    Next Y
'
'    For X = 0 To UBound(XValues)
'        If XParameters(X) = 0 Then XValues(X) = Trim(XValues(X)) & " " & xunit
'        If Len(XValues(X)) < xmax Then XValues(X) = Space(xmax - Len(XValues(X))) & XValues(X)
'    Next X
'    For Y = 0 To UBound(YValues)
'        If YParameters(Y) = 0 Then YValues(Y) = Trim(YValues(Y)) & " " & yunit
'        If Len(YValues(Y)) < ymax Then YValues(Y) = Space(ymax - Len(YValues(Y))) & YValues(Y)
'    Next Y
'
'    For site = 0 To DataObj.SiteDim
'        If TheExec.Sites.site(site).Active Then Call DisplayShmooAscii(XValues, YValues, site)
'    Next site
'
'End Function

Private Sub DisplayShmooAscii(XValues() As String, YValues() As String, site As Long)
    Dim YstrOut() As String
    Dim Num As Long
    Dim x As Long
    Dim y As Long
    Dim strOut As String
    Dim i As Integer
    Dim numspc As Integer
    numspc = 14
    
    OutToLog "  Site : " & site
    OutToLog Space(numspc + 1) & DataObj.XParameter
    ReDim YstrOut(DataObj.YDim)
    Num = 1
    For y = 0 To DataObj.YDim
        strOut = ""
        If y = 0 Then
            For i = 0 To Len(XValues(0))
                strOut = Space(numspc + 1)
                For x = 0 To DataObj.XDim
                    strOut = strOut & Mid(XValues(x), i + 1, 1)
                Next x
                OutToLog strOut
            Next i
            strOut = Space(numspc) & "+"
            For i = 0 To UBound(XValues)
                strOut = strOut & "-"
            Next i
            strOut = strOut & "+"
            OutToLog strOut
        End If
        
        YstrOut(y) = " " & YValues(y) & "  |"
        For x = 0 To DataObj.XDim
            Select Case ShmooData(site).Data(Num).result
            Case rtaNoTest
                YstrOut(y) = YstrOut(y) & " "
            Case rtaPass
                YstrOut(y) = YstrOut(y) & "*"
            Case rtaFail
                YstrOut(y) = YstrOut(y) & "-"
            Case rtaError
                YstrOut(y) = YstrOut(y) & "!"
            End Select
            If x = DataObj.XDim Then YstrOut(y) = YstrOut(y) & "|"
            Num = Num + 1
        Next x
    Next y
    
    Num = 1
    For i = UBound(YstrOut) To 0 Step -1
        If Len(DataObj.YParameter) >= Num Then
            strOut = " " & Mid(DataObj.YParameter, Num, 1)
        Else
            strOut = "  "
        End If
        Num = Num + 1
        OutToLog strOut & YstrOut(i)
    Next i
    If Len(DataObj.YParameter) >= Num Then
        strOut = " " & Mid(DataObj.YParameter, Num, 1)
    Else
        strOut = "  "
    End If
    Num = Num + 1
    strOut = strOut & Space(numspc - 2) & "+"
    For i = 0 To UBound(XValues)
        strOut = strOut & "-"
    Next i
    strOut = strOut & "+"
    OutToLog strOut
    
    If Len(DataObj.YParameter) >= Num Then
        For i = Num To Len(DataObj.YParameter)
            strOut = " " & Mid(DataObj.YParameter, Num, 1)
            OutToLog strOut
            Num = Num + 1
        Next i
    End If
    
    OutToLog " *=PASS, -=FAIL !=ERROR " '.=ASSUMED FAIL ^=ASSUMED PASS"
'    OutToLog " >=MIN @=VALUE <=MAX &=MIN+VALUE+MAX #=MIN+VALUE +=VALUE+MAX $=MIN+MAX"
    OutToLog ""
End Sub


Private Sub DisplayShmooAscii_conv(XValues() As String, YValues() As String, site As Long)
    Dim YstrOut() As String
    Dim Num As Long
    Dim x As Long
    Dim y As Long
    Dim strOut As String
    Dim i As Integer
    Dim numspc As Integer
    numspc = 14
    
    OutToLog "  Site : " & site
    OutToLog Space(numspc + 1) & DataObj.YParameter
    ReDim YstrOut(DataObj.XDim)
    Num = 1
    For y = 0 To DataObj.XDim
        strOut = ""
        If y = 0 Then
            For i = 0 To Len(XValues(0))
                strOut = Space(numspc + 1)
                For x = 0 To DataObj.YDim
                    strOut = strOut & Mid(XValues(x), i + 1, 1)
                Next x
                OutToLog strOut
            Next i
            strOut = Space(numspc) & "+"
            For i = 0 To UBound(XValues)
                strOut = strOut & "-"
            Next i
            strOut = strOut & "+"
            OutToLog strOut
        End If
        
        YstrOut(y) = " " & YValues(y) & "  |"
        Num = y + 1
        For x = 0 To DataObj.YDim
            Select Case ShmooData(site).Data(Num).result
            Case rtaNoTest
                YstrOut(y) = YstrOut(y) & " "
            Case rtaPass
                YstrOut(y) = YstrOut(y) & "*"
            Case rtaFail
                YstrOut(y) = YstrOut(y) & "-"
            Case rtaError
                YstrOut(y) = YstrOut(y) & "!"
            End Select
            If x = DataObj.YDim Then YstrOut(y) = YstrOut(y) & "|"
            Num = Num + DataObj.XDim + 1
        Next x
    Next y
    
    Num = 1
    For i = UBound(YstrOut) To 0 Step -1
        If Len(DataObj.XParameter) >= Num Then
            strOut = " " & Mid(DataObj.XParameter, Num, 1)
        Else
            strOut = "  "
        End If
        Num = Num + 1
        OutToLog strOut & YstrOut(i)
    Next i
    If Len(DataObj.XParameter) >= Num Then
        strOut = " " & Mid(DataObj.XParameter, Num, 1)
    Else
        strOut = "  "
    End If
    Num = Num + 1
    strOut = strOut & Space(numspc - 2) & "+"
    For i = 0 To UBound(XValues)
        strOut = strOut & "-"
    Next i
    strOut = strOut & "+"
    OutToLog strOut
    
    If Len(DataObj.XParameter) >= Num Then
        For i = Num To Len(DataObj.XParameter)
            strOut = " " & Mid(DataObj.XParameter, Num, 1)
            OutToLog strOut
            Num = Num + 1
        Next i
    End If
    
    OutToLog " *=PASS, -=FAIL !=ERROR " '.=ASSUMED FAIL ^=ASSUMED PASS"
'    OutToLog " >=MIN @=VALUE <=MAX &=MIN+VALUE+MAX #=MIN+VALUE +=VALUE+MAX $=MIN+MAX"
    OutToLog ""
End Sub

Private Sub OutToLog(strOut As String)
    Call TheExec.Datalog.WriteComment(strOut)
End Sub

Private Function FormatValue(Value As Double) As String
'    FormatValue = Engr(value, False)
    FormatValue = Engr(Value, True)
End Function

Private Function Engr(D As Double, Optional GreekExponent = False) As String
    Dim Exp As Integer, ndx As Integer, Mantissa As Double
    Dim S As String

    Const GreekPrefixes = "afpnum KMGT"

    ' The format can have up to six digits displayed (xxx.xxx), so
    ' get a total of 8 digits worth of rounded result.  This allows
    ' for the multiplication by 10 or 100 that may be needed.
    S = Format(D, "0.0000000E+00")  ' convert to scientific, let VB round it
    Exp = CInt(Right(S, 3))         ' grab the exponent
    Mantissa = CDbl(Left(S, 9))     ' and mantissa

    Select Case (Exp Mod 3)
    Case 0
        Engr = Format(Mantissa, "0.000;-0.000")
    Case 1, -2
        Mantissa = Mantissa * 10
        Exp = Exp - 1
        Engr = Format(Mantissa, "00.000;-00.000")
    Case 2, -1
        Mantissa = Mantissa * 100
        Exp = Exp - 2
        Engr = Format(Mantissa, "000.000;-000.000")
    End Select

    If GreekExponent Then
        ndx = Exp / 3 + 7
        If ndx > 0 And ndx < 12 Then
            Engr = Engr + " " + Mid(GreekPrefixes, ndx, 1)
            If Exp = 0 Then Engr = Engr + "  "
        Else
            Engr = Engr + "E" + Format(Exp, "+00;-00")
        End If
    Else
        If (Exp <> 0) Then Engr = Engr + "E" + Format(Exp, "+00;-00")
    End If

End Function
