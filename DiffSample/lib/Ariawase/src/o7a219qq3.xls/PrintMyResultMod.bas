Attribute VB_Name = "PrintMyResultMod"
Option Explicit

Public Sub printMyResult(ByVal lngActiveSite As Long, ByVal lngTestNumber As Long, _
                        ByVal lngTestStatus As Long, ByVal lngParmFlag As Long, _
                        ByVal strPinName As String, ByVal lngChannelNumber As Long, _
                        ByVal dblLoLimit As Double, ByVal dblTestResult As Double, _
                        ByVal dblHiLimit As Double, ByVal lngUnits As Long, _
                        ByVal dblForceValue As Double, ByVal lngForceUnits As Long, _
                        ByVal loc As Long)
                        
    Dim strArgList() As String
    Dim lngArgCnt As Long
    Dim strJudge As String
    Dim strUnit As String
    Dim strFormat As String
    Dim myString, myStrData, myTestStatus As String
    Dim myStrBuf, strTestNumber As String
    Dim myMsg As String
    Dim myNum, mySite, myTestName, myPin, myChan, _
        myLow, myMeasured, myHigh, myForce, myLoc As String
    Dim strLength, myI, myPeriodN, myInt, myDpoint As Long
    Dim delta1, myUnitN As Long
    Dim myK As Double

    '=== get arg0~5 for datalog ===
    Call TheExec.DataManager.GetArgumentList(strArgList, lngArgCnt)
    dblLoLimit = val(strArgList(5 * LimitSetIndex + 0)) 'low limit
    dblHiLimit = val(strArgList(5 * LimitSetIndex + 1)) 'high limit
    strJudge = val(strArgList(5 * LimitSetIndex + 2))   'for judgement
    strUnit = strArgList(5 * LimitSetIndex + 3)        'unit for datalog
    strFormat = strArgList(5 * LimitSetIndex + 4)      'format for datalog
    myTestName = strArgList(5 * LimitSetIndex + 5)      'format for datalog

    '=== analyze specified data format ===
    strLength = Len(strFormat)
    For myI = 1 To strLength
        myString = Mid(strFormat, myI, 1)
        If myString = "." Then
            myPeriodN = myI
            Exit For
        End If
    Next
    myInt = CLng(Mid(strFormat, 2, myPeriodN - 2))
    myStrData = Mid(strFormat, myPeriodN + 1, strLength - myPeriodN - 1)
    myDpoint = CLng(myStrData)
    delta1 = myInt - myDpoint
    
    
    '=== analyze specified unit ===
    Select Case strUnit     'analysis for unit
        Case "V":   myK = 1#
        Case "mV":  myK = 1000#
        Case "uV":  myK = 1000000#
        Case "A":   myK = 1#
        Case "mA":  myK = 1000#
        Case "uA":  myK = 1000000#
        Case "nA":  myK = 1000000000#
        Case "%":   myK = 100#
        Case "S":   myK = 1#
        Case "mS":  myK = 1000#
        Case "uS":  myK = 1000000#
        Case "nS":  myK = 1000000000#
        Case "ohm":  myK = 1#
        Case "LSB":  myK = 1#
        Case "dB":  myK = 1#
        Case "C":  myK = 1#
        Case "GHz":  myK = 0.000000001
        Case "MHz":  myK = 0.000001
        Case "KHz":  myK = 0.001
        Case "Sm":   myK = 1#
        Case "mSm":  myK = 1000#
        Case "uSm":  myK = 1000000#
        Case "nSm":  myK = 1000000000#
        Case "W":   myK = 1#
        Case "mW":  myK = 1000#
        Case "uW":  myK = 1000000#
        Case "nW":  myK = 1000000000#
        Case "Kr":  myK = 0.001
        Case Else:  myK = -1
    End Select
    myUnitN = Len(strUnit)
    Select Case myUnitN
        Case 1: strUnit = "  " & strUnit
        Case 2: strUnit = " " & strUnit
        Case Else: strUnit = "   "
    End Select


    '--- make test Number ---
    If lngTestNumber >= 10000000 Then
        strTestNumber = CStr(lngTestNumber)
    ElseIf lngTestNumber >= 1000000 Then
        strTestNumber = " " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 100000 Then
        strTestNumber = "  " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 10000 Then
        strTestNumber = "   " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 1000 Then
        strTestNumber = "    " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 100 Then
        strTestNumber = "     " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 10 Then
        strTestNumber = "      " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 1 Then
        strTestNumber = "       " & CStr(lngTestNumber)
    Else
    End If
    myNum = strTestNumber
    
    '--- make Site number ---
    mySite = "   " & CStr(lngActiveSite)

    '--- make Result ---
    If lngTestStatus = 0 Then
        myTestStatus = "    PASS"
    Else
        myTestStatus = "    FAIL"
    End If
    
    '--- make Test Name ---
    strLength = Len(myTestName)
    myStrBuf = ""
    For myI = 1 To 16 - strLength
        myStrBuf = myStrBuf & " "
    Next
    myTestName = myStrBuf & myTestName
    
    '--- make Pin ---
    myPin = "    Empty"
    
    '--- make Channel ---
    If lngChannelNumber >= 10000000 Then
        myChan = CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 1000000 Then
        myChan = " " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 100000 Then
        myChan = "  " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 10000 Then
        myChan = "   " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 1000 Then
        myChan = "    " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 100 Then
        myChan = "     " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 10 Then
        myChan = "      " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 1 Then
        myChan = "       " & CStr(lngChannelNumber)
    Else
        myChan = "       " & CStr(lngChannelNumber)
    End If
        
    '--- make Force ---
    myForce = "     0.0000"
    
    '--- make Loc ---
    myLoc = "        0"
    
    
    '--- make Measured ---
    myMeasured = reFormatMyValue(myK, dblTestResult, myDpoint, strUnit)

    '--- make Low ---
    myLow = reFormatMyValue(myK, dblLoLimit, myDpoint, strUnit)
    
    '--- make High ---
    myHigh = reFormatMyValue(myK, dblHiLimit, myDpoint, strUnit)

    '=== print my message into datalog window ===
    myMsg = myNum & mySite & myTestStatus _
            & myTestName & myPin & myChan _
            & myLow & myMeasured & myHigh _
            & myForce & myLoc
    TheExec.Datalog.WriteComment myMsg
    
End Sub

Public Sub printMyResult_Logic(ByVal lngActiveSite As Long, ByVal lngTestNumber As Long, _
                        ByVal lngTestStatus As Long, ByVal lngParmFlag As Long, _
                        ByVal strPinName As String, ByVal lngChannelNumber As Long, _
                        ByVal dblLoLimit As Double, ByVal dblTestResult As Double, _
                        ByVal dblHiLimit As Double, ByVal lngUnits As Long, _
                        ByVal dblForceValue As Double, ByVal lngForceUnits As Long, _
                        ByVal loc As Long)
                        
    Dim strArgList() As String
    Dim lngArgCnt As Long
    Dim strJudge As String
    Dim strUnit As String
    Dim strFormat As String
    Dim myString, myStrData, myTestStatus As String
    Dim myStrBuf, strTestNumber As String
    Dim myMsg As String
    Dim myNum, mySite, myTestName, myPin, myChan, _
        myLow, myMeasured, myHigh, myForce, myLoc As String
    Dim strLength, myI, myPeriodN, myInt, myDpoint As Long
    Dim delta1, myUnitN As Long
    Dim myK As Double

    '=== get arg0~5 for datalog ===
    Call TheExec.DataManager.GetArgumentList(strArgList, lngArgCnt)
    dblLoLimit = 1
    dblHiLimit = 1
    strJudge = 3
    strUnit = strArgList(5 * LimitSetIndex + 28)        'unit for datalog
    strFormat = strArgList(5 * LimitSetIndex + 29)      'format for datalog
    myTestName = strArgList(5 * LimitSetIndex + 30)      'format for datalog

    '=== analyze specified data format ===
    strLength = Len(strFormat)
    For myI = 1 To strLength
        myString = Mid(strFormat, myI, 1)
        If myString = "." Then
            myPeriodN = myI
            Exit For
        End If
    Next
    myInt = CLng(Mid(strFormat, 2, myPeriodN - 2))
    myStrData = Mid(strFormat, myPeriodN + 1, strLength - myPeriodN - 1)
    myDpoint = CLng(myStrData)
    delta1 = myInt - myDpoint
    
    
    '=== analyze specified unit ===
    Select Case strUnit     'analysis for unit
        Case "V":   myK = 1#
        Case "mV":  myK = 1000#
        Case "uV":  myK = 1000000#
        Case "A":   myK = 1#
        Case "mA":  myK = 1000#
        Case "uA":  myK = 1000000#
        Case "nA":  myK = 1000000000#
        Case "%":   myK = 100#
        Case "S":   myK = 1#
        Case "mS":  myK = 1000#
        Case "uS":  myK = 1000000#
        Case "nS":  myK = 1000000000#
        Case "Kr":  myK = 0.001
        '--- add ---
        Case "GHz":  myK = 0.000000001
        Case "MHz":  myK = 0.000001
        Case "KHz":  myK = 0.001
        Case "sec":  myK = 1#
        Case "msec":  myK = 1000#
        Case "usec":  myK = 1000000#
        Case "Sm":   myK = 1#
        Case "mSm":  myK = 1000#
        Case "uSm":  myK = 1000000#
        Case "nSm":  myK = 1000000000#
        Case "ohm":  myK = 1#
        Case "LSB":  myK = 1#
        Case "dB":  myK = 1#
        Case "C":  myK = 1#
        Case "W":   myK = 1#
        Case "mW":  myK = 1000#
        Case "uW":  myK = 1000000#
        Case "nW":  myK = 1000000000#
        Case Else:  myK = -1
    End Select
    myUnitN = Len(strUnit)
    Select Case myUnitN
        Case 1: strUnit = "  " & strUnit
        Case 2: strUnit = " " & strUnit
        Case Else: strUnit = "   "
    End Select


    '--- make test Number ---
    If lngTestNumber >= 10000000 Then
        strTestNumber = CStr(lngTestNumber)
    ElseIf lngTestNumber >= 1000000 Then
        strTestNumber = " " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 100000 Then
        strTestNumber = "  " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 10000 Then
        strTestNumber = "   " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 1000 Then
        strTestNumber = "    " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 100 Then
        strTestNumber = "     " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 10 Then
        strTestNumber = "      " & CStr(lngTestNumber)
    ElseIf lngTestNumber >= 1 Then
        strTestNumber = "       " & CStr(lngTestNumber)
    Else
    End If
    myNum = strTestNumber
    
    '--- make Site number ---
    mySite = "   " & CStr(lngActiveSite)

    '--- make Result ---
    If lngTestStatus = 0 Then
        myTestStatus = "    PASS"
    Else
        myTestStatus = "    FAIL"
    End If
    
    '--- make Test Name ---
    strLength = Len(myTestName)
    myStrBuf = ""
    For myI = 1 To 16 - strLength
        myStrBuf = myStrBuf & " "
    Next
    myTestName = myStrBuf & myTestName
    
    '--- make Pin ---
    myPin = "    Empty"
    
    '--- make Channel ---
    If lngChannelNumber >= 10000000 Then
        myChan = CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 1000000 Then
        myChan = " " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 100000 Then
        myChan = "  " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 10000 Then
        myChan = "   " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 1000 Then
        myChan = "    " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 100 Then
        myChan = "     " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 10 Then
        myChan = "      " & CStr(lngChannelNumber)
    ElseIf lngChannelNumber >= 1 Then
        myChan = "       " & CStr(lngChannelNumber)
    Else
        myChan = "       " & CStr(lngChannelNumber)
    End If
        
    '--- make Force ---
    myForce = "     0.0000"
    
    '--- make Loc ---
    myLoc = "        0"
    
    
    '--- make Measured ---
    myMeasured = reFormatMyValue(myK, dblTestResult, myDpoint, strUnit)

    '--- make Low ---
    myLow = reFormatMyValue(myK, dblLoLimit, myDpoint, strUnit)
    
    '--- make High ---
    myHigh = reFormatMyValue(myK, dblHiLimit, myDpoint, strUnit)

    '=== print my message into datalog window ===
    myMsg = myNum & mySite & myTestStatus _
            & myTestName & myPin & myChan _
            & myLow & myMeasured & myHigh _
            & myForce & myLoc
    TheExec.Datalog.WriteComment myMsg
    
End Sub

Public Function reFormatMyValue(ByVal myK As Double, ByVal dblTestResult As Double, ByVal myDpoint As Long, ByVal strUnit As String) As String
    Dim myStrData, myStrBuf, myStrExp As String
    Dim strLength, myD, myI, delta2, myExpFlag As Long
    Dim myResult, myExp, myReal As Double
    
    If myK <> -1 Then
        myResult = dblTestResult * myK
    Else
        myResult = dblTestResult
    End If
    
    '--- if the value includes "E" ---
    myStrData = CStr(myResult)
    strLength = Len(myStrData)
    
    myStrExp = ""
    myExpFlag = 0
    For myI = 1 To strLength
        If myExpFlag = 1 Then
            myStrExp = myStrExp & Mid(myStrData, myI, 1)
        End If
        If Mid(myStrData, myI, 1) = "E" Then
            myExpFlag = 1
            myReal = Mid(myStrData, 1, myI - 1)
        End If
    Next
    If myExpFlag = 1 Then
        myExp = CDbl(myStrExp)
        myD = 0
        For myI = 1 To strLength
            If Mid(myStrData, myI, 1) = "." Then
                myD = myI
                Exit For
            End If
        Next
        myStrBuf = "0."
        For myI = 1 To Abs(myExp) - 1
            myStrBuf = myStrBuf & "0"
        Next
        strLength = Len(myReal)
        If Mid(myReal, 1, 1) = "-" Then
            myStrBuf = "-" & myStrBuf & Mid(myReal, 2, 1) & Mid(myReal, 4, strLength - 4)
        Else
            myStrBuf = myStrBuf & Mid(myReal, 1, 1) & Mid(myReal, 3, strLength - 3)
        End If
        myStrData = myStrBuf
        strLength = Len(myStrData)
    End If
    
    
    myD = 0
    For myI = 1 To strLength
        If Mid(myStrData, myI, 1) = "." Then
            myD = myI
            Exit For
        End If
    Next
    
    If myDpoint <> 0 Then
        If myD = 0 Then
            myStrData = myStrData & "."
            myD = strLength + 1
            strLength = strLength + 1
        End If
    Else
        myD = 0
    End If
    
    delta2 = strLength - myD
    If delta2 = 0 Then
        For myI = 1 To myDpoint
            myStrData = myStrData & "0"
        Next
    ElseIf delta2 < myDpoint Then
        For myI = 1 To myDpoint - delta2
            myStrData = myStrData & "0"
        Next
    ElseIf myDpoint = 0 Then
        myStrData = myStrData
    Else
        myStrData = Mid(myStrData, 1, myD + myDpoint)
    End If
    
    strLength = Len(myStrData)
    myStrBuf = ""
    For myI = 1 To 16 - strLength - 3
        myStrBuf = myStrBuf & " "
    Next
    If Mid(myStrBuf, 1, 1) <> " " Then
        myStrBuf = myStrBuf & " "
    End If
    reFormatMyValue = myStrBuf & myStrData & strUnit
    
End Function

Public Sub printMyHeader()

    Dim myMsg As String
    myMsg = "  Number Site Result    Test Name     Pin   Channel     Low          Measured          High           Force        Loc"
    TheExec.Datalog.WriteComment myMsg

End Sub


