VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_OptCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ILLUM_MAKER_LIST_WRKSHT_NAME = "IllumLst"
Private Const DEFAULT_ILLUM_GPIB_ADDR_NIKON = 8         'NIKON Illuminator GPIB Addr No.8
Private Const DEFAULT_ILLUM_GPIB_ADDR_IA = 2            'InterAction GPIB Addr No.2

Public IllumMaker As String
Public IllumModel As String
Public IllumGpibAddr As Integer
Public IllumWedgeMin  As Integer
Public IllumWedgeMax  As Integer
Public IllumNdMin  As Integer
Public IllumNdMax  As Integer
Public IllumHMin  As Integer
Public IllumHMax  As Integer
Public IllumChartMin  As Integer
Public IllumChartMax  As Integer
Public IllumZMin  As Integer
Public IllumZMax  As Integer
Public IllumXMin  As Integer
Public IllumXMax  As Integer
Public IllumYMin  As Integer
Public IllumYMax  As Integer
Public IllumModzMin  As Integer
Public IllumModzMax  As Integer

Private xOptCondInfo() As csPALS_OptCondParams

Private colColor As Collection  'カラーフィルター設定
Private colShut As Collection   'シャッター設定
Private colMod As Collection    'モジュール設定
Private colModz As Collection   'ダウンポジション設定
Private colSlide As Collection  'スライド設定

Private Const IllumInfoMax As Integer = 1500

Private strColorFil(IllumInfoMax) As String
Private strShutter(IllumInfoMax) As String
Private strOptMod(IllumInfoMax) As String
Private strSlide(IllumInfoMax) As String
Private strOptModz(IllumInfoMax) As String

Private Sub sub_IllumInfoSet()

    Set colColor = Nothing
    Set colColor = New Collection
    Set colShut = Nothing
    Set colShut = New Collection
    Set colMod = Nothing
    Set colMod = New Collection
    Set colSlide = Nothing
    Set colSlide = New Collection
    Set colModz = Nothing
    Set colModz = New Collection

    
    Dim intIllumInfo As Integer
    
    Erase strColorFil
    Erase strShutter
    Erase strOptMod
    Erase strOptModz
    Erase strSlide
'///////////Illuminator Parameter///////////////////////
    Select Case IllumModel
        Case "N-SIS2":
                strColorFil(0) = "NotUse"
                strColorFil(1) = "WHITE"
                strColorFil(2) = "RED"
                strColorFil(3) = "GREEN"
                strColorFil(4) = "BLUE"
                strColorFil(5) = "LB"
                strColorFil(6) = "FL"
                                
                strShutter(0) = "NotUse"
                strShutter(1) = "OPEN"
                strShutter(2) = "CLOSE"
                
                strOptMod(0) = "INIT"
                strOptMod(1) = "THROUGH"
                strOptMod(2) = "PIN"
                strOptMod(3) = "F_UNIT"

                strOptModz(0) = "NotUse"

                IllumNdMin = 1
                IllumNdMax = 12
                
                IllumWedgeMin = 0
                IllumWedgeMax = 827
                
                IllumChartMin = 1
                IllumChartMax = 8
                
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999
                
        Case "N-SIS3", "N-SIS3KAI":
                strColorFil(0) = "NotUse"
                strColorFil(1) = "WHITE"
                strColorFil(2) = "RED"
                strColorFil(3) = "GREEN"
                strColorFil(4) = "BLUE"
                strColorFil(5) = "LB"
                strColorFil(6) = "FL"
                                
                strShutter(0) = "NotUse"
                strShutter(1) = "OPEN"
                strShutter(2) = "CLOSE"
                
                strOptMod(0) = "INIT"
                strOptMod(1) = "THROUGH"
                strOptMod(2) = "PIN"
                strOptMod(3) = "F_UNIT"

                strOptModz(0) = "NotUse"
                strOptModz(1) = "B1"
                strOptModz(2) = "Up"
                strOptModz(3) = "Down"

                IllumNdMin = 1
                IllumNdMax = 12
                
                IllumWedgeMin = 0
                IllumWedgeMax = 827
                
                IllumChartMin = 1
                IllumChartMax = 4
        
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999

                IllumModzMin = 1
                IllumModzMax = 3

        Case "N-SIS5", "N-SIS5KAI":
                strColorFil(0) = "NotUse"
                strColorFil(1) = "WHITE"
                strColorFil(2) = "RED"
                strColorFil(3) = "GREEN"
                strColorFil(4) = "BLUE"
                strColorFil(5) = "LB"
                strColorFil(6) = "FL"
                                
                strShutter(0) = "NotUse"
                strShutter(1) = "OPEN"
                strShutter(2) = "CLOSE"
                
                strOptMod(0) = "INIT"
                strOptMod(1) = "THROUGH"
                strOptMod(2) = "PIN"
                strOptMod(3) = "F_UNIT"

                strOptModz(0) = "NotUse"
                strOptModz(1) = "B1"
                strOptModz(2) = "Up"
                strOptModz(3) = "Down"
                strOptModz(4) = "Down2"
                strOptModz(5) = "Down3"
                strOptModz(6) = "Down4"
                strOptModz(7) = "Down5"
                strOptModz(8) = "Down6"
                
                IllumNdMin = 1
                IllumNdMax = 16
                
                IllumWedgeMin = 0
                IllumWedgeMax = 780
                
                IllumChartMin = -1
                IllumChartMax = -1
        
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999

                IllumModzMin = 1
                IllumModzMax = 8

        Case "CALSYS-CD13":
                strColorFil(0) = "WHITE"
                strColorFil(1) = "BLUE"
                strColorFil(2) = "GREEN"
                strColorFil(3) = "RED"
                strColorFil(4) = "FILTER4"
                strColorFil(5) = "FILTER5"
                                                
                strOptMod(0) = "F3.5"
                strOptMod(1) = "F1.2"
                strOptMod(2) = "F8"
                strOptMod(3) = "F16"
                strOptMod(4) = "F22"
                strOptMod(5) = "F32"
                
                IllumWedgeMin = 0
                IllumWedgeMax = 999
                
                IllumNdMin = 0
                IllumNdMax = 5
                
                IllumHMin = 0
                IllumHMax = 5
        
                IllumChartMin = -1
                IllumChartMax = -1
        
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999
        
        Case "CALSYS-CD15":
                strColorFil(0) = "WHITE"
                strColorFil(1) = "BLUE"
                strColorFil(2) = "GREEN"
                strColorFil(3) = "RED"
                strColorFil(4) = "FILTER4"
                strColorFil(5) = "FILTER5"
                                                
                strOptMod(0) = "F3.5"
                strOptMod(1) = "F1.2"
                strOptMod(2) = "F8"
                strOptMod(3) = "F16"
                strOptMod(4) = "F22"
                strOptMod(5) = "F32"
                
                IllumWedgeMin = 0
                IllumWedgeMax = 999
                
                IllumNdMin = 0
                IllumNdMax = 5
                
                IllumHMin = 0
                IllumHMax = 5
                
                IllumChartMin = -1
                IllumChartMax = -1
        
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999

        Case "IA-OPT226":
                strColorFil(0) = "FILTER0"
                strColorFil(1) = "FILTER1"
                strColorFil(2) = "FILTER2"
                strColorFil(3) = "FILTER3"
                strColorFil(4) = "FILTER4"
                strColorFil(5) = "FILTER5"
                                                
                strOptMod(0) = "0"
                strOptMod(1) = "1"
                
                IllumWedgeMin = 0
                IllumWedgeMax = 1200
                IllumNdMin = 0
                IllumNdMax = 5
                IllumHMin = -1
                IllumHMax = -1
                IllumChartMin = 0
                IllumChartMax = 3
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999

        Case "OPTPATH-2SP":
                strColorFil(0) = "FILTER0"
                strColorFil(1) = "FILTER1"
                strColorFil(2) = "FILTER2"
                strColorFil(3) = "FILTER3"
                strColorFil(4) = "FILTER4"
                strColorFil(5) = "FILTER5"
                                                
                strOptMod(0) = "F3"
                strOptMod(1) = "F11.8"
                strOptMod(2) = "F16"
                strOptMod(3) = "F1.2"
                strOptMod(4) = "F16-2"
                strOptMod(5) = "F16-3"

                IllumWedgeMin = 0
                IllumWedgeMax = 1200
                IllumNdMin = 0
                IllumNdMax = 5
                IllumHMin = -1
                IllumHMax = -1
                IllumChartMin = -1
                IllumChartMax = -1
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999

        Case "OPTPATH-2SPDR":
                strColorFil(0) = "FILTER0"
                strColorFil(1) = "FILTER1"
                strColorFil(2) = "FILTER2"
                strColorFil(3) = "FILTER3"
                strColorFil(4) = "FILTER4"
                strColorFil(5) = "FILTER5"
                                                
                strOptMod(0) = "F3"
                strOptMod(1) = "F11.8"
                strOptMod(2) = "F16"
                strOptMod(3) = "F1.2"
                strOptMod(4) = "F16-2"
                strOptMod(5) = "F16-3"
                
                IllumWedgeMin = 0
                IllumWedgeMax = 1200
                IllumNdMin = 0
                IllumNdMax = 5
                IllumHMin = -1
                IllumHMax = -1
                IllumChartMin = -1
                IllumChartMax = -1
                IllumZMin = 0
                IllumZMax = 1600
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999

        Case "OPTPATH-4SA":
                strColorFil(0) = "FILTER0"
                strColorFil(1) = "FILTER1"
                strColorFil(2) = "FILTER2"
                strColorFil(3) = "FILTER3"
                strColorFil(4) = "FILTER4"
                strColorFil(5) = "FILTER5"
                                                
                strOptMod(0) = "F2"
                strOptMod(1) = "F16"
                strOptMod(2) = "F24"
                
                IllumWedgeMin = 0
                IllumWedgeMax = 1200
                IllumNdMin = 0
                IllumNdMax = 5
                IllumHMin = -1
                IllumHMax = -1
                IllumChartMin = -1
                IllumChartMax = -1
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999

        Case "OPTPATH-4SPN":
                strColorFil(0) = "FILTER0"
                strColorFil(1) = "FILTER1"
                strColorFil(2) = "FILTER2"
                strColorFil(3) = "FILTER3"
                strColorFil(4) = "FILTER4"
                strColorFil(5) = "FILTER5"
                                                
                strOptMod(0) = "F2.4"
                strOptMod(1) = "F7"
                strOptMod(2) = "F8"
                
                IllumWedgeMin = 0
                IllumWedgeMax = 1200
                IllumNdMin = 0
                IllumNdMax = 5
                IllumHMin = -1
                IllumHMax = -1
                IllumChartMin = -1
                IllumChartMax = -1
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999

        Case "OPTPATH-8SPN2":
                strColorFil(0) = "FILTER0"
                strColorFil(1) = "FILTER1"
                strColorFil(2) = "FILTER2"
                strColorFil(3) = "FILTER3"
                strColorFil(4) = "FILTER4"
                strColorFil(5) = "FILTER5"
                                                
                strOptMod(0) = "F2.5"
                strOptMod(1) = "F7"
                strOptMod(2) = "F22"
                
                IllumWedgeMin = 0
                IllumWedgeMax = 1200
                IllumNdMin = 0
                IllumNdMax = 5
                IllumHMin = -1
                IllumHMax = -1
                IllumChartMin = -1
                IllumChartMax = -1
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999

        Case "OPTPATH-8SPN4":
                strColorFil(0) = "FILTER0"
                strColorFil(1) = "FILTER1"
                strColorFil(2) = "FILTER2"
                strColorFil(3) = "FILTER3"
                strColorFil(4) = "FILTER4"
                strColorFil(5) = "FILTER5"
                                                
                strOptMod(0) = "F10"
                strOptMod(1) = "F2.5"
                strOptMod(2) = "F22"
                
                IllumWedgeMin = 0
                IllumWedgeMax = 1200
                IllumNdMin = 0
                IllumNdMax = 5
                IllumHMin = -1
                IllumHMax = -1
                IllumChartMin = -1
                IllumChartMax = -1
                IllumZMin = -1
                IllumZMax = -1
                IllumXMin = -9999
                IllumXMax = -9999
                IllumYMin = -9999
                IllumYMax = -9999

        Case "OPT027P":
                strColorFil(0) = "FILTER0"
                strColorFil(1) = "FILTER1"
                strColorFil(2) = "FILTER2"
                strColorFil(3) = "FILTER3"
                strColorFil(4) = "FILTER4"
                strColorFil(5) = "FILTER5"
                                                
                strShutter(0) = "OPEN"
                strShutter(1) = "CLOSE"
                                                
                strSlide(0) = "IN"
                strSlide(1) = "OUT"
                
                Dim intIloop As Integer
                For intIloop = 0 To 1300
                    strOptMod(intIloop) = CStr(intIloop)
                Next
                IllumWedgeMin = 0
                IllumWedgeMax = 1300
                IllumNdMin = 0
                IllumNdMax = 5
                IllumHMin = -1
                IllumHMax = -1
                IllumChartMin = -1
                IllumChartMax = -1
                IllumZMin = 30
                IllumZMax = 500
                IllumXMin = -300
                IllumXMax = 300
                IllumYMin = -300
                IllumYMax = 300

        Case default:
                MsgBox "Unknown IllumModel (" & IllumModel & ")@sub_IllumInfoSet! " & vbCrLf & "ErrCode.4-3-01-4-01", vbExclamation, "Error@sub_illumInfoSet"

    End Select
'////////////////////////////////////////////////////////
        

    For intIllumInfo = 0 To IllumInfoMax
        If strColorFil(intIllumInfo) <> "" Then
            colColor.Add Item:=intIllumInfo, key:=strColorFil(intIllumInfo)
        End If
        If strShutter(intIllumInfo) <> "" Then
            colShut.Add Item:=intIllumInfo, key:=strShutter(intIllumInfo)
        End If
        If strOptMod(intIllumInfo) <> "" Then
            colMod.Add Item:=intIllumInfo, key:=strOptMod(intIllumInfo)
        End If
        If strOptModz(intIllumInfo) <> "" Then
            colModz.Add Item:=intIllumInfo, key:=strOptModz(intIllumInfo)
        End If
        If strSlide(intIllumInfo) <> "" Then
            colSlide.Add Item:=intIllumInfo, key:=strSlide(intIllumInfo)
        End If
    Next

End Sub

Private Sub Class_Initialize()
    Call ReadOptListInfo
    
    '>>>2011/6/13 M.IMAMURA CopyInfo Add.
    If g_blnUseCSV = True And g_blnOptAdjusting = False Then
        Sheets("Read CSV").Select
        Range("A7:Z2000").Select
        Application.CutCopyMode = False
        Selection.Copy
        If IllumMaker = NIKON Then
            Sheets(NIKON_WRKSHT_NAME).Select
        Else
            Sheets(IA_WRKSHT_NAME).Select
        End If
        Range("A7").Select
        Selection.PasteSpecial Paste:=xlPasteValues
    End If
    '<<<2011/6/13 M.IMAMURA CopyInfo Add.
    
    Call ReadOptCondInfo
End Sub

Private Sub Class_Terminate()
    Erase xOptCondInfo
End Sub

Public Property Get CondInfoI(ByVal Condcnt As Integer) As csPALS_OptCondParams
    Set CondInfoI = xOptCondInfo(Condcnt)
End Property

Public Property Get CondInfo(ByVal CondcntIdentifier As String) As csPALS_OptCondParams
    Dim Optcnt As Long
    For Optcnt = 0 To UBound(xOptCondInfo)
        If xOptCondInfo(Optcnt).OptIdentifier = CondcntIdentifier Then
            Set CondInfo = xOptCondInfo(Optcnt)
            Exit For
        End If
    Next
End Property
Public Function CondInfoNo(ByVal CondcntIdentifier As String) As Long
    Dim Optcnt As Long
    For Optcnt = 0 To UBound(xOptCondInfo)
        If xOptCondInfo(Optcnt).OptIdentifier = CondcntIdentifier Then
            CondInfoNo = Optcnt
            Exit For
        End If
    Next
End Function

Public Sub ReadOptCondInfo()
    Dim wkshtObj As Object
    Dim basePoint As Variant
    Dim nodeRow As Long
    Dim nodeColumn As Long
    Dim cmdRow As Long
    Dim cmdColumn As Long
    Dim axisCmd As String
    Dim cmdArgstr As String
    Dim axisCmdLst(30) As String
    Dim Optcnt As Long
    Dim strFnumber_buf As String
    Dim strColor_buf As String
   
    Dim intChkCnt As Long
    Dim intAxisNo As Integer
    intAxisNo = 1
    
    On Error GoTo errorLable

    axisCmdLst(0) = "Axis.Pattern"
    axisCmdLst(1) = "Axis.Level"
    axisCmdLst(2) = "Axis.RegisterAxis"
    axisCmdLst(3) = "Shutter"
    axisCmdLst(4) = "WedgeFilter"
    axisCmdLst(5) = "Pattern"
    axisCmdLst(6) = "Color"
    axisCmdLst(7) = "NDFilter"
    axisCmdLst(8) = "ColorTemperature"
    axisCmdLst(9) = "FNumberTurret"
    axisCmdLst(10) = "OptMod.FnumberTurret"
    axisCmdLst(11) = "Identifier"
    axisCmdLst(12) = "L"
    axisCmdLst(13) = "C"
    axisCmdLst(14) = "N"
    axisCmdLst(15) = "A"
    axisCmdLst(16) = "P"
    axisCmdLst(17) = "I"
    axisCmdLst(18) = "J"
    axisCmdLst(19) = "H"
    axisCmdLst(20) = "S"
    axisCmdLst(21) = "Z"
    axisCmdLst(22) = "X"
    axisCmdLst(23) = "Y"
    axisCmdLst(24) = "D"
    axisCmdLst(25) = "Opt_Judgment"
    axisCmdLst(26) = "OptModZ.Position"
    Optcnt = 0

    If IllumMaker = NIKON Then
        '******* NIKON *******
        '======= WorkSheet Select ========
        Set wkshtObj = ThisWorkbook.Sheets(NIKON_WRKSHT_NAME)
    ElseIf IllumMaker = INTERACTION Or IllumMaker = KESILLUM Then
        Set wkshtObj = ThisWorkbook.Sheets(IA_WRKSHT_NAME)
    Else
        MsgBox "Unknown IllumMaker!" & vbCrLf & "ErrCode.4-3-02-4-02", vbExclamation, "Error@ReadOptCondInfo"
        Exit Sub
    End If
    
    '======= WorkSheet ErrorProcess ========
    If IsEmpty(wkshtObj) Then
        MsgBox "Non OptSet WorkSheet [" & NIKON_WRKSHT_NAME & "]!" & vbCrLf & "ErrCode.4-3-02-8-03", vbExclamation, "Error@ReadOptCondInfo"
        Exit Sub
    End If

    '======= Base Point Find ========
    Set basePoint = wkshtObj.Range("A1:K100").Find("Sw_node")
    If basePoint Is Nothing Then
        MsgBox "Not Found Sw_node @[" & NIKON_WRKSHT_NAME & "] WorkSheet!" & vbCrLf & "ErrCode.4-3-02-8-04", vbExclamation, "Error@ReadOptCondInfo"
        Exit Sub
    End If

    '======= Search Node Start Addres Define ========
    nodeRow = basePoint.Row + 2      'Search "Sw_node" Start Row
    nodeColumn = basePoint.Column   '"Sw_node" Column
    cmdRow = basePoint.Row + 1        '"illuminator Cmd" Row

    Do While Not wkshtObj.Cells(nodeRow, nodeColumn) = vbNullString
        cmdColumn = basePoint.Column + 1 '"illuminator Cmd" Start Column

        '************* Sw_node Check *************
        If wkshtObj.Cells(nodeRow, nodeColumn) = Sw_Node Then
            ReDim Preserve xOptCondInfo(Optcnt)
            Set xOptCondInfo(Optcnt) = New csPALS_OptCondParams
            strFnumber_buf = ""
            strColor_buf = ""
            Do While Not (wkshtObj.Cells(cmdRow, cmdColumn) = vbNullString And wkshtObj.Cells(cmdRow - 1, cmdColumn) = vbNullString)
                
                axisCmd = wkshtObj.Cells(cmdRow, cmdColumn)
                cmdArgstr = wkshtObj.Cells(nodeRow, cmdColumn)
                
                If wkshtObj.Cells(cmdRow - 1, cmdColumn) Like "*Identifier*" Or wkshtObj.Cells(cmdRow - 1, cmdColumn) Like "*Opt_Judgment*" Then
                    axisCmd = wkshtObj.Cells(cmdRow - 1, cmdColumn)
                    cmdArgstr = wkshtObj.Cells(nodeRow, cmdColumn)
                End If

                With xOptCondInfo(UBound(xOptCondInfo))
                If cmdArgstr <> vbNullString Or axisCmd = axisCmdLst(2) Then
                    '#Axis.Pattern
                    If axisCmd = axisCmdLst(0) Then
                        .AxisPattern = CInt(cmdArgstr)
                    'Axis.Level  L
                    ElseIf (axisCmd = axisCmdLst(1) Or axisCmd = axisCmdLst(12)) Then
                        .AxisLevel = CDbl(cmdArgstr)
                    '#Axis.RegisterAxis
                    ElseIf axisCmd = axisCmdLst(2) Then
                        '>>> 2011/8/8 M.IMAMURA RegisterAxis Auto
'                        .AxisRegisterAxis = CInt(cmdArgstr)
                        If .AxisLevel > 0 Then
                            .AxisRegisterAxis = intAxisNo
                            If blnAutoAxisSet = True Then
                                wkshtObj.Cells(nodeRow, cmdColumn) = intAxisNo
                            End If
                            intAxisNo = intAxisNo + 1
                        '<<< 2011/8/8 M.IMAMURA RegisterAxis Auto
                        End If
                    '#Shutter
                    ElseIf axisCmd = axisCmdLst(3) Or axisCmd = axisCmdLst(20) Then
                        If IsNumeric(cmdArgstr) = True Then
                            .Shutter = CInt(cmdArgstr)
                            If .Shutter >= colShut.Count Or strShutter(.Shutter) = "NotUse" Then
                                MsgBox "Not Available [" & cmdArgstr & "]@Shutter,@" & .OptIdentifier & "!" & vbCrLf & "ErrCode.4-3-02-5-05", vbExclamation, "Error@ReadOptCondInfo"
                            End If
                        Else
                            .Shutter = colShut(cmdArgstr)
                        End If
                    '#WedgeFilter
                    ElseIf (axisCmd = axisCmdLst(4) Or axisCmd = axisCmdLst(15)) Then
                        .WedgeFilter = CInt(cmdArgstr)
                        If .WedgeFilter > IllumWedgeMax Or .WedgeFilter < IllumWedgeMin Then
                            MsgBox "Not Available [" & cmdArgstr & "]@WedgeFilter,@" & .OptIdentifier & vbCrLf & "Must Be " & CStr(IllumWedgeMin) & "-" & CStr(IllumWedgeMax) & "!" & vbCrLf & "ErrCode.4-3-02-5-06", vbExclamation, "Error@ReadOptCondInfo"
                        End If
                    '#Pattern P
                    ElseIf (axisCmd = axisCmdLst(5) Or axisCmd = axisCmdLst(16)) Then
                        .Pattern = CInt(cmdArgstr)
                        If .Pattern > IllumChartMax Or .Pattern < IllumChartMin Then
                            MsgBox "Not Available [" & cmdArgstr & "]@Pattern,@" & .OptIdentifier & vbCrLf & "Must Be " & CStr(IllumChartMin) & "-" & CStr(IllumChartMax) & "!" & vbCrLf & "ErrCode.4-3-02-5-07", vbExclamation, "Error@ReadOptCondInfo"
                        End If
                    '#Color C
                    ElseIf (axisCmd = axisCmdLst(6) Or axisCmd = axisCmdLst(13)) Then
                        If IsNumeric(cmdArgstr) = True Then
                            .color = CInt(cmdArgstr)
                            If .color >= colColor.Count Or strColorFil(.color) = "NotUse" Then
                                MsgBox "Not Available [" & cmdArgstr & "]@Color,@" & .OptIdentifier & "!" & vbCrLf & "ErrCode.4-3-02-5-08", vbExclamation, "Error@ReadOptCondInfo"
                            End If
                            strColor_buf = strColorFil(.color)
                        Else
                            .color = colColor(cmdArgstr)
                            strColor_buf = cmdArgstr
                        End If
                    '#NDFilter N
                    ElseIf (axisCmd = axisCmdLst(7) Or axisCmd = axisCmdLst(14)) Then
                        .NDFilter = CInt(cmdArgstr)
                        If .NDFilter > IllumNdMax Or .NDFilter < IllumNdMin Then
                            MsgBox "Not Available [" & cmdArgstr & "]@NDFilter,@" & .OptIdentifier & vbCrLf & "Must Be  " & CStr(IllumNdMin) & "-" & CStr(IllumNdMax) & "!" & vbCrLf & "ErrCode.4-3-02-5-09", vbExclamation, "Error@ReadOptCondInfo"
                        End If
                    '#ColorTemperature
                    ElseIf axisCmd = axisCmdLst(8) Then
                        .ColorTemperature = CInt(cmdArgstr)
                    '#FNumberTurret
                    ElseIf axisCmd = axisCmdLst(9) Then
                        .FNumberTurret = CInt(cmdArgstr)
                    '#OptMod.FnumberTurret
                    ElseIf axisCmd = axisCmdLst(10) Then
                        If IsNumeric(cmdArgstr) = True Then
                            .OptModFnumberTurret = CInt(cmdArgstr)
                            If .OptModFnumberTurret >= colMod.Count Or strOptMod(.OptModFnumberTurret) = "NotUse" Then
                                MsgBox "Not Available [" & cmdArgstr & "]@OptMod.FnumberTurret,@" & .OptIdentifier & "!" & vbCrLf & "ErrCode.4-3-02-5-10", vbExclamation, "Error@ReadOptCondInfo"
                            End If
                            strFnumber_buf = strOptMod(.OptModFnumberTurret)
                        Else
                            strFnumber_buf = cmdArgstr
                            .OptModFnumberTurret = colMod(cmdArgstr)
                        End If
                    '>>>2012/10/18 M.Imamura Add.
                    '#OptModZ.Position
                    ElseIf axisCmd = axisCmdLst(26) Then
                        If IsNumeric(cmdArgstr) = True Then
                            .OptModDownPosition = CInt(cmdArgstr)
                            If .OptModDownPosition >= colModz.Count Or strOptModz(.OptModDownPosition) = "NotUse" Then
                                MsgBox "Not Available [" & cmdArgstr & "]@OptModZ.Position,@" & .OptIdentifier & "!" & vbCrLf & "ErrCode.4-3-02-5-17", vbExclamation, "Error@ReadOptCondInfo"
                            End If
                        Else
                            .OptModDownPosition = colModz(cmdArgstr)
                        End If
                    '<<<2012/10/18 M.Imamura Add.
                    '#Identifier
                    ElseIf axisCmd = axisCmdLst(11) Then
                        .OptIdentifier = CStr(cmdArgstr)
                    '>>>2012/3/13 M.Imamura Add.
                    '#Opt_Judgment
                    ElseIf axisCmd = axisCmdLst(25) Then
                        .OptJudge = CStr(cmdArgstr)
                        If .OptJudge <> "ON" And .OptJudge <> "OFF" Then
                            MsgBox "Not Available [" & cmdArgstr & "]@Opt_Judgment,@" & .OptIdentifier & "!" & vbCrLf & "ErrCode.4-3-02-5-63", vbExclamation, "Error@ReadOptCondInfo"
                        End If
                    '<<<2012/3/13 M.Imamura Add.
                    '#I
                    ElseIf axisCmd = axisCmdLst(17) Then
                        If IsNumeric(cmdArgstr) = True Then
                            .EPD1 = CInt(cmdArgstr)
                            If .EPD1 >= colMod.Count Or strOptMod(.EPD1) = "NotUse" Then
                                MsgBox "Not Available [" & cmdArgstr & "]@Iris,@" & .OptIdentifier & "!" & vbCrLf & "ErrCode.4-3-02-5-11", vbExclamation, "Error@ReadOptCondInfo"
                            End If
                            strFnumber_buf = strOptMod(.EPD1)
                        Else
                            strFnumber_buf = cmdArgstr
                            .EPD1 = colMod(cmdArgstr)
                        End If
                    '#J
                    ElseIf axisCmd = axisCmdLst(18) Then
                        .EPD2 = CInt(cmdArgstr)
                    '#H
                    ElseIf axisCmd = axisCmdLst(19) Then
                        .Frosted = CInt(cmdArgstr)
                        If .Frosted > IllumHMax Or .Frosted < IllumHMin Then
                            MsgBox "Not Available [" & cmdArgstr & "]@Frosted,@" & .OptIdentifier & vbCrLf & "Must Be " & CStr(IllumHMin) & "-" & CStr(IllumHMax) & "!" & vbCrLf & "ErrCode.4-3-02-5-12", vbExclamation, "Error@ReadOptCondInfo"
                        End If
                    '#Z
                    ElseIf axisCmd = axisCmdLst(21) Then
                        .IrisPos = CInt(cmdArgstr)
                        If .IrisPos > IllumZMax Or .IrisPos < IllumZMin Then
                            MsgBox "Not Available [" & cmdArgstr & "]@IrisPos,@" & .OptIdentifier & vbCrLf & "Must Be " & CStr(IllumZMin) & "-" & CStr(IllumZMax) & "!" & vbCrLf & "ErrCode.4-3-02-5-13", vbExclamation, "Error@ReadOptCondInfo"
                        End If
                    '#X
                    ElseIf axisCmd = axisCmdLst(22) Then
                        .DeviceX = CInt(cmdArgstr)
                        If .DeviceX > IllumXMax Or .DeviceX < IllumXMin Then
                            MsgBox "Not Available [" & cmdArgstr & "]@DeviceX,@" & .OptIdentifier & vbCrLf & "Must Be " & CStr(IllumXMin) & "-" & CStr(IllumXMax) & "!" & vbCrLf & "ErrCode.4-3-02-5-14", vbExclamation, "Error@ReadOptCondInfo"
                        End If
                    '#Y
                    ElseIf axisCmd = axisCmdLst(23) Then
                        .DeviceY = CInt(cmdArgstr)
                        If .DeviceY > IllumYMax Or .DeviceY < IllumYMin Then
                            MsgBox "Not Available [" & cmdArgstr & "]@DeviceY,@" & .OptIdentifier & vbCrLf & "Must Be " & CStr(IllumYMin) & "-" & CStr(IllumYMax) & "!" & vbCrLf & "ErrCode.4-3-02-5-15", vbExclamation, "Error@ReadOptCondInfo"
                        End If
                    '#D
                    ElseIf axisCmd = axisCmdLst(24) Then
                        If IsNumeric(cmdArgstr) = True Then
                            .SlideINOUT = CInt(cmdArgstr)
                            If .SlideINOUT >= colSlide.Count Or strSlide(.SlideINOUT) = "NotUse" Then
                                MsgBox "Not Available [" & cmdArgstr & "]@SlideINOUT,@" & .OptIdentifier & "!" & vbCrLf & "ErrCode.4-3-02-5-16", vbExclamation, "Error@ReadOptCondInfo"
                            End If
                        Else
                            .SlideINOUT = colSlide(cmdArgstr)
                        End If
                    Else
                        MsgBox "Illegal Command[" & axisCmd & "]@Optini!" & vbCrLf & "ErrCode.4-3-02-4-17", vbExclamation, "Error@ReadOptCondInfo"
                    End If
                End If
                End With
                
                cmdColumn = cmdColumn + 1
            Loop

            With xOptCondInfo(UBound(xOptCondInfo))
            If IllumMaker = NIKON Then
                'OptMod.FnumberTurret Check
                If IsNumeric(strFnumber_buf) = False And Left(.OptIdentifier, 1) = "P" And strFnumber_buf <> "PIN" Then
                    MsgBox "UnMatch [" & .OptIdentifier & "]@Identifier - [" & strFnumber_buf & "]@OptMod.FnumberTurret" & vbCrLf & vbCrLf & "MustBe [PIN]@OptMod.FnumberTurret!!!!!" & vbCrLf & "ErrCode.4-3-02-5-18", vbExclamation, "Error@ReadOptCondInfo"
                ElseIf IsNumeric(strFnumber_buf) = False And Left(.OptIdentifier, 1) = "F" And strFnumber_buf <> "F_UNIT" Then
                    MsgBox "UnMatch [" & .OptIdentifier & "]@Identifier - [" & strFnumber_buf & "]@OptMod.FnumberTurret" & vbCrLf & vbCrLf & "MustBe [F_UNIT]@OptMod.FnumberTurret!!!!!" & vbCrLf & "ErrCode.4-3-02-5-19", vbExclamation, "Error@ReadOptCondInfo"
                End If
                
                'Color Check
                If IsNumeric(strColor_buf) = False And Right(.OptIdentifier, 4) = "_RED" And strColor_buf <> "RED" Then
                    MsgBox "UnMatch [" & .OptIdentifier & "]@Identifier - [" & strColor_buf & "]@Optset.Color" & vbCrLf & vbCrLf & "MustBe [RED]@Optset.Color!!!!!" & vbCrLf & "ErrCode.4-3-02-5-20", vbExclamation, "Error@ReadOptCondInfo"
                ElseIf IsNumeric(strColor_buf) = False And Right(.OptIdentifier, 6) = "_GREEN" And strColor_buf <> "GREEN" Then
                    MsgBox "UnMatch [" & .OptIdentifier & "]@Identifier - [" & strColor_buf & "]@Optset.Color" & vbCrLf & vbCrLf & "MustBe [GREEN]@Optset.Color!!!!!" & vbCrLf & "ErrCode.4-3-02-5-21", vbExclamation, "Error@ReadOptCondInfo"
                ElseIf IsNumeric(strColor_buf) = False And Right(.OptIdentifier, 5) = "_BLUE" And strColor_buf <> "BLUE" Then
                    MsgBox "UnMatch [" & .OptIdentifier & "]@Identifier - [" & strColor_buf & "]@Optset.Color" & vbCrLf & vbCrLf & "MustBe [BLUE]@Optset.Color!!!!!" & vbCrLf & "ErrCode.4-3-02-5-22", vbExclamation, "Error@ReadOptCondInfo"
                ElseIf IsNumeric(strColor_buf) = False And Right(.OptIdentifier, 3) = "_LB" And strColor_buf <> "LB" Then
                    MsgBox "UnMatch [" & .OptIdentifier & "]@Identifier - [" & strColor_buf & "]@Optset.Color" & vbCrLf & vbCrLf & "MustBe [LB]@Optset.Color!!!!!" & vbCrLf & "ErrCode.4-3-02-5-23", vbExclamation, "Error@ReadOptCondInfo"
                ElseIf IsNumeric(strColor_buf) = False And Right(.OptIdentifier, 3) = "_FL" And strColor_buf <> "FL" Then
                    MsgBox "UnMatch [" & .OptIdentifier & "]@Identifier - [" & strColor_buf & "]@Optset.Color" & vbCrLf & vbCrLf & "MustBe [FL]@Optset.Color!!!!!" & vbCrLf & "ErrCode.4-3-02-5-24", vbExclamation, "Error@ReadOptCondInfo"
                ElseIf IsNumeric(strColor_buf) = False And InStr(1, .OptIdentifier, "_") = 0 And strColor_buf <> "" And strColor_buf <> "WHITE" Then
                    MsgBox "UnMatch [" & .OptIdentifier & "]@Identifier - [" & strColor_buf & "]@Optset.Color" & vbCrLf & vbCrLf & "MustBe [WHITE]@Optset.Color!!!!!" & vbCrLf & "ErrCode.4-3-02-5-25", vbExclamation, "Error@ReadOptCondInfo"
                End If
            End If
            
            For intChkCnt = 0 To Optcnt - 1
                If .OptIdentifier = xOptCondInfo(intChkCnt).OptIdentifier Then
                    MsgBox "Duplicate [" & .OptIdentifier & "]@Identifier @Row=" & nodeRow & "!!!!!" & vbCrLf & "ErrCode.4-3-02-5-26", vbExclamation, "Error@ReadOptCondInfo"
                End If
                If .AxisRegisterAxis = xOptCondInfo(intChkCnt).AxisRegisterAxis And .AxisRegisterAxis <> -1 Then
                    MsgBox "Duplicate [" & .AxisRegisterAxis & "]@AxisRegisterAxis @" & .OptIdentifier & " @Row=" & nodeRow & "!!!!!" & vbCrLf & "ErrCode.4-3-02-5-27", vbExclamation, "Error@ReadOptCondInfo"
                End If
            Next
            If .AxisRegisterAxis > 150 Then
                MsgBox "> 150 [" & .AxisRegisterAxis & "]@AxisRegisterAxis  @" & .OptIdentifier & " @Row=" & nodeRow & "!!!!!" & vbCrLf & "ErrCode.4-3-02-5-28", vbExclamation, "Error@ReadOptCondInfo"
            End If

            End With

            Optcnt = Optcnt + 1
        End If
        '*****************************************
        nodeRow = nodeRow + 1
    Loop
    
    Set wkshtObj = Nothing 'Release Memory
    Exit Sub

errorLable:
        Set wkshtObj = Nothing 'Release Memory
        MsgBox "Not Available " & axisCmd & "[" & cmdArgstr & "]@ReadOptCondInfo" & vbCrLf & "ErrCode.4-3-02-0-29", vbExclamation, "Error@ReadOptCondInfo"

End Sub
 '=============================
'========== COMMON ===========
'=============================
Public Sub ReadOptListInfo()
    
    Dim wkshtObj As Object
    Dim basePoint As Variant
    
    Dim ModelPoint As Variant
    Dim AddressPoint As Variant
    Dim MakerPoint As Variant
    
    Dim nodeRow As Long
    Dim nodeColumn As Long

    On Error GoTo EndOpt

    '======= WorkSheet Select ========
    Set wkshtObj = ThisWorkbook.Sheets(ILLUM_MAKER_LIST_WRKSHT_NAME)
         
    '======= WorkSheet ErrorProcess ========
    If IsEmpty(wkshtObj) Then
        MsgBox "Non IllumLst WorkSheet For Sw_Node Select!" & vbCrLf & "ErrCode.4-3-03-8-30", vbExclamation, "Error@ReadOptListInfo"
        Exit Sub
    End If

    '======= Base Point Find ========
    Set basePoint = Worksheets(ILLUM_MAKER_LIST_WRKSHT_NAME).Range("A1:K100").Find("Sw_node")
    If basePoint Is Nothing Then
        MsgBox "Non IllumLst WorkSheet For Find Sw_node!" & vbCrLf & "ErrCode.4-3-03-8-31", vbExclamation, "Error@ReadOptListInfo"
        Exit Sub
    End If

    '======= Model Point Find ========
    Set MakerPoint = Worksheets(ILLUM_MAKER_LIST_WRKSHT_NAME).Range("A1:K100").Find("Maker")
    If MakerPoint Is Nothing Then
        MsgBox "Non IllumLst WorkSheet For Find Maker!" & vbCrLf & "ErrCode.4-3-03-8-32", vbExclamation, "Error@ReadOptListInfo"
        Exit Sub
    End If
    '======= Model Point Find ========
    Set ModelPoint = Worksheets(ILLUM_MAKER_LIST_WRKSHT_NAME).Range("A1:K100").Find("Model")
    If ModelPoint Is Nothing Then
        MsgBox "Non IllumLst WorkSheet For Find Model!" & vbCrLf & "ErrCode.4-3-03-8-33", vbExclamation, "Error@ReadOptListInfo"
        Exit Sub
    End If
    '======= Model Point Find ========
    Set AddressPoint = Worksheets(ILLUM_MAKER_LIST_WRKSHT_NAME).Range("A1:K100").Find("GPIB Address")
    If AddressPoint Is Nothing Then
        MsgBox "Non IllumLst WorkSheet For GPIB Address!" & vbCrLf & "ErrCode.4-3-03-8-34", vbExclamation, "Error@ReadOptListInfo"
        Exit Sub
    End If

    '======= Search Node Start Addres Define ========
    nodeRow = basePoint.Row + 1
    nodeColumn = basePoint.Column

    '======= Search Node ========
    Do While Not wkshtObj.Cells(nodeRow, nodeColumn) = vbNullString
        If wkshtObj.Cells(nodeRow, nodeColumn) = Sw_Node Then
            IllumMaker = wkshtObj.Cells(nodeRow, MakerPoint.Column)
            IllumModel = wkshtObj.Cells(nodeRow, ModelPoint.Column)
            IllumGpibAddr = wkshtObj.Cells(nodeRow, AddressPoint.Column)
            If IllumGpibAddr = 0 Then
                '======== Default Setting ========
                If IllumMaker = NIKON Then
                    IllumGpibAddr = DEFAULT_ILLUM_GPIB_ADDR_NIKON
                ElseIf IllumMaker = INTERACTION Or IllumMaker = KESILLUM Then
                    IllumGpibAddr = DEFAULT_ILLUM_GPIB_ADDR_IA
                End If
                '=================================
            End If
            
            Call sub_IllumInfoSet   'GetIllumInfo
            
            Exit Sub
        End If
        nodeRow = nodeRow + 1
    Loop

EndOpt:
    MsgBox "Sw_node Error@ ReadOptListInfo!" & vbCrLf & "ErrCode.4-3-03-0-35", vbSystemModal, "Sw_node Error"

End Sub

Public Function OptCondNum() As Integer

    OptCondNum = UBound(xOptCondInfo)

End Function

