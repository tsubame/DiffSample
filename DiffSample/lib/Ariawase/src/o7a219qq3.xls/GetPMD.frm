VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GetPMD 
   Caption         =   "Get PMD"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12915
   OleObjectBlob   =   "GetPMD.frx":0000
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
End
Attribute VB_Name = "GetPMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'Revision History:
'Jun 23/2005 N.TOGO @IS TEST : Created this Module
'May 30/2008 N.TOGO @IS TEST : Modify  for EeeJobV1.2 JOB
'May 01/2009 N.TOGO @IS TEST : Modify  for EeeJobV2.0 JOB

Const SheetName = "PMD Definition"

Private HeightArray(20) As Double
Private WidthArray(20) As Double
Private XstartArray(20) As Double
Private YstartArray(20) As Double
Private XendArray(20) As Double
Private YendArray(20) As Double
Private PMDName(20) As String

Const DrawWidth As Double = 750
Const InitialHeight As Double = 90
Const InitialWidth As Double = 650

Private Sub UserForm_Initialize()
    GetPMD.height = InitialHeight
    GetPMD.width = InitialWidth
End Sub

Private Sub CmdAddress_Click()

    Dim i As Integer
    Dim PMD As String
    Dim Xstart As Long
    Dim Ystart As Long
    Dim ColSize As Long
    Dim RowSize As Long
    Dim Xend As Long
    Dim Yend As Long

    Label1.Caption = "PMD Address for IP750"
    
    i = 0
    If Txstart.Text <> "" Then Xstart = Txstart.Text: i = i + 1
    If Txend.Text <> "" Then Xend = Txend.Text: i = i + 1
    If Twidth.Text <> "" Then ColSize = Twidth.Text: i = i + 1
    Select Case i
        Case 0, 1
            MsgBox "The input data is not enough!": Exit Sub
        Case 2
            If Xstart = 0 Then Xstart = Xend - ColSize + 1
            If Xend = 0 Then Xend = Xstart + ColSize - 1
            If ColSize = 0 Then ColSize = Xend - Xstart + 1
        Case 3
            Xend = Xstart + ColSize - 1
    End Select
        
    If Xstart <= 0 Then MsgBox "Error! Xstart <= 0": Exit Sub
    If Xend <= 0 Then MsgBox "Error! Xend <= 0": Exit Sub
    If ColSize <= 0 Then MsgBox "Error! Width <= 0": Exit Sub
    If Xend < Xstart Then MsgBox "Error! Xend < Xstart": Exit Sub
    
    i = 0
    If Tystart.Text <> "" Then Ystart = Tystart.Text: i = i + 1
    If Tyend.Text <> "" Then Yend = Tyend.Text: i = i + 1
    If Theight.Text <> "" Then RowSize = Theight.Text: i = i + 1
    Select Case i
        Case 0, 1
            MsgBox "The input data is not enough!": Exit Sub
        Case 2
            If Ystart = 0 Then Ystart = Yend - RowSize + 1
            If Yend = 0 Then Yend = Ystart + RowSize - 1
            If RowSize = 0 Then RowSize = Yend - Ystart + 1
        Case 3
            Yend = Ystart + RowSize - 1
    End Select
        
    If Ystart <= 0 Then MsgBox "Error! Ystart <= 0": Exit Sub
    If Yend <= 0 Then MsgBox "Error! Yend <= 0": Exit Sub
    If RowSize <= 0 Then MsgBox "Error! Height <= 0": Exit Sub
    If Yend < Ystart Then MsgBox "Error! Yend < Ystart": Exit Sub
        
    PMD = "" & Xstart & ", " & Ystart & ", " & ColSize & ", " & RowSize
    TxtPMD.Text = PMD
    Txstart.Text = Xstart
    Txend.Text = Xend
    Twidth.Text = ColSize
    Tystart.Text = Ystart
    Tyend.Text = Yend
    Theight.Text = RowSize

End Sub

Private Sub CmdExp_Click()

    Dim PMD As String
    Dim ExpBit As String
    Dim Xstart As Long
    Dim Ystart As Long
    Dim ColSize As Long
    Dim RowSize As Long
    Dim ExXstart As Long
    Dim ExYstart As Long
    Dim ExColsize As Long
    Dim ExRowsize As Long
    Dim Xend As Long
    Dim Yend As Long
    Dim PmdIn1 As String
    Dim PmdOut1 As String
    Dim PmdIn2 As String
    Dim PmdOut2 As String
    Dim BaseX As Long
    Dim BaseY As Long
    Dim baseWidth As Long
    Dim baseHeight As Long
    Dim Flg_OutCheck As Boolean

    PMD = TxtPMD.Text
    ExpBit = TxtExp.Text
    If IsNumeric(ExpBit) = False Or IsNull(ExpBit) = True Then
        MsgBox "Please Input ExpansionBit Number!"
        Exit Sub
    End If
    ExpBit = CLng(ExpBit)
    
    Label1.Caption = "PMD Name"
    Call GetPmdInfoTool(PMD, Xstart, Ystart, ColSize, RowSize, BaseX, BaseY, baseWidth, baseHeight)
    Xend = Xstart + ColSize - 1
    Yend = Ystart + RowSize - 1
    
    If OpbV.Value = True Then
        LPmdIn1.Caption = "Top IN"
        ExXstart = Xstart: ExYstart = Ystart: ExColsize = ColSize: ExRowsize = ExpBit
        PmdIn1 = "" & ExXstart & ", " & ExYstart & ", " & ExColsize & ", " & ExRowsize
        If ExXstart < 1 Or ExYstart < 1 Or (ExXstart + ExColsize - 1) > baseWidth _
        Or (ExYstart + ExRowsize - 1) > baseHeight Then LOut.Visible = True
        
        LPmdOut1.Caption = "Top OUT"
        ExXstart = Xstart: ExYstart = Ystart - ExpBit: ExColsize = ColSize: ExRowsize = ExpBit
        PmdOut1 = "" & ExXstart & ", " & ExYstart & ", " & ExColsize & ", " & ExRowsize
        If ExXstart < 1 Or ExYstart < 1 Or (ExXstart + ExColsize - 1) > baseWidth _
        Or (ExYstart + ExRowsize - 1) > baseHeight Then LOut.Visible = True
        
        LPmdIn2.Caption = "Bottom IN"
        ExXstart = Xstart: ExYstart = Yend - ExpBit + 1: ExColsize = ColSize: ExRowsize = ExpBit
        PmdIn2 = "" & ExXstart & ", " & ExYstart & ", " & ExColsize & ", " & ExRowsize
        If ExXstart < 1 Or ExYstart < 1 Or (ExXstart + ExColsize - 1) > baseWidth _
        Or (ExYstart + ExRowsize - 1) > baseHeight Then LOut.Visible = True
        
        LPmdOut2.Caption = "Bottom OUT"
        ExXstart = Xstart: ExYstart = Yend + 1: ExColsize = ColSize: ExRowsize = ExpBit
        PmdOut2 = "" & ExXstart & ", " & ExYstart & ", " & ExColsize & ", " & ExRowsize
        If ExXstart < 1 Or ExYstart < 1 Or (ExXstart + ExColsize - 1) > baseWidth _
        Or (ExYstart + ExRowsize - 1) > baseHeight Then LOut.Visible = True
        
    ElseIf OpbH.Value = True Then
        LPmdIn1.Caption = "Left IN"
        ExXstart = Xstart: ExYstart = Ystart: ExColsize = ExpBit: ExRowsize = RowSize
        PmdIn1 = "" & ExXstart & ", " & ExYstart & ", " & ExColsize & ", " & ExRowsize
        If ExXstart < 1 Or ExYstart < 1 Or (ExXstart + ExColsize - 1) > baseWidth _
        Or (ExYstart + ExRowsize - 1) > baseHeight Then LOut.Visible = True
        
        LPmdOut1.Caption = "Left OUT"
        ExXstart = Xstart - ExpBit: ExYstart = Ystart: ExColsize = ExpBit: ExRowsize = RowSize
        PmdOut1 = "" & ExXstart & ", " & ExYstart & ", " & ExColsize & ", " & ExRowsize
        If ExXstart < 1 Or ExYstart < 1 Or (ExXstart + ExColsize - 1) > baseWidth _
        Or (ExYstart + ExRowsize - 1) > baseHeight Then LOut.Visible = True
        
        LPmdIn2.Caption = "Right IN"
        ExXstart = Xend - ExpBit + 1: ExYstart = Ystart: ExColsize = ExpBit: ExRowsize = RowSize
        PmdIn2 = "" & ExXstart & ", " & ExYstart & ", " & ExColsize & ", " & ExRowsize
        If ExXstart < 1 Or ExYstart < 1 Or (ExXstart + ExColsize - 1) > baseWidth _
        Or (ExYstart + ExRowsize - 1) > baseHeight Then LOut.Visible = True
        
        LPmdOut2.Caption = "Right OUT"
        ExXstart = Xend + 1: ExYstart = Ystart: ExColsize = ExpBit: ExRowsize = RowSize
        PmdOut2 = "" & ExXstart & ", " & ExYstart & ", " & ExColsize & ", " & ExRowsize
        If ExXstart < 1 Or ExYstart < 1 Or (ExXstart + ExColsize - 1) > baseWidth _
        Or (ExYstart + ExRowsize - 1) > baseHeight Then LOut.Visible = True
        
    End If
    
    GetPMD.height = 183
    GetPMD.width = InitialWidth
    Lpad1.Visible = False
    Lpad2.Visible = False
    Lpad3.Visible = False
    Lpad4.Visible = False
    Lpad5.Visible = False
    Lpad6.Visible = False
    Lpad7.Visible = False
    Lpad8.Visible = False
    Lpad9.Visible = False
    Lpad10.Visible = False
    Lpad11.Visible = False
    Lpad12.Visible = False
    Lpad13.Visible = False
    NoMore.Visible = False

    LPmdIn1.Visible = True
    LPmdOut1.Visible = True
    LPmdIn2.Visible = True
    LPmdOut2.Visible = True
    TxtPmdIn1.Visible = True
    TxtPmdIn2.Visible = True
    TxtPmdOut1.Visible = True
    TxtPmdOut2.Visible = True
    TxtPmdIn1.Text = PmdIn1
    TxtPmdIn2.Text = PmdIn2
    TxtPmdOut1.Text = PmdOut1
    TxtPmdOut2.Text = PmdOut2
        
End Sub

Private Sub CmdGet_Click()

    Dim PMD As String
    Dim Xstart As Long
    Dim Ystart As Long
    Dim ColSize As Long
    Dim RowSize As Long
    Dim Xend As Long
    Dim Yend As Long
    Dim BaseX As Long
    Dim BaseY As Long
    Dim baseWidth As Long
    Dim baseHeight As Long

    PMD = UCase(TxtPMD.Text)
    
    Label1.Caption = "PMD Name"

    Call GetPmdInfoTool(PMD, Xstart, Ystart, ColSize, RowSize, BaseX, BaseY, baseWidth, baseHeight)
    If baseWidth = 0 Then Exit Sub
    
    Xend = Xstart + ColSize - 1
    Yend = Ystart + RowSize - 1
    
    Txstart.Text = Xstart
    Txend.Text = Xend
    Twidth.Text = ColSize
    Tystart.Text = Ystart
    Tyend.Text = Yend
    Theight.Text = RowSize
    
    LPmdIn1.Visible = False
    LPmdOut1.Visible = False
    LPmdIn2.Visible = False
    LPmdOut2.Visible = False
    TxtPmdIn1.Visible = False
    TxtPmdIn2.Visible = False
    TxtPmdOut1.Visible = False
    TxtPmdOut2.Visible = False
    LOut.Visible = False
    
    Call DrawArea(PMD, Xstart, Ystart, Xend, Yend, ColSize, RowSize, BaseX, BaseY, baseWidth, baseHeight)
    
End Sub

Private Sub DrawArea(ByVal PMD As String, ByVal Xstart As Long, ByVal Ystart As Long, ByVal Xend As Long, ByVal Yend As Long, _
                    ByVal ColSize As Long, ByVal RowSize As Long, _
                    ByVal BaseX As Long, ByVal BaseY As Long, ByVal baseWidth As Long, ByVal baseHeight As Long)

    Dim BasisX As Double
    Dim BasisY As Double
    Dim OnePix As Double
    
    BasisX = 6  'Draw X Start Address
    BasisY = 70 'Draw Y Start Address
    
    OnePix = DrawWidth / baseWidth
    GetPMD.width = BasisX + DrawWidth + 15
    GetPMD.height = BasisY + OnePix * baseHeight + 25
    
    'Base PMD Writing
    With Lpad1
        .Visible = True
        .Left = BasisX
        .Top = BasisY
        .width = OnePix * baseWidth
        .height = OnePix * baseHeight
        HeightArray(1) = baseHeight
        WidthArray(1) = baseWidth
        XstartArray(1) = BaseX
        YstartArray(1) = BaseY
        XendArray(1) = BaseX + baseWidth - 1
        YendArray(1) = BaseY + baseHeight - 1
        PMDName(1) = "BASE PMD"
    End With
    
    With Lpad2
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(2) = RowSize
        WidthArray(2) = ColSize
        XstartArray(2) = Xstart
        YstartArray(2) = Ystart
        XendArray(2) = Xend
        YendArray(2) = Yend
        PMDName(2) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad3
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(3) = RowSize
        WidthArray(3) = ColSize
        XstartArray(3) = Xstart
        YstartArray(3) = Ystart
        XendArray(3) = Xend
        YendArray(3) = Yend
        PMDName(3) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad4
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(4) = RowSize
        WidthArray(4) = ColSize
        XstartArray(4) = Xstart
        YstartArray(4) = Ystart
        XendArray(4) = Xend
        YendArray(4) = Yend
        PMDName(4) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad5
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(5) = RowSize
        WidthArray(5) = ColSize
        XstartArray(5) = Xstart
        YstartArray(5) = Ystart
        XendArray(5) = Xend
        YendArray(5) = Yend
        PMDName(5) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad6
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(6) = RowSize
        WidthArray(6) = ColSize
        XstartArray(6) = Xstart
        YstartArray(6) = Ystart
        XendArray(6) = Xend
        YendArray(6) = Yend
        PMDName(6) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad7
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(7) = RowSize
        WidthArray(7) = ColSize
        XstartArray(7) = Xstart
        YstartArray(7) = Ystart
        XendArray(7) = Xend
        YendArray(7) = Yend
        PMDName(7) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad8
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(8) = RowSize
        WidthArray(8) = ColSize
        XstartArray(8) = Xstart
        YstartArray(8) = Ystart
        XendArray(8) = Xend
        YendArray(8) = Yend
        PMDName(8) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad9
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(9) = RowSize
        WidthArray(9) = ColSize
        XstartArray(9) = Xstart
        YstartArray(9) = Ystart
        XendArray(9) = Xend
        YendArray(9) = Yend
        PMDName(9) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad10
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(10) = RowSize
        WidthArray(10) = ColSize
        XstartArray(10) = Xstart
        YstartArray(10) = Ystart
        XendArray(10) = Xend
        YendArray(10) = Yend
        PMDName(10) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad11
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(11) = RowSize
        WidthArray(11) = ColSize
        XstartArray(11) = Xstart
        YstartArray(11) = Ystart
        XendArray(11) = Xend
        YendArray(11) = Yend
        PMDName(11) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad12
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(12) = RowSize
        WidthArray(12) = ColSize
        XstartArray(12) = Xstart
        YstartArray(12) = Ystart
        XendArray(12) = Xend
        YendArray(12) = Yend
        PMDName(12) = PMD
        Exit Sub
        End If
    End With
        
    With Lpad13
        If .Visible = False Then
        .Visible = True
        .Caption = PMD
        .Left = BasisX + (OnePix * (Xstart - 1))
        .Top = BasisY + (OnePix * (Ystart - 1))
        .width = OnePix * ColSize
        .height = OnePix * RowSize
        HeightArray(13) = RowSize
        WidthArray(13) = ColSize
        XstartArray(13) = Xstart
        YstartArray(13) = Ystart
        XendArray(13) = Xend
        YendArray(13) = Yend
        PMDName(13) = PMD
        NoMore.Visible = True
        NoMore.Left = 5
        NoMore.Top = 0
        Exit Sub
        End If
    End With
        
End Sub

Private Sub Lpad1_Click()

    Txstart.Text = XstartArray(1)
    Txend.Text = XendArray(1)
    Twidth.Text = WidthArray(1)
    Tystart.Text = YstartArray(1)
    Tyend.Text = YendArray(1)
    Theight.Text = HeightArray(1)
    TxtPMD.Text = PMDName(1)
    
End Sub

Private Sub Lpad2_Click()

    Txstart.Text = XstartArray(2)
    Txend.Text = XendArray(2)
    Twidth.Text = WidthArray(2)
    Tystart.Text = YstartArray(2)
    Tyend.Text = YendArray(2)
    Theight.Text = HeightArray(2)
    TxtPMD.Text = PMDName(2)
    
End Sub

Private Sub Lpad3_Click()

    Txstart.Text = XstartArray(3)
    Txend.Text = XendArray(3)
    Twidth.Text = WidthArray(3)
    Tystart.Text = YstartArray(3)
    Tyend.Text = YendArray(3)
    Theight.Text = HeightArray(3)
    TxtPMD.Text = PMDName(3)
    
End Sub

Private Sub Lpad4_Click()

    Txstart.Text = XstartArray(4)
    Txend.Text = XendArray(4)
    Twidth.Text = WidthArray(4)
    Tystart.Text = YstartArray(4)
    Tyend.Text = YendArray(4)
    Theight.Text = HeightArray(4)
    TxtPMD.Text = PMDName(4)
    
End Sub

Private Sub Lpad5_Click()

    Txstart.Text = XstartArray(5)
    Txend.Text = XendArray(5)
    Twidth.Text = WidthArray(5)
    Tystart.Text = YstartArray(5)
    Tyend.Text = YendArray(5)
    Theight.Text = HeightArray(5)
    TxtPMD.Text = PMDName(5)
    
End Sub

Private Sub Lpad6_Click()

    Txstart.Text = XstartArray(6)
    Txend.Text = XendArray(6)
    Twidth.Text = WidthArray(6)
    Tystart.Text = YstartArray(6)
    Tyend.Text = YendArray(6)
    Theight.Text = HeightArray(6)
    TxtPMD.Text = PMDName(6)
    
End Sub

Private Sub Lpad7_Click()

    Txstart.Text = XstartArray(7)
    Txend.Text = XendArray(7)
    Twidth.Text = WidthArray(7)
    Tystart.Text = YstartArray(7)
    Tyend.Text = YendArray(7)
    Theight.Text = HeightArray(7)
    TxtPMD.Text = PMDName(7)
    
End Sub

Private Sub Lpad8_Click()

    Txstart.Text = XstartArray(8)
    Txend.Text = XendArray(8)
    Twidth.Text = WidthArray(8)
    Tystart.Text = YstartArray(8)
    Tyend.Text = YendArray(8)
    Theight.Text = HeightArray(8)
    TxtPMD.Text = PMDName(8)
    
End Sub

Private Sub Lpad9_Click()

    Txstart.Text = XstartArray(9)
    Txend.Text = XendArray(9)
    Twidth.Text = WidthArray(9)
    Tystart.Text = YstartArray(9)
    Tyend.Text = YendArray(9)
    Theight.Text = HeightArray(9)
    TxtPMD.Text = PMDName(9)
    
End Sub

Private Sub Lpad10_Click()

    Txstart.Text = XstartArray(10)
    Txend.Text = XendArray(10)
    Twidth.Text = WidthArray(10)
    Tystart.Text = YstartArray(10)
    Tyend.Text = YendArray(10)
    Theight.Text = HeightArray(10)
    TxtPMD.Text = PMDName(10)
    
End Sub

Private Sub Lpad11_Click()

    Txstart.Text = XstartArray(11)
    Txend.Text = XendArray(11)
    Twidth.Text = WidthArray(11)
    Tystart.Text = YstartArray(11)
    Tyend.Text = YendArray(11)
    Theight.Text = HeightArray(11)
    TxtPMD.Text = PMDName(11)
    
End Sub

Private Sub Lpad12_Click()

    Txstart.Text = XstartArray(12)
    Txend.Text = XendArray(12)
    Twidth.Text = WidthArray(12)
    Tystart.Text = YstartArray(12)
    Tyend.Text = YendArray(12)
    Theight.Text = HeightArray(12)
    TxtPMD.Text = PMDName(12)
    
End Sub

Private Sub Lpad13_Click()

    Txstart.Text = XstartArray(13)
    Txend.Text = XendArray(13)
    Twidth.Text = WidthArray(13)
    Tystart.Text = YstartArray(13)
    Tyend.Text = YendArray(13)
    Theight.Text = HeightArray(13)
    TxtPMD.Text = PMDName(13)
    
End Sub

Private Sub CmdClear_Click()
    
    Lpad1.Visible = False
    Lpad2.Visible = False
    Lpad3.Visible = False
    Lpad4.Visible = False
    Lpad5.Visible = False
    Lpad6.Visible = False
    Lpad7.Visible = False
    Lpad8.Visible = False
    Lpad9.Visible = False
    Lpad10.Visible = False
    Lpad11.Visible = False
    Lpad12.Visible = False
    Lpad13.Visible = False
    
    LPmdIn1.Visible = False
    LPmdOut1.Visible = False
    LPmdIn2.Visible = False
    LPmdOut2.Visible = False
    TxtPmdIn1.Visible = False
    TxtPmdIn2.Visible = False
    TxtPmdOut1.Visible = False
    TxtPmdOut2.Visible = False
    LOut.Visible = False
    
    TxtPMD.Text = ""
    Txstart.Text = ""
    Txend.Text = ""
    Twidth.Text = ""
    Tystart.Text = ""
    Tyend.Text = ""
    Theight.Text = ""
    TxtExp.Text = ""

    NoMore.Visible = False
    GetPMD.height = InitialHeight
    GetPMD.width = InitialWidth

End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub LH_Click()
    OpbH.Value = True
    OpbV.Value = False
End Sub

Private Sub LV_Click()
    OpbV.Value = True
    OpbH.Value = False
End Sub

Private Sub OpbH_Click()
    OpbH.Value = True
    OpbV.Value = False
End Sub

Private Sub OpbV_Click()
    OpbV.Value = True
    OpbH.Value = False
End Sub

Private Sub GetPmdInfoTool(ByVal Pad As String, Xstart As Long, Ystart As Long, ColSize As Long, RowSize As Long, _
         BaseX As Long, BaseY As Long, baseWidth As Long, baseHeight As Long)
    
    Dim Row As Long
    Dim nStartLine As Long
    Dim nEndLine As Long
    Dim WshtObject As Object
    Set WshtObject = ActiveWorkbook.Sheets(SheetName)
    
    nStartLine = 6
    Row = nStartLine
    Do While WshtObject.Cells(Row, 8) <> ""
        Row = Row + 1
    Loop
    nEndLine = Row - 1
    
    For Row = nStartLine To nEndLine
        If UCase(WshtObject.Cells(Row, 8)) = Pad Then
            Xstart = WshtObject.Cells(Row, 9)
            Ystart = WshtObject.Cells(Row, 10)
            ColSize = WshtObject.Cells(Row, 11)
            RowSize = WshtObject.Cells(Row, 12)
            
            'BASE PMD INFO GET
            Do While WshtObject.Cells(Row, 2) = ""
                Row = Row - 1
            Loop
            BaseX = WshtObject.Cells(Row, 9)
            BaseY = WshtObject.Cells(Row, 10)
            baseWidth = WshtObject.Cells(Row, 11)
            baseHeight = WshtObject.Cells(Row, 12)
          
            GoTo EndGetPad
        End If
    Next Row
    
    MsgBox "The PAD doesn't exist!! "

EndGetPad:
End Sub

