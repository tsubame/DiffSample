Attribute VB_Name = "ICUL1G_ReadSheet_Mod"
Option Explicit

Public Sub get_idelay(ByVal MipiKeyName As String)

    Dim wkshtObj As Object
    
    Dim nChans As Long
    Dim nSites As Long
    Dim errMsg As String
    Dim site As Long
    
    Dim Lane(4) As String

    Lane(0) = "DataLane00"
    Lane(1) = "DataLane01"
    Lane(2) = "DataLane02"
    Lane(3) = "DataLane03"
    Lane(4) = "ClockLane"
   
    '======= WorkSheet Select ========
    Set wkshtObj = ThisWorkbook.Sheets(MipiKeyName)
         
    '======= WorkSheet ErrorProcess ========
    If IsEmpty(wkshtObj) Then
        MsgBox "Non [" & MipiKeyName & "] WorkSheet"
        Exit Sub
    End If

    Dim StartPoint_Row As Long
    Dim StartPoint_Column As Long
    Dim Site_Row As Long
    Dim Site_Column As Long
    Dim nodePoint As Variant
    Dim LanePoint As Variant
    
    Dim i As Long
    StartPoint_Row = 16
    StartPoint_Column = 2
    
    For i = 0 To 4
        If Lane(i) <> "-" Then
            '======= SwNode Find ========
            Set nodePoint = Worksheets(MipiKeyName).Range(wkshtObj.Cells(StartPoint_Row, StartPoint_Column), wkshtObj.Cells(StartPoint_Row + 500, StartPoint_Column)).Find(Sw_Node)
            If nodePoint Is Nothing Then
                MsgBox "Search Error! Not Finding SwNodee @[" & MipiKeyName & "] WorkSheet"
                Exit Sub
            End If

            '======= ICUL1G_Board Find ========
            Set LanePoint = Worksheets(MipiKeyName).Range(wkshtObj.Cells(nodePoint.Row, nodePoint.Column + 1), wkshtObj.Cells(nodePoint.Row + 4, nodePoint.Column + 1)).Find(Lane(i))
            If LanePoint Is Nothing Then
                MsgBox "Search Error! Not Finding LVDS Namee @[" & MipiKeyName & "] WorkSheet"
                Exit Sub
            End If

            With MipiSetFor1G(getMipiNum(MipiKeyName))
                Select Case Lane(i)
                 Case "ClockLane"
                    For site = 0 To nSite
                        .UserDelayCLK(site) = wkshtObj.Cells(LanePoint.Row, LanePoint.Column + 1 + site)
                    Next site
                 Case "DataLane00"
                    For site = 0 To nSite
                        .UserDelay00(site) = wkshtObj.Cells(LanePoint.Row, LanePoint.Column + 1 + site)
                    Next site
                 Case "DataLane01"
                    For site = 0 To nSite
                        .UserDelay01(site) = wkshtObj.Cells(LanePoint.Row, LanePoint.Column + 1 + site)
                    Next site
                 Case "DataLane02"
                    For site = 0 To nSite
                        .UserDelay02(site) = wkshtObj.Cells(LanePoint.Row, LanePoint.Column + 1 + site)
                    Next site
                 Case "DataLane03"
                    For site = 0 To nSite
                        .UserDelay03(site) = wkshtObj.Cells(LanePoint.Row, LanePoint.Column + 1 + site)
                    Next site
                End Select
            End With
        End If
    Next i
                            
End Sub

Public Sub get_VOD(ByVal MipiKeyName As String)

    Dim wkshtObj As Object
    
    Dim nChans As Long
    Dim nSites As Long
    Dim errMsg As String
    Dim site As Long
    
    Dim Lane(4) As String

    Lane(0) = "DCK"
    Lane(1) = "DO0"
    Lane(2) = "DO1"
    Lane(3) = "DO2"
    Lane(4) = "DO3"
   
    '======= WorkSheet Select ========
    Set wkshtObj = ThisWorkbook.Sheets(MipiKeyName)
         
    '======= WorkSheet ErrorProcess ========
    If IsEmpty(wkshtObj) Then
        MsgBox "Non [" & MipiKeyName & "] WorkSheet"
        Exit Sub
    End If

    Dim StartPoint_Row As Long
    Dim StartPoint_Column As Long
    Dim Site_Row As Long
    Dim Site_Column As Long
    Dim nodePoint As Variant
    Dim LanePoint As Variant
    
    Dim i As Long
    StartPoint_Row = 26
    StartPoint_Column = 2
    
    For i = 0 To 4
        If Lane(i) <> "-" Then
            '======= SwNode Find ========
            Set nodePoint = Worksheets(MipiKeyName).Range(wkshtObj.Cells(StartPoint_Row, StartPoint_Column), wkshtObj.Cells(StartPoint_Row + 500, StartPoint_Column)).Find(Sw_Node)
            If nodePoint Is Nothing Then
                MsgBox "Search Error! Not Finding SwNodee @[" & MipiKeyName & "] WorkSheet"
                Exit Sub
            End If

            '======= ICUL1G_Board Find ========
            Set LanePoint = Worksheets(MipiKeyName).Range(wkshtObj.Cells(nodePoint.Row, nodePoint.Column + 1), wkshtObj.Cells(nodePoint.Row + 4, nodePoint.Column + 1)).Find(Lane(i))
            If LanePoint Is Nothing Then
                MsgBox "Search Error! Not Finding LVDS Namee @[" & MipiKeyName & "] WorkSheet"
                Exit Sub
            End If

            With MipiSetFor1G(getMipiNum(MipiKeyName))
                Select Case Lane(i)
                 Case "DCK"
                    For site = 0 To nSite
                        .VodSetCLK(site) = wkshtObj.Cells(LanePoint.Row, LanePoint.Column + 1 + site)
                    Next site
                 Case "DO0"
                    For site = 0 To nSite
                        .VodSet00(site) = wkshtObj.Cells(LanePoint.Row, LanePoint.Column + 1 + site)
                    Next site
                 Case "DO1"
                    For site = 0 To nSite
                        .VodSet01(site) = wkshtObj.Cells(LanePoint.Row, LanePoint.Column + 1 + site)
                    Next site
                 Case "DO2"
                    For site = 0 To nSite
                        .VodSet02(site) = wkshtObj.Cells(LanePoint.Row, LanePoint.Column + 1 + site)
                    Next site
                 Case "DO3"
                    For site = 0 To nSite
                        .VodSet03(site) = wkshtObj.Cells(LanePoint.Row, LanePoint.Column + 1 + site)
                    Next site
                End Select
            End With
        End If
    Next i
                            
End Sub
Public Sub get_Threshold(ByVal MipiKeyName As String)

    Dim wkshtObj As Object
    Dim ICUL1G_Board(5) As String
    Dim ListThreshold(5) As Double

    ICUL1G_Board(0) = "16"
    ICUL1G_Board(1) = "19"
    
    '======= WorkSheet Select ========
    Set wkshtObj = ThisWorkbook.Sheets(MipiKeyName)
         
    '======= WorkSheet ErrorProcess ========
    If IsEmpty(wkshtObj) Then
        MsgBox "Non [" & MipiKeyName & "] WorkSheet"
        Exit Sub
    End If

    Dim StartPoint_Row As Long
    Dim StartPoint_Column As Long
    Dim Site_Row As Long
    Dim Site_Column As Long
    Dim nodePoint As Variant
    Dim BoardPoint As Variant

    Dim i As Long
    StartPoint_Row = 36
    StartPoint_Column = 2
    
    For i = 0 To 1
        If ICUL1G_Board(i) <> "-" Then
            '======= SwNode Find ========
            Set nodePoint = Worksheets(MipiKeyName).Range(wkshtObj.Cells(StartPoint_Row, StartPoint_Column), wkshtObj.Cells(StartPoint_Row + 500, StartPoint_Column)).Find(Sw_Node)
            If nodePoint Is Nothing Then
                MsgBox "Search Error! Not Finding SwNode @[" & MipiKeyName & "] WorkSheet"
                Exit Sub
            End If

            '======= ICUL1G_Board Find ========
            Set BoardPoint = Worksheets(MipiKeyName).Range(wkshtObj.Cells(StartPoint_Row + 1, StartPoint_Column), wkshtObj.Cells(StartPoint_Row + 1, StartPoint_Column + 3)).Find(ICUL1G_Board(i))
            If BoardPoint Is Nothing Then
                MsgBox "Search Error! Not Finding LVDS Name @[" & MipiKeyName & "] WorkSheet"
                Exit Sub
            End If

            Site_Row = nodePoint.Row
            Site_Column = BoardPoint.Column

            ListThreshold(i) = wkshtObj.Cells(Site_Row, Site_Column)
        End If
    Next i

    With MipiSetFor1G(getMipiNum(MipiKeyName))
        .Threshold_Board16 = ListThreshold(0) 'Bord16
        .Threshold_Board19 = ListThreshold(1) 'Bord19
    End With
                
End Sub


