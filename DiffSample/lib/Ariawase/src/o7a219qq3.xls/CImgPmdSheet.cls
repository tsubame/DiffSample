VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPmdSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents m_Sheet As Worksheet
Attribute m_Sheet.VB_VarHelpID = -1
Private WithEvents m_CbChecker As CommandButton
Attribute m_CbChecker.VB_VarHelpID = -1

Private Const START_CELL = "B5"
Private Const PMDNAME_INDEX_CELL = "H5"
Private Const PLANENAME_CELL = "B6"
Private Const PMDNAME_CELL = "H6"
Private Const TOP_INDEX_CELLS = "B4:M4"
Private Const TOP_PLANE_CELLS = "B4:G4"
Private Const TOP_PMD_CELLS = "H4:M4"
Private Const INDEX_CELLS = "B5:M5"

Private Const START_ROW_NUM = 6
Private Const START_COLUMN_NUM = 2
Private Const START_ROW = "6"
Private Const END_ROW = "65536"
Private Const VALID_ROWS = START_ROW & ":" & END_ROW
Private Const VALID_CELLS = "B6:M6"
Private Const VALID_INDEX_CELLS = "B5:M5"
Private Const VALID_PMD_CELLS = "H6:M6"
Private Const END_ADDRESS_CELL = "N6"
Private Const PLANE_CELL_NUM = 6

Private m_ErrLog As String


Public Property Set targetSheet(ByRef RHS As Object)
    Set m_Sheet = RHS
    Set m_CbChecker = RHS.CB_Checker
End Property

Public Sub CreatePMD()

    Dim pmdCell As Range
    Dim Planes As IImgPlaneManager

    Set pmdCell = m_Sheet.Range(PMDNAME_CELL)
    While Not IsEmpty(pmdCell)
        If Not IsEmpty(pmdCell.offset(0, -5)) Then
            With pmdCell
                Set Planes = TheIDP.CreateBasePMD(.offset(0, -6).Value, .Value, .offset(0, 3).Value, .offset(0, 4).Value, .offset(0, -2).Value, .offset(0, -5).Value, .offset(0, -4).Value, .offset(0, -3).Value)
            End With
        Else
            With pmdCell
                Call Planes.CreateSubPMD(.Value, .offset(0, 1).Value, .offset(0, 2).Value, .offset(0, 3).Value, .offset(0, 4).Value)
            End With
        End If
        Set pmdCell = pmdCell.offset(1)
    Wend

End Sub

Public Sub CheckPmdName()
    'PMD名の重複チェック
    
    Dim table As Collection
    Set table = New Collection
    
    Dim cell As Range
    Set cell = m_Sheet.Range(PMDNAME_CELL)
    
    While Not IsEmpty(cell)
        Call CheckCollision(table, cell)
        
        Set cell = cell.offset(1)
    Wend

End Sub

Private Sub CheckCollision(ByRef pTable As Collection, ByRef pDst As Range)

    On Error GoTo COLLISION
    
    Call pTable.Add(pDst, pDst.Text)
    Exit Sub

COLLISION:
    pDst.Interior.color = RGB(255, 0, 0)
    pTable.Item(pDst.Text).Interior.color = RGB(255, 0, 0)
    
    Call AddErrorLog("""" & pDst.Value & """が衝突しています")
    
End Sub

Public Sub CheckPlaneName()
    'プレーン名の重複チェック
    
    Dim table As Collection
    Set table = New Collection
    
    Dim cell As Range
    Dim planeCell As Range
    
    Set cell = m_Sheet.Range(PMDNAME_CELL)
    Set planeCell = m_Sheet.Range(PLANENAME_CELL)
    
    While Not IsEmpty(cell)
        If Not IsEmpty(planeCell) Then
            Call CheckCollision(table, planeCell)
        End If
        
        Set cell = cell.offset(1)
        Set planeCell = planeCell.offset(1)
    Wend
    
End Sub

Public Sub CheckPmdSize()
    'PMDのサイズチェック
    
    Dim cell As Range
    Dim planeCell As Range
    
    Set cell = m_Sheet.Range(PMDNAME_CELL)
    Set planeCell = m_Sheet.Range(PLANENAME_CELL)
    
    Dim baseWidth As Long
    Dim baseHeight As Long
    Dim subLeft As Long
    Dim subTop As Long
    Dim subRight As Long
    Dim subBottom As Long
    
    While Not IsEmpty(cell)
        If Not IsEmpty(planeCell) Then
            With cell
                If .offset(0, 1).Value <> 1 Or .offset(0, 2).Value <> 1 Then
                    .Interior.color = RGB(255, 0, 0)
                    Call AddErrorLog("""" & .Value & """のサイズが不適当です")
                End If
            End With
            baseWidth = cell.offset(0, 3).Value
            baseHeight = cell.offset(0, 4).Value
        Else
            subLeft = cell.offset(0, 1).Value
            subTop = cell.offset(0, 2).Value
            subRight = subLeft + cell.offset(0, 3).Value - 1
            subBottom = subTop + cell.offset(0, 4).Value - 1
            If (subLeft < 1) Or (subTop < 1) Or (subRight > baseWidth) Or (subBottom > baseHeight) Then
                cell.Interior.color = RGB(255, 0, 0)
                Call AddErrorLog("""" & cell.Value & """のサイズが不適当です")
            End If
        End If
        
        Set cell = cell.offset(1)
        Set planeCell = planeCell.offset(1)
    Wend
    
End Sub

Public Sub DrawLine()
    '罫線を引く

    With m_Sheet
        '一度クリア
'        .Cells.Borders.LineStyle = xlLineStyleNone
        .Rows(VALID_ROWS).Borders.LineStyle = xlLineStyleNone        '見出しは修正しないように変更2010/12/03
        
        '内部を細線にする
        '全体を太線で囲む
        Call DrawLine_All
        
        'BaseとSubの間を中線にする
        Call DrawLine_Middle
        
        '見出しとデータの間を中線にする
        '各グループ間を中線にする
        Call DrawLine_byGroup
    End With

End Sub

Private Sub DrawLine_All()

    With m_Sheet
'        With .Range(.Range(TOP_INDEX_CELLS), .Range(TOP_PMD_CELLS).End(xlDown))
        With .Range(.Range(VALID_CELLS), .Range(VALID_PMD_CELLS).End(xlDown))
            .Borders.LineStyle = xlContinuous
            Call .BorderAround(xlContinuous, xlThick)
        End With
    End With
    
End Sub

Private Sub DrawLine_Middle()
    
    With m_Sheet
        With .Range(.Range(VALID_CELLS), .Range(VALID_PMD_CELLS).End(xlDown))
            .Borders(xlEdgeLeft).Weight = xlMedium
        End With
    End With

End Sub

Private Sub DrawLine_byGroup()

    Dim pmdCell As Range
    Dim planeCell As Range
    Dim dstCell As Range
    
    With m_Sheet
        Set pmdCell = .Range(PMDNAME_CELL)
        Set planeCell = .Range(PLANENAME_CELL)
        
        Set dstCell = .Range(VALID_INDEX_CELLS)
        While Not IsEmpty(pmdCell)
            If Not IsEmpty(planeCell) Then
                dstCell.Borders(xlEdgeBottom).Weight = xlMedium
                Set dstCell = .Range(planeCell, .Cells(planeCell.Row, 13))
            Else
                Set dstCell = Union(dstCell, .Range(planeCell, .Cells(planeCell.Row, 13)))
            End If
            
            Set pmdCell = pmdCell.offset(1)
            Set planeCell = planeCell.offset(1)
        Wend
    End With
    
End Sub

Public Sub PaintCells()
    'BasePMD欄の空白部に色を塗る
    
    With m_Sheet
        Dim pmdCell As Range
        Dim planeCell As Range
        
        Set pmdCell = .Range(PMDNAME_CELL)
        Set planeCell = .Range(PLANENAME_CELL)
        
        While Not IsEmpty(pmdCell)
            If IsEmpty(planeCell) Then
                .Range(planeCell, pmdCell.offset(0, -1)).Interior.color = RGB(128, 128, 128)
            End If
            
            Set pmdCell = pmdCell.offset(1)
            Set planeCell = planeCell.offset(1)
        Wend
    End With
    
End Sub

Public Sub GroupRows()
    'グループ化する
    Call UngroupRows
    
    With m_Sheet
        Dim pmdCell As Range
        Dim planeCell As Range
        Dim dstRows As Range
        
        Set pmdCell = .Range(PMDNAME_CELL)
        Set planeCell = .Range(PLANENAME_CELL)
        
        Set dstRows = Nothing
        
        .Outline.SummaryRow = xlSummaryAbove
        
        While Not IsEmpty(pmdCell)
            If Not IsEmpty(planeCell) Then
                If Not dstRows Is Nothing Then
                    Call dstRows.group
                    Set dstRows = Nothing
                End If
            Else
                If dstRows Is Nothing Then
                    Set dstRows = .Rows(pmdCell.Row)
                Else
                    Set dstRows = Union(dstRows, .Rows(pmdCell.Row))
                End If
            End If
            
            Set pmdCell = pmdCell.offset(1)
            Set planeCell = planeCell.offset(1)
        Wend
        
        If Not dstRows Is Nothing Then
            Call dstRows.group
        End If
    End With
    
End Sub

Private Sub UngroupRows()
    'グループ化の解除
    With m_Sheet
        On Error GoTo UNGROUP_END
        Do
            .Rows.Ungroup
        Loop
    End With

UNGROUP_END:
    '終了
    
End Sub

Public Sub PaintFormula()

    Dim cell As Range
    
    With m_Sheet
        For Each cell In .Range(.Range(VALID_CELLS), .Range(VALID_PMD_CELLS).End(xlDown))
            If Left$(cell.formula, 1) = "=" Then
                cell.Interior.color = RGB(255, 255, 0)
            End If
        Next cell
    End With
    
End Sub

Public Sub ClearFormat()

    Dim pHeight As Long
    Dim pRangeAddress As String
    Dim pTmpWrkSht As Worksheet

    Set pTmpWrkSht = ActiveSheet        '現在のアクティブシートを保持しておきます。
    
    If TypeName(Selection) = "Range" Then
        pRangeAddress = Selection.Address
    End If
        
    With m_Sheet
        
        m_Sheet.Select      '一時対象シート選択

        pHeight = .UsedRange.height             'SpecialCellsの誤作動対策ダミー。行や列を挿入や削除するとエクセルが認識する領域が実際とずれる。
        
        'Plane領域とPMD領域の罫線クリア
        With .Range(.Cells(START_ROW_NUM, START_COLUMN_NUM), .Cells(START_ROW_NUM, START_COLUMN_NUM).SpecialCells(xlCellTypeLastCell))
            .Borders.LineStyle = xlNone
'            .Interior.ColorIndex = xlNone
        End With
        
        '### 塗りつぶしをクリア ###
        'PMD領域は空白部分のみ書式クリア
        .Range(.Cells(START_ROW_NUM, START_COLUMN_NUM), .Cells(START_ROW_NUM, START_COLUMN_NUM).SpecialCells(xlCellTypeLastCell)).SpecialCells(xlCellTypeBlanks).ClearFormats
'        With .Range(.Cells(START_ROW_NUM, START_COLUMN_NUM), .Cells(START_ROW_NUM, START_COLUMN_NUM).SpecialCells(xlCellTypeLastCell)).SpecialCells(xlCellTypeBlanks)
'            .Interior.ColorIndex = xlNone
'        End With
        
        '### Plane領域の書式クリア ###
        .Range(.Cells(START_ROW_NUM, START_COLUMN_NUM), .Cells(START_ROW_NUM, START_COLUMN_NUM).SpecialCells(xlCellTypeLastCell)).Resize(, PLANE_CELL_NUM).ClearFormats
'        With .Range(.Cells(START_ROW_NUM, START_COLUMN_NUM), .Cells(START_ROW_NUM, START_COLUMN_NUM).SpecialCells(xlCellTypeLastCell)).Resize(, PLANE_CELL_NUM)
'            .Interior.ColorIndex = xlNone
'        End With
        
        .Cells.EntireRow.Hidden = False
        
        .Range(pRangeAddress).Select        'SpecialCellsのセルがactiveになるので戻す
        
        Call UngroupRows
    
        If pRangeAddress <> "" Then
            .Range(pRangeAddress).Select
        End If
    
    End With
    
    pTmpWrkSht.Select                   '元のアクティブシートに戻します。
    Set pTmpWrkSht = Nothing
    
End Sub

Private Sub WriteEndAddress()

    Dim pHeight As Long
    With m_Sheet
        .Range(END_ADDRESS_CELL, .Range(END_ADDRESS_CELL).End(xlDown)).Resize(, 2).Value = Empty
        
        pHeight = .Range(.Range(VALID_CELLS), .Range(VALID_PMD_CELLS).End(xlDown)).Rows.Count
        .Range(END_ADDRESS_CELL).Value = "=I6+K6-1"
        .Range(END_ADDRESS_CELL).Copy .Range(END_ADDRESS_CELL).Resize(pHeight, 2)
    End With

End Sub

Public Sub WriteCalcColumn()

    'サイズ計算行の式を書く
    
    With m_Sheet
        Dim pmdCell As Range
        Dim planeCell As Range
        
        Set pmdCell = .Range(PMDNAME_CELL)
        Set planeCell = .Range(PLANENAME_CELL)
        
        While Not IsEmpty(pmdCell)
            If Not IsEmpty(planeCell) Then
                pmdCell.offset(0, -1).FormulaR1C1 = "=2*RC[4]*RC[5]*(2*RC[-4]+4*RC[-3]+4*RC[-2])/2^20"
                pmdCell.offset(0, -1).NumberFormatLocal = "0.00_ "      '書式設定の追加(2010/12/28)
            Else
                pmdCell.offset(0, -1).Value = Empty
            End If
            
            Set pmdCell = pmdCell.offset(1)
            Set planeCell = planeCell.offset(1)
        Wend
    End With

End Sub

Private Sub AddErrorLog(ByVal pMessage As String)
    m_ErrLog = m_ErrLog & pMessage & vbCrLf
End Sub

Public Sub DisplayErrorLog()
    If m_ErrLog = "" Then
        Call MsgBox("OK")
    Else
        Call MsgBox(m_ErrLog)
    End If
End Sub

Private Sub m_CbChecker_Click()
        
    m_ErrLog = ""
    Application.ScreenUpdating = False
    
    Call ClearFormat
    Call CheckPmdName
    Call CheckPlaneName
    Call CheckPmdSize
    Call DrawLine
    Call PaintCells
    Call GroupRows
    Call WriteCalcColumn
    Call PaintFormula
    Call WriteEndAddress

    Application.ScreenUpdating = True
    
    Call DisplayErrorLog
    
End Sub

Private Sub m_Sheet_Change(ByVal target As Range)
    DestroyTheIDP
End Sub

