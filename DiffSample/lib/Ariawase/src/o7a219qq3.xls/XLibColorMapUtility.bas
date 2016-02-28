Attribute VB_Name = "XLibColorMapUtility"
Option Explicit

Private Const SHT_LABEL_COLOR_MAP_NAME As String = "Color Map Name"
Private Const SHT_LABEL_COLOR_MAP As String = "Color Map"
Private Const SHT_LABEL_COLOR As String = "Color"
Private Const SHT_LABEL_COLOR_ARG_START As Integer = 1
Private Const SHT_LABEL_COLOR_ARG_END As Integer = 8
Private Const SHT_LABEL_COMMENT As String = "Comment"

Private Const SHT_DATA_ROWSTART As Integer = 5
Private Const SHT_DATA_COLUMNSTART As Integer = 2
Private Const SHT_DATA_COLUMNEND As Integer = 10

Private Const SHT_DATA_WIDTH As String = "B:K"
Private Const SHT_DATA_START As String = "B5"

'OK
Public Sub ControlShtFormatColorMap()
'内容:
'   ColorMapシートの書式を整えます。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年01月11日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'   シートデータの整合チェック、シートデータのインスタンス生成は、CImgPlaneMapManagerが受け持ちます。

    '##### Sheetの書式を初期化 #####
    Dim tmpHeight As Long
    Dim pRangeAddress As String
    Dim pTmpWrkSht As Worksheet     'ActiveSheet保持用
    Dim pWorkSht As Worksheet       '書式整形用
    
    '#### 行のグループ化 ####
    Dim pGroupStart As Integer
    Dim pGroupEnd As Integer
    Dim pGroupInfo As Collection
    Dim pTmp As Variant
    
    Set pGroupInfo = New Collection
    
    '#### 行のデータ確認カウンタ ####
    Dim intStartRow As Integer
    Dim intDataCnt As Integer
    Dim intGroupRowCnt As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    Application.ScreenUpdating = False
    
    '#####  SheetReaderを利用して、ColorMapInfoシートを読み込む  #####
    Dim clsWrkShtRdr As CWorkSheetReader
    Set clsWrkShtRdr = GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeColorMap)
    
    Dim strSheetName As String
    strSheetName = GetWkShtReaderManagerInstance.GetActiveSheetName(eSheetType.shtTypeColorMap)

    Dim IFileStream As IFileStream
    Dim IParamReader As IParameterReader

    With clsWrkShtRdr
        Set IFileStream = .AsIFileStream
        Set IParamReader = .AsIParameterReader
    End With

    Set clsWrkShtRdr = Nothing

    Set pWorkSht = Worksheets(strSheetName)
    
    Set pTmpWrkSht = ActiveWorkbook.ActiveSheet       '現在のアクティブシートを保持しておきます。

    With pWorkSht
    
        pWorkSht.Select

        If TypeName(Selection) = "Range" Then
            pRangeAddress = Selection.Address
        End If
        
        tmpHeight = .UsedRange.height             'SpecialCells誤動作対策のダミー。行、列を削除したときに誤動作する。
        
        With .Range(SHT_DATA_START, .Range(SHT_DATA_START).Cells.SpecialCells(xlCellTypeLastCell))
            .Borders.LineStyle = xlNone
            .Interior.ColorIndex = xlNone
            .ClearOutline
        End With
        
        If pRangeAddress <> "" Then
            .Range(pRangeAddress).Select
        End If
    
    End With
    
    pTmpWrkSht.Select                   '元のアクティブシートに戻します。
    Set pTmpWrkSht = Nothing
    
    '#####  ColorMapInfoシートの情報を読み出し＆格納する  #####
    pWorkSht.Outline.SummaryRow = xlSummaryAbove
    Do While Not IFileStream.IsEOR
    
        intStartRow = SHT_DATA_ROWSTART + intDataCnt
        
        If IParamReader.ReadAsString(SHT_LABEL_COLOR_MAP_NAME) <> "" Then
            pWorkSht.Rows(CStr(intStartRow)).Columns(SHT_DATA_WIDTH).Borders(xlEdgeTop).Weight = xlMedium      'グループ始まりの罫線を引く
        End If
        
        '#####  ColorMapInfoシートに網掛け処理を実施  #####
        With pWorkSht
            For i = SHT_DATA_COLUMNSTART To SHT_DATA_COLUMNEND Step 1
                If i = 2 Then
                    'Color Map Name
                    If IParamReader.ReadAsString(SHT_LABEL_COLOR_MAP_NAME) <> "" Then
                        .Cells(intStartRow, i).Interior.Pattern = xlSolid
                        If pGroupEnd <> 0 And pGroupEnd >= pGroupStart Then
                            Call pGroupInfo.Add(pGroupStart & ":" & pGroupEnd)
                        End If
                        
                        .Cells(intStartRow, i).Interior.ColorIndex = xlNone
                        pGroupStart = intStartRow + 1
                        pGroupEnd = intStartRow
                    Else
                        .Cells(intStartRow, i).Interior.Pattern = xlGray8
                        .Cells(intStartRow, i).Interior.ColorIndex = 15
                        pGroupEnd = intStartRow
                    End If
                Else
                    'color Map 1-8
                    If IParamReader.ReadAsString(SHT_LABEL_COLOR & i - SHT_DATA_COLUMNSTART & "@" & SHT_LABEL_COLOR_MAP) <> "" Then
                        .Cells(intStartRow, i).Interior.Pattern = xlSolid
                        .Cells(intStartRow, i).Interior.ColorIndex = xlNone
                    Else
                        .Range(.Cells(intStartRow, i), .Cells(intStartRow, SHT_DATA_COLUMNEND)).Interior.Pattern = xlGray8
                        .Range(.Cells(intStartRow, i), .Cells(intStartRow, SHT_DATA_COLUMNEND)).Interior.ColorIndex = 15
                        Exit For
                    End If
                End If
            Next i
        End With
        
        intDataCnt = intDataCnt + 1
        intGroupRowCnt = intGroupRowCnt + 1
        IFileStream.MoveNext
    Loop

    If pGroupEnd <> 0 And pGroupEnd >= pGroupStart Then
        Call pGroupInfo.Add(pGroupStart & ":" & pGroupEnd)
    End If
    
    '##### カラーマップテーブルごとに罫線を整えます #####
    pWorkSht.Rows(CStr(SHT_DATA_ROWSTART & ":" & intStartRow)).Columns(SHT_DATA_WIDTH).BorderAround Weight:=xlThick
    
    '##### カラーマップテーブルごとに行をグループ化 #####
    For Each pTmp In pGroupInfo
        pWorkSht.Rows(CStr(pTmp)).group
    Next pTmp
    
    '#####  終了  #####
    Application.ScreenUpdating = True
    
    Set IFileStream = Nothing
    Set IParamReader = Nothing
    Set pGroupInfo = Nothing
    Set pWorkSht = Nothing
    
End Sub

