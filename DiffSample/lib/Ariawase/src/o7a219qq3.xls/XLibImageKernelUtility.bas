Attribute VB_Name = "XLibImageKernelUtility"
Option Explicit

Private Const SHT_LABEL_KERNEL_NAME As String = "Kernel Name"
Private Const SHT_LABEL_VAL As String = "Val"
Private Const SHT_LABEL_COMMENT As String = "Comment"

Private Const SHT_DATA_ROWSTART As Integer = 4
Private Const SHT_DATA_COLUMNSTART As Integer = 2
Private Const SHT_DATA_COLUMNEND As Integer = 73
'カーネルの引数
Private Const SHT_DATA_KERNELPARAMSTART As Integer = 2
Private Const SHT_DATA_KERNELPARAMEND As Integer = 8
'カーネルのデータ
Private Const SHT_DATA_KERNELDATASTART As Integer = 9
Private Const SHT_DATA_KERNELDATAEND As Integer = 72

Private Const SHT_DATA_WIDTH As String = "B:BU"
Private Const SHT_DATA_START As String = "B4"

Public Sub CreateKernelManagerIfNothing()
'内容:
'   カーネルマネージャーの情報の有無を見て、無ければカーネルシートを読みに行きます。
'   シートが無ければ何もしません。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月10日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'

    On Error GoTo Err_Handler
    If TheIDP.KernelManager.Count = 0 Then
        Call TheIDP.KernelManager.Init
'        Call ControlShtFormatKernel
    End If
    
    Exit Sub

Err_Handler:
    If TheIDP.KernelManager.IsErrIGXL = True Then
        'EeeJOBチェックはOKで、IG-XLでエラーなのでTheIDP.RemoveResources
        Call TheIDP.RemoveResources
    End If
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'OK
Public Sub ControlShtFormatKernel()
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
'

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
    
    '#### Kernelのデータ ####
    Dim pWidth As Integer
    Dim pHeight As Integer
    Dim pWidthAddr As String
    Dim pHeightAddr As String
    Dim pShiftR As Integer
    Dim pKernelType As IdpKernelType
    Dim pFirstChk As Boolean                'カーネル定義のパラメータチェック。これがNGならデータ域は無視。
    Dim pNameForChk As Collection
    
    '#### Kernel Anchorのデータ格納用 ####
    Dim pAnchorCnt As Long
    Dim pAnchorAddrCollect As Collection
    Dim pAnchorValCollect As Collection
    Set pAnchorAddrCollect = New Collection
    Set pAnchorValCollect = New Collection
    
    Dim i As Integer
    Dim j As Integer
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    '#####  SheetReaderを利用して、Kernelシートを読み込む  #####
    Dim clsWrkShtRdr As CWorkSheetReader
    Set clsWrkShtRdr = GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeKernel)
    
    Dim strSheetName As String
    strSheetName = GetWkShtReaderManagerInstance.GetActiveSheetName(eSheetType.shtTypeKernel)

    Dim IFileStream As IFileStream
    Dim IParamReader As IParameterReader

    With clsWrkShtRdr
        Set IFileStream = .AsIFileStream
        Set IParamReader = .AsIParameterReader
    End With

    Set clsWrkShtRdr = Nothing

    Set pWorkSht = Worksheets(strSheetName)
    
    Set pTmpWrkSht = ActiveWorkbook.ActiveSheet       '現在のアクティブシートを保持しておきます。
    pWorkSht.Select

    With pWorkSht

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
        
    '#####  Kernelシートの整形  #####
    pWorkSht.Outline.SummaryRow = xlSummaryAbove
    With pWorkSht
        Set pNameForChk = New Collection
        Do While Not IFileStream.IsEOR
    
            intStartRow = SHT_DATA_ROWSTART + intDataCnt
    
            '##### Kernel パラメータ部 #####
            If IParamReader.ReadAsString(SHT_LABEL_KERNEL_NAME) <> "" Then
                pWorkSht.Rows(CStr(intStartRow)).Columns(SHT_DATA_WIDTH).Borders(xlEdgeTop).Weight = xlMedium      'グループ始まりの罫線を引く
                i = 2
                '##### 前行までの定義終端処理 #####
                'Kernel高さが合っているかチェック
                If pHeight <> 0 And pHeight > intGroupRowCnt Then
                    .Cells(intStartRow - intGroupRowCnt, i + 2).Interior.ColorIndex = 3
                    .Range(pHeightAddr).Interior.ColorIndex = 3
                End If
                '前回の定義のグループ情報をadd
                If pGroupEnd <> 0 And pGroupEnd >= pGroupStart Then
                    Call pGroupInfo.Add(pGroupStart & ":" & pGroupEnd)
                End If

                '##### Kernel定義開始 #####
                .Cells(intStartRow, i).Interior.Pattern = xlSolid
                .Cells(intStartRow, i).Interior.ColorIndex = xlNone
                pGroupStart = intStartRow + 1
                pGroupEnd = intStartRow
                intGroupRowCnt = 0
                pFirstChk = True
                
                'シート内でのカーネル名重複チェック
                If IsKey(CStr(.Cells(intStartRow, i)), pNameForChk) = True Then
                    .Cells(intStartRow, i).Interior.ColorIndex = 3
                    pFirstChk = False
                Else
                    Call pNameForChk.Add(CStr(.Cells(intStartRow, i)), CStr(.Cells(intStartRow, i)))
                End If
                
                '幅チェック
                pWidth = .Cells(intStartRow, i + 1).Value
                pWidthAddr = .Cells(intStartRow, i + 1).Address
                If ChkSize(pWidth) = False Then
                    .Cells(intStartRow, i + 1).Interior.ColorIndex = 3
                    pFirstChk = False
                End If
                '高さチェック
                pHeight = .Cells(intStartRow, i + 2).Value
                pHeightAddr = .Cells(intStartRow, i + 2).Address
                If ChkSize(pHeight) = False Then
                    .Cells(intStartRow, i + 2).Interior.ColorIndex = 3
                    pFirstChk = False
                End If
                'Bitシフトチェック
                pShiftR = CInt(.Cells(intStartRow, i + 5).Value)
                If ChkShiftR(pShiftR) = False Then
                    .Cells(intStartRow, i + 5).Interior.ColorIndex = 3
                    pFirstChk = False
                End If
                'カーネルタイプチェック
                pKernelType = CIdpKernel(.Cells(intStartRow, i + 6).Value)
                If pKernelType = -1 Then
                    .Cells(intStartRow, i + 6).Interior.ColorIndex = 3
                    pFirstChk = False
                End If
                
                '記入パラメータがエラーの場合はデータは無視
                If pFirstChk = False Then
                    GoTo NEXT_DOLOOP        'VBAではCのContinueみたいなやつが無いので
                End If
                
                'X Anchor
'                    .Cells(intStartRow, i).Value = ((pWidth + 1) \ 2)      'Changeイベントでセルの位置が変わってしまうのでループ内では使用禁止
                pAnchorCnt = pAnchorCnt + 1
                Call pAnchorAddrCollect.Add(.Cells(intStartRow, i + 3).Address, CStr(pAnchorCnt))
                Call pAnchorValCollect.Add(((pWidth + 1) \ 2), CStr(pAnchorCnt))
                .Cells(intStartRow, i + 3).Interior.Pattern = xlSolid
                .Cells(intStartRow, i + 3).Interior.ColorIndex = 15
                
                'Y Anchor
'                    .Cells(intStartRow, i).Value = ((pHeight + 1) \ 2)      'Changeイベントでセルの位置が変わってしまうのでループ内では使用禁止
                pAnchorCnt = pAnchorCnt + 1
                Call pAnchorAddrCollect.Add(.Cells(intStartRow, i + 4).Address, CStr(pAnchorCnt))
                Call pAnchorValCollect.Add(((pHeight + 1) \ 2), CStr(pAnchorCnt))
                .Cells(intStartRow, i + 4).Interior.Pattern = xlSolid
                .Cells(intStartRow, i + 4).Interior.ColorIndex = 15
            
            Else
                
                '記入パラメータがエラーの場合はデータは無視
                If pFirstChk = False Then
                    GoTo NEXT_DOLOOP        'VBAではCのContinueみたいなやつが無いので
                End If
                
                If intGroupRowCnt >= pHeight Then
                    '設定より縦のデータが多すぎる。
                    .Range(.Cells(intStartRow, SHT_DATA_KERNELPARAMSTART), .Cells(intStartRow, SHT_DATA_KERNELDATAEND)).Interior.ColorIndex = 3
                    .Range(pHeightAddr).Interior.ColorIndex = 3
                End If
                'Kernelパラメータ領域を塗りつぶし
                .Range(.Cells(intStartRow, SHT_DATA_KERNELPARAMSTART), .Cells(intStartRow, SHT_DATA_KERNELPARAMEND)).Interior.Pattern = xlGray8
                .Range(.Cells(intStartRow, SHT_DATA_KERNELPARAMSTART), .Cells(intStartRow, SHT_DATA_KERNELPARAMEND)).Interior.ColorIndex = 15
                pGroupEnd = intStartRow
            End If
            
            '##### Kernel データ部 #####
            For i = SHT_DATA_KERNELDATASTART To SHT_DATA_KERNELDATAEND
                If i < SHT_DATA_KERNELDATASTART + pWidth Then
                    'Kernelデータの範囲外を塗りつぶし
                    If .Cells(intStartRow, i).Value = "" Then
                        .Cells(intStartRow, i).Interior.ColorIndex = 3
                        .Range(pWidthAddr).Interior.ColorIndex = 3
                    End If
                ElseIf i >= SHT_DATA_KERNELDATASTART + pWidth Then
                    'Kernelデータの範囲外を塗りつぶし
                    If .Cells(intStartRow, i).Value = "" Then
                        .Range(.Cells(intStartRow, i), .Cells(intStartRow, SHT_DATA_KERNELDATAEND)).Interior.Pattern = xlGray8
                        .Range(.Cells(intStartRow, i), .Cells(intStartRow, SHT_DATA_KERNELDATAEND)).Interior.ColorIndex = 15
                    Else
                        .Cells(intStartRow, i).Interior.ColorIndex = 3
                        .Range(pWidthAddr).Interior.ColorIndex = 3
                    End If
                End If
            Next i
            
NEXT_DOLOOP:
            intDataCnt = intDataCnt + 1
            intGroupRowCnt = intGroupRowCnt + 1
            IFileStream.MoveNext
        Loop
    
        'Anchorの値をセルに入力
        If pAnchorCnt > 0 Then
            For i = 1 To pAnchorCnt
                .Range(pAnchorAddrCollect.Item(CStr(i))).Value = pAnchorValCollect.Item(CStr(i))
            Next i
        End If
    
        '##### 前行までの定義終端処理 #####
        '前の定義のKernel高さが合っているかチェックして、足りなければ高さ指定を赤塗りつぶし
        If pHeight <> 0 And pHeight > intGroupRowCnt Then
            .Cells(intStartRow - intGroupRowCnt + 1, i + 2).Interior.ColorIndex = 3
            .Range(pHeightAddr).Interior.ColorIndex = 3
        End If
        '前の定義のグループ情報をadd
        If pGroupEnd <> 0 And pGroupEnd >= pGroupStart Then
            Call pGroupInfo.Add(pGroupStart & ":" & pGroupEnd)
        End If
    
    
        '##### 外枠罫線 #####
        If intStartRow > 0 Then
            .Rows(CStr(SHT_DATA_ROWSTART & ":" & intStartRow)).Columns(SHT_DATA_WIDTH).BorderAround Weight:=xlThick
        End If
        
        '##### カーネル定義ごとに行をグループ化 #####
        For Each pTmp In pGroupInfo
            .Rows(CStr(pTmp)).group
        Next pTmp
    
    End With
    
    pTmpWrkSht.Select                   '元のアクティブシートに戻します。
    Set pTmpWrkSht = Nothing

    '#####  終了  #####
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Set pAnchorAddrCollect = Nothing
    Set pAnchorValCollect = Nothing
    Set IFileStream = Nothing
    Set IParamReader = Nothing
    Set pGroupInfo = Nothing
    Set pWorkSht = Nothing
    
End Sub

Private Function ChkShiftR(ByVal pShiftRbit As Integer) As Boolean
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月3日
'パラメータ:
'   [pShiftRbit]    In/Out  1):
'戻り値:
'   Integer
'
'注意事項:
'
'
    If pShiftRbit >= 0 And pShiftRbit <= 16 Then
        ChkShiftR = True
    Else
        ChkShiftR = False
    End If
    
End Function

Private Function CIdpKernel(ByVal pKernelType As String) As IdpKernelType
'内容:
'   文字情報をidpKernelTypeに変換。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月3日
'パラメータ:
'   [pKernelType]   In/Out  1):
'戻り値:
'   IdpKernelType
'
'注意事項:
'   該当しない場合は-1

    pKernelType = UCase(pKernelType)    '大小文字無視
    
    If pKernelType = "INTEGER" Then
        CIdpKernel = idpKernelInteger
    ElseIf pKernelType = "FLOAT" Then
        CIdpKernel = idpKernelFloat
    Else
        CIdpKernel = -1
    End If

End Function

Private Function ChkSize(pSize As Integer) As Boolean
'内容:
'   大きさが１～２５かチェックして問題なければそのまま返す。エラーなら-1を返す。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月4日
'パラメータ:
'   [pSize] In/Out  1):
'戻り値:
'   Integer
'
'注意事項:
'
'
    If pSize >= 1 And pSize <= 25 Then
        ChkSize = True
    Else
        ChkSize = False
    End If

End Function

Private Function IsKey(ByVal pKey As String, ByRef pObj As Collection) As Boolean
'内容:
'   該当Collectionオブジェクトにキーが存在するか調べる
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月22日
'パラメータ:
'   なし
'戻り値:
'   Boolean :Trueすでにadd済み。存在します。       Falseまだ無し
'
'注意事項:
'
    On Error GoTo ALREADY_REG
    Call pObj.Item(pKey)
    IsKey = True
    Exit Function

ALREADY_REG:
    IsKey = False

End Function



