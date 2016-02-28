Attribute VB_Name = "XLibImageKernelUtility"
Option Explicit

Private Const SHT_LABEL_KERNEL_NAME As String = "Kernel Name"
Private Const SHT_LABEL_VAL As String = "Val"
Private Const SHT_LABEL_COMMENT As String = "Comment"

Private Const SHT_DATA_ROWSTART As Integer = 4
Private Const SHT_DATA_COLUMNSTART As Integer = 2
Private Const SHT_DATA_COLUMNEND As Integer = 73
'�J�[�l���̈���
Private Const SHT_DATA_KERNELPARAMSTART As Integer = 2
Private Const SHT_DATA_KERNELPARAMEND As Integer = 8
'�J�[�l���̃f�[�^
Private Const SHT_DATA_KERNELDATASTART As Integer = 9
Private Const SHT_DATA_KERNELDATAEND As Integer = 72

Private Const SHT_DATA_WIDTH As String = "B:BU"
Private Const SHT_DATA_START As String = "B4"

Public Sub CreateKernelManagerIfNothing()
'���e:
'   �J�[�l���}�l�[�W���[�̏��̗L�������āA������΃J�[�l���V�[�g��ǂ݂ɍs���܂��B
'   �V�[�g��������Ή������܂���B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��10��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'

    On Error GoTo Err_Handler
    If TheIDP.KernelManager.Count = 0 Then
        Call TheIDP.KernelManager.Init
'        Call ControlShtFormatKernel
    End If
    
    Exit Sub

Err_Handler:
    If TheIDP.KernelManager.IsErrIGXL = True Then
        'EeeJOB�`�F�b�N��OK�ŁAIG-XL�ŃG���[�Ȃ̂�TheIDP.RemoveResources
        Call TheIDP.RemoveResources
    End If
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'OK
Public Sub ControlShtFormatKernel()
'���e:
'   ColorMap�V�[�g�̏����𐮂��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N01��11��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'

    '##### Sheet�̏����������� #####
    Dim tmpHeight As Long
    Dim pRangeAddress As String
    Dim pTmpWrkSht As Worksheet     'ActiveSheet�ێ��p
    Dim pWorkSht As Worksheet       '�������`�p
    
    '#### �s�̃O���[�v�� ####
    Dim pGroupStart As Integer
    Dim pGroupEnd As Integer
    Dim pGroupInfo As Collection
    Dim pTmp As Variant
    
    Set pGroupInfo = New Collection
    
    '#### �s�̃f�[�^�m�F�J�E���^ ####
    Dim intStartRow As Integer
    Dim intDataCnt As Integer
    Dim intGroupRowCnt As Integer
    
    '#### Kernel�̃f�[�^ ####
    Dim pWidth As Integer
    Dim pHeight As Integer
    Dim pWidthAddr As String
    Dim pHeightAddr As String
    Dim pShiftR As Integer
    Dim pKernelType As IdpKernelType
    Dim pFirstChk As Boolean                '�J�[�l����`�̃p�����[�^�`�F�b�N�B���ꂪNG�Ȃ�f�[�^��͖����B
    Dim pNameForChk As Collection
    
    '#### Kernel Anchor�̃f�[�^�i�[�p ####
    Dim pAnchorCnt As Long
    Dim pAnchorAddrCollect As Collection
    Dim pAnchorValCollect As Collection
    Set pAnchorAddrCollect = New Collection
    Set pAnchorValCollect = New Collection
    
    Dim i As Integer
    Dim j As Integer
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    '#####  SheetReader�𗘗p���āAKernel�V�[�g��ǂݍ���  #####
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
    
    Set pTmpWrkSht = ActiveWorkbook.ActiveSheet       '���݂̃A�N�e�B�u�V�[�g��ێ����Ă����܂��B
    pWorkSht.Select

    With pWorkSht

        If TypeName(Selection) = "Range" Then
            pRangeAddress = Selection.Address
        End If
        
        tmpHeight = .UsedRange.height             'SpecialCells�듮��΍�̃_�~�[�B�s�A����폜�����Ƃ��Ɍ듮�삷��B
        
        With .Range(SHT_DATA_START, .Range(SHT_DATA_START).Cells.SpecialCells(xlCellTypeLastCell))
            .Borders.LineStyle = xlNone
            .Interior.ColorIndex = xlNone
            .ClearOutline
        End With
        
        If pRangeAddress <> "" Then
            .Range(pRangeAddress).Select
        End If
    
    End With
        
    '#####  Kernel�V�[�g�̐��`  #####
    pWorkSht.Outline.SummaryRow = xlSummaryAbove
    With pWorkSht
        Set pNameForChk = New Collection
        Do While Not IFileStream.IsEOR
    
            intStartRow = SHT_DATA_ROWSTART + intDataCnt
    
            '##### Kernel �p�����[�^�� #####
            If IParamReader.ReadAsString(SHT_LABEL_KERNEL_NAME) <> "" Then
                pWorkSht.Rows(CStr(intStartRow)).Columns(SHT_DATA_WIDTH).Borders(xlEdgeTop).Weight = xlMedium      '�O���[�v�n�܂�̌r��������
                i = 2
                '##### �O�s�܂ł̒�`�I�[���� #####
                'Kernel�����������Ă��邩�`�F�b�N
                If pHeight <> 0 And pHeight > intGroupRowCnt Then
                    .Cells(intStartRow - intGroupRowCnt, i + 2).Interior.ColorIndex = 3
                    .Range(pHeightAddr).Interior.ColorIndex = 3
                End If
                '�O��̒�`�̃O���[�v����add
                If pGroupEnd <> 0 And pGroupEnd >= pGroupStart Then
                    Call pGroupInfo.Add(pGroupStart & ":" & pGroupEnd)
                End If

                '##### Kernel��`�J�n #####
                .Cells(intStartRow, i).Interior.Pattern = xlSolid
                .Cells(intStartRow, i).Interior.ColorIndex = xlNone
                pGroupStart = intStartRow + 1
                pGroupEnd = intStartRow
                intGroupRowCnt = 0
                pFirstChk = True
                
                '�V�[�g���ł̃J�[�l�����d���`�F�b�N
                If IsKey(CStr(.Cells(intStartRow, i)), pNameForChk) = True Then
                    .Cells(intStartRow, i).Interior.ColorIndex = 3
                    pFirstChk = False
                Else
                    Call pNameForChk.Add(CStr(.Cells(intStartRow, i)), CStr(.Cells(intStartRow, i)))
                End If
                
                '���`�F�b�N
                pWidth = .Cells(intStartRow, i + 1).Value
                pWidthAddr = .Cells(intStartRow, i + 1).Address
                If ChkSize(pWidth) = False Then
                    .Cells(intStartRow, i + 1).Interior.ColorIndex = 3
                    pFirstChk = False
                End If
                '�����`�F�b�N
                pHeight = .Cells(intStartRow, i + 2).Value
                pHeightAddr = .Cells(intStartRow, i + 2).Address
                If ChkSize(pHeight) = False Then
                    .Cells(intStartRow, i + 2).Interior.ColorIndex = 3
                    pFirstChk = False
                End If
                'Bit�V�t�g�`�F�b�N
                pShiftR = CInt(.Cells(intStartRow, i + 5).Value)
                If ChkShiftR(pShiftR) = False Then
                    .Cells(intStartRow, i + 5).Interior.ColorIndex = 3
                    pFirstChk = False
                End If
                '�J�[�l���^�C�v�`�F�b�N
                pKernelType = CIdpKernel(.Cells(intStartRow, i + 6).Value)
                If pKernelType = -1 Then
                    .Cells(intStartRow, i + 6).Interior.ColorIndex = 3
                    pFirstChk = False
                End If
                
                '�L���p�����[�^���G���[�̏ꍇ�̓f�[�^�͖���
                If pFirstChk = False Then
                    GoTo NEXT_DOLOOP        'VBA�ł�C��Continue�݂����Ȃ�������̂�
                End If
                
                'X Anchor
'                    .Cells(intStartRow, i).Value = ((pWidth + 1) \ 2)      'Change�C�x���g�ŃZ���̈ʒu���ς���Ă��܂��̂Ń��[�v���ł͎g�p�֎~
                pAnchorCnt = pAnchorCnt + 1
                Call pAnchorAddrCollect.Add(.Cells(intStartRow, i + 3).Address, CStr(pAnchorCnt))
                Call pAnchorValCollect.Add(((pWidth + 1) \ 2), CStr(pAnchorCnt))
                .Cells(intStartRow, i + 3).Interior.Pattern = xlSolid
                .Cells(intStartRow, i + 3).Interior.ColorIndex = 15
                
                'Y Anchor
'                    .Cells(intStartRow, i).Value = ((pHeight + 1) \ 2)      'Change�C�x���g�ŃZ���̈ʒu���ς���Ă��܂��̂Ń��[�v���ł͎g�p�֎~
                pAnchorCnt = pAnchorCnt + 1
                Call pAnchorAddrCollect.Add(.Cells(intStartRow, i + 4).Address, CStr(pAnchorCnt))
                Call pAnchorValCollect.Add(((pHeight + 1) \ 2), CStr(pAnchorCnt))
                .Cells(intStartRow, i + 4).Interior.Pattern = xlSolid
                .Cells(intStartRow, i + 4).Interior.ColorIndex = 15
            
            Else
                
                '�L���p�����[�^���G���[�̏ꍇ�̓f�[�^�͖���
                If pFirstChk = False Then
                    GoTo NEXT_DOLOOP        'VBA�ł�C��Continue�݂����Ȃ�������̂�
                End If
                
                If intGroupRowCnt >= pHeight Then
                    '�ݒ���c�̃f�[�^����������B
                    .Range(.Cells(intStartRow, SHT_DATA_KERNELPARAMSTART), .Cells(intStartRow, SHT_DATA_KERNELDATAEND)).Interior.ColorIndex = 3
                    .Range(pHeightAddr).Interior.ColorIndex = 3
                End If
                'Kernel�p�����[�^�̈��h��Ԃ�
                .Range(.Cells(intStartRow, SHT_DATA_KERNELPARAMSTART), .Cells(intStartRow, SHT_DATA_KERNELPARAMEND)).Interior.Pattern = xlGray8
                .Range(.Cells(intStartRow, SHT_DATA_KERNELPARAMSTART), .Cells(intStartRow, SHT_DATA_KERNELPARAMEND)).Interior.ColorIndex = 15
                pGroupEnd = intStartRow
            End If
            
            '##### Kernel �f�[�^�� #####
            For i = SHT_DATA_KERNELDATASTART To SHT_DATA_KERNELDATAEND
                If i < SHT_DATA_KERNELDATASTART + pWidth Then
                    'Kernel�f�[�^�͈̔͊O��h��Ԃ�
                    If .Cells(intStartRow, i).Value = "" Then
                        .Cells(intStartRow, i).Interior.ColorIndex = 3
                        .Range(pWidthAddr).Interior.ColorIndex = 3
                    End If
                ElseIf i >= SHT_DATA_KERNELDATASTART + pWidth Then
                    'Kernel�f�[�^�͈̔͊O��h��Ԃ�
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
    
        'Anchor�̒l���Z���ɓ���
        If pAnchorCnt > 0 Then
            For i = 1 To pAnchorCnt
                .Range(pAnchorAddrCollect.Item(CStr(i))).Value = pAnchorValCollect.Item(CStr(i))
            Next i
        End If
    
        '##### �O�s�܂ł̒�`�I�[���� #####
        '�O�̒�`��Kernel�����������Ă��邩�`�F�b�N���āA����Ȃ���΍����w���ԓh��Ԃ�
        If pHeight <> 0 And pHeight > intGroupRowCnt Then
            .Cells(intStartRow - intGroupRowCnt + 1, i + 2).Interior.ColorIndex = 3
            .Range(pHeightAddr).Interior.ColorIndex = 3
        End If
        '�O�̒�`�̃O���[�v����add
        If pGroupEnd <> 0 And pGroupEnd >= pGroupStart Then
            Call pGroupInfo.Add(pGroupStart & ":" & pGroupEnd)
        End If
    
    
        '##### �O�g�r�� #####
        If intStartRow > 0 Then
            .Rows(CStr(SHT_DATA_ROWSTART & ":" & intStartRow)).Columns(SHT_DATA_WIDTH).BorderAround Weight:=xlThick
        End If
        
        '##### �J�[�l����`���Ƃɍs���O���[�v�� #####
        For Each pTmp In pGroupInfo
            .Rows(CStr(pTmp)).group
        Next pTmp
    
    End With
    
    pTmpWrkSht.Select                   '���̃A�N�e�B�u�V�[�g�ɖ߂��܂��B
    Set pTmpWrkSht = Nothing

    '#####  �I��  #####
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
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��3��
'�p�����[�^:
'   [pShiftRbit]    In/Out  1):
'�߂�l:
'   Integer
'
'���ӎ���:
'
'
    If pShiftRbit >= 0 And pShiftRbit <= 16 Then
        ChkShiftR = True
    Else
        ChkShiftR = False
    End If
    
End Function

Private Function CIdpKernel(ByVal pKernelType As String) As IdpKernelType
'���e:
'   ��������idpKernelType�ɕϊ��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��3��
'�p�����[�^:
'   [pKernelType]   In/Out  1):
'�߂�l:
'   IdpKernelType
'
'���ӎ���:
'   �Y�����Ȃ��ꍇ��-1

    pKernelType = UCase(pKernelType)    '�召��������
    
    If pKernelType = "INTEGER" Then
        CIdpKernel = idpKernelInteger
    ElseIf pKernelType = "FLOAT" Then
        CIdpKernel = idpKernelFloat
    Else
        CIdpKernel = -1
    End If

End Function

Private Function ChkSize(pSize As Integer) As Boolean
'���e:
'   �傫�����P�`�Q�T���`�F�b�N���Ė��Ȃ���΂��̂܂ܕԂ��B�G���[�Ȃ�-1��Ԃ��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��4��
'�p�����[�^:
'   [pSize] In/Out  1):
'�߂�l:
'   Integer
'
'���ӎ���:
'
'
    If pSize >= 1 And pSize <= 25 Then
        ChkSize = True
    Else
        ChkSize = False
    End If

End Function

Private Function IsKey(ByVal pKey As String, ByRef pObj As Collection) As Boolean
'���e:
'   �Y��Collection�I�u�W�F�N�g�ɃL�[�����݂��邩���ׂ�
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��22��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Boolean :True���ł�add�ς݁B���݂��܂��B       False�܂�����
'
'���ӎ���:
'
    On Error GoTo ALREADY_REG
    Call pObj.Item(pKey)
    IsKey = True
    Exit Function

ALREADY_REG:
    IsKey = False

End Function



