Attribute VB_Name = "XLibTestFlowMod"
'�T�v:
'   �e�X�g�t���[����p���C�u�����Q
'
'�ړI:
'   EnableWord�̈ꊇFlase�ݒ���s��
'   �����I�ɂ�EnableWord�̊Ǘ��������������E�E�E
'
'�쐬��:
'   SLSI��J

Option Explicit

Dim mEnableWords() As String

Public Sub DisableAllTest()
'���e:
'   FlowTable����EnableWord���擾���S��False�ɐݒ肷��
'
'�p�����[�^:
'
'���ӎ���:
'

    Dim wordIndex As Long

    If CheckEnableWord = False Then
        Exit Sub
    End If

    For wordIndex = 0 To UBound(mEnableWords)
        TheExec.Flow.EnableWord(mEnableWords(wordIndex)) = False
    Next wordIndex

End Sub

Public Function CheckEnableWord( _
) As Boolean
'���e:
'   EnableWord�z��Ƀf�[�^���i�[����Ă��邩�ǂ������m�F
'   �Ȃ���Ώ��������s��
'
'�p�����[�^:
'
'�߂�l�F
'   �������������������ǂ���
'
'���ӎ���:
'

    Const DATA_SHEET_NAME = "Flow Table"
    Dim workSheetObject As Worksheet

    If preInit = True Then

        Set workSheetObject = getWorkSheet(DATA_SHEET_NAME)

        If workSheetObject Is Nothing Then
            Exit Function
        End If

        If getEnableWord(workSheetObject) = False Then
            Exit Function
        End If

    End If

    CheckEnableWord = True

End Function

Private Function getEnableWord( _
    ByVal targetWorkSheet As Object _
) As Boolean

    Const FunctionName = "getEnableWord"

    Const ENABLE_COLUMN = 3
    Const OPCODE_COLUMN = 7
    Const ENABLE_LABEL = "Enable"
    Const OPCODE_LABEL = "Opcode"

    Dim testEnable As Range
    Dim testOpcode As Range
    Dim rowIndex As Long

    Dim tempWord As String
    Dim enableWordCount As Long
    Dim enableWords() As String
    Dim wordIndex As Long

    On Error GoTo errMsg

    With targetWorkSheet

        Set testEnable = .Columns(ENABLE_COLUMN).Find(ENABLE_LABEL)
        Set testOpcode = .Columns(OPCODE_COLUMN).Find(OPCODE_LABEL)

        rowIndex = testOpcode.Row + 1

        Do While .Cells(rowIndex, testOpcode.Column) <> ""
            If .Cells(rowIndex, testOpcode.Column) = "Test" Then
                If .Cells(rowIndex, testEnable.Column) <> "" Then
                    If tempWord <> .Cells(rowIndex, testEnable.Column) Then
                        tempWord = .Cells(rowIndex, testEnable.Column)
                        enableWordCount = enableWordCount + 1
                    End If
                End If
            End If
            rowIndex = rowIndex + 1
        Loop

        ReDim mEnableWords(enableWordCount - 1) As String

        rowIndex = testOpcode.Row + 1

        Do While .Cells(rowIndex, testOpcode.Column) <> ""
            If .Cells(rowIndex, testOpcode.Column) = "Test" Then
                If .Cells(rowIndex, testEnable.Column) <> "" Then
                    If tempWord <> .Cells(rowIndex, testEnable.Column) Then
                        mEnableWords(wordIndex) = .Cells(rowIndex, testEnable.Column)
                        tempWord = .Cells(rowIndex, testEnable.Column)
                        wordIndex = wordIndex + 1
                    End If
                End If
            End If
            rowIndex = rowIndex + 1
        Loop

    End With

    getEnableWord = True

    Exit Function

errMsg:

    Call DebugMsg(FunctionName & " Is Failed !")
    '���{��G���[���b�Z�[�W�o��
'    Call DebugMsg(FunctionName & " �Ɏ��s���܂���")

    getEnableWord = False

End Function

Private Function getWorkSheet( _
    ByRef SheetName As String _
) As Worksheet

    On Error GoTo errMsg

    Set getWorkSheet = Worksheets(SheetName)

    Exit Function

errMsg:

    Call DebugMsg("Not " & SheetName & " Sheet Exist !")
    '���{��G���[���b�Z�[�W�o��
'    Call DebugMsg(sheetName & " �V�[�g������܂���")

    Set getWorkSheet = Nothing

End Function

Private Function preInit() As Boolean

    Dim i As Long

    On Error GoTo UNTIL_EMPTY

    i = UBound(mEnableWords)

    preInit = False

    Exit Function

UNTIL_EMPTY:

    preInit = True

End Function



