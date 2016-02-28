Attribute VB_Name = "XLibTestFlowMod"
'概要:
'   テストフロー制御用ライブラリ群
'
'目的:
'   EnableWordの一括Flase設定を行う
'   将来的にはEnableWordの管理をさせたいが・・・
'
'作成者:
'   SLSI大谷

Option Explicit

Dim mEnableWords() As String

Public Sub DisableAllTest()
'内容:
'   FlowTableからEnableWordを取得し全てFalseに設定する
'
'パラメータ:
'
'注意事項:
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
'内容:
'   EnableWord配列にデータが格納されているかどうかを確認
'   なければ初期化を行う
'
'パラメータ:
'
'戻り値：
'   初期化が成功したかどうか
'
'注意事項:
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
    '日本語エラーメッセージ出力
'    Call DebugMsg(FunctionName & " に失敗しました")

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
    '日本語エラーメッセージ出力
'    Call DebugMsg(sheetName & " シートがありません")

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



