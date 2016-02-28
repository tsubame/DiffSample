Attribute VB_Name = "XLibStrings"
'概要:
'   文字列に関するチェックや変換を行うプロシージャ群
'
'目的:
'   パラメータクラス用の文字列チェック及び変換プロシージャ群
'   グローバルでも使用出来るよう共通化した
'
'作成者:
'   0145206097
'
Option Explicit

Public Function IsOneByte(ByVal strData As String) As Boolean
'内容:
'   文字列内全ての文字の全角/半角チェックを行う
'
'[strData]     IN String型:     チェック対象となる文字列
'
'戻り値：
'   全て半角の場合はTRUE
'   一つでも全角がある場合はFALSEを返す
'
'備考:
'   なし
'
    Dim strIndex As Long
    Dim maxIndex As Long
    Dim CHAR As String
    maxIndex = Len(strData)
    For strIndex = 1 To maxIndex
        CHAR = Mid$(strData, strIndex, 1)
        If Not (Len(CHAR) = LenB(StrConv(CHAR, vbFromUnicode))) Then
            Exit Function
        End If
    Next strIndex
    IsOneByte = True
End Function

Public Function IsNumber(ByVal CHAR As String) As Boolean
'内容:
'   文字の数字チェックを行う
'
'[char]     IN String型:     チェック対象となる文字
'
'戻り値：
'   数字である場合はTRUE
'   それ以外はFALSEを返す
'
'備考:
'   なし
'
    IsNumber = (CHAR >= "0") And (CHAR <= "9")
End Function

Public Function IsSymbol(ByVal CHAR As String) As Boolean
'内容:
'   文字の記号チェックを行う
'
'[char]     IN String型:     チェック対象となる文字
'
'戻り値：
'   マイナス/カンマ/アンダースコア/パーセントの場合はTRUE
'   それ以外はFALSEを返す
'
'備考:
'   なし
'
    IsSymbol = (CHAR = "-") Or (CHAR = ".") Or (CHAR = "_") Or (CHAR = "%")
End Function

Public Function IsAlphabet(ByVal CHAR As String) As Boolean
'内容:
'   文字のアルファベットチェックを行う
'
'[char]     IN String型:     チェック対象となる文字
'
'戻り値：
'   アルファベットの場合はTRUE
'   それ以外はFALSEを返す
'
'備考:
'   大文字小文字は問わない
'
    IsAlphabet = ((CHAR >= "a") And (CHAR <= "z")) Or ((CHAR >= "A") And (CHAR <= "Z"))
End Function

Public Function IsSubUnit(ByVal SubUnit As String) As Boolean
'内容:
'   文字の補助単位チェックを行う
'
'[subUnit]     IN String型:     チェック対象となる文字
'
'戻り値：
'   該当する補助単位であればTRUE
'   それ以外はFALSEを返す
'
'備考:
'   現在補助単位として[p/n/u/m/%/k/M/G]を準備している
'
    IsSubUnit = (SubUnit = "p") Or (SubUnit = "n") Or (SubUnit = "u") Or (SubUnit = "m") Or (SubUnit = "%") Or (SubUnit = "k") Or (SubUnit = "M") Or (SubUnit = "G")
End Function

Public Function IsOperator(ByVal operator As String) As Boolean
'内容:
'   文字列の演算子チェックを行う
'
'[subUnit]     IN String型:     チェック対象となる文字
'
'戻り値：
'   該当する演算子であればTRUE
'   それ以外はFALSEを返す
'
'備考:
'   現在演算子として[+|-|*|/|=]を準備している
'
    IsOperator = (operator = "+") Or (operator = "-") Or (operator = "*") Or (operator = "/") Or (operator = "=")
End Function

Public Sub CheckAsString(ByVal dataStr As String)
'内容:
'   文字列としての入力制限をするためエラー処理を行う
'
'[dataStr]     IN String型:     チェック対象となる文字列
'
'備考:
'   文字列に全角文字列や数字/アルファベット/記号以外が含まれている場合、
'   実行時エラーを生成する
'
    '文字列が全角でないことをチェック
    If Not IsOneByte(dataStr) Then
        TheError.Raise 9999, "checkAsString", "[" & dataStr & "]  - 2-Byte Characters In This String Are Invalid !"
    End If
    Dim strIndex As Integer
    Dim maxIndex As Integer
    Dim CHAR As String
    maxIndex = Len(dataStr)
    '文字列がアルファベット/数字/記号であることをチェック
    For strIndex = 1 To maxIndex
        CHAR = Mid$(dataStr, strIndex, 1)
        If Not (IsNumber(CHAR) Or IsAlphabet(CHAR) Or IsSymbol(CHAR)) Then
            TheError.Raise 9999, "CheckAsString", "[" & dataStr & "]  - This Parameter Description Is Invalid !"
        End If
    Next strIndex
End Sub



Public Function SubUnitToValue(ByVal SubUnit As String) As Double
'内容:
'   補助単位文字を10進数の数値に変換する
'
'[subUnit]     IN String型:     チェック対象となる文字列
'
'戻り値：
'   変換後の数値
'
'備考:
'   現状は[p/n/u/m/%/k/M/G]に対応
'   それ以外は実行時エラーを生成をする
'
    Select Case SubUnit
        Case "":
            SubUnitToValue = 1#
        Case "p":
            SubUnitToValue = 1 * 10 ^ (-12)
        Case "n":
            SubUnitToValue = 1 * 10 ^ (-9)
        Case "u":
            SubUnitToValue = 1 * 10 ^ (-6)
        Case "m":
            SubUnitToValue = 1 * 10 ^ (-3)
        Case "%":
            SubUnitToValue = 1 * 10 ^ (-2)
        Case "k":
            SubUnitToValue = 1 * 10 ^ 3
        Case "M":
            SubUnitToValue = 1 * 10 ^ 6
        Case "G":
            SubUnitToValue = 1 * 10 ^ 9
        Case Else
            TheError.Raise 9999, "SubUnitToValue()", "[" & SubUnit & "]  - Invalid Sub Unit !"
    End Select
End Function

Public Function GetUnit(ByVal unitStr As String) As String
'内容:
'   単位及び補助単位付文字列から単位文字のみを取り出す
'
'[unitStr]     IN String型:     単位及び補助単位付文字列
'
'戻り値：
'   単位文字
'
'備考:
'
    Dim SubUnit As String
    Dim SubValue As Double
    SplitUnitValue "999" & unitStr, GetUnit, SubUnit, SubValue
End Function

Public Function DecomposeStringList(ByVal strList As String) As Collection
'内容:
'   カンマで区切られた文字列リストを分解する
'
'[strList]     IN String型:     文字列リスト
'
'戻り値：
'   分解された文字列コレクション
'
'備考:
'
    Set DecomposeStringList = New Collection
    Dim strIndex As Long
    Dim strTemp As String
    For strIndex = 1 To Len(strList)
        If Mid$(strList, strIndex, 1) = "," Then
            DecomposeStringList.Add strTemp
            strTemp = ""
        Else
            strTemp = strTemp & Mid$(strList, strIndex, 1)
        End If
    Next strIndex
    DecomposeStringList.Add strTemp
End Function

Public Function ComposeStringList(ByVal strList As Collection) As String
'内容:
'   カンマで区切られた文字列リストを作成する
'
'[strList]     IN Collection型:     文字列コレクション
'
'戻り値：
'   作成された文字列リスト
'
'備考:
'
    Dim currStr As Variant
    Dim dataIndex As Long
    For Each currStr In strList
        If dataIndex = 0 Then
            ComposeStringList = currStr
        Else
            ComposeStringList = ComposeStringList & "," & currStr
        End If
        dataIndex = dataIndex + 1
    Next currStr
End Function




Public Sub SplitUnitValue(ByVal dataStr As String, ByRef MainUnit As String, ByRef SubUnit As String, ByRef SubValue As Double)
'内容:
'   単位付文字列を数字/補助単位/単位に分離して返す
'
'[dataStr]     IN String型:     チェック対象となる文字列
'[mainUnit]    OUT String型:     単位を表す文字
'[subUnit]     OUT String型:     補助単位を表す文字
'[subValue]    OUT Double型:     文字列の中の数字
'
'備考:
'   引数MainUnitは必ずしも正しい単位を返すとは限らないので注意
'   対象の文字列が意図した単位付文字列なのかどうかのチェックを外部で行う必要がある
'
    On Error GoTo ErrorHandler
    SplitUnitValueWithoutTheError dataStr, MainUnit, SubUnit, SubValue
    Exit Sub
ErrorHandler:
    TheError.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub SplitUnitValueWithoutTheError(ByVal dataStr As String, ByRef MainUnit As String, ByRef SubUnit As String, ByRef SubValue As Double)
'内容:
'   単位付文字列を数字/補助単位/単位に分離して返す
'
'[dataStr]     IN String型:     チェック対象となる文字列
'[mainUnit]    OUT String型:     単位を表す文字
'[subUnit]     OUT String型:     補助単位を表す文字
'[subValue]    OUT Double型:     文字列の中の数字
'
'備考:
'   引数MainUnitは必ずしも正しい単位を返すとは限らないので注意
'   対象の文字列が意図した単位付文字列なのかどうかのチェックを外部で行う必要がある
'
    Dim maxIndex As Integer
    On Error GoTo ErrorHandler
    maxIndex = Len(dataStr)
    '文字数のチェック
    If (maxIndex < 1) Then GoTo ErrorHandler
    If (maxIndex = 1) Then
        If IsNumeric(dataStr) Then
            MainUnit = ""
            SubUnit = ""
            SubValue = CDbl(dataStr)
            Exit Sub
        Else
            GoTo ErrorHandler
        End If
    End If
    
    '基本単位文字を取得
    If dataStr Like "*fps" Then
        MainUnit = Right$(dataStr, 3)
        maxIndex = maxIndex - 3
    ElseIf dataStr Like "*dB" Then
        MainUnit = Right$(dataStr, 2)
        maxIndex = maxIndex - 2
    ElseIf dataStr Like "*%" Then
        MainUnit = ""
    ElseIf IsAlphabet(Right$(dataStr, 1)) = False Then
        MainUnit = ""
        SubUnit = ""
        SubValue = CDbl(dataStr)
        Exit Sub
    Else
        MainUnit = Right$(dataStr, 1)
        maxIndex = maxIndex - 1
    End If

    '補助単位文字を取得
    SubUnit = ""
    SubValue = 0#
    Dim strIndex As Integer
    Dim CHAR As String
    strIndex = maxIndex
    Do While (strIndex > 0)
        CHAR = Mid$(dataStr, strIndex, 1)
        If Not IsSubUnit(CHAR) Then Exit Do
        strIndex = strIndex - 1
    Loop
    If strIndex < maxIndex Then
        If (maxIndex - strIndex) > 1 Then GoTo ErrorHandler
        SubUnit = Mid$(dataStr, strIndex + 1, maxIndex - strIndex)
        If SubUnit = "%" And MainUnit <> "" Then GoTo ErrorHandler
    End If
    SubValue = CDbl(Left$(dataStr, strIndex))
    Exit Sub
ErrorHandler:
    Err.Raise 9999, "SplitUnitValue()", "[" & dataStr & "]  - This Parameter Description Is Invalid !"
End Sub

Function IsStringWithUnit(ByVal pStr As String) As Boolean
    Dim MainUnit As String
    Dim SubUnit As String
    Dim SubValue As Double
    On Error GoTo illegal
    Call SplitUnitValueWithoutTheError(pStr, MainUnit, SubUnit, SubValue)
    IsStringWithUnit = True
    Exit Function
illegal:
    Err.Clear
    IsStringWithUnit = False
End Function
