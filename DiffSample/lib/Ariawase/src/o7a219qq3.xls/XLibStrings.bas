Attribute VB_Name = "XLibStrings"
'�T�v:
'   ������Ɋւ���`�F�b�N��ϊ����s���v���V�[�W���Q
'
'�ړI:
'   �p�����[�^�N���X�p�̕�����`�F�b�N�y�ѕϊ��v���V�[�W���Q
'   �O���[�o���ł��g�p�o����悤���ʉ�����
'
'�쐬��:
'   0145206097
'
Option Explicit

Public Function IsOneByte(ByVal strData As String) As Boolean
'���e:
'   ��������S�Ă̕����̑S�p/���p�`�F�b�N���s��
'
'[strData]     IN String�^:     �`�F�b�N�ΏۂƂȂ镶����
'
'�߂�l�F
'   �S�Ĕ��p�̏ꍇ��TRUE
'   ��ł��S�p������ꍇ��FALSE��Ԃ�
'
'���l:
'   �Ȃ�
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
'���e:
'   �����̐����`�F�b�N���s��
'
'[char]     IN String�^:     �`�F�b�N�ΏۂƂȂ镶��
'
'�߂�l�F
'   �����ł���ꍇ��TRUE
'   ����ȊO��FALSE��Ԃ�
'
'���l:
'   �Ȃ�
'
    IsNumber = (CHAR >= "0") And (CHAR <= "9")
End Function

Public Function IsSymbol(ByVal CHAR As String) As Boolean
'���e:
'   �����̋L���`�F�b�N���s��
'
'[char]     IN String�^:     �`�F�b�N�ΏۂƂȂ镶��
'
'�߂�l�F
'   �}�C�i�X/�J���}/�A���_�[�X�R�A/�p�[�Z���g�̏ꍇ��TRUE
'   ����ȊO��FALSE��Ԃ�
'
'���l:
'   �Ȃ�
'
    IsSymbol = (CHAR = "-") Or (CHAR = ".") Or (CHAR = "_") Or (CHAR = "%")
End Function

Public Function IsAlphabet(ByVal CHAR As String) As Boolean
'���e:
'   �����̃A���t�@�x�b�g�`�F�b�N���s��
'
'[char]     IN String�^:     �`�F�b�N�ΏۂƂȂ镶��
'
'�߂�l�F
'   �A���t�@�x�b�g�̏ꍇ��TRUE
'   ����ȊO��FALSE��Ԃ�
'
'���l:
'   �啶���������͖��Ȃ�
'
    IsAlphabet = ((CHAR >= "a") And (CHAR <= "z")) Or ((CHAR >= "A") And (CHAR <= "Z"))
End Function

Public Function IsSubUnit(ByVal SubUnit As String) As Boolean
'���e:
'   �����̕⏕�P�ʃ`�F�b�N���s��
'
'[subUnit]     IN String�^:     �`�F�b�N�ΏۂƂȂ镶��
'
'�߂�l�F
'   �Y������⏕�P�ʂł����TRUE
'   ����ȊO��FALSE��Ԃ�
'
'���l:
'   ���ݕ⏕�P�ʂƂ���[p/n/u/m/%/k/M/G]���������Ă���
'
    IsSubUnit = (SubUnit = "p") Or (SubUnit = "n") Or (SubUnit = "u") Or (SubUnit = "m") Or (SubUnit = "%") Or (SubUnit = "k") Or (SubUnit = "M") Or (SubUnit = "G")
End Function

Public Function IsOperator(ByVal operator As String) As Boolean
'���e:
'   ������̉��Z�q�`�F�b�N���s��
'
'[subUnit]     IN String�^:     �`�F�b�N�ΏۂƂȂ镶��
'
'�߂�l�F
'   �Y�����鉉�Z�q�ł����TRUE
'   ����ȊO��FALSE��Ԃ�
'
'���l:
'   ���݉��Z�q�Ƃ���[+|-|*|/|=]���������Ă���
'
    IsOperator = (operator = "+") Or (operator = "-") Or (operator = "*") Or (operator = "/") Or (operator = "=")
End Function

Public Sub CheckAsString(ByVal dataStr As String)
'���e:
'   ������Ƃ��Ă̓��͐��������邽�߃G���[�������s��
'
'[dataStr]     IN String�^:     �`�F�b�N�ΏۂƂȂ镶����
'
'���l:
'   ������ɑS�p������␔��/�A���t�@�x�b�g/�L���ȊO���܂܂�Ă���ꍇ�A
'   ���s���G���[�𐶐�����
'
    '�����񂪑S�p�łȂ����Ƃ��`�F�b�N
    If Not IsOneByte(dataStr) Then
        TheError.Raise 9999, "checkAsString", "[" & dataStr & "]  - 2-Byte Characters In This String Are Invalid !"
    End If
    Dim strIndex As Integer
    Dim maxIndex As Integer
    Dim CHAR As String
    maxIndex = Len(dataStr)
    '�����񂪃A���t�@�x�b�g/����/�L���ł��邱�Ƃ��`�F�b�N
    For strIndex = 1 To maxIndex
        CHAR = Mid$(dataStr, strIndex, 1)
        If Not (IsNumber(CHAR) Or IsAlphabet(CHAR) Or IsSymbol(CHAR)) Then
            TheError.Raise 9999, "CheckAsString", "[" & dataStr & "]  - This Parameter Description Is Invalid !"
        End If
    Next strIndex
End Sub



Public Function SubUnitToValue(ByVal SubUnit As String) As Double
'���e:
'   �⏕�P�ʕ�����10�i���̐��l�ɕϊ�����
'
'[subUnit]     IN String�^:     �`�F�b�N�ΏۂƂȂ镶����
'
'�߂�l�F
'   �ϊ���̐��l
'
'���l:
'   �����[p/n/u/m/%/k/M/G]�ɑΉ�
'   ����ȊO�͎��s���G���[�𐶐�������
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
'���e:
'   �P�ʋy�ѕ⏕�P�ʕt�����񂩂�P�ʕ����݂̂����o��
'
'[unitStr]     IN String�^:     �P�ʋy�ѕ⏕�P�ʕt������
'
'�߂�l�F
'   �P�ʕ���
'
'���l:
'
    Dim SubUnit As String
    Dim SubValue As Double
    SplitUnitValue "999" & unitStr, GetUnit, SubUnit, SubValue
End Function

Public Function DecomposeStringList(ByVal strList As String) As Collection
'���e:
'   �J���}�ŋ�؂�ꂽ�����񃊃X�g�𕪉�����
'
'[strList]     IN String�^:     �����񃊃X�g
'
'�߂�l�F
'   �������ꂽ������R���N�V����
'
'���l:
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
'���e:
'   �J���}�ŋ�؂�ꂽ�����񃊃X�g���쐬����
'
'[strList]     IN Collection�^:     ������R���N�V����
'
'�߂�l�F
'   �쐬���ꂽ�����񃊃X�g
'
'���l:
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
'���e:
'   �P�ʕt������𐔎�/�⏕�P��/�P�ʂɕ������ĕԂ�
'
'[dataStr]     IN String�^:     �`�F�b�N�ΏۂƂȂ镶����
'[mainUnit]    OUT String�^:     �P�ʂ�\������
'[subUnit]     OUT String�^:     �⏕�P�ʂ�\������
'[subValue]    OUT Double�^:     ������̒��̐���
'
'���l:
'   ����MainUnit�͕K�������������P�ʂ�Ԃ��Ƃ͌���Ȃ��̂Œ���
'   �Ώۂ̕����񂪈Ӑ}�����P�ʕt������Ȃ̂��ǂ����̃`�F�b�N���O���ōs���K�v������
'
    On Error GoTo ErrorHandler
    SplitUnitValueWithoutTheError dataStr, MainUnit, SubUnit, SubValue
    Exit Sub
ErrorHandler:
    TheError.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub SplitUnitValueWithoutTheError(ByVal dataStr As String, ByRef MainUnit As String, ByRef SubUnit As String, ByRef SubValue As Double)
'���e:
'   �P�ʕt������𐔎�/�⏕�P��/�P�ʂɕ������ĕԂ�
'
'[dataStr]     IN String�^:     �`�F�b�N�ΏۂƂȂ镶����
'[mainUnit]    OUT String�^:     �P�ʂ�\������
'[subUnit]     OUT String�^:     �⏕�P�ʂ�\������
'[subValue]    OUT Double�^:     ������̒��̐���
'
'���l:
'   ����MainUnit�͕K�������������P�ʂ�Ԃ��Ƃ͌���Ȃ��̂Œ���
'   �Ώۂ̕����񂪈Ӑ}�����P�ʕt������Ȃ̂��ǂ����̃`�F�b�N���O���ōs���K�v������
'
    Dim maxIndex As Integer
    On Error GoTo ErrorHandler
    maxIndex = Len(dataStr)
    '�������̃`�F�b�N
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
    
    '��{�P�ʕ������擾
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

    '�⏕�P�ʕ������擾
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
