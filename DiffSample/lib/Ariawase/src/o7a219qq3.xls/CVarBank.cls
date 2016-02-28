VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVarBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   Public�ϐ��̑�p�N���X
'�ړI:
'   Public�ϐ��̎g�p�ɂ��g���u�����p�ł��B
'   ���̃N���X�ɓo�^�����l�͌Œ�ƂȂ�A�s�p�ӂȕύX�̃��X�N������܂��B
'   �l���C�����悤�Ƃ���ƃG���[�ƂȂ�܂��B�l���C������K�v������ꍇ�́A�ݒ��ύX���Ă���C�����܂��B
'   ����TOPT�g�p���́A�v���V�[�W�������r���ŕʂ̃v���V�[�W���֏������ڂ邽�߁A�Ӑ}���Ȃ���������������\�������蒍�ӂ��K�v�ł��B
'���L����:
'   ##### �Ԃ�l�́A���͎��Ɠ����^�Ŏ󂯂Ă��������B�Ⴄ�^�Ŏ󂯂��ꍇ�A�S���Ⴄ�l�ƂȂ邱�Ƃ�����܂� #####
'   ��jDouble�^��Integer�^�Ŏ󂯂�
'
'   �L�[�����̑召�Ⴂ�͓����Ƃ݂Ȃ��܂��B
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��12��

Option Explicit

Implements IVarBank

Private m_VarBankElements As Collection         '�v�f��Collection
Private m_FixedVarBankElements As Collection    '�v�f��Collection Fixed�͕ʃI�u�W�F�N�g�ŕێ��BClear���\�b�h�ŏ�������Ȃ��B

'Like���Z�q�@�啶�������������I�v�V����
Option Compare Text

Private Property Get IVarBank_Value(VariableName As String) As Variant
'���e:
'   �o�^�ς݂̕ϐ����̒l���擾���܂��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]  In  1):�ϐ���
'�߂�l:
'   Variant
'
'���ӎ���:
'
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".GetValue")
    If Not (p_Element Is Nothing) Then
        IVarBank_Value = p_Element.val
    End If
    Set p_Element = Nothing
End Property

Private Property Let IVarBank_Value(VariableName As String, pValue As Variant)
'���e:
'   �o�^�ς݂̕ϐ��ɒl��ݒ肵�܂��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]  In  1):�ϐ���
'   [pValue]        In  2):�ݒ�l
'�߂�l:
'
'���ӎ���:
'   �ݒ肷��ɂ� ReadOnly �v���p�e�B��False �ɐݒ肵�Ȃ��Ƃ����܂���B
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".LetValue")
    If Not (p_Element Is Nothing) Then
        If p_Element.ReadOnly Then
            Call TheError.Raise(9999, TypeName(Me) & ".LetValue", "This Variable is Now Status ReadOnly! " & VariableName)
            Exit Property
        Else
            p_Element.val = pValue
        End If
    End If
    Set p_Element = Nothing
End Property

Private Property Get IVarBank_ReadOnly(VariableName As String) As Boolean
'���e:
'   �ϐ���ReadOnly�v���p�e�B��Ԃ��擾���܂��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]  In  1):�ϐ���
'�߂�l:
'   Boolean
'
'���ӎ���:
'
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".GetReadOnly")
    If Not (p_Element Is Nothing) Then
        IVarBank_ReadOnly = p_Element.ReadOnly
    End If
    Set p_Element = Nothing
End Property

Private Property Let IVarBank_ReadOnly(VariableName As String, pReadOnly As Boolean)
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]  In  1):�ϐ���
'   [pReadOnly]     In  2):�ݒ肷����
'�߂�l:
'
'���ӎ���:
'
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".LetReadOnly")
    If Not (p_Element Is Nothing) Then
        p_Element.ReadOnly = pReadOnly
    End If
    Set p_Element = Nothing
End Property

Private Property Get IVarBank_Fixed(VariableName As String) As Boolean
'���e:
'   �ێ���Ԃ��ǂ������擾���܂��B
'   �ێ���Ԃɂ���ƁAClear���\�b�h�ɂĔj������܂���B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]  In  1):�ϐ���
'�߂�l:
'   Boolean
'
'���ӎ���:
'   AllClear���\�b�h�ł͔j������܂��B
'

    If ExistFixed(VariableName) Then
        IVarBank_Fixed = True
        Exit Property
    ElseIf ExistNonFixed(VariableName) Then
        IVarBank_Fixed = False
        Exit Property
    Else
        Call TheError.Raise(9999, TypeName(Me) & ".GetFixed", "This Variable is not Exist! " & VariableName)
    End If
End Property

Private Property Let IVarBank_Fixed(VariableName As String, pFixed As Boolean)
'���e:
'   �ێ���Ԃ�ݒ肵�܂��B
'   �ێ���Ԃɂ���ƁAClear���\�b�h�ɂĔj������܂���B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]  In  1):�ϐ���
'   [pFixed]        In  2):�ݒ肷��ێ����
'�߂�l:
'
'���ӎ���:
'   AllClear���\�b�h�ł͔j������܂��B
'

    If ExistFixed(VariableName) Then
        If pFixed = False Then
            Call m_VarBankElements.Add(m_FixedVarBankElements.Item(VariableName), VariableName)
            Call m_FixedVarBankElements.Remove(VariableName)
            Exit Property
        End If
    ElseIf ExistNonFixed(VariableName) Then
        If pFixed = True Then
            Call m_FixedVarBankElements.Add(m_VarBankElements.Item(VariableName), VariableName)
            Call m_VarBankElements.Remove(VariableName)
            Exit Property
        End If
    Else
        Call TheError.Raise(9999, TypeName(Me) & ".LetFixed", "This Variable is not Exist! " & VariableName)
    End If
End Property

Private Property Get IVarBank_Comment(VariableName As String) As String
'���e:
'   �ϐ��̃R�����g���Q�Ƃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]  In  1):�ϐ���
'�߂�l:
'   String
'
'���ӎ���:
'
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".GetComment")
    If Not (p_Element Is Nothing) Then
        IVarBank_Comment = p_Element.Comment
    End If
    Set p_Element = Nothing
End Property

Private Property Let IVarBank_Comment(VariableName As String, pComment As String)
'���e:
'   �ϐ��̃R�����g��ύX���܂��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]  In  1):�ϐ���
'   [pComment]      In  2):�R�����g
'�߂�l:
'
'���ӎ���:
'
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".LetComment")
    If Not (p_Element Is Nothing) Then
        p_Element.Comment = pComment
    End If
    Set p_Element = Nothing
End Property

Private Property Get IVarBank_Count() As Long
'���e:
'   �o�^����Ă���ϐ��̐���Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Long
'
'���ӎ���:
'
'

    If Not (m_VarBankElements Is Nothing) Then                           '�o�^�ϐ��Q�����邩�ǂ���
        IVarBank_Count = m_VarBankElements.Count
    End If
    If Not (m_FixedVarBankElements Is Nothing) Then                   '�o�^�ϐ��Q(�Œ��)�����邩�ǂ���
        IVarBank_Count = IVarBank_Count + m_FixedVarBankElements.Count
    End If
End Property

Private Function IVarBank_Add(VariableName As String, FirstValue As Variant, Optional FixedStatus As Boolean, Optional Comment As String) As Boolean
'���e:
'   �ϐ����ƒl��o�^���܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��23��
'�p�����[�^:
'   [VariableName]       In  1):�o�^��
'   [Value]              In  2):�o�^�l
'   [Fixed]              In  3):�ێ����
'   [Comment]            In  4):�o�^�Ɋւ��ẴR�����g
'�߂�l:
'
'���ӎ���:
'

    If IVarBank_IsExist(VariableName) Then                                                          '�o�^�ϐ��Q�ɕϐ������邩�ǂ���
        Call TheError.Raise(9999, TypeName(Me) & ".Add", "This Variable Already is Exist! " & VariableName)
        IVarBank_Add = False
        Exit Function
    Else                                                                                    '�ϐ��ɓo�^����Ă��Ȃ�
        Dim pVariable As CVarBankElement
        Set pVariable = New CVarBankElement
        Call pVariable.Initialize(VariableName, FirstValue, Comment)
        If FixedStatus Then
            Call m_FixedVarBankElements.Add(pVariable, VariableName)
        Else
            Call m_VarBankElements.Add(pVariable, VariableName)
        End If
    End If
    
    IVarBank_Add = True
    
End Function

Private Function IVarBank_IsExist(VariableName As String) As Boolean
'���e:
'   �Y������VariableName���o�^�ς݂��ǂ�����Boolean�ŕԂ��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]       In  1):�ϐ���
'�߂�l:
'   Boolean
'���ӎ���:
'
    
    If ExistFixed(VariableName) Then
        IVarBank_IsExist = True
        Exit Function
    ElseIf ExistNonFixed(VariableName) Then
        IVarBank_IsExist = True
        Exit Function
    End If
    
    IVarBank_IsExist = False

End Function

Private Function IVarBank_Search(ByVal SearchPattern As String, Optional NotFlag As Boolean = False) As String
'���e:
'   �o�^����Ă���ϐ�����ASearchPattern�ɍ����ϐ���T���܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��15��
'�p�����[�^:
'   [SearchPattern] In  1):�T������
'   [NotFlag]       In  2):True:�܂܂Ȃ�������Ԃ��BFalse:�܂ޕ�����Ԃ��B
'�߂�l:
'   String
'
'���ӎ���:
'   �����̑召�Ⴂ�͖�������܂�
'   SearchPattern��Like���Z�q�Ŕ�r����܂�
'
'  ���T������SearchPattern��
'   �S���T�������ꍇ                  �F"*"
'   ����Test���t��������T�������ꍇ  �F"Test*"
'   ������7���t��������T�������ꍇ   �F"*7"
'   HL���܂ޕ�����T�������ꍇ        �F"*HL*"


    Dim i As Integer
    Dim var As Variant
    Dim keyName As String
    Dim retStr As String
    
    If Not (m_VarBankElements Is Nothing) Then
        For Each var In m_VarBankElements
            If i > 5000 Then
                MsgBox "SearchVariable Function. Too Much Variable!"
                Exit Function
            End If
            If (var.Name Like SearchPattern) Xor NotFlag Then
                retStr = retStr & var.Name & ","
                i = i + 1
            End If
        Next var
    End If
    
    If Not (m_FixedVarBankElements Is Nothing) Then
        For Each var In m_FixedVarBankElements
            If i > 5000 Then
                MsgBox "SearchVariable Function. Too Much Variable!"
                Exit Function
            End If
            If (var.Name Like SearchPattern) Xor NotFlag Then
                retStr = retStr & var.Name & ","
                i = i + 1
            End If
        Next var
    End If
    
    IVarBank_Search = retStr

End Function

Private Sub IVarBank_Clear()
'���e:
'   �o�^���ꂽ�ϐ��Q�����������܂��BFixed�ł͏��������܂���B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��14��
'�p�����[�^:
'�߂�l:
'
'���ӎ���:
'
    Set m_VarBankElements = Nothing
    Set m_VarBankElements = New Collection
End Sub

Private Sub IVarBank_AllClear()
'���e:
'   �o�^���ꂽ�ϐ��Q��S�ď��������܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��14��
'�p�����[�^:
'�߂�l:
'
'���ӎ���:
'
    Set m_VarBankElements = Nothing
    Set m_VarBankElements = New Collection
    Set m_FixedVarBankElements = Nothing
    Set m_FixedVarBankElements = New Collection
End Sub

Private Function ExistNonFixed(VariableName As String) As Boolean
'���e:
'   �Y������VariableName��NonFixed collection�ɓo�^�ς݂��ǂ�����Boolean�ŕԂ��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]       In  1):�ϐ���
'�߂�l:
'   Boolean
'���ӎ���:
'
    Dim p_Element As CVarBankElement
    
    On Error GoTo NOT_FOUND_NONFIX
    Set p_Element = m_VarBankElements.Item(VariableName)
    ExistNonFixed = True
    Set p_Element = Nothing
    Exit Function
    
NOT_FOUND_NONFIX:
    
    ExistNonFixed = False

End Function

Private Function ExistFixed(VariableName As String) As Boolean
'���e:
'   �Y������VariableName��Fixed collection�ɓo�^�ς݂��ǂ�����Boolean�ŕԂ��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]       In  1):�ϐ���
'�߂�l:
'   Boolean
'���ӎ���:
'
    Dim p_Element As CVarBankElement
    
    On Error GoTo NOT_FOUND_FIX
    Set p_Element = m_FixedVarBankElements.Item(VariableName)
    ExistFixed = True
    Set p_Element = Nothing
    Exit Function
    
NOT_FOUND_FIX:
    
    ExistFixed = False

End Function

Private Function GetVariable(VariableName As String, MethodName As String) As CVarBankElement
'���e:
'   �Y������ϐ������݂��邩�ǂ����B���݂���΃I�u�W�F�N�g��Ԃ��B�������Nothing
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [VariableName]       In  1):�ϐ���
'�߂�l:
'   CVarBankElement
'
'���ӎ���:
'

    If ExistFixed(VariableName) Then
        Set GetVariable = m_FixedVarBankElements.Item(VariableName)
        Exit Function
    ElseIf ExistNonFixed(VariableName) Then
        Set GetVariable = m_VarBankElements.Item(VariableName)
        Exit Function
    End If
    
    Call TheError.Raise(9999, TypeName(Me) & MethodName, "This Variable is not Exist! " & VariableName)
    Set GetVariable = Nothing

End Function

Private Sub Class_Initialize()
    Set m_VarBankElements = New Collection
    Set m_FixedVarBankElements = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_VarBankElements = Nothing
    Set m_FixedVarBankElements = Nothing
End Sub
