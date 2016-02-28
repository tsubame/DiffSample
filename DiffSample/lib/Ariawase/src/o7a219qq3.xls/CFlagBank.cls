VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFlagBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �v���O��������p�̃t���O�����Ǘ�����N���X
'
'   Revision History:
'       Data        Description
'       2010/10/07  �v���O��������p�̃t���O�����Ǘ�����@�\����������
'                   CVarBank�ꎮ�𗬗p���Ă���
'       2010/10/22  �R�����g����ǉ����ύX����
'       2010/10/28  Property���̂�ύX����(Condition �� State)
'       2011/03/04�@�s��C��(by 0145206097)
'                   ������(Initialize)�y�уR���X�g���N�^���W�b�N�̕ύX
'                   �C���^�[�Z�v�^�[�؂�ւ����W�b�N�̕ύX�y�у��\�b�h���̕ύX
'                   �ESaveLogFile��Save
'                   �ESaveMode��Dump
'                   �_���ے�(LNot)�̕����t���O���͎��̃G���[��ǉ�
'                   �r���I�_���a(LXor)�̒ǉ�
'                   �t���O�N���A(Clear)�̒ǉ�
'                   �G���[���C�Y����
'       2011/03/22  LAnd, LOr, LXor�ɑ΂��ċ󕶎�����w�肵���ꍇ�G���[�ɂȂ�悤�C��
'
'�쐬��:
'   0145184346
'

Option Explicit

Private Const ERR_NUMBER = 9999        ' Error�ԍ�
Private Const CLASS_NAME = "CFlagBank" ' Class���̂�ێ�����

'/** �t���O�o���N�����p����C���^�[�t�F�C�X **/
Private mIVarBank As IVarBank
'/** �t���O�o���N��p��VarBank�̎��� **/
Private mMyBank As CVarBank
'/** �f�t�H���g���O�t�@�C����**/
Private Const DEFAULT_FILE_NAME As String = "EeeJOBLogTheFlagBank.csv"

Public Function Add(ByRef pFlagName As String, ByRef pState As Boolean, Optional ByRef pComment As String = "")
'���e:
'   Flag���Ə����l��o�^����
'
'�p�����[�^:
'   [pFlagName] In  Flag��
'   [pState]    In  Flag�����_���l
'   [pComment]  In  �R�����g
'
'�߂�l:
'
'���ӎ���:
'   VarBank�ւ�Fixed�œo�^����
'
    If pFlagName = "" Then GoTo IsAnonymous
    On Error GoTo AlreadyExist
    Call mIVarBank.Add(pFlagName, pState, True, pComment)
    Exit Function
IsAnonymous:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".Add", "Unnamed Flag Added!"
AlreadyExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".Add", "This Flag Is Already Exist! " & pFlagName
End Function

Public Property Let state(ByRef pFlagName As String, ByRef pState As Boolean)
'���e:
'   Flag�̘_���l��ύX����
'
'�p�����[�^:
'   [pFlagName] In  Flag��
'   [pState]    In  Flag�_���l
'
'�߂�l:
'
'���ӎ���:
'   ����Flag�̓��͂��\�i�J���}��؂�j
'
    On Error GoTo NotExist
    Dim flgArray() As String
    Dim flgIndex As Integer
    flgArray = Split(Replace(pFlagName, " ", ""), ",")
    For flgIndex = 0 To UBound(flgArray)
        If (flgArray(flgIndex) <> "") Then
            With mIVarBank
                .ReadOnly(flgArray(flgIndex)) = False
                .Value(flgArray(flgIndex)) = pState
            End With
        End If
    Next flgIndex
    Exit Property
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LetState", "This Flag Is Not Exist! " & flgArray(flgIndex)
End Property

Public Property Get state(ByRef pFlagName As String) As Boolean
'���e:
'   Flag�̘_���l��Ԃ�
'
'�p�����[�^:
'   [pFlagName] In  Flag��
'
'�߂�l:
'   Flag�_���l
'
'���ӎ���:
'
    On Error GoTo NotExist
    state = mIVarBank.Value(pFlagName)
    Exit Property
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".GetState", "This Flag Is Not Exist! " & pFlagName
End Property

Public Property Let Comment(ByRef pFlagName As String, pComment As String)
'���e:
'   Flag�̃R�����g��ύX����
'
'�p�����[�^:
'   [pFlagName] In  Flag��
'   [pComment]  In  �R�����g
'
'�߂�l:
'
'���ӎ���:
'
    On Error GoTo NotExist
    mIVarBank.Comment(pFlagName) = pComment
    Exit Property
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LetComment", "This Flag Is Not Exist! " & pFlagName
End Property

Public Property Get Comment(ByRef pFlagName As String) As String
'���e:
'   Flag�̃R�����g��Ԃ�
'
'�p�����[�^:
'   [pFlagName] In  Flag��
'
'�߂�l:
'   �R�����g
'
'���ӎ���:
'
    On Error GoTo NotExist
    Comment = mIVarBank.Comment(pFlagName)
    Exit Property
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LetComment", "This Flag Is Not Exist! " & pFlagName
End Property

Public Property Get Count() As Long
'���e:
'   �o�^���Ă���Flag����Ԃ�
'
'�p�����[�^:
'
'�߂�l:
'   Flag��
'
'���ӎ���:
'
    Count = mIVarBank.Count
End Property

Public Property Get IsExist(ByRef pFlagName As String) As Boolean
'���e:
'   �w�肵��Flag���o�^����Ă��邩�ǂ����m�F����
'
'�p�����[�^:
'   [pFlagName] In  Flag��
'
'�߂�l:
'   Flag�̓o�^����(True)
'   Flag�̓o�^�Ȃ�(False)
'
'���ӎ���:
'
    IsExist = mIVarBank.IsExist(pFlagName)
End Property

Public Function Search(ByRef SearchPattern As String) As String
'���e:
'   �L�[���[�h�ɂ��Flag����
'
'�p�����[�^:
'   [SearchPattern] In  Flag�̌����L�[���[�h
'
'�߂�l:
'   Pattern�ƈ�v����Flag��(","��؂�)
'
'���ӎ���:
'
    Search = mIVarBank.Search(SearchPattern)
End Function

Public Function SearchState(ByRef pState As Boolean) As String
'���e:
'   �w�肵���_���l�Ɉ�v����Flag����Ԃ�
'
'�p�����[�^:
'   [pState]  In  Flag�_���l
'
'�߂�l:
'   �_���l�ƈ�v����Flag��(","��؂�)
'
'���ӎ���:
'
    Dim flgArray() As String
    Dim flgIndex As Integer
    flgArray = Split(Me.Search("*"), ",")
    For flgIndex = 0 To UBound(flgArray)
        If (flgArray(flgIndex) <> "") Then
            With mIVarBank
                If (pState = .Value(flgArray(flgIndex))) Then
                    SearchState = SearchState & flgArray(flgIndex) & ","
                End If
            End With
        End If
    Next flgIndex
End Function

Public Sub Clear()
'���e:
'   �o�^����Flag��S�ď�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'   2011/03/04�ǉ� 0145206097
'
    mIVarBank.AllClear
End Sub

Public Function LNot(ByRef pFlagName As String) As Boolean
'���e:
'   �o�^����Flag�̘_���ے��Ԃ��iNot�j
'
'�p�����[�^:
'   [pFlagName]  In  Flag��
'
'�߂�l:
'   �_���ے�
'
'���ӎ���:
'   �o�^����Ă���_���l�͕ύX����Ȃ�
'   ����Flag�̓��͕͂s��
'
    On Error GoTo NotExist
    LNot = Not (mIVarBank.Value(pFlagName))
    Exit Function
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LNot", "This Flag Is Not Exist! " & pFlagName
End Function

Public Function LAnd(ByRef pFlagName As String) As Boolean
'���e:
'   �o�^����Flag�̘_���ς�Ԃ��iAnd�j
'
'�p�����[�^:
'   [pFlagName] In  Flag����
'
'�߂�l:
'   �_����
'
'���ӎ���:
'   �o�^����Ă���_���l�͕ύX����Ȃ�
'   ����Flag�̓��͂��\�i�J���}��؂�j
'   �P��Flag�̏ꍇ�͂���Flag�̘_���l���Ԃ�
'
    On Error GoTo NotExist
    Dim flgArray() As String
    Dim flgIndex As Integer
    flgArray = Split(Replace(pFlagName, " ", ""), ",")
    LAnd = True
    
    If UBound(flgArray) < 0 Then
        On Error GoTo 0
        TheError.Raise ERR_NUMBER, TypeName(Me) & ".LAnd", "Please Input Flag Name!"
    End If

    For flgIndex = 0 To UBound(flgArray)
        If (flgArray(flgIndex) <> "") Then
            LAnd = LAnd And (mIVarBank.Value(flgArray(flgIndex)))
        End If
    Next flgIndex
    Exit Function
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LAnd", "This Flag Is Not Exist! " & flgArray(flgIndex)
End Function

Public Function LOr(ByRef pFlagName As String) As Boolean
'���e:
'   �o�^����Flag�̘_���a��Ԃ��iOr�j
'
'�p�����[�^:
'   [pFlagName] In  Flag����
'
'�߂�l:
'   �_���a
'
'���ӎ���:
'   �o�^����Ă���_���l�͕ύX����Ȃ�
'   ����Flag�̓��͂��\�i�J���}��؂�j
'   �P��Flag�̏ꍇ�͂���Flag�̘_���l���Ԃ�
'
    On Error GoTo NotExist
    Dim flgArray() As String
    Dim flgIndex As Integer
    flgArray = Split(Replace(pFlagName, " ", ""), ",")
    LOr = False

    If UBound(flgArray) < 0 Then
        On Error GoTo 0
        TheError.Raise ERR_NUMBER, TypeName(Me) & ".LOr", "Please Input Flag Name!"
    End If

    For flgIndex = 0 To UBound(flgArray)
        If (flgArray(flgIndex) <> "") Then
            LOr = LOr Or (mIVarBank.Value(flgArray(flgIndex)))
        End If
    Next flgIndex
    Exit Function
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LOr", "This Flag Is Not Exist! " & flgArray(flgIndex)
End Function

Public Function LXor(ByRef pFlagName As String) As Boolean
'���e:
'   �o�^����Flag�̔r���I�_���a��Ԃ��iLXor�j
'
'�p�����[�^:
'   [pFlagName] In  Flag����
'
'�߂�l:
'   �r���I�_���a
'
'���ӎ���:
'   2011/03/04�ǉ� 0145206097
'   �o�^����Ă���_���l�͕ύX����Ȃ�
'   ����Flag�̓��͂��\�i�J���}��؂�j
'   �P��Flag�̏ꍇ�͂���Flag�̘_���l���Ԃ�
'
    On Error GoTo NotExist
    Dim flgArray() As String
    Dim flgIndex As Integer
    flgArray = Split(Replace(pFlagName, " ", ""), ",")
    LXor = False
    
    If UBound(flgArray) < 0 Then
        On Error GoTo 0
        TheError.Raise ERR_NUMBER, TypeName(Me) & ".LXor", "Please Input Flag Name!"
    End If

    For flgIndex = 0 To UBound(flgArray)
        If (flgArray(flgIndex) <> "") Then
            LXor = LXor Xor (mIVarBank.Value(flgArray(flgIndex)))
        End If
    Next flgIndex
    Exit Function
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LXor", "This Flag Is Not Exist! " & flgArray(flgIndex)
End Function

Public Sub Initialize()
'���e:
'   TheFlagBank�̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Set mMyBank = New CVarBank
    Set mIVarBank = mMyBank
End Sub

Public Sub Dump(ByVal pDump As Boolean)
'���e:
'   TheFlagBank�̃_���v
'
'�p�����[�^:
'   [pDump]    In  �_���v���[�h�w��
'
'�߂�l:
'
'���ӎ���:
'
    Dim inDumping As Boolean
    inDumping = IsInterceptor
    If pDump And Not inDumping Then
        Dim myInterceptor As CVarBankInterceptor
        Set myInterceptor = New CVarBankInterceptor
        Call myInterceptor.Initialize(mMyBank)
        Set mIVarBank = myInterceptor
        TheExec.Datalog.WriteComment "Eee JOB Output Log! :TheFlagBank Condition"
    ElseIf Not pDump And inDumping Then
        Set mIVarBank = mMyBank
    End If
End Sub

Public Function Save(Optional saveFileName As String = DEFAULT_FILE_NAME) As Long
'���e:
'   �_���v�f�[�^���O���t�@�C���֏o�͂���
'
'�p�����[�^:
'   [SaveFileName]  In  �t�@�C�����w��i�I�v�V�����j
'
'�߂�l:
'
'���ӎ���:
'   �t�@�C�����̎w�肪�Ȃ��ꍇ�̓f�t�H���g���ŃJ�����g�t�H���_�ɏo�͂���
'
    If Not IsInterceptor Then Exit Function
    If saveFileName = "" Then saveFileName = DEFAULT_FILE_NAME
    Dim myInterceptor As CVarBankInterceptor
    Set myInterceptor = mIVarBank
    Call myInterceptor.SaveLogFile(saveFileName)
    Set mIVarBank = mMyBank
End Function

Private Function IsInterceptor() As Boolean
'/* ���ɃC���^�[�Z�v�^�[�𗘗p���Ă��邩�ǂ������^�`�F�b�N�Ŕ��f���� */
    Dim myInterceptor As CVarBankInterceptor
    On Error GoTo IsInterceptor
    Set myInterceptor = mIVarBank
    IsInterceptor = True
    Exit Function
IsInterceptor:
    IsInterceptor = False
End Function

Private Sub Class_Initialize()
    Me.Initialize
End Sub

Private Sub Class_Terminate()
    Set mMyBank = Nothing
    Set mIVarBank = Nothing
End Sub