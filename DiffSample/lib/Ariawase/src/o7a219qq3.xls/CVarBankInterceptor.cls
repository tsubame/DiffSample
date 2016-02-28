VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVarBankInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   CVarBank�N���X�̊��荞�ݗp�N���X
'�ړI:
'   CVarBank �N���X�̗��p�󋵂̏����W�p
'   �C���X�^���X�� Set �Ń��O�o�͂̃I���I�t�؂�ւ���B
'   ���O�o�̓I�t���̃I�[�o�[�w�b�h�������B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��22��

Option Explicit

Implements IVarBank

Private m_Container As IVarBank                                        '���̃N���X�����荞�ޑΏۃN���X
Private m_LogCollection As Collection                                       '���O�i�[�p

Private m_Counter As Long                                                   '���O�o�̓V���A���ԍ��p

Private m_SaveLogFile As CFileOut                                           '���O�t�@�C�����o�͐���p�C���X�^���X
Private Const LOG_FILE_NAME As String = "EeeJOBLogTheVarBank.csv"            '���̃N���X�̃��O�o�̓t�@�C����

Public Sub Initialize(p As CVarBank)
    Set m_Container = p
End Sub

Public Function AsIVarBank() As IVarBank
    Set AsIVarBank = Me
End Function

Public Sub SaveLogFile(Optional ByVal fileName As String)
    
    '#################### ���O�o�͂̏����ݒ� ####################
    Dim OutputString As String
    Dim RetFileName   As String
    Dim RetJobName   As String
    Dim i As Long
    
    If fileName = "" Then fileName = LOG_FILE_NAME
    
    '#################### �t�@�C���o�� ####################
    '#################### �w�b�_���� ####################
    Set m_SaveLogFile = New CFileOut
    Call m_SaveLogFile.SetFileName(fileName)
    RetFileName = TheExec.ExcelHandle.ActiveWorkbook.Path & "\" & TheExec.ExcelHandle.ActiveWorkbook.Name
    RetJobName = TheExec.CurrentJob                             'JobList �� JobName ���B������΃t�@�C����
    OutputString = "[RecordNumber], [RecordDate], [RecordTime], [TestInstanceName], [ExecuteMethod], [VariableName], [Value], [Status], [Fixed], [ReadOnly], [Comment]"
    '#################### �w�b�_�o�� ####################
    Call m_SaveLogFile.OpenFile
    Print #m_SaveLogFile.fileNum, "[Test Program],[" & RetFileName & "]"
    Print #m_SaveLogFile.fileNum, "[Job],[" & RetJobName & "]"
    Print #m_SaveLogFile.fileNum, OutputString
    
    '#################### �f�[�^�o�� ####################
    For i = 1 To m_LogCollection.Count
        Print #m_SaveLogFile.fileNum, m_LogCollection.Item(CStr(i))
    Next i
    Call m_SaveLogFile.CloseFile
    
    '#################### ���O������ ####################
    m_Counter = 0
    Set m_LogCollection = New Collection
    
End Sub

Private Property Get IVarBank_Value(VariableName As String) As Variant
    Dim i As Long       'TheError check�p
    i = CheckTheError
    
    IVarBank_Value = m_Container.Value(VariableName)
    
    '#################### ���O�o�^ ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                '���\�b�h���s�����ꍇ
            OutputString = LogFormat(VariableName, "-", "GetValue", "Error", "-", "-", "-")
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "GetValue", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), "-")
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Property

Private Property Let IVarBank_Value(VariableName As String, pValue As Variant)
    Dim i As Long       'TheError check�p
    i = CheckTheError
    
    m_Container.Value(VariableName) = pValue
    
    '#################### ���O�o�^ ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                '���\�b�h���s�����ꍇ
            OutputString = LogFormat(VariableName, CStr(pValue), "LetValue", "Error", "-", "-", "-")
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "LetValue", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), "-")
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Property

Private Property Get IVarBank_ReadOnly(VariableName As String) As Boolean
    IVarBank_ReadOnly = m_Container.ReadOnly(VariableName)
End Property

Private Property Let IVarBank_ReadOnly(VariableName As String, pReadOnly As Boolean)
    Dim i As Long       'TheError check�p
    i = CheckTheError
    
    m_Container.ReadOnly(VariableName) = pReadOnly
    
    '#################### ���O�o�^ ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                '���\�b�h���s�����ꍇ
            OutputString = LogFormat(VariableName, "-", "LetReadOnly", "Error", "-", CStr(pReadOnly), "-")
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "LetReadOnly", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), "-")
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Property

Private Property Get IVarBank_Fixed(VariableName As String) As Boolean
    IVarBank_Fixed = m_Container.Fixed(VariableName)
End Property

Private Property Let IVarBank_Fixed(VariableName As String, pFixed As Boolean)
    Dim i As Long       'TheError check�p
    i = CheckTheError
    
    m_Container.Fixed(VariableName) = pFixed
    
    '#################### ���O�o�^ ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                '���\�b�h���s�����ꍇ
            OutputString = LogFormat(VariableName, "-", "LetFixed", "Error", CStr(pFixed), "-", "-")
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "LetFixed", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), "-")
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Property

Private Property Get IVarBank_Comment(VariableName As String) As String
    IVarBank_Comment = m_Container.Comment(VariableName)
End Property

Private Property Let IVarBank_Comment(VariableName As String, pComment As String)
    Dim i As Long       'TheError check�p
    i = CheckTheError
    
    m_Container.Comment(VariableName) = pComment
    
    '#################### ���O�o�^ ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                '���\�b�h���s�����ꍇ
            OutputString = LogFormat(VariableName, "-", "LetComment", "Error", "-", "-", CStr(pComment))
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "LetComment", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), pComment)
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Property

Private Property Get IVarBank_Count() As Long
    IVarBank_Count = m_Container.Count
End Property

Private Function IVarBank_Add(VariableName As String, FirstValue As Variant, Optional FixedStatus As Boolean, Optional Comment As String) As Boolean
    Dim i As Long       'TheError check�p
    i = CheckTheError
    
    Call m_Container.Add(VariableName, FirstValue, FixedStatus, Comment)
    
    '#################### ���O�o�^ ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                '���\�b�h���s�����ꍇ
            OutputString = LogFormat(VariableName, CStr(FirstValue), "Add", "Error", CStr(FixedStatus), "-", CStr(Comment))
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "Add", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), CStr(.Comment(VariableName)))
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Function

Private Function IVarBank_IsExist(VariableName As String) As Boolean
    IVarBank_IsExist = m_Container.IsExist(VariableName)
End Function

Private Function IVarBank_Search(ByVal SearchPattern As String, Optional NotFlag As Boolean = False) As String
    IVarBank_Search = m_Container.Search(SearchPattern, NotFlag)
End Function

Private Sub IVarBank_Clear()
    
    Call m_Container.Clear

    '#################### ���O�o�^ ####################
    Dim OutputString As String
    OutputString = LogFormat("-", "-", "Clear", "-", "-", "-", "-")
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
    
End Sub

Private Sub IVarBank_AllClear()
    
    Call m_Container.AllClear

    '#################### ���O�o�^ ####################
    Dim OutputString As String
    OutputString = LogFormat("-", "-", "AllClear", "-", "-", "-", "-")
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
    
End Sub

'####### �ȉ��C���^�[�t�F�C�X�Ɗ֌W�Ȃ����\�b�h #######

Private Function LogFormat( _
    key As String, _
    Value As String, _
    Method As String, _
    status As String, _
    Fixed As String, _
    ReadOnly As String, _
    Comment As String _
    ) As String
'���e:
'   ���O���͗p�̏����󂯂āA�J���}��؂�Ő��^����String�^�ŕԂ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N4��26��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   String
'
'���ӎ���:

    Dim p_InstanceName As String
    
    If TheExec.Flow.IsRunning Then
        p_InstanceName = TheExec.DataManager.InstanceName
    Else
        p_InstanceName = "-"
    End If
    
    m_Counter = m_Counter + 1
    
    '����Ԃ��B
    LogFormat = _
                m_Counter & "," & _
                CStr(DateTime.Date) & "," & _
                CStr(DateTime.TIME) & "," & _
                p_InstanceName & "," & _
                Method & "," & _
                key & "," & _
                Value & "," & _
                status & "," & _
                Fixed & "," & _
                ReadOnly & "," & _
                Comment

End Function

Private Function CheckTheError() As Long
'TheError.LastError �� �V���A���ԍ���Ԃ��܂��B
'Nothing ��0��Ԃ��܂��B

    If TheError.LastError Is Nothing Then
        CheckTheError = 0
    Else
        CheckTheError = TheError.LastError.SerialNumber
    End If

End Function

Private Sub Class_Initialize()
'�R���X�g���N�^

    m_Counter = 0
    Set m_Container = New CVarBank
    Set m_LogCollection = New Collection
    
End Sub

Private Sub Class_Terminate()
    m_Counter = 0
    Set m_Container = Nothing
    Set m_SaveLogFile = Nothing
    Set m_LogCollection = Nothing
End Sub
