VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �ėp���[�N�V�[�gReader
'
'�ړI:
'   �Ώۃ��[�N�V�[�g��Ǎ���
'   �ŏ��ɃL���b�V�������A�ȍ~�L���b�V���𗘗p
'
'�쐬��:
'   a_oshima
'
'�O�����:
'   �O���[�o���I�u�W�F�N�g TheError:CErrManager �����݂��邱��
'
'���p�菇:
'   �@�C���X�^���X������AInitialize���\�b�h���R�[���B��_�Z���͏����ʒu�ɐݒ肳���
'       ��SetLocation�ɂ����PrimaryKey���T�[�`���Ċ�_�Z����ݒ肷�邱�Ƃ��\
'   �ADo While ��EndOfRow�𒴂��Ȃ��ԁA
'       �A-1 IparameterReader��ReadAs���\�b�h�ŕK�v�ȗ�̒l��Ǎ���
'       �A-2 MoveNext�Ŏ��̍s�ɐi��
'   �BDo While�I���
'

Option Explicit

Implements IParameterReader
Implements IFileStream

Private WithEvents mEventWatcher As CEventWatcher      '�ǂݏo���Ώۃ��[�N�V�[�g��
Attribute mEventWatcher.VB_VarHelpID = -1
Private mTargetWorkSheet As Worksheet      '�ǂݏo���Ώۃ��[�N�V�[�g��
Attribute mTargetWorkSheet.VB_VarHelpID = -1
Private mSheetInfo As CSheetInfo    '�f�[�^���x�����ƃZ���A�h���X�̑Ή���ێ�����B�萔�̂悤�ɗ��p�B�N���X��Initialize�C�x���g�ŏ������B
Private mEnableReportCurrentCell As Boolean

Private Const Comment As String = "Comment"

Private mCache As Variant                           '�L���b�V��
Private mCacheRowOffset As Long                     '�L���b�V���̃I�t�Z�b�g
Private mCacheColumnOffset As Long                  '�L���b�V���̃I�t�Z�b�g
Private mCacheRowCount As Long                      '�L���b�V���̑傫��
Private mCacheColumnCount As Long                   '�L���b�V���̑傫��
Private mCurrentRow As Long                         '��_�Z��(mDateCell)�̑���
Private mCurrentColumn As Long                      '��_�Z��(mDateCell)�̑���
Private mEndOfRow As Long                           'EndOfRow(mEorCell)�̑���(�f�[�^�̏I���̍s�̃Z��)

Public Event CacheUpdate(Msg As String)
Public Event ButtonClick(Msg As String)


Property Get sheetInfo() As CSheetInfo
'���e:
'   CSheetInfo�̃C���X�^���X��Ԃ��܂�
'
'�p�����[�^:
'   �Ȃ�
'
'�߂�l:
'    CSheetInfo�̃C���X�^���X
'
'��O:
'   SheetInfo�̃C���X�^���X�����쐬�̂Ƃ��͗�O9999�𓊂��܂�
'
'���ӎ���:

    If mSheetInfo Is Nothing Then
        Call TheError.Raise(9999, TypeName(Me) & ".SheetInfo", "WorkSheet Reader has not initialized !")
    Else
        Set sheetInfo = mSheetInfo
    End If
End Property



Property Let EnableReportCurrentCell(ByVal pFlagEnable As Boolean)
'���e:
'   �f�o�b�O�p�B��_�Z���񍐋@�\��On/Off���܂��B
'   �@�h��_�Z���񍐋@�\�h�F�Ǎ��݊�_�Z�����ړ�����x�Ɋ�_�Z���̈ʒu���C�~�f�B�G�C�g�E�C���h�E�ɏo�͂��܂��B
'
'�p�����[�^:
'   [pFlgEnable]  In  Boolean�^:  True�F��_�Z���񍐋@�\��On�AFalse�F��_�Z���񍐋@�\��Off
'
'�߂�l:
'   �Ȃ�
'
'���ӎ���:
'
    mEnableReportCurrentCell = pFlagEnable
End Property
Property Get EnableReportCurrentCell() As Boolean
'���e:
'   �f�o�b�O�p�B��_�Z���񍐋@�\�̌��ݐݒ��Ԃ��܂��B
'
'�p�����[�^:
'   �Ȃ�
'
'�߂�l:
'    True�F��_�Z���񍐋@�\On�AFalse�F��_�Z���񍐋@�\Off
'
'���ӎ���:
'
    EnableReportCurrentCell = mEnableReportCurrentCell
End Property

Private Sub Init()
    
    With mTargetWorkSheet.UsedRange
        mCache = .Value
        mCacheRowOffset = .Row - 1
        mCacheColumnOffset = .Column - 1
        mCacheRowCount = .Rows.Count
        mCacheColumnCount = .Columns.Count
    End With
    If mEnableReportCurrentCell = True Then
        Debug.Print "Cache Offset(" & mCacheRowOffset & "," & mCacheColumnOffset & ")"
        Debug.Print "Cache Size(" & mCacheRowCount & "," & mCacheColumnCount & ")"
    End If
    
    Me.AsIFileStream.Rewind
    updataEor
    
End Sub
Public Sub Initialize(ByRef pBook As Workbook, ByVal pSheetName As String, ByRef pSheetInfoFactory As ISheetInfoFactory)
'���e:
'   ���������\�b�h�B���L���s���B
'     �@�Ǎ��ݑΏۃ��[�N�V�[�g���������o�ϐ��ɐݒ�
'     �A�L���b�V���捞
'     �BRewind
'     �CEndOfRow�X�V
'
'�p�����[�^:
'   [pBook]       IN  Workbook�^:  �Ǎ��ݑΏۂ̃u�b�N�̃C���X�^���X
'   [pSheetName]  IN  String�^:    �Ǎ��ݑΏۂ̃��[�N�V�[�g��
'   [pSheetInfoFactory]  IN  CSheetInfoFactory�^:   CSheetInfoFactory�̃C���X�^���X
'
'�߂�l:
'   �Ȃ�
'
'��O:
'   �����������Ɏ��s����Ɨ�O9999�𓊂��܂�
'
'���ӎ���:
'
'
    
    On Error GoTo ErrorHandler
    Set mTargetWorkSheet = pBook.Worksheets(pSheetName)
    Set mSheetInfo = pSheetInfoFactory.CreateSheetInfo(mTargetWorkSheet)
    Set mEventWatcher = New CEventWatcher
    Set mEventWatcher.SetSheet = mTargetWorkSheet
    Call Init
    Exit Sub
ErrorHandler:
    Call TheError.Raise(9999, TypeName(Me) & ".Initialize", "Can Not Initialize WorkSheet Reader !")
End Sub

Private Sub Class_Initialize()
'    Set mSheetInfo = New CSheetInfo
End Sub

Private Sub Class_Terminate()
    mCache = Empty
    Set mEventWatcher = Nothing
    Set mSheetInfo = Nothing
    Set mTargetWorkSheet = Nothing
End Sub



Public Function AsIFileStream() As IFileStream
'���e:
'   IFileStream�^�ւ̃L���X�g
'
'�p�����[�^:
'   �Ȃ�
'
'�߂�l:
'    IFileStream�^�̎��I�u�W�F�N�g
'
    Set AsIFileStream = Me
End Function

Private Sub IFileStream_SetLocation(ByVal pDataName As String)
'���e:
'   PrimaryKey���T�[�`���Ċ�_�Z�����ړ������܂�
'   �iFind���\�b�h�ň����Ɏw�肵��������Ɉ�v����Z����T�������܂��j
'
'�p�����[�^:
'    [pDataName] IN  String�^:  Test Instance Name����T�[�`���镶����
'
'��O:
'    �T�[�`�̌��ʁA������Ȃ������ꍇ�͗�O9999�𓊂��܂�
'
'�߂�l:
'    �Ȃ�
'
    Call CheckInitializeCalled
    Dim col As String
    col = GetCacheColumn(GetAddress(mSheetInfo.GetPrimaryKey))
    Dim i As Long
    Dim flgFound As Boolean
    For i = 1 To mCacheRowCount
        If mCache(mCacheRowOffset + i, col) = pDataName Then
            flgFound = True
            mCurrentRow = i
            mCurrentColumn = col
        End If
    Next i
    If flgFound = False Then
        TheError.Raise 9999, TypeName(Me) & ".SetLocation", "Can Not Find The Primary Key [" & pDataName & "] !"
    End If
    
    If mEnableReportCurrentCell = True Then
        Debug.Print "SetLocation: Cache(" & mCurrentRow & "," & mCurrentColumn & ")"
    End If
End Sub

Private Sub IFileStream_MoveNext()
'���e:
'   ��_�Z����1�i�߂܂��i�s�����j
'
'�߂�l:
'    �Ȃ�
'
    Call CheckInitializeCalled
    mCurrentRow = mCurrentRow + 1
    If mEnableReportCurrentCell = True Then
        Debug.Print "MoveNext: Cache(" & mCurrentRow & "," & mCurrentColumn & ")"
    End If
End Sub

Private Sub IFileStream_Rewind()
'���e:
'   ��_�Z���������ʒu�ɖ߂��܂��i�����ʒu�FPrimary Key�̌��o���̒����̃f�[�^�Z���j
'
'�߂�l:
'    �Ȃ�
'
    Call CheckInitializeCalled
    mCurrentRow = GetCacheRow(GetAddress(mSheetInfo.GetPrimaryKey)) + 1
    mCurrentColumn = GetCacheColumn(GetAddress(mSheetInfo.GetPrimaryKey))
'�V�[�g�I�u�W�F�N�g�擾���̓���ɍ��킹�ă`�F�b�N���Ȃ��B2011/03/22
'    If mCurrentRow > mCacheRowOffset + mCacheRowCount Then
'        TheError.Raise 9999, TypeName(Me) & ".Rewind", "Can Not Find The Beginning Of Data !"
'    End If
    
    If mEnableReportCurrentCell = True Then
        Debug.Print "Rewind: Cache(" & mCurrentRow & "," & mCurrentColumn & ")"
    End If
End Sub

Private Function IFileStream_IsEOR() As Boolean
'���e:
'   ��_�Z����EndOfRow�𒴂����ǂ����𔻒肵�܂�
'
'�߂�l:
'    True�F��_�Z����EndOfRow�𒴂����AFalse�F��_�Z����EndOfRow�𒴂��Ă��Ȃ�
    Call CheckInitializeCalled
    IFileStream_IsEOR = (mCurrentRow > mEndOfRow)
End Function



Public Function AsIParameterReader() As IParameterReader
'���e:
'   IParameterReader�^�ւ̃L���X�g
'
'�p�����[�^:
'   �Ȃ�
'
'�߂�l:
'    IParameterReader�^�̎��I�u�W�F�N�g
'
    Set AsIParameterReader = Me
End Function

Private Function IParameterReader_ReadAsBoolean(ByVal pDataName As String) As Boolean
    Dim Value As Variant
    Value = mCache(mCurrentRow, dataColumn(pDataName))
    If IsNumeric(Value) Then
        IParameterReader_ReadAsBoolean = CBool(Value)
    Else
        TheError.Raise 9999, TypeName(Me) & ".ReadAsBoolean", "Does Not Numeric:" & Value & " Cache(" & mCurrentRow & "," & dataColumn(pDataName) & ")"
    End If
End Function

Private Function IParameterReader_ReadAsDouble(ByVal pDataName As String) As Double
    Dim Value As Variant
    Value = mCache(mCurrentRow, dataColumn(pDataName))
    If IsNumeric(Value) Then
        IParameterReader_ReadAsDouble = CDbl(Value)
    Else
        TheError.Raise 9999, TypeName(Me) & ".ReadAsDouble", "Does Not Numeric:" & Value & " Cache(" & mCurrentRow & "," & dataColumn(pDataName) & ")"
    End If
End Function

Private Function IParameterReader_ReadAsInteger(ByVal pDataName As String) As Integer
'pDataName��!Row���w�肳���ƃV�[�g�̍s�ԍ����Ԃ�
    Dim Value As Variant
    If pDataName = "!Row" Then
        Value = mCurrentRow
    Else
        Value = mCache(mCurrentRow, dataColumn(pDataName))
    End If
    If IsNumeric(Value) Then
        IParameterReader_ReadAsInteger = CInt(Value)
    Else
        TheError.Raise 9999, TypeName(Me) & ".ReadAsInteger", "Does Not Numeric:" & Value & " Cache(" & mCurrentRow & "," & dataColumn(pDataName) & ")"
    End If
End Function

Private Function IParameterReader_ReadAsLong(ByVal pDataName As String) As Long
'pDataName��!Row���w�肳���ƃV�[�g�̍s�ԍ����Ԃ�
    Dim Value As Variant
    If pDataName = "!Row" Then
        Value = mCurrentRow
    Else
        Value = mCache(mCurrentRow, dataColumn(pDataName))
    End If
    If IsNumeric(Value) Then
        IParameterReader_ReadAsLong = CLng(Value)
    Else
        TheError.Raise 9999, TypeName(Me) & ".ReadAsLong", "Does Not Numeric:" & Value & " Cache(" & mCurrentRow & "," & dataColumn(pDataName) & ")"
    End If
End Function

Private Function IParameterReader_ReadAsString(ByVal pDataName As String) As String
'���e:
'   �f�[�^���x������������o���A�l��ǂݎ���ĕ�����^�ŕԂ��܂�
'
'�p�����[�^:
'    [pDataName] IN  String�^:  �f�[�^���x���i���͑啶��Property�Ŏ擾�\�B�܂�Label�őS�Ă��擾�\�j
'
'��O:
'   �f�[�^���x�����z��O�̏ꍇ�A��O9999�𓊂��܂�
'
'�߂�l:
'    ���肳�ꂽ�Z���̒l�iValue���\�b�h�̖߂�l�j
    IParameterReader_ReadAsString = mCache(mCurrentRow, dataColumn(pDataName))
End Function


Private Function dataColumn(ByVal pDataName As String) As Long
    Call CheckInitializeCalled
    If IFileStream_IsEOR Then
        Call TheError.Raise(9999, TypeName(Me) & ".dataColumn", "Out Of Range !")
    End If
    If IsExist(pDataName) Then
        dataColumn = GetCacheColumn(GetAddress(pDataName))
    Else
        Call TheError.Raise(9999, TypeName(Me) & ".dataCell", "Can Not Find The Parameter [" & pDataName & "] !")
    End If
    If mEnableReportCurrentCell = True Then
        Debug.Print "dataCell: Cache(" & mCurrentRow & "," & GetCacheColumn(GetAddress(pDataName)) & ")"
    End If
End Function

Private Function IsExist(ByVal pKey As String) As Boolean
    IsExist = mSheetInfo.IsExist(pKey)
End Function

Private Sub updataEor()

    If mSheetInfo.IsExist(Comment) Then
        mCacheColumnCount = mTargetWorkSheet.Range(mSheetInfo.GetAddress(Comment)).Column - mCacheColumnOffset
    End If

    mEndOfRow = mCacheRowOffset + mCacheRowCount

    Dim i As Long
    For i = mCurrentRow To mCacheRowCount
        Dim flg_empty As Boolean
        flg_empty = True
        Dim j As Long
        For j = 1 To mCacheColumnCount
            If mCache(i, j) <> vbNullString Then
                flg_empty = False
                Exit For
            End If
        Next j
        If flg_empty = True Then
            mEndOfRow = i - 1 + mCacheRowOffset
            Exit For
        Else
            Dim col As String
            col = GetCacheColumn(GetAddress(mSheetInfo.GetPrimaryKey))
            If mCache(i, col) = "END" Then
                mEndOfRow = i - 1 + mCacheRowOffset
                Exit For
            End If
        End If
    Next i

    mCacheRowCount = mEndOfRow

    If mEnableReportCurrentCell = True Then
       Debug.Print "updateEor: Cache EndOfRow: " & mEndOfRow
    End If
End Sub

Private Sub CheckInitializeCalled()
    If mTargetWorkSheet Is Nothing Then
        TheError.Raise 9999, TypeName(Me) & ".CheckInitializeCalled", "Call Initialize method first!"
    End If
End Sub

Private Function GetCacheRow(ByVal pAddress As String) As Long
    Dim rowStr As Long
    If pAddress Like "[A-Z][A-Z]*" Then
        rowStr = Strings.Mid$(pAddress, 3)
    ElseIf pAddress Like "[A-Z]*" Then
        rowStr = Strings.Mid$(pAddress, 2)
    Else
        TheError.Raise 9999, TypeName(Me) & ".GetCacheRow", "Illegal Address:" & pAddress
    End If
    
    Dim Row As Long
    If Not IsNumeric(rowStr) Then
        TheError.Raise 9999, TypeName(Me) & ".GetCacheRow", "Illegal Address:" & pAddress
    End If
    Row = CLng(rowStr)
    
    GetCacheRow = Row - mCacheRowOffset
End Function

Private Function GetCacheColumn(ByVal pAddress As String) As Long
    Dim secondFigure As Long
    Dim firstFigure As Long
    If pAddress Like "[A-Z][A-Z]*" Then
        secondFigure = Asc(Strings.Left$(pAddress, 1)) - Asc("A") + 1
        firstFigure = Asc(Strings.Mid$(pAddress, 2, 1)) - Asc("A") + 1
    ElseIf pAddress Like "[A-Z]*" Then
        secondFigure = 0
        firstFigure = Asc(Strings.Left$(pAddress, 1)) - Asc("A") + 1
    Else
        TheError.Raise 9999, TypeName(Me) & ".GetCacheColumn", "Illegal Address:" & pAddress
    End If
    
    Dim col As Long
    col = secondFigure * 26 + firstFigure
      
    GetCacheColumn = col - mCacheColumnOffset
End Function
Private Function GetAddress(ByVal pName As String) As String
'���o���̃A�h���X��Ԃ�
    GetAddress = mSheetInfo.GetAddress(pName)
End Function



Private Sub mEventWatcher_ButtonClick(Msg As String)
    RaiseEvent ButtonClick(Msg)
End Sub

Private Sub mEventWatcher_WorksheetChange(ByRef TargetRng As Range)
    Call UpdateCache
    RaiseEvent CacheUpdate("EndOfRow:" & mEndOfRow)
End Sub


'Public Property Get WorkSheetCanged() As Boolean
'    WorkSheetCanged = False
'End Property

Private Sub UpdateCache()
    Erase mCache
    Call Init
    If mEnableReportCurrentCell = True Then
        MsgBox "Update Cache!" & vbCrLf _
            & "  EndOfRow:" & mEndOfRow & vbCrLf _
            & "  Cache Size:(" & mCacheRowCount & "," & mCacheColumnCount & ")"
    End If
End Sub

Public Function DataAddress(ByVal pDataName As String) As String
'���e:
'   ��_�Z���̂���s�̎w��f�[�^�̃A�h���X��Ԃ��܂�
'   �i�G���[���b�Z�[�W�̃Z���ʒu���񋟗p�j
'
'�p�����[�^:
'    [pDataName] IN  String�^:  �f�[�^���x��
'
'��O:
'   �f�[�^���x�����z��O�̏ꍇ�A��O9999�𓊂��܂�
'
'�߂�l:
'    ���肳�ꂽ�Z���̃A�h���X
    Dim addr As String
    addr = GetAddress(pDataName)
    Dim w As Long
    If addr Like "[A-Z][A-Z][0-9]*" Then
        w = 2
    ElseIf addr Like "[A-Z][0-9]*" Then
        w = 1
    End If
    DataAddress = Left$(addr, w) & mCurrentRow
End Function

Public Property Get endRow() As Long
'���e:
'   �ŏI�s��������

    endRow = mEndOfRow
End Property