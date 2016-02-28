VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcScenarioLoopOptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   DcSccenarioLoopOptionForm�Ǘ��N���X
'
'�ړI:
'   DcSccenarioLoopOptionForm�ɂ���
'   �E�\��
'   �E���擾
'   �E�{�^���C�x���g�Ǘ�
'   ���s��
'
'�쐬��:
'   0145184306
'
Option Explicit

Private Const FORM_CAPTION As String = "DC Test Scenario Looping Option"

Private mDcLoop As Boolean
Private mCategories As Collection

Private WithEvents mLoopOptionForm As DcScenarioLoopOptionForm
Attribute mLoopOptionForm.VB_VarHelpID = -1

Private WithEvents mRunButton As MSForms.CommandButton
Attribute mRunButton.VB_VarHelpID = -1
Private WithEvents mCancelButton As MSForms.CommandButton
Attribute mCancelButton.VB_VarHelpID = -1
Private WithEvents mMoveRightButton As MSForms.CommandButton
Attribute mMoveRightButton.VB_VarHelpID = -1
Private WithEvents mMoveRightAllButton As MSForms.CommandButton
Attribute mMoveRightAllButton.VB_VarHelpID = -1
Private WithEvents mMoveLeftButton As MSForms.CommandButton
Attribute mMoveLeftButton.VB_VarHelpID = -1
Private WithEvents mMoveLeftAllButton As MSForms.CommandButton
Attribute mMoveLeftAllButton.VB_VarHelpID = -1

Private mCategoriesList As MSForms.ListBox
Private mLoopCategoriesList As MSForms.ListBox

Private WithEvents mCountsOptionButton As MSForms.OptionButton
Attribute mCountsOptionButton.VB_VarHelpID = -1
Private WithEvents mTimeOptionButton As MSForms.OptionButton
Attribute mTimeOptionButton.VB_VarHelpID = -1
Private mParamText As MSForms.TextBox
Private mUnitLabel As MSForms.Label
Private Const COUNTS_LOWER_LIMIT As Long = 1
Private Const COUNTS_UPPER_LIMIT As Long = 5000
Private Const COUNTS_OF_LIMIT As Long = 50000
Private Const TIME_LOWER_LIMIT As Double = 1
Private Const TIME_UPPER_LIMIT As Double = 180
Private Const TIME_OF_LIMIT As Long = 3600

Private Const COUNTS_UNIT As String = "[times]"
Private Const TIME_UNIT As String = "[sec]"


Private WithEvents mFolderReferenceButton As MSForms.CommandButton
Attribute mFolderReferenceButton.VB_VarHelpID = -1
Private mFolderPathText As MSForms.TextBox

Public Function Show() As Boolean
'���e:
'   Form�̕\��
'
'�p�����[�^:
'
'�߂�l:
'   �ǂ̃{�^���������ďI��������    Boolean�^
'       [Run]�{�^��:True
'       ���̑�     :False
'���l:
'   Initialize�ŕ\������J�e�S���[��ݒ肵�Ă��Ȃ��ꍇ�A�\�������ɏI��
'
    
    Dim res As Boolean
    res = False
    
    If mCategories Is Nothing Then
        res = False
    Else
        '���̃��X�g�ɑS�A�C�e���\��
        Dim category As Variant
        For Each category In mCategories
            mCategoriesList.AddItem category
        Next category
        
        '�J��Ԃ�����������
        mParamText.Value = ""
        mCountsOptionButton.Value = True
        mUnitLabel = COUNTS_UNIT
        
        '�t�H���_�p�X�����l�ݒ�(JOB�t�@�C���p�X)
        mFolderPathText.Value = ThisWorkbook.Path
        
        mLoopOptionForm.Show
        
        res = mDcLoop
    End If
    
    Show = res
End Function

Public Sub Initialize(ByVal pCategories As Collection)
    Set mCategories = createCategoriesNameList(pCategories)
End Sub

Public Property Get LoopCondition() As LOOP_CONDITION
'���e:
'   �ݒ���e�̎擾
'
'�p�����[�^:
'
'�߂�l:
'   �ݒ���e    LOOP_CONDITION�^
'
'���l:
'
'
    Dim loopCategories As Collection
    Dim samplingParam As SAMPLING_PARAM
    Dim paramLimit As Double
    Dim folderPath As String
    
    'Loop�ΏۃJ�e�S���[
    Set loopCategories = New Collection
    Dim index As Integer
    For index = 0 To mLoopCategoriesList.ListCount - 1
        loopCategories.Add mLoopCategoriesList.List(index), mLoopCategoriesList.List(index)
    Next index
    
    'Counts/Time�I�����
    If mCountsOptionButton.Value = True Then
        samplingParam = PARAM_COUNTS
    Else
        samplingParam = PARAM_TIME
    End If
    
    '�T���v�����O���~�b�g
    On Error GoTo CAST_ERROR
    If mCountsOptionButton.Value = True Then
        Dim cast As Long
        cast = CLng(mParamText.Value)
        paramLimit = CDbl(cast)
    Else
        paramLimit = CDbl(mParamText.Value)
    End If
    
    '�t�H���_�p�X
    folderPath = mFolderPathText.Value
    
    Dim res As LOOP_CONDITION
    With res
        Set .LOOP_CATEGORIES = loopCategories
        .Param = samplingParam
        .PARAM_LIMIT = paramLimit
        .FOLDER = folderPath
    End With
    
    LoopCondition = res
    
    Set loopCategories = Nothing
    Exit Property
CAST_ERROR:
    Set loopCategories = Nothing
    Err.Raise 9999, "CDcScenarioLoopOptionForm.LoopCondition", ""
    
End Property

Private Sub mMoveRightButton_Click()
'[>]�{�^��������
    Call moveListItems(mCategoriesList, mLoopCategoriesList)
End Sub

Private Sub mMoveRightAllButton_Click()
'[>>]�{�^��������
    Call moveAllListItems(mCategoriesList, mLoopCategoriesList)
End Sub

Private Sub mMoveLeftButton_Click()
'[<]�{�^��������
    Call moveListItems(mLoopCategoriesList, mCategoriesList)
End Sub

Private Sub mMoveLeftAllButton_Click()
'[<<]�{�^��������
    Call moveAllListItems(mLoopCategoriesList, mCategoriesList)
End Sub

Private Sub mRunButton_Click()
'���e:
'   [Run]�{�^��������
'
'�p�����[�^:
'
'�߂�l:
'
'���l:
'   �t�H���_�A�N�Z�X�`�F�b�N�ƃT���v�����O�p�����[�^�̃`�F�b�N���s��
'
    If checkLimit = True Then
        If checkFolderAccess = True Then
            mDcLoop = True
            mLoopOptionForm.Hide
        End If
    End If
End Sub

Private Sub mCancelButton_Click()
'[Cancel]�{�^��������
    MsgBox "Canceled loop option." & vbCrLf & "Run on the normal mode.", vbInformation, FORM_CAPTION
    
    mDcLoop = False
    mLoopOptionForm.Hide
End Sub
Private Sub mLoopOptionForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'[X]�{�^��������
    If CloseMode = 0 Then
        mCancelButton_Click
    End If
End Sub


Private Sub mCountsOptionButton_Click()
'[Counts]�I�v�V�����I����
    mUnitLabel = COUNTS_UNIT
End Sub

Private Sub mTimeOptionButton_Click()
'[Time]�I�v�V�����I����
    mUnitLabel = TIME_UNIT
End Sub

Private Sub mFolderReferenceButton_Click()
'���e:
'   Folder�Q�ƃ{�^��
'
'�p�����[�^:
'
'�߂�l:
'
'���l:
'
'
'�t�H���_�I���{�^��������
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            mFolderPathText.Value = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub Class_Initialize()
'���e:
'   �N���X������
'
'���l:
'
'
    Set mCategories = Nothing
    mDcLoop = False
    
    Set mLoopOptionForm = DcScenarioLoopOptionForm
    With mLoopOptionForm
        Set mRunButton = .RunButton
        Set mCancelButton = .CancelButton
        
        Set mMoveRightButton = .MoveRightButton
        Set mMoveLeftButton = .MoveLeftButton
        Set mMoveRightAllButton = .MoveRightAllButton
        Set mMoveLeftAllButton = .MoveLeftAllButton
        
        Set mCategoriesList = .CategoriesList
        mCategoriesList.ColumnWidths = mCategoriesList.width + 10
        Set mLoopCategoriesList = .LoopCategoriesList
        mLoopCategoriesList.ColumnWidths = mLoopCategoriesList.width + 10
        
        Set mCountsOptionButton = .CountsOptionButton
        Set mTimeOptionButton = .TimeOptionButton
        Set mParamText = .ParamText
        Set mUnitLabel = .UnitLabel
        
        Set mFolderPathText = .FolderPathText
        Set mFolderReferenceButton = .FolderReferenceButton
    End With
    
End Sub

Private Sub Class_Terminate()
'���e:
'   �N���X�j��
'
'���l:
'
'
    Set mRunButton = Nothing
    Set mCancelButton = Nothing
    
    Set mMoveRightButton = Nothing
    Set mMoveLeftButton = Nothing
    Set mMoveRightAllButton = Nothing
    Set mMoveLeftAllButton = Nothing
    
    Set mCategoriesList = Nothing
    Set mLoopCategoriesList = Nothing
    
    Set mCountsOptionButton = Nothing
    Set mTimeOptionButton = Nothing
    Set mParamText = Nothing
    Set mUnitLabel = Nothing
    
    Set mFolderPathText = Nothing
    Set mFolderReferenceButton = Nothing
    
    Unload mLoopOptionForm
    Set mLoopOptionForm = Nothing
    
    Set mCategories = Nothing
    mDcLoop = False
End Sub

Private Function createCategoriesNameList(ByRef pCategories As Collection) As Collection
    Dim res As Collection
    
    If pCategories Is Nothing Then
        Set res = Nothing
    Else
        Set res = New Collection
        
        Dim cat As CDcCategory
        For Each cat In pCategories
            Call res.Add(cat.Name)
        Next cat
    End If
    
    Set createCategoriesNameList = res
    Set res = Nothing
    
End Function

Private Sub moveAllListItems(ByRef pFrom As MSForms.ListBox, ByRef pTo As MSForms.ListBox)
'���e:
'   ListBox�Ԃ̑S�A�C�e���ړ�����
'
'�p�����[�^:
'[pFrom]       IN/OUT MSForms.ListBox�^:  �ړ���ListBox
'[pTo]         IN/OUT MSForms.ListBox�^:  �ړ���ListBox
'
'�߂�l:
'
'���l:
'
    pFrom.Clear
    With pTo
        .Clear
        Dim category As Variant
        For Each category In mCategories
            .AddItem category
        Next category
    End With
End Sub

Private Sub moveListItems(ByRef pFrom As MSForms.ListBox, ByRef pTo As MSForms.ListBox)
'���e:
'   ListBox�Ԃ̑I���A�C�e���ړ�����
'
'�p�����[�^:
'[pFrom]       IN/OUT MSForms.ListBox�^:  �ړ���ListBox
'[pTo]         IN/OUT MSForms.ListBox�^:  �ړ���ListBox
'
'�߂�l:
'
'���l:
'
    With pFrom
        If .ListIndex >= 0 Then
            Dim index As Integer
            Dim removeIndex As Collection
            Set removeIndex = New Collection
            For index = 0 To .ListCount - 1
                If .Selected(index) Then
                    pTo.AddItem .List(index)
                    removeIndex.Add index
                    .Selected(index) = False
                End If
            Next index
            If removeIndex.Count > 0 Then
                Call GetRemoveList(removeIndex)
                Dim retIndex As Variant
                For Each retIndex In removeIndex
                    Call .RemoveItem(retIndex)
                Next retIndex
            End If
        End If
    End With
End Sub

Private Function checkFolderAccess() As Boolean
'���e:
'   �e�L�X�g�{�b�N�X�Ɏw�肳�ꂽ�t�H���_�̃A�N�Z�X�m�F
'   �A�N�Z�X���m�F�ׁ̈A���ۂɃt�@�C���������݂��s���B
'
'�߂�l:
'   �t�H���_�A�N�Z�X��    Boolean�^
'���l:
'
'
        
    Const ACCESS_TEST_FILE As String = "AccessTest"
    
    Dim tmpFolder As String
    tmpFolder = mFolderPathText.Value
    If tmpFolder = "" Then
        MsgBox "Selected folder cannot be accessed.", vbCritical, FORM_CAPTION
        checkFolderAccess = False
        Exit Function
    End If
    
    '�h���C�u�G���[���p
    On Error Resume Next
    If Dir(tmpFolder, vbDirectory) = "" Then
        On Error GoTo FolderCreateError
        MkDir tmpFolder
        checkFolderAccess = True
    Else
        Dim FileNo As Integer
        FileNo = FreeFile
        Dim checkFile As String
        checkFile = mFolderPathText.Value & Application.PathSeparator & ACCESS_TEST_FILE
    
        On Error GoTo FileAccessError
        Open checkFile For Output As FileNo
        On Error GoTo 0
        Close FileNo
        Kill checkFile
        checkFolderAccess = True
    
    End If
    Exit Function
FolderCreateError:
    MsgBox "Selected folder cannot be created.", vbCritical, FORM_CAPTION
    checkFolderAccess = False
    Exit Function

FileAccessError:
    MsgBox "Selected folder cannot be accessed.", vbCritical, FORM_CAPTION
    Close FileNo
    checkFolderAccess = False
End Function

Private Function checkLimit() As Boolean
'���e:
'   �e�L�X�g�{�b�N�X�Ɏw�肳�ꂽ�p�����[�^�̃`�F�b�N
'   �^�`�F�b�N�Ə��/�����l�`�F�b�N���s��
'
'�߂�l:
'   �`�F�b�N����    Boolean�^
'���l:
'
'
    Dim res As Boolean
    res = True
    
    Dim doOverLimit As VbMsgBoxResult
    
    'CountsOption�̏ꍇ
    If mCountsOptionButton.Value = True Then
        Dim checkCountVal As Long
        On Error GoTo OVERFLOW
        checkCountVal = CLng(mParamText.Value)
        'OverFlow�K��l�ȏ�
        If checkCountVal >= COUNTS_OF_LIMIT Then
            MsgBox mParamText.Value & " is invalid parameter.", vbCritical, FORM_CAPTION
            res = False
        '�����l�ȉ�
        ElseIf checkCountVal < COUNTS_LOWER_LIMIT Then
            MsgBox mParamText.Value & " is invalid parameter.", vbCritical, FORM_CAPTION
            res = False
        '����l�ȏ�
        ElseIf checkCountVal > COUNTS_UPPER_LIMIT Then
            doOverLimit = MsgBox(checkCountVal & " has exceeded the upper limit " & COUNTS_UPPER_LIMIT & "." & vbCrLf & _
                            "'OK' to run at this parameter." _
                                , vbOKCancel + vbExclamation, FORM_CAPTION)
            If doOverLimit <> vbOK Then res = False
        End If
    'TimeOption�̏ꍇ
    Else
        Dim checkTimeVal As Double
        On Error GoTo OVERFLOW
        checkTimeVal = CDbl(mParamText.Value)
        'OverFlow�K��l�ȏ�
        If checkTimeVal >= TIME_OF_LIMIT Then
            MsgBox mParamText.Value & " is invalid parameter.", vbCritical, FORM_CAPTION
            res = False
        '�����l�ȉ�
        ElseIf checkTimeVal < TIME_LOWER_LIMIT Then
            MsgBox mParamText.Value & " is invalid parameter.", vbCritical, FORM_CAPTION
            res = False
        '����l�ȏ�
        ElseIf checkTimeVal > TIME_UPPER_LIMIT Then
            doOverLimit = MsgBox(checkTimeVal & " has exceeded the upper limit " & TIME_UPPER_LIMIT & "." & vbCrLf & _
                            "'OK' to run at this parameter." _
                                , vbOKCancel + vbExclamation, FORM_CAPTION)
            If doOverLimit <> vbOK Then res = False
        End If
    End If
    
    checkLimit = res
    Exit Function
OVERFLOW:
    res = False
    MsgBox mParamText.Value & " is invalid parameter.", vbCritical, FORM_CAPTION
    checkLimit = res
    
End Function

Private Sub GetRemoveList(ByRef pIndexList As Collection)
'���e:
'   ������Index���X�g���\�[�g���ĕԂ�
'
'[pIndexList]  IN/OUT Collection�^:  Index���X�g
'
'���l:
'
'
    Dim removeIndexArr() As Long
    ReDim removeIndexArr(pIndexList.Count - 1)
    Dim retIndex As Variant
    Dim tgtIndex As Long
    tgtIndex = 0
    For Each retIndex In pIndexList
        removeIndexArr(tgtIndex) = retIndex
        tgtIndex = tgtIndex + 1
    Next retIndex
    Call DescendingSort(removeIndexArr, LBound(removeIndexArr), UBound(removeIndexArr))
    Set pIndexList = New Collection
    For tgtIndex = 0 To UBound(removeIndexArr)
        pIndexList.Add removeIndexArr(tgtIndex)
    Next tgtIndex
End Sub

Private Sub DescendingSort(ByRef pData As Variant, ByVal pLower As Long, ByVal pUpper As Long)
'���e:
'   �f�[�^�̃\�[�g���s��
'
'[pData]       IN/OUT Variant�^:    �\�[�g����f�[�^
'[pLower]      IN     Long�^:       �ŏ��l
'[pUpper]      IN     Long�^:       �ő�l
'
'���l:
'
'
    Dim i As Long
    Dim j As Long
    Dim S As Variant
    Dim tmp As Variant
        S = pData(Int((pLower + pUpper) / 2))
        i = pLower
        j = pUpper
        Do
            Do While pData(i) > S
                i = i + 1
            Loop
            Do While pData(j) < S
                j = j - 1
            Loop
            If i >= j Then Exit Do
            tmp = pData(i)
            pData(i) = pData(j)
            pData(j) = tmp
            i = i + 1
            j = j - 1
      Loop
      If (pLower < i - 1) Then DescendingSort pData, pLower, i - 1
      If (pUpper > j + 1) Then DescendingSort pData, j + 1, pUpper
 End Sub