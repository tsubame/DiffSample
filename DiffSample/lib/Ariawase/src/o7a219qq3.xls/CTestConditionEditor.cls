VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestConditionEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   Form�̃R���|�[�l���g�̋@�\�����p�N���X
'
'�ړI:
'   ���I�����ݒ�N���X�Ə����ݒ�ҏWForm�Ƃ̏�������肷��
'
'�쐬��:
'   0145184935 K.Hara
'
Option Explicit

Private mActive As Boolean  'TestConditionEditor�t�H�[����Active�t���O
Private mExecute As Boolean 'Execute���s�ς݃t���O

Private mArgTextBoxs As Collection              'Arg�e�L�X�g�{�b�N�X
Private mProcedureRow As MSComctlLib.ListItem   '�I������Procedure�s�̏��

Private mCurConditionName As String          '���݂�Condition��
Private mFunctionNameList As Collection      '�֐������X�g
Private mFunctionParamList As Collection     '�p�����[�^���X�g(Key�͊֐����j
Private mCurExecuteSettingName As String     '���݂̃}�N���O���[�v
Private Const MAX_ARG_NUMBER As Long = 9     'Arg�p�����[�^�̌�

'TestConditionControler�t�H�[���̊e�R���g���[���I�u�W�F�N�g
Private WithEvents mTestConditionController As TestConditionController 'TestConditionController�t�H�[����F�m����ׂ̃����o
Attribute mTestConditionController.VB_VarHelpID = -1
Private WithEvents mReloadButton As MSForms.CommandButton           'Reload�{�^��
Attribute mReloadButton.VB_VarHelpID = -1
Private WithEvents mExecutebutton As MSForms.CommandButton          'Execute�{�^��
Attribute mExecutebutton.VB_VarHelpID = -1
Private WithEvents mExitButton As MSForms.CommandButton            'Exit�{�^��
Attribute mExitButton.VB_VarHelpID = -1
Private WithEvents mProcedureRowListViewCtrl As MSComctlLib.ListView   'ListView�E�B���h�E
Attribute mProcedureRowListViewCtrl.VB_VarHelpID = -1

Private Const ENABLE_COLOR As Long = &H80000008 'ListView�s�̗L���F
Private Const DISABLE_COLOR As Long = &H808080  'ListView�s�̖����F

'ConditionParameterEditor�t�H�[���̊i�R���g���[���I�u�W�F�N�g
Private mConditionParameterEditor As ConditionParameterEditor 'ConditionParamterEdit�t�H�[����F�m����ׂ̃����o
Private WithEvents mApplyButton As MSForms.CommandButton    'Apply�{�^��
Attribute mApplyButton.VB_VarHelpID = -1
Private WithEvents mCancelButton As MSForms.CommandButton   'Cancel�{�^��
Attribute mCancelButton.VB_VarHelpID = -1

Private Const TERMINATE_CODE As String = "#EOP"     '�I�[������
Private Const ALERT_COLOR As Long = &HFFF           '���f��̕����F
Private Const ORIGINAL_COLOR As Long = &H80000008   '���f�O�̕����F

'���̃N���X�Ŕ���������C�x���g�BCDynamicSetCondition�N���X���L���b�`����B
Public Event LoadConditionInfo(ByVal pConditionInfoName As String)
Public Event ControlExecuteMacro(ByVal pCurConditionName As String, ByVal pFunctionNameList As Collection, ByVal pFunctionParamList As Collection)

Public Sub Show(ByVal pConditionNameList As Collection, ByVal pConditionName As String)
'���e:
'   �t�H�[���̏������`�\���܂ł��s��
'
'[pConditionNameList] IN Collection�^: �t�H�[���̃R���{�{�b�N�X�ɓo�^����
'[pConditionName] IN String�^: �t�H�[���̃R���{�{�b�N�X�̏���������
'
'���l:
'

    '��~����Condition����\������
    mTestConditionController.StatusLabel = "Stopping before Setting Conditions [" & pConditionName & "]."
    Call Load(pConditionName)
    mCurConditionName = pConditionName
    
    Call mTestConditionController.Show(vbModeless)
    
    '����ȍ~��TestConditionController�t�H�[���̕\���ʒu���蓮(�O��̕\���ʒu���L������)�ɕύX����B
    mTestConditionController.StartUpPosition = 0
    
    mActive = True
    While mActive = True
        DoEvents
    Wend
    
    Unload mTestConditionController
    
End Sub

Private Sub Load(ByVal pConditionName As String)
'���e:
'   Condition������Condition�����擾���ATreeView���쐬�E�\������
'
'�p�����[�^:
'   [pConditionName] IN String�^: �t�H�[���őI�����ꂽCondition��
'
'���l:
'

    'TreeView�\���p�̃N���[�������擾����B(�擾���͈ȉ��̃����o�ɐݒ肳���)
    ' �EmFunctionNameList
    ' �EmFunctionParamList
    
    RaiseEvent LoadConditionInfo(pConditionName)
    
    mProcedureRowListViewCtrl.ListItems.Clear  'ListView���N���A

    Dim FunctionCount As Long
    For FunctionCount = 1 To mFunctionNameList.Count Step 1
    
        With mProcedureRowListViewCtrl.ListItems.Add
        
            'Procedure����ݒ�
            Dim FunctionName As String
            FunctionName = mFunctionNameList(FunctionCount)
            .Text = FunctionName
            
            'Execute,Arg0~9��ݒ�
            Dim Parameter As Collection
            Set Parameter = mFunctionParamList(FunctionName & FunctionCount)
            
            'Execute
            If Parameter("Exec") = "enable" Then
                .Checked = True
            Else
                .Checked = False
            End If
            
            'Arg0~9
            Dim ParamCount As Long
            For ParamCount = 0 To MAX_ARG_NUMBER Step 1
                
                Dim ArgString As String
                ArgString = ""
                
                Dim SubItemNumber As Long
                SubItemNumber = ParamCount + 1  'SubItems�̊J�nIndex��1�ׁ̈A+1����
                
                If ParamCount < Parameter.Count - 1 Then 'Execute���������p�����^��
                    Dim ArgNumber As String
                    ArgNumber = "Arg" & CStr(ParamCount)
                    ArgString = Parameter(ArgNumber)
                End If
                
                .SubItems(SubItemNumber) = ArgString
                
            Next ParamCount
            
        End With
        
    Next FunctionCount
    
    'List�s�̗L����������
    Dim CurListItem As Long
    For CurListItem = 1 To mProcedureRowListViewCtrl.ListItems.Count Step 1
    
        Call ListItemEnabled(mProcedureRowListViewCtrl.ListItems(CurListItem))
    
    Next CurListItem
    
End Sub

Private Sub Execute()
'���e:
'   �t�H�[����ListView��������擾���A���s�ł���悤�ɐ��`����
'
'���l:
'

    Dim result As Long
    result = MsgBox("Execute condition settings?", vbOKCancel, "TestConditionController")
    If result = vbCancel Then
        mTestConditionController.ExecuteButton.SetFocus
        Exit Sub
    End If
    
    Dim FunctionNameList As Collection
    Dim FunctionParamList As Collection
    
    Set FunctionNameList = New Collection
    Set FunctionParamList = New Collection
    
    'TheCondition�Ŏ��s�ł���悤�ɐ��`����
    Dim CurListItem As Long
    For CurListItem = 1 To mProcedureRowListViewCtrl.ListItems.Count Step 1
    
        With mProcedureRowListViewCtrl.ListItems(CurListItem)
    
            'Function�������X�g�ɒǉ�����
            Dim FunctionName As String
            FunctionName = .Text
            FunctionNameList.Add (FunctionName)
            
            '���݂�Function�p�����[�^���擾����
            Dim ParamList As Collection
            Set ParamList = New Collection
            
            'Function��Execute�̒l��ǉ�����
            If .Checked = True Then
                Call ParamList.Add("enable", "Exec")
            Else
                Call ParamList.Add("disable", "Exec")
            End If
            
            'Arg0~9�̒l��ǉ�����
            Dim ArgCount
            For ArgCount = 0 To MAX_ARG_NUMBER Step 1
        
                Dim SubItemCount As Long
                SubItemCount = ArgCount + 1 'SubItems�̊J�nIndex��1�ׁ̈A+1����
                Dim argParam As String
                argParam = .SubItems(SubItemCount)
                Call ParamList.Add(argParam, "Arg" & ArgCount)
                
            Next ArgCount
            
            Call FunctionParamList.Add(ParamList, FunctionName & CurListItem)
       End With
    Next CurListItem
    
    '�}�N�����s���A�t�H�[���̃I�u�W�F�N�g�𖳌�������
    FormObjectEnabled (False)

    'CDynamicSetCondition�N���X�ɁATestCondition�}�N���̎��s���˗�����
    RaiseEvent ControlExecuteMacro(mCurConditionName, FunctionNameList, FunctionParamList)
    
    '���s��A���ɖ߂�
    FormObjectEnabled (True)
    
    'Execute�����s�����ꍇ�AExit�{�^����L���ɂ���
    If mExecute = False Then
        
        mExitButton.enabled = True
        mTestConditionController.Repaint
        mExecute = True 'Execute�����s�ς݂ɂ���
        
    End If
    
    mTestConditionController.StatusLabel = "Executed Setting Conditions [" & mCurConditionName & "]."
    
    mTestConditionController.ExecuteButton.SetFocus
    
End Sub

Public Property Set SetFunctionNameList(ByVal pFunctionNameList As Collection)
    
    Set mFunctionNameList = pFunctionNameList
    
End Property

Public Property Set SetFunctionParamList(ByVal pFunctionParamList As Collection)

    Set mFunctionParamList = pFunctionParamList
    
End Property

Public Property Get CurExecuteSettingName() As String

    CurExecuteSettingName = mCurExecuteSettingName

End Property

Public Property Let CurExecuteSettingName(ByVal pExecuteSettingName As String)

    mCurExecuteSettingName = pExecuteSettingName

End Property

Private Sub Class_Initialize()
'���e:
'   TestConditionController�t�H�[�����쐬����
'
'���l:
'

    mExecute = False 'Execute���s�ς݃t���O
    Set mArgTextBoxs = Nothing  'Arg�e�L�X�g�{�b�N�X
    Set mProcedureRow = Nothing '�I������Procedure�s�̏��
    Call CreateTestConditionController 'TestConditionController�t�H�[�����쐬����
    
End Sub

Private Sub CreateTestConditionController()
'���e:
'   Utility�N���X�ƃt�H�[�����ӎ������A�t�H�[���̃R���g���[����ݒ�A�擾����
'
'���l:
'

    'TestConditionController�t�H�[�����쐬����
    Set mTestConditionController = TestConditionController
    
    '�t�H�[����ListView��ݒ肷��
    Dim tmpListView As Control
    Set tmpListView = mTestConditionController.Controls.Add("MSComctlLib.ListViewCtrl", "ProcedureListView", True)
    Set mProcedureRowListViewCtrl = tmpListView
    
    With mProcedureRowListViewCtrl
        
        '���X�g�̊�{�ݒ�
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = False
        .FullRowSelect = True
        .Gridlines = True
        .HoverSelection = False
        .CheckBoxes = True
        .Font.Charset = 128
    
        '���X�g�̌��o���̐ݒ�
        .ColumnHeaders.Add , "_Procedure Name", "Execute / Procedure Name", 105
        .ColumnHeaders.Add , "_Arg0", "Arg0", 45
        .ColumnHeaders.Add , "_Arg1", "Arg1", 45
        .ColumnHeaders.Add , "_Arg2", "Arg2", 45
        .ColumnHeaders.Add , "_Arg3", "Arg3", 45
        .ColumnHeaders.Add , "_Arg4", "Arg4", 45
        .ColumnHeaders.Add , "_Arg5", "Arg5", 45
        .ColumnHeaders.Add , "_Arg6", "Arg6", 45
        .ColumnHeaders.Add , "_Arg7", "Arg7", 45
        .ColumnHeaders.Add , "_Arg8", "Arg8", 45
        .ColumnHeaders.Add , "_Arg9", "Arg9", 45
        
    End With
    
    With tmpListView
    
        .Top = 36
        .width = 492
        .Left = 12
        .height = 132
        .TabIndex = 1
        
    End With

    '�t�H�[���̊e�R���g���[���I�u�W�F�N�g���擾
    With mTestConditionController
        Set mReloadButton = .ReloadButton
        Set mExecutebutton = .ExecuteButton
        Set mExitButton = .ExitButton
    End With
    
    '�t�H�[������ʒ����ɔz�u����
    mTestConditionController.StartUpPosition = 2
    
    Set tmpListView = Nothing
    
    'Exit�{�^���𖳌�������(Excute�����s���Ȃ�����A�L���ɂ͂Ȃ�Ȃ�)
    Call FormObjectEnabled(True)
    
End Sub

Private Sub CreateConditionParameterEditor()
'���e:
'   Utility�N���X�ƃt�H�[�����ӎ������A�t�H�[���̃R���g���[����ݒ�A�擾����
'
'���l:
'

    'ConditionParameterEditor�t�H�[�����쐬����
    Set mConditionParameterEditor = ConditionParameterEditor
    
    '�t�H�[���̊e�R���g���[���I�u�W�F�N�g���擾
    With mConditionParameterEditor
        Set mApplyButton = .ApplyButton
        Set mCancelButton = .CancelButton
    End With
End Sub

Private Sub Class_Terminate()
'���e:
'  �t�H�[���̊e�R���g���[�����������
'
'���l:
'
    Set mArgTextBoxs = Nothing
    Set mProcedureRow = Nothing
    
    Set mFunctionNameList = Nothing
    Set mFunctionParamList = Nothing

    'TestConditionController�t�H�[��
    Set mTestConditionController = Nothing
    Set mReloadButton = Nothing
    Set mExecutebutton = Nothing
    Set mExitButton = Nothing
    Set mProcedureRowListViewCtrl = Nothing
    
    'ConditionParameterEditor�t�H�[��
    Set mConditionParameterEditor = Nothing
    Set mApplyButton = Nothing
    Set mCancelButton = Nothing

End Sub

Private Sub mExitButton_Click()
'���e:
'   TestConditionController�����
'   Condition�����s���Ȃ�
'
'���l:
'
    mActive = False

End Sub

Private Sub mApplyButton_Click()
'���e:
'   ListView�ւ̍X�V��v������
'
'���l:
'
    'Arg�̌`���`�F�b�N
    If CheckArgParameter = False Then
        Exit Sub
    End If
    
    Call UpdateProcedure(mProcedureRow)
    mConditionParameterEditor.Hide

End Sub

Private Sub UpdateProcedure(ByVal pProcedureRow As MSComctlLib.ListItem)
'���e:
'   ConditionParameterEidtor�ɕ\������Ă���Condition�p�����[�^��
'   TestConditionController��ListView�ɔ��f����
'
'[pProcedureRow]    IN MSComctlLib.ListItem�^: �ҏW����Procedure�s
'
'���l:
'
    'ListView��Arg���ڂ��X�V����
    Dim tmpArg As CTextBoxCtrl
    For Each tmpArg In mArgTextBoxs
        With pProcedureRow
            If .ListSubItems(tmpArg.id) <> tmpArg.Value Then
                If tmpArg.Value <> "" Then
                    .ListSubItems(tmpArg.id).ForeColor = ALERT_COLOR
                End If
                .ListSubItems(tmpArg.id) = tmpArg.Value
            Else
                .ListSubItems(tmpArg.id).ForeColor = ORIGINAL_COLOR
            End If
        End With
    Next
    Set tmpArg = Nothing
End Sub
Private Function CheckArgParameter() As Boolean
'���e:
'   ConditionParameterEidtor�ɕ\������Ă���Condition�p�����[�^��
'   �`�����`�F�b�N����
'
'�߂�l: Boolean�^
'
'���l:
'   �E#EOP(�I�[������)�����݂��Ȃ��ꍇ�A�G���[�ƂȂ�܂�
'   �E#EOP(�I�[������)���������݂���ꍇ�A�G���[�ƂȂ�܂�
'   �EArg0�`#EOP(�I�[������)�̊Ԃɋ󗓂�����ꍇ�A�G���[�ƂȂ�܂�
'
    
    CheckArgParameter = False
    
    Dim EmptyDataNo As Long
    Dim TermDataNo As Long
    
    '�t���O��Initial�l
    Const INIT_DAT_NO As Long = -1
    
    '�t���O�̏�����
    EmptyDataNo = INIT_DAT_NO
    TermDataNo = INIT_DAT_NO
    
    Dim ExistCount As Long  '�I�[������J�E���g
    ExistCount = 0
    
    Dim ParamCount As Long
    For ParamCount = 0 To mArgTextBoxs.Count - 1 Step 1
    
        Dim tmpArg As CTextBoxCtrl
        Set tmpArg = mArgTextBoxs(CStr(ParamCount))
        
        'Empty��Ԃ̊m�F
        If Len(tmpArg.Value) = 0 Then
            EmptyDataNo = ParamCount
        End If
        '�I�[�����̊m�F
        If StrComp(tmpArg.Value, TERMINATE_CODE) = 0 Then
            TermDataNo = ParamCount
             ExistCount = ExistCount + 1
        End If
        
        '���菈��
        If (EmptyDataNo <> INIT_DAT_NO) And (TermDataNo <> INIT_DAT_NO) Then '�����l��ԂłȂ������m�F
            '�G���[���[�h���菈��
            If EmptyDataNo < TermDataNo Then
                Call MsgBox("Arg" & EmptyDataNo & " is empty.", vbCritical, "ConditionParameterEditor")
                Exit Function
            End If
        End If
        
    Next ParamCount
    
    If (TermDataNo = INIT_DAT_NO) And (EmptyDataNo <> INIT_DAT_NO) Then
        Call MsgBox("'#EOP' is missing.", vbCritical, "ConditionParameterEditor")
        Exit Function
    ElseIf ExistCount > 1 Then
        Call MsgBox("Only one '#EOP' is allowed.", vbCritical, "ConditionParameterEditor")
        Exit Function
    End If
    
    CheckArgParameter = True    '�ُ�Ȃ�
    
End Function
Private Sub mCancelButton_Click()
'���e:
'   ConditionParameterEidtor�����
'   TestConditionController��ListView�ɂ͔��f����Ȃ�
'
'���l:
'

    Call mConditionParameterEditor.Hide

End Sub

Private Sub mExecutebutton_Click()
'���e:
'   ListView�ɕ\������Ă���Condition�p�����[�^�Ń}�N�������s����l�A�˗�����
'
'���l:
'

'TestParameterController�t�H�[����Execute�{�^�����������ꂽ���̏���

    Call Execute
    
End Sub

Private Sub mProcedureRowListViewCtrl_DblClick()
'���e:
'   ���X�g���_�u���N���b�N���ꂽ�ꍇ�̏���
'
'���l:
'

    '�_�u���N���b�N���ꂽ�ꏊ���擾����
    Dim ProcedureRow As MSComctlLib.ListItem
    Set ProcedureRow = mProcedureRowListViewCtrl.SelectedItem
    
    If ProcedureRow.Checked = False Then
        '�I������Procedure�s��Execute��disable�̏ꍇ�A�ҏW�ł��Ȃ�
        Call MsgBox("Cannot edit disabled condition.", vbExclamation, "TestConditionController")
    
    Else
        
        Call CreateConditionParameterEditor   'ConditionParameterEditor�t�H�[�����쐬����
        Call SetConditionParameterEditor(ProcedureRow)
        Call mConditionParameterEditor.Show
    
    End If
    
    ProcedureRow.Selected = False   '�I���s�𖢑I����Ԃɂ���
    
    '�ҏW�I����ATestConditionController�t�H�[�����A�N�e�B�u�ɂ���ׁA��xShow�������B
    Call mTestConditionController.Show(vbModeless)
    
End Sub

Private Sub SetConditionParameterEditor(ByVal pProcedureRow As MSComctlLib.ListItem)
'���e:
'   �I�����ꂽProcedure�s��Arg����ConditionParamterEidtor�ɐݒ肷��
'
'[pProcedureRowRow]    IN MSComctlLib.ListItem�^: �I�����ꂽProcedure�s
'
'���l:
'
    Set mArgTextBoxs = New Collection
    Dim ArgCount As Long
    For ArgCount = 0 To MAX_ARG_NUMBER
        
        Dim ctrlObj As CTextBoxCtrl
        Set ctrlObj = New CTextBoxCtrl
        
        Dim ArgTextBox As MSForms.TextBox
        Set ArgTextBox = mConditionParameterEditor.Controls("Arg" & ArgCount & "Text")
        
        Dim SubItemNumber As Long
        SubItemNumber = ArgCount + 1    'SubItems�̊J�nIndex��1�ׁ̈A+1����
        
        Call ctrlObj.Initialize(ArgTextBox, SubItemNumber)
        Call mArgTextBoxs.Add(ctrlObj, CStr(ArgCount))
        
    Next ArgCount
    
    Set mProcedureRow = pProcedureRow
    Call SetArguments(mProcedureRow)
    
End Sub
Private Sub SetArguments(ByVal pProcedureRowRow As MSComctlLib.ListItem)
'���e:
'   Procedure�s���e�L�X�g�{�b�N�X�ɏ�������
'
'[pProcedureRowRow]    IN MSComctlLib.ListItem�^: �I�����ꂽProcedure�s
'
'���l:
'
    Dim tmpArg As CTextBoxCtrl
    For Each tmpArg In mArgTextBoxs
        tmpArg.Value = pProcedureRowRow.ListSubItems(tmpArg.id)
    Next tmpArg
End Sub
Private Sub mProcedureRowListViewCtrl_ItemCheck(ByVal pProcedureRowRow As MSComctlLib.ListItem)
'���e:
'   ���X�g�̃`�F�b�N�{�b�N�X���ύX���ꂽ�ꍇ�̏���
'
'[pProcedureRowRow]    IN MSComctlLib.ListItem�^: �`�F�b�N�{�b�N�X���ύX���ꂽProcedure�s
'
'���l:
'

    '�`�F�b�N�{�b�N�X���ύX���ꂽProcedure�s��L����or����������
    Call ListItemEnabled(pProcedureRowRow)

End Sub

Private Sub mReloadButton_Click()
'���e:
'   ListView�ɍēxCondition����ǂݍ��ޗl�A�˗�����
'
'���l:
'

'TestParameterController�t�H�[����ReLoad�{�^�����������ꂽ���̏���
    
    'Condition���ǂݍ��ݎ��A�t�H�[���̃I�u�W�F�N�g�𖳌�������
    FormObjectEnabled (False)
    
    Call Load(mCurConditionName)
    
    '�I������猳�ɖ߂�
    FormObjectEnabled (True)
    
    mTestConditionController.StatusLabel = "Stopping before Setting Conditions [" & mCurConditionName & "]."
    
    mTestConditionController.ReloadButton.SetFocus

End Sub

Private Sub FormObjectEnabled(ByVal pEnable As Boolean)
'���e:
'   TestConditionController�I�u�W�F�N�g��Enable�؂�ւ�
'
'[pEnable]   IN Boolean�^: Enable�t���O
'
'���l:
'
    mReloadButton.enabled = pEnable    'Reload�{�^��
    mExecutebutton.enabled = pEnable   'Execute�{�^��
    
    'Exit�{�^��
    If mExecute = True Then
        mExitButton.enabled = pEnable    'Execute�����s�̏ꍇ�́A��ɖ�����
    Else
        mExitButton.enabled = False
    End If
    
    mProcedureRowListViewCtrl.enabled = pEnable   'ListView�E�B���h�E
    
    mTestConditionController.Repaint

End Sub

Private Sub ListItemEnabled(ByVal pProcedureRowRow As MSComctlLib.ListItem)
'���e:
'   enable/diable�ɂ���āA�s�̐F��ύX����
'
'[pProcedureRowRow]      IN MSComctlLib.ListItem�^: �L����or�����������Procedure�s
'
'���l:
'
    '�`�F�b�N�{�b�N�X���ύX���ꂽ�A�C�e�����擾
    Dim ConditionItem As MSComctlLib.ListItem
    Set ConditionItem = pProcedureRowRow
    Dim tmpsub As MSComctlLib.ListSubItem
    
    '�`�F�b�N��ԂŃA�C�e���̃t�H���g�F��ύX
    With ConditionItem
        If .Checked Then
            .ForeColor = ENABLE_COLOR
            For Each tmpsub In .ListSubItems
                tmpsub.ForeColor = ENABLE_COLOR
            Next tmpsub
        Else
            .ForeColor = DISABLE_COLOR
            For Each tmpsub In .ListSubItems
                tmpsub.ForeColor = DISABLE_COLOR
            Next tmpsub
        End If
    End With

End Sub

Private Sub mTestConditionController_QueryClose(ByRef Cancel As Integer, ByVal CloseMode As Integer)
'���e:
'   TestConditionController�t�H�[���́~�{�^�����������ꂽ�ꍇ�̃C�x���g
'   �~�{�^���ł͏I���ł��Ȃ��|��MsgBox�ŕ\������
'
'���l:
'
    If CloseMode = vbFormControlMenu Then
        Call MsgBox("Close button is not available.", , "TestConditionController")
        Cancel = True
    End If

End Sub
