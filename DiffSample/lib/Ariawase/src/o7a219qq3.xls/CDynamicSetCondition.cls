VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDynamicSetCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ���I�����ݒ�N���X
'
'�ړI:
'   TestCondition�V�[�g����擾����Condition���A���[�U���ꎞ�I�ɕύX�����s����N���X
'
'�쐬��:
'   0145184935 K.Hara
'

Option Explicit

Implements ISetCondition

Private WithEvents mTestConditionEditor As CTestConditionEditor  'Utility�N���X���ӎ�����B
Attribute mTestConditionEditor.VB_VarHelpID = -1


Private Sub Class_Initialize()

    Set mTestConditionEditor = Nothing

End Sub

Private Sub Class_Terminate()
'
    Set mTestConditionEditor = Nothing
    
End Sub

Public Sub IsetCondition_Execute(ByVal pFunctionInfo As CAcquireFunctionInfo)
'���e:
'   IsetCondition��Execute���s���ATestParameterEditor�t�H�[����\������
'
'[pFunctionInfo] IN CAcquireFunctionInfo�^:  �p�����[�^�l����ێ�����
'
'���l:
'
    'Utility�N���X���쐬�B
    Set mTestConditionEditor = New CTestConditionEditor
    
    'Condition���̃��X�g���擾����
    Dim ConditionNameList As Collection
    Set ConditionNameList = TheCondition.GetConditionNameList
    
    '���݂�FWSetCondition�s��Condition�����擾
    Dim ConditionName As String
    ConditionName = pFunctionInfo.Arg(0)
    
    Call mTestConditionEditor.Show(ConditionNameList, ConditionName)
    
End Sub

Private Sub mTestConditionEditor_ControlExecuteMacro(ByVal pCurConditionName As String, ByVal pFunctionNameList As Collection, ByVal pFunctionParamList As Collection)
'���e:
'   TestParameterEditor�t�H�[���ŕҏW���ꂽ�p�����[�^�Ń}�N�������s����
'
'[pCurConditionName] IN String�^:  ���݂�Condition��
'[pFunctionNameList] IN Collection�^:  ���s���ꂽCondition���ێ�����֐��̖��̃��X�g
'[pFunctionParamList] IN Collection�^:  ���s���ꂽCondition���ێ�����֐��̏�񃊃X�g
'
'���l:
'
        Dim FunctionNameCount As Long
        
        Dim ExecuteCount As Long
        ExecuteCount = 1
        For FunctionNameCount = 0 To pFunctionNameList.Count - 1 Step 1
        
            Dim SetFunctionInfo As CSetFunctionInfo
            Set SetFunctionInfo = New CSetFunctionInfo
            
            'Function�����擾
            Dim FunctionName As String
            FunctionName = pFunctionNameList(FunctionNameCount + 1)
            
            With SetFunctionInfo
        
                .ConditionName = pCurConditionName
                .FunctionName = FunctionName
                
                Dim ParamList As Collection
                Set ParamList = New Collection
                
                'Function�̃p�����[�^�擾
                Set ParamList = pFunctionParamList.Item(FunctionName & FunctionNameCount + 1)
                
                'Execute�ݒ�&���s�ԍ��ݒ�
                Dim MacroExec As Boolean
                If ParamList.Item("Exec") = "enable" Then
                    MacroExec = True
                    ExecuteCount = ExecuteCount + 1
                Else
                    MacroExec = False
                End If
                Let .IsEnableMacro = MacroExec
                
                ParamList.Remove ("Exec") 'ParamList�ɂ�Arg�p�����[�^�����ɂ���
                
                'Arg�ݒ�
                Dim ArgCount As Long
                For ArgCount = 0 To ParamList.Count - 1 Step 1
                    If ParamList("Arg" & ArgCount) = "#EOP" Then '�I�[�����񂪂�����I��
                        Exit For
                    End If
                    .AddParameter ParamList("Arg" & ArgCount)
                Next ArgCount
                
                .ExecuteSettingName = mTestConditionEditor.CurExecuteSettingName    '�}�N���O���[�v��
        
            End With
                                
            '�ҏW����Condition�}�N�������s
            Call TheCondition.ExecuteMacro(SetFunctionInfo)
                
        Next FunctionNameCount
        
End Sub

Private Sub mTestConditionEditor_LoadConditionInfo(ByVal pConditionInfoName As String)
'���e:
'   Condition������Condition���(�R�s�[)���擾���ATreeView�\���p�ɐ��`����
'
'[pConditionInfoName] IN String�^:  TestParameterEditor�t�H�[���őI�����ꂽCondition��
'
'���l:
'
'
    Dim CloneConditionInfoList As Collection
    Set CloneConditionInfoList = TheCondition.GetCloneConditionInfo(pConditionInfoName)
    
    Dim FunctionNameList As Collection
    Dim FunctionParamList As Collection
    
    Set FunctionNameList = New Collection   'Function�����X�g
    Set FunctionParamList = New Collection  'Function�p�����^���X�g(Exec�AArg0�`Arg9)
    
    'Function���̃��X�g��Function����Key�Ƃ���ArgList���쐬����
    Dim InfoCount As Long
    For InfoCount = 1 To CloneConditionInfoList.Count Step 1
    
        Dim FunctionInfo As CSetFunctionInfo
        Set FunctionInfo = CloneConditionInfoList(InfoCount)
        Call FunctionNameList.Add(FunctionInfo.FunctionName)
        
        Dim ParamList As Collection
        Set ParamList = New Collection
        
        'Function��Exec�l��ǉ�����B
        Dim EnableValue As String
        If FunctionInfo.IsEnableMacro = True Then
            EnableValue = "enable"
        Else
            EnableValue = "disable"
        End If
            
        Call ParamList.Add(EnableValue, "Exec")
        
        'Function����Arg��ǉ�����
        Dim ArgCount As Long
        For ArgCount = 0 To FunctionInfo.ArgParameterCount - 1 Step 1
            
            Dim strArg As String
            strArg = CStr(FunctionInfo.Arg(ArgCount))

            Call ParamList.Add(strArg, ("Arg") & CStr(ArgCount))
        
        Next ArgCount
        
        Call ParamList.Add("#EOP", ("Arg") & CStr(ArgCount)) '�Ō��"#EOP"��ǉ�����
        
        Call FunctionParamList.Add(ParamList, FunctionInfo.FunctionName & InfoCount)
    
    Next InfoCount
    
    'Utility�̏�����x�N���A����
    Set mTestConditionEditor.SetFunctionNameList = Nothing
    Set mTestConditionEditor.SetFunctionParamList = Nothing
    Let mTestConditionEditor.CurExecuteSettingName = ""
    
    'Utility�֏���Set����
    Set mTestConditionEditor.SetFunctionNameList = FunctionNameList
    Set mTestConditionEditor.SetFunctionParamList = FunctionParamList
    Dim TopFunctionInfo As CSetFunctionInfo
    Set TopFunctionInfo = CloneConditionInfoList(1)    'Condition�P�ʂłȂ�}�N���O���[�v���͋��ʂ̂͂��B�擪�̃}�N���O���[�v���擾�B
    Let mTestConditionEditor.CurExecuteSettingName = TopFunctionInfo.ExecuteSettingName
    
    'Clone�͍Ō�ɉ������
    Set CloneConditionInfoList = Nothing
    
End Sub
