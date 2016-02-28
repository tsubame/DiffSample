VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDynamicSetCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   動的条件設定クラス
'
'目的:
'   TestConditionシートから取得したConditionを、ユーザが一時的に変更し実行するクラス
'
'作成者:
'   0145184935 K.Hara
'

Option Explicit

Implements ISetCondition

Private WithEvents mTestConditionEditor As CTestConditionEditor  'Utilityクラスを意識する。
Attribute mTestConditionEditor.VB_VarHelpID = -1


Private Sub Class_Initialize()

    Set mTestConditionEditor = Nothing

End Sub

Private Sub Class_Terminate()
'
    Set mTestConditionEditor = Nothing
    
End Sub

Public Sub IsetCondition_Execute(ByVal pFunctionInfo As CAcquireFunctionInfo)
'内容:
'   IsetConditionのExecute実行時、TestParameterEditorフォームを表示する
'
'[pFunctionInfo] IN CAcquireFunctionInfo型:  パラメータ値情報を保持する
'
'備考:
'
    'Utilityクラスを作成。
    Set mTestConditionEditor = New CTestConditionEditor
    
    'Condition名のリストを取得する
    Dim ConditionNameList As Collection
    Set ConditionNameList = TheCondition.GetConditionNameList
    
    '現在のFWSetCondition行のCondition名を取得
    Dim ConditionName As String
    ConditionName = pFunctionInfo.Arg(0)
    
    Call mTestConditionEditor.Show(ConditionNameList, ConditionName)
    
End Sub

Private Sub mTestConditionEditor_ControlExecuteMacro(ByVal pCurConditionName As String, ByVal pFunctionNameList As Collection, ByVal pFunctionParamList As Collection)
'内容:
'   TestParameterEditorフォームで編集されたパラメータでマクロを実行する
'
'[pCurConditionName] IN String型:  現在のCondition名
'[pFunctionNameList] IN Collection型:  実行されたConditionが保持する関数の名称リスト
'[pFunctionParamList] IN Collection型:  実行されたConditionが保持する関数の情報リスト
'
'備考:
'
        Dim FunctionNameCount As Long
        
        Dim ExecuteCount As Long
        ExecuteCount = 1
        For FunctionNameCount = 0 To pFunctionNameList.Count - 1 Step 1
        
            Dim SetFunctionInfo As CSetFunctionInfo
            Set SetFunctionInfo = New CSetFunctionInfo
            
            'Function名を取得
            Dim FunctionName As String
            FunctionName = pFunctionNameList(FunctionNameCount + 1)
            
            With SetFunctionInfo
        
                .ConditionName = pCurConditionName
                .FunctionName = FunctionName
                
                Dim ParamList As Collection
                Set ParamList = New Collection
                
                'Functionのパラメータ取得
                Set ParamList = pFunctionParamList.Item(FunctionName & FunctionNameCount + 1)
                
                'Execute設定&実行番号設定
                Dim MacroExec As Boolean
                If ParamList.Item("Exec") = "enable" Then
                    MacroExec = True
                    ExecuteCount = ExecuteCount + 1
                Else
                    MacroExec = False
                End If
                Let .IsEnableMacro = MacroExec
                
                ParamList.Remove ("Exec") 'ParamListにはArgパラメータだけにする
                
                'Arg設定
                Dim ArgCount As Long
                For ArgCount = 0 To ParamList.Count - 1 Step 1
                    If ParamList("Arg" & ArgCount) = "#EOP" Then '終端文字列がきたら終了
                        Exit For
                    End If
                    .AddParameter ParamList("Arg" & ArgCount)
                Next ArgCount
                
                .ExecuteSettingName = mTestConditionEditor.CurExecuteSettingName    'マクログループ名
        
            End With
                                
            '編集したConditionマクロを実行
            Call TheCondition.ExecuteMacro(SetFunctionInfo)
                
        Next FunctionNameCount
        
End Sub

Private Sub mTestConditionEditor_LoadConditionInfo(ByVal pConditionInfoName As String)
'内容:
'   Condition名からCondition情報(コピー)を取得し、TreeView表示用に整形する
'
'[pConditionInfoName] IN String型:  TestParameterEditorフォームで選択されたCondition名
'
'備考:
'
'
    Dim CloneConditionInfoList As Collection
    Set CloneConditionInfoList = TheCondition.GetCloneConditionInfo(pConditionInfoName)
    
    Dim FunctionNameList As Collection
    Dim FunctionParamList As Collection
    
    Set FunctionNameList = New Collection   'Function名リスト
    Set FunctionParamList = New Collection  'Functionパラメタリスト(Exec、Arg0～Arg9)
    
    'Function名のリストとFunction名をKeyとしてArgListを作成する
    Dim InfoCount As Long
    For InfoCount = 1 To CloneConditionInfoList.Count Step 1
    
        Dim FunctionInfo As CSetFunctionInfo
        Set FunctionInfo = CloneConditionInfoList(InfoCount)
        Call FunctionNameList.Add(FunctionInfo.FunctionName)
        
        Dim ParamList As Collection
        Set ParamList = New Collection
        
        'FunctionのExec値を追加する。
        Dim EnableValue As String
        If FunctionInfo.IsEnableMacro = True Then
            EnableValue = "enable"
        Else
            EnableValue = "disable"
        End If
            
        Call ParamList.Add(EnableValue, "Exec")
        
        'Function内のArgを追加する
        Dim ArgCount As Long
        For ArgCount = 0 To FunctionInfo.ArgParameterCount - 1 Step 1
            
            Dim strArg As String
            strArg = CStr(FunctionInfo.Arg(ArgCount))

            Call ParamList.Add(strArg, ("Arg") & CStr(ArgCount))
        
        Next ArgCount
        
        Call ParamList.Add("#EOP", ("Arg") & CStr(ArgCount)) '最後に"#EOP"を追加する
        
        Call FunctionParamList.Add(ParamList, FunctionInfo.FunctionName & InfoCount)
    
    Next InfoCount
    
    'Utilityの情報を一度クリアする
    Set mTestConditionEditor.SetFunctionNameList = Nothing
    Set mTestConditionEditor.SetFunctionParamList = Nothing
    Let mTestConditionEditor.CurExecuteSettingName = ""
    
    'Utilityへ情報をSetする
    Set mTestConditionEditor.SetFunctionNameList = FunctionNameList
    Set mTestConditionEditor.SetFunctionParamList = FunctionParamList
    Dim TopFunctionInfo As CSetFunctionInfo
    Set TopFunctionInfo = CloneConditionInfoList(1)    'Condition単位でならマクログループ名は共通のはず。先頭のマクログループを取得。
    Let mTestConditionEditor.CurExecuteSettingName = TopFunctionInfo.ExecuteSettingName
    
    'Cloneは最後に解放する
    Set CloneConditionInfoList = Nothing
    
End Sub

