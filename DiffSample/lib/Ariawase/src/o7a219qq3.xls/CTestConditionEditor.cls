VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestConditionEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   Formのコンポーネントの機能実現用クラス
'
'目的:
'   動的条件設定クラスと条件設定編集Formとの情報をやり取りする
'
'作成者:
'   0145184935 K.Hara
'
Option Explicit

Private mActive As Boolean  'TestConditionEditorフォームのActiveフラグ
Private mExecute As Boolean 'Execute実行済みフラグ

Private mArgTextBoxs As Collection              'Argテキストボックス
Private mProcedureRow As MSComctlLib.ListItem   '選択したProcedure行の情報

Private mCurConditionName As String          '現在のCondition名
Private mFunctionNameList As Collection      '関数名リスト
Private mFunctionParamList As Collection     'パラメータリスト(Keyは関数名）
Private mCurExecuteSettingName As String     '現在のマクログループ
Private Const MAX_ARG_NUMBER As Long = 9     'Argパラメータの個数

'TestConditionControlerフォームの各コントローラオブジェクト
Private WithEvents mTestConditionController As TestConditionController 'TestConditionControllerフォームを認知する為のメンバ
Attribute mTestConditionController.VB_VarHelpID = -1
Private WithEvents mReloadButton As MSForms.CommandButton           'Reloadボタン
Attribute mReloadButton.VB_VarHelpID = -1
Private WithEvents mExecutebutton As MSForms.CommandButton          'Executeボタン
Attribute mExecutebutton.VB_VarHelpID = -1
Private WithEvents mExitButton As MSForms.CommandButton            'Exitボタン
Attribute mExitButton.VB_VarHelpID = -1
Private WithEvents mProcedureRowListViewCtrl As MSComctlLib.ListView   'ListViewウィンドウ
Attribute mProcedureRowListViewCtrl.VB_VarHelpID = -1

Private Const ENABLE_COLOR As Long = &H80000008 'ListView行の有効色
Private Const DISABLE_COLOR As Long = &H808080  'ListView行の無効色

'ConditionParameterEditorフォームの格コントローラオブジェクト
Private mConditionParameterEditor As ConditionParameterEditor 'ConditionParamterEditフォームを認知する為のメンバ
Private WithEvents mApplyButton As MSForms.CommandButton    'Applyボタン
Attribute mApplyButton.VB_VarHelpID = -1
Private WithEvents mCancelButton As MSForms.CommandButton   'Cancelボタン
Attribute mCancelButton.VB_VarHelpID = -1

Private Const TERMINATE_CODE As String = "#EOP"     '終端文字列
Private Const ALERT_COLOR As Long = &HFFF           '反映後の文字色
Private Const ORIGINAL_COLOR As Long = &H80000008   '反映前の文字色

'このクラスで発生させるイベント。CDynamicSetConditionクラスがキャッチする。
Public Event LoadConditionInfo(ByVal pConditionInfoName As String)
Public Event ControlExecuteMacro(ByVal pCurConditionName As String, ByVal pFunctionNameList As Collection, ByVal pFunctionParamList As Collection)

Public Sub Show(ByVal pConditionNameList As Collection, ByVal pConditionName As String)
'内容:
'   フォームの初期化～表示までを行う
'
'[pConditionNameList] IN Collection型: フォームのコンボボックスに登録する
'[pConditionName] IN String型: フォームのコンボボックスの初期文字列
'
'備考:
'

    '停止時のCondition情報を表示する
    mTestConditionController.StatusLabel = "Stopping before Setting Conditions [" & pConditionName & "]."
    Call Load(pConditionName)
    mCurConditionName = pConditionName
    
    Call mTestConditionController.Show(vbModeless)
    
    'これ以降のTestConditionControllerフォームの表示位置を手動(前回の表示位置を記憶する)に変更する。
    mTestConditionController.StartUpPosition = 0
    
    mActive = True
    While mActive = True
        DoEvents
    Wend
    
    Unload mTestConditionController
    
End Sub

Private Sub Load(ByVal pConditionName As String)
'内容:
'   Condition名からCondition情報を取得し、TreeViewを作成・表示する
'
'パラメータ:
'   [pConditionName] IN String型: フォームで選択されたCondition名
'
'備考:
'

    'TreeView表示用のクローン情報を取得する。(取得情報は以下のメンバに設定される)
    ' ・mFunctionNameList
    ' ・mFunctionParamList
    
    RaiseEvent LoadConditionInfo(pConditionName)
    
    mProcedureRowListViewCtrl.ListItems.Clear  'ListViewをクリア

    Dim FunctionCount As Long
    For FunctionCount = 1 To mFunctionNameList.Count Step 1
    
        With mProcedureRowListViewCtrl.ListItems.Add
        
            'Procedure名を設定
            Dim FunctionName As String
            FunctionName = mFunctionNameList(FunctionCount)
            .Text = FunctionName
            
            'Execute,Arg0~9を設定
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
                SubItemNumber = ParamCount + 1  'SubItemsの開始Indexが1の為、+1する
                
                If ParamCount < Parameter.Count - 1 Then 'Executeを除いたパラメタ数
                    Dim ArgNumber As String
                    ArgNumber = "Arg" & CStr(ParamCount)
                    ArgString = Parameter(ArgNumber)
                End If
                
                .SubItems(SubItemNumber) = ArgString
                
            Next ParamCount
            
        End With
        
    Next FunctionCount
    
    'List行の有効無効処理
    Dim CurListItem As Long
    For CurListItem = 1 To mProcedureRowListViewCtrl.ListItems.Count Step 1
    
        Call ListItemEnabled(mProcedureRowListViewCtrl.ListItems(CurListItem))
    
    Next CurListItem
    
End Sub

Private Sub Execute()
'内容:
'   フォームのListViewから情報を取得し、実行できるように整形する
'
'備考:
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
    
    'TheConditionで実行できるように整形する
    Dim CurListItem As Long
    For CurListItem = 1 To mProcedureRowListViewCtrl.ListItems.Count Step 1
    
        With mProcedureRowListViewCtrl.ListItems(CurListItem)
    
            'Function名をリストに追加する
            Dim FunctionName As String
            FunctionName = .Text
            FunctionNameList.Add (FunctionName)
            
            '現在のFunctionパラメータを取得する
            Dim ParamList As Collection
            Set ParamList = New Collection
            
            'FunctionのExecuteの値を追加する
            If .Checked = True Then
                Call ParamList.Add("enable", "Exec")
            Else
                Call ParamList.Add("disable", "Exec")
            End If
            
            'Arg0~9の値を追加する
            Dim ArgCount
            For ArgCount = 0 To MAX_ARG_NUMBER Step 1
        
                Dim SubItemCount As Long
                SubItemCount = ArgCount + 1 'SubItemsの開始Indexが1の為、+1する
                Dim argParam As String
                argParam = .SubItems(SubItemCount)
                Call ParamList.Add(argParam, "Arg" & ArgCount)
                
            Next ArgCount
            
            Call FunctionParamList.Add(ParamList, FunctionName & CurListItem)
       End With
    Next CurListItem
    
    'マクロ実行時、フォームのオブジェクトを無効化する
    FormObjectEnabled (False)

    'CDynamicSetConditionクラスに、TestConditionマクロの実行を依頼する
    RaiseEvent ControlExecuteMacro(mCurConditionName, FunctionNameList, FunctionParamList)
    
    '実行後、元に戻す
    FormObjectEnabled (True)
    
    'Executeを実行した場合、Exitボタンを有効にする
    If mExecute = False Then
        
        mExitButton.enabled = True
        mTestConditionController.Repaint
        mExecute = True 'Executeを実行済みにする
        
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
'内容:
'   TestConditionControllerフォームを作成する
'
'備考:
'

    mExecute = False 'Execute実行済みフラグ
    Set mArgTextBoxs = Nothing  'Argテキストボックス
    Set mProcedureRow = Nothing '選択したProcedure行の情報
    Call CreateTestConditionController 'TestConditionControllerフォームを作成する
    
End Sub

Private Sub CreateTestConditionController()
'内容:
'   Utilityクラスとフォームを意識させ、フォームのコントローラを設定、取得する
'
'備考:
'

    'TestConditionControllerフォームを作成する
    Set mTestConditionController = TestConditionController
    
    'フォームにListViewを設定する
    Dim tmpListView As Control
    Set tmpListView = mTestConditionController.Controls.Add("MSComctlLib.ListViewCtrl", "ProcedureListView", True)
    Set mProcedureRowListViewCtrl = tmpListView
    
    With mProcedureRowListViewCtrl
        
        'リストの基本設定
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = False
        .FullRowSelect = True
        .Gridlines = True
        .HoverSelection = False
        .CheckBoxes = True
        .Font.Charset = 128
    
        'リストの見出しの設定
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

    'フォームの各コントローラオブジェクトを取得
    With mTestConditionController
        Set mReloadButton = .ReloadButton
        Set mExecutebutton = .ExecuteButton
        Set mExitButton = .ExitButton
    End With
    
    'フォームを画面中央に配置する
    mTestConditionController.StartUpPosition = 2
    
    Set tmpListView = Nothing
    
    'Exitボタンを無効化する(Excuteを実行しない限り、有効にはならない)
    Call FormObjectEnabled(True)
    
End Sub

Private Sub CreateConditionParameterEditor()
'内容:
'   Utilityクラスとフォームを意識させ、フォームのコントローラを設定、取得する
'
'備考:
'

    'ConditionParameterEditorフォームを作成する
    Set mConditionParameterEditor = ConditionParameterEditor
    
    'フォームの各コントローラオブジェクトを取得
    With mConditionParameterEditor
        Set mApplyButton = .ApplyButton
        Set mCancelButton = .CancelButton
    End With
End Sub

Private Sub Class_Terminate()
'内容:
'  フォームの各コントローラを解放する
'
'備考:
'
    Set mArgTextBoxs = Nothing
    Set mProcedureRow = Nothing
    
    Set mFunctionNameList = Nothing
    Set mFunctionParamList = Nothing

    'TestConditionControllerフォーム
    Set mTestConditionController = Nothing
    Set mReloadButton = Nothing
    Set mExecutebutton = Nothing
    Set mExitButton = Nothing
    Set mProcedureRowListViewCtrl = Nothing
    
    'ConditionParameterEditorフォーム
    Set mConditionParameterEditor = Nothing
    Set mApplyButton = Nothing
    Set mCancelButton = Nothing

End Sub

Private Sub mExitButton_Click()
'内容:
'   TestConditionControllerを閉じる
'   Conditionを実行しない
'
'備考:
'
    mActive = False

End Sub

Private Sub mApplyButton_Click()
'内容:
'   ListViewへの更新を要請する
'
'備考:
'
    'Argの形式チェック
    If CheckArgParameter = False Then
        Exit Sub
    End If
    
    Call UpdateProcedure(mProcedureRow)
    mConditionParameterEditor.Hide

End Sub

Private Sub UpdateProcedure(ByVal pProcedureRow As MSComctlLib.ListItem)
'内容:
'   ConditionParameterEidtorに表示されているConditionパラメータを
'   TestConditionControllerのListViewに反映する
'
'[pProcedureRow]    IN MSComctlLib.ListItem型: 編集したProcedure行
'
'備考:
'
    'ListViewのArg項目を更新する
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
'内容:
'   ConditionParameterEidtorに表示されているConditionパラメータの
'   形式をチェックする
'
'戻り値: Boolean型
'
'備考:
'   ・#EOP(終端文字列)が存在しない場合、エラーとなります
'   ・#EOP(終端文字列)が複数存在する場合、エラーとなります
'   ・Arg0～#EOP(終端文字列)の間に空欄がある場合、エラーとなります
'
    
    CheckArgParameter = False
    
    Dim EmptyDataNo As Long
    Dim TermDataNo As Long
    
    'フラグのInitial値
    Const INIT_DAT_NO As Long = -1
    
    'フラグの初期化
    EmptyDataNo = INIT_DAT_NO
    TermDataNo = INIT_DAT_NO
    
    Dim ExistCount As Long  '終端文字列カウント
    ExistCount = 0
    
    Dim ParamCount As Long
    For ParamCount = 0 To mArgTextBoxs.Count - 1 Step 1
    
        Dim tmpArg As CTextBoxCtrl
        Set tmpArg = mArgTextBoxs(CStr(ParamCount))
        
        'Empty状態の確認
        If Len(tmpArg.Value) = 0 Then
            EmptyDataNo = ParamCount
        End If
        '終端文字の確認
        If StrComp(tmpArg.Value, TERMINATE_CODE) = 0 Then
            TermDataNo = ParamCount
             ExistCount = ExistCount + 1
        End If
        
        '判定処理
        If (EmptyDataNo <> INIT_DAT_NO) And (TermDataNo <> INIT_DAT_NO) Then '初期値状態でない事を確認
            'エラーモード判定処理
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
    
    CheckArgParameter = True    '異常なし
    
End Function
Private Sub mCancelButton_Click()
'内容:
'   ConditionParameterEidtorを閉じる
'   TestConditionControllerのListViewには反映されない
'
'備考:
'

    Call mConditionParameterEditor.Hide

End Sub

Private Sub mExecutebutton_Click()
'内容:
'   ListViewに表示されているConditionパラメータでマクロを実行する様、依頼する
'
'備考:
'

'TestParameterControllerフォームでExecuteボタンが押下された時の処理

    Call Execute
    
End Sub

Private Sub mProcedureRowListViewCtrl_DblClick()
'内容:
'   リストがダブルクリックされた場合の処理
'
'備考:
'

    'ダブルクリックされた場所を取得する
    Dim ProcedureRow As MSComctlLib.ListItem
    Set ProcedureRow = mProcedureRowListViewCtrl.SelectedItem
    
    If ProcedureRow.Checked = False Then
        '選択したProcedure行のExecuteがdisableの場合、編集できない
        Call MsgBox("Cannot edit disabled condition.", vbExclamation, "TestConditionController")
    
    Else
        
        Call CreateConditionParameterEditor   'ConditionParameterEditorフォームを作成する
        Call SetConditionParameterEditor(ProcedureRow)
        Call mConditionParameterEditor.Show
    
    End If
    
    ProcedureRow.Selected = False   '選択行を未選択状態にする
    
    '編集終了後、TestConditionControllerフォームをアクティブにする為、一度Showし直す。
    Call mTestConditionController.Show(vbModeless)
    
End Sub

Private Sub SetConditionParameterEditor(ByVal pProcedureRow As MSComctlLib.ListItem)
'内容:
'   選択されたProcedure行のArg情報をConditionParamterEidtorに設定する
'
'[pProcedureRowRow]    IN MSComctlLib.ListItem型: 選択されたProcedure行
'
'備考:
'
    Set mArgTextBoxs = New Collection
    Dim ArgCount As Long
    For ArgCount = 0 To MAX_ARG_NUMBER
        
        Dim ctrlObj As CTextBoxCtrl
        Set ctrlObj = New CTextBoxCtrl
        
        Dim ArgTextBox As MSForms.TextBox
        Set ArgTextBox = mConditionParameterEditor.Controls("Arg" & ArgCount & "Text")
        
        Dim SubItemNumber As Long
        SubItemNumber = ArgCount + 1    'SubItemsの開始Indexが1の為、+1する
        
        Call ctrlObj.Initialize(ArgTextBox, SubItemNumber)
        Call mArgTextBoxs.Add(ctrlObj, CStr(ArgCount))
        
    Next ArgCount
    
    Set mProcedureRow = pProcedureRow
    Call SetArguments(mProcedureRow)
    
End Sub
Private Sub SetArguments(ByVal pProcedureRowRow As MSComctlLib.ListItem)
'内容:
'   Procedure行をテキストボックスに書き込む
'
'[pProcedureRowRow]    IN MSComctlLib.ListItem型: 選択されたProcedure行
'
'備考:
'
    Dim tmpArg As CTextBoxCtrl
    For Each tmpArg In mArgTextBoxs
        tmpArg.Value = pProcedureRowRow.ListSubItems(tmpArg.id)
    Next tmpArg
End Sub
Private Sub mProcedureRowListViewCtrl_ItemCheck(ByVal pProcedureRowRow As MSComctlLib.ListItem)
'内容:
'   リストのチェックボックスが変更された場合の処理
'
'[pProcedureRowRow]    IN MSComctlLib.ListItem型: チェックボックスが変更されたProcedure行
'
'備考:
'

    'チェックボックスが変更されたProcedure行を有効化or無効化する
    Call ListItemEnabled(pProcedureRowRow)

End Sub

Private Sub mReloadButton_Click()
'内容:
'   ListViewに再度Condition情報を読み込む様、依頼する
'
'備考:
'

'TestParameterControllerフォームでReLoadボタンが押下された時の処理
    
    'Condition情報読み込み時、フォームのオブジェクトを無効化する
    FormObjectEnabled (False)
    
    Call Load(mCurConditionName)
    
    '終わったら元に戻す
    FormObjectEnabled (True)
    
    mTestConditionController.StatusLabel = "Stopping before Setting Conditions [" & mCurConditionName & "]."
    
    mTestConditionController.ReloadButton.SetFocus

End Sub

Private Sub FormObjectEnabled(ByVal pEnable As Boolean)
'内容:
'   TestConditionControllerオブジェクトのEnable切り替え
'
'[pEnable]   IN Boolean型: Enableフラグ
'
'備考:
'
    mReloadButton.enabled = pEnable    'Reloadボタン
    mExecutebutton.enabled = pEnable   'Executeボタン
    
    'Exitボタン
    If mExecute = True Then
        mExitButton.enabled = pEnable    'Execute未実行の場合は、常に無効化
    Else
        mExitButton.enabled = False
    End If
    
    mProcedureRowListViewCtrl.enabled = pEnable   'ListViewウィンドウ
    
    mTestConditionController.Repaint

End Sub

Private Sub ListItemEnabled(ByVal pProcedureRowRow As MSComctlLib.ListItem)
'内容:
'   enable/diableによって、行の色を変更する
'
'[pProcedureRowRow]      IN MSComctlLib.ListItem型: 有効化or無効化されるProcedure行
'
'備考:
'
    'チェックボックスが変更されたアイテムを取得
    Dim ConditionItem As MSComctlLib.ListItem
    Set ConditionItem = pProcedureRowRow
    Dim tmpsub As MSComctlLib.ListSubItem
    
    'チェック状態でアイテムのフォント色を変更
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
'内容:
'   TestConditionControllerフォームの×ボタンが押下された場合のイベント
'   ×ボタンでは終了できない旨をMsgBoxで表示する
'
'備考:
'
    If CloseMode = vbFormControlMenu Then
        Call MsgBox("Close button is not available.", , "TestConditionController")
        Cancel = True
    End If

End Sub

