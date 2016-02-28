VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcScenarioLoopOptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   DcSccenarioLoopOptionForm管理クラス
'
'目的:
'   DcSccenarioLoopOptionFormについて
'   ・表示
'   ・情報取得
'   ・ボタンイベント管理
'   を行う
'
'作成者:
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
'内容:
'   Formの表示
'
'パラメータ:
'
'戻り値:
'   どのボタンを押して終了したか    Boolean型
'       [Run]ボタン:True
'       その他     :False
'備考:
'   Initializeで表示するカテゴリーを設定していない場合、表示せずに終了
'
    
    Dim res As Boolean
    res = False
    
    If mCategories Is Nothing Then
        res = False
    Else
        '左のリストに全アイテム表示
        Dim category As Variant
        For Each category In mCategories
            mCategoriesList.AddItem category
        Next category
        
        '繰り返し条件初期化
        mParamText.Value = ""
        mCountsOptionButton.Value = True
        mUnitLabel = COUNTS_UNIT
        
        'フォルダパス初期値設定(JOBファイルパス)
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
'内容:
'   設定内容の取得
'
'パラメータ:
'
'戻り値:
'   設定内容    LOOP_CONDITION型
'
'備考:
'
'
    Dim loopCategories As Collection
    Dim samplingParam As SAMPLING_PARAM
    Dim paramLimit As Double
    Dim folderPath As String
    
    'Loop対象カテゴリー
    Set loopCategories = New Collection
    Dim index As Integer
    For index = 0 To mLoopCategoriesList.ListCount - 1
        loopCategories.Add mLoopCategoriesList.List(index), mLoopCategoriesList.List(index)
    Next index
    
    'Counts/Time選択情報
    If mCountsOptionButton.Value = True Then
        samplingParam = PARAM_COUNTS
    Else
        samplingParam = PARAM_TIME
    End If
    
    'サンプリングリミット
    On Error GoTo CAST_ERROR
    If mCountsOptionButton.Value = True Then
        Dim cast As Long
        cast = CLng(mParamText.Value)
        paramLimit = CDbl(cast)
    Else
        paramLimit = CDbl(mParamText.Value)
    End If
    
    'フォルダパス
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
'[>]ボタン押下時
    Call moveListItems(mCategoriesList, mLoopCategoriesList)
End Sub

Private Sub mMoveRightAllButton_Click()
'[>>]ボタン押下時
    Call moveAllListItems(mCategoriesList, mLoopCategoriesList)
End Sub

Private Sub mMoveLeftButton_Click()
'[<]ボタン押下時
    Call moveListItems(mLoopCategoriesList, mCategoriesList)
End Sub

Private Sub mMoveLeftAllButton_Click()
'[<<]ボタン押下時
    Call moveAllListItems(mLoopCategoriesList, mCategoriesList)
End Sub

Private Sub mRunButton_Click()
'内容:
'   [Run]ボタン押下時
'
'パラメータ:
'
'戻り値:
'
'備考:
'   フォルダアクセスチェックとサンプリングパラメータのチェックを行う
'
    If checkLimit = True Then
        If checkFolderAccess = True Then
            mDcLoop = True
            mLoopOptionForm.Hide
        End If
    End If
End Sub

Private Sub mCancelButton_Click()
'[Cancel]ボタン押下時
    MsgBox "Canceled loop option." & vbCrLf & "Run on the normal mode.", vbInformation, FORM_CAPTION
    
    mDcLoop = False
    mLoopOptionForm.Hide
End Sub
Private Sub mLoopOptionForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'[X]ボタン押下時
    If CloseMode = 0 Then
        mCancelButton_Click
    End If
End Sub


Private Sub mCountsOptionButton_Click()
'[Counts]オプション選択時
    mUnitLabel = COUNTS_UNIT
End Sub

Private Sub mTimeOptionButton_Click()
'[Time]オプション選択時
    mUnitLabel = TIME_UNIT
End Sub

Private Sub mFolderReferenceButton_Click()
'内容:
'   Folder参照ボタン
'
'パラメータ:
'
'戻り値:
'
'備考:
'
'
'フォルダ選択ボタン押下時
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            mFolderPathText.Value = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub Class_Initialize()
'内容:
'   クラス初期化
'
'備考:
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
'内容:
'   クラス破棄
'
'備考:
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
'内容:
'   ListBox間の全アイテム移動処理
'
'パラメータ:
'[pFrom]       IN/OUT MSForms.ListBox型:  移動元ListBox
'[pTo]         IN/OUT MSForms.ListBox型:  移動先ListBox
'
'戻り値:
'
'備考:
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
'内容:
'   ListBox間の選択アイテム移動処理
'
'パラメータ:
'[pFrom]       IN/OUT MSForms.ListBox型:  移動元ListBox
'[pTo]         IN/OUT MSForms.ListBox型:  移動先ListBox
'
'戻り値:
'
'備考:
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
'内容:
'   テキストボックスに指定されたフォルダのアクセス確認
'   アクセス権確認の為、実際にファイル書き込みを行う。
'
'戻り値:
'   フォルダアクセス可否    Boolean型
'備考:
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
    
    'ドライブエラー回避用
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
'内容:
'   テキストボックスに指定されたパラメータのチェック
'   型チェックと上限/下限値チェックを行う
'
'戻り値:
'   チェック結果    Boolean型
'備考:
'
'
    Dim res As Boolean
    res = True
    
    Dim doOverLimit As VbMsgBoxResult
    
    'CountsOptionの場合
    If mCountsOptionButton.Value = True Then
        Dim checkCountVal As Long
        On Error GoTo OVERFLOW
        checkCountVal = CLng(mParamText.Value)
        'OverFlow規定値以上
        If checkCountVal >= COUNTS_OF_LIMIT Then
            MsgBox mParamText.Value & " is invalid parameter.", vbCritical, FORM_CAPTION
            res = False
        '下限値以下
        ElseIf checkCountVal < COUNTS_LOWER_LIMIT Then
            MsgBox mParamText.Value & " is invalid parameter.", vbCritical, FORM_CAPTION
            res = False
        '上限値以上
        ElseIf checkCountVal > COUNTS_UPPER_LIMIT Then
            doOverLimit = MsgBox(checkCountVal & " has exceeded the upper limit " & COUNTS_UPPER_LIMIT & "." & vbCrLf & _
                            "'OK' to run at this parameter." _
                                , vbOKCancel + vbExclamation, FORM_CAPTION)
            If doOverLimit <> vbOK Then res = False
        End If
    'TimeOptionの場合
    Else
        Dim checkTimeVal As Double
        On Error GoTo OVERFLOW
        checkTimeVal = CDbl(mParamText.Value)
        'OverFlow規定値以上
        If checkTimeVal >= TIME_OF_LIMIT Then
            MsgBox mParamText.Value & " is invalid parameter.", vbCritical, FORM_CAPTION
            res = False
        '下限値以下
        ElseIf checkTimeVal < TIME_LOWER_LIMIT Then
            MsgBox mParamText.Value & " is invalid parameter.", vbCritical, FORM_CAPTION
            res = False
        '上限値以上
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
'内容:
'   引数のIndexリストをソートして返す
'
'[pIndexList]  IN/OUT Collection型:  Indexリスト
'
'備考:
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
'内容:
'   データのソートを行う
'
'[pData]       IN/OUT Variant型:    ソートするデータ
'[pLower]      IN     Long型:       最小値
'[pUpper]      IN     Long型:       最大値
'
'備考:
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
