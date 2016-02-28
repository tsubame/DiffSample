VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDcCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DCテストカテゴリクラス
'
'目的:
'   カテゴリごとにアクションのコレクションを所有する
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWritable

Dim mActionItems As Collection
Dim mCategoryID As CParamLong
Dim mCategoryName As CParamString
Dim mIsExamin As CParamBoolean
Dim mExaminMode As CParamString

Private Sub Class_Initialize()
    Set mActionItems = New Collection
    Set mCategoryID = CreateCParamLong
    mCategoryID.AsIParameter.Name = CATEGORY_ID
    Set mCategoryName = CreateCParamString
    mCategoryName.AsIParameter.Name = TEST_CATEGORY
    Set mIsExamin = CreateCParamBoolean
    mIsExamin.AsIParameter.Name = EXAMIN_FLAG
    Set mExaminMode = CreateCParamString
    mExaminMode.AsIParameter.Name = EXAMIN_MODE
End Sub

Public Sub Initialize(ByVal actions As Collection, ByVal reader As IParameterReader)
    Set mActionItems = actions
    mCategoryID.AsIParameter.Read reader
    mCategoryName.AsIParameter.Read reader
    mIsExamin.AsIParameter.Read reader
    mExaminMode.AsIParameter.Read reader
End Sub

Public Property Get Groups() As Collection
    Set Groups = mActionItems
End Property

Public Property Let Groups(ByVal actionItems As Collection)
    Set mActionItems = actionItems
End Property

Public Property Get Name() As String
    Name = mCategoryName.AsIParameter.AsString
End Property

Public Property Let Name(ByVal cName As String)
    mCategoryName.AsIParameter.AsString = cName
End Property

Public Property Get index() As String
    index = mCategoryID.AsIParameter.AsLong
End Property

Public Property Get IsExamination() As Boolean
    IsExamination = mIsExamin.AsIParameter.AsBoolean
End Property

Public Property Get ExamMode() As String
    ExamMode = mExaminMode.AsIParameter.AsString
End Property

Private Property Get IParameterWritable_ParameterList() As Collection
    Set IParameterWritable_ParameterList = New Collection
    With IParameterWritable_ParameterList
        .Add mCategoryName.AsIParameter.Name
        .Add mIsExamin.AsIParameter.Name
        .Add mExaminMode.AsIParameter.Name
    End With
End Property

Private Sub IParameterWritable_WriteOut(ByVal writer As IParameterWriter)
    mCategoryID.AsIParameter.WriteOut writer
    mCategoryName.AsIParameter.WriteOut writer
    mIsExamin.AsIParameter.WriteOut writer
    mExaminMode.AsIParameter.WriteOut writer
End Sub

Public Function AsIParameterWritable() As IParameterWritable
    Set AsIParameterWritable = Me
End Function

Private Sub Class_Terminate()
    Set mActionItems = Nothing
    Set mCategoryID = Nothing
    Set mCategoryName = Nothing
    Set mIsExamin = Nothing
    Set mExaminMode = Nothing
End Sub
