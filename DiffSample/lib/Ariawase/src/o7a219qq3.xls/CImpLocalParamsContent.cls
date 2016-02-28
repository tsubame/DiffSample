VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImpLocalParamsContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ローカルパラメータの情報を保持する
'
'目的:
'   パラメータクラスのメンバー編集の情報を保持する
'
'作成者:
'   0145184931
'
Option Explicit

Private mPath As String         'パラメータのパスを保持
Private mValue As String        'パラメータの値を保持
Private mTypeName As String     'パラメータのデータ型名を保持
Private mEditable As Boolean    'パラメータが編集可能かを保持
Private mLabel As String        'パラメータのラベル名を保持

Public Sub Initialize(ByVal pPath As String, ByVal pValue As String, ByVal pTypeName As String, ByVal pEditable As Boolean, ByVal pLabel As String)
'内容:
'   パラメータ情報を初期化する
'
'[pPath]        IN String型:    パラメータのパス
'[pValue]       IN String型:    パラメータの値
'[pTypeName]    IN String型:    パラメータのデータ型名を保持
'[pEditable]    IN Boolean型:   パラメータが編集可能かを保持
'[pLabel]       IN String型:    パラメータのラベル名
'
'備考:
'
    mPath = pPath
    mValue = pValue
    mTypeName = pTypeName
    mEditable = pEditable
    mLabel = pLabel
End Sub

Public Function Clone() As CImpLocalParamsContent
    Dim myClone As CImpLocalParamsContent
    Set myClone = New CImpLocalParamsContent
    Call myClone.Initialize(mPath, mValue, mTypeName, mEditable, mLabel)
    Set Clone = myClone
End Function

Public Property Get Path() As String
    Path = mPath
End Property

Public Property Let Path(ByVal pPath As String)
    mPath = pPath
End Property

Public Property Get Value() As String
    Value = mValue
End Property

Public Property Let Value(ByVal pValue As String)
    mValue = pValue
End Property

Public Property Get TypeName() As String
    TypeName = mTypeName
End Property

Public Property Let TypeName(ByVal pTypeName As String)
    mTypeName = pTypeName
End Property

Public Property Get Editable() As Boolean
    Editable = mEditable
End Property

Public Property Let Editable(ByVal pEditable As Boolean)
    mEditable = pEditable
End Property

Public Property Get Label() As String
    Label = mLabel
End Property

Public Property Let Label(ByVal pLabel As String)
    mLabel = pLabel
End Property
