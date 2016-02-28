VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   String�^�p�����[�^�N���X
'
'�ړI:
'   String�^�̃p�����[�^��ێ�����
'
'�쐬��:
'   0145206097
'
Option Explicit

Implements IParameter
Implements IProcParameter
Implements IProcParamReplaceable

Private Const PARAMETER_REPLACER_NAME As String = "Name"
Private Const PARAMETER_REPLACER_VALUE As String = "Value"
Private Const PARAMETER_REPLACER_COMMENT As String = "Comment"
Private Const PARAMETER_REPLACER_TYPE As String = "Type"

Dim mDataName As String
Dim mStrData As String

Dim mComment As String

Private Sub IParameter_Read(ByRef reader As IParameterReader)
    Me.AsIParameter.AsString = reader.ReadAsString(mDataName)
End Sub

Private Sub IParameter_WriteOut(ByRef writer As IParameterWriter)
    writer.WriteAsString mDataName, Me.AsIParameter.AsString
End Sub

Private Property Let IParameter_AsDouble(ByVal RHS As Double)
    TheError.Raise 9999, "CParamString.AsIParameter.AsDouble()", "[" & mDataName & "] Can Not Be Treated As A Double Value !"
End Property

Private Property Get IParameter_AsDouble() As Double
    TheError.Raise 9999, "CParamString.AsIParameter.AsDouble()", "[" & mDataName & "] Can Not Be Treated As A Double Value !"
End Property

Private Property Let IParameter_AsString(ByVal RHS As String)
    If Not IsOneByte(RHS) Then
        TheError.Raise 9999, "CParamString.AsIParameter.AsString", mDataName & " : [" & RHS & "]  - 2-Byte Characters In This String Are Invalid !"
    End If
    mStrData = RHS
End Property

Private Property Get IParameter_AsString() As String
    IParameter_AsString = mStrData
End Property

Private Property Let IParameter_AsBoolean(ByVal RHS As Boolean)
    TheError.Raise 9999, "CParamString.AsIParameter.AsBoolean()", "[" & mDataName & "] Can Not Be Treated As A Boolean Value !"
End Property

Private Property Get IParameter_AsBoolean() As Boolean
    TheError.Raise 9999, "CParamString.AsIParameter.AsBoolean()", "[" & mDataName & "] Can Not Be Treated As A Boolean Value !"
End Property

Private Property Let IParameter_AsLong(ByVal RHS As Long)
    TheError.Raise 9999, "CParamString.AsIParameter.AsLong()", "[" & mDataName & "] Can Not Be Treated As A Long Value !"
End Property

Private Property Get IParameter_AsLong() As Long
    TheError.Raise 9999, "CParamString.AsIParameter.AsLong()", "[" & mDataName & "] Can Not Be Treated As A Long Value !"
End Property

Private Property Let IParameter_Name(ByVal RHS As String)
    mDataName = RHS
End Property

Private Property Get IParameter_Name() As String
    IParameter_Name = mDataName
End Property

Private Property Let IParameter_UpperLimit(ByVal RHS As Double)
    TheError.Raise 9999, "CParamString.AsIParameter.UpperLimit()", "Does Not Support This Method !"
End Property

Private Property Get IParameter_UpperLimit() As Double
    TheError.Raise 9999, "CParamString.AsIParameter.UpperLimit()", "Does Not Support This Method !"
End Property

Private Property Let IParameter_LowerLimit(ByVal RHS As Double)
    TheError.Raise 9999, "CParamString.AsIParameter.LowerLimit()", "Does Not Support This Method !"
End Property

Private Property Get IParameter_LowerLimit() As Double
    TheError.Raise 9999, "CParamString.AsIParameter.LowerLimit()", "Does Not Support This Method !"
End Property

Public Function AsIParameter() As IParameter
    Set AsIParameter = Me
End Function



Public Function AsIProcParameter() As IProcParameter
    Set AsIProcParameter = Me
End Function

Private Sub IProcParamReplaceable_GetContents(ByVal pRoot As String, ByRef pRes As Collection)
'���e:
'   �p�����[�^�̐ݒ���e���擾����
'
'�p�����[�^:
'   [pRoot]     IN  String�^:  ���̃p�����[�^�N���X�̊K�w���
'   [pRes]      IN/OUT  Collection�^:  �p�����[�^�̐ݒ���e���i�[����R���N�V����
'
'���ӎ���:
'
'

    With pRes
        '   2012-01-27 (0145184931): Name��ҏW�s�ɕύX
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_NAME, mDataName, "String", False, PARAMETER_REPLACER_NAME)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_VALUE, mStrData, "String", True, PARAMETER_REPLACER_VALUE)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COMMENT, mComment, "String", True, PARAMETER_REPLACER_COMMENT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_TYPE, TypeName(Me), "String", False, PARAMETER_REPLACER_TYPE)
    End With
End Sub

Private Sub IProcParamReplaceable_ReplaceContent(ByVal pRoot As String, ByRef pParamContents As CImpLocalParamsContent)
'���e:
'   �p�����[�^�̓��e������������
'
'�p�����[�^:
'   [pRoot]             IN  String�^:  ���̃p�����[�^�N���X�̊K�w���
'   [pParamContents]    IN  CImpLocalParamsContent�^:  �����������s���p�����[�^�̓��e
'
'���ӎ���:
'
'

    With pParamContents
        Dim subParamLabel As String
        subParamLabel = XLibImpExperimentUtility.GetSubParamLabel(.Path, pRoot)
        Select Case subParamLabel
'   2012-01-27 (0145184931): Name��ҏW�s�ɕύX
'            Case PARAMETER_REPLACER_NAME
'                mDataName = .value
            Case PARAMETER_REPLACER_VALUE
                mStrData = .Value
            Case PARAMETER_REPLACER_COMMENT
                mComment = .Value
        End Select
    End With
End Sub

Private Sub AddParamContents(ByRef pRes As Collection, ByVal pPath As String, ByVal pValue As String, ByVal pTypeName As String, ByVal pEditable As Boolean, ByVal pLabel As String)
'���e:
'   �p�����[�^�̏����R���N�V�����ɒǉ�����
'
'�p�����[�^:
'   [pRes]          IN/OUT  Collection�^:  �p�����[�^�̐ݒ���e��ǉ�����R���N�V����
'   [pPath]         IN  String�^:  �p�����[�^�̊K�w���
'   [pValue]        IN  String�^:  �p�����[�^�̒l
'   [pTypeName]     IN  String�^:  �p�����[�^�̃f�[�^�^
'   [pEditable]     IN  Boolean�^:  �p�����[�^���ҏW�\���ǂ���
'   [pLabel]        IN  String�^:  �p�����[�^��
'
'���ӎ���:
'
'

    Dim myContents As New CImpLocalParamsContent
    Call myContents.Initialize(pPath, pValue, pTypeName, pEditable, pLabel)
    Call pRes.Add(myContents)
End Sub

Private Property Let IProcParameter_Comment(ByVal pComment As String)
    mComment = pComment
End Property
Private Property Get IProcParameter_Comment() As String
    IProcParameter_Comment = mComment
End Property

Private Property Let IProcParameter_Name(ByVal pName As String)
    Me.AsIParameter.Name = pName
End Property
Private Property Get IProcParameter_Name() As String
    IProcParameter_Name = Me.AsIParameter.Name
End Property

Private Sub IProcParameter_Dump(ByRef pWriter As IProcParamWriter)
    With pWriter
        .WriteAsString "Name", mDataName
        .WriteAsString "Type", TypeName(Me)
        .WriteAsString "Value", mStrData
        .WriteAsString "Comment", mComment
    End With
End Sub

Private Function IProcParameter_Clone() As IProcParameter
    Dim ret As CParamString
    Set ret = New CParamString
    With ret.AsIParameter
        .Name = mDataName
        .AsString = mStrData
    End With
    ret.AsIProcParameter.Comment = mComment
    Set IProcParameter_Clone = ret
End Function

Private Sub IProcParameter_Read(ByRef pReader As IProcParamReader)
End Sub