VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   Long�^�p�����[�^�N���X
'
'�ړI:
'   Long�^�̃p�����[�^��ێ�����
'
'�쐬��:
'   0145206097
'
Option Explicit

Implements IParameter
Implements IProcParameter
Implements ISiteArray
Implements IProcParamReplaceable

Private Const PARAMETER_REPLACER_NAME As String = "Name"
Private Const PARAMETER_REPLACER_VALUE As String = "Value"
Private Const PARAMETER_REPLACER_TYPE As String = "Type"
Private Const PARAMETER_REPLACER_UPPERLIMIT As String = "UpperLimit"
Private Const PARAMETER_REPLACER_LOWERLIMIT As String = "LowerLimit"
Private Const PARAMETER_REPLACER_COMMENT As String = "Comment"

Dim mDataName As String
Dim mDataValue As Long
Dim mHighLimit As Long
Dim mLowLimit As Long

Dim mComment As String

Private Sub Class_Initialize()
    mHighLimit = 99999
    mLowLimit = -99999
End Sub

Private Sub IParameter_Read(ByRef reader As IParameterReader)
    Me.AsIParameter.AsLong = reader.ReadAsLong(mDataName)
End Sub

Private Sub IParameter_WriteOut(ByRef writer As IParameterWriter)
    writer.WriteAsLong mDataName, Me.AsIParameter.AsLong
End Sub

Private Property Let IParameter_AsDouble(ByVal RHS As Double)
    checkLimits CDbl(RHS)
    mDataValue = CLng(RHS)
End Property

Private Property Get IParameter_AsDouble() As Double
    IParameter_AsDouble = CDbl(mDataValue)
End Property

Private Property Let IParameter_AsString(ByVal RHS As String)
    checkLimits CDbl(RHS)
    mDataValue = CLng(RHS)
End Property

Private Property Get IParameter_AsString() As String
    IParameter_AsString = CStr(mDataValue)
End Property

Private Property Let IParameter_AsBoolean(ByVal RHS As Boolean)
    TheError.Raise 9999, "CParamLong.AsIParameter.AsBoolean()", "[" & mDataName & "] Can Not Be Treated As A Boolean Value !"
End Property

Private Property Get IParameter_AsBoolean() As Boolean
    TheError.Raise 9999, "CParamLong.AsIParameter.AsBoolean()", "[" & mDataName & "] Can Not Be Treated As A Boolean Value !"
End Property

Private Property Let IParameter_AsLong(ByVal RHS As Long)
    checkLimits CDbl(RHS)
    mDataValue = CLng(RHS)
End Property

Private Property Get IParameter_AsLong() As Long
    IParameter_AsLong = mDataValue
End Property

Private Property Let IParameter_Name(ByVal RHS As String)
    mDataName = RHS
End Property

Private Property Get IParameter_Name() As String
    IParameter_Name = mDataName
End Property

Private Property Let IParameter_UpperLimit(ByVal RHS As Double)
    mHighLimit = CLng(RHS)
End Property

Private Property Get IParameter_UpperLimit() As Double
    IParameter_UpperLimit = CDbl(mHighLimit)
End Property

Private Property Let IParameter_LowerLimit(ByVal RHS As Double)
    mLowLimit = CLng(RHS)
End Property

Private Property Get IParameter_LowerLimit() As Double
    IParameter_LowerLimit = CDbl(mLowLimit)
End Property

Private Sub checkLimits(ByVal dataValue As Double)
    If (dataValue > mHighLimit) Then TheError.Raise 9999, "CParamLong.CheckLimits()", mDataName & " : [" & dataValue & "]  - It Exceeds The Upper Limit [" & mHighLimit & "] !"
    If (dataValue < mLowLimit) Then TheError.Raise 9999, "CParamLong.CheckLimits()", mDataName & " : [" & dataValue & "]  - It Falls Bellow The Lower Limit [" & mLowLimit & "] !"
End Sub

Public Function AsIParameter() As IParameter
    Set AsIParameter = Me
End Function



Public Function AsIProcParameter() As IProcParameter
    Set AsIProcParameter = Me
End Function

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
        .WriteAsLong "Value", mDataValue
        .WriteAsLong "LowerLimit", mLowLimit
        .WriteAsLong "UpperLimit", mHighLimit
        .WriteAsString "Comment", mComment
    End With
End Sub

Private Function IProcParameter_Clone() As IProcParameter
    Dim ret As CParamLong
    Set ret = New CParamLong
    With ret.AsIParameter
        .Name = mDataName
        .AsLong = mDataValue
        .LowerLimit = mLowLimit
        .UpperLimit = mHighLimit
    End With
    ret.AsIProcParameter.Comment = mComment
    Set IProcParameter_Clone = ret
End Function



Function AsSiteArray() As ISiteArray
    Set AsSiteArray = Me
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
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_VALUE, CStr(mDataValue), "Long", True, PARAMETER_REPLACER_VALUE)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_TYPE, TypeName(Me), "String", False, PARAMETER_REPLACER_TYPE)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_UPPERLIMIT, CStr(mHighLimit), "Long", True, PARAMETER_REPLACER_UPPERLIMIT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_LOWERLIMIT, CStr(mLowLimit), "Long", True, PARAMETER_REPLACER_LOWERLIMIT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COMMENT, mComment, "String", True, PARAMETER_REPLACER_COMMENT)
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
                mDataValue = CLng(.Value)
            Case PARAMETER_REPLACER_UPPERLIMIT
                mHighLimit = CLng(.Value)
            Case PARAMETER_REPLACER_LOWERLIMIT
                mLowLimit = CLng(.Value)
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

Private Function ISiteArray_AsDouble() As Double()
    Dim maxSiteIdx As Long
    maxSiteIdx = GetSiteCount - 1
    Dim ret() As Double
    ReDim ret(maxSiteIdx)
    Dim i As Long
    For i = 0 To maxSiteIdx
        ret(i) = mDataValue
    Next i
    ISiteArray_AsDouble = ret
End Function
Private Function GetSiteCount() As Long
    GetSiteCount = TheExec.sites.ExistingCount
End Function

Private Function ISiteArray_AsRaw() As Variant
    ISiteArray_AsRaw = mDataValue
End Function

Public Sub IProcParameter_Read(ByRef pReader As IProcParamReader)
End Sub