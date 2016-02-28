VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COffsetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   オフセットデータ管理クラス
'
'目的:
'   結果管理クラスの測定結果から補正データを作成する
'
'作成者:
'   0145206097
'
Option Explicit

Dim mOffsetDataReader As IParameterReader
Dim mCoefficient As Collection
Dim mConstant As Collection

Private Sub Class_Initialize()
    Set mCoefficient = New Collection
    Set mConstant = New Collection
End Sub

Private Sub Class_Terminate()
    Set mOffsetDataReader = Nothing
    Set mCoefficient = Nothing
    Set mConstant = Nothing
End Sub

Public Sub Initialize(ByVal reader As IParameterReader)
    Set mOffsetDataReader = reader
    Me.Load
End Sub

Public Sub Load()
    constructOffsetData mOffsetDataReader
End Sub

Public Sub Calculate(ByVal TestLabel As String, ByVal resultKeeper As CResultManager)
    On Error GoTo errHander
    If mOffsetDataReader.ReadAsBoolean(DATA_CHANGED) Then
        constructOffsetData mOffsetDataReader
    End If
    If IsContain(TestLabel) Then
        Dim tempdata() As Double
        TheResult.GetResult TestLabel, tempdata
        Dim SiteIndex As Long
        For SiteIndex = 0 To UBound(tempdata)
            tempdata(SiteIndex) = tempdata(SiteIndex) * _
                                  mCoefficient(TestLabel)(SiteIndex) + _
                                  mConstant(TestLabel)(SiteIndex)
        Next SiteIndex
        TheResult.Delete TestLabel
        TheResult.Add TestLabel, tempdata
    Else
        Err.Raise 9999, "COffsetManager.Calculate", "Not Define The Offset Data Of Test Name [" & TestLabel & "] !"
    End If
    Exit Sub
errHander:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "COffsetManager.Calculate", "Can Not Offset The Result [" & TestLabel & "] !"
End Sub

Private Function IsContain(ByVal keyName As String) As Boolean
    On Error GoTo notContain
    mCoefficient.Item (keyName)
    mConstant.Item (keyName)
    IsContain = True
    Exit Function
notContain:
    IsContain = False
End Function

Private Sub constructOffsetData(ByVal reader As IParameterReader)
    On Error GoTo ErrHandler
    Set mCoefficient = New Collection
    Set mConstant = New Collection
    '### ファイルストリームの準備 #########################
    Dim fileStream As IFileStream
    Set fileStream = reader
    fileStream.Rewind
    '### パラメータオブジェクトの作成 #####################
    Dim ParamLabel As CParamName
    Set ParamLabel = CreateCParamName
    ParamLabel.AsIParameter.Name = OFFSET_LABEL
    Dim paramCoef As CParamString
    Set paramCoef = CreateCParamString
    paramCoef.AsIParameter.Name = OFFSET_COEF
    Dim paramConst As CParamString
    Set paramConst = CreateCParamString
    paramConst.AsIParameter.Name = OFFSET_CONS
    '### データ読み込み実行 ###############################
    Do While Not fileStream.IsEOR
        '### テストラベルの読み込み #######################
        ParamLabel.AsIParameter.Read reader
        '### 係数データの読み込み #########################
        Dim coeDataList As New Collection
        With paramCoef.AsIParameter
            .Read reader
            Set coeDataList = DecomposeStringList(.AsString)
        End With
        Dim coeDataArray() As Double
        ReDim coeDataArray(coeDataList.Count - 1)
        Dim dataIndex As Long
        For dataIndex = 0 To UBound(coeDataArray)
            If coeDataList.Item(dataIndex + 1) = NOT_DEFINE Then
                coeDataArray(dataIndex) = 1
            Else
                coeDataArray(dataIndex) = CDbl(coeDataList.Item(dataIndex + 1))
            End If
        Next dataIndex
        '### 定数データの読み込み #########################
        Dim consDataList As New Collection
        With paramConst.AsIParameter
            .Read reader
            Set consDataList = DecomposeStringList(.AsString)
        End With
        Dim consDataArray() As Double
        ReDim consDataArray(consDataList.Count - 1)
        For dataIndex = 0 To UBound(consDataArray)
            If consDataList.Item(dataIndex + 1) = NOT_DEFINE Then
                consDataArray(dataIndex) = 0
            Else
                Dim MainUnit As String
                Dim SubUnit As String
                Dim SubValue As Double
                SplitUnitValue consDataList.Item(dataIndex + 1), MainUnit, SubUnit, SubValue
                Dim paramData As CParamStringWithUnit
                Set paramData = CreateCParamStringWithUnit
                With paramData
                    .Initialize MainUnit
                    With .AsIParameter
                        .AsString = consDataList.Item(dataIndex + 1)
                        consDataArray(dataIndex) = .AsDouble
                    End With
                End With
            End If
        Next dataIndex
        mCoefficient.Add coeDataArray, ParamLabel.AsIParameter.AsString
        mConstant.Add consDataArray, ParamLabel.AsIParameter.AsString
        fileStream.MoveNext
    Loop
    Exit Sub
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Err.Raise 9999, "COffsetManager.OffsetDataLoad", "Can Not Load The Offset Data !"
End Sub

