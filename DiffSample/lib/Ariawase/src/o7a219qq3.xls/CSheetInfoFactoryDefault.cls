VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSheetInfoFactoryDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   CSheetInfoのデフォルトFactory
'
'目的:
'   CSheetInfoはシートごとに専用のFactoryを用意するが、全てを用意するのは大変なのでデフォルトを用意
'   ワークシートの見出し情報から情報取得するので列が増減するシートにも対応
'
'作成者:
'   a_oshima


Option Explicit

Implements ISheetInfoFactory

Private Const mTITLE_ADDR = "B1"

Private mKeyAddress As Collection
Private mWorkSheetTypes As Collection

'CWorkSheetReaderFactoryの同一Constと同一内容であること
Private Const mSHTTYPE_AND_PKEYADDR As String = "Image Test Instance Table:B6," & _
                                 "Image Acquire Table:B4," & "Color Map Definition:B4," & _
                                 "Channel Map:B6," & "Image Kernel Definition:B3," & _
                                 "TEST Condition:B4," & "Device Configurations:B4," & _
                                 "Pin Map:C3," & _
                                 "Home:B3," & _
                                 "Job List:B4," & _
                                 "Image Pixel Map:B4," & _
                                 "Image Lookup Table:B3," & _
                                 "Image Kernel:B3," & _
                                 "Global Specs:B3," & _
                                 "DC Specs:B4," & _
                                 "AC Specs:B4," & _
                                 "Pattern Sets:B3," & _
                                 "Pattern Groups:B3," & _
                                 "Test Instances:B4," & _
                                 "Time Sets:B6," & _
                                 "Time Sets (Basic):B6," & _
                                 "Data Sheet Manager:B4," & _
                                 "DC Test Scenario:B6," & _
                                 "DC Playback Data:B5," & _
                                 "Offset Manager:B5," & _
                                 "Flow Table:H4," & _
                                 "PMD Definition:B5," & _
                                 "Errors:B3"


Private Sub Class_Terminate()
    Set mKeyAddress = Nothing
    Set mWorkSheetTypes = Nothing
End Sub

Private Sub Class_Initialize()
    Set mKeyAddress = New Collection
    Set mWorkSheetTypes = New Collection
    Dim pairs As Variant
    pairs = Split(mSHTTYPE_AND_PKEYADDR, ",")
    Dim pair As Variant
    For Each pair In pairs
        Dim Data As Variant
        Data = Split(pair, ":")
        mKeyAddress.Add Data(1), Data(0)
        mWorkSheetTypes.Add Data(0), Data(0)
    Next pair
End Sub

Private Function ISheetInfoFactory_CreateSheetInfo(ByRef pSheet As Worksheet) As CSheetInfo
    
    Dim idxRange As Range
    Dim primaryKeyRng As Range
    
    Dim wsType As Variant
    For Each wsType In mWorkSheetTypes
        If pSheet.Range(mTITLE_ADDR).Text = wsType Then
            Set primaryKeyRng = pSheet.Range(mKeyAddress(wsType))
            Set idxRange = getIdxRange(pSheet, primaryKeyRng)
        End If
    Next wsType
    If primaryKeyRng Is Nothing Then
        TheError.Raise 9999, TypeName(Me) & ".CreateSheetInfo", "Failed to specify sheet type: " & pSheet.Name
    End If
    
    Dim sheetInfo As CSheetInfo
    Set sheetInfo = New CSheetInfo
    
    Dim currRng As Range
    For Each currRng In idxRange
        If GetHeaderStr(currRng) <> vbNullString Then
            sheetInfo.SetKeyAddress(GetHeaderStr(currRng)) = Replace(currRng.Address, "$", "")
        End If
    Next currRng
    
    sheetInfo.SetPrimaryKey = GetHeaderStr(primaryKeyRng)
    Set currRng = Nothing
    Set primaryKeyRng = Nothing
    Set idxRange = Nothing

    
    Set ISheetInfoFactory_CreateSheetInfo = sheetInfo
End Function

Private Function getIdxRange(ByRef pSheet As Worksheet, ByRef pPrimaryKeyRng As Range) As Range
    Set getIdxRange = Application.Intersect(pSheet.Rows(pPrimaryKeyRng.Row), pSheet.UsedRange)
End Function


Private Function GetHeaderStr(ByRef rng As Range) As String
        Dim key As String
        If rng.MergeCells Then
            key = rng.MergeArea.Cells(1, 1).Value
            key = key & GetGroupHeaderStr(rng.MergeArea)
        Else
            key = rng.Text
            key = key & GetGroupHeaderStr(rng)
        End If
        GetHeaderStr = key
End Function

Private Function GetGroupHeaderStr(ByRef childRng As Range) As String
        If childRng.offset(-1, 0).MergeCells Then
            GetGroupHeaderStr = "@" & childRng.offset(-1, 0).MergeArea.Cells(1, 1).Value
        ElseIf childRng.offset(-1, 0).Value <> vbNullString Then
            GetGroupHeaderStr = "@" & childRng.offset(-1, 0).Value
        Else
            GetGroupHeaderStr = vbNullString
        End If
End Function

Public Function AsISheetInfoFactory() As ISheetInfoFactory
    Set AsISheetInfoFactory = Me
End Function

