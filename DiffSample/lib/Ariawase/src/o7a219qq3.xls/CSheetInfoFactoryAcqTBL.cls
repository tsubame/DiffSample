VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSheetInfoFactoryAcqTBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   CSheetInfoのFactory
'
'目的:
'   Image Acquire Table用
'
'作成者:
'   a_oshima

Option Explicit

Implements ISheetInfoFactory

Private Const mINDEX_PKEY As String = "Group Name"
Private Const mINDEX_LABELS As String = "Group Name:B4," & _
                                 "Macro Name:C4," & _
                                 "Auto Acquire:D4," & _
                                 "Instance Name:E4," & _
                                 "Arg0-Arg79@Parameters:F4," & _
                                 "Comment:CH4"


Private Sub Class_Terminate()
'
End Sub

Private Sub Class_Initialize()
'
End Sub

Private Function ISheetInfoFactory_CreateSheetInfo(ByRef pSheet As Worksheet) As CSheetInfo
    Dim sheetInfo As CSheetInfo
    Set sheetInfo = New CSheetInfo
    
    Dim pairs As Variant
    pairs = Split(mINDEX_LABELS, ",")
    Dim pair As Variant
    For Each pair In pairs
        Dim Data As Variant
        Data = Split(pair, ":")
        If Data(0) Like "Arg*-Arg*@*" Then
            Dim argss As Variant
            Dim args As Variant
            Dim fromNo As Long
            Dim toNo As Long
            argss = Split(Data(0), "@")
            args = Split(argss(0), "-")
            fromNo = CLng(Replace(args(0), "Arg", ""))
            toNo = CLng(Replace(args(1), "Arg", ""))
            Dim i As Long
            For i = fromNo To toNo
                sheetInfo.SetKeyAddress("Arg" & i & "@" & argss(1)) = Replace(pSheet.Range(Data(1)).offset(0, i - fromNo).Address, "$", "")
            Next i
        Else
            sheetInfo.SetKeyAddress(Data(0)) = Data(1)
        End If
    Next pair
    sheetInfo.SetPrimaryKey = mINDEX_PKEY
    
    Set ISheetInfoFactory_CreateSheetInfo = sheetInfo
End Function

Public Function AsISheetInfoFactory() As ISheetInfoFactory
    Set AsISheetInfoFactory = Me
End Function





