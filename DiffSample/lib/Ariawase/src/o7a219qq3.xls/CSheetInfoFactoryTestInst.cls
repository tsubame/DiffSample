VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSheetInfoFactoryTestInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'äTóv:
'   CSheetInfoÇÃFactory
'
'ñ⁄ìI:
'   Test Instancesóp
'
'çÏê¨é“:
'   a_oshima

Option Explicit

Implements ISheetInfoFactory

Private Const mINDEX_PKEY As String = "Test Name"
Private Const mINDEX_LABELS As String = "Test Name:B4," & _
                                 "Type@Test Procedure:C4," & _
                                 "Name@Test Procedure:D4," & _
                                 "Called As@Test Procedure:E4," & _
                                 "Category@DC Specs:F4," & _
                                 "Selector@DC Specs:G4," & _
                                 "Category@AC Specs:H4," & _
                                 "Selector@AC Specs:I4," & _
                                 "Time Sets@Sheet Parameters:J4," & _
                                 "Edge Sets@Sheet Parameters:K4," & _
                                 "PinLevels@Sheet Parameters:L4," & _
                                 "Overlay:M4," & _
                                 "Arg0-Arg79@Other Parameters:N4," & _
                                 "Comment:CP4"


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



