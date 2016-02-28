VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSheetInfoFactoryJobList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'äTóv:
'   CSheetInfoÇÃFactory
'
'ñ⁄ìI:
'   Job Listóp
'
'çÏê¨é“:
'   a_oshima

Option Explicit

Implements ISheetInfoFactory

Private Const mTITLE_ADDR = "B1"

Private Const mINDEX_PKEY As String = "Job Name"
Private Const mINDEX_LABELS As String = "Job Name:B4," & _
                                 "Pin Map@Sheet Parameters:C4," & _
                                 "Test Instances@Sheet Parameters:D4," & _
                                 "Flow Table@Sheet Parameters:E4," & _
                                 "AC Specs@Sheet Parameters:F4," & _
                                 "DC Specs@Sheet Parameters:G4," & _
                                 "Pattern Sets@Sheet Parameters:H4," & _
                                 "Pattern Groups@Sheet Parameters:I4," & _
                                 "Bin Table@Sheet Parameters:J4," & _
                                 "Characterization@Sheet Parameters:K4," & _
                                 "Test Procedures@Sheet Parameters:L4," & _
                                 "Comment:M4"



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

        sheetInfo.SetKeyAddress(Data(0)) = Data(1)

    Next pair
    sheetInfo.SetPrimaryKey = mINDEX_PKEY
    
    
    Set ISheetInfoFactory_CreateSheetInfo = sheetInfo
End Function


Public Function AsISheetInfoFactory() As ISheetInfoFactory
    Set AsISheetInfoFactory = Me
End Function




