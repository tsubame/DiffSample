VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkSheetReaderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   汎用ワークシートReaderを作成するFactory
'
'目的:
'   シート名を指定すればどの種類のシートかを自動判別して適切なReaderを作成する
'
'作成者:
'   a_oshima
'
'前提条件:
'   グローバルオブジェクト TheError:CErrManager が存在すること
'

Option Explicit

Private Const mTITLE_ADDR = "B1"

Private mKeyAddress As Collection
Private mWorkSheetTypes As Collection

'CSheetInfoFactoryDefaultの同一Constと同一内容であること
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

Function CreateReader(ByRef pBook As Workbook, ByVal pSheetName As String, Optional ByVal enableReport As Boolean) As CWorkSheetReader
'内容:
'   CWorkSheetReaderのインスタンスを返します
'
'パラメータ:
'   [pBook]  In  WorkBook型:  ワークブックインスタンス
'   [pSheetName]  In  String型:  対象ワークシート名
'   [enableReport]  In  Boolean型:  現在セルのレポート機能On/Off
'
'戻り値:
'    CWorkSheetReaderのインスタンス
'
'例外:
'   指定した名前のワークシートがないときは例外9999を投げます
'
'注意事項:
    
    Dim reader As CWorkSheetReader
    Set reader = New CWorkSheetReader
    
    reader.EnableReportCurrentCell = enableReport
    
    Dim sht As Worksheet
    Dim factorySht As ISheetInfoFactory
    
    On Error GoTo catch
    Set sht = pBook.Worksheets(pSheetName)
    On Error GoTo 0
    
    If sht.Range(mTITLE_ADDR).Text = "Test Instances" Then
        Set factorySht = New CSheetInfoFactoryTestInst
#If ITS <> 0 Then
    ElseIf sht.Range(mTITLE_ADDR).Text = "Image Test Scenario" Then
        Set factorySht = New CSheetInfoFactoryImageScenario
#End If
    ElseIf sht.Range(mTITLE_ADDR).Text = "Job List" Then
        Set factorySht = New CSheetInfoFactoryJobList
    ElseIf sht.Range(mTITLE_ADDR).Text = "Image Acquire Table" Then
        Set factorySht = New CSheetInfoFactoryAcqTBL
    Else
        Dim wsType As Variant
        For Each wsType In mWorkSheetTypes
            If sht.Range(mTITLE_ADDR).Text = wsType Then
                Set factorySht = New CSheetInfoFactoryDefault
            End If
        Next wsType
        If factorySht Is Nothing Then
            TheError.Raise 9999, TypeName(Me) & ".CreateReader", "Failed to specify sheet type: " & pSheetName
        End If
    End If
    
    Call reader.Initialize(pBook, pSheetName, factorySht)
            
    Set CreateReader = reader
    Exit Function
catch:
    TheError.Raise 9999, TypeName(Me) & ".CreateReader", "Cannot find Worksheet:" & pSheetName
End Function





