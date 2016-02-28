VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkSheetReaderManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   汎用ワークシートReaderを管理する
'
'目的:
'   WorkSheetReaderFactoryのCreateReaderメソッドにはシート名が必要だが
'   Activeなシート名を特定するのが大変
'   そこで、シートの種類をEnumで指定すれば適切なReaderのインスタンスが得られるようにする
'   また、同一シートのReaderのインスタンスが複数できないように一括管理する
'
'作成者:
'   a_oshima
'
'前提条件:
'   グローバルオブジェクト TheError:CErrManager が存在すること

Option Explicit

Private mReaderFactory As CWorkSheetReaderFactory
Private mReaders As Collection

Private mWorkSheetOfJobContext As Collection
Private mWorkSheetOfTestContext As Collection
Private mWorkSheetUniq As Collection
Private mWorkSheetCostom As Collection

Private Const mJOBLIST_MANAGING_SHEETS = "Pin Map,Test Instances,Flow Table,AC Specs,DC Specs,Pattern Sets,Pattern Groups,Bin Table,Characterization,Test Procedures"
Private Const mSHEETMANAGER_MANAGING_SHEETS = "DC Test Scenario,DC Playback Data,Offset Manager"
Private Const mUNIQ_SHEETS_TERADYNE = "Job List,Global Specs,Image Acquire Table"
Private Const mUNIQ_SHEETS_EEEJOB = "Image Test Instance Table,PMD Definition,TEST Condition,Image Test Scenario,Data Sheet Manager,Color Map Definition,Image Kernel Definition,Device Configurations"

'順番を変えないこと！
Public Enum eSheetType
    'Jobに1枚しか許されないもの
    shtTypeJobList
    shtTypeGlobalSpecs
    shtTypeAcquire
    
    'JobContextから取得
    shtTypeTestInstances
    shtTypeFlow
    shtTypePinMap
    shtTypeDCSpecs
    shtTypeACSpecs
    shtTypePatSet
    shtTypePatGrp
    shtTypeBinTable
    shtTypeCharacterization
    shtTypeTestProcedures

    'ChansはTheExecから
    shtTypeChans
    
    'TestContextから取得
    shtTypeTimeSets
    shtTypeEdgeSets
    shtTypePinLevels
    
    'カスタムシート（1枚しか許されない）
    shtTypeTestCond
    shtTypeDeviceConfigurations
    shtTypePMDDefinition
    shtTypeImgTestInstTBL
#If ITS <> 0 Then
    shtTypeImgTestScenario
#End If
    shtTypeDataShtMgr
    shtTypeKernel
    shtTypeColorMap
    
    'DataSheetManagerシートから（JobContext）
    shtTypeDcTestScenario
    shtTypeDcPlaybackData
    shtTypeOffsetManager
End Enum

Function GetActiveSheetName(ByVal pSheetType As eSheetType) As String
   
    Select Case pSheetType
    Case eSheetType.shtTypeJobList To eSheetType.shtTypeAcquire
        GetActiveSheetName = GetActiveSheetNameFromUniq(pSheetType)
    Case eSheetType.shtTypeTestInstances To eSheetType.shtTypeTestProcedures
        GetActiveSheetName = GetActiveSheetNameFromJobContext(pSheetType)
    Case eSheetType.shtTypeChans
        GetActiveSheetName = TheExec.CurrentChanMap()
    Case eSheetType.shtTypeTimeSets To eSheetType.shtTypePinLevels
        GetActiveSheetName = GetActiveSheetNameFromTestContext(pSheetType)
    
    Case eSheetType.shtTypeTestCond To eSheetType.shtTypeDataShtMgr
        GetActiveSheetName = GetActiveSheetNameCustomUniq(pSheetType)
        
    Case eSheetType.shtTypeTestCond To eSheetType.shtTypeColorMap
        GetActiveSheetName = GetActiveSheetNameCustomUniq(pSheetType)
        
    Case eSheetType.shtTypeDcTestScenario To eSheetType.shtTypeOffsetManager
        GetActiveSheetName = GetActiveSheetNameCustomContext(pSheetType)
    Case Else
        TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetName", "Illegal Sheet Type:" & pSheetType
    End Select

End Function

Private Function GetActiveSheetNameCustomUniq(ByVal pSheetType As eSheetType) As String

    Dim SheetName As String
    Select Case pSheetType
    Case eSheetType.shtTypeTestCond
        SheetName = "TEST Condition"
    Case eSheetType.shtTypePMDDefinition
        SheetName = "PMD Definition"
    Case eSheetType.shtTypeImgTestInstTBL
        SheetName = "Image Test Instance Table"
#If ITS <> 0 Then
    Case eSheetType.shtTypeImgTestScenario
        SheetName = "Image Test Scenario"
#End If
    Case eSheetType.shtTypeDataShtMgr
        SheetName = "Data Sheet Manager"
    Case eSheetType.shtTypeDeviceConfigurations
        SheetName = "Device Configurations"
    Case eSheetType.shtTypeKernel
        SheetName = "Image Kernel Definition"
    Case eSheetType.shtTypeColorMap
        SheetName = "Color Map Definition"
    Case Else
        TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetNameCustomUniq", "Illegal Sheet Type:" & pSheetType
    End Select
    
    On Error GoTo catch
    GetActiveSheetNameCustomUniq = mWorkSheetCostom.Item(SheetName)
    On Error GoTo 0
    Exit Function
catch:
    TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetNameCustomUniq", "Worksheet does not exist: " & SheetName
End Function

Private Function GetActiveSheetNameFromUniq(ByVal pSheetType As eSheetType) As String

    Dim SheetName As String
    Select Case pSheetType
    Case eSheetType.shtTypeJobList
        SheetName = "Job List"
    Case eSheetType.shtTypeGlobalSpecs
        SheetName = "Global Specs"
    Case eSheetType.shtTypeAcquire
        SheetName = "Image Acquire Table"
    Case Else
        TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetNameFromUniq", "Illegal Sheet Type:" & pSheetType
    End Select
    
    On Error GoTo catch
    GetActiveSheetNameFromUniq = mWorkSheetUniq.Item(SheetName)
    On Error GoTo 0
    Exit Function
catch:
    TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetNameFromUniq", "Worksheet does not exist: " & SheetName
End Function

Private Function GetActiveSheetNameFromTestContext(ByVal pSheetType As eSheetType) As String
    Dim activeTestName As String
    activeTestName = TheExec.DataManager.InstanceName()
        
    Dim sheetIdx As String
    Select Case pSheetType
    Case eSheetType.shtTypeTimeSets
        sheetIdx = "Time Sets@Sheet Parameters"
    Case eSheetType.shtTypeEdgeSets
        sheetIdx = "Edge Sets@Sheet Parameters"
    Case eSheetType.shtTypePinLevels
        sheetIdx = "Pin Levels@Sheet Parameters"
    Case Else
        TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetNameFromTestContext", "Illegal Sheet Type:" & pSheetType
    End Select
    
    On Error GoTo catch
    GetActiveSheetNameFromTestContext = mWorkSheetOfTestContext.Item(activeTestName).Item(sheetIdx)
    On Error GoTo 0
    Exit Function
catch:
    TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetNameFromTestContext", "Worksheet Name does not specified at Test Instances sheet: " & sheetIdx
End Function

Private Function GetActiveSheetNameFromJobContext(ByVal pSheetType As eSheetType) As String
    
    Dim activeJobName As String
    activeJobName = TheExec.CurrentJob()
 
    Dim sheetIdx As String
    Select Case pSheetType
    Case eSheetType.shtTypeTestInstances
        sheetIdx = "Test Instances@Sheet Parameters"
    Case eSheetType.shtTypeFlow
        sheetIdx = "Flow Table@Sheet Parameters"
    Case eSheetType.shtTypePinMap
        sheetIdx = "Pin Map@Sheet Parameters"
    Case eSheetType.shtTypeDCSpecs
        sheetIdx = "DC Specs@Sheet Parameters"
    Case eSheetType.shtTypeACSpecs
        sheetIdx = "AC Specs@Sheet Parameters"
    Case eSheetType.shtTypePatSet
        sheetIdx = "Pattern Sets@Sheet Parameters"
    Case eSheetType.shtTypePatGrp
        sheetIdx = "Pattern Groups@Sheet Parameters"
    Case eSheetType.shtTypeBinTable
        sheetIdx = "Bin Table@Sheet Parameters"
    Case eSheetType.shtTypeCharacterization
        sheetIdx = "Characterization@Sheet Parameters"
    Case eSheetType.shtTypeTestProcedures
        sheetIdx = "Test Procedures@Sheet Parameters"
    Case Else
        TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetNameFromJobContext", "Illegal Sheet Type:" & pSheetType
    End Select
    
    On Error GoTo catch
    GetActiveSheetNameFromJobContext = mWorkSheetOfJobContext.Item(activeJobName).Item(sheetIdx)
    On Error GoTo 0
    Exit Function
catch:
    TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetNameFromJobContext", "Worksheet Name does not specified at Job List sheet: " & sheetIdx
End Function

Private Function GetActiveSheetNameCustomContext(ByVal pSheetType As eSheetType) As String
    
    Dim activeJobName As String
    activeJobName = TheExec.CurrentJob()
 
    Dim sheetIdx As String
    Select Case pSheetType
    Case eSheetType.shtTypeDcTestScenario
        sheetIdx = "DC Test Scenario@Sheet Parameters"
    Case eSheetType.shtTypeDcPlaybackData
        sheetIdx = "DC Playback Data@Sheet Parameters"
    Case eSheetType.shtTypeOffsetManager
        sheetIdx = "Offset Manager@Sheet Parameters"
    Case Else
        TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetNameCustomContext", "Illegal Sheet Type:" & pSheetType
    End Select
    
    On Error GoTo catch
    GetActiveSheetNameCustomContext = mWorkSheetCostom.Item(activeJobName).Item(sheetIdx)
    On Error GoTo 0
    Exit Function
catch:
    TheError.Raise 9999, TypeName(Me) & ".GetActiveSheetNameCustomContext", "Worksheet Name does not specified at Data Sheet Manager sheet: " & sheetIdx
End Function

Function GetReaderInstance(ByVal pSheetType As eSheetType) As CWorkSheetReader
'Readerが既に作成されている場合はそのインスタンスを返す（Rewindする）
'Readerが未作成の場合は新規に作る
'シートが存在しない場合はNothingを返す

    Dim activeShtName As String
    activeShtName = GetActiveSheetName(pSheetType)
    If activeShtName = vbNullString Then
        Set GetReaderInstance = Nothing
        Exit Function
    End If
    On Error GoTo catch
    Set GetReaderInstance = mReaders.Item(activeShtName)
    GetReaderInstance.AsIFileStream.Rewind
    On Error GoTo 0
    Exit Function
catch:
    Dim reader As CWorkSheetReader
    Set reader = mReaderFactory.CreateReader(ThisWorkbook, activeShtName)
    Call mReaders.Add(reader, activeShtName)
    Set GetReaderInstance = reader
    Set reader = Nothing
End Function

Private Sub Class_Initialize()
    Set mWorkSheetOfJobContext = New Collection
    Set mWorkSheetOfTestContext = New Collection
    Set mWorkSheetUniq = New Collection
    Set mWorkSheetCostom = New Collection
    
    Set mReaders = New Collection

    Set mReaderFactory = New CWorkSheetReaderFactory
    
    
    Call specifyUniqueSheet
    Call specifyCostomSheet
    
    If IsExistJobList = True Then
        Call readSheetContext("Job Name", mWorkSheetUniq.Item("Job List"), mWorkSheetOfJobContext)
    Else
        Call specifyJobListManagingSheet
    End If
    Call readSheetContext("Test Name", GetActiveSheetName(shtTypeTestInstances), mWorkSheetOfTestContext)
    
    If IsExistSheetManager = True Then
        Call readSheetContext("Job Name", mWorkSheetCostom.Item("Data Sheet Manager"), mWorkSheetCostom)
    Else
        Call specifySheetManagerManagingSheet
    End If

End Sub

Private Sub specifyCostomSheet()
    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Worksheets
        Dim shtType As Variant
        For Each shtType In Split(mUNIQ_SHEETS_EEEJOB, ",")
            If sht.Range("B1") = shtType Then

                On Error GoTo catch
                Call mWorkSheetCostom.Add(sht.Name, shtType)
                On Error GoTo 0

            End If
        Next shtType
    Next sht
    Set sht = Nothing
    Exit Sub
catch:
    TheError.Raise 9999, TypeName(Me) & ".specifyCostomSheet", "Too many Worksheet: " & shtType
End Sub

Private Sub specifyUniqueSheet()
    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Range("A1") Like "DFF*" Then
            Dim shtType As Variant
            For Each shtType In Split(mUNIQ_SHEETS_TERADYNE, ",")
                If sht.Range("B1") = shtType Then

                    On Error GoTo catch
                    Call mWorkSheetUniq.Add(sht.Name, shtType)
                    On Error GoTo 0

                End If
            Next shtType
        End If
    Next sht
    Set sht = Nothing
    Exit Sub
catch:
    TheError.Raise 9999, TypeName(Me) & ".specifyUniqueSheet", "Too many Worksheet: " & shtType
End Sub

Private Function IsExistJobList() As Boolean
    On Error GoTo catch
    mWorkSheetUniq.Item ("Job List")
    On Error GoTo 0
    IsExistJobList = True
    Exit Function
catch:
    Err.Clear
    IsExistJobList = False
End Function

Private Function IsExistSheetManager() As Boolean
    On Error GoTo catch
    mWorkSheetCostom.Item ("Data Sheet Manager")
    On Error GoTo 0
    IsExistSheetManager = True
    Exit Function
catch:
    Err.Clear
    IsExistSheetManager = False
End Function


Private Sub specifyJobListManagingSheet()
    Dim context As Collection
    Set context = New Collection
    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Range("A1") Like "DFF*" Then
            Dim shtType As Variant
            For Each shtType In Split(mJOBLIST_MANAGING_SHEETS, ",")
                If sht.Range("B1") = shtType Then

                    On Error GoTo catch
                    Call context.Add(sht.Name, shtType & "@Sheet Parameters")
                    On Error GoTo 0

                End If
            Next shtType
        End If
    Next sht
    Call mWorkSheetOfJobContext.Add(context, ThisWorkbook.Name)
    Set sht = Nothing
    Set context = Nothing
    Exit Sub
catch:
    TheError.Raise 9999, TypeName(Me) & ".specifyJobListManagingSheet", "Too many Worksheet: " & shtType
End Sub

Private Sub specifySheetManagerManagingSheet()
    Dim context As Collection
    Set context = New Collection
    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Worksheets
        Dim shtType As Variant
        For Each shtType In Split(mSHEETMANAGER_MANAGING_SHEETS, ",")
            If sht.Range("B1") = shtType Then

                On Error GoTo catch
                Call context.Add(sht.Name, shtType & "@Sheet Parameters")
                On Error GoTo 0

            End If
        Next shtType
    Next sht
    Call mWorkSheetCostom.Add(context, ThisWorkbook.Name)
    Set sht = Nothing
    Set context = Nothing
    Exit Sub
catch:
    TheError.Raise 9999, TypeName(Me) & ".specifySheetManagerManagingSheet", "Too many Worksheet: " & shtType
End Sub

Private Sub readSheetContext(ByVal pPrimaryKey As String, ByVal pShtName As String, ByRef pContext As Collection)
    Dim shtReader As CWorkSheetReader
    Set shtReader = mReaderFactory.CreateReader(ThisWorkbook, pShtName)
    
    Dim fileStrm As IFileStream
    Set fileStrm = shtReader
    Do While fileStrm.IsEOR <> True
        Dim paramReader As IParameterReader
        Set paramReader = shtReader
        
        Dim primaryKeyValue As String
        Dim shtContext As Collection
        Set shtContext = New Collection
        
        Dim Label As Variant
        For Each Label In shtReader.sheetInfo.Keys
            
            If Label = pPrimaryKey Then
                primaryKeyValue = paramReader.ReadAsString(Label)
            ElseIf Label Like "*@Sheet Parameters" Then
                Call shtContext.Add(paramReader.ReadAsString(Label), Label)
            Else
                'Skip
            End If
            
        Next Label
        
        Call pContext.Add(shtContext, primaryKeyValue)
        Set shtContext = Nothing
        primaryKeyValue = vbNullString
        Set paramReader = Nothing
        
        Call fileStrm.MoveNext
    Loop

    Set fileStrm = Nothing
    
    Call mReaders.Add(shtReader, pShtName)
    Set shtReader = Nothing
    
End Sub



Private Sub Class_Terminate()
    Set mReaders = Nothing
        
    Set mWorkSheetCostom = Nothing
    Set mWorkSheetUniq = Nothing
    Set mWorkSheetOfTestContext = Nothing
    Set mWorkSheetOfJobContext = Nothing
    
    Set mReaderFactory = Nothing
End Sub
