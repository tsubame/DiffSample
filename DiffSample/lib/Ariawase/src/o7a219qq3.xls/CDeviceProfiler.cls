VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDeviceProfiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   DeviceConfigurationsシートに設定した情報を取得＆保持するクラス
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfigurationsシートの情報を取得＆保持する機能を実装した
'       2010/12/07  不要コードを削除した
'
'作成者:
'   0145184346
'

Option Explicit

Private Const ERR_NUMBER As Long = 9999                ' Error番号を保持する
Private Const CLASS_NAME As String = "CDeviceProfiler" ' Class名称を保持する

Private Const PARAMETER_START_RANGE As String = "B5"   ' データ設定開始セルを保持する
Private Const PARAMETER_RANGE  As String = "B5:K5"     ' データ設定可能セルを保持する
Private Const PARAMETER_MODE_RANGE  As String = "H5"   ' Modeパラメータ設定セルを保持する
Private Const PARAMETER_START_ROW As Long = 5          ' データ設定開始セルを保持する
Private Const PARAMETER_MODE_COLUMN As Long = 8        ' Modeパラメータ設定セルを保持する
Private Const PARAMETER_MODE_TITLE As String = "Mode"  ' 入力規則のパラメータ名称を保持する
Private Const MODE_STATIC_WORD As String = "Static"    ' 入力規則のパラメータ値を保持する
Private Const MODE_DYNAMIC_WORD As String = "Dynamic"  ' 入力規則のパラメータ値を保持する

Private mConfigurations As Collection ' 全てのConfigデータを保持する

Public Sub Initialize()
'内容:
'   DeviceConfigurationsデータを取得して保持する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'

    '#####  SheetReaderを利用して、DeviceConfigurationsシートを読み込む  #####
    Dim clsWrkShtRdr As CWorkSheetReader
    Set clsWrkShtRdr = GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeDeviceConfigurations)
    
    Dim IFileStream As IFileStream
    Dim IParamReader As IParameterReader

    With clsWrkShtRdr
        Set IFileStream = .AsIFileStream
        Set IParamReader = .AsIParameterReader
    End With

    Set clsWrkShtRdr = Nothing


    '#####  DeviceConfigurations情報を読み出し＆格納する  #####
    Dim clsDevCfg As CDeviceConfiguration
    Set mConfigurations = New Collection

    Do While Not IFileStream.IsEOR
        On Error GoTo ErrHandler
        Set clsDevCfg = New CDeviceConfiguration
        Call clsDevCfg.Initialize(IParamReader)
        On Error GoTo ErrHandlerInstanceName
        mConfigurations.Add clsDevCfg, clsDevCfg.InstanceName
        On Error GoTo 0
        IFileStream.MoveNext
    Loop

    Set clsDevCfg = Nothing
    Set IFileStream = Nothing
    Set IParamReader = Nothing

    '#####  終了  #####
    Exit Sub


ErrHandler:
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Sub
ErrHandlerInstanceName:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Already Set Instance Name.[" & clsDevCfg.InstanceName & "]")
    Exit Sub

End Sub

Public Property Get Configuration(ByRef InstanceName As String) As CDeviceConfiguration
'内容:
'   指定したInstance名称のDeviceConfigurationデータを戻す
'
'パラメータ:
'   [InstanceName]  In  Instance名称を保持する
'
'戻り値:
'   DeviceConfigurationデータ
'
'注意事項:
'

    '#####  指定したInstance名称のConfiguration情報を取得する  #####
    On Error GoTo ErrHandler
    Set Configuration = mConfigurations.Item(InstanceName)
    Exit Property

ErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Configuration", "Illegal TestInstance Name.")
    Exit Property
    
End Property

Public Property Get ConfigInfo(ByRef InstanceName As String) As CDeviceConfigInfo
'内容:
'   指定したInstance名称のDeviceConfigInfoデータを戻す
'
'パラメータ:
'   [InstanceName]  In  Instance名称を保持する
'
'戻り値:
'   DeviceConfigInfoデータ
'
'注意事項:
'

    '#####  指定したInstance名称のConfigInfo情報のみを取得する  #####
    Dim clsDevCfg As CDeviceConfiguration
    On Error GoTo ErrHandler
    Set clsDevCfg = mConfigurations.Item(InstanceName)
    On Error GoTo 0
    Set ConfigInfo = clsDevCfg.ConfigInfo
    Set clsDevCfg = Nothing
    Exit Property

ErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".ConfigInfo", "Illegal TestInstance Name.")
    Exit Property
    
End Property

Public Sub SelectionCellForm(ByRef rngTarget As Range)
'内容:
'   Cellに入力規則を付加する
'
'パラメータ:
'   [rngTarget]  In  選択されたセル情報
'
'戻り値:
'
'注意事項:
'

    '#####  入力規則をリスト化して設定する  #####
    Const SET_VALUE_LIST As String = MODE_STATIC_WORD & "," & MODE_DYNAMIC_WORD
    With rngTarget
        If .Column = PARAMETER_MODE_COLUMN And .Row >= PARAMETER_START_ROW And .Count = 1 Then
            Call addValidationCell(rngTarget, PARAMETER_MODE_TITLE, SET_VALUE_LIST)
        End If
    End With

End Sub

Private Sub addValidationCell(ByRef pTargetRange As Range, ByVal pDataName As String, _
                            ByVal pDataItemList As String, Optional ByVal pShowInputMessage As Boolean = False)
    
    '#####  一度現在の入力規則を削除  #####
    Call pTargetRange.Validation.Delete
    
    '#####  入力規則を新規追加  #####
    Call pTargetRange.Validation.Add(Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                                    operator:=xlEqual, Formula1:=pDataItemList)
    
    '#####  表示メッセージ等の設定  #####
    With pTargetRange.Validation
        .InputTitle = pDataName                     ' Balloonのタイトル値
        .InputMessage = "Please select"             ' Balloonのメッセージ値
        .ShowInput = pShowInputMessage              ' Balloonの表示設定
        .InCellDropdown = True                      ' Dropdownの表示設定
        .ErrorTitle = "Unknown " & pDataName & " !" ' Error Msgbox タイトル
        .ErrorMessage = "Input Value is unknown !"  ' Error Msgbox メッセージ
        .IgnoreBlank = False                        ' ブランク入力許容設定
        .IMEMode = xlIMEModeNoControl               ' IMEモード設定
        .ShowError = True
    End With

End Sub

Public Sub ChangeCellForm(ByRef strWorkSheetName As String)
'内容:
'   Cellに網掛け処理を実施する
'
'パラメータ:
'   [strWorkSheetName]  In  使用するシート名
'
'戻り値:
'
'注意事項:
'

    '#####  Sheetを取得  #####
    Dim shtDevCfg As Worksheet
    On Error GoTo ErrHandler
    Set shtDevCfg = Worksheets(strWorkSheetName)
    On Error GoTo 0

    '#####  網掛け処理を実施  #####
    Application.ScreenUpdating = False
    Call ClearDataSheetFormat(shtDevCfg)
    Call setDataSheetFormat(shtDevCfg)
    Application.ScreenUpdating = True
           
    '#####  終了  #####
    Set shtDevCfg = Nothing
    Exit Sub
        
ErrHandler:
    Call errorRaise("DeviceConfigurations WorkSheet is Nothing")
    Exit Sub

End Sub

Private Sub setDataSheetFormat(ByRef shtWork As Worksheet)
    Dim rngCondition As Range ' データ開始セル
    Dim rngData As Range      ' Modeパラメータ設定セル
    Dim rngTarget As Range    ' 変更先セル
    Set rngCondition = shtWork.Range(PARAMETER_START_RANGE)
    Set rngData = shtWork.Range(PARAMETER_MODE_RANGE)
    While IsEmpty(rngCondition) = False
        If rngData.Value = MODE_STATIC_WORD Then
            Set rngTarget = rngData.offset(0, 2)
        ElseIf rngData.Value = MODE_DYNAMIC_WORD Then
            Set rngTarget = rngData.offset(0, 1)
        Else
            Set rngTarget = Nothing
        End If
        If Not (rngTarget Is Nothing) Then
            With rngTarget.Interior
                .Pattern = xlGray8
                .ColorIndex = 15
            End With
        End If
        Set rngData = rngData.offset(1, 0)
        Set rngCondition = rngCondition.offset(1, 0)
    Wend
    Set rngCondition = Nothing
    Set rngData = Nothing
    Set rngTarget = Nothing
End Sub

Private Sub ClearDataSheetFormat(ByRef shtWork As Worksheet)
    Dim rngCondition As Range ' データ開始セル
    Dim rngData As Range      ' データ設定可能セル
    Set rngCondition = shtWork.Range(PARAMETER_START_RANGE)
    Set rngData = shtWork.Range(PARAMETER_RANGE)
    While IsEmpty(rngCondition) = False
        With rngData.Interior
            .Pattern = xlSolid
            .ColorIndex = xlNone
        End With
        Set rngData = rngData.offset(1, 0)
        Set rngCondition = rngCondition.offset(1, 0)
    Wend
    Set rngCondition = Nothing
    Set rngData = Nothing
End Sub

Private Sub errorRaise(ByVal pErrMsg As String)
    If Not (TheError Is Nothing) Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME, pErrMsg)
    Else
        Call Err.Raise(ERR_NUMBER, CLASS_NAME, pErrMsg)
    End If
End Sub

Private Sub Class_Initialize()
    Set mConfigurations = Nothing
End Sub

Private Sub Class_Terminate()
    Set mConfigurations = Nothing
End Sub
