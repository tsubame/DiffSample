Attribute VB_Name = "XLibJob"
'概要:
'   共通使用するObjectの生成と初期化
'
'       Revision History:
'           Date        Description
'           2013/6/11   TheDcTestの型をCDcScenario=>IDcScenarioに変更(0145184306)
'
'目的:
'   共通使用するObjectの生成と初期化の処理をまとめる
'   正しい手順、順序で各Objectの生成、初期化を実行する。
'
'作成者:
'   SLSI今手
'

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const CUB_UB_USE = 0    'CUB UBの設定          0：未使用、0以外：使用

'Public Object群
Public TheUB As New CUtyBitController 'UB設定Object
Public TheDC As New CVISVISrcSelector '電源設定Object
Public TheSnapshot As New CSnapIP750  'スナップショット機能Object

Public TheDcTest As IDcScenario
Public TheOffsetResult As COffsetManager

Dim mDataManagerReader As CDataSheetManager
Dim mJobListReader As CDataSheetManager
Dim mDcScenarioReader As CDcScenarioSheetReader
Dim mDcScenarioWriter As CDcScenarioSheetLogWriter
Dim mDcReplayDataReader As CDcPlaybackSheetReader
Dim mInstanceReader As CInstanceSheetReader
Dim mOffsetReader As COffsetSheetReader

Dim mDcLogReportWriter As CDcLogReportWriter

#If CUB_UB_USE <> 0 Then
Public CUBUtilBit As New CUBUtilityBits.UtilityBits 'CUB UB設定用Object
#End If

Public Sub InitJob()
'内容:
'   JOBの初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'

    Call InitTheDC       'DCボードセレクタの初期化
    Call InitTheUB       'UBコントローラの初期化
    Call InitTheSnapshot 'スナップショット機能の初期化

End Sub

Public Sub InitTheSnapshot()
'内容:
'   スナップショット機能の初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    With TheSnapshot
        .Initialize                      'スナップショット機能の初期化
        .LogFileName = GetSnapFilename   '外部TXT出力ファイル名
        .OutputPlace = snapTXT_FILE      '取得結果出力先
        .OutputSaveStatus = True         'スナップショット機能の動作状況をデータログに出力
        .SerialNumber = 1                'ログに出力するシリアル番号の初期値
    End With

End Sub

#If CUB_UB_USE <> 0 Then
Public Sub InitCub()
'内容:
'   CUB UBの初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   利用時には条件付コンパイル引数に CUB_UB_USE=1の記載が必要です。
'   CUB UB使用時には本初期化作業をテスターイニシャル時に実行する必要があります。
'
    
    With CUBUtilBit
        .SetTheHdw TheHdw
        .SetTheExec TheExec
        .Clear
    End With
    
    TheHdw.DIB.LeavePowerOn = True

End Sub
#End If

Public Sub InitTheDC()
'内容:
'   電源設定ボードセレクタの初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    
    Call TheDC.Initialize

End Sub

Public Sub InitTheUB()
'内容:
'   'UBコントローラの初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   CUBのUB利用時には条件付コンパイル引数に CUB_UB_USE=1の記載が必要です。
'

    '######### 初期化 #########################################
    Call TheUB.Initialize

    '######### 設定 #########################################
    'APMU UB
    With TheUB.AsAPMU
        Set .UBSetSht = Worksheets("APMU UB") '条件表のシート指定
        Call .LoadCondition                       '条件表のLoad
    End With

    #If CUB_UB_USE <> 0 Then
    'CUB UB
    With TheUB.AsCUB
        Set .UBSetSht = Worksheets("CUB UB") '条件表のシート指定
        Call .LoadCondition                      '条件表のLoad
    End With
    #End If

End Sub

Public Sub InitControlShtReader()
'内容:
'   JOBリスト等のコントロールシートリーダーを初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   アクティブなデータシート指定が変更された場合と
'   テラダインデータツール上のデータ変更された場合の
'   バリデーションの際に行う
'
    Set mDataManagerReader = Nothing
    Set mJobListReader = Nothing
End Sub

Public Sub InitTestScenario()
'内容:
'   各オブジェクトの初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
'
    On Error GoTo ErrHandler
    '### 測定結果マネージャ初期化 #########################
    XLibResultManagerUtility.InitResult
    '### 各オブジェクトのの初期化 #########################
    If mDataManagerReader Is Nothing Or mJobListReader Is Nothing Then
        InitActiveDataSheet
        InitTheDcScenario
        InitTheOffsetResult
    Else
        ReInitTheDcScenario
    End If
    '### 測定結果のクリア #################################
    TheDcTest.ClearContainer
    TheDcTest.ResultManager = TheResult
        
    '### DcLoopOption設定
    Call XLibDcScenarioLoopOption.ApplyDcScenarioLoopOptionMode

    Exit Sub
ErrHandler:
    InitControlShtReader
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    DisableAllTest
End Sub

Private Sub InitActiveDataSheet()
    '### JOBリストマネージャの準備 ########################
    Set mJobListReader = CreateCDataSheetManager
    mJobListReader.Initialize JOB_LIST_TOOL
    '### アクティブテストインスタンスシートの取得 #########
    Set mInstanceReader = CreateCInstanceSheetReader
    mInstanceReader.Initialize mJobListReader.GetActiveDataSht(TEST_INSTANCES_TOOL).Name
    '### シートマネージャの準備 ###########################
    Set mDataManagerReader = CreateCDataSheetManager
    mDataManagerReader.Initialize SHEET_MANAGER_TOOL
    '### アクティブDCシナリオシートの取得 #################
    Dim activeSht As Worksheet
    Set activeSht = mDataManagerReader.GetActiveDataSht(DC_SCENARIO_TOOL)
    '### アクティブDCシナリオリーダーの準備 ###############
    If activeSht Is Nothing Then
        Set mDcScenarioReader = Nothing
        Set mDcScenarioWriter = Nothing
    Else
        Set mDcScenarioReader = CreateCDcScenarioSheetReader
        mDcScenarioReader.Initialize activeSht.Name
        If mDcScenarioReader.AsIParameterReader.ReadAsBoolean(IS_VALIDATE) Then
            Set mDcLogReportWriter = New CDcLogReportWriter
            If XLibDcScenarioLoopOption.DcLoop = False Then mDcLogReportWriter.Initialize
        Else
            Set mDcLogReportWriter = Nothing
        End If
        Set mDcScenarioWriter = CreateCDcScenarioSheetLogWriter
        mDcScenarioWriter.Initialize activeSht.Name, GetSiteCount
    End If
    '### アクティブDC再生データリーダーの準備 #############
    Set activeSht = mDataManagerReader.GetActiveDataSht(DC_PLAYBACK_TOOL)
    If activeSht Is Nothing Then
        Set mDcReplayDataReader = Nothing
    Else
        Set mDcReplayDataReader = CreateCDcPlaybackSheetReader
        mDcReplayDataReader.Initialize activeSht.Name, GetSiteCount
    End If
    '### オフセットリーダーの準備 #########################
    Set activeSht = mDataManagerReader.GetActiveDataSht(OFFSET_TOOL)
    If activeSht Is Nothing Then
        Set mOffsetReader = Nothing
    Else
        Set mOffsetReader = CreateCOffsetSheetReader
        mOffsetReader.Initialize activeSht.Name, GetTesterNum, GetSiteCount
    End If
End Sub

Private Sub InitTheDcScenario()
    '### DCテストシナリオ実行エンジンの初期化 #############
    If mDcScenarioReader Is Nothing Then
        Set TheDcTest = Nothing
    Else
        Dim dcPerformer As IDcTest
        If Not mDcReplayDataReader Is Nothing Then
            Dim replayDc As CPlaybackDc
            Set replayDc = CreateCPlaybackDc
            replayDc.Initialize mDcReplayDataReader
            Set dcPerformer = replayDc
        Else
            Set dcPerformer = CreateVISConnector '電源設定にVISクラスを使用する
'            Set dcPerformer = CreateCStdDCLibV01
        End If
        Set TheDcTest = CreateCDCScenario
        TheDcTest.Initialize dcPerformer, mDcScenarioReader, mInstanceReader, mDcScenarioWriter, mDcLogReportWriter
    End If
End Sub

Private Sub InitTheOffsetResult()
    '### オフセットマネージャの初期化 #####################
    If mOffsetReader Is Nothing Then
        Set TheOffsetResult = Nothing
    Else
        Set TheOffsetResult = CreateCOffsetManager
        TheOffsetResult.Initialize mOffsetReader
    End If
End Sub

Private Sub ReInitTheDcScenario()
    '### DCテストシナリオ実行エンジンの初期化 #############
    If mDcScenarioReader Is Nothing Then
        Set TheDcTest = Nothing
        Exit Sub
    Else
        With mDcScenarioReader
            If .AsIParameterReader.ReadAsBoolean(DATA_CHANGED) Then
                TheDcTest.Load
                mDcScenarioWriter.AsIActionStream.Rewind
            End If
            If Not mDcLogReportWriter Is Nothing Then
                If .AsIParameterReader.ReadAsBoolean(IS_VALIDATE) Then
                    If XLibDcScenarioLoopOption.DcLoop = False Then mDcLogReportWriter.Initialize
                Else
                    Set mDcLogReportWriter = Nothing
                End If
            End If
        End With
    End If
End Sub

Public Sub CloseDcLogReportWriter()
    If Not mDcLogReportWriter Is Nothing Then
        If XLibDcScenarioLoopOption.DcLoop = False Then mDcLogReportWriter.AsIFileStream.IsEOR
    End If
End Sub
