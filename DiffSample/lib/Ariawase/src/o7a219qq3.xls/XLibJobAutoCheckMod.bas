Attribute VB_Name = "XLibJobAutoCheckMod"
Option Explicit
'{
'厚木自動検証モード判定用FLAG  0:NOT自動検証、1:自動検証 (とりあえず、ココに入れているが後で見直し)
Public Flg_JobAutoCheck As Integer

Private Const IDP_EXTENSION_NAME = ".idp" '画像ファイルの拡張子

'外部ファイル読込先ラベル名称定義用
Private Const JOBAUTOCHK_SHEETNAME = "JobAutoCheck"           'ワークシートの名前
Private Const LABEL_IDP_LOAD_PATH = "IDP_LOAD_PATH"           '画像ファイルのラベル
Private Const LABEL_SNAPSHOT_FILENAME = "SNAPSHOT_FILE_NAME"  'スナップショットのラベル
Private Const LABEL_COLUMN = 1                                'ラベル記入行
Private Const VALUE_COLUMN = 2                                '値記入行

'スナップショット取得用変数
Private sampleNumber As Long
Private snapFileName As String        'テスタ情報スナップショット保存用ファイル名

'外部ファイルの保存先情報用変数
Private idpLoadPath As String          '取り込み画像、ADCEOEFの保存先
Private snapShotLogSavePath As String  'スナップショットを外部ファイルに出力するときの出力

'取り込み画像の保存先のPath情報を公開
Public Function GetIdpPath() As String
    GetIdpPath = idpLoadPath
End Function

'ユーザーが入力したサンプル番号値を公開
Public Function GetSampleNumber() As Long
    GetSampleNumber = sampleNumber
End Function

'JobAutoCheckスタート設定
Public Sub InitJobAutoCheck(ByVal MinChipNumber As Long, ByVal Max_ChipNumber As Long)
    
'テスタモードを取得して、実機かシミュレータかを確認する。
'testModeOffline:シミュレータ、testModeOnline：テスター実機
    If TheExec.TesterMode = testModeOffline Then

        Flg_JobAutoCheck = 1 '暫定施策 これはあとで何とかしたい
            
        '検証に使用する外部ファイルの定義をシートから取得
        Call mf_GetFileLoadPath
            
        Dim inputChipNumber As Variant
        inputChipNumber = 0
        'インプットボックスに、使用するダミー画像のサンプル番号を入力。
        Do
            inputChipNumber = InputBox("Enter CHIP number (" & MinChipNumber & "-" & Max_ChipNumber & "): ", "CHIP Number Input")
        Loop While (inputChipNumber <= 0 Or inputChipNumber > Max_ChipNumber Or inputChipNumber < MinChipNumber)
        
        sampleNumber = CLng(inputChipNumber)
                         
        If mf_Get_JAC_SheetVal("TESTER_SNAPSHOT_SAVE") Then
            Call mf_deleteFile(snapFileName)    'スナップショット保存用ファイルをお掃除
        End If
    Else
        Flg_JobAutoCheck = 0 '暫定施策 これはあとで何とかしたい
    End If

End Sub

'Measure時になにかやりたい時用(今は、テスタSnapのみ)
Public Sub CheckMeasureStatus(Optional idLabel As String = "")
    DoEvents
    Call mf_saveSnapShot(idLabel)  'テスタスナップショットを取得させる
End Sub

'ダミー画像のファイル名をフルパス指定で返します
Public Sub GetIdpFileName(ByVal siteNum As Long, ByVal TestLabel As String, ByRef IdpFilePath As String)
    IdpFilePath = idpLoadPath & sampleNumber & "\" & TestLabel & "_" & siteNum & IDP_EXTENSION_NAME
End Sub

'イメージプレーンのデータをファイルから読み込みます
Public Sub ReadIdpFile(InputPlaneName As String, _
    BasePmdName As String, _
    IdpFileName As String, _
    Optional InputSiteNumber As Long = ALL_SITE, _
    Optional IdpFileType As IdpFileFormat = idpFileBinary)

    Dim idpLogMsg As String

    TheHdw.IDP.SetPMD InputPlaneName, BasePmdName
    TheHdw.IDP.ReadFile InputSiteNumber, InputPlaneName, idpColorFlat, IdpFileName, IdpFileType
    TheHdw.IDP.SetPMD InputPlaneName, BasePmdName

    '画像ファイル読み込み時のログ出力用
    #If IDP_READ_LOG = 1 Then
        idpLogMsg = "IDP_READ," & "Instances=" & _
        TheExec.DataManager.InstanceName & _
        ",Plane=" & InputPlaneName & _
        ",Site=" & InputSiteNumber & _
        ",File=" & IdpFileName
        Call WriteComment(idpLogMsg)
    #End If

End Sub

'データログwindowに出力する、Debug情報を規定のFormat（自動比較用）で出力するためのサブルーチン 結果用
Public Sub OutputDebugInfo(ByVal testCategory As String, _
    ByVal testName As String, _
    ByVal SiteNumber As Long, _
    ByVal OutputValue As Double, _
    Optional UnitLabel As String = "")

    Dim outPutMsg As String

    outPutMsg = "#" & testCategory & ":" & testName & ":" & SiteNumber & ":" & " = " & OutputValue & "" & UnitLabel

    Call WriteComment(outPutMsg)

End Sub

'データログwindowに、光源の設定状態を出力するためのサブルーチン
Public Sub OutputOptsetInfo(ByVal category As String, ByVal TestInstanceName As String, ByVal CommandListIdentifier As String)
    
    Dim outPutMsg As String
    
    outPutMsg = "#" & category & ":" & TestInstanceName & ":" & " = " & CommandListIdentifier
    
    Call WriteComment(outPutMsg)

End Sub

'自動検証用APMU SnapShot実行用
Private Sub mf_saveSnapShot(Optional idLabel As String = "")
    
    If mf_Get_JAC_SheetVal("TESTER_SNAPSHOT_SAVE") Then
        'テスタスナップショットを取得し、結果をファイルに保存する
        Call TheSnapshot.GetSnapshot(idLabel)
    Else
'        MsgBox ("スナップショットを保存しろと言われたが" & "TESTER_SNAPSHOT_SAVE" & "が1ではないので何もしない")
    End If

End Sub

'JobAutoCheckシートの指定ラベルの隣の値を取得してくる
Private Function mf_Get_JAC_SheetVal(ByVal LabelName As String) As Long
    mf_Get_JAC_SheetVal = mf_GetCellValue(JOBAUTOCHK_SHEETNAME, mf_SearchWordRow(LabelName), VALUE_COLUMN)
End Function

'対象ワークシートの存在確認
Private Function mf_ChkJacSheet() As Integer

    Dim sheetChkFlg As Boolean
    Dim sheetCnt As Long

    For sheetCnt = 1 To ThisWorkbook.Worksheets.Count
        ' シートが見つかったらフラグ設定
        If ThisWorkbook.Worksheets(sheetCnt).Name = JOBAUTOCHK_SHEETNAME Then
            sheetChkFlg = True
            mf_ChkJacSheet = 0
        End If
    Next
    
    'シートが見つからなかった場合の処理
    If sheetChkFlg = False Then
        MsgBox JOBAUTOCHK_SHEETNAME & " ワークシートが存在しません", vbCritical, "SHEET CHECK ERROR"
        mf_ChkJacSheet = 1
        Stop
    End If

End Function

'指定キーワードが存在するCellの行番号を取得
Private Function mf_SearchWordRow(ByVal searchWord As String) As Long
    
    Dim tmpRange As Range
    
    With Worksheets(JOBAUTOCHK_SHEETNAME).Columns(LABEL_COLUMN)
        
        Set tmpRange = .Find(searchWord)
        
        If Not tmpRange Is Nothing Then
           'MsgBox "検索したキーワードが存在するのは" & tmpRange.Cells.Row & "行目です"
            mf_SearchWordRow = tmpRange.Cells.Row
        Else
            MsgBox JOBAUTOCHK_SHEETNAME & "シートのA列に" & searchWord & "が存在しません", vbCritical, "SHEET SEARCH ERROR"
            Stop
        End If
        
    End With
    
    Set tmpRange = Nothing

End Function

'指定シートの指定CELLの値を入手
Private Function mf_GetCellValue(ByVal workSheetName As String, ByVal cellsColumn As Long, ByVal cellsRow As Long) As Variant
    mf_GetCellValue = Worksheets(workSheetName).Cells(cellsColumn, cellsRow)
End Function

'JobAutoCheckワークシートのファイルの保存先値を取得し変数に格納
Private Sub mf_GetFileLoadPath()

    Dim idpRow As Long
    Dim snapRow As Long
    
    Call mf_ChkJacSheet
    
    idpRow = mf_SearchWordRow(LABEL_IDP_LOAD_PATH)     'ダミー画像の保存先の定義行
    snapRow = mf_SearchWordRow(LABEL_SNAPSHOT_FILENAME) 'スナップショット保存ファイルの定義行
        
    idpLoadPath = mf_GetCellValue(JOBAUTOCHK_SHEETNAME, idpRow, VALUE_COLUMN)
    snapFileName = mf_GetCellValue(JOBAUTOCHK_SHEETNAME, snapRow, VALUE_COLUMN)

End Sub

'スナップショットのログをお掃除する
Private Sub mf_deleteFile(ByVal delFileName As String)
    
    On Error GoTo FileDelErr
    
    Call Kill(delFileName)
    Exit Sub

FileDelErr:
'    MsgBox "指定されたファイルはなかったのでお掃除しなくて済んだ"

End Sub

'スナップショット用のログをファイルに出力する。
Private Sub mf_OutPutLog(ByVal LogFileName As String, outPutMessage As String)
    Dim fp As Integer
    On Error GoTo OUT_PUT_LOG_ERR
    
    fp = FreeFile
    Open LogFileName For Append As fp
    Print #fp, outPutMessage
    Close fp
    
    Exit Sub

OUT_PUT_LOG_ERR:
    Call MsgBox(LogFileName & " MsgOutPut Error", vbFalse Or vbCritical, "@mf_OutPutLog")
    Stop

End Sub

'スナップショットの保存先を教える
Public Function GetSnapFilename() As String
    
'    If snapFileName = "" Then
        Call mf_GetFileLoadPath
'    End If
    
    GetSnapFilename = snapFileName

End Function

'スナップショット取得フラグの確認
Public Function IsSnapshotOn() As Boolean

    If mf_Get_JAC_SheetVal("TESTER_SNAPSHOT_SAVE") <> 0 Then
        IsSnapshotOn = True
    Else
        IsSnapshotOn = False
    End If

End Function

'}

