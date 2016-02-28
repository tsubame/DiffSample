VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestConditionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   測定条件設定用クラス
'
'目的:
'   ワークシートの定義値を使用した測定条件設定機能提供
'   セット関数の直接実行機能
'   測定条件設定履歴情報の記録と公開
'
'作成者:
'   SLSI今手
'
'注意点:
'
Option Explicit

'条件表ワークシートに関する定義
Private Const CONDITION_SHEET_START_CELL = "B5"          '条件表ワークシートIndex のスタートセル定義(1番最初のCondition Nameのデータのある場所を指定)
Private Const SHEET_TITLE = "TEST Condition"             '測定条件表のタイトル文字列
Private Const SHEET_TITLE_CELL = "B1"                    '測定条件表のタイトルの場所
Private Const CONDITION_TITLE = "Condition Name"         'ConditionNameセルのタイトル文字列
Private Const CONDITION_TITLE_CELL = "B3"                'ConditionName 見出しセルの場所
Private Const SET_FUNCTION_TITLE = "Procedure Name"      'Set Function Nameセルのタイトル文字列
Private Const SET_FUNCTION_TITLE_CELL = "C3"             'Set Function Name 見出しセルの場所
Private Const ARG_RANGE = "D:M"                          'Argの範囲指定 グループ化用
Private Const ARG_START_COLUMN = 4                       'Argの開始列
Private Const ARG_END_COLUMN = 13                        'Argの終了列
Private Const DATA_START_RANGE As String = "B5:M5"       '表のデータ開始範囲行

'マクログループ定義用
Private Const EXECUTE_SETTINGLIST As String = "C2"              'マクログループ選択セルの位置
Private Const EXECUTE_SETTINGLIST_ROW As Long = 2               'マクログループ選択セルのRow位置
Private Const EXECUTE_SETTINGLIST_COL As Long = 3               'マクログループ選択セルのColumn位置 C=4
Private Const MACRO_SETTING_DEFAULT_TITLE As String = "Default" '条件設定マクロ定義ラベルのタイトル文字列
Private Const MACRO_SETTING_DEFAULT As String = "N4"            '条件設定マクロ有効/無効の開始位置
Private Const MACRO_SETTING_START_COLUMN As Long = 14           '条件設定マクロ有効/無効の開始Column位値 N=14
Private Const MACRO_SETTING_START_ROW As Long = 4               '条件設定マクロ有効/無効の開始Row位値
Private Const MACRO_SETTING_END_COLUMN As Long = 18             '条件設定マクロ有効/無効の終了Column位値 R=18

'内部情報用
Private Const ERR_NUMBER As Long = 9999                     'エラー時に渡すエラー番号
Private Const MAX_ARG_NUMBER As Long = 10                   'セット関数のパラメータの最大要素数
Private Const PARAMETER_END_LABEL1 As String = "#EOP"       'パラメータの終端を示すラベル その1
Private Const PARAMETER_END_LABEL2 As String = "#DATAEND"   'パラメータの終端を示すラベル その2
Private Const MACRO_ENABLE_WORD As String = "enable"        'マクロ有効設定 Word
Private Const MACRO_DISABLE_WORD As String = "disable"      'マクロ無効設定 Word
Private Const CONDITION_TO_ARG_OFFSET As Long = 2           '条件名からArgまでのOffset値
Private Const PROCEDURE_TO_ARG_OFFSET As Long = 1           'マクロ名セルからARG定義セルまでのOffset値
Private Const SHEET_DATA_START_COLUMN As Long = 2           '測定条件表データ開始Column位置    B=2
Private Const SHEET_DATA_START_ROW As Long = 5              '測定条件表データ開始Row位置
Private Const SHEET_DATA_END_COLUMN As Long = 14            '測定条件表データ終了Column位置    N=14

Private WithEvents m_TestConditionSheet As Worksheet        '測定条件表ワークシートObject(このシートのイベントも取得)
Attribute m_TestConditionSheet.VB_VarHelpID = -1
Private m_Conditions As Collection                          '測定条件表のキャッシュ用コレクション
Private m_ConditionNameList As Collection                   'Condition名リスト
Private m_HistoryRecordFlag As Boolean                      '設定履歴保存機能有効フラグ
Private m_TermCharacters() As String                        '終端文字列として取り扱うデータの詰め合わせ
Private m_JobWorkBookName As String                         '使用しているJOBの名前（セット関数があるJOBの名前）
Private m_SetConditionLogger As CSetConditionLogger         '測定条件設定LoggerObject
Private m_SettingNameColList As Collection                  'Column位置のリスト(KeyはSettingName)
Private m_ExecSettingDataList As Collection                 'マクログループ定義の有効無効リスト
Private m_SelecteExecSetting As String                      '選択されたマクログループ定義
Private m_BackUpGroupName As String                         '変更前のマクログループ定義

Public Sub ModifyConditionSheet(Optional ByVal pIsUnFormatMode As Boolean = False)
'内容:
'   設定されている測定条件表を整形する
'
'パラメータ:
'   [pIsUnFormatMode]  In  Boolean型： 表を未整形状態にするかどうかの指定(オプション)
'
'戻り値:
'
'注意事項:
'   ユーザーによる、すべてのグループ化が
'   本命令の実行により解除されます。
'

    Dim ConditionRng As Range
    Dim ArgRng As Range
    Dim ArgCnt As Long
    Dim ModifyCnt As Long
    Dim ConditionName As String
    Dim GroupStartCnt As Long
    
    Const COLOR_INDEX_GRAY As Long = 15 '灰色
    Const OUTLINE_LEVEL2 As Long = 2 'グループ化されている
    Const OUTLINE_LEVEL1 As Long = 1 'グループ化されていない
    
    'Argセルをグループ化するフラグ
    Dim GroupFlag As Boolean
    GroupFlag = False
    
    'グループ化の範囲チェック
    Dim CurArgCell As Long
    For CurArgCell = ARG_START_COLUMN To MACRO_SETTING_END_COLUMN Step 1
    
        'Arg0~Arg9までの範囲がGroup化されているかチェック
        If CurArgCell <= ARG_END_COLUMN And _
           m_TestConditionSheet.Columns(CurArgCell).OutlineLevel <> OUTLINE_LEVEL2 Then
            GroupFlag = True
            Exit For
        End If
        
        'ExecuteSettingListのDefault行～最終列までがグループ化されていないかチェック
        If CurArgCell > ARG_END_COLUMN And _
           m_TestConditionSheet.Columns(CurArgCell).OutlineLevel <> OUTLINE_LEVEL1 Then
            GroupFlag = True
            Exit For
        End If
           
    Next CurArgCell
    
    If GroupFlag = True Then
        'すべてのアウトラインの削除
        m_TestConditionSheet.Cells.ClearOutline
    End If
                                                
    '表の書式全て初期化
    Call ClearDataSheetFormat

    If pIsUnFormatMode = False Then
        
        'マクロ有効無効による書式設定処理
        Call MakeDatasheetFormat
                
        '初期値の設定
        Set ConditionRng = m_TestConditionSheet.Range(CONDITION_SHEET_START_CELL)
        Set ArgRng = ConditionRng.offset(0, CONDITION_TO_ARG_OFFSET)
        With ConditionRng
            ConditionName = .Value
            GroupStartCnt = .Row + 1
        End With
    
        'Condition Index LOOP
        While Not IsEmpty(ConditionRng)
            '○Argの網掛け処理 LOOP
            For ArgCnt = 1 To MAX_ARG_NUMBER Step 1
                If IsTermCharacter(ArgRng.Value) Then
                    For ModifyCnt = ArgCnt To MAX_ARG_NUMBER Step 1
                        With ArgRng.Interior
                            .Pattern = xlGray8
                            .ColorIndex = COLOR_INDEX_GRAY
                        End With
                        Set ArgRng = ArgRng.offset(0, 1)
                    Next ModifyCnt
                    Exit For
                End If
                'ひとつ右へ
                Set ArgRng = ArgRng.offset(0, 1)
            Next ArgCnt
    '==================================================================================================
    '        '○条件別グループ化処理(かえって見にくい予感なので封印)
    '        If (ConditionRng.Value <> ConditionName) Then
    '            If (((ConditionRng.Row - 1) - GroupStartCnt) >= 0) Then
    '                '条件名でグループ化
    '                Call m_TestConditionSheet.Range(GroupStartCnt & ":" & ConditionRng.Row - 1).Group
    '            End If
    '            With ConditionRng
    '                ConditionName = .Value
    '                GroupStartCnt = .Row + 1
    '            End With
    '        End If
    '==================================================================================================
            'インデックスをひとつ下へ
            Set ConditionRng = ConditionRng.offset(1, 0)
            Set ArgRng = ConditionRng.offset(0, CONDITION_TO_ARG_OFFSET)
        Wend
    '==================================================================================================
    '    '○最終ライン部のグループ化処理(かえって見にくい予感なので封印)
    '    If (ConditionRng.Value <> ConditionName) Then
    '        If (((ConditionRng.Row - 1) - GroupStartCnt) >= 0) Then
    '            '条件名でグループ化
    '            Call m_TestConditionSheet.Range(GroupStartCnt & ":" & ConditionRng.Row - 1).Group
    '        End If
    '    End If
    '==================================================================================================
    End If
    
    If GroupFlag = True Then
        '○アウトラインによるArg部のグループ化処理
        Call m_TestConditionSheet.Range(ARG_RANGE).group
    End If

End Sub

Public Sub Initialize()
'内容:
'   初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   すべてのデータが初期化されます
'
    Set m_TestConditionSheet = Nothing
    Set m_Conditions = Nothing
    Set m_ConditionNameList = Nothing
    Set m_SettingNameColList = Nothing
    Set m_ExecSettingDataList = Nothing
    Set m_SetConditionLogger = New CSetConditionLogger
    Call m_SetConditionLogger.Initialize
    m_SetConditionLogger.MaxArgNumber = MAX_ARG_NUMBER - 1
    m_HistoryRecordFlag = False
    m_JobWorkBookName = ""
    m_SelecteExecSetting = ""
    m_BackUpGroupName = ""
    
    '終端文字ルールの作成（あまり増やさないので今はべた書き）
    'ルールをメンテするときはココを編集する必要が有る
    ReDim m_TermCharacters(1)
    m_TermCharacters(0) = PARAMETER_END_LABEL1  '終端文字データの詰め込み1
    m_TermCharacters(1) = PARAMETER_END_LABEL2  '終端文字データの詰め込み2
        
End Sub

Public Property Let JobWorkbookName(ByVal pWorkbookName As String)
'内容:
'   JOBのExcelWorkbook名の設定
'
'パラメータ:
'   [pWorkbookName]  In  String型： 使用するJOB Workbook名
'
'戻り値:
'
'注意事項:
'
    m_JobWorkBookName = pWorkbookName
End Property

Public Property Get JobWorkbookName() As String
'内容:
'   JOBのExcelWorkbook名
'
'パラメータ:
'
'戻り値:
'   JOBのExcelWorkbook名
'
'注意事項:
'
    
    JobWorkbookName = m_JobWorkBookName

End Property

Public Property Let TestConditionSheet(pWorksheetName As String)
'内容:
'   TestConditionワークシート
'
'パラメータ:
'   [pWorkSheetName]  In  String型： 使用するTestConditionワークシート
'
'戻り値:
'
'注意事項:
'   TestConditionワークシートを指定して下さい
'
    
    Dim pWorkSheetobj As Worksheet
    
    On Error GoTo ErrHandler
    Set pWorkSheetobj = Worksheets(pWorksheetName)
    'これ以降のエラーが全部"シートがない"になってしまうのでエラーハンドラをリセット
    On Error GoTo 0
    
    'シート状態確認処理
    Call CheckConditionSht(pWorkSheetobj)
    
    '状態確認がOKだったら条件表Objとして設定
    Set m_TestConditionSheet = pWorkSheetobj
    Set pWorkSheetobj = Nothing
            
    'データキャッシュの消去
    Call ClearConditions
            
    Exit Property
        
ErrHandler:
    '指定ワークシートの存在確認とシートObjectの設定
    Call RaiseError("TestCondition worksheet is nothing")
    MsgBox "TestCondition worksheet is nothing"
    Exit Property

End Property

Public Property Get TestConditionSheet() As String
'内容:
'   TestConditionワークシート
'
'パラメータ:
'
'戻り値:
'   TestConditionワークシート名
'
'注意事項:
'   TestConditionワークシートが未設定の場合はNothingとなります
'
    
    TestConditionSheet = m_TestConditionSheet.Name

End Property

Public Property Let CanHistoryRecord(ByVal pRecordFlag As Boolean)
'内容:
'   条件設定履歴記録機能の設定状態
'
'パラメータ:
'   [pRecordFlag] In Boolean: 有効(True)／無効(False)
'
'戻り値:
'
'注意事項:
'   Initialize後のDefault値はFalse、履歴記録機能無効状態です。
'
    '#### ProductionMode チェックして、設定するModeを決定する ####
    Dim YesOrNo As VbMsgBoxResult
    
    '現在の設定値を格納
    Dim LastFlagValue As Boolean
    LastFlagValue = m_HistoryRecordFlag
    
    If TheExec.RunMode = runModeProduction Then
        If pRecordFlag = True Then
            YesOrNo = MsgBox("Now ProductionMode! But SetLogModeTheCondition is True!" & vbCrLf & _
                            "EeeJOB will output DataLog." & vbCrLf & _
                            "Please make sure." & vbCrLf & _
                            "When You don't want to output This MsgBox, Please change runModeProduction => runModeDebug" & vbCrLf & _
                            "" & vbCrLf & _
                            "Output DataLog?" & vbCrLf _
                            , vbYesNo + vbQuestion, "Confirm Output DataLog(SetLogModeTheCondition)")
            If YesOrNo = vbYes Then
                m_HistoryRecordFlag = True
            Else
                m_HistoryRecordFlag = False
            End If
        Else
            m_HistoryRecordFlag = False
        End If
    Else
        m_HistoryRecordFlag = pRecordFlag
    End If
    
    If m_HistoryRecordFlag = True Then
        TheExec.Datalog.WriteComment "Eee JOB Output Log! :TheCondition"
    End If
    
    'フラグの設定値が前回設定値のまま-->切り替えは無し
    If LastFlagValue = m_HistoryRecordFlag Then
        '何もしない
    Else
        '履歴取得ON-->Off
        If LastFlagValue = True And m_HistoryRecordFlag = False Then
            Call ChangeInterceptorSettings(False)
        '履歴取得OFF->ON
        Else
            Call ChangeInterceptorSettings(True)
        End If
    End If

End Property

Public Property Get CanHistoryRecord() As Boolean
'内容:
'   条件設定履歴記録機能の設定状態
'
'パラメータ:
'
'戻り値:
'   条件設定履歴記録機能の設定状態
'
'注意事項:
'   設定状態の値は、True=有効／False=無効です
'
    CanHistoryRecord = m_HistoryRecordFlag

End Property

Public Property Get ExecHistory(ByVal pHistoryNumber As Long) As CSetConditionHistory
'内容:
'   条件設定履歴情報
'
'パラメータ:
'   [pHistoryNumber]　In Long型：  取得する履歴番号の指定
'
'戻り値:
'   条件設定履歴情報Object(CSetConditionHistory型)
'
'注意事項:
'   [pHistoryNumber]で指定された番号に対応する情報が
'    存在しない場合の取得値はNothingとなります
'
    Set ExecHistory = m_SetConditionLogger.Record(pHistoryNumber)

End Property

Public Property Get HistoryCount() As Long
'内容:
'   記録されている条件設定履歴数
'
'パラメータ:
'
'戻り値:
'   記録されている条件設定履歴数
'
'注意事項:
'

    HistoryCount = m_SetConditionLogger.RecordCount

End Property
Public Property Let SelectExecSetting(ByVal pSelectExecSetting As String)
'内容:
'   選択されたマクログループ定義名
'
'パラメータ:
'   [pSelectExecSetting]　In String型：  選択されたマクログループ定義名
'
'戻り値:
'
'注意事項:
'
    m_SelecteExecSetting = pSelectExecSetting

End Property

Public Property Get SelectExecSetting() As String
'内容:
'   選択されたマクログループ定義名
'
'パラメータ:
'
'戻り値:
'   選択されたマクログループ定義名
'
'注意事項:
'
    SelectExecSetting = m_SelecteExecSetting

End Property



Public Sub ClearExecHistory()
'内容:
'   条件設定履歴情報をすべて削除
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Call m_SetConditionLogger.ClearRecord

End Sub

Public Sub SaveHistoryLog(ByVal pFileName As String)
'内容:
'   条件設定履歴情報ログをCSV形式で指定ファイルに出力
'
'パラメータ:
'   [pFileName]  In  String型: 条件設定履歴を出力するFullpathファイル名
'
'戻り値:
'
'注意事項:
'   ファイルの指定はFullpath指定して下さい（例:"d:\DataLog\SettingLog.csv"）
'   指定ファイルがすでに存在する場合は追記記録(append)となります。
'   情報は、CSV形式で出力されます。
'   履歴情報が記録されていない状態で実行すると何もしません。
'

    Call m_SetConditionLogger.SaveRecord(pFileName)
    
End Sub

Public Sub RunFunction(ByVal pFunctionName As String, _
Optional pArg0 As Variant = Empty, _
Optional pArg1 As Variant = Empty, _
Optional pArg2 As Variant = Empty, _
Optional pArg3 As Variant = Empty, _
Optional pArg4 As Variant = Empty, _
Optional pArg5 As Variant = Empty, _
Optional pArg6 As Variant = Empty, _
Optional pArg7 As Variant = Empty, _
Optional pArg8 As Variant = Empty, _
Optional pArg9 As Variant = Empty)
'内容:
'   セット関数の直接実行
'
'パラメータ:
'   [pFunctionName]  In  String型:   実行するセット関数名称
'   [pArg0]  In  Valiant型:  セット関数の第1パラメータ値(オプション)
'   [pArg1]  In  Valiant型:  セット関数の第2パラメータ値(オプション)
'   [pArg2]  In  Valiant型:  セット関数の第3パラメータ値(オプション)
'   [pArg3]  In  Valiant型:  セット関数の第4パラメータ値(オプション)
'   [pArg4]  In  Valiant型:  セット関数の第5パラメータ値(オプション)
'   [pArg5]  In  Valiant型:  セット関数の第6パラメータ値(オプション)
'   [pArg6]  In  Valiant型:  セット関数の第7パラメータ値(オプション)
'   [pArg7]  In  Valiant型:  セット関数の第8パラメータ値(オプション)
'   [pArg8]  In  Valiant型:  セット関数の第9パラメータ値(オプション)
'   [pArg9]  In  Valiant型:  セット関数の第10パラメータ値(オプション)
'
'戻り値:
'
'注意事項:
'   [pFunctionName]で指定されたセット関数が存在しない場合はエラーとなります
'   セット関数のパラメータpArg*の値にObjectは設定できません
'   最終パラメータには終端文字を必ず指定してください。
'

    'パラメータ情報を一旦テンポラリの情報に格納
    Dim tmpParameterInfo As CSetFunctionInfo
    Set tmpParameterInfo = New CSetFunctionInfo
    With tmpParameterInfo
        .ConditionName = "DirectRun"
        .FunctionName = pFunctionName
        .AddParameter pArg0
        .AddParameter pArg1
        .AddParameter pArg2
        .AddParameter pArg3
        .AddParameter pArg4
        .AddParameter pArg5
        .AddParameter pArg6
        .AddParameter pArg7
        .AddParameter pArg8
        .AddParameter pArg9
    End With
    
    'セット関数のパラメータ状態チェック処理
    Call ChkArgValue(tmpParameterInfo)
    
    'パラメータ情報Objectの生成
    Dim ParameterInfo As CSetFunctionInfo
    Set ParameterInfo = New CSetFunctionInfo
    
    'パラメータ情報Objectに情報を登録する
    Dim Argnum As Long
    With ParameterInfo
        .ConditionName = tmpParameterInfo.ConditionName '条件名称登録
        .FunctionName = tmpParameterInfo.FunctionName   'セット関数名称登録
        .IsEnableMacro = True                           'マクロ有効設定値
        'セット関数のパラメータ登録処理
        For Argnum = 0 To tmpParameterInfo.ArgParameterCount - 1 Step 1
            If ChkRunParameter(tmpParameterInfo.Arg(Argnum)) Then
                Call .AddParameter(tmpParameterInfo.Arg(Argnum))
            End If
        Next Argnum
    End With
        
    'マクロ実行Objectを設定
    Set ParameterInfo.MacroController = CreateMacroController(ParameterInfo)
        
    'セット関数の実行
    Call ParameterInfo.RunMacro

End Sub

Public Sub ExecuteMacro(ByVal pSetFunctionInfo As CSetFunctionInfo)
'内容:
'   TestParameterEditorフォームで編集したConditionマクロを実行する
'
'[pSetFunctionInfo] IN CSetFunctionInfo型:  フォームで編集されたSetFunctionInfo
'
'備考:
'
'

    'マクロ実行Objectを設定
    Set pSetFunctionInfo.MacroController = CreateMacroController(pSetFunctionInfo)
    
    'セット関数の実行
    Call pSetFunctionInfo.RunMacro
        
End Sub

Public Sub SetCondition(ByVal pConditionName As String)
'内容:
'   TestConditionシートの情報を使用したセット関数の実行
'
'パラメータ:
'   [pConditionName]  In  String型:  設定する条件名称
'
'戻り値:
'
'注意事項:
'   [pConditionName]で指定された条件が定義されていない場合はエラーとなります
'   TestConditionシートに記載されているセット関数が存在しない場合はエラーとなります
'
    Dim SetConditions As Collection
    Dim SetInfomation As CSetFunctionInfo
    Dim SetCnt As Long
        
    'データキャッシュの状態確認(存在しない場合は条件表のLoadを行い作成)
    If m_Conditions Is Nothing Then
        Me.LoadCondition
    End If
    
    '指定された条件名の、条件データセット取り出し
    
    Dim ConditionCount As Long
    On Error GoTo CONDITION_NOTHING_ERR
        ConditionCount = m_Conditions.Item(pConditionName).Count
    On Error GoTo 0
            
    '設定関数LOOP
    Dim ExecuteCount As Long
    ExecuteCount = 1
    For SetCnt = 1 To ConditionCount Step 1
       
        '条件名称に対応する、設定条件情報Objectの取得
        Call m_Conditions.Item(pConditionName).Item(SetCnt).RunMacro
    Next SetCnt

    Exit Sub

CONDITION_NOTHING_ERR:
'指定された条件名称に対応するデータが存在しない時のエラー処理
    Call RaiseError("Condition name = " & pConditionName & " is nothing")

End Sub

Public Function GetConditionInfo(ByVal pConditionName As String) As Collection
'内容:
'   指定された条件名称で定義されている測定条件Collectionを取得
'
'パラメータ:
'   [pConditionName]  In  String型:  取得する条件名称
'
'戻り値:
'   測定条件データ一式 (Collection)
'
'注意事項:
'   [pConditionName]で指定された条件が定義されていない場合はエラーとなります
'

    '指定された条件名の、条件データセット取り出し
    On Error GoTo CONDITION_NOTHING_ERR
    Set GetConditionInfo = m_Conditions.Item(pConditionName)
    On Error GoTo 0

    Exit Function
    
CONDITION_NOTHING_ERR:
'指定された条件名称に対応するデータが存在しない時のエラー処理
    Call RaiseError("Condition name = " & pConditionName & " is nothing")

End Function

Public Function GetCloneConditionInfo(ByVal pConditionName As String) As Collection
'内容:
'   指定された条件名称で定義されている測定条件Collectionのコピーを作成する。
'
'パラメータ:
'   [pConditionName]  In  String型:  取得する条件名称
'
'戻り値:
'   測定条件データ一式 (Collection)
'
'注意事項:
'   [pConditionName]で指定された条件が定義されていない場合はエラーとなります
'
    '指定された条件名の、条件データセット取り出し
    On Error GoTo CONDITION_NOTHING_ERR
    Dim ConditionInfo As Collection
    Set ConditionInfo = m_Conditions.Item(pConditionName)
    On Error GoTo 0
    
    'Condition情報のコピーを作成
    Dim CloneConditionInfo As Collection    'コピーConditionInfo
    Set CloneConditionInfo = New Collection
    Dim CloneFunctionInfo As CSetFunctionInfo    'コピーFunctionInfo
    Dim InfoCount As Long
    
    For InfoCount = 1 To ConditionInfo.Count Step 1
    
        Set CloneFunctionInfo = ConditionInfo.Item(InfoCount)
        Call CloneConditionInfo.Add(CloneFunctionInfo)
        Set CloneFunctionInfo = Nothing
    
    Next InfoCount
    
    Set GetCloneConditionInfo = CloneConditionInfo

    Exit Function

CONDITION_NOTHING_ERR:
'指定された条件名称に対応するデータが存在しない時のエラー処理
    Call RaiseError("Condition name = " & pConditionName & " is nothing")

End Function

Public Sub LoadCondition()
'内容:
'   TestConditionシートデータの読み込み
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   TestConditionシートが未設定の場合はエラーとなります
'   ConditionNameの値が重複定義されているとエラーとなります
'   Defaultラベルの位置が違う場合はエラーとなります(位置はMacroExecuteSettingの一番左端)
'

'===SetFunctionInfoに情報を読み込む前のエラーチェック処理 ===============================
    'TestConditionシートの確認処理
    If m_TestConditionSheet Is Nothing Then
        Call RaiseError("TestCondition worksheet is nothing!" & vbCrLf _
          & "Plese set TestConditionSheet property")
        Exit Sub
    End If
        
    'TestConditionシートのマクロ実行設定値の確認処理
    Dim ErrorAddress As String
    If ChkMacroExecuteSetting(ErrorAddress) = False Then
        Call RaiseError("Sheet = " & m_TestConditionSheet.Name & vbCrLf _
            & "Cell = " & ErrorAddress & " is unknown value" & vbCrLf _
            & "Please set " & MACRO_ENABLE_WORD & " or " & MACRO_DISABLE_WORD & " value")
            Exit Sub
    End If
    
    Call UpdateExecSettingList  'マクログループ選択セルのリスト更新
    Call CheckDefaultLabel  'シートのDefaultラベルのエラーチェック
    Call CheckExecuteSetting   'マクログループ定義のエラーチェック

'=========================================================================================
    
    Call CreateExecSettingDataList 'マクログループ定義の有効無効リストを作成
    
    Dim NowCondName As String       '現在のIndexが指している条件名称
    Dim ConditionIdx As Range       '条件名称のIndex
    Dim SetFunctionIdx As Range     '設定関数のIndex
    Dim ParameterIdx As Range       '設定関数のパラメータIndex
    Dim tmpCondInfo As Collection   '1条件分の条件情報格納用コレクション
    Dim tmpInfo As CSetFunctionInfo '条件情報型のクラス
    Dim ArgNumber As Long           'Arg*パラメータカウンタ

    'セルのIndexスタート位置設定
    Set ConditionIdx = m_TestConditionSheet.Range(CONDITION_SHEET_START_CELL) '条件設定表の開始Index指定
    Set SetFunctionIdx = ConditionIdx.offset(0, 1)   '設定関数の開始Index指定
        
    'ALL条件一覧データ用の箱の準備
    Set m_Conditions = Nothing
    Set m_ConditionNameList = Nothing
    Set m_Conditions = New Collection
    Set m_ConditionNameList = New Collection

    '条件表キャッシュ作成LOOP
    Dim ConditionRowCount As Long
    ConditionRowCount = 1
    While Not IsEmpty(ConditionIdx.Cells)
        NowCondName = ConditionIdx.Value '現在指し示している場所の条件名称を取得
        Set tmpCondInfo = New Collection '設定条件Object格納用コレクション
        
        While (ConditionIdx.Value = NowCondName)
            'パラメータCellのIndex設定（SetFunction Cellの2つ右隣に設定）
            Set ParameterIdx = SetFunctionIdx.offset(0, PROCEDURE_TO_ARG_OFFSET)
            '測定条件情報用Objectを用意
            Set tmpInfo = New CSetFunctionInfo
            '測定条件情報用Objectにデータを設定
            With tmpInfo
                .ConditionName = CStr(ConditionIdx.Value)
                .FunctionName = CStr(SetFunctionIdx.Value)
                .ConditionSheetName = m_TestConditionSheet.Name
                .ExecuteSettingName = m_SelecteExecSetting
                ArgNumber = 1
                While ((IsTermCharacter(ParameterIdx.Cells.Value) = False) And (ArgNumber <= MAX_ARG_NUMBER))
                    'パラメータデータの途中空欄チェック処理
                    If IsEmpty(ParameterIdx.Cells) = True Then
                        Call ClearConditions
                        With ParameterIdx
                            Call RaiseError(tmpInfo.FunctionName & " parameter is Empty" & vbCrLf & _
                            "Worksheet = " & m_TestConditionSheet.Name & "  Cell = " & .Address)
                            Exit Sub
                        End With
                    End If
                    'パラメータデータの登録処理
                    .AddParameter (ParameterIdx.Cells.Value)
                    ArgNumber = ArgNumber + 1
                    Set ParameterIdx = ParameterIdx.offset(0, 1)
                Wend
            End With
            
            '=== ユーザマクロの有効無効判定処理追加 ================================================
            'マクログループ選択セルで選択したグループ列に設定
            Dim ExecSettingData As Collection
            Set ExecSettingData = m_ExecSettingDataList.Item(m_SelecteExecSetting)
            
            Dim ExecSettingValue As String
            ExecSettingValue = ExecSettingData.Item(ConditionRowCount)
            
            If StrComp(ExecSettingValue, MACRO_ENABLE_WORD, vbTextCompare) = 0 Then
                'マクロ有効時(Excelマクロ実行機能を持たせる)
                With tmpInfo
                    .IsEnableMacro = True
                    Set .MacroController = CreateMacroController(tmpInfo)
                End With
            Else
                'マクロ無効時(偽物マクロ実行機能を持たせる)
                With tmpInfo
                    .IsEnableMacro = False
                     Set .MacroController = CreateMacroController(tmpInfo)
                End With
            End If
            
            '========================================================================================
                                    
            '○○条件時の測定条件情報を格納
            Call tmpCondInfo.Add(tmpInfo)
                        
            '現在の条件名称を保存し、Indexを進める(条件名称のセルを基準とする)
            With ConditionIdx
                NowCondName = .Value
                Set ConditionIdx = .offset(1, 0)             '条件名称Cellを現在の位置から1つ下に進める
            End With
            Set SetFunctionIdx = ConditionIdx.offset(0, 1)   'SetFunction Cellを現在の位置から条件名称Cellの1つ右に設定
            ConditionRowCount = ConditionRowCount + 1
        Wend
        
        '条件名称重複エラーハンドル
        On Error GoTo CONDITION_NAME_ERROR
        Call m_Conditions.Add(tmpCondInfo, NowCondName)   '条件名称ラベルをキーに対応する条件データ群を登録
        Call m_ConditionNameList.Add(NowCondName)   '条件名称のリストに追加。重複はしていないはず。
        On Error GoTo 0
    Wend
    Exit Sub

CONDITION_NAME_ERROR:
'Condition Name重複時のエラー処理
    Call ClearConditions
    Call TheError.Raise(ERR_NUMBER, TypeName(Me), "Condition Name = " & NowCondName & " is overlapping ?")
    
End Sub

Public Sub AddCondShtValidation(ByRef pTargetRange As Range)
'内容:
'   測定条件ワークシートに入力規則を追加する
'
'パラメータ:
'   [pTargetRange]  In  Range型:  変更対象
'
'戻り値:
'
'注意事項:
'
    
    With pTargetRange
    
        'マクログループ選択セルにフォーカスした場合
        If .Column = EXECUTE_SETTINGLIST_COL And .Row = EXECUTE_SETTINGLIST_ROW Then
            'マクログループ選択セルのリスト更新
            Call UpdateExecSettingList
        End If
        
        'マクログループ定義ラベルにフォーカスした場合
        If .Column >= MACRO_SETTING_START_COLUMN And _
        .Column <= MACRO_SETTING_END_COLUMN And _
        .Row = MACRO_SETTING_START_ROW And _
        .Count = 1 Then
        
            'マクログループ選択セルのリスト更新
            Call UpdateExecSettingList
        
            '変更前のGroup名を保持する
            m_BackUpGroupName = pTargetRange.Value
        End If
    
        'enble/disable入力セルが選択された場合
        If .Column >= MACRO_SETTING_START_COLUMN And _
        .Column <= MACRO_SETTING_END_COLUMN And _
        .Row >= SHEET_DATA_START_ROW And _
        .Count = 1 Then
            Call AddExecuteColumnValidation(pTargetRange) 'セルの入力規則による入力支援
'           Call AddExecuteColumnListBox(pTargetRange)   'ListBoxによる入力支援
        End If
    End With

End Sub

Public Sub AddCondShtFormat(ByRef pTargetRange As Range)
'内容:
'   測定条件ワークシートのセル書式を追加する
'
'パラメータ:
'   [pTargetRange]  In  Range型:  変更対象
'
'戻り値:
'
'注意事項:
'

    If m_SettingNameColList Is Nothing Then
        Call UpdateExecSettingList  'マクログループ選択セルのリスト更新
    End If
    
    Dim ExecuteSettingList As Range
    Set ExecuteSettingList = m_TestConditionSheet.Range(EXECUTE_SETTINGLIST)
    
    'ここでは、マクログループ選択リストを参照してチェックを行う
    On Error GoTo EXECUTE_SETTING_NAME_ERROR
    Call m_SettingNameColList.Item(CStr(ExecuteSettingList.Value))
    On Error GoTo 0
    
    With pTargetRange
        
        'マクログループ定義ラベル名の変更があった場合
        If .Column >= MACRO_SETTING_START_COLUMN And _
        .Column <= MACRO_SETTING_END_COLUMN And _
        .Row = MACRO_SETTING_START_ROW And _
        .Count = 1 Then
        
            '選択中のマクログループと同じマクログループ定義欄を変更した場合
            Dim ExecSettingCol As Long
            ExecSettingCol = m_SettingNameColList.Item(m_SelecteExecSetting)
            If .Column = ExecSettingCol And pTargetRange.Value <> m_SelecteExecSetting Then
            
                Call MsgBox("'" & m_SelecteExecSetting & "'" & " is active.", vbCritical, "TestConditionManager")
                pTargetRange.Value = m_SelecteExecSetting '入力をキャンセルする
                Exit Sub
            End If
            
            '変更したマクログループ名の前後にスペース(半角/全角)、又はカンマを含む入力をした場合
            If (StrComp(Left(pTargetRange.Value, 1), " ") = 0) Or _
                (StrComp(Right(pTargetRange.Value, 1), " ") = 0) Or _
                (StrComp(Left(pTargetRange.Value, 1), "　") = 0) Or _
                (StrComp(Right(pTargetRange.Value, 1), "　") = 0) Or _
                (InStr(pTargetRange.Value, ",")) Then
                
                Call MsgBox("Comma or blank can no be used.", , "TestConditionManager")
                pTargetRange.Value = m_BackUpGroupName '入力をキャンセルする
                Exit Sub
            End If
        End If
    
        'マクログループ選択セルの値変更があった場合
        If .Column = EXECUTE_SETTINGLIST_COL And .Row = EXECUTE_SETTINGLIST_ROW Then
        
            If TheExec.Flow.IsRunning = False Then

                If ChangeExecuteSetting(.Value) = False Then
                    .Value = m_SelecteExecSetting
                    Exit Sub
                End If
            End If
            
            'マクログループ選択セルのリスト更新
            Call UpdateExecSettingList
            Call Me.ModifyConditionSheet
        End If
        
        '測定条件表が使用しているセル範囲の値変更が合った場合セルの書式を設定する
        If .Column = m_SettingNameColList(m_SelecteExecSetting) And _
        .Row >= SHEET_DATA_START_ROW Then
            With Excel.Application
                .ScreenUpdating = False
                Call Me.ModifyConditionSheet
                .ScreenUpdating = True
            End With
        End If
    End With
    Exit Sub
    
EXECUTE_SETTING_NAME_ERROR:
    '切り替えマクログループ定義名は定義されていない
    Call MsgBox("""" & ExecuteSettingList.Value & """" & " is not defined.", , "TestConditionManager")
    ExecuteSettingList.Value = m_SelecteExecSetting
End Sub

Private Sub UpdateExecSettingList()
'内容:
'   ExecuteSetting選択セルの更新を行う
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'

    'マクログループ名の行番号リストを作成する
    Set m_SettingNameColList = Nothing
    Set m_SettingNameColList = New Collection
    
    'マクログループ名リスト
    Dim SettingNameList As Collection
    Set SettingNameList = New Collection
    
    'マクログループ選択セル入力用グループ名リスト
    Dim NameList As String
    NameList = ""
    
    'データ読み込み開始位置
    Dim CurrentRange As Range
    Set CurrentRange = m_TestConditionSheet.Range(MACRO_SETTING_DEFAULT)
    
    While Not CurrentRange.Column > MACRO_SETTING_END_COLUMN
    
        'グループ名が未入力の場合は無視する
        If CurrentRange.Value <> "" Then
        
            Dim Registration As Boolean
            Registration = True

            'マクログループ名の重複チェック
            Dim SettingNameCount As Long
            For SettingNameCount = 1 To SettingNameList.Count Step 1
                
                If StrComp(CurrentRange.Value, SettingNameList.Item(SettingNameCount)) = 0 Then
                    Registration = False
                End If
            Next SettingNameCount
            
            If Registration = True Then
            
                'マクログループ名リストに追加
                SettingNameList.Add (CurrentRange.Value)
                
                'マクログループ選択セル入力用グループ名リストに追加
                If NameList = "" Then
                    NameList = CurrentRange.Value
                Else
                    NameList = NameList & "," & CurrentRange.Value
                End If
                
                'グループのColumn位置リストに追加。
                Call m_SettingNameColList.Add(CurrentRange.Column, CStr(CurrentRange.Value))
            End If
        
        End If
        
        Set CurrentRange = CurrentRange.offset(0, 1) '横に移動
    Wend
    
    Dim ExecuteSettingList As Range
    Set ExecuteSettingList = m_TestConditionSheet.Range(EXECUTE_SETTINGLIST)
    Call addValidation(ExecuteSettingList, "", NameList)
    m_SelecteExecSetting = ExecuteSettingList.Value
    
End Sub

'コンストラクタ
Private Sub Class_Initialize()
    Me.Initialize
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Set m_TestConditionSheet = Nothing
    
    Set m_ConditionNameList = Nothing
    Set m_SettingNameColList = Nothing
    Set m_ExecSettingDataList = Nothing
    
End Sub

'インターセプターの有無を切り替える関数
Private Sub ChangeInterceptorSettings(ByVal pIsInterceptorDisabeToEnable As Boolean)
    
   'データキャッシュの状態確認
   '(存在しない場合はエラー 初期化未実行か、シートを変更して未LOAD状態)
    If m_Conditions Is Nothing Then
        Call RaiseError("TestCondition data sheet is not loaded!")
        Exit Sub
    End If

    'Instanceの参照切り替え処理
    Dim RetConditionCollection As Collection    '条件毎のマクロインスタンスセット
    Dim RetSetfunction As CSetFunctionInfo      '単独マクロインスタンス
    Dim InterceptorObj As CMacroExecInterceptor 'Interceptor型変換用
        
    For Each RetConditionCollection In m_Conditions
        For Each RetSetfunction In RetConditionCollection
            '--- Interceptor無効-->有効 ---
            If pIsInterceptorDisabeToEnable = True Then
                Set InterceptorObj = New CMacroExecInterceptor
                With InterceptorObj
                    Call .Initialize
                    Set .SetConditionLogger = m_SetConditionLogger
                    Set .MacroController = RetSetfunction.MacroController
                End With
                'Interceptorの参照を設定
                Set RetSetfunction.MacroController = InterceptorObj
            '------------------------------
            
            '--- Interceptor有効-->無効 ---
            Else
                'InterceptorからMacroControllerの参照を取得(I/F切り替え処理)
                Set InterceptorObj = RetSetfunction.MacroController
                'Interceptorの持っていたMacroControllerの参照を設定
                Set RetSetfunction.MacroController = InterceptorObj.MacroController
            End If
            '------------------------------
        Next
    Next

End Sub

'マクロ実行機能Objectの生成
Private Function CreateMacroController(ByRef pMacroInfo As CSetFunctionInfo) As IMacroController
    If m_HistoryRecordFlag = True Then
        If pMacroInfo.IsEnableMacro = True Then
            'LOG=ON,MACRO=ENABLE
            Set CreateMacroController = CreateMacroExecInterceptor(CreateExcelMacroController)
        Else
            'LOG=ON,MACRO=DISABLE
            Set CreateMacroController = CreateMacroExecInterceptor(Nothing)
        End If
    Else
        If pMacroInfo.IsEnableMacro = True Then
            'LOG=OFF,MACRO=ENABLE
            Set CreateMacroController = CreateExcelMacroController
        Else
            'LOG=OFF,MACRO=DISABLE
            Set CreateMacroController = CreateDmyMacroController
        End If
    End If
End Function

'マクロ実行Interceptor Objectの生成
Private Function CreateMacroExecInterceptor(ByRef pMacroController As IMacroController) As CMacroExecInterceptor
    Set CreateMacroExecInterceptor = New CMacroExecInterceptor
    With CreateMacroExecInterceptor
        Call .Initialize
        Set .SetConditionLogger = m_SetConditionLogger
        Set .MacroController = pMacroController
    End With
End Function

'マクロ実行機能 Objectの生成
Private Function CreateExcelMacroController() As CExcelMacroController
    Set CreateExcelMacroController = New CExcelMacroController
    With CreateExcelMacroController
        Call .Initialize
        .TargetBookName = m_JobWorkBookName
    End With
End Function

'ダミーマクロ実行機能 Objectの生成
Private Function CreateDmyMacroController() As CDmyMacroController
    Set CreateDmyMacroController = New CDmyMacroController
End Function

'Argパラメータの状態がルールに従っているかを確認する関数
Private Sub ChkArgValue(ByRef pChkParameter As CSetFunctionInfo)
    
    Dim ParamCnt As Long
    Dim EmptyDataNo As Long
    Dim TermDataNo As Long
    
    'フラグのInitial値
    Const INIT_DAT_NO As Long = -1
    
    'フラグの初期化
    EmptyDataNo = INIT_DAT_NO
    TermDataNo = INIT_DAT_NO
    
    'データチェックLOOP
    With pChkParameter
        For ParamCnt = 0 To .ArgParameterCount - 1 Step 1
            'Empty状態の確認
            If IsEmpty(.Arg(ParamCnt)) = True Then
                EmptyDataNo = ParamCnt
            End If
            '終端文字の確認
            If IsTermCharacter(.Arg(ParamCnt)) = True Then
                TermDataNo = ParamCnt
            End If
            '判定処理
            If (EmptyDataNo <> INIT_DAT_NO) And (TermDataNo <> INIT_DAT_NO) Then '初期値状態でない事を確認
                'エラーモード判定処理
                If EmptyDataNo < TermDataNo Then
                    Call RaiseError(.FunctionName & " Arg = " & EmptyDataNo & " data is Empty!")
                End If
            End If
        Next ParamCnt
        '終端文字の存在判定処理(終端文字がない、かつ、Emptyのデータが存在する)
        If (TermDataNo = INIT_DAT_NO) And (EmptyDataNo <> INIT_DAT_NO) Then
            Call RaiseError("Data termination character is nothing")
        End If
    End With
    
End Sub

'セット関数のパラメータの値確認用関数(パラメータObject作成用)
Private Function ChkRunParameter(ByVal pArgParameter As Variant) As Boolean
    If (pArgParameter <> Empty) And (IsTermCharacter(pArgParameter) = False) Then
        ChkRunParameter = True
    Else
        ChkRunParameter = False
    End If
End Function

'任意の文字列が、終端文字かどうかを判定する関数（戻り値 終端文字=True）
Private Function IsTermCharacter(ByVal pChkCharacter As String) As Boolean
    Dim RetTermChar As Variant
    For Each RetTermChar In m_TermCharacters
        If UCase(pChkCharacter) = RetTermChar Then
            IsTermCharacter = True
            Exit Function
        Else
            IsTermCharacter = False
        End If
    Next RetTermChar
End Function

'表のデータキャッシュの消去
Private Sub ClearConditions()
    Set m_Conditions = Nothing
End Sub

'条件表ワークシートの状態確認処理
Private Sub CheckConditionSht(ByVal pTargetWkSht As Worksheet)
    With pTargetWkSht
        'チェック1(シートタイトルを確認)
        If .Range(SHEET_TITLE_CELL).Value <> SHEET_TITLE Then
            Call RaiseError(.Name & " worksheet is not a test condition worksheet")
            Exit Sub
        End If
        'チェック2(Condition Nameデータの位置確認)
        If .Range(CONDITION_TITLE_CELL).Value <> CONDITION_TITLE Then
            Call RaiseError("The " & CONDITION_TITLE & " position of the " & .Name & " is abnormal")
            Exit Sub
        End If
        'チェック3(Set Function Nameデータの位置確認)
        If .Range(SET_FUNCTION_TITLE_CELL).Value <> SET_FUNCTION_TITLE Then
            Call RaiseError("The " & SET_FUNCTION_TITLE & " position of the " & .Name & " is abnormal")
            Exit Sub
        End If
    End With
End Sub

'ワークシートの値変更時のイベント取得(シートが変更されたらデータキャッシュを消去する)
Private Sub m_TestConditionSheet_Change(ByVal pTarget As Range)

    If TheExec.Flow.IsRunning = False Then
    
        Call ClearConditions
    End If
End Sub

'エラーレイズ用関数
Private Sub RaiseError(ByVal pErrMsg As String)
    If Not (TheError Is Nothing) Then
        'エラーマネージャへ
        Call TheError.Raise(ERR_NUMBER, TypeName(Me), pErrMsg)
    Else
        'VBのErrへ
        Call Err.Raise(ERR_NUMBER, TypeName(Me), pErrMsg)
    End If
End Sub

'マクロの有効無効設定値に合わせてセル書式を設定する関数
Private Sub MakeDatasheetFormat()
    Const COLOR_INDEX_GRAY As Long = 15 '灰色
    Dim ConditionRng As Range
    Set ConditionRng = m_TestConditionSheet.Range(CONDITION_SHEET_START_CELL)
    Dim DataRng As Range
    Set DataRng = m_TestConditionSheet.Range(DATA_START_RANGE)
    Dim MacroSettingRng As Range
    Dim OffsetColumn As Long '横移動の変数
    OffsetColumn = m_SettingNameColList(m_SelecteExecSetting) - ConditionRng.Column
    Set MacroSettingRng = ConditionRng.offset(0, OffsetColumn)
    While Not IsEmpty(ConditionRng)
        If StrComp(MacroSettingRng.Value, MACRO_ENABLE_WORD, vbTextCompare) = 0 Or IsEmpty(MacroSettingRng) Then
            'マクロ有効(何もしない)
        Else
            'マクロ無効(グレー網掛け)
            With DataRng.Interior
                .Pattern = xlGray8
                .ColorIndex = COLOR_INDEX_GRAY
            End With
        End If
        Set ConditionRng = ConditionRng.offset(1, 0)
        Set MacroSettingRng = MacroSettingRng.offset(1, 0)
        Set DataRng = DataRng.offset(1, 0)
    Wend
    Set DataRng = Nothing
    Set MacroSettingRng = Nothing
    Set ConditionRng = Nothing
End Sub

'測定条件表のセル書式設定を規定値にする関数
Private Sub ClearDataSheetFormat()
    Dim ConditionRange As Range
    Dim DataRange As Range
    Set ConditionRange = m_TestConditionSheet.Range(CONDITION_SHEET_START_CELL)
    Set DataRange = m_TestConditionSheet.Range(DATA_START_RANGE)
    While Not IsEmpty(ConditionRange)
        With DataRange.Interior
            .Pattern = xlSolid
            .ColorIndex = xlNone
        End With
        Set DataRange = DataRange.offset(1, 0)
        Set ConditionRange = ConditionRange.offset(1, 0)
    Wend
End Sub

Private Function ChkMacroExecuteSetting(ByRef pErrCellAddress As String) As Boolean
'内容:
'   TestConditionシート全体のマクロ有効無効指定値のチェック関数
'
'パラメータ:
'   [pErrCellAddress]  Out  String型:     enable/disable以外の値が入力されたセルのアドレス
'
'戻り値:
'   マクロの有効無効指定値の異常の有無  Boolean型
'
'注意事項:
'

    'マクログループ定義の左端(Default)からスタート
    Dim ExecSettingRange As Range
    Set ExecSettingRange = m_TestConditionSheet.Range(MACRO_SETTING_DEFAULT)
    
    'TestConditionシートに記述されているConditionと有効無効値は対になっている筈、
    'Conditionを有効無効値の全体数と考える
    Dim ConditionRange As Range
    Set ConditionRange = m_TestConditionSheet.Range(CONDITION_SHEET_START_CELL)
    
    While Not ExecSettingRange.Column > MACRO_SETTING_END_COLUMN
    
        Dim ConditionDataRange As Range
        Set ConditionDataRange = ConditionRange
    
        'マクログループ名が未入力の場合は無視
        If ExecSettingRange.Value <> "" Then
            
            Dim ExecSettingDataRange As Range
            Set ExecSettingDataRange = ExecSettingRange.offset(1, 0)
            
            While Not IsEmpty(ConditionDataRange)
            
                With ExecSettingDataRange
                    If StrComp(.Value, MACRO_ENABLE_WORD, vbTextCompare) = 0 Or _
                    StrComp(.Value, MACRO_DISABLE_WORD, vbTextCompare) = 0 Then
                        'PASS(何もしない)
                    Else
                        'FAIL(エラーセルのアドレスを返す)
                        pErrCellAddress = .Address
                        ChkMacroExecuteSetting = False
                        Exit Function
                    End If
                End With
                    
                Set ExecSettingDataRange = ExecSettingDataRange.offset(1, 0)
                Set ConditionDataRange = ConditionDataRange.offset(1, 0)
                
            Wend
        End If
        
        Set ExecSettingRange = ExecSettingRange.offset(0, 1) '次のグループ列へ移動
    Wend
    pErrCellAddress = ""
    ChkMacroExecuteSetting = True
End Function

'Execute Columにリストボックスを表示
Private Sub AddExecuteColumnListBox(ByRef pTargetRange As Range)
    Dim SetValueList As Collection
    Set SetValueList = New Collection
    With SetValueList
        Call .Add(MACRO_ENABLE_WORD)
        Call .Add(MACRO_DISABLE_WORD)
    End With
    Call CreateListBox(pTargetRange, SetValueList)
    Set SetValueList = Nothing
End Sub

'Execute Columnに入力規則追加
Private Sub AddExecuteColumnValidation(ByRef pTargetRange As Range)
    Const SET_VALUE_LIST As String = MACRO_ENABLE_WORD & "," & MACRO_DISABLE_WORD
    Dim column_name As String
    column_name = m_TestConditionSheet.Cells(MACRO_SETTING_START_ROW, pTargetRange.Column).Value
    Call addValidation(pTargetRange, column_name, SET_VALUE_LIST)
End Sub

'セルに入力規則を付与する為の関数
Private Sub addValidation(ByRef pTargetRange As Range, ByVal pDataName As String, _
ByVal pDataItemList As String, Optional ByVal pShowInputMessage As Boolean = False)

    'TestConditionシートがアクティブ時のみ入力規則を設定する。
    If IsActive = False Then
        Exit Sub
    End If

    '一度現在の入力規則を削除
    Call pTargetRange.Validation.Delete
    '入力規則を新規追加
    Dim ActionItemList As String
    ActionItemList = pDataItemList
    Call pTargetRange.Validation.Add( _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        operator:=xlEqual, _
        Formula1:=ActionItemList)
    '表示メッセージ等の設定
    With pTargetRange.Validation
        .InputTitle = pDataName                      'Balloonのタイトル値
        .InputMessage = "Please select"              'Balloonのメッセージ値
        .ShowInput = pShowInputMessage               'Balloonの表示設定
        .InCellDropdown = True                       'Dropdownの表示設定
        .ErrorTitle = "Unknown " & pDataName & " !"  'Error Msgbox タイトル
        .ErrorMessage = "Input Value is unknown !"   'Error Msgbox メッセージ
        .IgnoreBlank = False                         'ブランク入力許容設定
        .IMEMode = xlIMEModeNoControl                'IMEモード設定
        .ShowError = False                           'エラーダイアログを非表示にする
    End With
End Sub

'TestConditionシートのCondition名リストを取得する。
Public Property Get GetConditionNameList() As Collection

    Set GetConditionNameList = m_ConditionNameList
    
End Property

Public Function ChangeExecuteSetting(ByVal pExecuteSettingName As String) As Boolean
'内容:
'   マクログループ定義を切り替える
'
'パラメータ:
'   [pExecuteSettingName]  In  String型:  切り替えマクログループ定義名
'
'戻り値:
'   Boolean
'
'注意事項:
'   Job実行中はユーザーがCallしない限り、ここには来ない。
'
    On Error GoTo EXECUTE_SETTING_NAME_ERROR
        Call m_SettingNameColList.Item(pExecuteSettingName)
    On Error GoTo 0
    
    Dim ExecSettingList As Range
    Set ExecSettingList = m_TestConditionSheet.Range(EXECUTE_SETTINGLIST)
    
    
    If ExecSettingList.Value <> pExecuteSettingName Then
    
        'TheConditionのメンバにマクログループを設定する
        If Not (TheCondition Is Nothing) Then
        
            TheCondition.SelectExecSetting = pExecuteSettingName
        
            If TheExec.Flow.IsRunning = True Then 'ユーザーから呼ばれた場合
            
                ExecSettingList.Value = TheCondition.SelectExecSetting
                Call TheCondition.LoadCondition
            Else
            
                ExecSettingList.Value = pExecuteSettingName
            
            End If
        End If
    End If
    
    ChangeExecuteSetting = True

    Exit Function
    
EXECUTE_SETTING_NAME_ERROR:
    '切り替えマクログループ定義名は定義されていない
    Call MsgBox("""" & pExecuteSettingName & """" & " is not defined.", , "TestConditionManager")
    ChangeExecuteSetting = False
End Function

Private Sub CheckExecuteSetting()
'内容:
'   マクログループ定義の形式チェックを行う
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'

    On Error GoTo EXECUTE_SETTING_NAME_ERROR
        Call m_SettingNameColList.Item(m_SelecteExecSetting)
    On Error GoTo 0

    'マクログループ名リスト
    Dim SettingNameList As Collection
    Set SettingNameList = New Collection
    
    'データ読み込み開始位置
    Dim CurrentRange As Range
    Set CurrentRange = m_TestConditionSheet.Range(MACRO_SETTING_DEFAULT)
    
    While Not CurrentRange.Column > MACRO_SETTING_END_COLUMN + 1
        'グループ名が未入力の場合は無視する
        If CurrentRange.Value <> "" Then
        
            '6行目にグループ名があるかチェック
            If CurrentRange.Column = MACRO_SETTING_END_COLUMN + 1 Then '6行目にグループ名がある場合
                    Call RaiseError("Too many 'Execute Setting List' definitions (max:5).")
                Exit Sub
            End If
            
            'マクログループ名の重複チェック
            Dim SettingNameCount As Long
            For SettingNameCount = 1 To SettingNameList.Count Step 1
            
                If StrComp(CurrentRange.Value, SettingNameList.Item(SettingNameCount)) = 0 Then
                    Call RaiseError("""" & CurrentRange.Value & """" & " is duplicated.")
                    Exit Sub
                End If
            Next SettingNameCount
            
            'マクログループ名リストに追加
            SettingNameList.Add (CurrentRange.Value)
        End If
        
        Set CurrentRange = CurrentRange.offset(0, 1) '横に移動
    Wend
    Exit Sub
    
EXECUTE_SETTING_NAME_ERROR:
    '切り替えマクログループ定義名は定義されていない
    Call TheError.Raise(ERR_NUMBER, TypeName(Me), """" & m_SelecteExecSetting & """" & " is not defined.")
    
End Sub

Private Sub CheckDefaultLabel()
'内容:
'   マクログループ定義のDefaultラベルのエラーチェックを行う
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    'Defaultの定義名、位置が変更されていないかチェック
    Dim DefaultLabel As Range
    Set DefaultLabel = m_TestConditionSheet.Range(MACRO_SETTING_DEFAULT)
    If DefaultLabel.Value <> MACRO_SETTING_DEFAULT_TITLE Then
        Call TheError.Raise(ERR_NUMBER, TypeName(Me), "Cell $N$4 is reserved for label 'Default'.")
    End If
End Sub

Private Sub CreateExecSettingDataList()
'内容:
'   マクログループ定義にある各グループの有効無効リストを作成する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   マクログループ定義の値は全て、enable/disableのみ

    '初期化
    Set m_ExecSettingDataList = Nothing
    Set m_ExecSettingDataList = New Collection
    
    'マクログループ定義の左端(Default)からスタート
    Dim ExecSettingRange As Range
    Set ExecSettingRange = m_TestConditionSheet.Range(MACRO_SETTING_DEFAULT)
    While Not ExecSettingRange.Column > MACRO_SETTING_END_COLUMN
    
        Dim ExecSettingData As Collection
        Set ExecSettingData = New Collection
    
        'マクログループ名が未入力の場合は無視
        If ExecSettingRange.Value <> "" Then
            
            Dim ExecSettingDataRange As Range
            Set ExecSettingDataRange = ExecSettingRange.offset(1, 0)
            
            While Not IsEmpty(ExecSettingDataRange)
                
                Call ExecSettingData.Add(ExecSettingDataRange.Value)
                Set ExecSettingDataRange = ExecSettingDataRange.offset(1, 0)
                
            Wend
            
            'リストに追加(Keyはマクログループ定義名)
            Call m_ExecSettingDataList.Add(ExecSettingData, CStr(ExecSettingRange.Value))
        End If
        Set ExecSettingRange = ExecSettingRange.offset(0, 1)
    Wend
End Sub

Public Sub ChangeDefaultSetting()
'内容:
'   選択しているグループをDefaultに変更する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'

    'Validation実行時の処理
    Call Me.ChangeExecuteSetting(MACRO_SETTING_DEFAULT_TITLE)  'マクログループ定義を"Default"に設定する
    Call Me.LoadCondition

End Sub

Private Function IsActive() As Boolean
'内容:
'   ActiveSheetがTestConditionシートか判定
'
'備考:
'
'

    Dim readerManager As CWorkSheetReaderManager
    
    On Error GoTo ErrHandling

    Call XLibWkShtReaderUtility.CreateReaderManagerIfNothing
    Set readerManager = XLibWkShtReaderUtility.GetWkShtReaderManagerInstance
    
    If readerManager.GetActiveSheetName(eSheetType.shtTypeTestCond) = ActiveSheet.Name Then
        IsActive = True
    Else
        IsActive = False
    End If

    Set readerManager = Nothing
    Exit Function
ErrHandling:
    IsActive = False
    Set readerManager = Nothing
End Function








