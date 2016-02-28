VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVarBankInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   CVarBankクラスの割り込み用クラス
'目的:
'   CVarBank クラスの利用状況の情報収集用
'   インスタンスの Set でログ出力のオンオフ切り替える。
'   ログ出力オフ時のオーバーヘッドが無い。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月22日

Option Explicit

Implements IVarBank

Private m_Container As IVarBank                                        'このクラスが割り込む対象クラス
Private m_LogCollection As Collection                                       'ログ格納用

Private m_Counter As Long                                                   'ログ出力シリアル番号用

Private m_SaveLogFile As CFileOut                                           'ログファイル入出力制御用インスタンス
Private Const LOG_FILE_NAME As String = "EeeJOBLogTheVarBank.csv"            'このクラスのログ出力ファイル名

Public Sub Initialize(p As CVarBank)
    Set m_Container = p
End Sub

Public Function AsIVarBank() As IVarBank
    Set AsIVarBank = Me
End Function

Public Sub SaveLogFile(Optional ByVal fileName As String)
    
    '#################### ログ出力の初期設定 ####################
    Dim OutputString As String
    Dim RetFileName   As String
    Dim RetJobName   As String
    Dim i As Long
    
    If fileName = "" Then fileName = LOG_FILE_NAME
    
    '#################### ファイル出力 ####################
    '#################### ヘッダ生成 ####################
    Set m_SaveLogFile = New CFileOut
    Call m_SaveLogFile.SetFileName(fileName)
    RetFileName = TheExec.ExcelHandle.ActiveWorkbook.Path & "\" & TheExec.ExcelHandle.ActiveWorkbook.Name
    RetJobName = TheExec.CurrentJob                             'JobList の JobName より。無ければファイル名
    OutputString = "[RecordNumber], [RecordDate], [RecordTime], [TestInstanceName], [ExecuteMethod], [VariableName], [Value], [Status], [Fixed], [ReadOnly], [Comment]"
    '#################### ヘッダ出力 ####################
    Call m_SaveLogFile.OpenFile
    Print #m_SaveLogFile.fileNum, "[Test Program],[" & RetFileName & "]"
    Print #m_SaveLogFile.fileNum, "[Job],[" & RetJobName & "]"
    Print #m_SaveLogFile.fileNum, OutputString
    
    '#################### データ出力 ####################
    For i = 1 To m_LogCollection.Count
        Print #m_SaveLogFile.fileNum, m_LogCollection.Item(CStr(i))
    Next i
    Call m_SaveLogFile.CloseFile
    
    '#################### ログ初期化 ####################
    m_Counter = 0
    Set m_LogCollection = New Collection
    
End Sub

Private Property Get IVarBank_Value(VariableName As String) As Variant
    Dim i As Long       'TheError check用
    i = CheckTheError
    
    IVarBank_Value = m_Container.Value(VariableName)
    
    '#################### ログ登録 ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                'メソッド失敗した場合
            OutputString = LogFormat(VariableName, "-", "GetValue", "Error", "-", "-", "-")
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "GetValue", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), "-")
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Property

Private Property Let IVarBank_Value(VariableName As String, pValue As Variant)
    Dim i As Long       'TheError check用
    i = CheckTheError
    
    m_Container.Value(VariableName) = pValue
    
    '#################### ログ登録 ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                'メソッド失敗した場合
            OutputString = LogFormat(VariableName, CStr(pValue), "LetValue", "Error", "-", "-", "-")
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "LetValue", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), "-")
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Property

Private Property Get IVarBank_ReadOnly(VariableName As String) As Boolean
    IVarBank_ReadOnly = m_Container.ReadOnly(VariableName)
End Property

Private Property Let IVarBank_ReadOnly(VariableName As String, pReadOnly As Boolean)
    Dim i As Long       'TheError check用
    i = CheckTheError
    
    m_Container.ReadOnly(VariableName) = pReadOnly
    
    '#################### ログ登録 ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                'メソッド失敗した場合
            OutputString = LogFormat(VariableName, "-", "LetReadOnly", "Error", "-", CStr(pReadOnly), "-")
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "LetReadOnly", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), "-")
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Property

Private Property Get IVarBank_Fixed(VariableName As String) As Boolean
    IVarBank_Fixed = m_Container.Fixed(VariableName)
End Property

Private Property Let IVarBank_Fixed(VariableName As String, pFixed As Boolean)
    Dim i As Long       'TheError check用
    i = CheckTheError
    
    m_Container.Fixed(VariableName) = pFixed
    
    '#################### ログ登録 ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                'メソッド失敗した場合
            OutputString = LogFormat(VariableName, "-", "LetFixed", "Error", CStr(pFixed), "-", "-")
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "LetFixed", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), "-")
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Property

Private Property Get IVarBank_Comment(VariableName As String) As String
    IVarBank_Comment = m_Container.Comment(VariableName)
End Property

Private Property Let IVarBank_Comment(VariableName As String, pComment As String)
    Dim i As Long       'TheError check用
    i = CheckTheError
    
    m_Container.Comment(VariableName) = pComment
    
    '#################### ログ登録 ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                'メソッド失敗した場合
            OutputString = LogFormat(VariableName, "-", "LetComment", "Error", "-", "-", CStr(pComment))
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "LetComment", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), pComment)
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Property

Private Property Get IVarBank_Count() As Long
    IVarBank_Count = m_Container.Count
End Property

Private Function IVarBank_Add(VariableName As String, FirstValue As Variant, Optional FixedStatus As Boolean, Optional Comment As String) As Boolean
    Dim i As Long       'TheError check用
    i = CheckTheError
    
    Call m_Container.Add(VariableName, FirstValue, FixedStatus, Comment)
    
    '#################### ログ登録 ####################
    Dim OutputString As String
    With m_Container
        If i <> CheckTheError Then                'メソッド失敗した場合
            OutputString = LogFormat(VariableName, CStr(FirstValue), "Add", "Error", CStr(FixedStatus), "-", CStr(Comment))
        Else
            OutputString = LogFormat(VariableName, CStr(.Value(VariableName)), "Add", "-", CStr(.Fixed(VariableName)), CStr(.ReadOnly(VariableName)), CStr(.Comment(VariableName)))
        End If
    End With
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
End Function

Private Function IVarBank_IsExist(VariableName As String) As Boolean
    IVarBank_IsExist = m_Container.IsExist(VariableName)
End Function

Private Function IVarBank_Search(ByVal SearchPattern As String, Optional NotFlag As Boolean = False) As String
    IVarBank_Search = m_Container.Search(SearchPattern, NotFlag)
End Function

Private Sub IVarBank_Clear()
    
    Call m_Container.Clear

    '#################### ログ登録 ####################
    Dim OutputString As String
    OutputString = LogFormat("-", "-", "Clear", "-", "-", "-", "-")
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
    
End Sub

Private Sub IVarBank_AllClear()
    
    Call m_Container.AllClear

    '#################### ログ登録 ####################
    Dim OutputString As String
    OutputString = LogFormat("-", "-", "AllClear", "-", "-", "-", "-")
    Call m_LogCollection.Add(OutputString, CStr(m_Counter))
    
End Sub

'####### 以下インターフェイスと関係ないメソッド #######

Private Function LogFormat( _
    key As String, _
    Value As String, _
    Method As String, _
    status As String, _
    Fixed As String, _
    ReadOnly As String, _
    Comment As String _
    ) As String
'内容:
'   ログ入力用の情報を受けて、カンマ区切りで成型したString型で返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月26日
'パラメータ:
'   なし
'戻り値:
'   String
'
'注意事項:

    Dim p_InstanceName As String
    
    If TheExec.Flow.IsRunning Then
        p_InstanceName = TheExec.DataManager.InstanceName
    Else
        p_InstanceName = "-"
    End If
    
    m_Counter = m_Counter + 1
    
    '情報を返す。
    LogFormat = _
                m_Counter & "," & _
                CStr(DateTime.Date) & "," & _
                CStr(DateTime.TIME) & "," & _
                p_InstanceName & "," & _
                Method & "," & _
                key & "," & _
                Value & "," & _
                status & "," & _
                Fixed & "," & _
                ReadOnly & "," & _
                Comment

End Function

Private Function CheckTheError() As Long
'TheError.LastError の シリアル番号を返します。
'Nothing は0を返します。

    If TheError.LastError Is Nothing Then
        CheckTheError = 0
    Else
        CheckTheError = TheError.LastError.SerialNumber
    End If

End Function

Private Sub Class_Initialize()
'コンストラクタ

    m_Counter = 0
    Set m_Container = New CVarBank
    Set m_LogCollection = New Collection
    
End Sub

Private Sub Class_Terminate()
    m_Counter = 0
    Set m_Container = Nothing
    Set m_SaveLogFile = Nothing
    Set m_LogCollection = Nothing
End Sub

