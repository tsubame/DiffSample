VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSnapIP750"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@@@IP750EX Support
'概要:
'   IP750スナップショット機能を提供
'
'目的:
'   Ⅰ:スナップショット機能を取替え可能に
'   Ⅱ:ISnapshotの実装による機能の提供
'
'作成者:
'   SLSI今手
'
'XlibTesterInfoModをラッピング
'
'Code Checked
'Comment Checked
'

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1                '1CUL1Gボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>
#Const HDVIS_USE = 0      'HVDISリソースの使用　 0：未使用、0以外：使用  <PALSとEeeAuto両方で使用>
#Const CUB_UB_USE = 0    'CUB UBの設定          0：未使用、0以外：使用

Implements ISnapshot
Private mLogFileName As String
Private mSerialNo As Long

#If UB_SNAPSHOT_USE <> 0 Then
Private m_UBSetObj As CUtyBitController
#End If

#If ICUL1G_USE <> 0 Then
Private m_Icul1gSnapshot As CIcul1gSnapshotProvider
#End If

Private mLogOutputPlace As OUTPUT_PLACE
Private mOutputStatus As Boolean

'スナップショットボード
Enum SNAPSHOT_BOARD
    snapALL = 0
    snapAPMU
    snapPPMU
    snapDPS
    snapBPMU
    snapD_CH
    #If HDVIS_USE <> 0 Then
    snapHDVIS
    #End If
    
    #If ICUL1G_USE <> 0 Then
    snapICUL1G
    #End If
    
    'UB設定問題が解決するまではいったん封印する
    #If UB_SNAPSHOT_USE <> 0 Then
    snapAPMU_UB
    #If CUB_UB_USE <> 0 Then
    snapCUB_UB
    #End If
    #End If
End Enum

'スナップログ出力先の定義
Enum OUTPUT_PLACE
    snapTXT_FILE = 0
    snapIGXL_DATALOG = 1
    snapTXT_AND_DATALOG = 2
End Enum

'#Pass-Ok
Public Sub Initialize()
'内容:
'   スナップショット機能の初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:

    Call InitSnapIP750

End Sub

#If UB_SNAPSHOT_USE <> 0 Then
'#Pass-Ok UBの設定不整合問題が解決するまでは封印
Public Property Set UBObj(ByVal UBController As CUtyBitController)
'内容:
'    ユーティリティBit設定Object
'
'パラメータ:
'    [UBController]   In  ユーティリティBit設定Object
'
'戻り値:
'
'注意事項:

    Set m_UBSetObj = UBController

End Property
#End If

'#Pass-Ok
Public Property Let OutputSaveStatus(ByVal outPutMsg As Boolean)
'内容:
'    データログへのスナップショット取得メッセージの出力切り替え
'
'パラメータ:
'    [outputMsg]   In   TRUE:メッセージ出力ON、FALSE:メッセージ出力OFF
'
'戻り値:
'
'注意事項:
'

    mOutputStatus = outPutMsg

End Property

'#Pass-Ok
Public Function AsISnapshot() As ISnapshot
'内容:
'   ISnapshot OBJ取得
'
'パラメータ:
'
'戻り値:
'   ISnapshot オブジェクト
'
'注意事項:
'
    Set AsISnapshot = Me

End Function

'#Pass-Ok
Public Property Let SerialNumber(ByVal Number As Long)
'内容:
'    スナップショットログのラベルに表示するシリアル番号
'
'パラメータ:
'    [Number]   In  シリアル番号
'
'戻り値:
'
'注意事項:
'   GetSnapshotを実行する毎に値が現在値より+1されます。
'
    mSerialNo = Number

End Property

'#Pass-Ok
Public Property Get SerialNumber() As Long
'内容:
'    スナップショットログのラベルに表示するシリアル番号
'
'パラメータ:
'
'戻り値:
'   現在のシリアル番号値
'
'注意事項:
'   GetSnapshotを実行する毎に値が現在の値より+1されます。
'
    SerialNumber = mSerialNo

End Property

'#Pass-Ok
Public Property Let LogFileName(ByVal FullPathFileName As String)
'内容:
'    スナップショットログをTXTファイルに出力するときのファイル設定
'
'パラメータ:
'    [FullPathFileName]   In  ログファイル名
'
'戻り値:
'
'注意事項:
'   ファイル名はFullPathで指定してください。
'   (例) "D:\Jobs\Imx000\Log\SnapshotLog.txt"
'
    mLogFileName = FullPathFileName

End Property

'#Pass-Ok
Public Property Get LogFileName() As String
'内容:
'    スナップショットログをTXTファイルに出力するときのファイル設定
'
'パラメータ:
'
'戻り値:
'   ログ出力ファイル名
'
'注意事項:
'
    LogFileName = mLogFileName

End Property

'#Pass-Ok
Public Property Let OutputPlace(ByVal Place As OUTPUT_PLACE)
'内容:
'    スナップショットログの出力先の設定
'
'パラメータ:
'    [Place]   In  ログの出力場所
'
'戻り値:
'
'注意事項:
'   出力先にテキストファイルを指定した時に、LogFile
'   の設定が行われていない場合はエラーとなります｡
'   先にLogFileの設定を行ってください。
'
    If (Place = snapTXT_AND_DATALOG Or Place = snapTXT_FILE) And (mLogFileName = "") Then
        Call OutputErrMsg("SnapshotLog file setting is nothing")
    Else
        mLogOutputPlace = Place
    End If

End Property

'#Pass-Ok
Public Sub GetSnapshot(Optional ByVal Comment As String = "", Optional ByVal ip750board As SNAPSHOT_BOARD = snapALL)
'内容:
'    スナップショットを取得し結果を出力する
'
'パラメータ:
'    [Comment]   In   スナップショットログに表示するコメント
'    [ip750board] In  スナップショットを取得するボードの指定
'
'戻り値:
'
'注意事項:
'   実行後、スナップショットシリアル番号が +1 されます
'
    Dim outputflg As Long
    Dim LogFileName As String
    Dim snapIdMsg As String
    Dim dataLogMsg As String
    
    Call makeSnapshotSetting(outputflg, LogFileName)
    snapIdMsg = makeSnapIdMsg(Comment)
    
    Select Case ip750board
        Case snapALL
            Call GetTesterInfo(snapIdMsg, outputflg, LogFileName)
            #If ICUL1G_USE <> 0 Then
            Call m_Icul1gSnapshot.CreatePpmuInfo(snapIdMsg, outputflg, LogFileName)
            #End If

            'UB設定問題が解決するまではいったん封印する
            #If UB_SNAPSHOT_USE <> 0 Then
            If Not m_UBSetObj.AsAPMU Is Nothing Then
                Call GetApmuUbSnapshot(snapIdMsg, outputflg, LogFileName)
            End If
            #If CUB_UB_USE <> 0 Then
            If Not m_UBSetObj.AsCUB Is Nothing Then
                Call GetCubUbSnapshot(snapIdMsg, outputflg, LogFileName)
            End If
            #End If
            #End If
        
        Case snapAPMU
            Call CreateApmuInfo(snapIdMsg, outputflg, LogFileName)
        
        Case snapPPMU
            Call CreatePpmuInfo(snapIdMsg, outputflg, LogFileName)
        
        Case snapD_CH
            Call CreatePeInfo(snapIdMsg, outputflg, LogFileName)
        
        Case snapBPMU
            Call CreateBpmuInfo(snapIdMsg, outputflg, LogFileName)
        
        Case snapDPS
            Call CreateDpsInfo(snapIdMsg, outputflg, LogFileName)
        
        #If HDVIS_USE <> 0 Then
        Case snapHDVIS
            Call CreateHdvisInfo(snapIdMsg, outputflg, LogFileName)
        #End If
        
                
        #If ICUL1G_USE <> 0 Then
        Case snapICUL1G
            Call m_Icul1gSnapshot.CreatePpmuInfo(snapIdMsg, outputflg, LogFileName)
        #End If
                
        'UB設定問題が解決するまではいったん封印する
        #If UB_SNAPSHOT_USE <> 0 Then
        Case snapAPMU_UB
            If Not m_UBSetObj.AsAPMU Is Nothing Then
                Call GetApmuUbSnapshot(snapIdMsg, outputflg, LogFileName)
            Else
                Call OutputErrMsg("APMU UB object is Nothing! Please set UBObj Property")
            End If

        #If CUB_UB_USE <> 0 Then
        Case snapCUB_UB
            If Not m_UBSetObj.AsCUB Is Nothing Then
                Call GetCubUbSnapshot(snapIdMsg, outputflg, LogFileName)
            Else
                Call OutputErrMsg("CUB UB object is Nothing! Please set UBObj Property")
            End If
        #End If
        #End If
        
    
    End Select
    
    
    'スナップショットの取得を行ったことをデータログに表示するときの処理
    If mOutputStatus = True Then
        If Comment <> "" Then
            dataLogMsg = "#" & TheExec.DataManager.InstanceName & "." & Comment & "." & Me.SerialNumber
        Else
            dataLogMsg = "#" & TheExec.DataManager.InstanceName & "." & Me.SerialNumber
        End If
        TheExec.Datalog.WriteComment "TESTER_SS_SAVED = " & dataLogMsg
    End If
        
    'スナップショットのID用のシリアル番号 increment
    mSerialNo = mSerialNo + 1

End Sub

'UB設定問題が解決するまではいったん封印する
#If UB_SNAPSHOT_USE <> 0 Then
'#Pass-Ok
Private Sub GetApmuUbSnapshot(ByVal CommentMsg As String, Optional OutputDatalog As Long = 1, _
    Optional outputLogName As String = "")
'APMU UBのスナップショット取得用

    Dim snapMsg As String

    snapMsg = CommentMsg & ",[APMU_UB]," & "Board=" & m_UBSetObj.AsAPMU.BoardNumber & _
                            ",Hex=" & m_UBSetObj.AsAPMU.UBSnapshotInHex

    If OutputDatalog = 1 Then
        Call TheExec.Datalog.WriteComment(snapMsg)
        Call TheExec.Datalog.WriteComment("")
    End If
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, snapMsg)
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub

'#Pass-Ok
#If CUB_UB_USE <> 0 Then
Private Sub GetCubUbSnapshot(ByVal CommentMsg As String, Optional OutputDatalog As Long = 1, _
    Optional outputLogName As String = "")
'CUB UBのスナップショット取得用

    Dim snapMsg As String

    snapMsg = CommentMsg & ",[CUB_UB]," & "Hex=" & m_UBSetObj.AsCUB.UBSnapshotInHex

    If OutputDatalog = 1 Then
        Call TheExec.Datalog.WriteComment(snapMsg)
        Call TheExec.Datalog.WriteComment("")
    End If
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, snapMsg)
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub
#End If

#End If


'#Pass-Ok
Private Function makeSnapIdMsg(ByVal Comment As String) As String
'スナップショットのコメントメッセージ作成
        
    'スナップショットID用のラベル作成
    If Comment <> "" Then
        makeSnapIdMsg = "#" & TheExec.DataManager.InstanceName & "." & Comment & "." & mSerialNo
    Else
        makeSnapIdMsg = "#" & TheExec.DataManager.InstanceName & "." & mSerialNo
    End If

End Function

'#Pass-Ok
Private Sub makeSnapshotSetting(ByRef outputflg As Long, ByRef LogFileName As String)
'スナップショット実行前のプロパティの設定状況の確認

    Select Case mLogOutputPlace
        Case snapTXT_FILE
            outputflg = 0
            LogFileName = mLogFileName
                    
        Case snapIGXL_DATALOG
            outputflg = 1
            LogFileName = ""
        
        Case snapTXT_AND_DATALOG
            outputflg = 1
            LogFileName = mLogFileName
    End Select

End Sub

'#Pass-Ok
Private Sub mf_OutPutLog(ByVal LogFileName As String, outPutMessage As String)
'スナップショット用のログをファイルに出力する。
    Dim fp As Integer
    On Error GoTo OUT_PUT_LOG_ERR

    fp = FreeFile
    Open LogFileName For Append As fp
    Print #fp, outPutMessage
    Close fp

    Exit Sub

OUT_PUT_LOG_ERR:
    Call OutputErrMsg(LogFileName & " MsgOutPut Error" & " @mf_OutPutLog")
    Stop

End Sub

'#Pass-Ok
Private Sub InitSnapIP750()
'初期化
    mLogFileName = ""
    mLogOutputPlace = snapIGXL_DATALOG
    mSerialNo = 1
    
    #If UB_SNAPSHOT_USE <> 0 Then
    Set m_UBSetObj = Nothing
    #End If
    
    mOutputStatus = False

    #If ICUL1G_USE <> 0 Then
    Set m_Icul1gSnapshot = New CIcul1gSnapshotProvider
    #End If

End Sub

'#Pass-Ok
Private Sub Class_Initialize()
    
    Call InitSnapIP750

End Sub




'######################## 以下 ISnapshotの実装 #############################

'#Pass-Ok
Private Property Get ISnapshot_FunctionName() As String
    
    ISnapshot_FunctionName = "IP750Snapshot"

End Property

'#Pass-Ok
Private Sub ISnapshot_SaveSnapshot(Optional ByVal CommentMsg As String = "")
    
    Call Me.GetSnapshot(CommentMsg)

End Sub

