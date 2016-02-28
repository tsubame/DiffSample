VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgIDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   TheHdw.IDPのラッピング
'
'目的:
'   IDP関連のクラス、プロパティ、メソッドを統括する。
'
'作成者:
'   0145184004
'
Option Explicit

Private m_PlaneManagers As Collection
Private m_PlaneBank As IImgPlaneBank

'For MultiCapture
Private Const MC_MAX = 2
Private Type strMultiCapture
    acqIndex As Long
    acqNum As Long
    acqPlane As String
    acqDepth As IdpBitDepth
    acqPmd As String
End Type

Private m_CaptureIndex As Long
Private m_TransferIndex As Long
Private m_MultiCaptureParam(MC_MAX - 1) As strMultiCapture

Private m_ImgPlaneList As CImgPlaneList
Private m_ImgPlaneMapManager As CImgPlaneMapManager

'ログ出力用
Private Const LOG_FILE_NAME As String = "EeeJOBLogIDPAction.csv"                'このクラスのディフォルトで出力するログファイル名
Private m_LogFileName As String
Private m_SaveMode As Boolean

'For Kernel,LUT     'V3.0 No use
'Private m_Kernels As Collection
'Private m_LutDefinitions As Collection
'Private m_IsExistLUT As Boolean

Private m_LutManager As CImgLUTManager
Private m_KernelManager As CImgKernelManager

'/* ======= 2012/01/18 Add 0145206097 ここから ============== */
#If IDP_40IP02_USE <> 0 Then
Private Const SAME_COLOR As String = "SAME_COLOR"
#End If
'/* ======= 2012/01/18 Add 0145206097 ここまで ============== */

Friend Function DumpUsingPlane() As Collection
    Set DumpUsingPlane = New Collection

    Dim mng As IImgPlaneManager
    For Each mng In m_PlaneManagers
        Call mng.GetPlanes(DumpUsingPlane)
    Next mng
End Function

'=======2009/03/31 Add Maruyama ここから==============
Friend Function DumpPlaneGroup() As Collection
    Set DumpPlaneGroup = New Collection

    Dim mng As IImgPlaneManager
    For Each mng In m_PlaneManagers
        Call DumpPlaneGroup.Add(mng.Name)
    Next mng
End Function
'=======2009/03/31 Add Maruyama ここまで==============

Private Sub Class_Terminate()

    Set m_ImgPlaneMapManager = Nothing
    Set m_PlaneManagers = Nothing
    Set m_PlaneBank = Nothing
    Set m_ImgPlaneList = Nothing
    Set m_LutManager = Nothing
    Set m_KernelManager = Nothing

End Sub

Private Sub Class_Initialize()
    Call Initialize
End Sub

Public Sub Initialize()
'内容:
'   全体の初期化
'
    Call RemoveResources
    Call ResetMultiCapture
    
    Set m_ImgPlaneList = New CImgPlaneList

End Sub

'#FlagExpansion
Public Sub ResetTest()
'内容:
'   テストごとの初期化
'       PlaneBankのクリア(Fixedのプレーン以外)
'       各マネージャのフラグプレーンのクリア
'       マルチキャプチャのリセット
'
    Call m_PlaneBank.Clear
    Dim mng As IImgPlaneManager
    
    For Each mng In m_PlaneManagers
        Call mng.ClearAllFlag
    Next mng
    
    Call ResetMultiCapture
End Sub

Public Property Get PlaneManager(ByVal pName As Variant) As IImgPlaneManager
'内容:
'   プレーンマネージャ取得
'
'[pName]       IN   String型:     プレーン名(グループ名)
'
'備考:
'   エラー条件
'       1,指定された名前のマネージャがない
'
'   サンプルコード
'    With TheIDP.PlaneManager("vmcu")
'        ………
'    End With
'
On Error GoTo NOT_FOUND
    Set PlaneManager = m_PlaneManagers(pName)
    Exit Property
    
NOT_FOUND:
    Set PlaneManager = Nothing
    
    'エラー
    Call TheError.Raise(999, "PlaneManager", ErrMsgNameDoesntExist(pName))
    
End Property

Public Property Get PlaneBank() As IImgPlaneBank
'内容:
'   プレーンバンクを取得
'
'備考:
'   サンプルコード
'    With TheIDP.PlaneBank
'        ………
'    End With
'
    Set PlaneBank = m_PlaneBank
End Property

Public Property Get PMD(ByVal pName As String) As CImgPmdInfo
'内容:
'   指定した名前のPMDの情報を取得
'
'[pName]       IN   String型:     PMD名
'
'備考:
'   エラー条件
'       1,指定された名前のPMDがない
'
'   サンプルコード
'    With TheIDP.PMD("ZONE2D")
'        ………
'    End With
'
    Set PMD = PMD_(pName)
    
    'エラー
    If PMD Is Nothing Then
        Call TheError.Raise(999, "PMD", ErrMsgNameDoesntExist(pName))
    End If
    
End Property

Private Property Get PMD_(ByVal pName As String) As CImgPmdInfo
'PMDの補助
    Dim tmp As IImgPlaneManager
    
    For Each tmp In m_PlaneManagers
        Set PMD_ = tmp.PMD(pName)
        If Not PMD_ Is Nothing Then Exit Property
    Next tmp

End Property

Public Property Get PlaneMapManager() As CImgPlaneMapManager
'内容:
'   カラーマップテーブルのインスタンスを返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月11日
'パラメータ:
'   [pName] In/Out  1):
'戻り値:
'   CImgPlaneMapManager
'
'注意事項:
'
'

    Set PlaneMapManager = m_ImgPlaneMapManager
    
End Property

Public Property Get Kernel(ByVal pName As String) As CImgKernelInfo
'内容:
'   指定した名前のカーネルの情報を取得
'
'[pName]       IN   String型:     PMD名
'
'備考:
'   エラー条件
'       1,指定された名前のカーネルがない
'
'   サンプルコード
'    With TheIDP.Kernel
'        ………
'    End With
'
    On Error GoTo NOT_FOUND
    Set Kernel = m_KernelManager.Kernel(pName)
    
    Exit Property

NOT_FOUND:
    Call TheError.Raise(999, "Kernel", ErrMsgNameDoesntExist(pName))

End Property

Public Property Let SaveMode(ByVal mode As Boolean)
'内容:
'   画像確保関連の命令
'作成者:
'  tomoyoshi.takase
'作成日: 2010年6月2日
'パラメータ:
'   [mode]  In  1):
'戻り値:
'
'注意事項:
'   前提条件として、すでにCreatePMD済み
'
    '#### ProductionMode チェックして、設定するModeを決定する ####
    Dim YesOrNo As VbMsgBoxResult
    Dim tmpMode As Boolean
    
    If TheExec.RunMode = runModeProduction Then
        If mode = True Then
            YesOrNo = MsgBox("Now ProductionMode! But SetLogModeTheIDP is True!" & vbCrLf & _
                            "EeeJOB will output DataLog." & vbCrLf & _
                            "Please make sure." & vbCrLf & _
                            "When You don't want to output This MsgBox, Please change runModeProduction => runModeDebug" & vbCrLf & _
                            "" & vbCrLf & _
                            "Output DataLog?" & vbCrLf _
                            , vbYesNo + vbQuestion, "Confirm Output DataLog(SetLogModeTheIDP)")
            If YesOrNo = vbYes Then
                tmpMode = True
            Else
                tmpMode = False
            End If
        Else
            tmpMode = False
        End If
    Else
        tmpMode = mode
    End If
    
    '########### インスタンス群の切り替え ##########
    If m_SaveMode <> tmpMode Then           '現状のインスタンス構成と違うか否か
        
        Dim p_MngCollection As Collection
        Dim p_Mng As CImgPlaneManager
        Dim p_MngInterceptor As CImgPlaneManagerInterceptor
        Dim p_ImgBankInterceptor As CImgPlaneBankInterceptor
        Dim p_ImgFlagInterceptor As CImgFlagInterceptor
        
        Set p_MngCollection = New Collection
        
        If tmpMode = True Then
        '########### インスタンス群に割り込みクラスを入れる ##########
            
            '##### PlaneManager #####
            For Each p_Mng In m_PlaneManagers
                '##### Flag #####
                Call p_Mng.EnableInterceptor(tmpMode)
            
                '##### PlaneManager #####
                Set p_MngInterceptor = New CImgPlaneManagerInterceptor
                Set p_MngInterceptor.TrueInstance = p_Mng
                Call p_MngCollection.Add(p_MngInterceptor.AsIImgPlaneManager, p_MngInterceptor.AsIImgPlaneManager.Name)
                
            Next p_Mng
            Set m_PlaneManagers = Nothing
            Set m_PlaneManagers = p_MngCollection
            
            '##### PlaneBank #####
            Set p_ImgBankInterceptor = New CImgPlaneBankInterceptor
            Set p_ImgBankInterceptor.TrueInstance = m_PlaneBank
            Set m_PlaneBank = p_ImgBankInterceptor.AsIImgPlaneBank
            
        Else
        '######### インスタンス群から割り込みクラスを外す #########
            
            '##### PlaneManager #####
            For Each p_MngInterceptor In m_PlaneManagers
                Set p_Mng = p_MngInterceptor.TrueInstance
                Set p_MngInterceptor.TrueInstance = Nothing
                '##### Flag #####
                Call p_Mng.EnableInterceptor(tmpMode)
                
                '##### PlaneManager #####
                Call p_MngCollection.Add(p_Mng.AsIImgPlaneManager, p_Mng.AsIImgPlaneManager.Name)
                
            Next p_MngInterceptor
            Set m_PlaneManagers = Nothing
            Set m_PlaneManagers = p_MngCollection
            
            '##### PlaneBank #####
            Set p_ImgBankInterceptor = m_PlaneBank                  'Interceptorにキャスト
            Set m_PlaneBank = Nothing
            Set m_PlaneBank = p_ImgBankInterceptor.TrueInstance
            Set p_ImgBankInterceptor.TrueInstance = Nothing
        End If
        
        Set p_MngCollection = Nothing
        Set p_Mng = Nothing
        Set p_MngInterceptor = Nothing
        Set p_ImgBankInterceptor = Nothing
        Set p_ImgFlagInterceptor = Nothing
    End If
    
    If tmpMode = True Then
        TheExec.Datalog.WriteComment "Eee JOB Output Log! :TheIDP"
    End If
    
    m_SaveMode = tmpMode

End Property

Public Property Get SaveMode() As Boolean
    SaveMode = m_SaveMode
End Property

Public Property Let saveFileName(fileName As String)
    m_LogFileName = fileName
End Property

Public Property Get saveFileName() As String
    'ユーザー設定ファイル名が空ならディフォルトファイル名を返す
    If m_LogFileName <> "" Then
        saveFileName = m_LogFileName
    Else
        saveFileName = LOG_FILE_NAME
    End If
End Property

Public Function CreateBasePMD(ByVal pName As String, ByVal pBasePmdName As String, ByVal pWidth As Long, ByVal pHeight As Long, ByVal pColorMapTableName As String, ByVal pNumS16 As Long, ByVal pNumS32 As Long, ByVal pNumF32 As Long) As IImgPlaneManager
'内容:
'   ベースPMD作成、プレーンマネージャの登録
'   TheHdw.IDP.CreateBasePMDに準ずる。
'
'[pName]        IN  String型:   プレーン名
'[pBasePmdName] IN  String型:   ベースPMD名
'[pWidth]       IN  Long型:     幅
'[pHeight]      IN  Long型:     高さ
'[pColorMapTableName]    IN  String型:   カラーマップ定義テーブル名
'[pNumS16]      IN  Long型:     16Bit整数プレーンの枚数
'[pNumS32]      IN  Long型:     32Bit整数プレーンの枚数
'[pNumF32]      IN  Long型:     32Bit浮動小数プレーンの枚数
'
'備考:
'   エラー条件
'       1,pNameで指定されたプレーンが既に存在している時
'       2,pBasePmdNameで指定したPMDが既に存在している時
'       3.TheHdw.IDP.CreateBasePMDに対する不正なパラメータを入力した時
'
'   サンプルコード
'        Call TheIDP.CreateBasePMD("vmcu", "FULL", 800, 600, "(0,1)(7,2)", 10, 5, 3)
'

    If isExistingPlane(pName) Then
        'エラー
        Call TheError.Raise(999, "CreateBasePMD", ErrMsgNameAlreadyExists(pName))
        Exit Function
    End If
    
    Dim newManager As IImgPlaneManager
    Set newManager = New CImgPlaneManager
    
    Dim pColorMap As CImgPlaneMap
    Set pColorMap = m_ImgPlaneMapManager.planeMap(pColorMapTableName)
'    Call pColorMap.SetLimitPmdSize(pWidth, pHeight)
    
    Call newManager.CreateBasePMD(pName, pBasePmdName, pWidth, pHeight, pColorMap, pNumS16, pNumS32, pNumF32)
    Set pColorMap = Nothing
    
    Call m_PlaneManagers.Add(newManager, pName)
    
    Set CreateBasePMD = newManager
    
End Function

Public Function isExistingPlane(ByVal pName As String) As Boolean
'内容:
'   指定したプレーンマネージャがあるかどうかの確認
'
'[pName]       IN   String型:       プレーンマネージャの名前指定
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.isExistingPlane("vmcu")
'
    Dim tmp As IImgPlaneManager
    
    On Error GoTo NOT_FOUND
    Set tmp = m_PlaneManagers.Item(pName)
    isExistingPlane = True
    
    Exit Function
    
NOT_FOUND:
    isExistingPlane = False
    
End Function

Public Function isExistingPMD(ByVal pName As String) As Boolean
'内容:
'   指定したPMDがあるかどうかの確認
'
'[pName]       IN   String型:       PMDの名前指定
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.isExistingPMD("ZONE2D")
'

    If PMD_(pName) Is Nothing Then
        isExistingPMD = False
    Else
        isExistingPMD = True
    End If
    
End Function

Public Sub MultiCapture( _
    ByVal pPins As String, ByVal pNumFrame As Long, ByVal pAverageMode As IdpAverageMode, ByVal pAcqMode As IdpAcquireMode, _
    ByVal pPlaneGroup As String, Optional ByVal pBitDepth As IdpBitDepth = idpDepthS16, Optional ByVal pZone As String = "", _
    Optional ByVal pNumFields As Long = 2, Optional ByVal pSite As Long = -1 _
)
'内容:
'   マルチキャプチャを実行
'
'[pPins]        IN  String型:           対象のピンリスト
'[pNumFrame]    IN  Long型:             撮り込みフレーム数指定
'[pAverageMode] IN  IdpAverageMode型:   アベレージモード指定(加算 or 平均)
'[pAcqMode]     IN  IdpAcquireMode型:   撮り込みモード指定(インターレース or ノンインターレース)
'[pPlaneGroup]  IN  String型:           撮り込み先のプレーンマネージャ名
'[pBitDepth]    IN  IdpBitDepth型:      撮り込み先のビット深さ
'[pZone]        IN  String型:           撮り込み先のゾーン名
'[pNumFields]   IN  Long型:             フィールド数(インターレースの時)
'[pSite]        IN  Long型:             サイト
'
'備考:
'   異なるプレーン、ゾーンは指定不可
'   そういうものが出るのであれば、対応を考える。
'
'   サンプルコード
'    Call TheIDP.MultiCapture("VOUT1,VOUT2", 10, idpAverage, idpAcqNonInterlace, "vmcu", idpDepthS16, "FULL")
'
    Dim pinArray() As String
    pinArray = Split(Replace(pPins, " ", ""), ",")
    
    With m_MultiCaptureParam(m_CaptureIndex)
        If .acqPlane <> "" Then
            Call TheError.Raise(999, "MultiCapture", "Transfering doesn't end")
        End If
        
        .acqNum = UBound(pinArray) + 1
        .acqPlane = pPlaneGroup
        .acqDepth = pBitDepth
        .acqPmd = pZone
'
'        Dim tmpPlane As CImgPlane
'        Set tmpPlane = TheIDP.PlaneManager(.acqPlane).GetFreePlane(.acqDepth)
'        If .acqPmd <> "" Then
'            Call tmpPlane.SetPMD(.acqPmd)
'        End If
'
'        Dim pmdArray() As String
'        ReDim pmdArray(.acqNum - 1)
'
'        Dim i As Long
'        For i = 0 To UBound(pmdArray)
'            pmdArray(i) = tmpPlane.CurrentPmdName
'        Next i
'
'        Call TheHdw.IDP.MultiCaptureEx(pinArray, UBound(pinArray) + 1, pNumFrame, pAverageMode, pAcqMode, .acqIndex, pNumFields, pSite, pmdArray)
        
'        If .acqPmd <> "" Then
'            Dim pmdArray() As String
'            Dim i As Long
'
'            ReDim pmdArray(.acqNum - 1)
'            For i = 0 To UBound(pmdArray)
'                pmdArray(i) = .acqPmd & Postfix(.acqDepth)
'            Next i
'
'            Call TheHdw.IDP.MultiCaptureEx(pinArray, UBound(pinArray) + 1, pNumFrame, pAverageMode, pAcqMode, .acqIndex, pNumFields, pSite, pmdArray)
'        Else
'            Call TheHdw.IDP.MultiCaptureEx(pinArray, UBound(pinArray) + 1, pNumFrame, pAverageMode, pAcqMode, .acqIndex, pNumFields, pSite)
'        End If

        If .acqPmd = "" Then .acqPmd = TheIDP.PlaneManager(.acqPlane).BasePMD.Name
        
        Dim pmdArray() As String
        Dim i As Long

        ReDim pmdArray(.acqNum - 1)
        For i = 0 To UBound(pmdArray)
            pmdArray(i) = .acqPmd & Postfix(.acqDepth)
        Next i
        
        Call TheHdw.IDP.MultiCaptureEx(pinArray, UBound(pinArray) + 1, pNumFrame, pAverageMode, pAcqMode, .acqIndex, pNumFields, pSite, pmdArray)
    End With
    
    m_CaptureIndex = (m_CaptureIndex + 1) Mod MC_MAX
    
End Sub

Public Sub MultiCaptureTransfer(ByRef pPlaneArray() As CImgPlane)
'内容:
'   マルチキャプチャで取り込んだ画を転送。
'
'[pPlaneArray]  OUT     CImgPlane型:        撮り込み画像を保存するプレーンの配列
'
'備考:
'   必要なパラメータはマルチキャプチャ時に与えられているので、ここではそれにしたがって結果をとってくるだけ
'   pPlaneArrayはCImgPlane型の動的配列を渡す。
'
'   サンプルコード
'    Dim rawPlanes() As CImgPlane
'    Call TheIDP.MultiCaptureTransfer(rawPlanes)
'
    
    With m_MultiCaptureParam(m_TransferIndex)
        Dim planeNameArray() As String
        Dim i As Long
        ReDim pPlaneArray(.acqNum - 1)
        ReDim planeNameArray(.acqNum - 1)
        For i = 0 To UBound(planeNameArray)
            Set pPlaneArray(i) = PlaneManager(.acqPlane).GetFreePlane(.acqDepth)
'            If .acqPmd <> "" Then
                Call pPlaneArray(i).SetPMD(.acqPmd)
'            End If
            planeNameArray(i) = pPlaneArray(i).Name
        Next i
        
        Call TheHdw.IDP.MultiCaptureTransferEx(planeNameArray, .acqNum, .acqIndex)
    End With
    
    Dim emptyParam As strMultiCapture
    m_MultiCaptureParam(m_TransferIndex) = emptyParam
    
    m_TransferIndex = (m_TransferIndex + 1) Mod MC_MAX

End Sub

Private Function Postfix(ByVal pDepth As IdpBitDepth) As String

    Select Case pDepth
    Case idpDepthS16
        Postfix = "_S"
        
    Case idpDepthS32
        Postfix = "_L"
    
    Case idpDepthF32
        Postfix = "_F"
    
    End Select
    
End Function

'************************** 以下TheHdw.IDPのプロパティ **************************
Public Property Get CaptureTimeOut() As Double
'内容:
'   キャプチャータイムアウト時間取得
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.CaptureTimeOut
'
    CaptureTimeOut = TheHdw.IDP.CaptureTimeOut
End Property

Public Property Let CaptureTimeOut(ByVal RHS As Double)
'内容:
'   キャプチャータイムアウト時間設定
'
'備考:
'   サンプルコード
'    TheIDP.CaptureTimeOut = 10
'
    On Error GoTo ERROR_IGXL
    TheHdw.IDP.CaptureTimeOut = RHS
    Exit Property

ERROR_IGXL:
    Call ErrMsgIGXL("CaptureTimeOut")
    
End Property

Public Property Get IEEE1394Performance() As Idp1394PerformanceType
'内容:
'   データ転送のパフォーマンス状態取得
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.IEEE1394Performance
'
    IEEE1394Performance = TheHdw.IDP.IEEE1394Performance
End Property

Public Property Let IEEE1394Performance(ByVal RHS As Idp1394PerformanceType)
'内容:
'   データ転送のパフォーマンス状態設定
'
'備考:
'   サンプルコード
'    TheIDP.IEEE1394Performance = idp1394PerformanceHigh
'
    On Error GoTo ERROR_IGXL
    TheHdw.IDP.IEEE1394Performance = RHS
    Exit Property

ERROR_IGXL:
    Call ErrMsgIGXL("IEEE1394Performance")
    
End Property

Public Property Get IsCapturing() As Boolean
'内容:
'   キャプチャ状態かどうか取得
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.IsCapturing
'
    IsCapturing = TheHdw.IDP.IsCapturing
End Property

Public Property Get IsTransfering() As Boolean
'内容:
'   転送状態かどうか取得
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.IsTransfering
'
    IsTransfering = TheHdw.IDP.IsTransfering
End Property

Public Property Get TransferMode() As IdpTransferMode
'内容:
'   転送モード取得
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.TransferMode
'
    TransferMode = TheHdw.IDP.TransferMode
End Property

Public Property Let TransferMode(ByVal RHS As IdpTransferMode)
'内容:
'   転送モード設定
'
'備考:
'   サンプルコード
'    TheIDP.TransferMode = idpTransferHighest
'
    On Error GoTo ERROR_IGXL
    TheHdw.IDP.TransferMode = RHS
    Exit Property

ERROR_IGXL:
    Call ErrMsgIGXL("TransferMode")
    
End Property

'*************************** 以下TheHdw.IDPのメソッド ***************************
Public Sub CreateIndexLUT( _
    ByVal pLUTName As String, ByVal pInVal1 As Long, ByVal pInVal2 As Long, _
    ByVal pOutVal1 As Long, ByVal pOutVal2 As Long, Optional ByVal pValidBit As Long = 12 _
)
'内容:
'   LUTの作成
'
'[pLutName]    IN   String型:       LUT名
'[pInVal1]     IN   Long型:         入力値1
'[pInVal2]     IN   Long型:         入力値2
'[pOutVal1]    IN   Long型:         出力値1
'[pOutVal2]    IN   Long型:         出力値2
'[pValidBit]   IN   Long型:         有効ビット幅
'
'備考:
'   予約となっているパラメータは省略。
'
'   サンプルコード
'    Call TheIDP.CreateIndexLUT("ABS", -1024, 0, 1024, 0, 11)
'
'2010/12/03 メソッド実行をCImgLUTManagerに移動

    Call m_LutManager.CreateIndexLUT(pLUTName, pInVal1, pInVal2, pOutVal1, pOutVal2, pValidBit)
    
End Sub

'Public Sub TerminateLUT()
''内容:
''   LUT定義の終端処理
''   LUTの定義が終了したことを明示して、定義終了後の同名の再定義は許可しません。
''   終端処理前は再定義(上書き)可能です。
''作成者:
''  tomoyoshi.takase
''作成日: 2010年11月30日
''パラメータ:
''   なし
''戻り値:
''
''注意事項:
''
''サンプルコード
''    Call TheIDP.CreateIndexLUT("ABS", -1024, 0, 1024, 0, 11)
''    Call TheIDP.CreateIndexLUT("ABS", 0, 1024, 0, 1024, 11)
''    Call TheIDP.TerminateLUT
''
'
'    Call m_LutManager.TerminateLUT
'
'End Sub
'
Public Sub CreateKernel( _
    pKernelName As String, ByVal pKernelType As IdpKernelType, _
    ByVal pNumCol As Long, ByVal pNumRow As Long, ByVal pShiftR As Long, ByVal pKernelVal As String, Optional ByVal pComment As String = "" _
)
'内容:
'   ディジタルフィルタのカーネル作成
'
'[pKernelName] IN   String型:           カーネル名
'[pKernelType] IN   IdpKernelType型:    カーネルタイプ
'[pNumCol]     IN   Long型:             幅
'[pNumRow]     IN   Long型:             高さ
'[pShiftR]     IN   Long型:             右シフト数
'[pKernelVal]  IN   String型:           カーネル係数定義
'
'備考:
'   サンプルコード
'    Call TheIDP.CreateKernel("HPF", idpKernelInteger, 3, 3, 2, "0 -1 0 -1 4 -1 0 -1 0")
'
'2010/12/03 メソッド実行をCImgKernelManagerに移動
    
    Call m_KernelManager.CreateKernel(pKernelName, pKernelType, pNumCol, pNumRow, pShiftR, pKernelVal, pComment)
    
End Sub

Public Sub PostPatStart()
'内容:
'   Acquireの後でパターンを走らせる場合に使う。
'
'備考:
'   PrePatStartと対で使う
'
'   サンプルコード
'    Call TheIDP.PostPatStart
'
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.PostPatStart
    Exit Sub
    
ERROR_IGXL:
    Call ErrMsgIGXL("PostPatStart")
    
End Sub

Public Sub PrePatStart()
'内容:
'   Acquireの後でパターンを走らせる場合に使う。
'
'備考:
'   PostPatStartと対で使う
'
'   サンプルコード
'    Call TheIDP.PrePatStart
'
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.PrePatStart
    Exit Sub

ERROR_IGXL:
    Call ErrMsgIGXL("PrePatStart")
    
End Sub

Public Sub RemoveResources()
'内容:
'   PMD,Kernel,LUTの定義情報を削除。CImgIDPの初期化
'
'備考:
'
'   サンプルコード
'    Call TheIDP.RemoveResources
'
    Call TheHdw.IDP.RemoveResources
    Set m_ImgPlaneMapManager = New CImgPlaneMapManager
    Set m_PlaneManagers = New Collection
    
    Dim p_PlaneBank As CImgPlanePublicBank
    Set p_PlaneBank = New CImgPlanePublicBank
    Call p_PlaneBank.Init(TheParameterBank)
    Set m_PlaneBank = p_PlaneBank
    
'    Set m_PlaneBank = New CImgPlaneBank

'V3.0 No use
'    Set m_Kernels = New Collection
'    m_IsExistLUT = False
    Set m_LutManager = New CImgLUTManager
    Set m_KernelManager = New CImgKernelManager

End Sub

Public Sub ResetMultiCapture()
'内容:
'   マルチキャプチャのリセット
'
'備考:
'   サンプルコード
'    Call TheIDP.ResetMultiCapture
'
    Erase m_MultiCaptureParam
    m_CaptureIndex = 0
    m_TransferIndex = 0
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.ResetMultiCapture
    Exit Sub

ERROR_IGXL:
    Call ErrMsgIGXL("ResetMultiCapture")
    
End Sub

Public Sub WaitCaptureCompletion()
'内容:
'   キャプチャ終了まで待機
'
'備考:
'   サンプルコード
'    Call TheIDP.WaitCaptureCompletion
'
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.WaitCaptureCompletion
    Exit Sub

ERROR_IGXL:
    Call ErrMsgIGXL("WaitCaptureCompletion")
    
End Sub

Public Sub WaitTransferCompletion()
'内容:
'   転送終了まで待機
'
'備考:
'   サンプルコード
'    Call TheIDP.WaitTransferCompletion
'
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.WaitTransferCompletion
    Exit Sub
    
ERROR_IGXL:
    Call ErrMsgIGXL("WaitTransferCompletion")
    
End Sub

Private Function ErrMsgNameDoesntExist(ByVal pName As String)
    ErrMsgNameDoesntExist = """" & pName & """ doesn't exist."
End Function

Private Function ErrMsgNameAlreadyExists(ByVal pName As String)
    ErrMsgNameAlreadyExists = """" & pName & """ already exists."
End Function

Private Sub ErrMsgIGXL(ByVal pErrSource As String)
    Call TheError.Raise(999, pErrSource, Err.Description)
End Sub

Public Sub DisplayUsingPlaneList()
'内容:
'   使用中のプレーンのリストを表示する
'
    Dim dsp As New UsingPlaneDisplay
    Call dsp.Display
    
End Sub

Public Property Get LUTManager() As CImgLUTManager
    Set LUTManager = m_LutManager
End Property

Public Property Get KernelManager() As CImgKernelManager
    Set KernelManager = m_KernelManager
End Property

Public Property Get PlaneList() As CImgPlaneList
    Set PlaneList = m_ImgPlaneList
End Property

Public Function PlaneManagerCount() As Long
    PlaneManagerCount = m_PlaneManagers.Count
End Function

Public Function IsExistKernel() As Boolean
    IsExistKernel = CBool(m_KernelManager.Count)
End Function
Public Function IsExistLUT() As Boolean
    IsExistLUT = CBool(m_LutManager.Count)
End Function
Public Function IsRemovedResources() As Long
    IsRemovedResources = Not (IsExistLUT Or IsExistKernel)
End Function

'/* ======= 2012/01/18 Add 0145206097 ここから ============== */
#If IDP_40IP02_USE <> 0 Then
Public Property Let ReadWriteFileExTimeout(ByVal pTimeout As Long)
'内容:
'   WaitReadWriteFileExCompletionで使用されるタイムアウト時間の設定。
'
'パラメータ:
'   [pTimeout]  In  タイムアウトとなる時間
'
'備考:
'   タイムアウトの設定単位は「秒」。初期値は180秒。
'
'   サンプルコード
'   Call TheIDP.ReadWriteFileExTimeout = 100
'
    Const FUNCTION_NAME = "ReadWriteFileExTimeout"
    On Error GoTo ErrHandler
    TheHdw.IDP.ReadWriteFileExTimeout = pTimeout
    Exit Property

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Property

Public Property Get ReadWriteFileExTimeout() As Long
'内容:
'   WaitReadWriteFileExCompletionで使用されるタイムアウト時間の取得。
'
'戻り値:
'   タイムアウトとなる時間
'
'備考:
'   タイムアウトの設定単位は「秒」。初期値は180秒。
'
'   サンプルコード
'   Debug.Print TheIDP.ReadWriteFileExTimeout
'
    ReadWriteFileExTimeout = TheHdw.IDP.ReadWriteFileExTimeout
End Property

Public Sub WaitReadWriteFileExCompletion()
'内容:
'   ReadFileEx、またはWriteFileExを実行した後、ファイルの読み込み、または画像データの書き込みが終了するまで待つ
'
'備考:
'   予め設定されたタイムアウト時間を超えると次の処理に移る。
'
'   サンプルコード
'   Call TheIDP.WaitReadWriteFileExCompletion
'
    Const FUNCTION_NAME = "ReadWriteFileExTimeout"
    On Error GoTo ErrHandler
    Call TheHdw.IDP.WaitReadWriteFileExCompletion
    Exit Sub

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Sub

Public Sub TemporalStatisticsAll( _
    ByRef pDst() As CImgPlane, ByRef pSrc() As CImgPlane, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "" _
)
'内容:
'   複数のプレーン2~Nの各ピクセルごとに串刺し統計演算を行い、演算結果をプレーンに入れて返す。
'
'[pDst()]   In  演算結果保存先のプレーン配列
'[pSrc()]   In  統計演算対象のプレーン配列
'[pColorD]  In  演算結果保存先のプレーン色指定
'[pColorS]  In  統計演算対象プレーンの色指定
'[pFlgName] In  フラグ名
'
'注意事項:
'   IG-XL制約事項:以下のエラー検出はIG-XL APIに委ねる
'   【入力プレーン制約】
'   ①入力プレーン配列の要素数は2以上100以下であること
'   ②入力プレーン配列内にNothingが含まれていないこと
'   ③入力プレーンが全て同じBit長であること
'   ④入力プレーンが全て同じBasePMDであること
'   ⑤入力プレーンに結果格納プレーンが含まれていないこと
'
'   【結果格納プレーン制約】
'   ⑥入力プレーン配列の要素数は3であること
'   ⑦結果格納プレーンが②～⑤を満たすこと
'   ⑧結果格納プレーンのBit長が入力プレーンのBit長と同じか、それより深いこと
'　（例：入力が32Bit整数型であれば、出力はそれと同じか、32Bit浮動小数点型であることが必要）
'
'   【カラー指定制約】
'   ⑨入力プレーンの色指定がColorAll、ColorFlatの場合、自分自身の色指定もそれぞれColorAll、ColorFlatであること
'   ⑩色指定がColorAllの場合は自分自身が入力プレーンの持つ色を全て含んでいること
'　注）⑩以外の場合、Color Map DefinitionでのidpColorTypeの割り当てによっては演算可能なケースが発生するが動作保障外なので注意
'   ⑪入力プレーンの色指定が単独色の場合、自分自身の色指定も単独色であること（指定する色は異なっても構わない））
'
'備考:
'   サンプルコード
'   Dim src(2) As CImgPlane
'   Set src(0) = srcPlane1
'   Set src(1) = srcPlane2
'   Set src(2) = srcPlane3
'
'   Dim dst(2) As CImgPlane
'   Set dst(0) = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'   Set dst(1) = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'   Set dst(2) = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'   Call TheIDP.TemporalStatisticsAll(dst, src)
'
    Const FUNCTION_NAME = "TemporalStatisticsAll"
    On Error GoTo ErrHandler
    '/* 各プレーン配列の最小インデックス番号の要素だけチェック */
    '/* CImgPlane型でなかったらエラー（Nothing含む）*/
    '/* プレーン配列に、CImgPlane型でないオブジェクトを一度バリアントに突っ込んでCImgPlane型にキャストされた変数が渡されるとなぜか型の不一致が起こらないケースに対応 */
    If TypeName(pSrc(LBound(pSrc))) <> "CImgPlane" Then
        On Error GoTo 0
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pSrc()"))
    End If
    If TypeName(pDst(LBound(pDst))) <> "CImgPlane" Then
        On Error GoTo 0
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pDst()"))
    End If

    '/* カラータイプ型変換 */
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    '/* プレーンマップは各プレーン配列の先頭を参照する */
    pIdpColorTypeD = GetIdpColorType(pDst(LBound(pDst)).planeMap, pColorD)
    pIdpColorTypeS = GetSrcIdpColorType(pSrc(LBound(pSrc)).planeMap, pColorS, pColorD)

    '/* 串刺し演算 */
    Call TheHdw.IDP.TemporalStatistics(GetArrayPlaneNames(pSrc), pIdpColorTypeS, GetArrayPlaneNames(pDst, True), pIdpColorTypeD, idpTemporalStatisticsAll, _
                                                        GetOptFlagName(pDst(LBound(pDst)), pFlgName), GetOptFlagBit(pDst(LBound(pDst)), pFlgName))
    Exit Sub

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Sub

Public Sub TemporalMinMaxAll( _
    ByRef pDst() As CImgPlane, ByRef pSrc() As CImgPlane, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlagPlane As Variant = "", Optional ByRef pFlgName As String = "" _
)
'内容:
'   複数のプレーン2~Nの各ピクセルから最大及び最小値を持つピクセルを抽出し、ピクセル値を結果プレーンに入れて返す。
'
'[pDst()]       In  抽出結果保存先のプレーン配列
'[pSrc()]       In  最大及び最小値抽出対象のプレーン配列
'[pColorD]      In  抽出結果保存先のプレーン色指定
'[pColorS]      In  最大及び最小値抽出対象プレーンの色指定
'[pFlgPlane]    In  出力フラグプレーン
'[pFlgName]     In  フラグ名
'
'注意事項:
'   IG-XL制約事項:以下のエラー検出はIG-XL APIに委ねる
'   【入力プレーン制約】
'   ①入力プレーン配列の要素数は2以上32以下であること
'   ②入力プレーン配列にNothingが含まれていないこと
'   ③入力プレーンが全て同じBit長であること
'   ④入力プレーンが全て同じBasePMDであること
'   ⑤入力プレーンに結果格納プレーンが含まれていないこと
'
'   【結果格納プレーン制約】
'   ⑥入力プレーン配列の要素数は2であること
'   ⑦結果格納プレーンが②～⑤を満たすこと
'   ⑧結果格納プレーンのBit長が入力プレーンのBit長と同じであること
'
'   【カラー指定制約】
'   ⑨入力プレーンの色指定がColorAll、ColorFlatの場合、結果格納プレーンの色指定もそれぞれColorAll、ColorFlatであること
'   ⑩入力プレーンの色指定がColorAllの場合、結果格納プレーンが入力プレーンの持つ色を全て含んでいること
'　注）⑩以外の場合、Color Map DefinitionでのidpColorTypeの割り当てによっては演算可能なケースが発生するが動作保障外なので注意
'   ⑪入力プレーンの色指定が単独色の場合、自分自身の色指定も単独色であること（指定する色は異なっても構わない）
'
'   【出力フラグプレーン制約】
'   ⑩出力フラグプレーンを利用する場合は、2次元配列が入力されていること（要素番号0をMin、要素番号1をMax用に割り当てる）
'   ⑪出力フラグプレーンを利用する場合は、入力プレーンが16枚以下の場合は1枚、17枚以上であれば2枚が最大値と最小値用にそれぞれ用意されていること
'   ⑫出力フラグプレーンを利用する場合は、Bit長が16Bit整数型であること
'   ⑬出力フラグプレーンを利用する場合は、入力プレーンに含まれていないこと
'   ⑭出力フラグプレーンを利用する場合は、結果格納プレーンに含まれていないこと
'   ⑮出力フラグプレーンを利用する場合は、その配列が②～④を満たすこと
'
'備考:
'   サンプルコード
'   Dim src(2) As CImgPlane
'   Set src(0) = srcPlane1
'   Set src(1) = srcPlane2
'   Set src(2) = srcPlane3
'
'   Dim dst(1) As CImgPlane
'   Set dst(0) = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'   Set dst(1) = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'   Call TheIDP.TemporalMinMaxAll(dst, src)
'
    Const FUNCTION_NAME = "TemporalStatisticsAll"
    On Error GoTo ErrHandler
    '/* 各プレーン配列の最小インデックス番号の要素だけチェック */
    '/* CImgPlane型でなかったらエラー（Nothing含む）*/
    '/* プレーン配列に、CImgPlane型でないオブジェクトを一度バリアントに突っ込んでCImgPlane型にキャストされた変数が渡されるとなぜか型の不一致が起こらないケースに対応 */
    If TypeName(pSrc(LBound(pSrc))) <> "CImgPlane" Then
        On Error GoTo 0
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pSrc()"))
    End If
    If TypeName(pDst(LBound(pDst))) <> "CImgPlane" Then
        On Error GoTo 0
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pDst()"))
    End If

    '/* カラータイプ型変換 */
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    '/* プレーンマップは各プレーン配列の先頭を参照する */
    pIdpColorTypeD = GetIdpColorType(pDst(LBound(pDst)).planeMap, pColorD)
    pIdpColorTypeS = GetSrcIdpColorType(pSrc(LBound(pSrc)).planeMap, pColorS, pColorD)

    '/* 串刺し演算 */
    Call TheHdw.IDP.TemporalMinMax(GetArrayPlaneNames(pSrc), pIdpColorTypeS, GetArrayPlaneNames(pDst, True), pIdpColorTypeD, idpTemporalMinMaxAll, _
                                                    GetArrayOptPlaneNames2D(pFlagPlane), GetOptFlagName(pDst(LBound(pDst)), pFlgName), GetOptFlagBit(pDst(LBound(pDst)), pFlgName))
    Exit Sub

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Sub

'/* 以下ユーティリティ関数 */
Private Function GetArrayOptPlaneNames2D(ByRef pOptSrc As Variant) As Variant
'内容:
'   オプショナルのプレーン情報（文字列、プレーンオブジェクト配列）を文字列情報に変換する。
'   CImgPlaneクラスのGetArrayOptPlaneNamesメソッド2次元配列限定版
'
'[pOptSrc]           In  オプションプレーン指定
'
'戻り値:
'   空白、またはプレーン名の文字列配列
'
'概要:
'   ①入力は文字列、プレーンオブジェクト配列以外はエラーにする
'   ②文字列入力は空白以外はエラーにする
'   ③プレーンが書込み禁止だったらエラーにする
'
'/* 2012/3/1 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetArrayOptPlaneNames2D"
    On Error GoTo ErrHandler
    '/* 入力フラグプレーンが文字列だった場合 */
    If TypeName(pOptSrc) = "String" Then
        If pOptSrc = "" Then
            GetArrayOptPlaneNames2D = ""
        '/* 文字列入力は空白以外は許可しない */
        Else
            GoTo errTypeMismatch
        End If
    '/* 入力フラグプレーンがオブジェクト型配列だった場合 */
    ElseIf TypeName(pOptSrc) = "Object()" Then
        '/* CImgPlane型配列でない場合はキャストでエラー */
        Dim tmpNames2D() As String
        ReDim tmpNames2D(UBound(pOptSrc, 1) - LBound(pOptSrc, 1), UBound(pOptSrc, 2) - LBound(pOptSrc, 2))
        Dim dIndex As Long
        For dIndex = 0 To UBound(tmpNames2D, 1)
            Dim tmpSrc() As CImgPlane
            ReDim tmpSrc(UBound(tmpNames2D, 2))
            Dim pIndex As Long
            For pIndex = 0 To UBound(tmpSrc)
                Set tmpSrc(pIndex) = pOptSrc(LBound(pOptSrc, 1) + dIndex, LBound(pOptSrc, 2) + pIndex)
            Next pIndex
            '/* 書込み禁止チェックをONにする */
            Dim tmpNames() As String
            tmpNames = GetArrayPlaneNames(tmpSrc, True)
            For pIndex = 0 To UBound(tmpNames)
                tmpNames2D(dIndex, pIndex) = tmpNames(pIndex)
            Next pIndex
        Next dIndex
        GetArrayOptPlaneNames2D = tmpNames2D
    '/* それ以外の変数型の入力は許可しない */
    Else
        GoTo errTypeMismatch
    End If
    Exit Function
ErrHandler:
    If Err.Number = 9 Then
        Call TheError.Raise(Err.Number, FUNCTION_NAME, ErrMsgInvalidArrayType)
    ElseIf Err.Number = 13 Then
        Call TheError.Raise(Err.Number, FUNCTION_NAME, ErrMsgTypeMismatch("pOptSrc"))
    Else
        Call TheError.Raise(Err.Number, FUNCTION_NAME, Err.Description)
    End If
errTypeMismatch:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pOptSrc"))
End Function

'#FlagExpansion
Private Function GetOptFlagName(ByRef pSrc As CImgPlane, ByVal pFlgName As String) As String
'内容:
'   このプレーンの共用フラグプレーン名を返す
'
'[pFlgName] In  フラグ名
'
'戻り値:
'   空白文字列または共用フラグプレーン名
'
'概要:
'   空白文字列が入力された場合はそのまま空白文字列を返す
'   CImgPlaneクラスのGetOptFlagNameメソッドに引数を追加しただけ
'
'/* 2012/3/1 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetOptFlagName"
    On Error GoTo ErrHandler
    If pFlgName = "" Then
        GetOptFlagName = ""
    Else
        With pSrc
            '/* プレーンのPMDをオブジェクトで渡すとCustomPMD扱いになりPMDが可変PMDになってしまう問題を回避 */
            GetOptFlagName = .GetSharedFlagPlane(pFlgName).SetPMD(.CurrentPMD.Name).Name
        End With
    End If
    Exit Function
ErrHandler:
    Call TheError.Raise(Err.Number, FUNCTION_NAME, Err.Description)
End Function

'#FlagExpansion
Private Function GetOptFlagBit(ByRef pSrc As CImgPlane, ByVal pFlgName As String) As Long
'内容:
'   フラグ名が割り当てられているフラグビット番号を返す
'
'[pFlgName] In  フラグ名
'
'戻り値:
'   フラグビット番号
'
'概要:
'   空白文字列が入力された場合は0を返す
'   CImgPlaneクラスのGetOptFlagBitメソッドに引数を追加しただけ
'
'/* 2012/3/1 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetOptFlagBit"
    On Error GoTo ErrHandler
    If pFlgName = "" Then
        GetOptFlagBit = 0
    Else
        GetOptFlagBit = pSrc.GetSharedFlagPlane(pFlgName).FlagBit(pFlgName)
        If GetOptFlagBit = 0 Then GoTo errFlagName
    End If
    Exit Function
ErrHandler:
    Call TheError.Raise(Err.Number, FUNCTION_NAME, Err.Description)
errFlagName:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgNotExistSharedFlag(pFlgName))
End Function

'/* 以下CImgPlaneに追加したユーティリティメソッドと全く同じ */
Private Function GetArrayPlaneNames(ByRef pSrc() As CImgPlane, Optional ByVal pCheckReadOnly As Boolean = False) As String()
'内容:
'   プレーンオブジェクト配列をプレーン名の文字列配列に変換する。
'
'[pSrc()]           In  プレーンオブジェクト配列
'[pCheckReadOnly]   In  プレーンの書き込み禁止プロパティチェックフラグ
'
'戻り値:
'   プレーン名の文字列配列
'
'概要:
'   ①配列の最小インデックス番号は問わない
'   ②配列要素内にNothingがあった場合は空白にする
'   ③書込み禁止チェックがONの時にプレーンが書込み禁止だったらエラーにする
'
'/* 2012/2/27 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetArrayPlaneNames"
    On Error GoTo errArrayMismatch
    Dim pNames() As String
    ReDim pNames(UBound(pSrc) - LBound(pSrc))
    Dim pIndex As Long
    Dim tmpName As String
    For pIndex = 0 To UBound(pNames)
        '/* 配列要素のNothingは無視する */
        If Not pSrc(pIndex + LBound(pSrc)) Is Nothing Then
            '/* 書き込み禁止プロパティチェック */
            If pCheckReadOnly Then
                If pSrc(pIndex + LBound(pSrc)).ReadOnly Then
                    tmpName = pSrc(pIndex + LBound(pSrc)).Name
                    GoTo errReadOnly
                End If
            End If
            pNames(pIndex) = pSrc(pIndex + LBound(pSrc)).Name
        End If
    Next pIndex
    GetArrayPlaneNames = pNames
    Exit Function
errArrayMismatch:
    Call TheError.Raise(Err.Number, FUNCTION_NAME, ErrMsgInvalidArrayType)
errReadOnly:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx(tmpName))
End Function

Private Function GetIdpColorType(ByRef pRefMap As CImgPlaneMap, ByRef pColor As Variant) As IdpColorType
'内容:
'   カラー情報（文字列、カラーオブジェクト）をIdpColorTypeに変換する。
'
'[pRefMap]  In  演算対象のプレーンが持つマップオブジェクト
'[pColor」  In  演算対象プレーンに指定するカラー
'
'戻り値:
'   idpColorType列挙型
'
'概要:
'   ①idpColorTypeの入力はエラーにする
'   ②単独色指定はpRefMapに存在しない色だとエラーになる
'
'/* 2012/2/28 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetIdpColorType"
    On Error GoTo ErrHandler
    Dim tmpColor As Variant
    '/* 色指定がCImgColor型だった場合 */
    If TypeName(pColor) = "CImgColor" Then
        tmpColor = pColor.Name
        GetIdpColorType = pRefMap.color(pColor.Name).IdpColor
    '/* 色指定が文字列だった場合 */
    ElseIf TypeName(pColor) = "String" Then
        '/* カラーオールの場合 */
        If pColor = EEE_COLOR_ALL Then
            GetIdpColorType = idpColorAll
        '/* カラーフラットの場合 */
        ElseIf pColor = EEE_COLOR_FLAT Then
            GetIdpColorType = idpColorFlat
        '/* 単独色指定の場合 */
        Else
            tmpColor = pColor
            GetIdpColorType = pRefMap.color(CStr(pColor)).IdpColor
        End If
    '/* それ以外の変数型の入力は許可しない */
    Else
        GoTo errTypeMismatch
    End If
    Exit Function
ErrHandler:
    If Err.Number = 9999 Then
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgNotExistColor(pRefMap.MapName, CStr(tmpColor)))
    Else
        Call TheError.Raise(Err.Number, FUNCTION_NAME, Err.Description)
    End If
errTypeMismatch:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pColor"))
End Function

Private Function GetSrcIdpColorType(ByRef pRefMap As CImgPlaneMap, ByRef pColorS As Variant, ByRef pColorD As Variant) As IdpColorType
'内容:
'   GetIdpColorType関数の拡張。
'   pColorSにSAME_COLORが入力されるとpColorDに置き換えて変換する。
'
'[pRefMap]  In  演算対象のプレーンが持つマップオブジェクト
'[pColorS]  In  演算対象プレーンに指定するカラー
'[pColorD]  In  置き換え用のカラー
'
'戻り値:
'   idpColorType列挙型
'
'概要:
'   pColorS及びpColorDの型判定、単独色指定のカラーマップへの存在の有無などはGetIdpColorType関数に任せる
'
'/* 2012/2/28 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetSrcIdpColorType"
    On Error GoTo ErrHandler
    If TypeName(pColorS) = "String" Then
        '/* pColorSの色指定がSAME_COLORならpColorDに置き換える */
        If pColorS = SAME_COLOR Then
            GetSrcIdpColorType = GetIdpColorType(pRefMap, pColorD)
        Else
            GetSrcIdpColorType = GetIdpColorType(pRefMap, pColorS)
        End If
    Else
        GetSrcIdpColorType = GetIdpColorType(pRefMap, pColorS)
    End If
    Exit Function
ErrHandler:
    Call TheError.Raise(Err.Number, FUNCTION_NAME, Err.Description)
End Function

'/* 以下エラーメッセージ集 */
Private Function ErrMsgReadOnlyForEx(Optional ByVal pName As String = "Destination") As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgReadOnlyForEx = "[" & pName & "] is Read Only."
End Function

Private Function ErrMsgTypeMismatch(ByVal pType As String) As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgTypeMismatch = "Type of " & pType & " is mismatch."
End Function

Private Function ErrMsgInvalidArrayType() As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgInvalidArrayType = "Invalid array type."
End Function

Private Function ErrMsgNotExistColor(ByVal pMapName As String, ByVal pColorName As String) As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgNotExistColor = "The color [" & pColorName & "] is not included in plane map [" & pMapName & "]."
End Function

Private Function ErrMsgNotExistSharedFlag(ByVal pFlgName As String) As String
'/* 2012/3/1確認済み 0145206097 */
    ErrMsgNotExistSharedFlag = "Flag [" & pFlgName & "] does not exist."
End Function
#End If
'/* ======= 2012/01/18 Add 0145206097 ここまで ============== */
