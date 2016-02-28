VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_TestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' クラス名  ：  csLoopTestInfo.cls
' 概要      ： 各テスト項目のパラメータの格納、処理を行うクラス
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/10　新規作成        K.Sumiyashiki
'==========================================================================

'フローテーブルデータ
Public tname     As String
Public Parameter As String
Public Bin       As Integer
Public OpCode    As String
Public Label     As String

'テストインスタンスデータ
Public HighLimit As Double
Public LowLimit  As Double
Public Unit      As String
Public arg2      As Integer

'LOOP調整用データ
Public CapCategory1   As String
Public CapCategory2   As String
Public LoopJudgeLimit As Double

'光量調整用データ
Public OptTarget     As Double
Public OptIdentifier As String
Public OptJudgeLimit As Double

'Wait調整用データ
Public WaitAdjFlg    As String

Private m_xSite(nSite) As csPALS_TestInfoParams     '各サイトの特性値情報を格納する為のクラス定義


'********************************************************************************************
' 名前: init_csPALS_TestInfo
' 内容: 各変数の初期化
'　　　 m_xSiteはサイト数分のインスタンスを生成
' 引数: なし
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/18　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub init_csPALS_TestInfo()

On Error GoTo errPALSinit_csPALS_TestInfo

    tname = vbNullString
    Parameter = vbNullString
    Bin = -1
    OpCode = vbNullString
    Label = vbNullString
    
    HighLimit = -1
    LowLimit = -1
    Unit = vbNullString
    
    OptTarget = -1
    OptIdentifier = vbNullString
    OptJudgeLimit = -1

    CapCategory1 = vbNullString
    CapCategory2 = vbNullString
    LoopJudgeLimit = -1

    arg2 = -1

    Dim sitez As Long       'ループカウンタ
    'サイト数分のインスタンスを生成
    For sitez = 0 To nSite
        Set m_xSite(sitez) = New csPALS_TestInfoParams
    Next sitez

Exit Sub

errPALSinit_csPALS_TestInfo:
    Call sub_errPALS("csPALS_TestInfo not initialize at 'init_csPALS_TestInfo'", "0-6-01-0-22")

End Sub


'*********************************************************************************
' 関数名    ：  Class_Initialize
' 概要      ：  インスタンス生成時の初期化
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()
    '初期化関数の呼び出し
    Call init_csPALS_TestInfo
End Sub


'*********************************************************************************
' 関数名    ：  Class_Terminate
' 概要      ：  インスタンス開放時の処理
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Terminate()
    'インスタンスが解放された際に実行される関数
    Erase m_xSite
End Sub


'*********************************************************************************
' プロパティ名  ：  Site
' 概要          ：  各Siteの特性値情報を格納しているSiteオブジェクトを返すプロパティ
' 引数          ：  sitez :サイト番号
' 備考          ：  なし
' 更新履歴      ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Public Property Get site(ByVal sitez As Long) As csPALS_TestInfoParams
    
On Error GoTo errPALSSiteGet
    
    'ここにsiteプロパティの値を取得する処理を書く
    Set site = m_xSite(sitez)

Exit Property

errPALSSiteGet:
    Call sub_errPALS("Get site data error at 'csPALS_TestInfo -> Get Site'", "0-6-02-0-23")

End Property


Public Property Let site(ByVal sitez As Long, ByVal vNewValue As csPALS_TestInfoParams)
    
On Error GoTo errPALSSiteLet

    'ここにsiteプロパティに値をセットする処理を書く
    Set m_xSite(sitez) = vNewValue

Exit Property

errPALSSiteLet:
    Call sub_errPALS("Input site data error at 'csPALS_TestInfo -> Let Site'", "0-6-03-0-24")

End Property


Public Property Set site(ByVal sitez As Long, ByVal vNewValue As csPALS_TestInfoParams)
    
On Error GoTo errPALSSiteSet

    'ここにsiteプロパティにオブジェクトをセットする処理を書く
    Set m_xSite(sitez) = vNewValue

Exit Property

errPALSSiteSet:
    Call sub_errPALS("Set site data error at 'csPALS_TestInfo -> Set Site'", "0-6-04-0-25")

End Property

