VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   システムに関わる情報のセットを保持するクラス
'
'目的:
'   テスターID、WaferIDなどシステムに関わる情報のセットを保持する。
'
'作成者:
'   0145184306
'

Option Explicit

Private Const ERR_NUMBER = 9999         ' エラー番号を保持する
Private Const CLASS_NAME = "CSystemInfo" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default値"Empty"を保持する

Private m_JobName As String
Private m_LotNo As Long
Private m_MapName As String
Private m_NodeNo As Long
Private m_ProductName As String
Private m_SystemName As String
Private m_WaferNo As Long
Private m_ChipNo() As Long
Private m_XAddress() As Long
Private m_YAddress() As Long

Public Property Get chipNo() As Long()
'内容:
'   保持したパラメータを返す
'
'戻り値:
'   保持したパラメータ(Long()型)
'
'備考:
'
'
    chipNo = m_ChipNo
End Property

Public Property Let chipNo(pChipNo() As Long)
'内容:
'   パラメータを保持する
'
'[pChipNo()]    IN Long型:     パラメータ
'
'備考:
'
'
    m_ChipNo = pChipNo
End Property

Public Property Get JobName() As String
'内容:
'   保持したパラメータを返す
'
'戻り値:
'   保持したパラメータ(String型)
'
'備考:
'
'
    JobName = m_JobName
End Property

Public Property Let JobName(pJobName As String)
'内容:
'   パラメータを保持する
'
'[pJobName]    IN String型:     パラメータ
'
'備考:
'
'
    m_JobName = pJobName
End Property

Public Property Get LotNo() As Long
'内容:
'   保持したパラメータを返す
'
'戻り値:
'   保持したパラメータ(Long型)
'
'備考:
'
'
    LotNo = m_LotNo
End Property

Public Property Let LotNo(pLotNo As Long)
'内容:
'   パラメータを保持する
'
'[pLotNo]    IN Long型:     パラメータ
'
'備考:
'
'
    m_LotNo = pLotNo
End Property

Public Property Get NodeNo() As Long
'内容:
'   保持したパラメータを返す
'
'戻り値:
'   保持したパラメータ(Long型)
'
'備考:
'
'
    NodeNo = m_NodeNo
End Property

Public Property Let NodeNo(pNodeNo As Long)
'内容:
'   パラメータを保持する
'
'[pNodeNo]    IN Long型:     パラメータ
'
'備考:
'
'
    m_NodeNo = pNodeNo
End Property

Public Property Get ProductName() As String
'内容:
'   保持したパラメータを返す
'
'戻り値:
'   保持したパラメータ(String型)
'
'備考:
'
'
    ProductName = m_ProductName
End Property

Public Property Let ProductName(pProductName As String)
'内容:
'   パラメータを保持する
'
'[pProductName]    IN String型:     パラメータ
'
'備考:
'
'
    m_ProductName = pProductName
End Property

Public Property Get SystemName() As String
'内容:
'   保持したパラメータを返す
'
'戻り値:
'   保持したパラメータ(String型)
'
'備考:
'
'
    SystemName = m_SystemName
End Property

Public Property Let SystemName(pSystemName As String)
'内容:
'   パラメータを保持する
'
'[pSystemName]    IN String型:     パラメータ
'
'備考:
'
'
    m_SystemName = pSystemName
End Property

Public Property Get WaferMapName() As String
'内容:
'   保持したパラメータを返す
'
'戻り値:
'   保持したパラメータ(Stirng型)
'
'備考:
'
'
    WaferMapName = m_MapName
End Property

Public Property Let WaferMapName(pMapName As String)
'内容:
'   パラメータを保持する
'
'[pMapName]    IN String型:     パラメータ
'
'備考:
'
'
    m_MapName = pMapName
End Property

Public Property Get WaferNo() As Long
'内容:
'   保持したパラメータを返す
'
'戻り値:
'   保持したパラメータ(Long型)
'
'備考:
'
'
    WaferNo = m_WaferNo
End Property

Public Property Let WaferNo(pWaferNo As Long)
'内容:
'   パラメータを保持する
'
'[pWaferNo]    IN Long型:     パラメータ
'
'備考:
'
'
    m_WaferNo = pWaferNo
End Property

Public Property Get xAddress() As Long()
'内容:
'   保持したパラメータを返す
'
'戻り値:
'   保持したパラメータ(Long()型)
'
'備考:
'
'
    xAddress = m_XAddress
End Property

Public Property Let xAddress(pXAddress() As Long)
'内容:
'   パラメータを保持する
'
'[pXAddress()]    IN Long型:     パラメータ
'
'備考:
'
'
    m_XAddress = pXAddress
End Property

Public Property Get yAddress() As Long()
'内容:
'   保持したパラメータを返す
'
'戻り値:
'   保持したパラメータ(Long()型)
'
'備考:
'
'
    yAddress = m_YAddress
End Property

Public Property Let yAddress(pYAddress() As Long)
'内容:
'   パラメータを保持する
'
'[pYAddress()]    IN Long型:     パラメータ
'
'備考:
'
'
    m_YAddress = pYAddress
End Property

Public Function GetChipNoOfSite(pSite As Long) As Long
'内容:
'   指定サイトのXアドレスを返す
'
'引数:
'[pSite]       IN Long型:       サイト番号
'
'戻り値:
'   指定サイトのChipNo(Long型)
'
'備考:
'
'
    On Error GoTo ErrHundler
    GetChipNoOfSite = m_ChipNo(pSite)
    Exit Function
ErrHundler:
    GetChipNoOfSite = -1
End Function

Public Function GetXAddressOfSite(pSite As Long) As Long
'内容:
'   指定サイトのXアドレスを返す
'
'引数:
'[pSite]       IN Long型:       サイト番号
'
'戻り値:
'   指定サイトのXアドレス(Long型)
'
'備考:
'
'
    On Error GoTo ErrHundler
    GetXAddressOfSite = m_XAddress(pSite)
    Exit Function
ErrHundler:
    GetXAddressOfSite = -1
End Function

Public Function GetYAddressOfSite(pSite As Long) As Long
'内容:
'   指定サイトのYアドレスを返す
'
'引数:
'[pSite]       IN Long型:       サイト番号
'
'戻り値:
'   指定サイトのYアドレス(Long型)
'
'備考:
'
'
    On Error GoTo ErrHundler
    GetYAddressOfSite = m_YAddress(pSite)
    Exit Function
ErrHundler:
    GetYAddressOfSite = -1
End Function

Private Sub Class_Initialize()
'内容:
'   コンストラクタ
'
'備考:
'
'
    m_JobName = INITIAL_EMPTY_VALUE
    m_LotNo = 0
    m_MapName = INITIAL_EMPTY_VALUE
    m_NodeNo = 0
    m_ProductName = INITIAL_EMPTY_VALUE
    m_SystemName = INITIAL_EMPTY_VALUE
    m_WaferNo = 0
    Erase m_ChipNo(), m_XAddress(), m_YAddress()
End Sub

Private Sub Class_Terminate()
'内容:
'   デストラクタ
'
'備考:
'
'
    m_JobName = INITIAL_EMPTY_VALUE
    m_LotNo = 0
    m_MapName = INITIAL_EMPTY_VALUE
    m_NodeNo = 0
    m_ProductName = INITIAL_EMPTY_VALUE
    m_SystemName = INITIAL_EMPTY_VALUE
    m_WaferNo = 0
    Erase m_ChipNo(), m_XAddress(), m_YAddress()
End Sub
