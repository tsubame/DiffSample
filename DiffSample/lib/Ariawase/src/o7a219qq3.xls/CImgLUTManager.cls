VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgLUTManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   LUTを管理
'目的:
'   LUTの情報管理用
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日

Option Explicit

Private Const ERR_NUMBER = 9999                         ' Error番号
Private Const CLASS_NAME = "CImgLUTManager"             ' Class名称

Private m_LutInfos As Collection

Public Sub CreateIndexLUT( _
    ByVal pLUTName As String, ByVal pInVal1 As Long, ByVal pInVal2 As Long, _
    ByVal pOutVal1 As Long, ByVal pOutVal2 As Long, Optional ByVal pValidBit As Integer = 12 _
)
'内容:
'   LUTの作成
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   [pLUTName]  In  1):LUT定義名
'   [pInVal1]   In  2):InputValue1
'   [pInVal2]   In  3):InputValue2
'   [pOutVal1]  In  4):OutputValue1
'   [pOutVal2]  In  5):OutputValue2
'   [pValidBit] In  6):Bit
'戻り値:
'
'注意事項:
'
    Dim pLut As CImgLUTInfo
    
    'IG-XL LUTシートエラー仕様と同等のエラー処理
    If pOutVal1 < -32768 Or pOutVal1 > 32767 Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CreateIndexLUT", _
                            pLUTName & "'s OutputValue is Too big or Too Small")
    End If
    If pOutVal2 < -32768 Or pOutVal2 > 32767 Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CreateIndexLUT", _
                            pLUTName & "'s OutputValue is Too big or Too Small")
    End If
    
    If IsExist(pLUTName) Then
        Set pLut = m_LutInfos.Item(pLUTName)
        If pLut.Bit <> pValidBit Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CreateIndexLUT", _
                                pLUTName & "'s LUT bit is different.")
        End If
    Else
        Set pLut = New CImgLUTInfo
        Call pLut.Initialize(pLUTName, pValidBit)
        Call m_LutInfos.Add(pLut, pLUTName)
    End If
    Call pLut.Add(pLUTName, pInVal1, pInVal2, pOutVal1, pOutVal2)
    Set pLut = Nothing

    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.CreateIndexLUT(pLUTName, pInVal1, pInVal2, pOutVal1, pOutVal2, 0, pValidBit)
    
    Exit Sub
    
ERROR_IGXL:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CreateIndexLUT", "CreateIndexLUT")
    
End Sub

'IG-XLの仕様に合わせて、LUT定義の上書きを可能とするため廃止。
'ソースコード上に散在するLUT名は、ユーザーが管理しないといけない。
'Public Sub TerminateLUT()
''内容:
''   LUT定義の終端処理
''   LUTの定義が終了したことを明示して、定義終了後の同名の再定義は許可しません。
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
''2010/12/03 メソッド実行をCImgLUTManagerに移動
'
'    If m_tmpLutInfo.Count = 0 Then Exit Sub
'
'    Dim pLut As CImgLUTInfo
'    Set pLut = New CImgLUTInfo
'    Call pLut.Initialize(m_CurrentLutName, m_LatestBitNumber, m_tmpLutInfo)
'
'    On Error GoTo ALREADY_ADD
'    Call m_LutInfos.Add(pLut, m_CurrentLutName)
'
'    Set m_tmpLutInfo = Nothing
'    Set m_tmpLutInfo = New Collection
'
'    Set pLut = Nothing
'
'    Exit Sub
'ALREADY_ADD:
'    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".TerminateLUT", "")
'
'End Sub

Public Function IsExist(pLUTName As String) As Boolean
'内容:
'   対象のLUTの有無を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   [pLutName]  In  1):
'戻り値:
'   Boolean
'
'注意事項:
'
    On Error GoTo NOT_FOUND
    Call m_LutInfos.Item(pLUTName)
    IsExist = True
    Exit Function
NOT_FOUND:
    IsExist = False
End Function

Public Function LUT(pLUTName As String) As CImgLUTInfo
'内容:
'   対象のLUTのインスタンスを返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年12月2日
'パラメータ:
'   [pLUTName]  In  1):LUT名
'戻り値:
'   CImgLUTInfo
'
'注意事項:
'

    On Error GoTo NOT_FOUND
    Set LUT = m_LutInfos.Item(pLUTName)
    Exit Function
NOT_FOUND:
    Call TheError.Raise(9999, "", "This LUT Name is not found!" & vbCrLf & "LUT Name:" & pLUTName)
End Function

'Public Property Get LutCollection() As Collection
''内容:
''   このクラスが持つLutのCollectionを返します。
''作成者:
''  tomoyoshi.takase
''作成日: 2010年12月2日
''パラメータ:
''   なし
''戻り値:
''   Collection
''
''注意事項:
'
'    Set LutCollection = m_LutInfos
'
'End Property

Public Property Get Count() As Long
'内容:
'   このクラスが持つLutの数を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年12月2日
'パラメータ:
'   なし
'戻り値:
'   Long
'
'注意事項:

    Count = m_LutInfos.Count

End Property

Private Sub Class_Initialize()
'内容:
'   コンストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'
'
    Set m_LutInfos = New Collection

End Sub

Private Sub Class_Terminate()
'内容:
'   デストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'
    Set m_LutInfos = Nothing

End Sub
