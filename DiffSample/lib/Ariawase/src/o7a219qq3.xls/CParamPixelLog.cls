VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamPixelLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ピクセルログを結果セットを保持するクラス
'
'目的:
'   各ピクセルのピクセルログを保持する
'
'作成者:
'   0145184306
'
Option Explicit

Implements IProcParamReplaceable

Private Const ERR_NUMBER As Long = 9999                   ' Error番号を保持する
Private Const CLASS_NAME As String = "CParamPixelLog"     ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty       ' Default値"Empty"を保持する

Private Const PARAMETER_REPLACER_COLOR As String = "Color"
Private Const PARAMETER_REPLACER_VALUE As String = "Value"
Private Const PARAMETER_REPLACER_X As String = "x"
Private Const PARAMETER_REPLACER_Y As String = "y"

Private m_Color As CImgColor
Private m_Level As Double
Private m_XAddress As Long
Private m_YAddress As Long

Public Sub Initialize(ByVal pXAdd As Long, ByVal pYAdd As Long, ByVal pLevel As Double, ByRef pColor As CImgColor)
'内容:
'   初期化
'
'引数:
'[pXAdd]     IN     Long型:  Xアドレス
'[pYAdd]     IN     Long型:  Yアドレス
'[pLevel]    IN     Double型:  レベル
'[pColor]    IN/OUT CImgColor型:  色情報
'
'備考:
'
'
    Set m_Color = pColor
    m_Level = pLevel
    m_XAddress = pXAdd
    m_YAddress = pYAdd
End Sub

Public Property Get color() As CImgColor
'内容:
'   保持した色情報を返す
'
'戻り値:
'   保持した色情報
'
'備考:
'
    Set color = m_Color
End Property

Public Property Get Level() As Double
'内容:
'   保持したレベル情報を返す
'
'戻り値:
'   保持したレベル情報
'
'備考:
'
    Level = m_Level
End Property

Public Property Get x() As Long
'内容:
'   保持したXアドレスを返す
'
'戻り値:
'   保持したXアドレス
'
'備考:
'
    x = m_XAddress
End Property

Public Property Get y() As Long
'内容:
'   保持したYアドレスを返す
'
'戻り値:
'   保持したYアドレス
'
'備考:
'
    y = m_YAddress
End Property

Private Sub Class_Initialize()
    Set m_Color = Nothing
    m_Level = 0
    m_XAddress = 0
    m_YAddress = 0
End Sub

Private Sub Class_Terminate()
    Set m_Color = Nothing
    m_Level = 0
    m_XAddress = 0
    m_YAddress = 0
End Sub

Private Sub IProcParamReplaceable_GetContents(ByVal pRoot As String, ByRef pRes As Collection)
'内容:
'   パラメータの設定内容を取得する
'
'パラメータ:
'   [pRoot]     IN  String型:  このパラメータクラスの階層情報
'   [pRes]      IN/OUT  Collection型:  パラメータの設定内容を格納するコレクション
'
'注意事項:
'
'

    With pRes
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COLOR, m_Color.Name, "String", False, PARAMETER_REPLACER_COLOR)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_VALUE, CStr(m_Level), "Double", True, PARAMETER_REPLACER_VALUE)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_X, CStr(m_XAddress), "Long", True, PARAMETER_REPLACER_X)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_Y, CStr(m_YAddress), "Long", True, PARAMETER_REPLACER_Y)
    End With
End Sub

Private Sub IProcParamReplaceable_ReplaceContent(ByVal pRoot As String, ByRef pParamContents As CImpLocalParamsContent)
'内容:
'   パラメータの内容を書き換える
'
'パラメータ:
'   [pRoot]             IN  String型:  このパラメータクラスの階層情報
'   [pParamContents]    IN  CImpLocalParamsContent型:  書き換えを行うパラメータの内容
'
'注意事項:
'
'

    With pParamContents
        Dim subParamLabel As String
        subParamLabel = XLibImpExperimentUtility.GetSubParamLabel(.Path, pRoot)
        Select Case subParamLabel
            Case PARAMETER_REPLACER_VALUE
                m_Level = CDbl(.Value)
            Case PARAMETER_REPLACER_X
                m_XAddress = CLng(.Value)
            Case PARAMETER_REPLACER_Y
                m_YAddress = CLng(.Value)
        End Select
    End With
End Sub

Private Sub AddParamContents(ByRef pRes As Collection, ByVal pPath As String, ByVal pValue As String, ByVal pTypeName As String, ByVal pEditable As Boolean, ByVal pLabel As String)
'内容:
'   パラメータの情報をコレクションに追加する
'
'パラメータ:
'   [pRes]          IN/OUT  Collection型:  パラメータの設定内容を追加するコレクション
'   [pPath]         IN  String型:  パラメータの階層情報
'   [pValue]        IN  String型:  パラメータの値
'   [pTypeName]     IN  String型:  パラメータのデータ型
'   [pEditable]     IN  Boolean型:  パラメータが編集可能かどうか
'   [pLabel]        IN  String型:  パラメータ名
'
'注意事項:
'
'

    Dim myContents As New CImpLocalParamsContent
    Call myContents.Initialize(pPath, pValue, pTypeName, pEditable, pLabel)
    Call pRes.Add(myContents)
End Sub
