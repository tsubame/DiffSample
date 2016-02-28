VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   色、またはidpColorTypeごとの管理クラス
'目的:
'   一つColorMapにおける、idpColorType定義１つ分の情報を管理します。
'   対応色の該当アドレスはCollectionで全て持ちます。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年1月31日

Option Explicit

Implements IProcParamReplaceable

Private Const PARAMETER_REPLACER_NAME As String = "Name"
Private Const PARAMETER_REPLACER_IDPCOLOR As String = "IdpColor"
Private Const PARAMETER_REPLACER_PIXELS As String = "Pixels"
Private Const PARAMETER_REPLACER_MAPNAME As String = "MapName"

Private m_ColorName As String
Private m_IdpColorType As IdpColorType
Private m_Pixels As Collection
Private m_ParentName As String

Public Sub Init(p_ParentName As String, p_ColorName As String, p_IdpColorType As IdpColorType, p_Pixels As Collection)
'内容:
'   このクラスの情報を設定します。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月1日
'パラメータ:
'   [p_ParentName]      In  1):この親クラスの名前
'   [p_ColorName]       In  2):このクラスの色の名前
'   [p_IdpColorType]    In  3):このクラスのidpColorType
'   [p_Pixels]          In  4):このクラスの色が該当する全アドレスのCollection
'戻り値:
'
'注意事項:
'
    m_ParentName = p_ParentName
    m_ColorName = p_ColorName
    m_IdpColorType = p_IdpColorType
    Set m_Pixels = p_Pixels
End Sub

Public Function Clone() As CImgColor
'内容:
'   このクラスのクローンを返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月3日
'パラメータ:
'   なし
'戻り値:
'   CImgColor
'
'注意事項:
'
'

    Dim p_Color As CImgColor
    Set p_Color = New CImgColor
    
    Dim p_Pixel As Variant
    Dim p_Pixels As Collection
    
    Set p_Pixels = New Collection
    
    For Each p_Pixel In m_Pixels
        Call p_Pixels.Add(p_Pixel.Clone)
    Next p_Pixel
    
    Call p_Color.Init(m_ParentName, m_ColorName, m_IdpColorType, p_Pixels)
    Set Clone = p_Color
    
End Function

Public Property Get Pixels() As Collection
    Set Pixels = m_Pixels
End Property

Public Property Get MapName() As String
    MapName = m_ParentName
End Property

Public Property Get Name() As String
    Name = m_ColorName
End Property

Public Property Get IdpColor() As IdpColorType
    IdpColor = m_IdpColorType
End Property

Private Sub Class_Terminate()
    Set m_Pixels = Nothing
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

    Dim bufSubParam As IProcParamReplaceable
    Dim itemCounter As Long
    With pRes
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_NAME, m_ColorName, "String", False, PARAMETER_REPLACER_NAME)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_IDPCOLOR, CStr(m_IdpColorType), "IdpColorType", False, PARAMETER_REPLACER_IDPCOLOR)
        If Not m_Pixels Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELS, vbNullString, "Collection", False, PARAMETER_REPLACER_PIXELS)
            itemCounter = 0
            For Each bufSubParam In m_Pixels
                If Not bufSubParam Is Nothing Then
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELS & "\Item(" & itemCounter & ")", vbNullString, "CImgPixel", False, "Item(" & itemCounter & ")")
                    Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_PIXELS & "\Item(" & itemCounter & ")", pRes)
                Else
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELS & "\Item(" & itemCounter & ")", "Nothing", "CImgPixel", False, "Item(" & itemCounter & ")")
                End If
                itemCounter = itemCounter + 1
            Next bufSubParam
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELS, "Nothing", "Collection", False, PARAMETER_REPLACER_PIXELS)
        End If
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_MAPNAME, m_ParentName, "String", False, PARAMETER_REPLACER_MAPNAME)
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
