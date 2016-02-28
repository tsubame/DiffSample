VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamSitePixelLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   サイト毎のピクセルログを管理するクラス
'
'目的:
'   サイト毎のピクセルログを管理する
'
'作成者:
'   0145184306
'
Option Explicit

Implements IProcParamReplaceable

Private Const ERR_NUMBER As Long = 9999                   ' Error番号を保持する
Private Const CLASS_NAME As String = "CParamSitePixelLog" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty       ' Default値"Empty"を保持する

Private Const PARAMETER_REPLACER_SITE As String = "Site"
Private Const PARAMETER_REPLACER_PIXELLOGS As String = "PixelLogs"

Private m_PixelLogs As Collection
Private m_Site As Long

Public Sub Add(ByRef pPixelLog As CParamPixelLog)
'内容:
'   CParamPixelLogを登録する
'
'引数:
'[pPixelLog]     IN/OUT CParamPixelLog型:  各アドレスのピクセルログ
'
'備考:
'
'
    Call m_PixelLogs.Add(pPixelLog)
End Sub

Public Sub Initialize(ByVal pSite As Long)
'内容:
'   初期化
'
'引数:
'[pSite]     IN Long型:  サイト番号
'
'備考:
'
'
    m_Site = pSite
End Sub

Public Property Get PixelLogs() As Collection
'内容:
'   ピクセルログを纏めて返す
'
'戻り値:
'   ピクセルログコレクション
'
'備考:
'
'
    Set PixelLogs = m_PixelLogs
End Property

Public Property Get Count() As Long
'内容:
'   ピクセルログの数を返す
'
'戻り値:
'   ピクセルログ数
'
'備考:
'
'
    Count = m_PixelLogs.Count
End Property

Public Property Get site() As Long
'内容:
'   サイト番号を返す
'
'戻り値:
'   サイト番号
'
'備考:
'
'
    site = m_Site
End Property
Private Sub Class_Initialize()
    Set m_PixelLogs = New Collection
    m_Site = 0
End Sub

Private Sub Class_Terminate()
    Set m_PixelLogs = Nothing
    m_Site = 0
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
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SITE, CStr(m_Site), "Long", False, PARAMETER_REPLACER_SITE)
        If Not m_PixelLogs Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELLOGS, vbNullString, "Collection", False, PARAMETER_REPLACER_PIXELLOGS)
            itemCounter = 0
            For Each bufSubParam In m_PixelLogs
                If Not bufSubParam Is Nothing Then
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELLOGS & "\Item(" & itemCounter & ")", vbNullString, "CParamPixelLog", False, "Item(" & itemCounter & ")")
                    Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_PIXELLOGS & "\Item(" & itemCounter & ")", pRes)
                Else
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELLOGS & "\Item(" & itemCounter & ")", "Nothing", "CParamPixelLog", False, "Item(" & itemCounter & ")")
                End If
                itemCounter = itemCounter + 1
            Next bufSubParam
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PIXELLOGS, "Nothing", "Collection", False, PARAMETER_REPLACER_PIXELLOGS)
        End If
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

    Dim targetSubParam As IProcParamReplaceable
    Dim targetIndex As Long
    With pParamContents
        Dim subParamLabel As String
        subParamLabel = XLibImpExperimentUtility.GetSubParamLabel(.Path, pRoot)
        Select Case subParamLabel
            Case PARAMETER_REPLACER_PIXELLOGS
                targetIndex = XLibImpExperimentUtility.GetSubParamIndex(.Path, pRoot)
                Set targetSubParam = m_PixelLogs.Item(targetIndex + 1)
                Call targetSubParam.ReplaceContent(pRoot & "\" & subParamLabel & "\Item(" & targetIndex & ")", pParamContents)
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
