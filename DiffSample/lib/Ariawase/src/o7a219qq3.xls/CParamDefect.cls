VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamDefect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ピクセルログのセットを扱うクラス
'
'目的:
'   ピクセルログを管理する
'
'作成者:
'   0145184306
'
Option Explicit

Implements IProcParameter
Implements IProcParamReplaceable

Private Const ERR_NUMBER As Long = 9999                  ' Error番号を保持する
Private Const CLASS_NAME As String = "CParamDefect"       ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default値"Empty"を保持する

Private Const PARAMETER_REPLACER_NAME As String = "Name"
Private Const PARAMETER_REPLACER_DEVICEINFO As String = "DeviceInfo"
Private Const PARAMETER_REPLACER_SITES As String = "Sites"
Private Const PARAMETER_REPLACER_COMMENT As String = "Comment"

Private m_Comment As String
Private m_Name As String
Private m_DeviceInfo As CDeviceConfigInfo
Private m_Sites As Collection

Public Function AsIProcParameter() As IProcParameter
    Set AsIProcParameter = Me
End Function

Private Function IProcParameter_Clone() As IProcParameter
'内容:
'   自身のクローンを返す
'
'戻り値:
'   IProcParameter型クラス
'
'備考:
'
'
    
    '##### 自身が持っているDeviceInfoのクローンを作る #####
    Dim objDevInfo As CDeviceConfigInfo
    If Not Me.DeviceConfig Is Nothing Then
        Set objDevInfo = Me.DeviceConfig.AsIProcParameter.Clone
    Else
        Set objDevInfo = Nothing
    End If
        
    '#####  自身のCloneを作成して戻す  #####
    Dim objClone As CParamDefect
    Set objClone = New CParamDefect

    objClone.AsIProcParameter.Name = m_Name
    Set objClone.DeviceConfig = objDevInfo
    objClone.AsIProcParameter.Comment = m_Comment
    
    Dim lSite As Long
    For lSite = 1 To m_Sites.Count
        If m_Sites.Item(lSite) Is Nothing Then
            Call objClone.Add(Nothing, lSite - 1)
        Else
            Call objClone.Add(m_Sites.Item(lSite), m_Sites.Item(lSite).site)
        End If
    Next lSite
    
    Set IProcParameter_Clone = objClone
    
    Set objClone = Nothing
    Set objDevInfo = Nothing

End Function

Private Sub IProcParameter_Dump(ByRef pWriter As IProcParamWriter)
'内容:
'   自身の情報をライターにダンプする
'
'引数:
'[pWriter]     IN/OUT IProcParamWriter型:  ライター
'
'備考:
'
'
    
    pWriter.WriteAsString "Parameter Type", m_Name
    If Not m_DeviceInfo Is Nothing Then
        m_DeviceInfo.AsIProcParameter.Dump pWriter
    End If
        
    Dim lSite As Long
    Dim vPixelLog As Variant
    
    For lSite = 0 To TheExec.sites.ExistingCount - 1
        If Me.site(lSite) Is Nothing Then
            pWriter.WriteAsString "Site = ", lSite
            pWriter.WriteAsString "PixelLogs", "Nothing"
        Else
            pWriter.WriteAsString "Site", Me.site(lSite).site
            For Each vPixelLog In Me.site(lSite).PixelLogs
                pWriter.WriteAsString m_Name, "x = " & vPixelLog.x
                pWriter.WriteAsString m_Name, "y = " & vPixelLog.y
                pWriter.WriteAsString m_Name, "Color = " & vPixelLog.color.Name
                pWriter.WriteAsString m_Name, "Value = " & vPixelLog.Level
            Next vPixelLog
        End If
    Next lSite
    
End Sub

Public Sub Add(ByRef pPixelSiteLog As CParamSitePixelLog, ByVal pSite As Long)
'内容:
'   自身の情報をライターにダンプする
'
'引数:
'[pPixelSiteLog]     IN/OUT CParamSitePixelLog型:  各サイトのピクセルログ
'
'備考:
'
'
    On Error GoTo err_handle
    Call m_Sites.Add(pPixelSiteLog, CStr(pSite))
    Exit Sub
err_handle:
    TheError.Raise ERR_NUMBER, CLASS_NAME, "This Site(" & pPixelSiteLog.site & ") has alreadey added."
End Sub

Public Function CountSites() As Long
'内容:
'   登録されたサイト数を返す
'
'備考:
'
'
    CountSites = m_Sites.Count
End Function
Public Function site(ByVal pSite As Long) As CParamSitePixelLog
'内容:
'   指定サイトのピクセルログを返す
'
'引数:
'[pSite]     IN Long型:  サイト番号
'
'戻り値:
'   CParamSitePixelLog型クラス
'
'備考:
'
'
    On Error GoTo err_handle
    Set site = m_Sites.Item(CStr(pSite))
    Exit Function
err_handle:
    TheError.Raise ERR_NUMBER, CLASS_NAME, "There is not CParamSitePixelLog as Site(" & pSite & ")"
End Function

Public Property Get DeviceConfig() As CDeviceConfigInfo
'内容:
'   保持したデバイス情報を返す
'
'戻り値:
'   保持したデバイス情報
'
'備考:
'
    Set DeviceConfig = m_DeviceInfo
End Property

Public Property Set DeviceConfig(ByRef pDeviceConfig As CDeviceConfigInfo)
'内容:
'   デバイス情報を保持する
'
'引数:
'[pDeviceInfo]   IN/OUT CDeviceConfigInfo型:     保持するデバイス情報を指定する
'
'備考:
'
    Set m_DeviceInfo = pDeviceConfig
End Property

Private Property Let IProcParameter_Name(ByVal pName As String)
'内容:
'   パラメータ名を保持する
'
'[pName]   IN String型:     保持するパラメータ名を指定する
'
'備考:
'

    m_Name = pName

End Property

Private Property Get IProcParameter_Name() As String
'内容:
'   保持したパラメータ名を返す
'
'戻り値:
'   保持したパラメータ名
'
'備考:
'

    IProcParameter_Name = m_Name

End Property

Private Property Let IProcParameter_Comment(ByVal pComment As String)
'内容:
'   コメントを保持する
'
'[pComment]   IN String型:     保持するコメントを指定する
'
'備考:
'

    m_Comment = pComment

End Property

Private Property Get IProcParameter_Comment() As String
'内容:
'   保持したコメントを返す
'
'戻り値:
'   保持したコメント
'
'備考:
'

    IProcParameter_Comment = m_Comment

End Property

Private Sub Class_Initialize()
    m_Comment = INITIAL_EMPTY_VALUE
    m_Name = INITIAL_EMPTY_VALUE
    Set m_DeviceInfo = Nothing
    Set m_Sites = New Collection
End Sub

Private Sub Class_Terminate()
    m_Comment = INITIAL_EMPTY_VALUE
    m_Name = INITIAL_EMPTY_VALUE
    Set m_DeviceInfo = Nothing
    Set m_Sites = Nothing
End Sub

Public Sub IProcParameter_Read(ByRef pReader As IProcParamReader)
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
        '   2012-01-27 (0145184931): Nameを編集不可に変更
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_NAME, m_Name, "String", False, PARAMETER_REPLACER_NAME)
        Set bufSubParam = m_DeviceInfo
        If Not bufSubParam Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, vbNullString, "CDeviceConfigInfo", False, PARAMETER_REPLACER_DEVICEINFO)
            Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, pRes)
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, "Nothing", "CDeviceConfigInfo", False, PARAMETER_REPLACER_DEVICEINFO)
        End If
        If Not m_Sites Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SITES, vbNullString, "Collection", False, PARAMETER_REPLACER_SITES)
            itemCounter = 0
            For Each bufSubParam In m_Sites
                If Not bufSubParam Is Nothing Then
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SITES & "\Item(" & itemCounter & ")", vbNullString, "CParamSitePixelLog", False, "Item(" & itemCounter & ")")
                    Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_SITES & "\Item(" & itemCounter & ")", pRes)
                Else
                    Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SITES & "\Item(" & itemCounter & ")", "Nothing", "CParamSitePixelLog", False, "Item(" & itemCounter & ")")
                End If
                itemCounter = itemCounter + 1
            Next bufSubParam
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_SITES, "Nothing", "Collection", False, PARAMETER_REPLACER_SITES)
        End If
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COMMENT, m_Comment, "String", True, PARAMETER_REPLACER_COMMENT)
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
'   2012-01-27 (0145184931): Nameを編集不可に変更
'            Case PARAMETER_REPLACER_NAME
'                m_Name = .value
            Case PARAMETER_REPLACER_SITES
                targetIndex = XLibImpExperimentUtility.GetSubParamIndex(.Path, pRoot)
                Set targetSubParam = m_Sites.Item(targetIndex + 1)
                Call targetSubParam.ReplaceContent(pRoot & "\" & subParamLabel & "\Item(" & targetIndex & ")", pParamContents)
            Case PARAMETER_REPLACER_COMMENT
                m_Comment = .Value
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
