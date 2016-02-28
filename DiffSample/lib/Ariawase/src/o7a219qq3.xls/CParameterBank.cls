VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParameterBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   パラメータ置場。
'
'目的:
'   パラメータが持っている名前付きで登録される。
'   異なるスコープごとに個別のインスタンスを生成する｡
'
'作成者:
'   0145184304
'

Option Explicit

Implements IParameterBank

Private Const ERR_NUMBER As Long = 9999                     ' Error番号を保持する
Private Const CLASS_NAME As String = "CParameterBank" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty        ' Default値"Empty"を保持する

'=======
' paramNameをKeyにしてIProcParameterをいれる
'=======
Private m_Keys As Collection          ' キーのCollection
Private m_FixedKeys As Collection          ' キーのCollection
Private m_Parameters As Collection          ' 要素のCollection
Private m_FixedParameters As Collection     ' Fixされた要素のCollection
Private m_OverWriteMode As Boolean          ' 上書モード

Private Sub IParameterBank_Clear()
'内容:
'   m_Parametersを全部クリア
'
'備考:
'
    Set m_Parameters = New Collection
    Set m_Keys = New Collection
    
End Sub

Private Function IParameterBank_Count() As Long
'内容:
'   m_ParametersのCount数をかえす
'
'戻り値:
'    m_ParametersのCount数
'
'備考:
'

    IParameterBank_Count = m_Parameters.Count + m_FixedParameters.Count

End Function

Private Sub IParameterBank_Delete(ByVal paramName As String)
'内容:
'   m_ParamSetから指定のparamNameを削除
'
'[paramName]   IN String型:     削除するparamNameを指定
'
'備考:
'
    
    If Not GetParam(paramName, m_Parameters) Is Nothing Then
        ' NormalParametersにあったら削除
        Call m_Parameters.Remove(paramName)
        Call m_Keys.Remove(paramName)
    ElseIf Not GetParam(paramName, m_FixedParameters) Is Nothing Then
        ' FixedParametersにあったら削除
        Call m_FixedParameters.Remove(paramName)
        Call m_FixedKeys.Remove(paramName)
    Else
        ' Normal/Fixedともになければエラー
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Delete", """" & paramName & """ doesn't exist.")
    End If
End Sub

Private Function IParameterBank_FixedKeys() As Collection
    Set IParameterBank_FixedKeys = m_FixedKeys
End Function

Private Function IParameterBank_Item(ByVal paramName As String) As IProcParameter
'内容:
'   m_Parametersより指定のparamNameに一致するIProcParameterを返す
'
'[paramName]   IN String型:     GetするparamNameを指定
'
'戻り値:
'    指定のparamNameに一致するIProcParameter
'
'備考:
'

    ' NormalParametersから取得
    If Me.AsIParameterBank.IsExist(paramName) Then
        Set IParameterBank_Item = GetParam(paramName, m_Parameters)
        If IParameterBank_Item Is Nothing Then
            ' NormalParametersになければFixedParametersから取得
            Set IParameterBank_Item = GetParam(paramName, m_FixedParameters)
        End If
    Else
        ' Normal/Fixedともになければエラー
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Item", """" & paramName & """ doesn't exist.")
    End If

End Function

Private Function IParameterBank_Keys() As Collection
    Set IParameterBank_Keys = m_Keys
End Function

Private Function IParameterBank_Move(ByVal paramName As String) As IProcParameter
'内容:
'   m_Parametersより指定のparamNameに一致するIProcParameterを返す
'   返したParameterはm_Parametersより削除する
'
'[paramName]   IN String型:     GetするparamNameを指定
'
'戻り値:
'    指定のparamNameに一致するIProcParameter
'
'備考:
'

    ' m_Parametersから検索して取得
    Set IParameterBank_Move = GetParam(paramName, m_Parameters)
    
    If IParameterBank_Move Is Nothing Then
        ' m_Parametersになかった場合、m_FixedParametersを検索して取得
        Set IParameterBank_Move = GetParam(paramName, m_FixedParameters)
    Else
        ' m_Parametersにあった場合、m_Parametersから削除
        Call m_Parameters.Remove(paramName)
        Call m_Keys.Remove(paramName)
        Exit Function
    End If

    'エラー
    If IParameterBank_Move Is Nothing Then
        ' m_FixedParametersになかった場合、エラー
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Item", """" & paramName & """ doesn't exist.")
    Else
        ' m_FixedParametersにあった場合、m_FixedParametersから削除
        Call m_FixedParameters.Remove(paramName)
        Call m_FixedKeys.Remove(paramName)
    End If
    
End Function

Private Sub IParameterBank_Add(ByRef paramObj As IProcParameter, ByVal paramName As String, Optional ByVal pFixedPlane As Boolean = False)
'内容:
'   m_Parametersに指定のparamNameをKeyにIProcParameterをAddする。
'
'[paramObj]   IN IProcParameter型:     SetするIPocParameterを指定
'[pFixedPlane] IN   Boolean型:      固定プレーンオプション
'
'備考:
'

    If IParameterBank_IsOverwriteMode = False Then
        ' 上書き禁止の時、すでに存在しているか確認
        If IParameterBank_IsExist(paramName) = True Then
            ' 上書き禁止ですでに存在していた場合はエラー
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Add", paramName & " already set")
            Exit Sub
        End If
    Else
        ' 上書き可能の場合は存在している場合は一度消す
        If IsExistParam(paramName, m_FixedParameters) = True Then
            Call m_FixedParameters.Remove(paramName)
            Call m_FixedKeys.Remove(paramName)
        ElseIf IsExistParam(paramName, m_Parameters) = True Then
            Call m_Parameters.Remove(paramName)
            Call m_Keys.Remove(paramName)
        End If
        '登録無ければ何もしない
    End If

    If pFixedPlane = True Then
        'FixedParametersに登録
        Call m_FixedParameters.Add(paramObj, paramName)     'Keyはparamの名前
        Call m_FixedKeys.Add(paramName, paramName)
    Else
        'NormalParametersに登録
        Call m_Parameters.Add(paramObj, paramName)
        Call m_Keys.Add(paramName, paramName)
    End If

End Sub
Private Function IParameterBank_List() As String
'内容:
'   登録されたパラメータのリストを取得
'
'備考:
'   カンマ区切りの文字列で出力。
'   登録名,プレーン名,…　の順
'
'出力サンプル
' [pParam1(Fixed),CParamPlane,pParam2,CParamPlane,]

    Dim tmp As Variant
    For Each tmp In m_FixedKeys
'        IParameterBank_List = IParameterBank_List & tmp.Name & "(Fixed)" & "," & TypeName(tmp) & ","
        IParameterBank_List = IParameterBank_List & tmp & "(Fixed)" & "," & TypeName(m_FixedParameters.Item(tmp)) & ","
    Next tmp
    
    For Each tmp In m_Keys
'        IParameterBank_List = IParameterBank_List & tmp.Name & "," & TypeName(tmp) & ","
        IParameterBank_List = IParameterBank_List & tmp & "," & TypeName(m_Parameters.Item(tmp)) & ","
    Next tmp
    
End Function
Private Property Get IParameterBank_IsOverwriteMode() As Boolean
    IParameterBank_IsOverwriteMode = m_OverWriteMode
End Property

Private Property Let IParameterBank_IsOverwriteMode(ByRef mode As Boolean)
    m_OverWriteMode = mode
End Property

Private Function IParameterBank_IsExist(ByVal paramName As String) As Boolean
'内容:
'   m_Parametersより指定のparamNameに一致するIProcParameterがあった場合Trueを返す
'
'[paramName]   IN String型:     ある/なしを確認するparamNameを指定
'
'戻り値:
'    指定のparamNameに一致するIProcParameterがあった場合 True
'    指定のparamNameに一致するIProcParameterがなかった場合 False
'
'備考:
'
    
    If IsExistParam(paramName, m_Parameters) Then
        IParameterBank_IsExist = True
        Exit Function
    ElseIf IsExistParam(paramName, m_FixedParameters) Then
        IParameterBank_IsExist = True
        Exit Function
    End If
    
    IParameterBank_IsExist = False

End Function

Private Function IsExistParam(ByVal paramName As String, ByRef pParam As Collection) As Boolean
'内容:
'   指定のコレクションより指定のparamNameに一致するKeyがあった場合Trueを返す
'
'[paramName]   IN String型:     ある/なしを確認するparamNameを指定
'[pParam]      IN Collection型: 確認対象のCollection
'
'戻り値:
'    指定のparamNameに一致するKeyがあった場合 True
'    指定のparamNameに一致するKeyがなかった場合 False
'
'備考:
'

    Dim p_Param As IProcParameter
    
    On Error GoTo NOT_FOUND_PARAM
    
    Set p_Param = pParam.Item(paramName)
    IsExistParam = True
    
    Set p_Param = Nothing
    Exit Function
    
NOT_FOUND_PARAM:
    
    IsExistParam = False

End Function

Private Function GetParam(ByVal pName As String, ByRef pParam As Collection) As IProcParameter
'内容:
'   指定のコレクションより指定のpNameに一致するKeyがあった場合、アイテム(IProcParameter)を返す
'
'[paramName]   IN String型:     ある/なしを確認するparamNameを指定
'[pParam]      IN Collection型: 確認対象のCollection
'
'戻り値:
'    指定のparamNameに一致するアイテム IProcParameter型
'
'備考:
'

On Error GoTo NOT_FOUND
    Set GetParam = pParam.Item(pName)
    GetParam.Name = pName
    Exit Function

NOT_FOUND:
    Set GetParam = Nothing
    
End Function

Private Sub Class_Initialize()
    Set m_Keys = New Collection
    Set m_FixedKeys = New Collection
    Set m_Parameters = New Collection
    Set m_FixedParameters = New Collection
    m_OverWriteMode = False
End Sub

Private Sub Class_Terminate()

    Set m_Parameters = Nothing
    Set m_Keys = Nothing
    Set m_FixedKeys = Nothing
    Set m_FixedParameters = Nothing
    m_OverWriteMode = False
    
End Sub

Public Function AsIParameterBank() As IParameterBank
    Set AsIParameterBank = Me
End Function

