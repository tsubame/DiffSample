VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDeviceConfigInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   DeviceConfigurationsシートに設定した情報を保持するクラス
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfiguration情報を保持する機能を実装した
'       2011/11/01  Clone,Dumpを実装した
'
'作成者:
'   0145184346
'

Option Explicit

Implements IProcParameter
Implements IProcParamReplaceable

Private Const ERR_NUMBER As Long = 9999                  ' Error番号を保持する
Private Const CLASS_NAME As String = "CDeviceConfigInfo" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default値"Empty"を保持する

Private Const PARAMETER_REPLACER_INSTANCENAME As String = "InstanceName"
Private Const PARAMETER_REPLACER_FRAMERATE As String = "FrameRate"
Private Const PARAMETER_REPLACER_ACCTIME As String = "AccTime"
Private Const PARAMETER_REPLACER_GAIN As String = "Gain"
Private Const PARAMETER_REPLACER_LSB As String = "LSB"

Private mInstanceName As String        ' Instance名称を保持する
Private mFrameRate As CParamFrameRate  ' FrameRateデータを保持する
Private mAccTime As CParamAccTime      ' AccTime_Vデータを保持する
Private mGain As CParamGain            ' Gainデシベルデータを保持する
Private mLSB As CParamLsb              ' LSB値を保持する
Private mConfigName As String         ' Config名称を保持する
Private mComment As String            ' Commentデータを保持す


Public Property Let InstanceName(ByVal strInstanceName As String)
'内容:
'   Instance名称を設定する
'
'パラメータ:
'   [strInstanceName]  In  Instance名称を保持する
'
'戻り値:
'
'注意事項:
'   書き込み処理は、1回のみ有効(Write At Once)
'

    Static bolWriteFlag As Boolean ' 書き込み実施確認用フラグ

    '#####  1回のみ書き込み可能  #####
    If bolWriteFlag = False Then
        mInstanceName = strInstanceName
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[InstanceName]")
    End If

End Property

Public Property Get InstanceName() As String
'内容:
'   Instance名称を戻す
'
'パラメータ:
'
'戻り値:
'   Instance名称を戻す
'
'注意事項:
'
    InstanceName = mInstanceName
End Property

Public Property Set FrameRate(ByRef cFrameRate As CParamFrameRate)
'内容:
'   FrameRate情報を設定する
'
'パラメータ:
'   [cFrameRate]  In  FrameRate情報を保持する
'
'戻り値:
'
'注意事項:
'   書き込み処理は、1回のみ有効(Write At Once)
'

    Static bolWriteFlag As Boolean ' 書き込み実施確認用フラグ
    
    '#####  1回のみ書き込み可能  #####
    If bolWriteFlag = False Then
        Set mFrameRate = cFrameRate
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[FrameRate]")
    End If

End Property

Public Property Get FrameRate() As IDeviceParameter
'内容:
'   FrameRate情報を戻す
'
'パラメータ:
'
'戻り値:
'   FrameRate情報を戻す
'
'注意事項:
'
    Set FrameRate = mFrameRate
End Property

Public Property Set AccTime(ByRef cAccTime As CParamAccTime)
'内容:
'   AccTime情報を設定する
'
'パラメータ:
'   [cAccTime]  In  AccTime情報を保持する
'
'戻り値:
'
'注意事項:
'   書き込み処理は、1回のみ有効(Write At Once)
'

    Static bolWriteFlag As Boolean ' 書き込み実施確認用フラグ
    
    '#####  1回のみ書き込み可能  #####
    If bolWriteFlag = False Then
        Set mAccTime = cAccTime
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[AccTime_V]")
    End If

End Property

Public Property Get AccTime() As IDeviceParameter
'内容:
'   AccTime情報を戻す
'
'パラメータ:
'
'戻り値:
'   AccTime情報を戻す
'
'注意事項:
'
    Set AccTime = mAccTime
End Property

Public Property Set gain(ByRef cGain As CParamGain)
'内容:
'   Gain情報を設定する
'
'パラメータ:
'   [cGainDecibel]  In  Gain情報を保持する
'
'戻り値:
'
'注意事項:
'   書き込み処理は、1回のみ有効(Write At Once)
'

    Static bolWriteFlag As Boolean ' 書き込み実施確認用フラグ

    '#####  1回のみ書き込み可能  #####
    If bolWriteFlag = False Then
        Set mGain = cGain
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[Gain]")
    End If

End Property

Public Property Get gain() As IDeviceParameter
'内容:
'   Gain情報を戻す
'
'パラメータ:
'
'戻り値:
'   Gain情報を戻す
'
'注意事項:
'
    Set gain = mGain
End Property

Public Property Set Lsb(ByRef cLSB As CParamLsb)
'内容:
'   LSB情報を設定する
'
'パラメータ:
'   [cLSB]  In  LSB情報を保持する
'
'戻り値:
'
'注意事項:
'   書き込み処理は、1回のみ有効(Write At Once)
'

    Static bolWriteFlag As Boolean ' 書き込み実施確認用フラグ
    
    '#####  1回のみ書き込み可能  #####
    If bolWriteFlag = False Then
        Set mLSB = cLSB
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[LSB]")
    End If

End Property

Public Property Get Lsb() As IDeviceParameter
'内容:
'   LSB情報を戻す
'
'パラメータ:
'
'戻り値:
'   LSB情報を戻す
'
'注意事項:
'
    Set Lsb = mLSB
End Property

Private Sub errorRaise(ByRef strErrMsg As String)
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, strErrMsg)
End Sub

Private Sub Class_Initialize()
    mInstanceName = INITIAL_EMPTY_VALUE
    Set mFrameRate = Nothing
    Set mAccTime = Nothing
    Set mGain = Nothing
    Set mLSB = Nothing
End Sub

Private Sub Class_Terminate()
    mInstanceName = INITIAL_EMPTY_VALUE
    Set mFrameRate = Nothing
    Set mAccTime = Nothing
    Set mGain = Nothing
    Set mLSB = Nothing
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

    Dim site As Long

    Dim tmpAcc() As String
    Dim tmpLsb() As String
    With pRes
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_INSTANCENAME, mInstanceName, "String", False, PARAMETER_REPLACER_INSTANCENAME)
        If Not mFrameRate Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_FRAMERATE, composeString(mFrameRate.AsIDeviceParameter.AsString), "IDeviceParameter", False, PARAMETER_REPLACER_FRAMERATE)
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_FRAMERATE, "Nothing", "IDeviceParameter", False, PARAMETER_REPLACER_FRAMERATE)
        End If
        If Not mAccTime Is Nothing Then
            tmpAcc = mAccTime.AsIDeviceParameter.AsStringAccTimeV
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_ACCTIME, vbNullString, "IDeviceParameter", False, PARAMETER_REPLACER_ACCTIME)
            For site = 0 To UBound(tmpAcc)
                Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_ACCTIME & "\Site" & site, tmpAcc(site), "IDeviceParameter", False, "Site" & site)
            Next site
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_ACCTIME, "Nothing", "IDeviceParameter", False, PARAMETER_REPLACER_ACCTIME)
        End If
        If Not mLSB Is Nothing Then
            tmpLsb = mLSB.AsIDeviceParameter.AsString
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_LSB, vbNullString, "IDeviceParameter", False, PARAMETER_REPLACER_LSB)
            For site = 0 To UBound(tmpLsb)
                Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_LSB & "\Site" & site, tmpLsb(site), "IDeviceParameter", False, "Site" & site)
            Next site
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_LSB, "Nothing", "IDeviceParameter", False, PARAMETER_REPLACER_LSB)
        End If
        If Not mGain Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_GAIN, composeString(mGain.AsIDeviceParameter.AsString), "IDeviceParameter", False, PARAMETER_REPLACER_GAIN)
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_GAIN, "Nothing", "IDeviceParameter", False, PARAMETER_REPLACER_GAIN)
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

Public Function IProcParameter_Clone() As IProcParameter
'内容:
'   Cloneを作成する
'
'戻り値:
'   CDeviceConfigInfoを返す（クローン）
'
'備考:
'
    
    '##### pDeviceInfoのクローンを作る #####
    Dim objFrameRate As CParamFrameRate
    Dim objAccTime As CParamAccTime
    Dim objGain As CParamGain
    Dim objLSB As CParamLsb
    
    Dim objDevInfo As CDeviceConfigInfo
    If Not Me Is Nothing Then
        Set objDevInfo = New CDeviceConfigInfo
        
        'InstanceName
        objDevInfo.InstanceName = Me.InstanceName
    
        'ConfigName
        objDevInfo.ConfigName = Me.ConfigName
        
        'FrameRate
        Set objFrameRate = Me.FrameRate
        Set objDevInfo.FrameRate = objFrameRate.Clone
        
        'AccTime
        Set objAccTime = Me.AccTime
        Set objDevInfo.AccTime = objAccTime.Clone
        
        'Gain
        Set objGain = Me.gain
        Set objDevInfo.gain = objGain.Clone
        
        'LSB
        Set objLSB = Me.Lsb
        Set objDevInfo.Lsb = objLSB.Clone
        
        'Comment
        objDevInfo.AsIProcParameter.Comment = Me.AsIProcParameter.Comment
    End If
    
    Set IProcParameter_Clone = objDevInfo
    
End Function

Public Sub IProcParameter_Dump(ByRef pWriter As IProcParamWriter)
'内容:
'   DeviceConfig情報を出力する
'
'[pWriter]     IN/OUT IProcParamWriter型:   Writer
'
'備考:
'
'
    If Not Me Is Nothing Then
        pWriter.WriteAsString "", "DeviceConfigInfo {"
        With pWriter
            .WriteAsBlockStart "DeviceConfigInfo"
            .WriteAsString "Instance Name", Me.InstanceName
            .WriteAsString "Config Name", Me.ConfigName
            If Not Me.FrameRate Is Nothing Then
                .WriteAsString "Frame Rate", composeString(Me.FrameRate.AsString)
            Else
                .WriteAsString "Frame Rate", "Nothing"
            End If
            If Not Me.AccTime Is Nothing Then
                .WriteAsString "XHS/V", composeString(Me.AccTime.AsStringAccTimeH)
            Else
                .WriteAsString "XHS/V", "Nothing"
            End If
            If Not Me.gain Is Nothing Then
                .WriteAsString "Gain", composeString(Me.gain.AsString)
            Else
                .WriteAsString "Gain", "Nothing"
            End If
            If Not Me.AccTime Is Nothing Then
                ' Acc Time Block
                Call writeSiteInfo(pWriter, "Acc Time", Me.AccTime.AsStringAccTimeV)
            End If
            If Not Me.Lsb Is Nothing Then
                ' Basic Lsb Block
                Call writeSiteInfo(pWriter, "Basic Lsb", Me.Lsb.AsString)
            End If
            .WriteAsString "", "Comment: " & Me.AsIProcParameter.Comment
            .WriteAsBlockEnd "DeviceConfigInfo"
        End With
        pWriter.WriteAsString "", "}"
    Else
        pWriter.WriteAsBlockStart "DeviceConfigInfo"
        pWriter.WriteAsString "DeviceConfigInfo", "Nothing"
        pWriter.WriteAsBlockEnd "DeviceConfigInfo"
    End If
End Sub

Private Property Let IProcParameter_Name(ByVal pName As String)
End Property

Private Property Get IProcParameter_Name() As String
End Property

Private Property Let IProcParameter_Comment(ByVal pComment As String)
'内容:
'   Comment情報を設定する
'
'パラメータ:
'   [strComment]  In  Comment情報を保持する
'
'戻り値:
'
'注意事項:
'   書き込み処理は、1回のみ有効(Write At Once)
'

    Static bolWriteFlag As Boolean ' 書き込み実施確認用フラグ

    '#####  1回のみ書き込み可能  #####
    If bolWriteFlag = False Then
        mComment = pComment
    Else
        Call errorRaise("This Parameter is Write At Once.[Comment]")
    End If

End Property

Private Property Get IProcParameter_Comment() As String
'内容:
'   Comment情報を戻す
'
'パラメータ:
'
'戻り値:
'   Comment情報を戻す
'
'注意事項:
'
    IProcParameter_Comment = mComment
End Property

Public Function AsIProcParameter() As IProcParameter
    Set AsIProcParameter = Me
End Function

Private Function composeString(ByRef strList() As String) As String
'内容:
'   Dump用文字列を作成する
'
'[strList]     IN String型: Dump用文字列の配列
'
'戻り値:
'   Stringを返す
'
'備考:
'
    Dim dataIndex As Long
    For dataIndex = 0 To UBound(strList)
        If dataIndex = 0 Then
            composeString = strList(dataIndex)
            Exit Function
        End If
    Next dataIndex
End Function

Public Sub IProcParameter_Read(ByRef pReader As IProcParamReader)
End Sub

Private Sub writeSiteInfo(ByRef pWriter As IProcParamWriter, ByVal pTitle As String, ByRef pStrList() As String)
'内容:
'   Site情報を出力する
'
'[pWriter]     IN IProcParamWriter型: Writer
'[pTitle]      IN String型: ブロックタイトル名称
'[pStrList]    IN String型: Dump用文字列の配列
'
'戻り値:
'
'
'備考:
'
    Dim dataIndex As Long
    
    With pWriter
        .WriteAsString "", pTitle & " {"
        .WriteAsBlockStart "DeviceConfigInfo"
        For dataIndex = 0 To UBound(pStrList)
            .WriteAsString "Site" & dataIndex, pStrList(dataIndex)
        Next dataIndex
        .WriteAsBlockEnd "DeviceConfigInfo"
        .WriteAsString "", "}"
    End With
End Sub

Public Property Let ConfigName(ByVal strConfigName As String)
'内容:
'   LSB値の選択情報を設定する
'
'パラメータ:
'   [strConfigName]  In  Config名称を保持する
'
'戻り値:
'
'注意事項:
'   書き込み処理は、1回のみ有効(Write At Once)
'

    Static bolWriteFlag As Boolean ' 書き込み実施確認用フラグ

    '#####  1回のみ書き込み可能  #####
    If bolWriteFlag = False Then
        mConfigName = strConfigName
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[ConfigName]")
    End If

End Property

Public Property Get ConfigName() As String
'内容:
'   Config名称を戻す
'
'パラメータ:
'
'戻り値:
'   Config名称を戻す
'
'注意事項:
'
    ConfigName = mConfigName
End Property
