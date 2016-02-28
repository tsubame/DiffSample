VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDeviceConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   DeviceConfigurationsシートに設定した情報を保持するクラス
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfigurationsシートの情報を取得する機能を実装した
'       2010/11/30  ConfigInfo取得時に、Cloneを利用するように変更した
'
'作成者:
'   0145184346
'

Option Explicit

Private Const ERR_NUMBER As Long = 9999                     ' Error番号を保持する
Private Const CLASS_NAME As String = "CDeviceConfiguration" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty         ' Default値"Empty"を保持する

Private Const STATIC_LSB As String = "Static"   ' LSB選択Mode名称を保持する
Private Const DYNAMIC_LSB As String = "Dynamic" ' LSB選択Mode名称を保持する
Private Const ACCTIME_UNIT_V As String = "V"    ' AccTimeの単位名称を保持する
Private Const ACCTIME_UNIT_H As String = "H"    ' AccTimeの単位名称を保持する

Private Const PARAM_INSTANCE_NAME As String = "Instance Name"           ' パラメータ参照キーを保持する
Private Const PARAM_CONFIG_NAME As String = "Config Name"               ' パラメータ参照キーを保持する
Private Const PARAM_FRAME_RATE As String = "Frame Rate"                 ' パラメータ参照キーを保持する
Private Const PARAM_XHSPV_NAME As String = "XHS/V"                      ' パラメータ参照キーを保持する
Private Const PARAM_ACCTIME_NAME As String = "Acc Time"                 ' パラメータ参照キーを保持する
Private Const PARAM_GAIN_NAME As String = "Gain"                        ' パラメータ参照キーを保持する
Private Const PARAM_MODE_LSB As String = "Mode@LSB Config"              ' パラメータ参照キーを保持する
Private Const PARAM_BASIC_LSB As String = "Basic Lsb@LSB Config"        ' パラメータ参照キーを保持する
Private Const PARAM_DYNAMIC_LSB As String = "Procedure Name@LSB Config" ' パラメータ参照キーを保持する
Private Const PARAM_COMMENT_DATA As String = "Comments"                 ' パラメータ参照キーを保持する

Private mInstanceName As String       ' Instance名称を保持する
Private mConfigName As String         ' Config名称を保持する
Private mFrameRate As CParamFrameRate ' FrameRateデータを保持する
Private mAccTime As CParamAccTime     ' AccTimeデータを保持する
Private mGain As CParamGain           ' Gainデータを保持する
Private mSelectLSB As String          ' LSB値の選択情報を保持する
Private mLSB As CParamLsb             ' FixedLSB値を保持する
Private mComment As String            ' Commentデータを保持する

Public Sub Initialize(ByRef reader As IParameterReader)
'内容:
'   DeviceConfigurationsシート読み込み＆初期化する
'
'パラメータ:
'   [reader]  In  DeviceConfigurationsシートの情報を保持する
'
'戻り値:
'
'注意事項:
'

    '#####  DeviceConfigurationsシートの情報を格納  #####
    '#####  文字列で保持するパラメータを格納        #####
    With reader
        mInstanceName = .ReadAsString(PARAM_INSTANCE_NAME)
        mConfigName = .ReadAsString(PARAM_CONFIG_NAME)
        Call paramIsEmpty(reader, PARAM_MODE_LSB, True)
        mSelectLSB = .ReadAsString(PARAM_MODE_LSB)
        mComment = .ReadAsString(PARAM_COMMENT_DATA)
    End With

    '#####  LSB選択を確認する  #####
    If (mSelectLSB <> STATIC_LSB) And (mSelectLSB <> DYNAMIC_LSB) Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Initialize", "UnSupported Select LSB.[" & mSelectLSB & "]")
    End If

    '#####  FrameRateパラメータを格納  #####
    If paramIsEmpty(reader, PARAM_FRAME_RATE) = False Then
        Call paramIsVariable(reader, PARAM_FRAME_RATE, True)
        Set mFrameRate = New CParamFrameRate
        mFrameRate.ReadParameter reader, PARAM_FRAME_RATE
    End If
    
    '#####  Gainパラメータを格納  #####
    If paramIsEmpty(reader, PARAM_GAIN_NAME) = False Then
        Call paramIsVariable(reader, PARAM_GAIN_NAME, True)
        Set mGain = New CParamGain
        mGain.ReadParameter reader, PARAM_GAIN_NAME
    End If

    '#####  AccumulateTimeパラメータを格納  #####
    If (paramIsEmpty(reader, PARAM_XHSPV_NAME) = False) Or (paramIsEmpty(reader, PARAM_ACCTIME_NAME) = False) Then
        Call paramIsVariable(reader, PARAM_XHSPV_NAME, True)
        Set mAccTime = New CParamAccTime
        mAccTime.ReadParameter reader, PARAM_XHSPV_NAME, PARAM_ACCTIME_NAME
    End If

    '#####  パラメータの組み合わせを確認して、LSBパラメータを格納する  #####
    Dim clsPrmLsb As CParamLsb
    Set clsPrmLsb = New CParamLsb

    Select Case mSelectLSB
        Case STATIC_LSB
            Call paramIsEmpty(reader, PARAM_GAIN_NAME, True)
            Call paramIsEmpty(reader, PARAM_BASIC_LSB, True)
            clsPrmLsb.ReadParameter reader, PARAM_BASIC_LSB
        Case DYNAMIC_LSB
            Call paramIsEmpty(reader, PARAM_DYNAMIC_LSB, True)
            If paramIsVariable(reader, PARAM_DYNAMIC_LSB) = False Then
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME, _
                                "UnSupported Number " & "[" & PARAM_DYNAMIC_LSB & " : " & reader.ReadAsString(PARAM_DYNAMIC_LSB) & "]")
            End If
            clsPrmLsb.ReadParameter reader, PARAM_DYNAMIC_LSB
    End Select

    Set mLSB = clsPrmLsb
    Set clsPrmLsb = Nothing

End Sub

Private Function paramIsEmpty(ByRef reader As IParameterReader, ByRef strParamName As String, Optional ByRef cErrorRaise As Boolean = False) As Boolean
    If reader.ReadAsString(strParamName) = "" Then
        paramIsEmpty = True
        If cErrorRaise = True Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".paramIsEmpty", "This Parameter is Empty " & "[" & strParamName & "]")
        End If
    Else
        paramIsEmpty = False
    End If
End Function

Private Function paramIsVariable(ByRef reader As IParameterReader, ByRef strParamName As String, Optional ByRef cErrorRaise As Boolean = False) As Boolean
    If IsAlphabet(reader.ReadAsString(strParamName)) = True Then
        paramIsVariable = True
        If cErrorRaise = True Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".paramIsVariable", "UnSupported Variable " & "[" & strParamName & " : " & reader.ReadAsString(strParamName) & "]")
        End If
    Else
        paramIsVariable = False
    End If
End Function

Public Property Get InstanceName() As String
'内容:
'   Instance名称を戻す
'
'パラメータ:
'
'戻り値:
'   Instance名称
'
'注意事項:
'
    InstanceName = mInstanceName
End Property

Public Property Get ConfigName() As String
'内容:
'   ConfigName名称を戻す
'
'パラメータ:
'
'戻り値:
'   ConfigName名称
'
'注意事項:
'
    ConfigName = mConfigName
End Property

Public Property Get SelectLSB() As String
'内容:
'   SelectLSB名称を戻す
'
'パラメータ:
'
'戻り値:
'   SelectLSB名称
'
'注意事項:
'
    SelectLSB = mSelectLSB
End Property

Public Property Get Comment() As String
'内容:
'   Commentデータを戻す
'
'パラメータ:
'
'戻り値:
'   Commentデータ
'
'注意事項:
'
    Comment = mComment
End Property

Public Property Get ConfigInfo() As CDeviceConfigInfo
'内容:
'   DeviceConfigInfoを戻す
'
'パラメータ:
'
'戻り値:
'   DeviceConfigInfo
'
'注意事項:
'

    '#####  各種パラメータをConfigInfoに格納する  #####
    Dim clsDevCfgInf As CDeviceConfigInfo ' ConfigInfoを保持する
    Set clsDevCfgInf = New CDeviceConfigInfo

    With clsDevCfgInf
        .InstanceName = mInstanceName
        If Not (mFrameRate Is Nothing) Then Set .FrameRate = mFrameRate.Clone
        If Not (mGain Is Nothing) Then Set .gain = mGain.Clone
        If Not (mAccTime Is Nothing) Then
            mAccTime.ArrangeParameter
            Set .AccTime = mAccTime.Clone
        End If
        If Not (mLSB Is Nothing) Then
            If mSelectLSB = STATIC_LSB Then mLSB.CalcLsb mGain.AsIDeviceParameter.AsRatio
            If mSelectLSB = DYNAMIC_LSB Then mLSB.CalcLsb
            Set .Lsb = mLSB.Clone
        End If
        .ConfigName = mConfigName
        .AsIProcParameter.Comment = mComment
    End With

    '#####  終了  #####
    Set ConfigInfo = clsDevCfgInf
    Set clsDevCfgInf = Nothing

End Property

Private Sub Class_Initialize()
    mInstanceName = INITIAL_EMPTY_VALUE
    mConfigName = INITIAL_EMPTY_VALUE
    Set mFrameRate = Nothing
    Set mAccTime = Nothing
    Set mGain = Nothing
    mSelectLSB = INITIAL_EMPTY_VALUE
    Set mLSB = Nothing
    mComment = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    mInstanceName = INITIAL_EMPTY_VALUE
    mConfigName = INITIAL_EMPTY_VALUE
    Set mFrameRate = Nothing
    Set mAccTime = Nothing
    Set mGain = Nothing
    mSelectLSB = INITIAL_EMPTY_VALUE
    Set mLSB = Nothing
    mComment = INITIAL_EMPTY_VALUE
End Sub
