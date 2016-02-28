VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDeviceConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   DeviceConfigurations�V�[�g�ɐݒ肵������ێ�����N���X
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfigurations�V�[�g�̏����擾����@�\����������
'       2010/11/30  ConfigInfo�擾���ɁAClone�𗘗p����悤�ɕύX����
'
'�쐬��:
'   0145184346
'

Option Explicit

Private Const ERR_NUMBER As Long = 9999                     ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CDeviceConfiguration" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty         ' Default�l"Empty"��ێ�����

Private Const STATIC_LSB As String = "Static"   ' LSB�I��Mode���̂�ێ�����
Private Const DYNAMIC_LSB As String = "Dynamic" ' LSB�I��Mode���̂�ێ�����
Private Const ACCTIME_UNIT_V As String = "V"    ' AccTime�̒P�ʖ��̂�ێ�����
Private Const ACCTIME_UNIT_H As String = "H"    ' AccTime�̒P�ʖ��̂�ێ�����

Private Const PARAM_INSTANCE_NAME As String = "Instance Name"           ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_CONFIG_NAME As String = "Config Name"               ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_FRAME_RATE As String = "Frame Rate"                 ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_XHSPV_NAME As String = "XHS/V"                      ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_ACCTIME_NAME As String = "Acc Time"                 ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_GAIN_NAME As String = "Gain"                        ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_MODE_LSB As String = "Mode@LSB Config"              ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_BASIC_LSB As String = "Basic Lsb@LSB Config"        ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_DYNAMIC_LSB As String = "Procedure Name@LSB Config" ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_COMMENT_DATA As String = "Comments"                 ' �p�����[�^�Q�ƃL�[��ێ�����

Private mInstanceName As String       ' Instance���̂�ێ�����
Private mConfigName As String         ' Config���̂�ێ�����
Private mFrameRate As CParamFrameRate ' FrameRate�f�[�^��ێ�����
Private mAccTime As CParamAccTime     ' AccTime�f�[�^��ێ�����
Private mGain As CParamGain           ' Gain�f�[�^��ێ�����
Private mSelectLSB As String          ' LSB�l�̑I������ێ�����
Private mLSB As CParamLsb             ' FixedLSB�l��ێ�����
Private mComment As String            ' Comment�f�[�^��ێ�����

Public Sub Initialize(ByRef reader As IParameterReader)
'���e:
'   DeviceConfigurations�V�[�g�ǂݍ��݁�����������
'
'�p�����[�^:
'   [reader]  In  DeviceConfigurations�V�[�g�̏���ێ�����
'
'�߂�l:
'
'���ӎ���:
'

    '#####  DeviceConfigurations�V�[�g�̏����i�[  #####
    '#####  ������ŕێ�����p�����[�^���i�[        #####
    With reader
        mInstanceName = .ReadAsString(PARAM_INSTANCE_NAME)
        mConfigName = .ReadAsString(PARAM_CONFIG_NAME)
        Call paramIsEmpty(reader, PARAM_MODE_LSB, True)
        mSelectLSB = .ReadAsString(PARAM_MODE_LSB)
        mComment = .ReadAsString(PARAM_COMMENT_DATA)
    End With

    '#####  LSB�I�����m�F����  #####
    If (mSelectLSB <> STATIC_LSB) And (mSelectLSB <> DYNAMIC_LSB) Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Initialize", "UnSupported Select LSB.[" & mSelectLSB & "]")
    End If

    '#####  FrameRate�p�����[�^���i�[  #####
    If paramIsEmpty(reader, PARAM_FRAME_RATE) = False Then
        Call paramIsVariable(reader, PARAM_FRAME_RATE, True)
        Set mFrameRate = New CParamFrameRate
        mFrameRate.ReadParameter reader, PARAM_FRAME_RATE
    End If
    
    '#####  Gain�p�����[�^���i�[  #####
    If paramIsEmpty(reader, PARAM_GAIN_NAME) = False Then
        Call paramIsVariable(reader, PARAM_GAIN_NAME, True)
        Set mGain = New CParamGain
        mGain.ReadParameter reader, PARAM_GAIN_NAME
    End If

    '#####  AccumulateTime�p�����[�^���i�[  #####
    If (paramIsEmpty(reader, PARAM_XHSPV_NAME) = False) Or (paramIsEmpty(reader, PARAM_ACCTIME_NAME) = False) Then
        Call paramIsVariable(reader, PARAM_XHSPV_NAME, True)
        Set mAccTime = New CParamAccTime
        mAccTime.ReadParameter reader, PARAM_XHSPV_NAME, PARAM_ACCTIME_NAME
    End If

    '#####  �p�����[�^�̑g�ݍ��킹���m�F���āALSB�p�����[�^���i�[����  #####
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
'���e:
'   Instance���̂�߂�
'
'�p�����[�^:
'
'�߂�l:
'   Instance����
'
'���ӎ���:
'
    InstanceName = mInstanceName
End Property

Public Property Get ConfigName() As String
'���e:
'   ConfigName���̂�߂�
'
'�p�����[�^:
'
'�߂�l:
'   ConfigName����
'
'���ӎ���:
'
    ConfigName = mConfigName
End Property

Public Property Get SelectLSB() As String
'���e:
'   SelectLSB���̂�߂�
'
'�p�����[�^:
'
'�߂�l:
'   SelectLSB����
'
'���ӎ���:
'
    SelectLSB = mSelectLSB
End Property

Public Property Get Comment() As String
'���e:
'   Comment�f�[�^��߂�
'
'�p�����[�^:
'
'�߂�l:
'   Comment�f�[�^
'
'���ӎ���:
'
    Comment = mComment
End Property

Public Property Get ConfigInfo() As CDeviceConfigInfo
'���e:
'   DeviceConfigInfo��߂�
'
'�p�����[�^:
'
'�߂�l:
'   DeviceConfigInfo
'
'���ӎ���:
'

    '#####  �e��p�����[�^��ConfigInfo�Ɋi�[����  #####
    Dim clsDevCfgInf As CDeviceConfigInfo ' ConfigInfo��ێ�����
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

    '#####  �I��  #####
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