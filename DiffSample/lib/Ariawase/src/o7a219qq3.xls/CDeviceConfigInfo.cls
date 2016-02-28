VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDeviceConfigInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   DeviceConfigurations�V�[�g�ɐݒ肵������ێ�����N���X
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfiguration����ێ�����@�\����������
'       2011/11/01  Clone,Dump����������
'
'�쐬��:
'   0145184346
'

Option Explicit

Implements IProcParameter
Implements IProcParamReplaceable

Private Const ERR_NUMBER As Long = 9999                  ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CDeviceConfigInfo" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default�l"Empty"��ێ�����

Private Const PARAMETER_REPLACER_INSTANCENAME As String = "InstanceName"
Private Const PARAMETER_REPLACER_FRAMERATE As String = "FrameRate"
Private Const PARAMETER_REPLACER_ACCTIME As String = "AccTime"
Private Const PARAMETER_REPLACER_GAIN As String = "Gain"
Private Const PARAMETER_REPLACER_LSB As String = "LSB"

Private mInstanceName As String        ' Instance���̂�ێ�����
Private mFrameRate As CParamFrameRate  ' FrameRate�f�[�^��ێ�����
Private mAccTime As CParamAccTime      ' AccTime_V�f�[�^��ێ�����
Private mGain As CParamGain            ' Gain�f�V�x���f�[�^��ێ�����
Private mLSB As CParamLsb              ' LSB�l��ێ�����
Private mConfigName As String         ' Config���̂�ێ�����
Private mComment As String            ' Comment�f�[�^��ێ���


Public Property Let InstanceName(ByVal strInstanceName As String)
'���e:
'   Instance���̂�ݒ肷��
'
'�p�����[�^:
'   [strInstanceName]  In  Instance���̂�ێ�����
'
'�߂�l:
'
'���ӎ���:
'   �������ݏ����́A1��̂ݗL��(Write At Once)
'

    Static bolWriteFlag As Boolean ' �������ݎ��{�m�F�p�t���O

    '#####  1��̂ݏ������݉\  #####
    If bolWriteFlag = False Then
        mInstanceName = strInstanceName
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[InstanceName]")
    End If

End Property

Public Property Get InstanceName() As String
'���e:
'   Instance���̂�߂�
'
'�p�����[�^:
'
'�߂�l:
'   Instance���̂�߂�
'
'���ӎ���:
'
    InstanceName = mInstanceName
End Property

Public Property Set FrameRate(ByRef cFrameRate As CParamFrameRate)
'���e:
'   FrameRate����ݒ肷��
'
'�p�����[�^:
'   [cFrameRate]  In  FrameRate����ێ�����
'
'�߂�l:
'
'���ӎ���:
'   �������ݏ����́A1��̂ݗL��(Write At Once)
'

    Static bolWriteFlag As Boolean ' �������ݎ��{�m�F�p�t���O
    
    '#####  1��̂ݏ������݉\  #####
    If bolWriteFlag = False Then
        Set mFrameRate = cFrameRate
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[FrameRate]")
    End If

End Property

Public Property Get FrameRate() As IDeviceParameter
'���e:
'   FrameRate����߂�
'
'�p�����[�^:
'
'�߂�l:
'   FrameRate����߂�
'
'���ӎ���:
'
    Set FrameRate = mFrameRate
End Property

Public Property Set AccTime(ByRef cAccTime As CParamAccTime)
'���e:
'   AccTime����ݒ肷��
'
'�p�����[�^:
'   [cAccTime]  In  AccTime����ێ�����
'
'�߂�l:
'
'���ӎ���:
'   �������ݏ����́A1��̂ݗL��(Write At Once)
'

    Static bolWriteFlag As Boolean ' �������ݎ��{�m�F�p�t���O
    
    '#####  1��̂ݏ������݉\  #####
    If bolWriteFlag = False Then
        Set mAccTime = cAccTime
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[AccTime_V]")
    End If

End Property

Public Property Get AccTime() As IDeviceParameter
'���e:
'   AccTime����߂�
'
'�p�����[�^:
'
'�߂�l:
'   AccTime����߂�
'
'���ӎ���:
'
    Set AccTime = mAccTime
End Property

Public Property Set gain(ByRef cGain As CParamGain)
'���e:
'   Gain����ݒ肷��
'
'�p�����[�^:
'   [cGainDecibel]  In  Gain����ێ�����
'
'�߂�l:
'
'���ӎ���:
'   �������ݏ����́A1��̂ݗL��(Write At Once)
'

    Static bolWriteFlag As Boolean ' �������ݎ��{�m�F�p�t���O

    '#####  1��̂ݏ������݉\  #####
    If bolWriteFlag = False Then
        Set mGain = cGain
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[Gain]")
    End If

End Property

Public Property Get gain() As IDeviceParameter
'���e:
'   Gain����߂�
'
'�p�����[�^:
'
'�߂�l:
'   Gain����߂�
'
'���ӎ���:
'
    Set gain = mGain
End Property

Public Property Set Lsb(ByRef cLSB As CParamLsb)
'���e:
'   LSB����ݒ肷��
'
'�p�����[�^:
'   [cLSB]  In  LSB����ێ�����
'
'�߂�l:
'
'���ӎ���:
'   �������ݏ����́A1��̂ݗL��(Write At Once)
'

    Static bolWriteFlag As Boolean ' �������ݎ��{�m�F�p�t���O
    
    '#####  1��̂ݏ������݉\  #####
    If bolWriteFlag = False Then
        Set mLSB = cLSB
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[LSB]")
    End If

End Property

Public Property Get Lsb() As IDeviceParameter
'���e:
'   LSB����߂�
'
'�p�����[�^:
'
'�߂�l:
'   LSB����߂�
'
'���ӎ���:
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
'���e:
'   �p�����[�^�̐ݒ���e���擾����
'
'�p�����[�^:
'   [pRoot]     IN  String�^:  ���̃p�����[�^�N���X�̊K�w���
'   [pRes]      IN/OUT  Collection�^:  �p�����[�^�̐ݒ���e���i�[����R���N�V����
'
'���ӎ���:
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
'���e:
'   �p�����[�^�̓��e������������
'
'�p�����[�^:
'   [pRoot]             IN  String�^:  ���̃p�����[�^�N���X�̊K�w���
'   [pParamContents]    IN  CImpLocalParamsContent�^:  �����������s���p�����[�^�̓��e
'
'���ӎ���:
'
'

End Sub

Private Sub AddParamContents(ByRef pRes As Collection, ByVal pPath As String, ByVal pValue As String, ByVal pTypeName As String, ByVal pEditable As Boolean, ByVal pLabel As String)
'���e:
'   �p�����[�^�̏����R���N�V�����ɒǉ�����
'
'�p�����[�^:
'   [pRes]          IN/OUT  Collection�^:  �p�����[�^�̐ݒ���e��ǉ�����R���N�V����
'   [pPath]         IN  String�^:  �p�����[�^�̊K�w���
'   [pValue]        IN  String�^:  �p�����[�^�̒l
'   [pTypeName]     IN  String�^:  �p�����[�^�̃f�[�^�^
'   [pEditable]     IN  Boolean�^:  �p�����[�^���ҏW�\���ǂ���
'   [pLabel]        IN  String�^:  �p�����[�^��
'
'���ӎ���:
'
'

    Dim myContents As New CImpLocalParamsContent
    Call myContents.Initialize(pPath, pValue, pTypeName, pEditable, pLabel)
    Call pRes.Add(myContents)
End Sub

Public Function IProcParameter_Clone() As IProcParameter
'���e:
'   Clone���쐬����
'
'�߂�l:
'   CDeviceConfigInfo��Ԃ��i�N���[���j
'
'���l:
'
    
    '##### pDeviceInfo�̃N���[������� #####
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
'���e:
'   DeviceConfig�����o�͂���
'
'[pWriter]     IN/OUT IProcParamWriter�^:   Writer
'
'���l:
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
'���e:
'   Comment����ݒ肷��
'
'�p�����[�^:
'   [strComment]  In  Comment����ێ�����
'
'�߂�l:
'
'���ӎ���:
'   �������ݏ����́A1��̂ݗL��(Write At Once)
'

    Static bolWriteFlag As Boolean ' �������ݎ��{�m�F�p�t���O

    '#####  1��̂ݏ������݉\  #####
    If bolWriteFlag = False Then
        mComment = pComment
    Else
        Call errorRaise("This Parameter is Write At Once.[Comment]")
    End If

End Property

Private Property Get IProcParameter_Comment() As String
'���e:
'   Comment����߂�
'
'�p�����[�^:
'
'�߂�l:
'   Comment����߂�
'
'���ӎ���:
'
    IProcParameter_Comment = mComment
End Property

Public Function AsIProcParameter() As IProcParameter
    Set AsIProcParameter = Me
End Function

Private Function composeString(ByRef strList() As String) As String
'���e:
'   Dump�p��������쐬����
'
'[strList]     IN String�^: Dump�p������̔z��
'
'�߂�l:
'   String��Ԃ�
'
'���l:
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
'���e:
'   Site�����o�͂���
'
'[pWriter]     IN IProcParamWriter�^: Writer
'[pTitle]      IN String�^: �u���b�N�^�C�g������
'[pStrList]    IN String�^: Dump�p������̔z��
'
'�߂�l:
'
'
'���l:
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
'���e:
'   LSB�l�̑I������ݒ肷��
'
'�p�����[�^:
'   [strConfigName]  In  Config���̂�ێ�����
'
'�߂�l:
'
'���ӎ���:
'   �������ݏ����́A1��̂ݗL��(Write At Once)
'

    Static bolWriteFlag As Boolean ' �������ݎ��{�m�F�p�t���O

    '#####  1��̂ݏ������݉\  #####
    If bolWriteFlag = False Then
        mConfigName = strConfigName
        bolWriteFlag = True
    Else
        Call errorRaise("This Parameter is Write At Once.[ConfigName]")
    End If

End Property

Public Property Get ConfigName() As String
'���e:
'   Config���̂�߂�
'
'�p�����[�^:
'
'�߂�l:
'   Config���̂�߂�
'
'���ӎ���:
'
    ConfigName = mConfigName
End Property