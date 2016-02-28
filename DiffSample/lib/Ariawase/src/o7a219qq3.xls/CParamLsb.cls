VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamLsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   DeviceConfigurations�V�[�g��FixedLSB/DynamicLSB����ێ�����N���X
'
'   Revision History:
'       Data        Description
'       2010/11/19  FixedLSB/DynamicLSB����ێ�����@�\����������
'       2010/11/30  Clone/Initialize�@�\����������
'
'�쐬��:
'   0145184346
'

Option Explicit

Implements IDeviceParameter ' DeviceParameter�pI/F��錾����

Private Const ERR_NUMBER As Long = 9999             ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CParamLsb"    ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty ' Default�l"Empty"��ێ�����

Private Const PARAM_BASIC_LSB As String = "Basic Lsb@LSB Config"        ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_DYNAMIC_LSB As String = "Procedure Name@LSB Config" ' �p�����[�^�Q�ƃL�[��ێ�����
Private Const PARAM_NAME As String = "LSB"                              ' �p�����[�^���̂�ێ�����

Private Const MAIN_UNIT As String = "V" ' �P�ʖ��̂�ێ�����

Private mLsbSnapper As ILSBSnapper ' LSB�擾�p�C���^�[�t�F�[�X��ێ�����

Private mMainUnit As String             ' �P�ʏ���ێ�����
Private mSubUnit As String              ' �T�u�P�ʏ���ێ�����
Private mValue As String                ' �p�����[�^�l��ێ�����
Private mFlagDynamic As Boolean         ' �p�����[�^��ʂ�ێ�����
Private mDataArray As CParamDoubleArray ' �z��^�f�[�^��ێ�����

Public Sub Initialize(ByVal strData As String, Optional ByRef clsPrmDblAry As CParamDoubleArray = Nothing, Optional ByVal DynamicMode As Boolean = False)
'���e:
'   �p�����[�^�����i�[����
'
'�p�����[�^:
'   [strValue]      In  �p�����[�^�l����ێ�����
'   [clsPrmDblAry]  In  �z��Double�^�f�[�^��ێ�����
'   [DynamicMode]   In  �p�����[�^��ʏ���ێ�����
'
'�߂�l:
'
'���ӎ���:
'

    '#####  �p�����[�^���P�ʕt�����l�̏ꍇ�́A�P�ʂƐ��l�ɕ������i�[����  #####
    Call XLibTheDeviceProfilerUtility.DivideDeviceParameter(strData, mMainUnit, mSubUnit, mValue)
    mFlagDynamic = DynamicMode
    If Not (clsPrmDblAry Is Nothing) Then
        Set mDataArray = New CParamDoubleArray
        mDataArray.SetArrayAsDouble clsPrmDblAry.ReadAsDouble
    Else
        Set mDataArray = clsPrmDblAry
    End If
End Sub

Public Function Clone() As CParamLsb
'���e:
'   Clone���쐬����
'
'�p�����[�^:
'
'�߂�l:
'   CParamLsb(Clone)
'
'���ӎ���:
'

    '#####  ���g��Clone���쐬���Ė߂�  #####
    Dim cloneobj As CParamLsb
    Set cloneobj = New CParamLsb
    cloneobj.Initialize Me.Parameter, mDataArray, mFlagDynamic
    Set Clone = cloneobj
    Set cloneobj = Nothing

End Function

Public Sub ReadParameter(ByRef paramReader As IParameterReader, ByRef strParamName As String)
'���e:
'   �p�����[�^�����擾���i�[����
'
'�p�����[�^:
'   [paramReader]   In  DeviceConfigurations�V�[�g�̏���ێ�����
'   [strParamName]  In  �擾����p�����[�^�Q�ƃL�[�̏���ێ�����
'
'�߂�l:
'
'���ӎ���:
'

    '#####  �p�����[�^��P�ʂƐ��l�ɕ������āA�擾����  #####
    Call XLibTheDeviceProfilerUtility.GetDeviceParameter(paramReader, strParamName, MAIN_UNIT, mMainUnit, mSubUnit, mValue)

    '#####  LSB��ʂɂ���āALSB�l���擾����N���X��؂�ւ���  #####
    If mValue <> "" Then
        Select Case strParamName
            Case PARAM_BASIC_LSB ' �Œ�l���ϐ����̂̏ꍇ
                Dim clsFixedLSB As CFixedLsb
                Set clsFixedLSB = New CFixedLsb
                clsFixedLSB.Initialize mValue
                Set mLsbSnapper = clsFixedLSB
                Set clsFixedLSB = Nothing
                mFlagDynamic = False
            Case PARAM_DYNAMIC_LSB ' �}�N�����s����LSB�l���擾����ꍇ
                Dim clsDynamicLsb As CDynamicLsb
                Set clsDynamicLsb = New CDynamicLsb
                clsDynamicLsb.Initialize mValue
                Set mLsbSnapper = clsDynamicLsb
                Set clsDynamicLsb = Nothing
                mFlagDynamic = True
        End Select
    End If

End Sub

Public Property Get Parameter() As String
'���e:
'   �擾�����p�����[�^����߂�
'
'�p�����[�^:
'
'�߂�l:
'   �擾�����p�����[�^���
'
'���ӎ���:
'

    '#####  ���l�̏ꍇ�́A�T�u�P�ʂ𗘗p���āA�P�ʊ��Z�����{����  #####
    Parameter = mValue
    If IsNumeric(mValue) = True Then
        Parameter = CStr(CDbl(mValue) / SubUnitToValue(mSubUnit)) & mSubUnit & mMainUnit
    End If

End Property

Public Sub CalcLsb(Optional ByVal dblGainRatio As Variant)
'���e:
'   LSB�l���Z�o����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'

    '#####  ���l�f�[�^�̗L�����m�F����  #####
    If mLsbSnapper Is Nothing Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CalcLsb", "Parameter is Nothing.[" & PARAM_NAME & "]")
    End If

    '#####  Static�̏ꍇ�́AGain�{���Ŋ���߂�  #####
    '#####  Dynamic�̏ꍇ�́A�������Ȃ�         #####
    Dim dblArray() As Double
    If mFlagDynamic = False Then
        Dim i As Long
        Dim dblLsb() As Double
        ReDim dblArray(UBound(dblGainRatio))
        dblLsb = mLsbSnapper.Lsb.ReadAsDouble
        For i = 0 To UBound(dblGainRatio) Step 1
            If dblGainRatio(i) <> 0 Then
                dblArray(i) = dblLsb(i) / dblGainRatio(i)
            Else
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CalcLsb", "Couldn't Calculated LSB.[GainRatio = 0]")
            End If
        Next i
    Else
        dblArray = mLsbSnapper.Lsb.ReadAsDouble
    End If
    
    '#####  �f�[�^���i�[����  #####
    Set mDataArray = New CParamDoubleArray
    mDataArray.SetArrayAsDouble dblArray

End Sub

Private Function calcUnit(ByVal dblValue As Double) As String

    '#####  �T�u�P�ʂ��K�v���m�F����  #####
    If dblValue = 0 Then
        calcUnit = ""
        Exit Function
    ElseIf dblValue <= 1 * SubUnitToValue("p") Then
        calcUnit = "p"
        Exit Function
    End If

    '#####  �œK�ȃT�u�P�ʂ����߂�  #####
    Dim strUnit As Variant
    Dim dblCalc As Double
    For Each strUnit In Array("m", "u", "n", "p")
        dblCalc = dblValue / SubUnitToValue(strUnit)
        If dblCalc > 1 Then
            calcUnit = strUnit
            Exit Function
        End If
    Next

End Function

Private Sub Class_Initialize()
    mMainUnit = INITIAL_EMPTY_VALUE
    mSubUnit = INITIAL_EMPTY_VALUE
    mValue = INITIAL_EMPTY_VALUE
    mFlagDynamic = False
    Set mDataArray = Nothing
    Set mLsbSnapper = Nothing
End Sub

Private Sub Class_Terminate()
    mMainUnit = INITIAL_EMPTY_VALUE
    mSubUnit = INITIAL_EMPTY_VALUE
    mValue = INITIAL_EMPTY_VALUE
    mFlagDynamic = False
    Set mDataArray = Nothing
    Set mLsbSnapper = Nothing
End Sub

Private Function IDeviceParameter_AsAccTimeH() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsAccTimeH", "UnSupported Data Type.[AccTimeH]")
End Function

Private Function IDeviceParameter_AsAccTimeV() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsAccTimeV", "UnSupported Data Type.[AccTimeV]")
End Function

Private Function IDeviceParameter_AsDecibel() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsDecibel", "UnSupported Data Type.[Decibel]")
End Function

Private Function IDeviceParameter_AsDouble() As Double()

    '#####  ���l�f�[�^�̗L�����m�F����  #####
    If mDataArray Is Nothing Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsDouble", "Parameter is Nothing.[" & PARAM_NAME & "]")
    End If
    IDeviceParameter_AsDouble = mDataArray.ReadAsDouble

End Function

Private Function IDeviceParameter_AsRatio() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsRatio", "UnSupported Data Type.[Ratio]")
End Function

Private Function IDeviceParameter_AsString() As String()
    
    '#####  ���l�f�[�^�̗L�����m�F����  #####
    If mDataArray Is Nothing Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsString", "Parameter is Nothing.[" & PARAM_NAME & "]")
    End If

    '#####  ���l�f�[�^���擾����  #####
    Dim dblValue() As Double
    dblValue = mDataArray.ReadAsDouble
    
    '#####  �T�u�P�ʂ𗘗p���āA�P�ʊ��Z�������l���Z�o���āA�P�ʕt�����l�ɕϊ�����  #####
    Dim i As Integer
    Dim strValue() As String
    Dim strSubUnit As String
    ReDim strValue(UBound(dblValue))
    strSubUnit = mSubUnit ' �T�u�P�ʂ��i�[����
    For i = 0 To UBound(dblValue) Step 1
        '#####  �����l���A�ϐ����̂�}�N�����̂̏ꍇ�́A�T�u�P�ʂ����߂�  #####
        If IsAlphabet(mValue) = True Then
            strSubUnit = calcUnit(dblValue(i))
        End If
        '#####  �T�u�P�ʂ𗘗p���āA�P�ʊ��Z�������l���Z�o���āA�P�ʕt�����l�ɕϊ�����  #####
        strValue(i) = CStr(Format((dblValue(i) / SubUnitToValue(strSubUnit)), "0.0###")) & strSubUnit & MAIN_UNIT
    Next i
    IDeviceParameter_AsString = strValue

End Function

Private Function IDeviceParameter_AsStringAccTimeH() As String()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsStringAccTimeH", "UnSupported Data Type.[StringAccTimeH]")
End Function

Private Function IDeviceParameter_AsStringAccTimeV() As String()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsStringAccTimeV", "UnSupported Data Type.[StringAccTimeV]")
End Function

Public Function AsIDeviceParameter() As IDeviceParameter
    Set AsIDeviceParameter = Me
End Function