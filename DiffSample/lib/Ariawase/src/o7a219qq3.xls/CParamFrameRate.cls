VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamFrameRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   DeviceConfigurations�V�[�g��FrameRate����ێ�����N���X
'
'   Revision History:
'       Data        Description
'       2010/11/19  FrameRate����ێ�����@�\����������
'       2010/11/30  Clone/Initialize�@�\����������
'
'�쐬��:
'   0145184346
'

Option Explicit

Implements IDeviceParameter ' DeviceParameter�pI/F��錾����

Private Const ERR_NUMBER As Long = 9999                ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CParamFrameRate" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty    ' Default�l"Empty"��ێ�����

Private Const MAIN_UNIT As String = "fps"         ' �P�ʖ��̂�ێ�����
Private Const PARAM_NAME As String = "Frame Rate" ' �p�����[�^���̂�ێ�����

Private mMainUnit As String             ' �P�ʏ���ێ�����
Private mSubUnit As String              ' �T�u�P�ʏ���ێ�����
Private mValue As String                ' �p�����[�^�l��ێ�����
Private mDataArray As CParamDoubleArray ' �z��^�f�[�^��ێ�����

Public Sub Initialize(ByVal strData As String, Optional ByRef clsPrmDblAry As CParamDoubleArray = Nothing)
'���e:
'   �p�����[�^�����i�[����
'
'�p�����[�^:
'   [strValue]      In  �p�����[�^�l����ێ�����
'   [clsPrmDblAry]  In  �z��Double�^�f�[�^��ێ�����
'
'�߂�l:
'
'���ӎ���:
'

    '#####  �p�����[�^���P�ʕt�����l�̏ꍇ�́A�P�ʂƐ��l�ɕ������i�[����  #####
    Call XLibTheDeviceProfilerUtility.DivideDeviceParameter(strData, mMainUnit, mSubUnit, mValue)
    If Not (clsPrmDblAry Is Nothing) Then
        Set mDataArray = New CParamDoubleArray
        mDataArray.SetArrayAsDouble clsPrmDblAry.ReadAsDouble
    Else
        Set mDataArray = clsPrmDblAry
    End If
End Sub

Public Function Clone() As CParamFrameRate
'���e:
'   Clone���쐬����
'
'�p�����[�^:
'
'�߂�l:
'   CParamFrameRate(Clone)
'
'���ӎ���:
'

    '#####  ���g��Clone���쐬���Ė߂�  #####
    Dim cloneobj As CParamFrameRate
    Set cloneobj = New CParamFrameRate
    cloneobj.Initialize Me.Parameter, mDataArray
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

    '#####  ���l��Site�W�J���āA�i�[����  #####
    If mValue <> "" Then
        Set mDataArray = New CParamDoubleArray
        mDataArray.SetAsDouble mValue
    End If

End Sub

Public Property Get Parameter() As String
'���e:
'   �ێ����Ă���p�����[�^����߂�
'
'�p�����[�^:
'
'�߂�l:
'   �ێ����Ă���p�����[�^���
'
'���ӎ���:
'

    '#####  �T�u�P�ʂ𗘗p���āA�P�ʊ��Z�����{����  #####
    Parameter = mValue
    If IsNumeric(mValue) = True Then
        Parameter = CStr(CDbl(mValue) / SubUnitToValue(mSubUnit)) & mSubUnit & mMainUnit
    End If

End Property

Private Sub Class_Initialize()
    mMainUnit = INITIAL_EMPTY_VALUE
    mSubUnit = INITIAL_EMPTY_VALUE
    mValue = INITIAL_EMPTY_VALUE
    Set mDataArray = Nothing
End Sub

Private Sub Class_Terminate()
    mMainUnit = INITIAL_EMPTY_VALUE
    mSubUnit = INITIAL_EMPTY_VALUE
    mValue = INITIAL_EMPTY_VALUE
    Set mDataArray = Nothing
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
    ReDim strValue(UBound(dblValue))
    For i = 0 To UBound(dblValue) Step 1
        strValue(i) = CStr((dblValue(i) / SubUnitToValue(mSubUnit))) & mSubUnit & mMainUnit
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