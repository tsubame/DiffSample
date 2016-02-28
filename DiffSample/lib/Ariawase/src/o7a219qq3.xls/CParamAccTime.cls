VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamAccTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   DeviceConfigurations�V�[�g��AccTime����ێ�����N���X
'
'   Revision History:
'       Data        Description
'       2010/11/19  AccTime����ێ�����@�\����������
'       2010/11/30  Clone/Initialize�@�\����������
'
'�쐬��:
'   0145184346
'

Option Explicit

Implements IDeviceParameter ' DeviceParameter�pI/F��錾����

Private Const ERR_NUMBER As Long = 9999              ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CParamAccTime" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty  ' Default�l"Empty"��ێ�����

Private Const MAIN_UNIT_V As String = "V"               ' �P�ʖ��̂�ێ�����
Private Const MAIN_UNIT_H As String = "H"               ' �P�ʖ��̂�ێ�����
Private Const PARAM_NAME_ACCTIME As String = "Acc Time" ' �p�����[�^���̂�ێ�����
Private Const PARAM_NAME_XHSpV As String = "XHS/V"      ' �p�����[�^���̂�ێ�����

Private mMainUnit_XHSpV As String                ' �P�ʏ���ێ�����
Private mSubUnit_XHSpV As String                 ' �T�u�P�ʏ���ێ�����
Private mValue_XHSpV As String                   ' �p�����[�^�l��ێ�����
Private mMainUnit_AccTime As String              ' �P�ʏ���ێ�����
Private mSubUnit_AccTime As String               ' �T�u�P�ʏ���ێ�����
Private mValue_AccTime As String                 ' �p�����[�^�l��ێ�����
Private mDataArray_AccTimeH As CParamDoubleArray ' �z��^�f�[�^��ێ�����
Private mDataArray_AccTimeV As CParamDoubleArray ' �z��^�f�[�^��ێ�����
Private mArrange_Unit As String                  ' �v�Z���O�Ɏ擾����AccTime�̒P��
Private mArrange_Value() As Double               ' �v�Z���O�Ɏ擾����AccTime�̒l

Public Sub Initialize(ByVal strXHSpV As String, ByVal strAccTime As String, ByRef dblParam() As Double, ByVal strPrmUnit As String)
'���e:
'   �p�����[�^�����i�[����
'
'�p�����[�^:
'   [strXHSpV]    In  XHS/V�p�����[�^�l����ێ�����
'   [strAccTime]  In  AccTime�p�����[�^�l����ێ�����
'   [dblParam]    In  �Z�o�O�p�����[�^�̔z��Double�^�f�[�^��ێ�����
'   [strPrmUnit]  In  �Z�o�O�p�����[�^�̒P�ʖ��̂�ێ�����
'
'�߂�l:
'
'���ӎ���:
'

    '#####  �p�����[�^���P�ʕt�����l�̏ꍇ�́A�P�ʂƐ��l�ɕ������i�[����  #####
    Call XLibTheDeviceProfilerUtility.DivideDeviceParameter(strXHSpV, mMainUnit_XHSpV, mSubUnit_XHSpV, mValue_XHSpV)
    Call XLibTheDeviceProfilerUtility.DivideDeviceParameter(strAccTime, mMainUnit_AccTime, mSubUnit_AccTime, mValue_AccTime)
    mArrange_Value = dblParam
    mArrange_Unit = strPrmUnit

End Sub

Public Function Clone() As CParamAccTime
'���e:
'   Clone���쐬����
'
'�p�����[�^:
'
'�߂�l:
'   CParamAccTime(Clone)
'
'���ӎ���:
'

    '#####  ���g��Clone���쐬���Ė߂�  #####
    Dim cloneobj As CParamAccTime
    Set cloneobj = New CParamAccTime
    cloneobj.Initialize Me.ParameterXHSpV, Me.ParameterAccTime, mArrange_Value, mArrange_Unit
    Set Clone = cloneobj
    Set cloneobj = Nothing

End Function

Public Sub ReadParameter(ByRef paramReader As IParameterReader, ByRef strXHSpVName As String, ByRef strAccTimeName As String)
'���e:
'   �p�����[�^�����擾���i�[����
'
'�p�����[�^:
'   [paramReader]     In  DeviceConfigurations�V�[�g�̏���ێ�����
'   [strXHSpVName]    In  �擾����p�����[�^�Q�ƃL�[�̏���ێ�����
'   [strAccTimeName]  In  �擾����p�����[�^�Q�ƃL�[�̏���ێ�����
'
'�߂�l:
'
'���ӎ���:
'

    '#####  �p�����[�^��P�ʂƐ��l�ɕ������āA�擾����  #####
    Call XLibTheDeviceProfilerUtility.GetDeviceParameter(paramReader, strXHSpVName, MAIN_UNIT_H, mMainUnit_XHSpV, mSubUnit_XHSpV, mValue_XHSpV)
    Call XLibTheDeviceProfilerUtility.GetDeviceParameter(paramReader, strAccTimeName, "", mMainUnit_AccTime, mSubUnit_AccTime, mValue_AccTime)

    '#####  �P�ʖ���(V/H)���m�F����  #####
    If mValue_AccTime <> "" Then
        If IsNumeric(mValue_AccTime) = True Then
            If (mMainUnit_AccTime <> MAIN_UNIT_V) And (mMainUnit_AccTime <> MAIN_UNIT_H) Then
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".SetupParameter", _
                                strAccTimeName & " : [" & paramReader.ReadAsString(strAccTimeName) & "]  - This Is Not Appropriate Main Unit !")
            End If
        End If
    End If

    '#####  �T�u�P�ʐݒ�������Ȃ��̂ŁA�G���[�`�F�b�N�����{����  #####
    If mSubUnit_XHSpV <> "" Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".SetupParameter", _
                        strXHSpVName & " : [" & paramReader.ReadAsString(strXHSpVName) & "]  - This Is Not Appropriate Sub Unit !")
    ElseIf mSubUnit_AccTime <> "" Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".SetupParameter", _
                        strAccTimeName & " : [" & paramReader.ReadAsString(strAccTimeName) & "]  - This Is Not Appropriate Sub Unit !")
    End If

End Sub

Public Property Get ParameterXHSpV() As String
'���e:
'   �ێ����Ă��p�����[�^���(XHS/V)��߂�
'
'�p�����[�^:
'
'�߂�l:
'   �ێ����Ă��p�����[�^���
'
'���ӎ���:
'

    '#####  �P�ʂ�t������  #####
    ParameterXHSpV = mValue_XHSpV & mMainUnit_XHSpV

End Property

Public Property Get ParameterAccTime() As String
'���e:
'   �ێ����Ă���p�����[�^���(AccTime)��߂�
'
'�p�����[�^:
'
'�߂�l:
'   �ێ����Ă���p�����[�^���
'
'���ӎ���:
'

    '#####  �P�ʂ�t������  #####
    ParameterAccTime = mValue_AccTime & mMainUnit_AccTime

End Property

Private Sub paramAccTimeIsEmpty()
    If mValue_AccTime = "" Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".paramAccTimeIsEmpty", "Parameter is Nothing.[" & PARAM_NAME_ACCTIME & "]")
    End If
End Sub

Private Sub paramXHSpVIsEmpty()
    If mValue_XHSpV = "" Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".paramXHSpVIsEmpty", "Parameter is Nothing.[" & PARAM_NAME_XHSpV & "]")
    End If
End Sub

Private Sub paramXHSpVIsNumeric()
    If IsNumeric(mValue_XHSpV) = False Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".paramXHSpVIsNumeric", "Unsupported Variable.[" & mValue_XHSpV & "]")
    End If
End Sub

Public Sub ArrangeParameter()
'���e:
'   AccTime�Z�o���邽�߂ɁA�ϐ����̂��琔�l���P�ʂ��擾����
'   ���l�̏ꍇ�́A����̂ݎ��{����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    
    '#####  �ϐ����̂̏ꍇ�́A���l���P�ʂ��擾����  #####
    If IsAlphabet(mValue_AccTime) = True Then
        Call XLibTheDeviceProfilerUtility.GetAccTimeParam(mValue_AccTime, mArrange_Unit, mArrange_Value) ' �ϐ��l�ƒP�ʂ��擾����
    ElseIf IsNumeric(mValue_AccTime) = True Then
        Dim i As Integer
        ReDim mArrange_Value(TheExec.sites.ExistingCount - 1)
        For i = 0 To (TheExec.sites.ExistingCount - 1) Step 1
            mArrange_Value(i) = CDbl(mValue_AccTime)
        Next i
        mArrange_Unit = mMainUnit_AccTime
    End If

End Sub

Private Sub calcAccTimeByUnit(ByRef strCalcType As String)
'���e:
'   �擾�����p�����[�^�𗘗p���āA�w�肵���P�ʂł�AccTime���Z�o����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'

    '#####  �f�[�^���ݒ肳��Ă��Ȃ����߁A�G���[�Ƃ���  #####
    Call paramAccTimeIsEmpty

    '#####  �P�ʂ��قȂ�ꍇ�́A���Z���K�v�ƂȂ�              #####
    '#####  �ϐ����̂̏ꍇ�́A�i�[�悩��l���擾���Ċ��Z����  #####
    Dim i As Long
    Dim dblCalcValue() As Double
    ReDim dblCalcValue(UBound(mArrange_Value))
    If (strCalcType = MAIN_UNIT_H) And (mArrange_Unit = MAIN_UNIT_V) Then ' "H"��ŎZ�o���P�ʂ�"V"�̏ꍇ�ɎZ�o����
        Call paramXHSpVIsEmpty
        Call paramXHSpVIsNumeric
        For i = 0 To UBound(mArrange_Value) Step 1
            dblCalcValue(i) = mArrange_Value(i) * CDbl(mValue_XHSpV)
        Next i
    ElseIf (strCalcType = MAIN_UNIT_V) And (mArrange_Unit = MAIN_UNIT_H) Then ' "V"��ŎZ�o���P�ʂ�"H"�̏ꍇ�ɎZ�o����
        Call paramXHSpVIsEmpty
        Call paramXHSpVIsNumeric
        For i = 0 To UBound(mArrange_Value) Step 1
            If CDbl(mValue_XHSpV) <> 0 Then
                dblCalcValue(i) = mArrange_Value(i) / CDbl(mValue_XHSpV)
            Else
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".calcAccTimeByUnit", "Couldn't Calculated AccTime.[XHS/V = 0]")
            End If
        Next i
    Else
        dblCalcValue = mArrange_Value ' �P�ʂ���v���Ă���ꍇ�́A�������Ȃ�
    End If
    
    
    '#####  �Z�o�����f�[�^���i�[����  #####
    Dim clsDataArray As CParamDoubleArray
    Set clsDataArray = New CParamDoubleArray
    clsDataArray.SetArrayAsDouble dblCalcValue

    '#####  �Z�o�����f�[�^�����W���[���ϐ��Ɋi�[����  #####
    Select Case strCalcType
        Case MAIN_UNIT_H
            Set mDataArray_AccTimeH = clsDataArray
        Case MAIN_UNIT_V
            Set mDataArray_AccTimeV = clsDataArray
    End Select
    Set clsDataArray = Nothing

End Sub

Private Sub Class_Initialize()
    mMainUnit_XHSpV = INITIAL_EMPTY_VALUE
    mSubUnit_XHSpV = INITIAL_EMPTY_VALUE
    mValue_XHSpV = INITIAL_EMPTY_VALUE
    mMainUnit_AccTime = INITIAL_EMPTY_VALUE
    mSubUnit_AccTime = INITIAL_EMPTY_VALUE
    mValue_AccTime = INITIAL_EMPTY_VALUE
    Set mDataArray_AccTimeH = Nothing
    Set mDataArray_AccTimeV = Nothing
    mArrange_Unit = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    mMainUnit_XHSpV = INITIAL_EMPTY_VALUE
    mSubUnit_XHSpV = INITIAL_EMPTY_VALUE
    mValue_XHSpV = INITIAL_EMPTY_VALUE
    mMainUnit_AccTime = INITIAL_EMPTY_VALUE
    mSubUnit_AccTime = INITIAL_EMPTY_VALUE
    mValue_AccTime = INITIAL_EMPTY_VALUE
    Set mDataArray_AccTimeH = Nothing
    Set mDataArray_AccTimeV = Nothing
    mArrange_Unit = INITIAL_EMPTY_VALUE
End Sub

Private Function IDeviceParameter_AsAccTimeH() As Double()

    '#####  ���l�f�[�^�̗L�����m�F����  #####
    If mDataArray_AccTimeH Is Nothing Then
        Call calcAccTimeByUnit("H")
    End If
    IDeviceParameter_AsAccTimeH = mDataArray_AccTimeH.ReadAsDouble

End Function

Private Function IDeviceParameter_AsAccTimeV() As Double()

    '#####  ���l�f�[�^�̗L�����m�F����  #####
    If mDataArray_AccTimeV Is Nothing Then
        Call calcAccTimeByUnit("V")
    End If
    IDeviceParameter_AsAccTimeV = mDataArray_AccTimeV.ReadAsDouble

End Function

Private Function IDeviceParameter_AsDecibel() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsDecibel", "UnSupported Data Type.[Decibel]")
End Function

Private Function IDeviceParameter_AsDouble() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsDouble", "UnSupported Data Type.[Double]")
End Function

Private Function IDeviceParameter_AsRatio() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsRatio", "UnSupported Data Type.[Ratio]")
End Function

Private Function IDeviceParameter_AsString() As String()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsString", "UnSupported Data Type.[String]")
End Function

Private Function IDeviceParameter_AsStringAccTimeH() As String()

    '#####  ���l�f�[�^�̗L�����m�F����  #####
    If mDataArray_AccTimeH Is Nothing Then
        Call calcAccTimeByUnit("H")
    End If

    '#####  ���l�f�[�^���擾����  #####
    Dim dblValue() As Double
    dblValue = mDataArray_AccTimeH.ReadAsDouble

    '#####  �T�u�P�ʂ𗘗p���āA�P�ʊ��Z�������l���Z�o���āA�P�ʕt�����l�ɕϊ�����  #####
    Dim i As Integer
    Dim strValue() As String
    ReDim strValue(UBound(dblValue))
    For i = 0 To UBound(dblValue) Step 1
        '#####  �����_�ȉ��̌������݂��鐔�l�́A�����_�ȉ�4���Ƃ���  #####
        If dblValue(i) Like "*.*" Then
            strValue(i) = CStr(Format(dblValue(i), "0.0###")) & MAIN_UNIT_H
        Else
            strValue(i) = CStr(dblValue(i)) & MAIN_UNIT_H
        End If
    Next i
    IDeviceParameter_AsStringAccTimeH = strValue

End Function

Private Function IDeviceParameter_AsStringAccTimeV() As String()

    '#####  ���l�f�[�^�̗L�����m�F����  #####
    If mDataArray_AccTimeV Is Nothing Then
        Call calcAccTimeByUnit("V")
    End If

    '#####  ���l�f�[�^���擾����  #####
    Dim dblValue() As Double
    dblValue = mDataArray_AccTimeV.ReadAsDouble

    '#####  �T�u�P�ʂ𗘗p���āA�P�ʊ��Z�������l���Z�o���āA�P�ʕt�����l�ɕϊ�����  #####
    Dim i As Integer
    Dim strValue() As String
    ReDim strValue(UBound(dblValue))
    For i = 0 To UBound(dblValue) Step 1
        '#####  �����_�ȉ��̌������݂��鐔�l�́A�����_�ȉ�4���Ƃ���  #####
        If dblValue(i) Like "*.*" Then
            strValue(i) = CStr(Format(dblValue(i), "0.0###")) & MAIN_UNIT_V
        Else
            strValue(i) = CStr(dblValue(i)) & MAIN_UNIT_V
        End If
    Next i
    IDeviceParameter_AsStringAccTimeV = strValue

End Function

Public Function AsIDeviceParameter() As IDeviceParameter
    Set AsIDeviceParameter = Me
End Function