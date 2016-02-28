Attribute VB_Name = "XLibTheDeviceProfilerUtility"
'�T�v:
'   TheDeviceProfiler�̃��[�e�B���e�B
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfig��Utility�@�\����������
'       2010/11/30  DivideDeviceParameter��ǉ�����
'       2010/12/07  �s�v�R�[�h���폜����
'
'�쐬��:
'   0145184346
'

Option Explicit

Public TheDeviceProfiler As CDeviceProfiler ' DeviceProfiler��錾����

Private Const ERR_NUMBER = 9999                           ' Error�ԍ���ێ�����
Private Const CLASS_NAME = "XLibTheDeviceProfilerUtility" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty       ' Default�l"Empty"��ێ�����

Private mLSBParam As Collection         ' LSB�ϐ���ێ�����
Private mAccTimeParam As Collection     ' AccTime�ϐ���ێ�����
Private mAccTimeParamUnit As Collection ' AccTime�ϐ��ɑΉ�����P�ʂ�ێ�����

Public Sub CreateTheDeviceProfilerIfNothing()
'���e:
'   TheDeviceProfiler������������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    On Error GoTo ErrHandler
    If TheDeviceProfiler Is Nothing Then
        Set TheDeviceProfiler = New CDeviceProfiler
        Call TheDeviceProfiler.Initialize
    End If
    Exit Sub
ErrHandler:
    Set TheDeviceProfiler = Nothing
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

Public Sub InitializeTheDeviceProfiler()
'���e:
'   TheDeviceProfiler������������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Call TheDeviceProfiler.Initialize
End Sub

Public Sub DestroyTheDeviceProfiler()
'���e:
'   TheDeviceProfiler��j������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Set TheDeviceProfiler = Nothing
    Set mLSBParam = Nothing
    Set mAccTimeParam = Nothing
    Set mAccTimeParamUnit = Nothing
End Sub

Public Function RunAtJobEnd() As Long
End Function

Public Sub SetLSBParam(ByVal strName As String, ByRef dblValue() As Double)
'���e:
'   �ϐ������ꂽLSB�l���i�[����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'

    '#####  �ϐ��l���i�[����  #####
    If mLSBParam Is Nothing Then
        Set mLSBParam = New Collection
    End If
    On Error GoTo ErrHandler
    mLSBParam.Remove strName
    mLSBParam.Add dblValue, strName
    Exit Sub
ErrHandler:
    mLSBParam.Add dblValue, strName
    Exit Sub
End Sub

Public Function GetLSBParam(ByRef strName As String) As Double()
'���e:
'   �ϐ������ꂽLSB�l���擾����
'
'�p�����[�^:
'
'�߂�l:
'   �z��Double�^(LSB�l)
'
'���ӎ���:
'

    '#####  �ϐ��l���擾����  #####
    On Error GoTo ErrHandler
    GetLSBParam = mLSBParam.Item(strName)
    Exit Function
ErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".GetLSBParam", "Nothing Parameter Name.[" & strName & "]")
    Exit Function
End Function

Public Sub SetAccTimeParam(ByVal strName As String, ByVal pramValue As Variant, Optional ByVal strUnit As String = "H")
'���e:
'   �ϐ������ꂽAccTime�l���i�[����
'
'�p�����[�^:
'   [strName]    In  �ϐ����̂�ێ�����
'   [pramValue]  In  �ϐ��̐��l��ێ�����
'   [strUnit]    In  �ϐ��l�̒P�ʖ��̂�ێ�����
'
'�߂�l:
'
'���ӎ���:
'

    '#####�@�P�ʖ��̂��m�F����  #####
    If (strUnit <> "H") And (strUnit <> "V") Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".SetAccTimeParam", "Illegal Unit. [""H""or""V""]")
    End If

    '#####  �ϐ��l�ƒP�ʖ��̂��i�[����  #####
    If (mAccTimeParam Is Nothing) Or (mAccTimeParamUnit Is Nothing) Then
        Set mAccTimeParam = New Collection
        Set mAccTimeParamUnit = New Collection
    End If
    On Error GoTo ErrHandler
    mAccTimeParam.Remove strName
    mAccTimeParam.Add pramValue, strName
    mAccTimeParamUnit.Remove strName
    mAccTimeParamUnit.Add strUnit, strName
    Exit Sub
ErrHandler:
    mAccTimeParam.Add pramValue, strName
    mAccTimeParamUnit.Add strUnit, strName
    Exit Sub
End Sub

Public Sub GetAccTimeParam(ByRef strName As String, ByRef strUnit As String, ByRef dblValue() As Double)
'���e:
'   �ϐ������ꂽAccTime�l���擾����
'
'�p�����[�^:
'   [strName]   In  �ϐ����̂�ێ�����
'   [strUnit]   In  �ϐ��l�̒P�ʖ��̂�ێ�����
'   [dblValue]  In  �ϐ����̂̐��l��ێ�����
'
'�߂�l:
'   CDeviceParamArray(�N���X)
'
'���ӎ���:
'

    '#####  �ϐ��l�ƒP�ʖ��̂��擾����  #####
    On Error GoTo ErrHandler
    strUnit = mAccTimeParamUnit.Item(strName)
    dblValue = mAccTimeParam.Item(strName)
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".GetAccTimeParam", "Nothing Parameter Name.[" & strName & "]")
    Exit Sub

End Sub

Public Sub GetDeviceParameter(ByRef paramReader As IParameterReader, ByRef strParamName As String, ByRef strFixedMainUnit As String, _
                            ByRef retMainUnit As String, ByRef retSubUnit As String, ByRef retValue As String)
'���e:
'   DeviceConfigurations�V�[�g�̃p�����[�^�l�𕪉����i�[����
'
'�p�����[�^:
'   [paramReader]       In  DeviceConfigurations�V�[�g�̏���ێ�����
'   [strParamName]      In  �擾����p�����[�^���̂�ێ�����
'   [strFixedMainUnit]  In  �擾����p�����[�^�̒P�ʖ��̂�ێ�����
'
'�߂�l:
'   [retMainUnit]  Out  �擾�����p�����[�^�̒P�ʖ��̂�߂�
'   [retSubUnit]   Out  �擾�����p�����[�^�̃T�u�P�ʖ��̂�߂�
'   [retValue]     Out  �擾�����p�����[�^�̒l��߂�
'
'���ӎ���:
'

    '#####  �p�����[�^���擾���āA��������m�F����  #####
    Dim strData As String
    strData = paramReader.ReadAsString(strParamName)
    Call CheckAsString(strData)

    '#####  �p�����[�^���P�ʕt�����l�̏ꍇ�́A�P�ʂƐ��l�ɕ�������  #####
    If IsNumeric(strData) = True Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".GetDeviceParameter", _
                        strParamName & " : [" & strData & "]  - This Is Not Appropriate Main Unit !")
    End If
    Call DivideDeviceParameter(strData, retMainUnit, retSubUnit, retValue)
    If (IsAlphabet(strData) = False) And (IsNumeric(strData) = False) Then
        '#####  �z�肵�Ă���P�ʂƐݒ肵���P�ʂ��s��v�Ȃ�΁A�G���[�Ƃ���  #####
        If (strFixedMainUnit <> retMainUnit) And (strFixedMainUnit <> "") Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".GetDeviceParameter", _
                            strParamName & " : [" & strData & "]  - This Is Not Appropriate Main Unit !")
        End If
    End If

End Sub

Public Sub DivideDeviceParameter(ByRef strData As String, ByRef retMainUnit As String, ByRef retSubUnit As String, ByRef retValue As String)
'���e:
'   DeviceConfigurations�V�[�g�̃p�����[�^�l�𕪉����i�[����
'
'�p�����[�^:
'   [strData]  In  �擾����p�����[�^���̂�ێ�����
'
'�߂�l:
'   [retMainUnit]  Out  �擾�����p�����[�^�̒P�ʖ��̂�߂�
'   [retSubUnit]   Out  �擾�����p�����[�^�̃T�u�P�ʖ��̂�߂�
'   [retValue]     Out  �擾�����p�����[�^�̒l��߂�
'
'���ӎ���:
'

    '#####  �p�����[�^���P�ʕt�����l�̏ꍇ�́A�P�ʂƐ��l�ɕ������i�[����  #####
    Dim strMainUnit As String
    Dim strSubUnit As String
    Dim dblSubValue As Double
    
    If (IsNumeric(strData) = True) Or (IsAlphabet(strData) = True) Or (strData = "") Then
        retMainUnit = INITIAL_EMPTY_VALUE
        retSubUnit = INITIAL_EMPTY_VALUE
        retValue = strData
    Else
        Call SplitUnitValue(CStr(strData), strMainUnit, strSubUnit, dblSubValue)
        retMainUnit = strMainUnit
        retSubUnit = strSubUnit
        retValue = CStr(dblSubValue * SubUnitToValue(strSubUnit))
    End If

End Sub
