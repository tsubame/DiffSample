VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFixedLsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �f�W�^���R�[�h���Z�pLSB�l�̃f�[�^���擾���ێ�����N���X
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfigurations�V�[�g��FixedLSB����ێ�����@�\����������
'
'�쐬��:
'   0145184346
'

Option Explicit

Implements ILSBSnapper ' LSB�擾�pI/F��錾����

Private Const ERR_NUMBER As Long = 9999             ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CFixedLsb"    ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty ' Default�l"Empty"��ێ�����

Private mLSBValue As String ' LSB�l��ێ�����

Private Property Get ILSBSnapper_LSB() As CParamDoubleArray

    '#####  LSB�l��LSB�ݒ�ϐ����̂���f�[�^���擾����  #####
    If mLSBValue <> "" Then
        Dim clsPrmDblAry As CParamDoubleArray ' �p�����[�^����ێ�����
        Set clsPrmDblAry = New CParamDoubleArray
        If IsNumeric(mLSBValue) = True Then
            clsPrmDblAry.SetAsDouble CDbl(mLSBValue) ' ���l�̏ꍇ
        Else
            clsPrmDblAry.SetArrayAsDouble XLibTheDeviceProfilerUtility.GetLSBParam(mLSBValue) ' �ϐ����̂̏ꍇ
        End If
    Else
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".LSB", "Nothing LSB Value.")
    End If

    '#####  �I��  #####
    Set ILSBSnapper_LSB = clsPrmDblAry
    Set clsPrmDblAry = Nothing

End Property

Public Sub Initialize(ByRef strLSBData As String)
    mLSBValue = strLSBData
End Sub

Private Sub Class_Initialize()
    mLSBValue = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    mLSBValue = INITIAL_EMPTY_VALUE
End Sub