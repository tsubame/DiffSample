VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDynamicLsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �f�W�^���R�[�h���Z�pLSB�l�̃f�[�^���擾���ێ�����N���X
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfigurations�V�[�g��DynamicLSB����ێ�����@�\����������
'
'�쐬��:
'   0145184346
'

Option Explicit

Implements ILSBSnapper ' LSB�擾�pI/F��錾����

Private Const ERR_NUMBER As Long = 9999             ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CDynamicLsb"  ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty ' Default�l"Empty"��ێ�����

Private mMacroName As String ' LSB�l�擾�p��Macro���̂�ێ�����

Private Property Get ILSBSnapper_LSB() As CParamDoubleArray

    '#####  LSB�擾�p��UserMacro�����s���āALSB�l���擾����  #####
    If mMacroName <> "" Then

        '#####  UserMacro�����s���f�[�^���i�[����  #####
        If IsAlphabet(mMacroName) = True Then

            '#####  UserMacro�����s����  #####
            Dim dblArray() As Double ' Macro����̖߂�l��ێ�����
            On Error GoTo ErrHandler
            dblArray = Application.Run(mMacroName)
            On Error GoTo 0
    
            '#####  �f�[�^���i�[����  #####
            Dim clsPrmDblAry As CParamDoubleArray ' �p�����[�^����ێ�����
            Set clsPrmDblAry = New CParamDoubleArray
            clsPrmDblAry.SetArrayAsDouble dblArray
            Set ILSBSnapper_LSB = clsPrmDblAry
            Set clsPrmDblAry = Nothing

        End If

    Else
        Set ILSBSnapper_LSB = Nothing
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".LSB", "Nothing Macro Name.")
    End If

    '#####  �I��  #####
    Exit Property

ErrHandler:
    Set ILSBSnapper_LSB = Nothing
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".LSB", "Mismatch Macro Name.[" & mMacroName & "]")
    Exit Property
End Property

Public Sub Initialize(ByRef strMacroName As String)
    mMacroName = strMacroName
End Sub

Private Sub Class_Initialize()
    mMacroName = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    mMacroName = INITIAL_EMPTY_VALUE
End Sub