VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamDoubleArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   Double�^�̃f�[�^��z��Double�^�ŕێ�����N���X
'
'   Revision History:
'       Data        Description
'       2010/11/19  Double�^�̏���z��Double�^�ŕێ�����@�\����������
'
'�쐬��:
'   0145184346
'

Option Explicit

Private Const ERR_NUMBER As Long = 9999                  ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CParamDoubleArray" ' Class���̂�ێ�����

Private mDouble() As Double ' �z��Ő��l�f�[�^��ێ�����

Public Sub SetArrayAsDouble(ByRef dblArray() As Double)
'���e:
'   �z��Double�^�̃f�[�^��z��Double�^�̃f�[�^�Ŋi�[����
'
'�p�����[�^:
'   [dblArray]  In  �z��Double�^�f�[�^��ێ�����
'
'�߂�l:
'
'���ӎ���:
'

    '#####  �f�[�^���i�[����  #####
    ReDim mDouble(UBound(dblArray))
    mDouble = dblArray

End Sub

Public Sub SetAsDouble(ByVal dblValue As Double)
'���e:
'   Double�^�̃f�[�^��z��Double�^�̃f�[�^�Ŋi�[����
'
'�p�����[�^:
'   [dblValue]  In  Double�^�f�[�^��ێ�����
'
'�߂�l:
'
'���ӎ���:
'

    '#####  Site�������߂āA�z��Double�^�ɓW�J���i�[����  #####
    Dim i As Integer
    ReDim mDouble(TheExec.sites.ExistingCount - 1)
    For i = 0 To (TheExec.sites.ExistingCount - 1) Step 1
        mDouble(i) = dblValue
    Next i

End Sub

Public Property Get ReadAsDouble() As Double()
'���e:
'   �ێ����Ă���z��Double�^�̃f�[�^��߂�
'
'�p�����[�^:
'
'�߂�l:
'   �ێ����Ă���z��Double�^�̃f�[�^
'
'���ӎ���:
'

    '#####  �ێ����Ă���z��Double�^�̃f�[�^��߂�  #####
    ReadAsDouble = mDouble

End Property
