VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLimitTypeLow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �X���C�X���x����Include����������
'
'�ړI:
'   Include�����̂��߁A������-1���ĕԂ�
'
'�쐬��:
'   0145184306
'
Option Explicit

Implements ILimitTypeOperator

Private Const ERR_NUMBER As Long = 9999                  ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CLimitTypeLow"       ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default�l"Empty"��ێ�����

Public Function AsIlimitTypeOperator() As ILimitTypeOperator
    Set AsIlimitTypeOperator = Me
End Function

Public Function IlimitTypeOperator_Include(ByVal pLevel As Long) As Long
'���e:
'   Include����������
'
'�߂�l:
'   pLevel - 1
'
'���l:
'
    IlimitTypeOperator_Include = pLevel - 1
End Function


