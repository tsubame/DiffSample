VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUtyBitController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ���[�e�B���e�BBit�̐ݒ�@�\�̒�
'
'�ړI:
'   APMU/CUB���[�e�B���e�BBit�ݒ�@�\�̒�
'
'�쐬��:
'   SLSI����
'
'���ӎ���:
'   APMU_UB�ACUB_UB�ɑΉ����܂��B
'   CUB UB�̐ݒ�@�\���p���ɂͤ�����t�R���p�C��������
'   CUB_UB_USE = 1�̋L�q���K�v�ł��B
'
'   ��������A�g�p�J�n�O�ɁAAPMU�ACUB�̃��[�e�B���e�BBit�ݒ��
'   �g�p����Utility Bit Setting���[�N�V�[�g�̎w�肪�K�v�ł�
'
'Code Checked
'
'

Option Explicit

'Tool�Ή���ɃR�����g�O���Ď��������ɂ���B�@2013/03/07 H.Arikawa
#Const CUB_UB_USE = 0    'CUB UB�̐ݒ�          0�F���g�p�A0�ȊO�F�g�p

Private m_ApmuUbSet As CApmuUBSet

#If CUB_UB_USE <> 0 Then
Private m_CubUbSet As CCubUBSet
#End If

'#Pass-Ok
Public Function AsAPMU() As CApmuUBSet
'���e:
'   APMU��UB�ݒ�@�\OBJ��I��
'
'�p�����[�^:
'
'�߂�l:
'   APMU��UB�ݒ�@�\Object
'
'���ӎ���:
'
    
    Set AsAPMU = m_ApmuUbSet

End Function

'#Pass-Ok
#If CUB_UB_USE <> 0 Then
Public Function AsCUB() As CCubUBSet
'���e:
'   CUB��UB�ݒ�@�\OBJ��I��
'
'�p�����[�^:
'
'�߂�l:
'   CUB��UB�ݒ�@�\Object
'
'���ӎ���:
'   ���p���ɂ́A�����t�R���p�C��������
'   CUB_UB_USE = 1�̋L�q���K�v�ł��B
'

    Set AsCUB = m_CubUbSet

End Function
#End If

'#Pass-Ok
Private Sub Class_Initialize()
    
    Me.Initialize

End Sub

'#Pass-Ok
Public Sub Initialize()
'���e:
'   UB�R���g���[���̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    
    Call InitApmuUb
    Call InitCubUb

End Sub

'#Pass-Ok
Private Sub InitApmuUb()
'APMU UB�ݒ�Object�̏�����
    
    Set m_ApmuUbSet = Nothing
    Set m_ApmuUbSet = New CApmuUBSet

End Sub

'#Pass-Ok
Private Sub InitCubUb()
'CUB UB�ݒ�Object�̏�����

    #If CUB_UB_USE <> 0 Then
    Set m_CubUbSet = Nothing
    Set m_CubUbSet = New CCubUBSet
    Set m_CubUbSet.CubUBObj = CUBUtilBit
    #End If

End Sub