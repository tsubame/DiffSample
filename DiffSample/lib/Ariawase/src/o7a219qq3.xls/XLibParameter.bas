Attribute VB_Name = "XLibParameter"
'�T�v:
'   �p�����[�^�N���X�p�I�u�W�F�N�g�쐬���W���[��
'
'�ړI:
'   �p�����[�^�I�N���X�u�W�F�N�g���쐬���Ԃ��i���C�u�����쐬���[���Ɋ�Â��j
'
'�쐬��:
'   0145206097
'
Option Explicit







Public Function CreateCParamDouble() As CParamDouble
    Set CreateCParamDouble = New CParamDouble
End Function

Public Function CreateCParamLong() As CParamLong
    Set CreateCParamLong = New CParamLong
End Function

Public Function CreateCParamBoolean() As CParamBoolean
    Set CreateCParamBoolean = New CParamBoolean
End Function

Public Function CreateCParamName() As CParamName
    Set CreateCParamName = New CParamName
End Function

Public Function CreateCParamString() As CParamString
    Set CreateCParamString = New CParamString
End Function

Public Function CreateCParamStringWithUnit() As CParamStringWithUnit
    Set CreateCParamStringWithUnit = New CParamStringWithUnit
End Function
