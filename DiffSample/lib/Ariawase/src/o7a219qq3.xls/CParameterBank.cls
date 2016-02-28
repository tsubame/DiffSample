VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParameterBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �p�����[�^�u��B
'
'�ړI:
'   �p�����[�^�������Ă��閼�O�t���œo�^�����B
'   �قȂ�X�R�[�v���ƂɌʂ̃C���X�^���X�𐶐�����
'
'�쐬��:
'   0145184304
'

Option Explicit

Implements IParameterBank

Private Const ERR_NUMBER As Long = 9999                     ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CParameterBank" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty        ' Default�l"Empty"��ێ�����

'=======
' paramName��Key�ɂ���IProcParameter�������
'=======
Private m_Keys As Collection          ' �L�[��Collection
Private m_FixedKeys As Collection          ' �L�[��Collection
Private m_Parameters As Collection          ' �v�f��Collection
Private m_FixedParameters As Collection     ' Fix���ꂽ�v�f��Collection
Private m_OverWriteMode As Boolean          ' �㏑���[�h

Private Sub IParameterBank_Clear()
'���e:
'   m_Parameters��S���N���A
'
'���l:
'
    Set m_Parameters = New Collection
    Set m_Keys = New Collection
    
End Sub

Private Function IParameterBank_Count() As Long
'���e:
'   m_Parameters��Count����������
'
'�߂�l:
'    m_Parameters��Count��
'
'���l:
'

    IParameterBank_Count = m_Parameters.Count + m_FixedParameters.Count

End Function

Private Sub IParameterBank_Delete(ByVal paramName As String)
'���e:
'   m_ParamSet����w���paramName���폜
'
'[paramName]   IN String�^:     �폜����paramName���w��
'
'���l:
'
    
    If Not GetParam(paramName, m_Parameters) Is Nothing Then
        ' NormalParameters�ɂ�������폜
        Call m_Parameters.Remove(paramName)
        Call m_Keys.Remove(paramName)
    ElseIf Not GetParam(paramName, m_FixedParameters) Is Nothing Then
        ' FixedParameters�ɂ�������폜
        Call m_FixedParameters.Remove(paramName)
        Call m_FixedKeys.Remove(paramName)
    Else
        ' Normal/Fixed�Ƃ��ɂȂ���΃G���[
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Delete", """" & paramName & """ doesn't exist.")
    End If
End Sub

Private Function IParameterBank_FixedKeys() As Collection
    Set IParameterBank_FixedKeys = m_FixedKeys
End Function

Private Function IParameterBank_Item(ByVal paramName As String) As IProcParameter
'���e:
'   m_Parameters���w���paramName�Ɉ�v����IProcParameter��Ԃ�
'
'[paramName]   IN String�^:     Get����paramName���w��
'
'�߂�l:
'    �w���paramName�Ɉ�v����IProcParameter
'
'���l:
'

    ' NormalParameters����擾
    If Me.AsIParameterBank.IsExist(paramName) Then
        Set IParameterBank_Item = GetParam(paramName, m_Parameters)
        If IParameterBank_Item Is Nothing Then
            ' NormalParameters�ɂȂ����FixedParameters����擾
            Set IParameterBank_Item = GetParam(paramName, m_FixedParameters)
        End If
    Else
        ' Normal/Fixed�Ƃ��ɂȂ���΃G���[
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Item", """" & paramName & """ doesn't exist.")
    End If

End Function

Private Function IParameterBank_Keys() As Collection
    Set IParameterBank_Keys = m_Keys
End Function

Private Function IParameterBank_Move(ByVal paramName As String) As IProcParameter
'���e:
'   m_Parameters���w���paramName�Ɉ�v����IProcParameter��Ԃ�
'   �Ԃ���Parameter��m_Parameters���폜����
'
'[paramName]   IN String�^:     Get����paramName���w��
'
'�߂�l:
'    �w���paramName�Ɉ�v����IProcParameter
'
'���l:
'

    ' m_Parameters���猟�����Ď擾
    Set IParameterBank_Move = GetParam(paramName, m_Parameters)
    
    If IParameterBank_Move Is Nothing Then
        ' m_Parameters�ɂȂ������ꍇ�Am_FixedParameters���������Ď擾
        Set IParameterBank_Move = GetParam(paramName, m_FixedParameters)
    Else
        ' m_Parameters�ɂ������ꍇ�Am_Parameters����폜
        Call m_Parameters.Remove(paramName)
        Call m_Keys.Remove(paramName)
        Exit Function
    End If

    '�G���[
    If IParameterBank_Move Is Nothing Then
        ' m_FixedParameters�ɂȂ������ꍇ�A�G���[
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Item", """" & paramName & """ doesn't exist.")
    Else
        ' m_FixedParameters�ɂ������ꍇ�Am_FixedParameters����폜
        Call m_FixedParameters.Remove(paramName)
        Call m_FixedKeys.Remove(paramName)
    End If
    
End Function

Private Sub IParameterBank_Add(ByRef paramObj As IProcParameter, ByVal paramName As String, Optional ByVal pFixedPlane As Boolean = False)
'���e:
'   m_Parameters�Ɏw���paramName��Key��IProcParameter��Add����B
'
'[paramObj]   IN IProcParameter�^:     Set����IPocParameter���w��
'[pFixedPlane] IN   Boolean�^:      �Œ�v���[���I�v�V����
'
'���l:
'

    If IParameterBank_IsOverwriteMode = False Then
        ' �㏑���֎~�̎��A���łɑ��݂��Ă��邩�m�F
        If IParameterBank_IsExist(paramName) = True Then
            ' �㏑���֎~�ł��łɑ��݂��Ă����ꍇ�̓G���[
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Add", paramName & " already set")
            Exit Sub
        End If
    Else
        ' �㏑���\�̏ꍇ�͑��݂��Ă���ꍇ�͈�x����
        If IsExistParam(paramName, m_FixedParameters) = True Then
            Call m_FixedParameters.Remove(paramName)
            Call m_FixedKeys.Remove(paramName)
        ElseIf IsExistParam(paramName, m_Parameters) = True Then
            Call m_Parameters.Remove(paramName)
            Call m_Keys.Remove(paramName)
        End If
        '�o�^������Ή������Ȃ�
    End If

    If pFixedPlane = True Then
        'FixedParameters�ɓo�^
        Call m_FixedParameters.Add(paramObj, paramName)     'Key��param�̖��O
        Call m_FixedKeys.Add(paramName, paramName)
    Else
        'NormalParameters�ɓo�^
        Call m_Parameters.Add(paramObj, paramName)
        Call m_Keys.Add(paramName, paramName)
    End If

End Sub
Private Function IParameterBank_List() As String
'���e:
'   �o�^���ꂽ�p�����[�^�̃��X�g���擾
'
'���l:
'   �J���}��؂�̕�����ŏo�́B
'   �o�^��,�v���[����,�c�@�̏�
'
'�o�̓T���v��
' [pParam1(Fixed),CParamPlane,pParam2,CParamPlane,]

    Dim tmp As Variant
    For Each tmp In m_FixedKeys
'        IParameterBank_List = IParameterBank_List & tmp.Name & "(Fixed)" & "," & TypeName(tmp) & ","
        IParameterBank_List = IParameterBank_List & tmp & "(Fixed)" & "," & TypeName(m_FixedParameters.Item(tmp)) & ","
    Next tmp
    
    For Each tmp In m_Keys
'        IParameterBank_List = IParameterBank_List & tmp.Name & "," & TypeName(tmp) & ","
        IParameterBank_List = IParameterBank_List & tmp & "," & TypeName(m_Parameters.Item(tmp)) & ","
    Next tmp
    
End Function
Private Property Get IParameterBank_IsOverwriteMode() As Boolean
    IParameterBank_IsOverwriteMode = m_OverWriteMode
End Property

Private Property Let IParameterBank_IsOverwriteMode(ByRef mode As Boolean)
    m_OverWriteMode = mode
End Property

Private Function IParameterBank_IsExist(ByVal paramName As String) As Boolean
'���e:
'   m_Parameters���w���paramName�Ɉ�v����IProcParameter���������ꍇTrue��Ԃ�
'
'[paramName]   IN String�^:     ����/�Ȃ����m�F����paramName���w��
'
'�߂�l:
'    �w���paramName�Ɉ�v����IProcParameter���������ꍇ True
'    �w���paramName�Ɉ�v����IProcParameter���Ȃ������ꍇ False
'
'���l:
'
    
    If IsExistParam(paramName, m_Parameters) Then
        IParameterBank_IsExist = True
        Exit Function
    ElseIf IsExistParam(paramName, m_FixedParameters) Then
        IParameterBank_IsExist = True
        Exit Function
    End If
    
    IParameterBank_IsExist = False

End Function

Private Function IsExistParam(ByVal paramName As String, ByRef pParam As Collection) As Boolean
'���e:
'   �w��̃R���N�V�������w���paramName�Ɉ�v����Key���������ꍇTrue��Ԃ�
'
'[paramName]   IN String�^:     ����/�Ȃ����m�F����paramName���w��
'[pParam]      IN Collection�^: �m�F�Ώۂ�Collection
'
'�߂�l:
'    �w���paramName�Ɉ�v����Key���������ꍇ True
'    �w���paramName�Ɉ�v����Key���Ȃ������ꍇ False
'
'���l:
'

    Dim p_Param As IProcParameter
    
    On Error GoTo NOT_FOUND_PARAM
    
    Set p_Param = pParam.Item(paramName)
    IsExistParam = True
    
    Set p_Param = Nothing
    Exit Function
    
NOT_FOUND_PARAM:
    
    IsExistParam = False

End Function

Private Function GetParam(ByVal pName As String, ByRef pParam As Collection) As IProcParameter
'���e:
'   �w��̃R���N�V�������w���pName�Ɉ�v����Key���������ꍇ�A�A�C�e��(IProcParameter)��Ԃ�
'
'[paramName]   IN String�^:     ����/�Ȃ����m�F����paramName���w��
'[pParam]      IN Collection�^: �m�F�Ώۂ�Collection
'
'�߂�l:
'    �w���paramName�Ɉ�v����A�C�e�� IProcParameter�^
'
'���l:
'

On Error GoTo NOT_FOUND
    Set GetParam = pParam.Item(pName)
    GetParam.Name = pName
    Exit Function

NOT_FOUND:
    Set GetParam = Nothing
    
End Function

Private Sub Class_Initialize()
    Set m_Keys = New Collection
    Set m_FixedKeys = New Collection
    Set m_Parameters = New Collection
    Set m_FixedParameters = New Collection
    m_OverWriteMode = False
End Sub

Private Sub Class_Terminate()

    Set m_Parameters = Nothing
    Set m_Keys = Nothing
    Set m_FixedKeys = Nothing
    Set m_FixedParameters = Nothing
    m_OverWriteMode = False
    
End Sub

Public Function AsIParameterBank() As IParameterBank
    Set AsIParameterBank = Me
End Function
