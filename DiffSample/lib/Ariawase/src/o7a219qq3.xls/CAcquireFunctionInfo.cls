VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcquireFunctionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   Image ACQTBL �V�[�g�̏��i�[�N���X
'�ړI:
'   Image Acquire Table �V�[�g�ɂ���������i�[���AAction�N���X�Ɋ֘A�t���A���[�U�[�}�N������A�N�Z�X�\�ɂ���B
'   ���[�U�[�}�N���ɂ��̃C���X�^���X��n�����ƂŁA���[�U�[��Property Get�ŃV�[�g�̏����擾�\�B
'   �C�j�V�������̂ݏ������݉\�B���̑��͕s�i�V�[�g�̏��Ƃ̕s���������邽�߁j
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N3��11��
'���ӎ���:
'   ImageACQTBL �V�[�g�� Instance Name �͕K�����j�[�N�łȂ���΂Ȃ�Ȃ��B����Instance Name�Ƃ̏d���͂ł��܂���B
'
Option Explicit

Private Const ERR_NUMBER = 9999                         '�G���[���ɓn���G���[�ԍ�
Private Const CLASS_NAME = "CAcquireFunctionInfo"       '���̃N���X�̖��O

Private Const m_MAX_ARG_NO As Long = 79                 'ImageACQTBL�V�[�g�̍ő�Arg��0~79

Private m_GroupName As String                           'ImageACQTBL�V�[�g��GroupName
Private m_MacroName As String                           'ImageACQTBL�V�[�g��MacroName
Private m_AutoAcquire As String                         'ImageACQTBL�V�[�g��AutoAcquire
Private m_InstanceName As String                        'ImageACQTBL�V�[�g��InstanceName
Private m_Arg(m_MAX_ARG_NO) As String                   'ImageACQTBL�V�[�g��Arg�Q
Private m_ArgParameterMaxIndex As Integer               'ImageACQTBL�V�[�g��Arg�̍Ō�̐�
Private m_Comment As String                             'ImageACQTBL�V�[�g��Comment

Property Get GroupName() As String
'���e:
'
'�p�����[�^:
'
'�߂�l:
'   Group Name �ɑΉ����郉�x����
'
'���ӎ���:
'
    GroupName = m_GroupName
End Property
Property Get macroName() As String
    macroName = m_MacroName
End Property
Property Get AutoAcquire() As String
    AutoAcquire = m_AutoAcquire
End Property
Property Get InstanceName() As String
    InstanceName = m_InstanceName
End Property
Property Get Arg(ByVal index As Long) As Variant
'���e:
'   ImageACQTBL�V�[�g��Arg0-79�̏���Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N3��15��
'�p�����[�^:
'   [index] In  1):Arg�ɑΉ����鐔
'�߂�l:
'   Variant
'
'���ӎ���:
'
    If index < 0 Or index > m_ArgParameterMaxIndex Then
        Arg = Empty
    Else
        Arg = m_Arg(index)
    End If
End Property
Property Get Comment() As String
    Comment = m_Comment
End Property

Property Get ArgParameterCount() As Long
    ArgParameterCount = m_ArgParameterMaxIndex + 1
End Property


Property Get ToCSVString() As String
'���e:
'   ���̃C���X�^���X�����SARG����CSV�ŕԂ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   String
'���ӎ���:
'
    Dim str As String
    Dim i  As Integer

    For i = 0 To m_ArgParameterMaxIndex
        str = str & m_Arg(i) & ","
    Next i
    
    ToCSVString = str

End Property

Public Function Clone() As CAcquireFunctionInfo
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��25��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   CAcquireFunctionInfo
'
'���ӎ���:
'

    Dim pClone As CAcquireFunctionInfo
    Set pClone = New CAcquireFunctionInfo
    Call pClone.CopyObj(Me)
    Set Clone = pClone

End Function

Public Sub CopyObj(OrgObj As CAcquireFunctionInfo)
'���e:
'   OrgObj �̓��e�����̃C���X�^���X�̓��e�ɔ��f���܂��B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [OrgObj]    In  1):���e���擾������ CAcquireFunctionInfo
'�߂�l:
'
'���ӎ���:
'
    Dim i As Integer
    
    If m_GroupName = "" And m_MacroName = "" And m_AutoAcquire = "" Then
        With OrgObj
            m_GroupName = .GroupName                              'OrgObj��GroupName
            m_MacroName = .macroName                              'OrgObj��MacroName
            m_AutoAcquire = .AutoAcquire                          'OrgObj��AutoAcquire
            m_InstanceName = .InstanceName                        'OrgObj��InstanceName
            For i = 0 To m_MAX_ARG_NO
                If .Arg(i) <> "" Then
                    m_Arg(i) = .Arg(i)                             'OrgObj��Arg�Q
                    m_ArgParameterMaxIndex = i
                End If
            Next i
            m_Comment = .Comment                                   'OrgObj��Comment
        End With
    Else
        Call TheError.Raise(9999, TypeName(Me), "CAcquireFunctionInfo can write at once.")
    End If

End Sub

Public Sub Initialize(reader As IParameterReader)
'���e:
'   ImageACQTBL�V�[�g�̏����i�[���܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N3��15��
'�p�����[�^:
'   [Reader]    In/Out  1):ImageACQTBL�̃p�����[�^���[�_�[�C���^�[�t�F�[�X
'�߂�l:
'
'���ӎ���:
'

    Dim i As Integer
    
    With reader
        m_GroupName = .ReadAsString("Group Name")                              'ImageACQTBL�V�[�g��GroupName
        m_MacroName = .ReadAsString("Macro Name")                              'ImageACQTBL�V�[�g��MacroName
        m_AutoAcquire = .ReadAsString("Auto Acquire")                          'ImageACQTBL�V�[�g��AutoAcquire
        m_InstanceName = .ReadAsString("Instance Name")                        'ImageACQTBL�V�[�g��InstanceName
        For i = 0 To m_MAX_ARG_NO
            If .ReadAsString("Arg" & i & "@Parameters") <> "" Then
                m_Arg(i) = .ReadAsString("Arg" & i & "@Parameters")                          'ImageACQTBL�V�[�g��Arg�Q
                m_ArgParameterMaxIndex = i
            End If
        Next i
        m_Comment = .ReadAsString("Comment")                                 'ImageACQTBL�V�[�g��Comment
    End With

End Sub

Private Sub Class_Initialize()
'���e:
'�@���̃N���X�̃I�u�W�F�N�g�������̏����B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:

m_GroupName = ""                            'ImageACQTBL�V�[�g��GroupName
m_MacroName = ""                            'ImageACQTBL�V�[�g��MacroName
m_AutoAcquire = ""                          'ImageACQTBL�V�[�g��AutoAcquire
m_InstanceName = ""                         'ImageACQTBL�V�[�g��InstanceName
m_Comment = ""                              'ImageACQTBL�V�[�g��Comment
Erase m_Arg                                 'ImageACQTBL�V�[�g��Arg�Q
m_ArgParameterMaxIndex = 0                  'ImageACQTBL�V�[�g��Arg�̍Ō�̐�

End Sub

Private Sub Class_Terminate()
'���e:
'�@���̃N���X�̃I�u�W�F�N�g�j�����̏����B
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:

m_GroupName = ""                            'ImageACQTBL�V�[�g��GroupName
m_MacroName = ""                            'ImageACQTBL�V�[�g��MacroName
m_AutoAcquire = ""                          'ImageACQTBL�V�[�g��AutoAcquire
m_InstanceName = ""                         'ImageACQTBL�V�[�g��InstanceName
m_Comment = ""                              'ImageACQTBL�V�[�g��Comment
Erase m_Arg                                 'ImageACQTBL�V�[�g��Arg�Q
m_ArgParameterMaxIndex = 0                  'ImageACQTBL�V�[�g��Arg�̍Ō�̐�

End Sub

