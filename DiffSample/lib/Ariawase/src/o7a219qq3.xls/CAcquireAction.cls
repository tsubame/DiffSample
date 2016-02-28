VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcquireAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   AcquireFrameWork�iFWImageAcquire�j�̏���ێ������s����N���X
'
'   Revision History:
'       Data        Description
'       2010/04/28  AcquireFrameWork�iFWImageAcquire�j�����s����@�\����������
'       2010/06/11  �f�[�^�\���������̂��߁A�v���O�����R�[�h��ύX����
'       2010/06/22  ���s�O��̏���ǉ������im_Status��ύX�j
'                   ���s��񃊃Z�b�g�@�\����������
'                   Macro���s�G���[������ǉ�����
'
'�쐬��:
'   tomoyoshi.takase
'   0145184346
'

Option Explicit

Implements IAcquireAction

Private Const ERR_NUMBER = 9999             ' �G���[�ԍ���ێ�����
Private Const ERR_NUMBER_NOTFOUND_MACRO = 424
Private Const CLASS_NAME = "CAcquireAction" ' Class���̂�ێ�����

Private m_FunctionInfo As CAcquireFunctionInfo ' ImageACQTBL�V�[�g1�s�̃f�[�^��ێ�����
'Private m_ActionLogger As CActionLogger        ' ���sLog�f�[�^���_���v����ActionLogger��ێ�����
Private m_InstanceName As String               ' TestInstance���̂�ێ�����
Private m_ActionName As String                 ' AcquireFrameWork�����s����UserMacro���̂�ێ�����
Private m_Status As Long                       ' ���s���ʂ�ێ�����i-1�F���s�O�@0�F����I���@1�F�G���[�����j

Public Property Get AsIAcquireAction() As IAcquireAction
    Set AsIAcquireAction = Me
End Property

Private Function IAcquireAction_Execute() As Integer
'���e:
'   ImageAcquire��UserMacro�����s����
'
'�p�����[�^:
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'
'   Arg(0):�C���X�^���X��
'   Arg(1):FW��
'   Arg(2):�]����v���[���O���[�v
'   Arg(3):�]����v���[��PMD��
'   Arg(4):BitDepth
'   �ȍ~���[�U�̈�

    On Error GoTo VBAErrHandler

    Dim capPlane As CImgPlane
    Set capPlane = GetFreePlaneForAcq(m_FunctionInfo.Arg(2), ConvertStrToDepth(m_FunctionInfo.Arg(3)), , m_FunctionInfo.Arg(0) & " Capture Plane")

    Dim devInfo As CDeviceConfigInfo
    Set devInfo = TheDeviceProfiler.ConfigInfo(m_FunctionInfo.Arg(0))
    
    Dim capParam As CParamPlane
    Set capParam = New CParamPlane
    Set capParam.plane = capPlane
    Set capParam.DeviceConfigInfo = devInfo
    
    m_Status = Application.Run(m_ActionName, m_FunctionInfo, capParam)

    'GetFreePlane�������B
    If Not capPlane Is Nothing Then
        Call TheParameterBank.Add(capParam, m_FunctionInfo.Arg(0))
    End If

    On Error GoTo ErrHandler

    '���[�U�[�}�N������̕Ԃ�l��TL_SUCCESS�ɕύX����̂�CAcquirenstance
    IAcquireAction_Execute = m_Status
    
    If m_Status = TL_ERROR Then
        Call TheError.Raise(ERR_NUMBER, m_ActionName, "User Macro returned Error!")
    ElseIf m_Status <> TL_ERROR Then
        m_Status = TL_SUCCESS
    End If
    
    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
VBAErrHandler:
    m_Status = TL_ERROR
    If Err.Number = ERR_NUMBER_NOTFOUND_MACRO Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Execute", _
                        "Can Not Find Acquire Macro." & "(" & m_ActionName & ")")
    Else
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Execute", _
                        "[" & Err.Number & ":" & Err.Source & "]" & Err.Description & "(" & m_ActionName & ")")
    End If
    Exit Function
ErrHandler:
    m_Status = TL_ERROR
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Execute", "Image Acquire Error : " & m_InstanceName)
    Exit Function

End Function

Private Function ConvertStrToDepth(ByVal pStr As String) As IdpBitDepth
    If pStr Like "*S16" Then
        ConvertStrToDepth = idpDepthS16
    ElseIf pStr Like "*S32" Then
        ConvertStrToDepth = idpDepthS32
    ElseIf pStr Like "*F32" Then
        ConvertStrToDepth = idpDepthF32
    Else
        Err.Raise 9999, TypeName(Me), "Illegal IdpBitDepth: " & pStr
    End If
End Function

Private Function GetFreePlaneForAcq(ByVal pType As String, ByVal pBitDepth As IdpBitDepth, _
        Optional ByVal pClear As Boolean = False, Optional pComment As String) As CImgPlane
    
    Dim ret As CImgPlane
    On Error GoTo ErrExit
    Set ret = TheIDP.PlaneManager(pType).GetFreePlane(pBitDepth, pClear)
    ret.Comment = pComment
    
    Set GetFreePlaneForAcq = ret
    Exit Function
    
ErrExit:
    Dim Err As CErrInfo
    Set Err = TheError.LastError
    TheExec.Datalog.WriteComment pType & " : There is no free plane."
    TheExec.Datalog.WriteComment Err.Message
    Set GetFreePlaneForAcq = ret
    Exit Function

End Function

Private Function IAcquireAction_ToptStart() As Integer
'���e:
'   TestOptimizer�@�\���g���āAAcquireFrameWork�����s����
'
'�p�����[�^:
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    '#####  AcquireFrameWork�����s����  #####
    With m_FunctionInfo
        If .macroName <> "" Then
            On Error GoTo IGXLErrHandler
            Call TheHdw.TOPT.Start(.GroupName, .macroName, .InstanceName)
        Else
            On Error GoTo ErrHandler
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".ToptStart", "Can Not Found Acquire Frame Work Name ! ( " _
                                & .GroupName & " / " & .InstanceName & " )")
        End If
    End With


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
IGXLErrHandler:
    IAcquireAction_ToptStart = TL_ERROR
    m_Status = IAcquireAction_ToptStart
    Exit Function
ErrHandler:
    IAcquireAction_ToptStart = TL_ERROR
    m_Status = IAcquireAction_ToptStart
    Exit Function


End Function

Private Function IAcquireAction_ClearStatus() As Integer
'���e:
'   AcquireFrameWork�C���X�^���X�̎��s�������폜����
'
'�p�����[�^:
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWork�C���X�^���X�̎��s�������폜����  #####
    Call TheHdw.TOPT.ClearStatus(m_FunctionInfo.GroupName)


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    IAcquireAction_ClearStatus = TL_ERROR
    Exit Function


End Function

Private Property Get IAcquireAction_GetFrameWorkName() As String
    IAcquireAction_GetFrameWorkName = m_FunctionInfo.macroName
End Property

Private Property Get IAcquireAction_GetStatus() As Integer
    IAcquireAction_GetStatus = m_Status
End Property

Private Sub IAcquireAction_Reset()
    m_Status = -1
End Sub

Private Property Get IAcquireAction_AcquireFunctionInfo() As CAcquireFunctionInfo
    Set IAcquireAction_AcquireFunctionInfo = m_FunctionInfo
End Property

Private Sub IAcquireAction_Initialize(reader As IParameterReader)
    Set m_FunctionInfo = New CAcquireFunctionInfo
    Call m_FunctionInfo.Initialize(reader)
    m_InstanceName = reader.ReadAsString("Arg0@Parameters")
    m_ActionName = reader.ReadAsString("Arg1@Parameters")
    m_Status = -1
End Sub

Private Sub Class_Initialize()
    Set m_FunctionInfo = Nothing
    m_InstanceName = ""
    m_ActionName = ""
    m_Status = -1
End Sub

Private Sub Class_Terminate()
    Set m_FunctionInfo = Nothing
    m_InstanceName = ""
    m_ActionName = ""
    m_Status = -1
End Sub

