VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcquireInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   AcquireInstance�̏���ێ������s����N���X
'
'   Revision History:
'       Data        Description
'       2010/06/11  AcquireInstance�����s����@�\����������
'       2010/06/22  AcquireInstance��CheckTermination�@�\����������
'                   ���s��񃊃Z�b�g�@�\����������
'       2010/07/02  Action��Count���Q�Ƃ���@�\����������
'       2011/02/10  Interceptor�AReadImage�ASaveImage�̎����B
'                   �����Acq�V�[�g�AEnable�`���\�b�h�V���[�Y�̂Q�ʂ�B�ݒ��Enable�`�V���[�Y���D�悳��܂��B
'
'�쐬��:
'   0145184346
'

Option Explicit

Private Const CLASS_NAME = "CAcquireInstance" ' Class���̂�ێ�����

Private mAcquireInstanceAction As Collection ' AcquireInstance����ێ�����

Private m_ActionLogger As CActionLogger

'Enable Flag ���L�����ǂ���
Private m_ValidEnableReadImage As Boolean
Private m_ValidEnableSaveImage As Boolean
Private m_ValidEnableShowImage As Boolean
'Private m_ValidEnableInterceptor As Boolean

'Enable Flag �̐ݒ�
Private m_FlagEnableReadImage As Boolean
Private m_FlagEnableSaveImage As Boolean
Private m_FlagEnableShowImage As Boolean
Private m_FlagEnableInterceptor As Boolean

Public Function Execute(ByRef strAcquireFrameWorkName As String) As Integer
'���e:
'   AcquireFrameWork�����s����
'
'�p�����[�^:
'   [strAcquireFrameWorkName]  In  AcquireFrameWork���̂�ێ�����
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    '#####  AcquireFrameWork���擾����  #####
    On Error GoTo VBAErrHandler
    Dim clsAcqAct As IAcquireAction

    Set clsAcqAct = mAcquireInstanceAction.Item(strAcquireFrameWorkName)
    
    
    '#####  AcquireFrameWork�����s����  #####
    On Error GoTo ErrHandler
'    Execute = clsAcqAct.Execute
    '���[�U�[�}�N������̕Ԃ�l��IG-XL�ɕԂ��Ȃ��B
    clsAcqAct.Execute

    Set clsAcqAct = Nothing


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
VBAErrHandler:
    Execute = TL_ERROR
    Call TheError.Raise(9999, CLASS_NAME & ".Execute", "Can Not Found Frame Work Instance : " & strAcquireFrameWorkName)
    Exit Function
ErrHandler:
    Execute = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function ToptStart(Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'���e:
'   AcquireFrameWork�C���X�^���X�̎��s���m�F����
'
'�p�����[�^:
'   [strAcquireFrameWorkName]  In  AcquireFrameWork���̂�ێ�����
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWork�C���X�^���X�̎��s���m�F����  #####
    Dim lngExeCnt As Integer ' AcquireFrameWork���s����ێ�����
    Dim clsAcqAct As IAcquireAction

    For Each clsAcqAct In mAcquireInstanceAction
        With clsAcqAct
            If strAcquireFrameWorkName = "" Or strAcquireFrameWorkName = .GetFrameWorkName Then
                lngExeCnt = lngExeCnt + 1
                Call .ToptStart
                ToptStart = .GetStatus
                If ToptStart = TL_ERROR Then
                    Exit Function
                End If
            End If
        End With
    Next clsAcqAct

    Set clsAcqAct = Nothing


    '#####  AcquireFrameWork�C���X�^���X�����s�L���̊m�F  #####
    If strAcquireFrameWorkName <> "" And lngExeCnt = 0 Then
        ToptStart = TL_ERROR
        Call TheError.Raise(9999, CLASS_NAME & ".ToptStart", "Can Not Execute Acquire Frame Work ! " _
                            & "You Have To Check AcquireFrameWork Name. ( " & strAcquireFrameWorkName & " )")
    End If


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    ToptStart = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function ClearStatus(Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'���e:
'   AcquireFrameWork�C���X�^���X�̎��s�������폜����
'
'�p�����[�^:
'   [strAcquireFrameWorkName]  In  AcquireFrameWork���̂�ێ�����
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWork�C���X�^���X�̎��s�������폜����  #####
    Dim lngExeCnt As Integer ' AcquireFrameWork���s����ێ�����
    Dim clsAcqAct As IAcquireAction

    For Each clsAcqAct In mAcquireInstanceAction
        With clsAcqAct
            If strAcquireFrameWorkName = "" Or strAcquireFrameWorkName = .GetFrameWorkName Then
                lngExeCnt = lngExeCnt + 1
                ClearStatus = .ClearStatus
                If ClearStatus = TL_ERROR Then
                    Exit Function
                End If
            End If
        End With
    Next clsAcqAct

    Set clsAcqAct = Nothing


    '#####  AcquireFrameWork�C���X�^���X�����s�L���̊m�F  #####
    If strAcquireFrameWorkName <> "" And lngExeCnt = 0 Then
        ClearStatus = TL_ERROR
        Call TheError.Raise(9999, CLASS_NAME & ".ClearStatus", "Can Not Execute Clear Status ! " _
                            & "You Have To Check AcquireFrameWork Name. ( " & strAcquireFrameWorkName & " )")
    End If


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    ClearStatus = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function CheckTermination(Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'���e:
'   AcquireFrameWork�C���X�^���X�̎��s�������m�F����
'
'�p�����[�^:
'   [strAcquireFrameWorkName]  In  AcquireFrameWork���̂�ێ�����
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWork�C���X�^���X�̎��s�������m�F����  #####
    Dim lngExeCnt As Integer ' AcquireFrameWork���s����ێ�����
    Dim clsAcqAct As IAcquireAction

    For Each clsAcqAct In mAcquireInstanceAction
        With clsAcqAct
            If strAcquireFrameWorkName = "" Or strAcquireFrameWorkName = .GetFrameWorkName Then
                lngExeCnt = lngExeCnt + 1
                If .GetStatus = -1 Then
                    CheckTermination = .ToptStart
                    If CheckTermination = TL_ERROR Then
                        Exit Function
                    End If
                End If
            End If
        End With
    Next clsAcqAct

    Set clsAcqAct = Nothing


    '#####  AcquireFrameWork�C���X�^���X�����s�L���̊m�F  #####
    If strAcquireFrameWorkName <> "" And lngExeCnt = 0 Then
        CheckTermination = TL_ERROR
        Call TheError.Raise(9999, CLASS_NAME & ".CheckTermination", "Can Not Execute CheckTermination ! " _
                            & "You Have To Check AcquireFrameWork Name. ( " & strAcquireFrameWorkName & " )")
    End If


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    CheckTermination = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Sub Reset()
'���e:
'   AcquireFrameWork�C���X�^���X�̎��s�������N���A����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'


    '#####  AcquireFrameWork�C���X�^���X�̎��s�������폜����  #####
    Dim clsAcqAct As IAcquireAction

    For Each clsAcqAct In mAcquireInstanceAction
        clsAcqAct.Reset
    Next clsAcqAct

    Set clsAcqAct = Nothing


    '#####  �I��  #####
    Exit Sub


End Sub

Public Property Let SetAction(ByRef IAcquireAction As IAcquireAction)
    If mAcquireInstanceAction Is Nothing Then
        Set mAcquireInstanceAction = New Collection
    End If
    Set IAcquireAction = InsertWrap(IAcquireAction, IAcquireAction.AcquireFunctionInfo)
    mAcquireInstanceAction.Add IAcquireAction, IAcquireAction.GetFrameWorkName
End Property

Public Property Get Count() As Long
    If Not mAcquireInstanceAction Is Nothing Then
        Count = mAcquireInstanceAction.Count
    Else
        Count = 0
    End If
End Property

Public Sub Initialize()
    Set mAcquireInstanceAction = Nothing
End Sub

Private Function ChangeConstructionObj()
'�v��
'�E�Q�Ƃ������ւ��iAcquireActionInstance�j
'�E�C���X�^���X�̏��Ԃ�����������ׂ邱�Ƃ��\�B���Ԃ������ĂȂ��Ɩ��Ӗ��ȓ���ƂȂ�B
'�@�Q�Ƃ̎�����
'�@Acquire��Interceptor��Dump��Fake������
'�E�Q�Ƃ�ύX���邱�ƂŎ��s�t���[���ύX�����
'
'���ӎ���
'Interceptor�͔C�ӂ̃^�C�~���O�ŃI�u�W�F�N�g�\����ύX�ł���K�v������B
'Dump�AFake�N���X�͂Ƃ肠�����AInit���ɍ\�z��z��B�C�ӂ̃^�C�~���O�ŕύX�\�ɂ��邩�ǂ����E�E�E
'�C���X�^���X�̃��X�g�\����CAcquireInstance���R���g���[������B
'�������ۂ����C���X�^���X�����邩�ۂ�
'�@Fake�AInterceptor
'�������ۂ����C���X�^���X�����邩�ۂ��{�t���O����(Show�Adump)
'�@dump
'
'���s���e
'�E���X�g�\���ɂȂ��Ă�C���X�^���X����A���̂����o���B
'�@���󂪂ǂ������\���ɂȂ��Ă��邩�͍l�����Ȃ��B
'�@���̃N���X�̌^�ɂ���āA���荞�ރN���X�ɐ������K�v�B
'�@�@�@�@�@�@�@�@�@�@�@Interceptor�@�@�@Dump�@�@�@�@�@Fake
'CSetAction�@�@�@�@�@�@�@�@���@�@�@�@�@�@�~�@�@�@�@�@�@�~
'CAcquireAction�@�@�@�@�@�@���@�@�@�@�@�@���@�@�@�@�@�@��(CAcquireFakeAction����)
'CPostAcquireAction�@�@�@�@���@�@�@�@�@�@���@�@�@�@�@�@��(CPostAcquireFakeAction����)

    Dim pFlagRead As Boolean
    Dim pFlagInterceptor As Boolean
    Dim pFlagSave As Boolean
    Dim pFlagShow As Boolean
    
    Dim FWmacro As IAcquireAction

    Dim pCollection As Collection
    Set pCollection = New Collection
    
    'CAcquireFunctionInfo�ApFlagInterceptor�̏��ɍ��킹�čč\��
    Dim pSubstance As IAcquireAction
    Dim pInfo As CAcquireFunctionInfo
    
    For Each FWmacro In mAcquireInstanceAction
        
        Set pSubstance = SearchSubstance(FWmacro)
        Set pInfo = pSubstance.AcquireFunctionInfo
        Set pSubstance = InsertWrap(pSubstance, pInfo)
        
        'Refer Collection�������Ă���Q�Ƃ��C���B
        Call pCollection.Add(pSubstance, pSubstance.GetFrameWorkName)
    
    Next FWmacro
    
    Set mAcquireInstanceAction = Nothing
    Set mAcquireInstanceAction = pCollection

End Function

Private Function InsertWrap(ByRef pSubstance As IAcquireAction, ByRef pInfo As CAcquireFunctionInfo) As IAcquireAction

    '���荞�ނ��ۂ�
    Dim pReadImage As Boolean
    Dim pSaveImage As Boolean
    Dim pShowImage As Boolean
    Dim pInterceptor As Boolean

    Dim pFakeObj As CAcquireFakeAction
    Dim pPostFakeObj As CPostAcquireFakeAction
    Dim pInterceptorObj As CAcquireActionInterceptor
    Dim pDumpObj As CAcquireDumpAction
    
    Dim pSetObj As IAcquireActionInstance       '�C���X�^���X�̃Z�b�^�[�A�Q�b�^�[�p

    Call FixFlag(pInfo, pReadImage, pSaveImage, pShowImage, pInterceptor)
    
    '##### �e���荞�݃C���X�^���X���� #####
    '### Fake ###
    'Fake(ACQ Sheet Arg5)
    If Not pInfo.macroName = "FWSetCondition" Then      'SetCondition�̎��͋L�q�Ɋ֌W�Ȃ�����
        If pReadImage = True Then
            If pInfo.macroName = "FWImageAcquire" Then
                Set pFakeObj = New CAcquireFakeAction
            ElseIf pInfo.macroName = "FWPostImageAcquire" Then
                Set pPostFakeObj = New CPostAcquireFakeAction
            Else
                Call TheError.Raise(9999, "CAcquireInstance", "FrameWorkName is different.")
            End If
        Else
            Set pFakeObj = Nothing
            Set pPostFakeObj = Nothing
        End If
    End If
    '### Dump ###
    'Dump(ACQ Sheet Arg4 or Arg7)
    If Not pInfo.macroName = "FWSetCondition" Then      'SetCondition�̎��͋L�q�Ɋ֌W�Ȃ�����
        If pSaveImage = True Or pShowImage = True Then
            Set pDumpObj = New CAcquireDumpAction
            pDumpObj.FlagSaveImage = pSaveImage
            pDumpObj.FlagShow = pShowImage
        Else
            Set pDumpObj = Nothing
        End If
    End If
    
    '### Interceptor ###
    If pInterceptor = True Then
        Set pInterceptorObj = New CAcquireActionInterceptor
        Set pInterceptorObj.actionLogger = m_ActionLogger
    Else
        Set pInterceptorObj = Nothing
    End If
        
        
    '##### ���X�g�\���𐶐� #####
    'Fake(ACQ Sheet Arg5)
    If Not pFakeObj Is Nothing Then
        Set pSetObj = pFakeObj                              '�L���X�g
        Set pSetObj.AcquireActionInstance = pSubstance      '���̎Q�Ƃ��Z�b�g
        Set pSubstance = pSetObj                            '
    End If
    'PostFake(ACQ Sheet Arg5)
    If Not pPostFakeObj Is Nothing Then
        Set pSetObj = pPostFakeObj                              '�L���X�g
        Set pSetObj.AcquireActionInstance = pSubstance      '���̎Q�Ƃ��Z�b�g
        Set pSubstance = pSetObj                            '
    End If
    'Dump(ACQ Sheet Arg4 or Arg7)
    If Not pDumpObj Is Nothing Then
        Set pSetObj = pDumpObj                              '�L���X�g
        Set pSetObj.AcquireActionInstance = pSubstance      '���̎Q�Ƃ��Z�b�g
        Set pSubstance = pSetObj                            '
    End If
    'Interceptor(pFlagInterceptor)
    If Not pInterceptorObj Is Nothing Then
        Set pSetObj = pInterceptorObj                       '�L���X�g
        Set pSetObj.AcquireActionInstance = pSubstance      '���̎Q�Ƃ��Z�b�g
        Set pSubstance = pSetObj                            '
    End If
    
    Set InsertWrap = pSubstance

End Function

Private Sub FixFlag(pInfo As CAcquireFunctionInfo, ByRef pReadImage As Boolean, ByRef pSaveImage As Boolean, ByRef pShowImage As Boolean, ByRef pInterceptor As Boolean)
'�N���X�����荞�ނ��ǂ��������肷��
        
    Dim pAfiArg4 As String
    Dim pAfiArg5 As String
    Dim pAfiArg6 As String
    Dim pAfiArg7 As String
    Dim pAfiArg8 As String

    'pInfo.Arg(4):Show
    'pInfo.Arg(5):Read
    'pInfo.Arg(6):FilePath
    'pInfo.Arg(7):Write
    'pInfo.Arg(8):FilePath
    
    '�召������ʂ��Ȃ�����UCase
    pAfiArg4 = UCase(pInfo.Arg(4))
    pAfiArg5 = UCase(pInfo.Arg(5))
    pAfiArg6 = pInfo.Arg(6)
    pAfiArg7 = UCase(pInfo.Arg(7))
    pAfiArg8 = pInfo.Arg(8)
    
    '�V�[�g���A���W���[���t���O���܂߂čŏI�I�Ɋ��荞�ރN���X�����肷��B
    '### Fake ###
    If m_ValidEnableReadImage = True Then
        pReadImage = m_FlagEnableReadImage
    Else
        If pAfiArg5 = "READ" Then
            pReadImage = True
        Else
            pReadImage = False
        End If
    End If
    
    '### Dump ###
    If m_ValidEnableShowImage = True Then
        pShowImage = m_FlagEnableShowImage
    Else
        If pAfiArg4 = "SHOW" Then
            pShowImage = True
        Else
            pShowImage = False
        End If
    End If
        
    If m_ValidEnableSaveImage = True Then
        pSaveImage = m_FlagEnableSaveImage
    Else
        If pAfiArg7 = "WRITE" Then
            pSaveImage = True
        Else
            pSaveImage = False
        End If
    End If
    
    '### Interceptor ###
    If m_FlagEnableInterceptor = True Then
        pInterceptor = True
    Else
        pInterceptor = False
    End If

End Sub

Public Sub EnableOffImageDebugger()
    m_ValidEnableReadImage = False
    m_ValidEnableSaveImage = False
    m_ValidEnableShowImage = False
    Call ChangeConstructionObj
End Sub

Public Sub EnableReadImage(ByVal pEnableFlag As Boolean, ByVal pPathName As String)
    m_ValidEnableReadImage = True
    m_FlagEnableReadImage = pEnableFlag
    Call ChangeConstructionObj
End Sub

Public Sub EnableSaveImage(ByVal pEnableFlag As Boolean, ByVal pPathName As String)
    m_ValidEnableSaveImage = True
    m_FlagEnableSaveImage = pEnableFlag
    Call ChangeConstructionObj
End Sub

Public Sub EnableShowImage(ByVal pEnableFlag As Boolean)
    m_ValidEnableShowImage = True
    m_FlagEnableShowImage = pEnableFlag
    Call ChangeConstructionObj
End Sub

Public Sub EnableInterceptor(ByVal pEnableFlag As Boolean, ByRef pLogger As CActionLogger)
    m_FlagEnableInterceptor = pEnableFlag
    Set m_ActionLogger = pLogger
    Call ChangeConstructionObj
End Sub

Private Function SearchSubstance(ByRef pObj As IAcquireAction) As IAcquireAction
'Set pSetObj = pIAA�����s����Ǝ��̂Ȃ̂ŁApIAA��Ԃ��B

    Dim pSetObj As IAcquireActionInstance       '�C���X�^���X�̃Z�b�^�[�A�Q�b�^�[�p
    Dim pIAA As IAcquireAction
    
    Set pIAA = pObj                             '�Ƃ肠�����A���[�J���̃I�u�W�F�N�g�ϐ��ɓn��
    
    On Error GoTo NON_WRAP
    
    Do
        '�����Ă���IAcquireAction�C���X�^���X�������Ȃ�܂ŌJ��Ԃ�Set�B
        Set pSetObj = pIAA                              '����Set�Ŏ��s������pIAA������
        Set pIAA = pSetObj.AcquireActionInstance        '����Set�͐�������
    Loop
    
NON_WRAP:
    Set SearchSubstance = pIAA

End Function

Private Sub Class_Initialize()
    Set mAcquireInstanceAction = Nothing
End Sub

Private Sub Class_Terminate()
    Set mAcquireInstanceAction = Nothing
End Sub