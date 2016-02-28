VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImageEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �B���e�X�g�C���X�^���X�𐶐�����N���X
'
'   Revision History:
'       Data        Description
'       2010/04/28  �B���e�X�g�C���X�^���X�𐶐������s����@�\����������
'       2010/05/12  �v���O�����R�[�h�𐮗�����
'       2010/05/19  SheetReader�Q�ƃL�[���[�h���Œ胉�x���ɕύX����
'       2010/05/25  SheetReader���O������󂯎��悤�ɕύX����
'       2010/05/31  Error������ύX����
'       2010/06/11  �v���O�����R�[�h�𐮗�����
'
'�쐬��:
'   0145184346
'

Option Explicit

Private Const CLASS_NAME  As String = "CImageEngine" ' Class���̂�ێ�����

Private m_FlagInterceptor As Boolean

Private mImageTestInstance As Collection       ' �S�Ă̎B���e�X�g�C���X�^���X��ێ�����
Private mImageTestAction As CImageTestInstance ' ����̎B���e�X�g�C���X�^���X��ێ�����
Private mActionLogger As CActionLogger         ' ���sLog�f�[�^���_���v����ActionLogger��ێ�����
Private mSheetTestInstance As CWorkSheetReader ' TestInstance�V�[�g��ێ�����

Public Function CreateScenario() As Integer
'���e:
'   �B���e�X�g�C���X�^���X�𐶐�����
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


    '#####  �B���e�X�g�C���X�^���X�𐶐�����  #####
    CreateScenario = loadTestInstance


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    CreateScenario = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Strings.Chr$(13) & Strings.Chr$(13) & Err.Description
    Exit Function


End Function

Private Function loadTestInstance() As Integer
'���e:
'   TestInstance�V�[�g��ǂݍ��݁A�B���e�X�g�C���X�^���X�𐶐�����
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


    '#####  Sheet���́�Template���̂�ݒ肷��  #####
    Dim strToptTemplateTool As String ' Template���̂�ێ�����
    strToptTemplateTool = "XLibToptFrameWork_T"


    '#####  TestInstance�V�[�g�̃f�[�^���擾����  #####
    Dim IFileStream As IFileStream
    Dim IParamReader As IParameterReader
    
    With mSheetTestInstance
        Set IFileStream = .AsIFileStream
        Set IParamReader = .AsIParameterReader
    End With


    '#####  �B���e�X�g�C���X�^���X�ɁA�ݒ肵�Ă���f�[�^���擾���ێ�����  #####
    Dim colAcqIns As Collection            ' AcquireInstance���̂�ێ�����
    Dim strAcquireInstanceName() As String ' AcquireInstance���̂�ێ�����
    Dim strAcqInsName As Variant           ' AcquireInstance���̂�ێ�����
    Dim strTestInstanceName As String      ' TestInstance���̂�ێ�����
    Dim clsUsrMcrAct As IUserMacroAction
    Set mImageTestInstance = New Collection

    Do While Not IFileStream.IsEOR
        
        If IParamReader.ReadAsString("Name@Test Procedure") = strToptTemplateTool Then

            '#####  TestInstance���̂��擾����  #####
            Set mImageTestAction = New CImageTestInstance

            strTestInstanceName = IParamReader.ReadAsString("Test Name")
            With mImageTestAction
                .Initialize
                .SetTestInstanceName = strTestInstanceName
            End With

            '#####  AcquireInstance���̂��擾����  #####
            Set colAcqIns = New Collection
            
            If IParamReader.ReadAsString("Arg0@Other Parameters") <> "" Then
                strAcquireInstanceName = Split(Replace(IParamReader.ReadAsString("Arg0@Other Parameters"), " ", ""), ",")
                For Each strAcqInsName In strAcquireInstanceName
                    colAcqIns.Add strAcqInsName
                Next strAcqInsName
                mImageTestAction.SetAcquireInstance = colAcqIns
            End If
            
            Set colAcqIns = Nothing
            
            '#####  UserMacro���̂��擾����  #####
            Set clsUsrMcrAct = New CUserMacroAction
            
            If IParamReader.ReadAsString("Arg1@Other Parameters") <> "" Then
                clsUsrMcrAct.SetUserMacroName = IParamReader.ReadAsString("Arg1@Other Parameters")
                mImageTestAction.SetUserMacro = clsUsrMcrAct
            End If
            
            Set clsUsrMcrAct = Nothing

            '#####  �B���e�X�g�C���X�^���X����ێ�����  #####
            mImageTestInstance.Add mImageTestAction, mImageTestAction.GetTestInstanceName
            
            Set mImageTestAction = Nothing
        
        End If

        '#####  ���̍s�ֈړ�����  #####
        IFileStream.MoveNext

    Loop

    Set IFileStream = Nothing
    Set IParamReader = Nothing


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    loadTestInstance = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Strings.Chr$(13) & Strings.Chr$(13) & Err.Description
    Exit Function


End Function

Public Function SetScenario() As Integer
'���e:
'   �B���e�X�g�����s���邽�߂̏���������
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


    '#####  �B���e�X�g��񂪖�����΁A�B���e�X�g���𐶐�����  #####
    If mImageTestInstance Is Nothing Then
        SetScenario = Me.CreateScenario
        If SetScenario = TL_ERROR Then
            Set mImageTestInstance = Nothing
            Exit Function
        End If
    End If


    '#####  �B���e�X�g���s����������  #####
    Set mImageTestAction = mImageTestInstance.Item(TheExec.DataManager.InstanceName)


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:

    '#####  �e���v���[�g���g�p�̂��߁A�B���e�X�g�C���X�^���X�𐶐�����  #####
    SetScenario = createTestInstance
    Exit Function


End Function

Private Function createTestInstance() As Integer
'���e:
'   �e���v���[�g���g�p���̎B���e�X�g�C���X�^���X�𐶐�����
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


    '#####  �p�����[�^���擾����  #####
    Dim strArg() As String ' Parameter��ێ�����
    Dim lngArgCnt As Long  ' Parameter����ێ�����

    Call TheExec.DataManager.GetArgumentList(strArg, lngArgCnt)
    
    If lngArgCnt = 0 Then
        Call TheError.Raise(9999, CLASS_NAME & ".createTestInstance", "Can Not Create Image Test Instance ! " & vbCrLf _
                            & "You Have To Check Test Instance Sheet. ( " & TheExec.DataManager.InstanceName & " )")
    ElseIf lngArgCnt > 2 Then
        If strArg(0) = "" And strArg(1) = "" Then
            Call TheError.Raise(9999, CLASS_NAME & ".createTestInstance", "Can Not Create Image Test Instance ! " & vbCrLf _
                                & "You Have To Check Test Instance Sheet. ( " & TheExec.DataManager.InstanceName & " )")
        End If
    End If


    '#####  �e�X�g�C���X�^���X���̂��擾����  #####
    Dim clsImgTstIns As CImageTestInstance
    Set clsImgTstIns = New CImageTestInstance
    
    With clsImgTstIns
        .Initialize
        .SetTestInstanceName = TheExec.DataManager.InstanceName
    End With


    '#####  Acquire�C���X�^���X�𐶐�����  #####
    Dim colAcqIns As Collection       ' AcquireInstance���̂�ێ�����
    Dim strAcqireInstance() As String ' AcquireInstance���̂�ێ�����
    Dim strTmp As Variant             ' AcquireInstance���̂�ێ�����
    Set colAcqIns = New Collection
    
    If strArg(0) <> "" Then
        strAcqireInstance = Split(Replace(strArg(0), " ", ""), ",")
        For Each strTmp In strAcqireInstance
            colAcqIns.Add strTmp
        Next strTmp
        clsImgTstIns.SetAcquireInstance = colAcqIns
    End If
    
    Set colAcqIns = Nothing
    
    
    '#####  UserMacro�C���X�^���X�𐶐�����  #####
    Dim clsUsrMcr As IUserMacroAction
    Set clsUsrMcr = New CUserMacroAction

    If lngArgCnt >= 2 Then
        If strArg(1) <> "" Then
            clsUsrMcr.SetUserMacroName = strArg(1)
            clsImgTstIns.SetUserMacro = clsUsrMcr
        End If
    End If

    Set clsUsrMcr = Nothing


    '#####  �B���e�X�g�C���X�^���X���i�[����  #####
    Set mImageTestAction = clsImgTstIns
    mImageTestInstance.Add clsImgTstIns, clsImgTstIns.GetTestInstanceName

    Set clsImgTstIns = Nothing


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    createTestInstance = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function Execute() As Integer
'���e:
'   �B���e�X�g�����s����
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


    '#####  �B���e�X�g�C���X�^���X�̑��݂��m�F����  #####
    If mImageTestAction Is Nothing Then
        Execute = Me.SetScenario
        If Execute = TL_ERROR Then
            Set mImageTestAction = Nothing
            Exit Function
        End If
    End If


    '#####  �B���e�X�g��AcquireFrameWork�����s����  #####
    Dim i As Integer

    With mImageTestAction
        If .GetAcquireInstanceCount <> 0 Then
            For i = 1 To .GetAcquireInstanceCount Step 1
                Execute = ExecuteAcquireAction(.GetAcquireInstanceName(i))
                If Execute = TL_ERROR Then
                    Exit Function
                End If
            Next i
        End If
    End With


    '#####  �B���e�X�g��UserMacro�����s����  #####
    With mImageTestAction
        If .GetUserMacroCount <> 0 Then
            Execute = .Execute
            If Execute = TL_ERROR Then
                Exit Function
            End If
        End If
    End With

    Set mImageTestAction = Nothing


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    Execute = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function RetryAcquire(ByRef strAcquireInstanceName As String, Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'���e:
'   ���s�����폜���āAAcquireFrameWork���Ď��s����
'
'�p�����[�^:
'   [strAcquireInstanceName]   In  AcquireInstance���̂�ێ�����
'   [strAcquireFrameWorkName]  In  AcquireFrameWork���̂�ێ�����
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWork���Ď��s����  #####
    If strAcquireInstanceName <> "" Then
        RetryAcquire = RetryAcquireAction(strAcquireInstanceName, strAcquireFrameWorkName)
        If RetryAcquire = TL_ERROR Then
            Call TheError.Raise(9999, CLASS_NAME & ".RetryAcquire", "Can Not Execute RetryAcquire ! " & vbCrLf _
                                & "You Have To Check Parameters. ( " & strAcquireInstanceName & " / " & strAcquireFrameWorkName & " )")
        End If

    Else
        Call TheError.Raise(9999, CLASS_NAME & ".RetryAcquire", "Can Not Execute RetryAcquire ! " & vbCrLf _
                            & "You Have To Set Acquire Instance Name.")
    End If


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    RetryAcquire = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Sub EnableInterceptor(ByVal pFlag As Boolean, ByRef pLogger As CActionLogger)
'�܂�
    
    Dim pObj As CImageTestInstance
    For Each pObj In mImageTestInstance
        Call pObj.EnableInterceptor(pFlag, pLogger)
    Next pObj

End Sub

Public Sub Initialize(ByRef clsActionLogger As CActionLogger, ByRef clsWorkSheetReader As CWorkSheetReader)
    Set mImageTestInstance = Nothing
    Set mImageTestAction = Nothing
    Set mActionLogger = clsActionLogger
    Set mSheetTestInstance = clsWorkSheetReader
End Sub

Private Sub Class_Initialize()
    Set mImageTestInstance = Nothing
    Set mImageTestAction = Nothing
    Set mActionLogger = Nothing
    Set mSheetTestInstance = Nothing
End Sub

Private Sub Class_Terminate()
    Set mImageTestInstance = Nothing
    Set mImageTestAction = Nothing
    Set mActionLogger = Nothing
    Set mSheetTestInstance = Nothing
End Sub
