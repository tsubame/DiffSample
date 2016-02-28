Attribute VB_Name = "XLibTOPT_FW"
'�T�v:
'   TOPT�Ή�����FrameWork
'   Image ACQTBL �V�[�g�̂���̌Ăяo�����W���[���A����� Image ACQTBL �V�[�g�̏��Ɋ�Â��I�u�W�F�N�g�Q�̐���
'
'   Revision History:
'       Data        Description
'       2010/04/28  AcquireFrameWork�����s����@�\����������
'                   �iFWSetCondition / FWImageAcquire / FWPostImageAcquire�j
'       2010/06/11  �f�[�^�\���������̂��߁A�v���O�����R�[�h��ύX����
'       2010/06/22  ���s��񃊃Z�b�g�@�\����������
'                   CheckTermination�@�\����������
'                   Arg0,Arg1�̃G���[�����@�\����������
'                   FW�̃G���[�����@�\����������
'       2010/07/02  AcquireInstance�̍쐬���̕s����C�������iCount<>0�œo�^�j
'
'�쐬��:
'   tomoyoshi.takase
'   0145184346
'

Option Explicit

Private Const ERR_NUMBER = 9999                  '�G���[���ɓn���G���[�ԍ�
Private Const CLASS_NAME = "XLibTOPT_FW"         '���̃N���X�̖��O
Private Const FW_KEY = "ImageACQTBL Sheet(ARG0)" '�t���[�����[�N�̓��쌈�肷����̏ꏊ�B�G���[�o�͗p

Private mActionLogger As CActionLogger ' ActionLogger��ێ�����
Private mAcquireInstance As Collection ' AcquireFrameWork��ێ�����
Private mImageCheckErrorMsg As String
Private mImageCheckCounter As Long

Public Sub AcquireInitialize(ByRef clsActionLogger As CActionLogger, ByRef reader As CWorkSheetReader)
'���e:
'�@ImageACQTBL �V�[�g��������W���āA�e�C���X�^���X�̐������s���܂��B
'
'�p�����[�^:
'   [clsActionLogger]  In  ActionLogger��ێ�����
'   [reader]           In  ImageAcquireTable�V�[�g����ێ�����
'
'�߂�l:
'
'���ӎ���:
'


    On Error GoTo ErrHandler


    '#####  ActionLogger�C���X�^���X����  #####
    Set mActionLogger = clsActionLogger


    '#####  ImageAcquireTable�����擾����  #####
    Dim strmReader As IFileStream
    Dim paramReader As IParameterReader
    Set strmReader = reader
    Set paramReader = reader


    '#####  AcquireInstance�𐶐�����  #####
    Dim strFrameworkName As String
    Dim strInstanceName As String
    Dim strAutoAcquire As String
    Dim strLastInsName As String
    Dim strArg0 As String
    Dim strArg1 As String
    Dim strArg2 As String
    Dim strArg3 As String
    Dim clsAcqAct As IAcquireAction
    Dim clsAcqIns As CAcquireInstance
    Set clsAcqIns = New CAcquireInstance
    Set mAcquireInstance = New Collection

    mImageCheckCounter = 0
    mImageCheckErrorMsg = ""
    
    Do While strmReader.IsEOR <> True
        strLastInsName = strInstanceName
        strFrameworkName = paramReader.ReadAsString("Macro Name")
        strInstanceName = paramReader.ReadAsString("Instance Name")
        strAutoAcquire = paramReader.ReadAsString("Auto Acquire")

        '#####  Instance���̂̊m�F  #####
        If strInstanceName = "" Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                                "Can Not Found Instance Name ! => " & strFrameworkName & "(ImageACQTBL Sheet)")
        ElseIf strLastInsName <> "" And strInstanceName <> strLastInsName Then
            On Error GoTo VBAErrHandler
            If clsAcqIns.Count <> 0 Then
                mAcquireInstance.Add clsAcqIns, strLastInsName
            End If
            Set clsAcqIns = New CAcquireInstance
            On Error GoTo ErrHandler
        End If

        '#####  Arg0,Arg1�̊m�F  #####
        strArg0 = paramReader.ReadAsString("Arg0@Parameters")
        If strArg0 = "" Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                                "Can Not Found Instance Name !(Arg0) => " & strFrameworkName & _
                                "(" & strInstanceName & ")" & " (ImageACQTBL Sheet)")
        End If
        If strArg0 <> strInstanceName Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                                "Acquire Instance Name and Arg0 is not same! => " & strFrameworkName & _
                                "(Inst:" & strInstanceName & " <=> Arg0:" & strArg0 & ")" & _
                                " (ImageACQTBL Sheet)")
        End If
        If strFrameworkName = "FWImageAcquire" Or strFrameworkName = "FWPostImageAcquire" Then
            strArg1 = paramReader.ReadAsString("Arg1@Parameters")
            If strArg1 = "" Then
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                                    "Can Not Found Acquire Macro Name !(Arg1) => " & strFrameworkName & _
                                    "(" & strInstanceName & ")" & " (ImageACQTBL Sheet)")
            End If
        End If

        '#####  FrameWork���Ƃ̏���  #####
        If strAutoAcquire <> "nop" Then
            If strFrameworkName = "FWSetCondition" Then
                Set clsAcqAct = New CSetAction
'                Call clsAcqAct.Initialize(paramReader, mActionLogger)
                Call clsAcqAct.Initialize(paramReader)
                clsAcqIns.SetAction = clsAcqAct
            ElseIf strFrameworkName = "FWImageAcquire" Then
                Set clsAcqAct = New CAcquireAction
'                Call clsAcqAct.Initialize(paramReader, mActionLogger)
                Call clsAcqAct.Initialize(paramReader)
                clsAcqIns.SetAction = clsAcqAct
            ElseIf strFrameworkName = "FWPostImageAcquire" Then
                Set clsAcqAct = New CPostAcquireAction
'                Call clsAcqAct.Initialize(paramReader, mActionLogger)
                Call clsAcqAct.Initialize(paramReader)
                clsAcqIns.SetAction = clsAcqAct
            Else
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                                    "Illegal Frame Work Name ! => " & strFrameworkName & "(ImageACQTBL Sheet)")
            End If
        End If

        '#####  Arg2,3�̉摜�p�����[�^�̃`�F�b�N  #####
        If strFrameworkName = "FWImageAcquire" Then
            strArg2 = paramReader.ReadAsString("Arg2@Parameters")
            strArg3 = paramReader.ReadAsString("Arg3@Parameters")
            If ChkImageParamter(strArg2, strArg3) = False Then
                Call StockErr(strFrameworkName, strInstanceName, strArg2, strArg3)
            End If
        End If
        
        '#####  ���̍s�Ɉړ�  #####
        Call strmReader.MoveNext

        '#####  �ŏI�s�̏���  #####
        If strmReader.IsEOR = True Then
            If clsAcqIns.Count <> 0 Then
                mAcquireInstance.Add clsAcqIns, strInstanceName
            End If
            Set clsAcqAct = Nothing
            Set clsAcqIns = Nothing
        End If
    
    Loop

    If mImageCheckCounter > 0 Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME, _
                            "ImageParameterCheck Error!" & vbCrLf & "ImageAcquireSheet Arg2-3" & vbCrLf & vbCrLf & mImageCheckErrorMsg)
    End If
    
    Set strmReader = Nothing
    Set paramReader = Nothing

    Exit Sub

'#####  �G���[���b�Z�[�W�������I��  #####
VBAErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AcquireInitialize", _
                        "This Instance Name is already setting." & "(" & strLastInsName & ")")
    Exit Sub
ErrHandler:
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub

End Sub

Private Sub StockErr(ByVal pMacroName As String, ByVal pInstanceName As String, ByVal pBaseName As String, ByVal pStrBitDepth As String)
    mImageCheckErrorMsg = mImageCheckErrorMsg & _
                            "FrameWorkname : " & pMacroName & vbCrLf & _
                            "InstanceName  : " & pInstanceName & vbCrLf & _
                            "ImageBaseName : " & pBaseName & vbCrLf & _
                            "ImageBitDepth : " & pStrBitDepth & vbCrLf
    mImageCheckCounter = mImageCheckCounter + 1
End Sub

Private Function ChkImageParamter(ByVal pBaseName As String, ByVal pStrBitDepth As String) As Boolean
'�p�����[�^�`�F�b�N���Ă��߂Ȃ�False
    Dim pBitDepth As CIdpBitDepth
    Dim PNum As Long
    On Error GoTo NOTHING_IMAGE
    Set pBitDepth = New CIdpBitDepth
    Call pBitDepth.SetValue(pStrBitDepth)
    PNum = TheIDP.PlaneManager(pBaseName).Count(pBitDepth.GetValue)
    If PNum > 0 Then
        ChkImageParamter = True
    Else
        ChkImageParamter = False
    End If
    Exit Function
NOTHING_IMAGE:
    ChkImageParamter = False
End Function

Private Function GetAcqIns(pFuncName As String, pArg0 As String) As CAcquireInstance
    On Error GoTo VBAErrHandler
    Set GetAcqIns = mAcquireInstance.Item(pArg0)
    Exit Function
VBAErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & "." & pFuncName, _
                        "Illegal Instance Name." & "(" & pArg0 & ")" & "(ImageACQTBL Sheet)")
End Function

Public Function FWSetCondition() As Long
'���e:
'�@TOPT.GetArgumentList�œ�������SetCondition�����s���܂��B
'�@TOPT.Start �܂��� TOPT Auto Acquire �Ŏ��s����܂��B
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
    
    '#####  ImageAcquireTable�̃p�����[�^���擾  #####
    Dim ArgImageAcqtbl() As String

    Call TheHdw.TOPT.GetArgumentList(ArgImageAcqtbl)
    If ArgImageAcqtbl(0) = "" Then
        Call TheError.Raise(ERR_NUMBER, "FWSetCondition", FW_KEY & " is Nothing!")
    End If
    
'    TheExec.Datalog.WriteComment ArgImageAcqtbl(0)
'    Call StartTime
    '#####  AcquireFrameWork�̃C���X�^���X���擾����  #####
    Dim clsAcqIns As CAcquireInstance
    
    Set clsAcqIns = GetAcqIns("FWSetCondition", ArgImageAcqtbl(0))

    '#####  �����Ă���TOPT���̒񎦗p  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = ArgImageAcqtbl(0)
    TheIDP.PlaneList.ACQTBL_Arg1 = ArgImageAcqtbl(1)

    '#####  AcquireFrameWork�̃C���X�^���X�����s����  #####
    FWSetCondition = clsAcqIns.Execute("FWSetCondition")

    Set clsAcqIns = Nothing

    '#####  TOPT FW �̏I�����  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = "-"
    TheIDP.PlaneList.ACQTBL_Arg1 = "-"
    
'    Call StopTime
    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    FWSetCondition = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function FWImageAcquire() As Long
'���e:
'�@TOPT.GetArgumentList�œ�������ImageAcquire�����s���܂��B
'�@TOPT.Start �܂��� TOPT Auto Acquire �Ŏ��s����܂��B
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

    Dim Acqsite As Long
    
    For Acqsite = 0 To nSite
        If TheExec.sites.site(Acqsite).Active = False Then
            TheExec.sites.site(Acqsite).Active = True
            If Flg_FailSiteImage(Acqsite) = False Then
                '@@@ DUT��񖈂ɒ��g��ւ��Ȃ��Ƃ����Ȃ��B@@@
                Call DisconnectAllDevicePins(Acqsite)                 'FailSite All OPEN   '2012/11/16 175JobMakeDebug
                Call GND_DisConnect(Acqsite)                          '2012/11/16 175JobMakeDebug
                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                Flg_FailSiteImage(Acqsite) = True
            End If
            TheExec.sites.site(Acqsite).Active = False
        End If
    Next

    '#####  ImageAcquireTable�̃p�����[�^���擾  #####
    Dim ArgImageAcqtbl() As String

    Call TheHdw.TOPT.GetArgumentList(ArgImageAcqtbl)
    If ArgImageAcqtbl(0) = "" Then
        Call TheError.Raise(ERR_NUMBER, "FWImageAcquire", FW_KEY & " is Nothing!")
    End If


    '#####  AcquireFrameWork�̃C���X�^���X���擾����  #####
    Dim clsAcqIns As CAcquireInstance

    Set clsAcqIns = GetAcqIns("FWImageAcquire", ArgImageAcqtbl(0))

    '#####  �����Ă���TOPT���̒񎦗p  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = ArgImageAcqtbl(0)
    TheIDP.PlaneList.ACQTBL_Arg1 = ArgImageAcqtbl(1)

    '#####  AcquireFrameWork�̃C���X�^���X�����s����  #####
    FWImageAcquire = clsAcqIns.Execute("FWImageAcquire")

    Set clsAcqIns = Nothing

    '#####  TOPT FW �̏I�����  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = "-"
    TheIDP.PlaneList.ACQTBL_Arg1 = "-"

    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    FWImageAcquire = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function FWPostImageAcquire() As Long
'���e:
'�@TOPT.GetArgumentList�œ�������PostImageAcquire�����s���܂��B
'�@TOPT.Start �܂��� TOPT Auto Acquire �Ŏ��s����܂��B
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


    '#####  ImageAcquireTable�̃p�����[�^���擾����  #####
    Dim ArgImageAcqtbl() As String

    Call TheHdw.TOPT.GetArgumentList(ArgImageAcqtbl)
    If ArgImageAcqtbl(0) = "" Then
        Call TheError.Raise(ERR_NUMBER, "FWPostImageAcquire", FW_KEY & " is Nothing!")
    End If


    '#####  AcquireFrameWork�̃C���X�^���X���擾����  #####
    Dim clsAcqIns As CAcquireInstance
    
    Set clsAcqIns = GetAcqIns("FWPostImageAcquire", ArgImageAcqtbl(0))

    '#####  �����Ă���TOPT���̒񎦗p  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = ArgImageAcqtbl(0)
    TheIDP.PlaneList.ACQTBL_Arg1 = ArgImageAcqtbl(1)

    '#####  AcquireFrameWork�̃C���X�^���X�����s����  #####
    FWPostImageAcquire = clsAcqIns.Execute("FWPostImageAcquire")

    Set clsAcqIns = Nothing

    '#####  TOPT FW �̏I�����  #####
    TheIDP.PlaneList.ACQTBL_Arg0 = "-"
    TheIDP.PlaneList.ACQTBL_Arg1 = "-"

    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    FWPostImageAcquire = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function ExecuteAcquireAction(ByRef strAcquireInstanceName As String, Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'���e:
'   AcquireFrameWork�C���X�^���X�����s����
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

    '#####  AcquireFrameWork�C���X�^���X���擾����  #####
    Dim clsAcqIns As CAcquireInstance
    Set clsAcqIns = GetAcqIns("ExecuteAcquireAction", strAcquireInstanceName)

    '#####  AcquireFrameWork�����s����  #####
    ExecuteAcquireAction = clsAcqIns.ToptStart(strAcquireFrameWorkName)
    If ExecuteAcquireAction = TL_ERROR Then
        Exit Function
    End If

    Set clsAcqIns = Nothing


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    ExecuteAcquireAction = TL_ERROR
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function


End Function

Public Function RetryAcquireAction(ByRef strAcquireInstanceName As String, ByRef strAcquireFrameWorkName As String) As Integer
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


    '#####  AcquireFrameWork�̎��s�������폜����  #####
    RetryAcquireAction = StartClearStatus(strAcquireInstanceName, strAcquireFrameWorkName)
    If RetryAcquireAction = TL_ERROR Then
        Exit Function
    End If
    

    '#####  AcquireFrameWork���Ď��s����  #####
    TheIDP.PlaneBank.IsOverwriteMode = True
    RetryAcquireAction = ExecuteAcquireAction(strAcquireInstanceName, strAcquireFrameWorkName)
'    TheIDP.PlaneBank.IsOverWriteMode = False


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    RetryAcquireAction = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function StartClearStatus(ByRef strAcquireInstanceName As String, Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'���e:
'   AcquireFrameWork�C���X�^���X�̎��s�������폜����
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


    '#####  AcquireFrameWork���擾����  #####
    Dim clsAcqIns As CAcquireInstance
    Set clsAcqIns = GetAcqIns("StartClearStatus", strAcquireInstanceName)


    '#####  AcquireFrameWork�����s����  #####
    StartClearStatus = clsAcqIns.ClearStatus(strAcquireFrameWorkName)
    If StartClearStatus = TL_ERROR Then
        Exit Function
    End If

    Set clsAcqIns = Nothing


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    StartClearStatus = TL_ERROR
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function


End Function

Public Sub ResetStatus()
'���e:
'   AcquireFrameWork�C���X�^���X�̎��s�������N���A����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'


    '#####  AcquireFrameWork���擾����  #####
    Dim i As Integer
    Dim clsAcqIns As CAcquireInstance
    
    For i = 1 To mAcquireInstance.Count
        Set clsAcqIns = mAcquireInstance.Item(i)
        clsAcqIns.Reset
    Next i

    Set clsAcqIns = Nothing


    '#####  �I��  #####
    Exit Sub


End Sub

Public Function CheckTermination(ByRef strAcquireInstanceName As String, Optional ByRef strAcquireFrameWorkName As String = "") As Integer
'���e:
'   AcquireFrameWork�C���X�^���X�̎��s���m�F����
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


    '#####  AcquireFrameWork�C���X�^���X���擾����  #####
    Dim clsAcqIns As CAcquireInstance
    Set clsAcqIns = GetAcqIns("CheckTermination", strAcquireInstanceName)


    '#####  AcquireFrameWork�C���X�^���X��Status�ɂ���āATOPT�����s����  #####
    CheckTermination = clsAcqIns.CheckTermination(strAcquireFrameWorkName)

    Set clsAcqIns = Nothing


    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    CheckTermination = TL_ERROR
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function


End Function

Public Sub DestroyTOPTFW()
    Set mActionLogger = Nothing
    Set mAcquireInstance = Nothing
End Sub

Public Sub CreateTOPTFWIfNothing()
'���e:
'   As New�̑�ւƂ��ď���ɌĂ΂���C���X�^���X��������
'
'�p�����[�^:
'   �Ȃ�
'
'���ӎ���:
'
    
    On Error GoTo ErrHandler
    If mAcquireInstance Is Nothing Then
        Call AcquireInitialize(GetActionLoggerInstance, GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeAcquire))
    End If
    Exit Sub
ErrHandler:
    Call DestroyTOPTFW
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub EnableReadImage(ByVal pFlag As Boolean, ByVal pPathName As String)
    Dim pAcqObj As CAcquireInstance
    For Each pAcqObj In mAcquireInstance
        Call pAcqObj.EnableReadImage(pFlag, pPathName)
    Next pAcqObj
End Sub

Public Sub EnableSaveImage(ByVal pFlag As Boolean, ByVal pPathName As String)
    Dim pAcqObj As CAcquireInstance
    For Each pAcqObj In mAcquireInstance
        Call pAcqObj.EnableSaveImage(pFlag, pPathName)
    Next pAcqObj
End Sub

Public Sub EnableShowImage(ByVal pFlag As Boolean)
    Dim pAcqObj As CAcquireInstance
    For Each pAcqObj In mAcquireInstance
        Call pAcqObj.EnableShowImage(pFlag)
    Next pAcqObj
End Sub

Public Sub EnableInterceptor(ByVal pFlag As Boolean, ByRef pLogger As CActionLogger)
    Dim pAcqObj As CAcquireInstance
    For Each pAcqObj In mAcquireInstance
        Call pAcqObj.EnableInterceptor(pFlag, pLogger)
    Next pAcqObj
End Sub
