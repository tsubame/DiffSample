Attribute VB_Name = "XLibToptFrameWorkUtility"
'�T�v:
'   ToptFrameWork�̃��[�e�B���e�B
'
'�ړI:
'   XLibIGXLEvents��On�`���\�b�h�ŌĂ΂������֐����`����
'
'�쐬��:
'   a_oshima

Option Explicit

Private mFailedJobInitialize As Boolean

Public Function ResetEeeJobObjects() As Long
'�T�v:
'   OnProgramLoaded�ŌĂ΂������֐�
'
    On Error GoTo ErrHandler
    Call DestroyAllEeeJobObjects

    Exit Function
ErrHandler:
    MsgBox "Failed at Eee-Job Initialize Routine: " & CStr(Err.Number) & " - " & Err.Source & vbCrLf & vbCrLf & Err.Description
End Function

Public Function ResetEeeJobSheetObjects() As Long
'�T�v:
'   OnProgramValidated�ŌĂ΂������֐�
'

    On Error GoTo ErrHandler
    Call DestroyAllEeeJobSheetObjects

    Exit Function
ErrHandler:
    MsgBox "Failed at Eee-Job Initialize Routine: " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
End Function

Public Function RunAtValidated() As Long
'�T�v:
'   OnProgramValidated�ŌĂ΂������֐�
'

    On Error GoTo ErrHandler

    Call CreateTheErrorIfNothing
    Call CreateReaderManagerIfNothing
    Call CreateTheParameterBankIfNothing        'add 20110209
    Call CreateTheIDPIfNothing
    Call CreatePlaneMapIfNothing
    Call CreateKernelManagerIfNothing
    Call CreateTheImageTestIfNothing
    Call CreateTheConditionIfNothing
    Call XLibSetConditionUtility.ChangeDefaultSettingTheCondition
#If ITS <> 0 Then
    Call XLibImpUIControllerUtility.RunAtValidated
#End If
    Exit Function
ErrHandler:
    MsgBox "Failed at Eee-Job Initialize Routine: " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
End Function

Public Function RunAtJobStart() As Long
'�T�v:
'   OnProgramStarted�ŌĂ΂������֐�
'
    mFailedJobInitialize = False

    On Error GoTo ErrHandler
    '### �C���X�^���X��������ΐ��� #######################
    Call CreateTheErrorIfNothing
    Call CreateReaderManagerIfNothing
    Call CreateTheParameterBankIfNothing        'add 20110209
    Call CreateTheSystemInfoIfNothing
    Call CreateTheIDPIfNothing
    Call CreateTheVarBankIfNothing
    Call CreateTheFlagBankIfNothing

    Call CreatePlaneMapIfNothing
    Call CreatePMDIfNothing
    Call CreateKernelManagerIfNothing
    Call CreateTheImageTestIfNothing
#If ITS <> 0 Then
    Call CreateScenarioBuilderIfNothing
    Call CreateTheImgTestScenarioIfNothing
#End If
    Call CreateTheConditionIfNothing
    Call CreateTOPTFWIfNothing
    Call CreateTheDeviceProfilerIfNothing
    Call TheParameterBank.Clear
    Call InitTestScenario

    '### �e�X�g���̃��Z�b�g��� ###########################
    TheIDP.ResetTest
    XLibTOPT_FW.ResetStatus

    '### �e���K�[��Disable�ɐݒ� ##########################
    'TheVarBank�̃��O�@�\�͊J���Ҍ����Ƃ��ĕ���
'    Call SetLogModeActionLogger(True)
'    Call SetLogModeTheIDP(True)
'    Call SetLogModeTheCondition(True)

'    Call SaveModeTheVarBank(True)

    XLibActionLoggerUtility.ApplyLogModeActionLogger
    
#If ITS <> 0 Then
    Call XLibImpUIControllerUtility.RunAtJobStart
    Call PrepareDumpDirectory(ThisWorkbook.Path & Application.PathSeparator & "Dump")
#End If

    Exit Function
ErrHandler:
    MsgBox "Failed at Job Initialize Routine: " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Call AbortJob
End Function

Public Function RunAtJobEnd() As Long
'�T�v:
'   OnProgramEnded�ŌĂ΂������֐�

'{ ���O���|�[�g���C�^�[�̌㏈��
    Call CloseDcLogReportWriter
'}
    Call XLibDcScenarioLoopOption.RunAtJobEnd
    
    Call XLibActionLoggerUtility.RunAtJobEnd
    Call XLibImageEngineUtility.RunAtJobEnd
#If ITS <> 0 Then
    Call XLibImSceEngineUtility.RunAtJobEnd
    Call XLibScenarioUtility.RunAtJobEnd
#End If
    Call XLibTheVarBankUtility.RunAtJobEnd
    Call XLibImgUtility.RunAtJobEnd
    Call XLibSetConditionUtility.RunAtJobEnd
    Call XLibTheFlagBankUtility.RunAtJobEnd
    Call XLibTheDeviceProfilerUtility.RunAtJobEnd
    Call XLibTheParameterBankUtility.RunAtJobEnd

    Unload ScenarioParameterViewer
#If ITS <> 0 Then
    Call XLibImpUIControllerUtility.RunAtJobEnd
#End If

'    TheIDP.ResetTest
    '### TheError�̃��O�@�\�͊J���Ҍ����Ƃ��ĕ��� #########
'    Call XLibErrManangerUtility.RunAtJObEnd(pSaveFileName:="TheErrorLog.csv")

End Function

Public Sub AbortJob()
    mFailedJobInitialize = True
End Sub

Public Function FailedJobInitialize() As Boolean
    FailedJobInitialize = mFailedJobInitialize
End Function

Public Sub DestroyAllEeeJobObjects()
    Call DestroyTheError
    Call DestroyTheIDP
    Call DestroyPMDSheet
    Call DestroyActionLogger
#If ITS <> 0 Then
    Call DestroyImpUIController
#End If

    Call DestroyAllEeeJobSheetObjects
End Sub

Public Sub DestroyAllEeeJobSheetObjects()
    Call DestroyTheVarBank
    Call DestroyTheFlagBank
    Call DestroyWkShtReaderManager
#If ITS <> 0 Then
    Call DestroyTheImgTestScenario
    Call DestroyScenarioBuilder
#End If
    Call DestroyTheImageTest
    Call DestroyTestCondition
    Call DestroyTOPTFW
    Call DestroyTheDeviceProfiler
End Sub

