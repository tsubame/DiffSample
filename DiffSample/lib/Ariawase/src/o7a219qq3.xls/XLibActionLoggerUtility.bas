Attribute VB_Name = "XLibActionLoggerUtility"
'�T�v:
'   ActionLogger�̃��[�e�B���e�B
'
'�ړI:
'   ActionLogger�̏�����/�j���̃��[�e���e�B���`����
'
'�쐬��:
'   a_oshima

Option Explicit

Private mEnableActionLogger As Boolean
Private mActionLogger As CActionLogger   ' ���sLog�f�[�^���_���v����ActionLogger��ێ�����
Private mActionLogFileName As String
Private Const DEFAULT_LOGNAME As String = "EeeJOBLogActionLogger.csv"

Public Sub CreateActionLoggerIfNothing()
'���e:
'   As New�̑�ւƂ��ď���ɌĂ΂���C���X�^���X��������
'   �����̃N���A���s��
'
'�p�����[�^:
'   �Ȃ�
'
'���ӎ���:
'
    On Error GoTo ErrHandler
    If mActionLogger Is Nothing Then
        Set mActionLogger = New CActionLogger
    End If
    Call mActionLogger.Initialize
    Exit Sub
ErrHandler:
    Set mActionLogger = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub SetLogModeActionLogger(ByVal pEnableLoggingActionLogger As Boolean, Optional ByVal pSaveActionLogFileName As String = DEFAULT_LOGNAME)

    If TheExec.JobIsValid = False Then
        MsgBox "The log mode cannot be set before Validation."
        Exit Sub
    End If

    Dim YesOrNo As VbMsgBoxResult
    If TheExec.RunMode = runModeProduction Then
        If pEnableLoggingActionLogger = True Then
            YesOrNo = MsgBox("Now ProductionMode! But SetLogModeActionLogger is True!" & vbCrLf & _
                            "EeeJOB will output DataLog." & vbCrLf & _
                            "Please make sure." & vbCrLf & _
                            "When You don't want to output This MsgBox, Please change runModeProduction => runModeDebug" & vbCrLf & _
                            "" & vbCrLf & _
                            "Output DataLog?" & vbCrLf _
                            , vbYesNo + vbQuestion, "Confirm Output DataLog(SetLogModeActionLogger)")
            If YesOrNo = vbYes Then
                mEnableActionLogger = True
            Else
                mEnableActionLogger = False
            End If
        Else
            mEnableActionLogger = False
        End If
    Else
        mEnableActionLogger = pEnableLoggingActionLogger
    End If

    If mEnableActionLogger = True Then
        TheExec.Datalog.WriteComment "Eee JOB Output Log! :Action Logger"
    End If

    mActionLogFileName = pSaveActionLogFileName

    If TheExec.Flow.IsRunning = True Then
        ApplyLogModeActionLogger
    End If

End Sub

Sub ApplyLogModeActionLogger()

    If mEnableActionLogger = True And mActionLogger Is Nothing Then
            CreateActionLoggerIfNothing
            Call XLibTOPT_FW.EnableInterceptor(mEnableActionLogger, mActionLogger)
            Call XLibImageEngineUtility.EnableInterceptor(mEnableActionLogger, mActionLogger)
#If ITS <> 0 Then
            Call EnableScenarioActionLogger(mEnableActionLogger, mActionLogger)
#End If
    ElseIf mEnableActionLogger = False And Not mActionLogger Is Nothing Then
            Call XLibTOPT_FW.EnableInterceptor(mEnableActionLogger, mActionLogger)
            Call XLibImageEngineUtility.EnableInterceptor(mEnableActionLogger, mActionLogger)
#If ITS <> 0 Then
            Call EnableScenarioActionLogger(mEnableActionLogger, mActionLogger)
#End If
            DestroyActionLogger
    End If
End Sub

Public Sub DestroyActionLogger()
    Set mActionLogger = Nothing
End Sub

Public Function GetActionLoggerInstance() As CActionLogger
'���e:
'   ActionLogger�̃C���X�^���X��Ԃ�
'
'�p�����[�^:
'   �Ȃ�
'
'�߂�l:
'   ActionLogger�̃C���X�^���X
'
'��O:
'   �����������ɌĂ΂���VBA��O����
'  �i�p�t�H�[�}���X���P�̂���AsNew�̑�ւƂ��ėp�ӂ��Ă���ANothing�`�F�b�N�͍s��Ȃ��j
'
'���ӎ���:
'   �������������ɌĂсA�C���X�^���X����������Ă��邱��

    'CreateActionLoggerIfNothing()
    Set GetActionLoggerInstance = mActionLogger
End Function

Public Function RunAtJobEnd() As Long
    If Not mActionLogger Is Nothing Then
        If mEnableActionLogger = True Then
            SaveActionLoggerHistoryLog
            Call XLibTOPT_FW.EnableInterceptor(False, mActionLogger)
            Call XLibImageEngineUtility.EnableInterceptor(False, mActionLogger)
#If ITS <> 0 Then
            Call EnableScenarioActionLogger(False, mActionLogger)
#End If
            DestroyActionLogger
            mEnableActionLogger = False
        End If
    End If
End Function



'ImageEngine����C���|�[�g
Public Function GetActionLoggerEnable() As Boolean
'�Ƃ肠�����A���True
'    GetActionLoggerEnable = mActionLogger.EnableLogging
    GetActionLoggerEnable = mEnableActionLogger
End Function

Public Function GetActionLogFileName() As String
    GetActionLogFileName = mActionLogFileName
End Function

Public Property Let SetActionLogFileName(ByRef strName As String)
    mActionLogFileName = strName
End Property

Public Function SaveActionLoggerHistoryLog()
    If mActionLogFileName <> "" Then
        Call mActionLogger.SaveHistoryLog(mActionLogFileName)
    Else
        Call mActionLogger.SaveHistoryLog(makeFileName)
    End If
End Function

Public Function ClearActionLoggerHistoryLog()
    Call mActionLogger.ClearHistories
End Function

Private Function makeFileName() As String
    makeFileName = getCurrentDirectory & "\" & getToday & "_" & "TOPT_ActionLog" & ".csv"
End Function

Private Function getToday() As String
    getToday = Format$(DateTime.Now, "yyyymmdd") & "_" & Format$(DateTime.Now, "hhnnss")
End Function

Private Function getCurrentDirectory() As String
    getCurrentDirectory = ActiveWorkbook.Path
End Function


