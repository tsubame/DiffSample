VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CActionLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �C���[�W�֘A��Action���s�������Ǘ�����Utility
'
'�ړI:
'   �C���[�W�֘A��Action���s�������i�[����
'   �S���s������CSV�`���ňꊇ�o�͂���
'
'�쐬��:
'   a_oshima
'
'�O�����:
'   CActionInfo�N���X�ACActionExecutionHistory�N���X�ACEeeTimer�N���X�𗘗p
'   �O���[�o���I�u�W�F�N�g TheError:CErrManager �����݂��邱��
'   �ŏ���Initialize���\�b�h���R�[�����邱�Ɓi�ᔽ���͗�O9999�����j
'
'���p�菇:
'   �@New�����炷��Initialize���R�[��
'   �AAction.Execute()���s���ƂɁA
'   �@�@�A-1 Action.Execute()�̃��[�U�}�N�����s�O��LogBeforeExecution���R�[�����Ă��炤
'   �@�@�A-2 ���[�U�}�N�����s
'   �@�@�A-3 Action.Execute()�̃��[�U�}�N�����s���LogAfterExecution���R�[�����Ă��炤
'   �BCSV�ꊇ�o��
'
'Action.Execute()�̏�����:
'   Private Sub DummyExec(ByRef pInfo As CActionInfo)
'       Call gActionLogger.LogBeforeExecution(pInfo)
'       Dim ret As Long
'       ret = Application.Run("DummyUserMacro_f")
'       Call gActionLogger.LogAfterExecution(pInfo, ret)
'   End Sub
'
'CSV�\��:
'

Option Explicit

Private mActionExecutionHistories As Collection '�������F�����o��CActionExecutionHistory�^
Private mEnableLogging As Boolean
Private mTimer As CEeeTimer
Private mCurrentHistoryID As Long
Private mMaxAcquireArgCount As Long   'AcquireTable��Arg�����΂�΂�̂Ƃ���CSV�񐔒����p

Private Const mSUCCESS As Long = 0&

Property Get GetHistory(ByRef pActionInfo As CActionInfo, ByRef pTrigger As String) As CActionExecutionHistory
'���e:
'�@�@�A�N�V�������, Trigger�ɍ��v���闚������Ԃ�
'
'�p�����[�^:
'   [pActionInfo]   IN  CActionInfo�^:�@�@�A�N�V�������
'   [pTrigger]   IN  String�^:�@�@�g���K�[���(Start/End)
'
'�߂�l:
'   CActionExecutionHistory�^  �Y�����闚��1��
'
'��O:
'   �w�肵��ActionInfo, Trigger�ɍ��v���闚�����Ȃ��ꍇ�͗�O����

    Dim flgFound As Boolean
    flgFound = False
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        If historyObj.Trigger = pTrigger Then
            If pActionInfo.IsSameName(historyObj.ActionInfo) Then
                flgFound = True
                Exit For
            End If
        End If
    Next historyObj

    If flgFound = True Then
        Set GetHistory = historyObj
    Else
        Call TheError.Raise(9999, TypeName(Me) & ".GetHistory", "Not found History:" & pActionInfo.ToString & " Trigger:" & pTrigger)
    End If

End Property
Property Get HistoryCount() As Long
    HistoryCount = mActionExecutionHistories.Count
End Property

Property Get EnableLogging() As Boolean
    EnableLogging = mEnableLogging
End Property

Private Sub Class_Terminate()
    Set mActionExecutionHistories = Nothing
    Set mTimer = Nothing
End Sub

Private Sub Class_Initialize()
'���e:
'�@�@�����������������A�^�C�}�[���X�^�[�g
    mEnableLogging = False
    Set mTimer = New CEeeTimer
    Call Initialize
End Sub

Public Sub Initialize()
'���e:
'�@�@�����������������A�^�C�}�[���X�^�[�g
    Call ClearHistories
    mTimer.Initialize
    mTimer.StartTimer

End Sub

Public Sub ClearHistories()
'���e:
'   ��������S�Ĕj��
    Set mActionExecutionHistories = New Collection
    mCurrentHistoryID = 0&
    mMaxAcquireArgCount = 0&
End Sub

Public Sub LogBeforeExecution(ByRef pActionInfo As CActionInfo)
'���e:
'   Action.Exec�̎��s�O�̗�����������
'   (EnableLogging��False�̂Ƃ��͉������Ȃ�)
'
'�p�����[�^:
'   [pActionInfo]  IN CActionInfo�^:�@�i�[����ActionInfo�i�i�[����̂̓N���[���j
'
'�O�����:
'   Action.Exec�Ń��[�U�}�N�����s���O�ɃR�[�������
'
    Call CheckInitializeCalled
'Interceptor�^�ɏC�����Ă��̃��\�b�h�����s�����O�o�͂ɕύX
'    If mEnableLogging = False Then
'        Exit Sub
'    End If

    mCurrentHistoryID = mCurrentHistoryID + 1

    Dim newHistory As CActionExecutionHistory
    Set newHistory = New CActionExecutionHistory
    newHistory.RecordNumber = CStr(mCurrentHistoryID)
    newHistory.Trigger = "Start"
    newHistory.RecordDate = Date$
    newHistory.RecordTime = TIME$
    newHistory.LapTime = mTimer.LapTimer
    Set newHistory.ActionInfo = pActionInfo.Clone()

    If mMaxAcquireArgCount < pActionInfo.Parameters.ArgParameterCount Then
        mMaxAcquireArgCount = pActionInfo.Parameters.ArgParameterCount
    End If

    On Error GoTo catch
    Call mActionExecutionHistories.Add(newHistory, CStr(mCurrentHistoryID))
    On Error GoTo 0
    Exit Sub
catch:
    If Err.Number = 457 Then
        Call TheError.Raise(9999, TypeName(Me) & ".LogBeforeExecution", "HistoryID appeared towice:" & mCurrentHistoryID)
    Else
        Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

Public Sub LogAfterExecution(ByRef pActionInfo As CActionInfo, Optional ByRef pErrorNo As Long = mSUCCESS)
'���e:
'   Action.Exec�̎��s��̗�����������
'   (EnableLogging��False�̂Ƃ��͉������Ȃ�)
'
'�p�����[�^:
'   [pActionInfo]  IN CActionInfo�^:�@�i�[����ActionInfo�i�i�[����̂̓N���[���j
'   [pErrorNo]   IN  Long�^:   Application.Run�ŕԂ����Ԓl�����̂܂ܓn��
'
'�O�����:
'   Action.Exec�Ń��[�U�}�N�����s����ɃR�[�������B

    Call CheckInitializeCalled
'Interceptor�^�ɏC�����Ă��̃��\�b�h�����s�����O�o�͂ɕύX
'    If mEnableLogging = False Then
'        Exit Sub
'    End If

    mCurrentHistoryID = mCurrentHistoryID + 1

    Dim newHistory As CActionExecutionHistory
    Set newHistory = New CActionExecutionHistory
    newHistory.RecordNumber = CStr(mCurrentHistoryID)
    newHistory.Trigger = "End"
    newHistory.RecordDate = Date$
    newHistory.RecordTime = TIME$
    newHistory.LapTime = mTimer.LapTimer
    newHistory.ErrorNo = pErrorNo
    Set newHistory.ActionInfo = pActionInfo.Clone()

    If mMaxAcquireArgCount < pActionInfo.Parameters.ArgParameterCount Then
        mMaxAcquireArgCount = pActionInfo.Parameters.ArgParameterCount
    End If

    On Error GoTo catch
    Call mActionExecutionHistories.Add(newHistory, CStr(mCurrentHistoryID))
    On Error GoTo 0
    Exit Sub
catch:
    If Err.Number = 457 Then
        Call TheError.Raise(9999, TypeName(Me) & ".LogAfterExecution", "HistoryID appeared towice:" & mCurrentHistoryID)
    Else
        Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub


Public Function IsStarted(ByRef pActionInfo As CActionInfo) As Boolean
'���e:
'   Action�����s�J�n�������ǂ����̖₢���킹
'
'[pActionInfo] IN CActionInfo�^:  LogBeforeExecution�Ŏw�肵��Action�Ɠ����e�X�g���E�}�N����������
'
'�߂�l:
'   Boolean�^�@�@True:���s�J�n�ρA�܂��͎��s�����A�܂��̓G���[�@�@False:�����s
'
    Dim histObj As CActionExecutionHistory
    On Error GoTo catch
    Set histObj = GetHistory(pActionInfo, "Start")
    On Error GoTo 0
    IsStarted = True
    Exit Function
catch:
    IsStarted = False
End Function


Public Function IsCompleted(ByRef pActionInfo As CActionInfo) As Boolean
'���e:
'   Action�����s�������Ă��邩�ǂ����̖₢���킹
'
'[pActionInfo] IN CActionInfo�^:  LogBeforeExecution�Ŏw�肵��Action�Ɠ����e�X�g���E�}�N����������
'
'�߂�l:
'   Boolean�^�@�@True:���s�����@�@False:���s�������A�G���[������
'
'�O�����:
'   LogBeforeExecution�����s�̏ꍇFalse
'
    If IsStarted(pActionInfo) Then
        Dim historyObj As CActionExecutionHistory
        On Error GoTo catch
        Set historyObj = GetHistory(pActionInfo, "End")
        On Error GoTo 0

        If historyObj.ErrorNo = mSUCCESS Then
            IsCompleted = True
        Else
            IsCompleted = False
        End If
        Set historyObj = Nothing
    Else
        IsCompleted = False
    End If
    Exit Function
catch:
    IsCompleted = False
End Function

Public Function IsError(ByRef pActionInfo As CActionInfo) As Boolean
'���e:
'   Action���G���[���ǂ����̖₢���킹
'
'[pActionInfo] IN CActionInfo�^:  LogBeforeExecution�Ŏw�肵��Action�Ɠ����e�X�g���E�}�N����������
'
'�߂�l:
'   Boolean�^�@�@True:�G���[�@�@False:�G���[���������A�܂��͎��s��
'
'�O�����:
'   LogBeforeExecution�����s�̏ꍇFalse
'
    If IsStarted(pActionInfo) Then
        Dim historyObj As CActionExecutionHistory
        On Error GoTo catch
        Set historyObj = GetHistory(pActionInfo, "End")
        On Error GoTo 0

        If historyObj.ErrorNo <> mSUCCESS Then
            IsError = True
        Else
            IsError = False
        End If
        Set historyObj = Nothing
    Else
        IsError = False
    End If
    Exit Function
catch:
    IsError = False
End Function

Private Function FileAndJobIndex() As String
    Dim RetFileName As String
    Dim RetJobName As String

    RetFileName = TheExec.ExcelHandle.ActiveWorkbook.Path & "\" & TheExec.ExcelHandle.ActiveWorkbook.Name
    RetJobName = TheExec.CurrentJob                             'JobList �� JobName ���B������΃t�@�C����

    '�t�@�C������}��
    FileAndJobIndex = """[Test Program]"",""[" & RetFileName & "]""" & vbCrLf & _
         """[Job]"",""[" & RetJobName & "]""" & vbCrLf
End Function

Public Function CSVIndex() As String
'���e:
'   CSV�̕\���Ԃ�
'
'�߂�l:
'   String�^   CSV�̕\��
'
'�O�����:
'   ������1���ȏ�L�^����Ă��邱��
'
    Dim str As String
    str = ""
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        str = historyObj.CSVIndex(mMaxAcquireArgCount) _
                & vbCrLf
        Exit For
    Next historyObj
    CSVIndex = str
End Function

Public Function ToCSVString() As String
'���e:
'   CSV�̓��e��Ԃ�
'
'�߂�l:
'   String�^   CSV�̓��e�i�����s�F1�s1���j
'
'�O�����:
'   ������1���ȏ�L�^����Ă��邱��
'
    Dim str As String
    str = ""
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        str = str & historyObj.ToCSVString(mMaxAcquireArgCount) _
                & vbCrLf
    Next historyObj
    ToCSVString = str
End Function

Public Function ToString() As String
'���e:
'   �����̓��e��Ԃ�
'
'�߂�l:
'   String�^   �����̓��e
'
'�O�����:
'   ������1���ȏ�L�^����Ă��邱��
'
    Dim str As String
    str = ""
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        str = str & "History[" & vbCrLf
        str = str & historyObj.ToString(4)
        str = str & "]" & vbCrLf
    Next historyObj
    ToString = str
End Function

Public Function CsvInformation() As String
    CsvInformation = ToCSVString()
End Function
Public Function Information() As String
    Information = ToString()
End Function

Public Sub ReportHistoriesWithMsgbox()
'���e:
'   ��������1�������b�Z�[�WBOX�֏o�͂���
'
'���ӎ���:
'   ���e�̕������������ꍇ�́A���ׂĕ\������Ȃ��ꍇ������܂�
'
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        MsgBox "Report:" & vbCrLf & historyObj.ToString()
    Next historyObj
End Sub
Public Sub ReportHistoriesWithDebugPrint()
'���e:
'   ��������1�����C�~�f�B�G�C�g�E�C���h�E�֏o�͂���
'
'���ӎ���:
'   ���e�̕������������ꍇ�́A���ׂĕ\������Ȃ��ꍇ������܂�
'
    Dim historyObj As CActionExecutionHistory
    For Each historyObj In mActionExecutionHistories
        Debug.Print "Report:" & vbCrLf & historyObj.ToString()
    Next historyObj
End Sub


Private Sub CheckInitializeCalled()
    If mTimer Is Nothing Then
        TheError.Raise 9999, TypeName(Me) & ".CheckInitializeCalled", "Call Initialize method first!"
    End If
End Sub

Public Function SaveHistoryLog(ByVal pFullFilePath As String) As Boolean
'���e:
'   �����̓��e��CSV�t�@�C���ɏo�́i�ǋL���[�h�j
'   �t�@�C�������݂���ꍇ�͐V�K�쐬
'
'[pFullFilePath] IN String�^:  �o�̓t�@�C���̖��O�i�t���p�X�j
'
'�߂�l:
'   Boolean�^   True:�L�^����    False:�L�^���s
'
'�O�����:
'   ������1���ȏ�L�^����Ă��邱��
'   �t�@�C���p�X��256�����ȓ��iDir�֐��̐���j
'
    If mActionExecutionHistories.Count > 0 Then
        SaveHistoryLog = WriteTxtFile(pFullFilePath, FileAndJobIndex & CSVIndex & ToCSVString)
    Else
        SaveHistoryLog = False
    End If
End Function

Private Function WriteTxtFile(ByVal pOutFilePath As String, ByRef pStr As String) As Boolean

    If Len(pOutFilePath) > 256 Then
        Call TheError.Raise(9999, TypeName(Me) & ".WriteTxtFile", "FilePath is too long (max 256):" & pOutFilePath)
    End If
    If Dir(pOutFilePath) = vbNullString Then
         If CreateTxtFile(pOutFilePath) = False Then
            Call TheError.Raise(9999, TypeName(Me) & ".WriteTxtFile", "Failed to create file:" & pOutFilePath)
         End If
    End If
    WriteTxtFile = AppendTxtFile(pOutFilePath, pStr)
End Function

Private Function CreateTxtFile(ByVal pOutFilePath As String) As Boolean

    Dim FileNo As Integer
    On Error GoTo catch
    FileNo = FreeFile
    Open pOutFilePath For Output As #FileNo
    Close #FileNo
    On Error GoTo 0

    CreateTxtFile = True
    Exit Function

catch:
    CreateTxtFile = False
    Call TheError.Raise(9999, TypeName(Me) & ".CreateTxtFile", "Output File Error:" & pOutFilePath)

End Function


Private Function AppendTxtFile(ByVal appendFileName As String, outPutMsg As String) As Boolean
'�w��TXT�t�@�C���ւ̕�����A�ǋL�o�͗p�̊֐�

    Dim fileNum As Integer
    Const FUNCTION_NAME = "AppendTxtFile"

    On Error GoTo OUT_PUT_LOG_ERR
    fileNum = FreeFile
    Open appendFileName For Append As fileNum
    Print #fileNum, outPutMsg
    Close fileNum
    AppendTxtFile = True
    On Error GoTo 0
    Exit Function

OUT_PUT_LOG_ERR:
    AppendTxtFile = False
    Call TheError.Raise(9999, FUNCTION_NAME, appendFileName & " Output File Error")

End Function





