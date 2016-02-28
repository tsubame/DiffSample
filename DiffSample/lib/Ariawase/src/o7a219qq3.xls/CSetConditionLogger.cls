VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetConditionLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ��������ݒ胍�O�񋟋@�\�N���X
'
'�ړI:
'   ��������ݒ�}�N���̓��샍�O��o�^�ۑ�����
'

Option Explicit

Private m_ExecuteLogs As Collection           '�ݒ藚��ۑ��p�R���N�V����
Private m_MaxArgNumber As Long                'Arg�p�����[�^�̍ő�ԍ�
Private Const ERROR_NUMBER As Long = 9999     '�G���[�������ɕԂ��G���[�ԍ�

Public Property Let MaxArgNumber(ByVal pArgNumber As Long)
'���e:
'   Arg�ԍ��̍ő�l
'
'�p�����[�^:
'   [pArgNumber]  In  Long�^: Arg�ԍ��ő�l
'
'�߂�l:
'
'���ӎ���:
'
    
    m_MaxArgNumber = pArgNumber

End Property

Public Property Get MaxArgNumber() As Long
'���e:
'   Arg�ԍ��̍ő�l
'
'�p�����[�^:
'
'�߂�l:
'   Arg�ԍ��ő�l
'
'���ӎ���:
'
    
    MaxArgNumber = m_MaxArgNumber

End Property

Public Sub AddRecord(ByVal pExecTime As Double, ByRef pMacroInfo As CSetFunctionInfo)
'���e:
'   ���O���̒ǉ��o�^
'
'�p�����[�^:
'   [pExecTime]     In  Double�^�F�}�N���̎��s����[sec]
'   [pMacroInfo]    In  CSetFunctionInfo�^�F�}�N�����Object
'
'�߂�l:
'
'���ӎ���:
'

    Dim SetHistoryObj As CSetConditionHistory
    Set SetHistoryObj = New CSetConditionHistory
       
    With SetHistoryObj
        .RecordDate = CStr(Date)
        .RecordTime = CStr(TIME)
        .ExecTestInstanceName = TheExec.DataManager.InstanceName   '�Z�b�g�֐����s���̃e�X�g�C���X�^���X���L�^
        .ExecTime = pExecTime                                      '�Z�b�g�֐��̎��s���ԋL�^
        Set .SetFunctionInfo = pMacroInfo.Clone                    '�Z�b�g�֐��̏��̃R�s�[���L�^
    End With

    Call m_ExecuteLogs.Add(SetHistoryObj)

End Sub

Public Sub Initialize()
'���e:
'   �@�\������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    
    Set m_ExecuteLogs = New Collection

End Sub

Public Function Record(ByVal pRecordNumber As Long) As CSetConditionHistory
'���e:
'   ����ԍ��w��ɂ�郍�O���̌��J
'
'�p�����[�^:
'   [pRecordNumber]     In  Long�^�F����ԍ�
'
'�߂�l:
'   �����ݒ藚�����(CSetConditionHistory)
'
'���ӎ���:
'   �w��ԍ��̗�����񂪑��݂��Ȃ��ꍇ�́A�߂�l�� Nothing �ɂȂ�܂��B
'
    
    If m_ExecuteLogs.Count > 0 Then
        Set Record = m_ExecuteLogs.Item(pRecordNumber)
    Else
        Set Record = Nothing
    End If

End Function

Public Property Get RecordCount() As Long
'���e:
'   �L�^�Ϗ��v�f�����J
'
'�p�����[�^:
'
'�߂�l:
'   �L�^�Ϗ��v�f��
'
'���ӎ���:
'
    
    RecordCount = m_ExecuteLogs.Count

End Property

Public Sub ClearRecord()
'���e:
'   �L�^����Ă���S�Ă̏��̏���
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    
    Set m_ExecuteLogs = New Collection

End Sub

Public Sub SaveRecord(ByVal pOutputFileName As String)
'���e:
'   �L�^����Ă��郍�O��TXT�t�@�C���֒ǋL����
'
'�p�����[�^:
'   [pOutputFileName]     In  String�^�F�o�̓t�@�C����(FullPath)
'
'�߂�l:
'
'���ӎ���:
'
    
    If m_ExecuteLogs.Count = 0 Then
        Call AppendTxtFile(pOutputFileName, "There is no SetCondition history")
        Exit Sub
    End If
    
    '�f�[�^�o��LOOP
    Dim RecordCnt As Long
    For RecordCnt = 1 To m_ExecuteLogs.Count
        On Error GoTo SAVE_HISTORY_LOG_ERR
        Call AppendTxtFile(pOutputFileName, MakeCsvHistoryInfo(RecordCnt))
        On Error GoTo 0
    Next RecordCnt
    
    Exit Sub

SAVE_HISTORY_LOG_ERR:
'�ݒ藚�����O�̃t�@�C���o�͎��̃G���[����
    Call RaiseError(pOutputFileName & " write error" & vbCrLf & _
    "The output to the file of the history log was not done")

End Sub

Private Function MakeCsvHistoryInfo(ByVal pHistoryNumber As Long) As String
'���e:
'   '�����ݒ藚������CSV Format�쐬�֐�
'
'�p�����[�^:
'   [pHistoryNumber]�@In Long�^�F  �擾���闚��ԍ��̎w��
'
'�߂�l:
'   �o�͂�����s���O
'���ӎ���:
'

    Dim makeMsg As String
    Dim MsgHeader As String
    Dim ArrayNumber As Long
    Dim HistoryInfo As CSetConditionHistory
    
    On Error GoTo GET_HISTORY_ERR
    Set HistoryInfo = Me.Record(pHistoryNumber)
    On Error GoTo 0
    
    Const MAX_ARGCOUNT As Long = 10 'Arg�̍ő��
    
    'CSV�`�����쐬����
    With HistoryInfo
        makeMsg = ""

        makeMsg = makeMsg & _
        pHistoryNumber & "," & _
        "" & .RecordDate & "" & "," & _
        "" & .RecordTime & "" & "," & _
        "" & .ExecTestInstanceName & "" & "," & _
        "" & MakeStringMsg(.SetFunctionInfo.ConditionName) & "" & "," & _
        "" & MakeStringMsg(.SetFunctionInfo.FunctionName) & "" & "," & _
        "" & MacroSetting2String(.SetFunctionInfo.IsEnableMacro) & "" & "," & _
            .ExecTime
        '�p�����[�^���쐬
        Dim ArgCount As Long
        ArgCount = 0
        For ArrayNumber = 0 To .SetFunctionInfo.ArgParameterCount - 1 Step 1
            makeMsg = makeMsg & "," & MakeStringMsg(.SetFunctionInfo.Arg(ArrayNumber))
            ArgCount = ArgCount + 1
        Next ArrayNumber
        
        '�}�N���O���[�v���x���܂�
        Dim ToExecSettingName As Long
        For ToExecSettingName = 1 To (MAX_ARGCOUNT - ArgCount) Step 1
            makeMsg = makeMsg & ","
        Next ToExecSettingName
        makeMsg = makeMsg & "," & .SetFunctionInfo.ExecuteSettingName
    End With
    
    If pHistoryNumber = 1 Then
        '�w�b�_��񂠂�
        Dim RetFileName   As String
        Dim RetJobName   As String
        RetFileName = TheExec.ExcelHandle.ActiveWorkbook.Path & "\" & TheExec.ExcelHandle.ActiveWorkbook.Name
        RetJobName = TheExec.CurrentJob                             'JobList �� JobName ���B������΃t�@�C����
        MsgHeader = "[Test Program],[" & RetFileName & "]" & vbCrLf & _
                    "[Job],[" & RetJobName & "]" & vbCrLf & _
                    "[RecordNumber],[RecordDate],[RecordTime],[TestInstanceName],[ConditionName],[ProcedureName],[Execute],[ExecTime(s)]"
        Dim ArgNumber As Long
        For ArgNumber = 0 To m_MaxArgNumber Step 1
            MsgHeader = MsgHeader & "," & "[Arg" & ArgNumber & "]"
        Next ArgNumber
        MsgHeader = MsgHeader & "," & "[ExecuteSetting]"
        MakeCsvHistoryInfo = MsgHeader & vbCrLf & makeMsg
    Else
        '�w�b�_���Ȃ�
        MakeCsvHistoryInfo = makeMsg
    End If
    
    Exit Function

GET_HISTORY_ERR:
'�����f�[�^���o���G���[����
    Call RaiseError("HistoryNo=" & pHistoryNumber & " data is nothing" & vbCrLf & _
    "As for the number, there is the need that is a value from 1" & " to " & Me.RecordCount)

End Function

'�e�L�X�g�t�@�C���ɒǋL�o�͂���ׂ̊֐�
Private Function AppendTxtFile(ByVal pTgtFileName As String, ByVal pOutPutMsg As String) As Boolean
    
    Dim fileNum As Integer
    Dim errFunctionName As String
    
    On Error GoTo OUT_PUT_TXT_ERR
    
    fileNum = FreeFile
    Open pTgtFileName For Append As fileNum
    Print #fileNum, pOutPutMsg
    Close fileNum
    
    AppendTxtFile = True
    
    Exit Function

OUT_PUT_TXT_ERR:
    Close fileNum
    Call RaiseError(pTgtFileName & " file write error")
    AppendTxtFile = False

End Function

Private Function MakeStringMsg(ByVal pMessage As String) As String
'CSV�̕�����v�f���쐬����֐�
    Dim makeMsg As String
        
    If Len(pMessage) > 0 Then
        '1:CR�ALF�ACR+LF���̉��s������u��������
        '2:������̃J���}�� \, �ɒu��������
        '3:������̃_�u���R�[�e�[�V������ ""�ɒu��������
        makeMsg = ReplaceDoubleQuotation(ReplaceComma(ReplaceCrLf(pMessage)))
        
        '������p�����[�^�̒�����1�ȏ�̏ꍇ�́A��������_�u���R�[�e�[�V�����ł�����
        makeMsg = Chr(34) & makeMsg & Chr(34)
        MakeStringMsg = makeMsg
    
    Else
        MakeStringMsg = pMessage
    End If
    
End Function

Private Function ReplaceCrLf(ByVal pMessage As String) As String
'������Ɋ܂܂�Ă�����s������u�������ĕԂ��֐�
'vbCrLf-->\r\n ,vbCr --> \r, vbLf --> \n
    Dim makeMsg As String
    makeMsg = Replace(Replace(Replace(pMessage, vbCrLf, "\r\n"), vbCr, "\r"), vbLf, "\n")
    ReplaceCrLf = makeMsg
End Function

Private Function ReplaceDoubleQuotation(ByVal pMessage As String) As String
'������Ɋ܂܂�Ă���DoubleQuotation��u�������ĕԂ��֐�
'" --> ""
    Dim makeMsg As String
    makeMsg = Replace(pMessage, """", """" & """")
    ReplaceDoubleQuotation = makeMsg
End Function

Private Function ReplaceComma(ByVal pMessage As String) As String
'������Ɋ܂܂�Ă���Comma��u�������ĕԂ��֐�
', --> \,
    Dim makeMsg As String
    makeMsg = Replace(pMessage, ",", "\,")
    ReplaceComma = makeMsg
End Function

'�G���[���C�Y�p�֐�
Private Sub RaiseError(ByVal pErrMsg As String)
    If Not (TheError Is Nothing) Then
        '�G���[�}�l�[�W����
        Call TheError.Raise(ERROR_NUMBER, TypeName(Me), pErrMsg)
    Else
        'VB��Err��
        Call Err.Raise(ERROR_NUMBER, TypeName(Me), pErrMsg)
    End If
End Sub

'�}�N���̗L�������ݒ肩�當�����Ԃ�
Private Function MacroSetting2String(ByVal pIsEnable As Boolean) As String
    If pIsEnable = True Then
        MacroSetting2String = "enable"
    Else
        MacroSetting2String = "disable"
    End If
End Function

