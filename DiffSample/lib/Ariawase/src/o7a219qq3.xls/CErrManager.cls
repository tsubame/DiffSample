VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CErrManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �G���[���Ǘ��p�N���X
'
'�ړI:
'   �G���[�����W�߂�
'   �G���[���̗��������W����
'
'�쐬��:
'   SLSI����
'
'���ӓ_:
'   �G���[���̗����͈ȉ��̂悤�ȏꍇ��������܂�
'   �E�������������� ClearHistory�����s
'   �EVB�̃��Z�b�g�{�^��������
'   �E�G���[�}�l�[�W��Object��p������
'   �G���[���N���X(CErrInfo)���K�v�ƂȂ�܂�
'
Option Explicit

'�G���[���~�σR���N�V����
Private m_ErrInfoStorage As Collection
'�G���[���������݂��Ȃ��Ƃ��ɕ񍐗v�����ꂽ�Ƃ��ɕԂ����b�Z�[�W
Private Const NO_HISTORY_MSG = "There is no error history"

'#Pass-OK
Private Sub Class_Initialize()
'�����������i�������j
    Me.ClearHistory
End Sub

'#Pass-OK
Private Sub Class_Terminate()
'�p��������
    Set m_ErrInfoStorage = Nothing
End Sub

'#Pass-OK
Public Property Get ErrorHistory(ByVal HistoryNumber As Long) As CErrInfo
'���e:
'   �w�藚��ԍ��̃G���[���Object��񋟂���
'
'�p�����[�^:
'   [HistoryNumber]  In  ����ԍ�
'
'�߂�l:
'   �w��ԍ��̃G���[���Object
'
'���ӎ���:
'   �w��ԍ��̃G���[�����������Ƃ��߂�l�́ANothing�ƂȂ�܂��B
'
    With m_ErrInfoStorage
        If (HistoryNumber > 0) And (HistoryNumber <= .Count) Then
            Set ErrorHistory = .Item(HistoryNumber)
        Else
            Set ErrorHistory = Nothing
        End If
    End With

End Property

'#Pass-OK
Public Property Get LastError() As CErrInfo
'���e:
'   �ŏI�i�ŐV�j�G���[���Object�̌��J
'
'�p�����[�^:
'
'�߂�l:
'   �ŏI�i�ŐV�j�G���[���object
'
'���ӎ���:
'   �G���[���������݂��Ȃ����߂�l�́ANothing�ɂȂ�܂��B
'
    With m_ErrInfoStorage
        If .Count > 0 Then
            Set LastError = .Item(.Count) '���������݂���Ƃ��Ō�̃G���[�������J
        Else
            Set LastError = Nothing       '���������݂��Ȃ��Ƃ���Nothing��Ԃ�
        End If
    End With

End Property

'#Pass-OK
Public Property Get ErrorCount() As Long
'���e:
'   �ۑ�����Ă���G���[���̐������J����
'
'�p�����[�^:
'
'�߂�l:
'   �ۑ�����Ă���G���[���
'
'���ӎ���:
'   �G���[��񂪕ۑ�����Ă��Ȃ��Ƃ��̒l��0�ƂȂ�܂�
'
    ErrorCount = m_ErrInfoStorage.Count

End Property

'#Pass-OK
Public Sub ClearHistory()
'���e:
'   �G���[�������̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Set m_ErrInfoStorage = New Collection

End Sub

'#Pass-OK
Public Sub Raise(ByVal ErrNumber As Long, ByVal ErrSource As String, ByVal errMsg As String, _
Optional ByVal HelpFile As String = "", Optional ByVal HelpContext As Long = 0)
'���e:
'   �G���[�̃��C�Y
'
'�p�����[�^:
'   [ErrNumber]  In  �G���[�ԍ�
'   [ErrSource]  In  �����\�[�X����
'   [ErrMsg]     In  ���b�Z�[�W������
'   [HelpFile]   In  �G���[���ɕ\������w���v�t�@�C�����i�I�v�V�����j
'   [HelpContext]  In �G���[���ɕ\������w���v�t�@�C���̃w���vID�ԍ��i�I�v�V�����j
'
'�߂�l:
'
'���ӎ���:
'   �{���߂����s����ƃG���[���̋L�^��A
'   ���͂��ꂽ�p�����[�^���g�p����VBA��Err.Raise�����s����܂��B
'   �G���[���ɕ\������w���v�t�@�C�����i�I�v�V�����j��Full Path�Ŏw�肵�ĉ�����
'
    Dim tmpErr As CErrInfo
    Set tmpErr = New CErrInfo

    '�����G���[Object�ɐݒ�
    With tmpErr
        .SerialNumber = m_ErrInfoStorage.Count + 1
        .Number = ErrNumber
        .Source = ErrSource
        .Message = errMsg
        .HelpFile = HelpFile
        .HelpContext = HelpContext
        .ErrDate = Date
        .ErrTime = TIME
    End With

    '�G���[�����̓o�^
    Call m_ErrInfoStorage.Add(tmpErr)  '�G���[�����ɍ���̃G���[����ǉ�
        
    'VBA��ErrObject�ɃG���[��������
    With tmpErr
        Call Err.Raise(.Number, .Source, .Message, .HelpFile, .HelpContext)
    End With

End Sub

'#Pass-OK
Public Sub ReportDebugPrint()
'���e:
'   �G���[���������C�~�f�B�G�C�gWindow�֏o�͂���
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Dim retErrInfo As CErrInfo

    If m_ErrInfoStorage.Count > 0# Then
        '����Record������Ƃ��̃��b�Z�[�W�쐬
        For Each retErrInfo In m_ErrInfoStorage
            With retErrInfo
                .DebugPrint
            End With
        Next retErrInfo
    Else
        '����Record�������Ƃ��̃��b�Z�[�W
        Debug.Print NO_HISTORY_MSG
    End If

End Sub

'#Pass-OK
Public Sub ReportMsgBox()
'���e:
'   �G���[�����������b�Z�[�WBOX�֏o�͂���
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'   �G���[�����̐��������ꍇ�́A���ׂẴG���[���\������Ȃ��ꍇ������܂�
'
    Call MsgBox(MakeErrHistoryMsg, vbInformation, "JOB error history")

End Sub

'#Pass-OK
Public Sub ReportTxtFile(ByVal fileName As String)
'���e:
'   �G���[���������w��TXT�t�@�C���֒ǋL�o�͂���
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'   �o�͐�̃t�@�C���́AFullPath�Ŏw�肵�ĉ�����
'   �w��t�@�C�������łɑ��݂��Ă���ꍇ�ɂ́A���t�@�C���ɒǋL�o�͂���܂��B
'   ���́ACSV�`���ŏo�͂���܂��B
'
    Call AppendTxtFile(fileName, MakeCsvErrHistoryMsg)

End Sub

'#Pass-OK
Private Function MakeErrHistoryMsg() As String
'��񃁃b�Z�[�W�쐬�p�֐�(MsgBox�p)

    Dim retErrInfo As CErrInfo
    Dim PrintMsg As String
    
    If m_ErrInfoStorage.Count > 0# Then
        '����Record������Ƃ��̃��b�Z�[�W�쐬
        PrintMsg = ""
        For Each retErrInfo In m_ErrInfoStorage
            With retErrInfo
                PrintMsg = PrintMsg & .Information & vbCrLf
            End With
        Next retErrInfo
        MakeErrHistoryMsg = PrintMsg
    Else
        '����Record�������Ƃ��̃��b�Z�[�W
        MakeErrHistoryMsg = NO_HISTORY_MSG
    End If

End Function

'#Pass-OK
Private Function MakeCsvErrHistoryMsg() As String
'CSV�`���̃G���[���b�Z�[�W�쐬�p�֐�

    Dim retErrInfo As CErrInfo
    Dim PrintMsg As String
        
    If m_ErrInfoStorage.Count > 0# Then
        '����Record������Ƃ��̃��b�Z�[�W�쐬
        PrintMsg = ""
        For Each retErrInfo In m_ErrInfoStorage
            With retErrInfo
                PrintMsg = PrintMsg & .CsvInformation
            End With
        Next retErrInfo
        MakeCsvErrHistoryMsg = PrintMsg
    Else
        '����Record�������Ƃ��̃��b�Z�[�W
        MakeCsvErrHistoryMsg = NO_HISTORY_MSG & vbCrLf
    End If

End Function

'#Pass-OK
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
    Call Err.Raise(9999, FUNCTION_NAME, appendFileName & " Output File Error")

End Function
