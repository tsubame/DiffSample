VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CErrInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �G���[���\���̗p�N���X
'
'�ړI:
'   �G���[���̗v�f���܂Ƃ߂�
'   �G���[���Ə��񋟋@�\�����킹�ĊǗ�����
'
'�쐬��:
'   SLSI����
'
'���ӓ_:
'   �G���[�}�l�[�W��Object���{�N���X�𗘗p���܂��B
'
Option Explicit

Private m_SerialNumber As Long   '�G���[�L�^�V���A���ԍ�
Private m_Number As Long         '�G���[�ԍ�
Private m_Source As String       '�G���[�\�[�X
Private m_Message As String      '�G���[���b�Z�[�W
Private m_HelpFile As String     '�G���[���ɕ\������w���v�t�@�C��
Private m_HelpContext As Long '�w���v�t�@�C���̃g�s�b�N��\���R���e�L�X�gID���܂ސ���
Private m_Date As String         '�G���[������
Private m_Time As String         '�G���[��������
Private Const ThisClassName = "CErrInfo" '�N���X����
'#Pass-OK
Public Property Let SerialNumber(ByVal ErrRecordNumber As Long)
'���e:
'   �G���[���V���A���ԍ��ݒ�
'
'�p�����[�^:
'   [ErrRecordNumber]  In  �G���[�L�^�ԍ�
'
'�߂�l:
'
'���ӎ���:
'
    m_SerialNumber = ErrRecordNumber

End Property

'#Pass-OK
Public Property Get SerialNumber() As Long
'���e:
'   �G���[���V���A���ԍ��擾
'
'�p�����[�^:
'
'�߂�l:
'   �G���[�L�^���̃V���A���ԍ�
'
'���ӎ���:
'
    SerialNumber = m_SerialNumber

End Property

'#Pass-OK
Public Property Let ErrDate(ByVal ErrDate As String)
'���e:
'   �G���[�������ݒ�
'
'�p�����[�^:
'   [ErrDate]  In  �G���[������������
'
'�߂�l:
'
'���ӎ���:
'
    m_Date = ErrDate

End Property

'#Pass-OK
Public Property Get ErrDate() As String
'���e:
'   �G���[�������擾
'
'�p�����[�^:
'
'�߂�l:
'   �G���[������������
'
'���ӎ���:
'
    ErrDate = m_Date

End Property

'#Pass-OK
Public Property Let ErrTime(ByVal ErrTime As String)
'���e:
'   �G���[�������Ԑݒ�
'
'�p�����[�^:
'   [ErrTime]  In  �G���[�������ԕ�����
'
'�߂�l:
'
'���ӎ���:
'
    m_Time = ErrTime

End Property

'#Pass-OK
Public Property Get ErrTime() As String
'���e:
'   �G���[�������Ԏ擾
'
'�p�����[�^:
'
'�߂�l:
'   �G���[�������ԕ�����
'
'���ӎ���:
'
    ErrTime = m_Time

End Property

'#Pass-OK
Public Property Let HelpContext(ByVal HelpContextNumber As Long)
'���e:
'   �G���[���w���v�t�@�C���̃w���vID�ݒ�
'
'�p�����[�^:
'   [HelpContextNumber]  In  �w���vID�ԍ�
'
'�߂�l:
'
'���ӎ���:
'
    m_HelpContext = HelpContextNumber

End Property

'#Pass-OK
Public Property Get HelpContext() As Long
'���e:
'   �G���[���w���v�t�@�C���̃w���vID�擾
'
'�p�����[�^:
'
'�߂�l:
'   �w���vID�ԍ�
'
'���ӎ���:
'
    HelpContext = m_HelpContext

End Property

'#Pass-OK
Public Property Let HelpFile(ByVal HelpFile As String)
'���e:
'   �G���[���ɕ\������w���v�t�@�C�����ݒ�
'
'�p�����[�^:
'   [HelpFile]  In  �w���v�t�@�C����������
'
'�߂�l:
'
'���ӎ���:
'   �w���v�t�@�C�����́AFullPath�Őݒ肵�܂�
'   (��) d:\HelpFiles\JobErrHelp.mht
'
    m_HelpFile = HelpFile

End Property

'#Pass-OK
Public Property Get HelpFile() As String
'���e:
'   �G���[���ɕ\������w���v�t�@�C�����擾
'
'�p�����[�^:
'
'�߂�l:
'   �w���v�t�@�C��������
'
'���ӎ���:
'
    HelpFile = m_HelpFile

End Property

'#Pass-OK
Public Property Let Number(ByVal ErrNumber As Long)
'���e:
'   �G���[�ԍ��ݒ�
'
'�p�����[�^:
'   [ErrNumber]  In  �G���[�ԍ��l
'
'�߂�l:
'
'���ӎ���:
'   �G���[�ԍ��ɂ́A65535�ȉ��̒l��ݒ肵�ĉ������B
'
    
    If ErrNumber > 65535 Then
        Call Err.Raise(9999, ThisClassName, _
        "error number=" & ErrNumber & " is a value that cannot be set.  [error number <= 65535]")
    End If
    
    m_Number = ErrNumber

End Property

'#Pass-OK
Public Property Get Number() As Long
'���e:
'   �G���[�ԍ��擾
'
'�p�����[�^:
'
'�߂�l:
'   �G���[�ԍ��l
'
'���ӎ���:
'
    Number = m_Number

End Property

'#Pass-OK
Public Property Let Message(ByVal ErrorMessage As String)
'���e:
'   �G���[���b�Z�[�W
'
'�p�����[�^:
'   [ErrorMessage]  In  �G���[���̃��b�Z�[�W
'
'�߂�l:
'
'���ӎ���:
'
    m_Message = ErrorMessage

End Property

'#Pass-OK
Public Property Get Message() As String
'���e:
'   �G���[���b�Z�[�W
'
'�p�����[�^:
'
'�߂�l:
'   �G���[���̃��b�Z�[�W
'
'���ӎ���:
'
    Message = m_Message

End Property

'#Pass-OK
Public Property Let Source(ByVal ErrSource As String)
'���e:
'   �G���[�\�[�X��
'
'�p�����[�^:
'   [ErrSource]  In  �G���[�\�[�X����
'
'�߂�l:
'
'���ӎ���:
'
    m_Source = ErrSource

End Property

'#Pass-OK
Public Property Get Source() As String
'���e:
'   �G���[�\�[�X��
'
'�p�����[�^:
'
'�߂�l:
'   �G���[�\�[�X����
'
'���ӎ���:
'
    Source = m_Source

End Property

'#Pass-OK
Public Property Get CsvInformation() As String
'���e:
'   �G���[����CSV�i�J���}��؂�`���j������Œ񋟂���
'
'�p�����[�^:
'
'�߂�l:
'   CSV�`���G���[���
'
'���ӎ���:
'  �o��Format�d�l
'   [SerialNumber],"[Date]","[Time]",[Number],"[Source]","[Message]","[HelpFile]",[HelpContext]
'
'    �����s�����A�J���}�A�_�u���R�[�e�[�V�����̓G�X�P�[�v�����ɒu���������܂��B
'      vbCrLf-->\r\n ,vbCr --> \r, vbLf --> \n
'      , --> \,
'      " --> ""
'
    Dim CsvMsg As String
            
    With Me
        CsvMsg = .SerialNumber & "," & _
         MakeStringMsg(.ErrDate) & "," & _
         MakeStringMsg(.ErrTime) & "," & _
        .Number & "," & _
         MakeStringMsg(.Source) & "," & _
         MakeStringMsg(.Message) & "," & _
         MakeStringMsg(.HelpFile) & "," & _
        .HelpContext & vbCrLf
    End With
    
    CsvInformation = CsvMsg

End Property

'#Pass-OK
Public Property Get Information() As String
'���e:
'   �G���[���񋟗p�֐�
'
'�p�����[�^:
'
'�߂�l:
'   �G���[���\���p���b�Z�[�W
'
'���ӎ���:
'   �o�͂����v�f��
'   �����A�V���A���ԍ��A�G���[�ԍ��A�G���[�\�[�X�A�G���[���b�Z�[�W�ƂȂ�܂�
'
    Dim PrintMsg As String
    
    '�G���[���b�Z�[�W�쐬
    With Me
        PrintMsg = "[Date] " & .ErrDate & " " & .ErrTime & vbCrLf & _
        "[Serial] " & .SerialNumber & vbCrLf & _
        "[ErrNum] " & .Number & vbCrLf & _
        "[Src] " & MakeStringMsg(.Source) & vbCrLf & _
        "[Msg] " & MakeStringMsg(.Message) & vbCrLf
    End With
    
    Information = PrintMsg

End Property

'#Pass-OK
Public Sub DebugPrint()
'���e:
'   �G���[�����C�~�f�B�G�C�gWindow�ɏo�͂���
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Debug.Print Me.Information

End Sub

'#Pass-OK
Public Sub DispMsgBox()
'���e:
'   �G���[�������b�Z�[�WBOX�ɏo�͂���
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Call MsgBox(Me.Information, vbInformation, "Error information")

End Sub

'-- �ȉ��APrivate�@�\�Q -------------------------------------------------------

'#Pass-OK
Private Sub Class_Initialize()
'constructor�F�����f�[�^������
    m_SerialNumber = 0
    m_Number = 0
    m_Source = ""
    m_Message = ""
    m_HelpFile = ""
    m_HelpContext = 0
    m_Date = ""
    m_Time = ""
End Sub

'#Pass-OK
Private Function MakeStringMsg(Message As String) As String
'CSV�̕�����v�f���쐬����֐�
    Dim makeMsg As String
        
    If Len(Message) > 0 Then
        '1:CR�ALF�ACR+LF���̉��s������u��������
        '2:������̃J���}�� \, �ɒu��������
        '3:������̃_�u���R�[�e�[�V������ ""�ɒu��������
        makeMsg = ReplaceDoubleQuotation(ReplaceComma(ReplaceCrLf(Message)))
        
        '������p�����[�^�̒�����1�ȏ�̏ꍇ�́A��������_�u���R�[�e�[�V�����ł�����
        makeMsg = Chr(34) & makeMsg & Chr(34)
        MakeStringMsg = makeMsg
    
    Else
        MakeStringMsg = Message
    End If
    
End Function

'#Pass-OK
Private Function ReplaceCrLf(Message As String) As String
'������Ɋ܂܂�Ă�����s������u�������ĕԂ��֐�
'vbCrLf-->\r\n ,vbCr --> \r, vbLf --> \n
    Dim makeMsg As String
    makeMsg = Replace(Replace(Replace(Message, vbCrLf, "\r\n"), vbCr, "\r"), vbLf, "\n")
    ReplaceCrLf = makeMsg
End Function

'#Pass-OK
Private Function ReplaceDoubleQuotation(Message As String) As String
'������Ɋ܂܂�Ă���DoubleQuotation��u�������ĕԂ��֐�
'" --> ""
    Dim makeMsg As String
    makeMsg = Replace(Message, """", """" & """")
    ReplaceDoubleQuotation = makeMsg
End Function

'#Pass-OK
Private Function ReplaceComma(Message As String) As String
'������Ɋ܂܂�Ă���Comma��u�������ĕԂ��֐�
', --> \,
    Dim makeMsg As String
    makeMsg = Replace(Message, ",", "\,")
    ReplaceComma = makeMsg
End Function

'Private Function ReplaceBackSlash(Message As String) As String
''������Ɋ܂܂�Ă���Backslash��u�������ĕԂ��֐�(���g�p)
''\ --> \\
'    Dim makeMsg As String
'    makeMsg = Replace(Message, "\", "\\")
'    ReplaceBackSlash = makeMsg
'End Function


