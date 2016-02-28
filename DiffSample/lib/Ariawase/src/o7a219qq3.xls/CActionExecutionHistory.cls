VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CActionExecutionHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   Action�̎��s������1�������̏��
'
'�ړI:
'   Action�̎��s������1�������̏����i�[����f�[�^�I�u�W�F�N�g�B���L��ێ��B
'     �E�L�^�ԍ�
'     �E�L�^��
'     �E�L�^����
'     �E����
'     �E�g���K�[�iStart/End�j
'     �EActionInfo�^�I�u�W�F�N�g
'     �EError�ԍ��i0:Success�ANot 0:Error�j
'     �E���s�X�e�[�^�X�iCSV�ɂ͏o�͂���Ȃ��j��ԁF�����s�A�J�n�ς݁A���s�ς݁A�G���[
'
'�쐬��:
'   a_oshima
'
'�O�����:
'   CActionInfo�N���X�𗘗p
'   �O���[�o���I�u�W�F�N�g TheError:CErrManager �����݂��邱��
'
'���p�菇:
'


Option Explicit

Private mRecordNumber As String
Private mTrigger As String
Private mRecordDate As String
Private mRecordTime As String
Private mLapTime As Single
Private mActionInfo As CActionInfo
Private mErrorNo As Long

Property Let Trigger(ByVal pTrigger As String)
    mTrigger = pTrigger
End Property
Property Get Trigger() As String
    Trigger = mTrigger
End Property

Property Get RecordNumber() As String
    RecordNumber = mRecordNumber
End Property
Property Let RecordNumber(ByRef pId As String)
    mRecordNumber = pId
End Property

Property Get RecordDate() As String
    RecordDate = mRecordDate
End Property
Property Let RecordDate(ByRef pDay As String)
'�p�����[�^:
'   [pDay]  IN  String�^:  ���t��\��������B���t�łȂ��ꍇ�͗�O����

    If Not IsDate(pDay) Then
        Call TheError.Raise(9999, TypeName(Me) & ".RecordDate", "Not Date:" & pDay)
    End If
    mRecordDate = pDay
End Property

Property Get RecordTime() As String
    RecordTime = mRecordTime
End Property
Property Let RecordTime(ByRef pTim As String)
    mRecordTime = pTim
End Property

Property Get LapTime() As Single
    LapTime = mLapTime
End Property
Property Let LapTime(ByRef pTim As Single)
'�p�����[�^:
'   [pTim]  IN  Single�^:  �͈́F0�܂��͐��̒l�B���̏ꍇ�͗�O����
    
    If pTim < 0! Then
        Call TheError.Raise(9999, TypeName(Me) & ".LapTime", "Tim must 0 or positive value:" & pTim)
    End If
    mLapTime = pTim
End Property

Property Get ActionInfo() As CActionInfo
    Set ActionInfo = mActionInfo
End Property
Property Set ActionInfo(ByRef pActionInfo As CActionInfo)
    Set mActionInfo = pActionInfo
End Property

Property Get ErrorNo() As Long
    ErrorNo = mErrorNo
End Property

Property Let ErrorNo(ByRef pErrorNo As Long)
    mErrorNo = pErrorNo
End Property

Property Get ReturnCode() As String
    If mTrigger = "Start" Then
        ReturnCode = "-"
    Else
        ReturnCode = CStr(mErrorNo)
    End If
End Property

Private Sub Class_Initialize()
    Set mActionInfo = New CActionInfo
End Sub

Private Sub Class_Terminate()
    Set mActionInfo = Nothing
End Sub

Public Function CSVIndex(Optional ByVal ArgCount As Long = 0&) As String
'���e:
'   CSV�̃w�b�_��Ԃ��B
'
'�߂�l:
'   CSV�̃w�b�_�i�J���}��؂�̕�����B�Ō�ɉ��s�͕t�����Ȃ��j
'
'���l:
'   �e�l��MakeStringMsg�œ��ꕶ�����G�X�P�[�v�ς�
'
    Dim str As String
    str = MakeStringMsg("[RecordNumber]") _
            & "," & MakeStringMsg("[RecordDate]") _
            & "," & MakeStringMsg("[RecordTime]") _
            & "," & MakeStringMsg("[LapTime(s)]") _
            & "," & MakeStringMsg("[Trigger]") _
            & "," & mActionInfo.CSVIndex(ArgCount) _
            & "," & MakeStringMsg("[ReturnCode]")
            
    CSVIndex = str
End Function

Public Function ToCSVString(Optional ByVal ArgCount As Long = 0&) As String
'���e:
'   CSV�̃f�[�^��Ԃ��B
'
'�߂�l:
'   CSV�̃f�[�^�i�J���}��؂�̕�����B�Ō�ɉ��s�͕t�����Ȃ��j
'
'���l:
'   �e�l��MakeStringMsg�œ��ꕶ�����G�X�P�[�v�ς�
'
    Dim str As String
    str = MakeStringMsg(mRecordNumber) _
            & "," & MakeStringMsg(mRecordDate) _
            & "," & MakeStringMsg(mRecordTime) _
            & "," & mLapTime _
            & "," & MakeStringMsg(mTrigger) _
            & "," & mActionInfo.ToCSVString(ArgCount) _
            & "," & MakeStringMsg(ReturnCode())
    
    ToCSVString = str
End Function


Public Function ToString(Optional ByVal pIndentWidth As Long = 0&) As String
'���e:
'   �f�o�b�O�p�B�f�[�^�𕶎���ɏ����o���B
'
'�p�����[�^:
'   [pIndentWidth]  IN  Long�^:  �C���f���g���i�f�t�H���g�F0�j
'
'�߂�l:
'   ������ɐ��`�ς݂̃f�[�^
'
'���l:
'
    Dim str As String
    str = "RecordNumber:" & mRecordNumber & vbCrLf _
            & "RecordDate:" & mRecordDate & vbCrLf _
            & "RecordTime:" & mRecordTime & vbCrLf _
            & "LapTime(s):" & CStr(mLapTime) & vbCrLf _
            & "Trigger:" & mTrigger & vbCrLf _
            & "ActionInfo:[" & vbCrLf _
                & mActionInfo.ToString(4) _
            & "]" & vbCrLf _
            & "ReturnCode:" & ReturnCode() & vbCrLf

    If pIndentWidth > 0 Then
        ToString = Indent(str, pIndentWidth)
    Else
        ToString = str
    End If
End Function

Public Function CsvInformation(Optional ByVal ArgCount As Long = 0&) As String
'���e:
'   CSV�̃f�[�^��Ԃ��B
'
'�߂�l:
'   CSV�̃f�[�^�i�J���}��؂�̕�����B�Ō�ɉ��s��t���j
'
'���l:
'   �e�l��MakeStringMsg�œ��ꕶ�����G�X�P�[�v�ς�
'
    CsvInformation = ToCSVString(ArgCount) & vbCrLf
End Function
Public Function Information() As String
'���e:
'   �f�o�b�O�p�B�f�[�^�𕶎���ɏ����o���B
'
'�߂�l:
'   ������ɐ��`�ς݂̃f�[�^�i�Ō�ɕ������t���j
'
'���l:
'
    Information = ToString() & vbCrLf
End Function




Private Function Indent(ByVal pMessage As String, ByVal width As Long) As String
    Dim lines As Variant
    lines = Split(pMessage, vbCrLf)
    Dim i As Long
    Dim Msg As String
    For i = LBound(lines) To UBound(lines) - 1
        Msg = Msg & Space$(width) & lines(i) & vbCrLf
    Next i
    If lines(UBound(lines)) <> vbNullString Then
        Msg = Msg & Space$(width) & lines(UBound(lines))
    End If
    Indent = Msg
End Function


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
        makeMsg = Chr$(34) & makeMsg & Chr$(34)
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



