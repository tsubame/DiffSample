VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CActionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   Action�̏��
'
'�ړI:
'   Action�̏����i�[����f�[�^�I�u�W�F�N�g�B���L�������B
'     �E�e�X�g�C���X�^���X��
'     �E�e�X�g�v���V�[�W�����iUserMacro�j
'     �EAction���i�t���[�����[�N���j
'     �EAcquireGroup��
'     �EAcquire�C���X�^���X��
'     �ECAcquireFunctionInfo�^�I�u�W�F�N�g�i�}�N���̃p�����[�^��ێ��j
'
'�쐬��:
'   a_oshima
'
'�O�����:
'   CAcquireFunctionInfo�N���X�𗘗p
'   CAcquireFunctionInfo�N���X�̉��L���\�b�h�𗘗p
'      �EPublic Sub CopyObj(OrgObj As CAcquireFunctionInfo)
'�@�@�@�EPublic Property Get ArgParameterCount() As Long
'�@�@�@�EPublic Property Get Arg(ByVal pArgNumber As Long) As Variant
'
'���p�菇:



Option Explicit

Private mTestInstanceName As String
Private mMacroName As String
Private mFrameworkName As String
Private mAcquireGroupName As String
Private mAcquireInstanceName As String
Private mParameters As CAcquireFunctionInfo




Property Get TestInstanceName() As String
    TestInstanceName = mTestInstanceName
End Property
Property Let TestInstanceName(ByRef pName As String)
    mTestInstanceName = pName
End Property

Property Get FrameworkName() As String
    FrameworkName = mFrameworkName
End Property
Property Let FrameworkName(ByRef pName As String)
    mFrameworkName = pName
End Property

Property Get AcquireGroupName() As String
    AcquireGroupName = mAcquireGroupName
End Property
Property Let AcquireGroupName(ByRef pName As String)
    mAcquireGroupName = pName
End Property

Property Get AcquireInstanceName() As String
    AcquireInstanceName = mAcquireInstanceName
End Property
Property Let AcquireInstanceName(ByRef pName As String)
    mAcquireInstanceName = pName
End Property

Property Get macroName() As String
    macroName = mMacroName
End Property
Property Let macroName(ByRef pName As String)
    mMacroName = pName
End Property

Property Get Parameters() As CAcquireFunctionInfo
    Set Parameters = mParameters
End Property
Property Set Parameters(ByRef pParameters As CAcquireFunctionInfo)
'    '�R�s�[���i�[
'    Set mParameters = New CAcquireFunctionInfo
'    Call mParameters.CopyObj(pParameters)
    '�Q�ƌ����i�[
    Set mParameters = pParameters
End Property

Private Sub Class_Initialize()
    Set mParameters = New CAcquireFunctionInfo
End Sub

Private Sub Class_Terminate()
    Set mParameters = Nothing
End Sub

Public Function Clone() As CActionInfo
'���e:
'   ���I�u�W�F�N�g�̕�����Ԃ��iDeep Copy�j
'
'�߂�l:
'   ���I�u�W�F�N�g�̕���
'
'���l:
'
'
    Dim objCopy As CActionInfo
    Set objCopy = New CActionInfo
    objCopy.TestInstanceName = mTestInstanceName
    objCopy.FrameworkName = mFrameworkName
    objCopy.AcquireGroupName = mAcquireGroupName
    objCopy.AcquireInstanceName = mAcquireInstanceName
    objCopy.macroName = mMacroName
    Set objCopy.Parameters = mParameters.Clone
    
    Set Clone = objCopy
End Function

Public Function IsSameName(ByRef pTargetActionInfo As CActionInfo) As Boolean
'���e:
'   ���^�I�u�W�F�N�g�ƊeProperty�����ꂼ���r����
'   �i�p�����[�^�̏��͔�r���Ȃ��j
'
'�p�����[�^:
'   [pTargetActionInfo]  IN  CActionInfo�^:  ��r�Ώۂ�ActionInfo
'
'�߂�l:
'   True�F�p�����[�^�������SProperty�������AFalse�FProperty���P�ł��قȂ�
'
'���l:
'
'
    If mTestInstanceName = pTargetActionInfo.TestInstanceName _
            And mFrameworkName = pTargetActionInfo.FrameworkName _
            And mAcquireGroupName = pTargetActionInfo.AcquireGroupName _
            And mAcquireInstanceName = pTargetActionInfo.AcquireInstanceName _
            And mMacroName = pTargetActionInfo.macroName Then
        IsSameName = True
    Else
        IsSameName = False
    End If
End Function

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
    str = MakeStringMsg("[TestInstanceName]") _
            & "," & MakeStringMsg("[TestProcedureName]") _
            & "," & MakeStringMsg("[ActionName]") _
            & "," & MakeStringMsg("[GroupName]") _
            & "," & MakeStringMsg("[InstanceName]")

    Dim loopEnd As Long
    loopEnd = ArgCount - 1
    If ArgCount < mParameters.ArgParameterCount Then
        loopEnd = mParameters.ArgParameterCount - 1
    End If
    
    Dim i As Long
    For i = 0& To loopEnd
        str = str & "," & MakeStringMsg("[Arg" & i & "]")
    Next i
    
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
    str = MakeStringMsg(mTestInstanceName) _
            & "," & MakeStringMsg(mMacroName) _
            & "," & MakeStringMsg(mFrameworkName) _
            & "," & MakeStringMsg(mAcquireGroupName) _
            & "," & MakeStringMsg(mAcquireInstanceName)

    Dim i As Long
    For i = 0& To mParameters.ArgParameterCount - 1
        str = str & "," & MakeStringMsg(mParameters.Arg(i))
    Next i
    
    If ArgCount > mParameters.ArgParameterCount Then
        For i = mParameters.ArgParameterCount To ArgCount - 1
            str = str & ","
        Next i
    End If
    
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
'
    Dim str As String
    str = "TestInstanceName:" & mTestInstanceName & vbCrLf _
            & "TestProcedureName:" & mMacroName & vbCrLf _
            & "ActionName:" & mFrameworkName & vbCrLf _
            & "GroupName:" & mAcquireGroupName & vbCrLf _
            & "InstanceName:" & mAcquireInstanceName & vbCrLf _
            & "Param:[" & vbCrLf
    
    Dim i As Long
    For i = 0& To mParameters.ArgParameterCount - 1
        str = str & Space$(4) & "Arg" & i & ":" & mParameters.Arg(i) & vbCrLf
    Next i
    
    str = str & "]" & vbCrLf
    
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


