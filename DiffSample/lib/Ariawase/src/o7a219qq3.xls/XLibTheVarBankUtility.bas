Attribute VB_Name = "XLibTheVarBankUtility"
'�T�v:
'   TheVarBank�̃��[�e�B���e�B
'
'�ړI:
'   TheVarBank:CVarBank�̏�����/�j���̃��[�e���e�B���`����
'
'�쐬��:
'   a_oshima

Option Explicit

Public TheVarBank As IVarBank

Private mSaveMode As Boolean
Private mSaveFileName As String
Private mTheVarBank As CVarBank
Private mTheVarBankInterceptor As CVarBankInterceptor


Public Sub CreateTheVarBankIfNothing()
'���e:
'   As New�̑�ւƂ��ď���ɌĂ΂���C���X�^���X��������
'
'�p�����[�^:
'   �Ȃ�
'
'���ӎ���:
'
    
    On Error GoTo ErrHandler
    If TheVarBank Is Nothing Then
        Set mTheVarBank = New CVarBank
        Set TheVarBank = mTheVarBank
        mSaveMode = False
    End If
    Exit Sub
ErrHandler:
    Set TheVarBank = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub SaveModeTheVarBank(ByVal pEnableLoggingTheVar As Boolean, Optional saveFileName As String)
    mSaveFileName = saveFileName
    If mSaveMode <> pEnableLoggingTheVar Then
        If pEnableLoggingTheVar = True Then
            '���荞�݂����
            Set mTheVarBankInterceptor = New CVarBankInterceptor
            Call mTheVarBankInterceptor.Initialize(mTheVarBank)
            Set TheVarBank = mTheVarBankInterceptor.AsIVarBank
            mSaveMode = True
            TheExec.Datalog.WriteComment "Eee JOB Output Log! :TheCondition"
        Else
            '���荞�݊O��
            Call mTheVarBankInterceptor.Initialize(Nothing)
            Set mTheVarBankInterceptor = Nothing
            Set TheVarBank = mTheVarBank
            mSaveMode = False
        End If
    End If
End Sub

Public Sub DestroyTheVarBank()
    Set TheVarBank = Nothing
    Set mTheVarBank = Nothing
    Set mTheVarBankInterceptor = Nothing
End Sub

Public Function RunAtJobEnd() As Long
    
    If Not (TheVarBank Is Nothing) Then
        TheVarBank.Clear
        If Not (mTheVarBankInterceptor Is Nothing) Then
            Call mTheVarBankInterceptor.SaveLogFile(mSaveFileName)
        End If
        Call SaveModeTheVarBank(False)            '�e�X�g�I����False
    End If

End Function
