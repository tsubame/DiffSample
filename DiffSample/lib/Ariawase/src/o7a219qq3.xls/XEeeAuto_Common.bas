Attribute VB_Name = "XEeeAuto_Common"
'�T�v:
'   EeeAuto���ŕ��L���g�p����A�ق��ł��g�p�����ł��낤�֐��Q
'
'�ړI:
'
'
'�쐬��:
'   2012/02/03 Ver0.1 D.Maruyama
'   2012/04/09 Ver0.2 D.Maruyama�@�@FW_SeparateFailSiteGnd�Ŏg�p����VarBank���̒�`��ǉ�
'   2012/10/19 Ver1.2 K.Tokuyoshi   �ȉ��̊֐���ǉ�
'                                   �Em_GetLimit
'                                   �Emf_GetResult
'                                   �EngCapture_Judge_f
'                                   �EEnableFlag_False_f
'   2012/12/25 Ver1.3 H.Arikawa     "ngCap"��"ngCap1"�ɕύX
'   2012/12/26 Ver1.4 H.Arikawa     "ngCap2-5"��ǉ�
'   2013/03/15 Ver1.5 H.Arikawa     �s�v�����폜
'   2013/10/28 Ver1.6 H.Arikawa     �����ݒ�ȗ��̃t���O��

Option Explicit

Public Const JOB_KUMAMOTO_S As Long = 0
Public Const JOB_NAGASAKI_200_S As Long = 1
Public Const JOB_NAGASAKI_300_S As Long = 2

Public Const PIN_NAME_VDDSUB As String = "__VDDSUB_PIN_NAME__"
Public Const GND_SEPARATE_APMU_UB As String = "__GND_SEPARATE_APMU_UB__"
Public Const GND_SEPARATE_CUB_UB As String = "__GND_SEPARATE_CUB_UB__"
Public Const EEE_AUTO_NOUSE_STBSUB As String = "-"
Public Const EEE_AUTO_NOUSE_RELAY As String = "-"

Public Function mf_div(ByVal val1 As Double, ByVal val2 As Double, Optional ByVal errVal As Double = 0) As Double

    If val2 <> 0# Then
        mf_div = val1 / val2
    Else
        mf_div = errVal
    End If

End Function

Public Sub m_GetLimit(ByRef dblLoLimit As Double, ByRef dblHiLimit As Double)

    Dim strArgList() As String
    Dim lngArgCnt As Long
    
    Call TheExec.DataManager.GetArgumentList(strArgList, lngArgCnt)
    dblLoLimit = val(strArgList(5 * LimitSetIndex + 0))
    dblHiLimit = val(strArgList(5 * LimitSetIndex + 1))
    
End Sub

Public Function mf_GetResult(ByVal strKey As String, ByRef pResult() As Double) As Double

    On Error GoTo ErrorExit

    Call TheDcTest.GetTempResult(strKey, pResult)

    Exit Function
    
ErrorExit:
    Call TheResult.GetResult(strKey, pResult)

End Function

'��
'��������������������������������������
'NgCapture_Test�p:Start
'��������������������������������������
'��

'���e:
'
'
'�p�����[�^:
'[Arg1]         In  �Ώ�TestLabel
'[Arg2]         In  �Ώ�LoLimit
'[Arg3]         In  �Ώ�HiLimit
'[Arg4]         In  �Ώ�LimitValid
'
'
Public Function ngCapture_Judge_f() As Double  '2012/11/16 175JobMakeDebug

    On Error GoTo ErrorExit

    Dim site As Long

    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Err.Raise 9999, "ngCapturel_Judge_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
    End If

    'Capture����
    Dim tmpValue1() As Double
    Dim dblLoLimit As Double
    Dim dblHiLimit As Double
    Dim dblLimValid As Double

    Call mf_GetResult(ArgArr(0), tmpValue1)
    dblLoLimit = CDbl(ArgArr(1))
    dblHiLimit = CDbl(ArgArr(2))
    dblLimValid = CDbl(ArgArr(3))

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then

            Select Case dblLimValid
                Case 1
                    If tmpValue1(site) < dblLoLimit Then
                        TheExec.Flow.EnableWord("ngCap1") = True
                        TheExec.Flow.EnableWord("ngCap2") = True
                        TheExec.Flow.EnableWord("ngCap3") = True
                        TheExec.Flow.EnableWord("ngCap4") = True
                        TheExec.Flow.EnableWord("ngCap5") = True
                    End If
                Case 2
                    If tmpValue1(site) > dblHiLimit Then
                        TheExec.Flow.EnableWord("ngCap1") = True
                        TheExec.Flow.EnableWord("ngCap2") = True
                        TheExec.Flow.EnableWord("ngCap3") = True
                        TheExec.Flow.EnableWord("ngCap4") = True
                        TheExec.Flow.EnableWord("ngCap5") = True
                    End If
                Case 3
                    If tmpValue1(site) < dblLoLimit And tmpValue1(site) > dblHiLimit Then
                        TheExec.Flow.EnableWord("ngCap1") = True
                        TheExec.Flow.EnableWord("ngCap2") = True
                        TheExec.Flow.EnableWord("ngCap3") = True
                        TheExec.Flow.EnableWord("ngCap4") = True
                        TheExec.Flow.EnableWord("ngCap5") = True
                    End If
                Case Else
            End Select

        End If
    Next site


    Exit Function

ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function

Public Function EnableFlag_False_f() As Double

    TheExec.Flow.EnableWord("ngCap1") = False
    TheExec.Flow.EnableWord("ngCap2") = False
    TheExec.Flow.EnableWord("ngCap3") = False
    TheExec.Flow.EnableWord("ngCap4") = False
    TheExec.Flow.EnableWord("ngCap5") = False

End Function
'��
'��������������������������������������
'NgCapture_Test�p:End
'��������������������������������������
'��

'���e:
'   PowerDown��Disconnect���s���B
'
'�p�����[�^:
'    [Arg0]      In   ������ PowerSuppluyVoltage�V�[�g�ł̖���
'    [Arg1]      In   �V�[�P���X���@PowerSequence�V�[�g�ł̖���
'    [Arg2]      In   ���s��̃E�F�C�g�^�C��(�ȗ��\ �ȗ���Wait�Ȃ�)
'    [Arg3]      In   �s�����i�s���O���[�v���j
'    [ArgN-1]
'    [ArgN-1]    In   �T�C�g�ԍ�(�ȗ����ꂽ�ꍇ�͑S�T�C�g)
'�߂�l:
'
'���ӎ���:
'
Public Sub PowerDownAndDisconnect()
       
    Call PowerDown4ApmuUnderShoot
    'Pin�̐؂藣�����s���B
    Call DisconnectAllDevicePins
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

Public Sub PatRun( _
    ByVal patGroupName As String, _
    Optional ByVal startLabel As String = "START", _
    Optional timeSetName As String = "", _
    Optional categoryName As String = "", _
    Optional selectorName As String = "" _
)
    Call RunPattern(patGroupName)

End Sub


Public Sub WaitSet(ByVal waitTime As Double)
    If TheExec.RunOptions.AutoAcquire = True Then
        Call TheHdw.TOPT.WAIT(toptTimer, waitTime * 1000)
    Else
        Call TheHdw.WAIT(waitTime)
    End If
End Sub

Public Sub InitializeEeeAutoModules()

    Call InitializeDefectInformation '�d���ݒ�}�N���̏�����
    Call InitializePowerCondition '���׍\���̂̏�����
    
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call InitializeAutoConditionModify '����TestCondition�̕␳
    End If

End Sub

Public Sub UnInitializeEeeAutoModules()

    Call UninitializeDefectInformation
    Call UninitializePowerCondition
    
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call UninitializeAutoConditionModify
    End If
    
End Sub


