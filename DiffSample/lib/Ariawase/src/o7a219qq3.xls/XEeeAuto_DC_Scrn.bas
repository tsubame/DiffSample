Attribute VB_Name = "XEeeAuto_DC_Scrn"
'�T�v:
'
'
'�ړI:
'   ���d��SCRN���s�����߂̃��W���[��
'
'�쐬��:
'   2012/01/23 Ver0.1 D.Maruyama
'   2012/02/14 Ver0.2 D.Maruyama�@SV125�̒l��SCRN�t���O�ɂ�炸Add����悤�ɕύX
'   2012/03/07 Ver0.3 D.Maruyama�@TestInstance����ForceTime��̃p�^�[���ݒ�����悤�ɕύX
'                                 SetMV���Wait��TestCondition������炤�悤�ɕύX
'   2012/10/19 Ver0.4 K.Tokuyoshi �啝�ɏC��
'   2012/10/26 Ver0.5 K.Tokuyoshi �ȉ��̊֐���ǉ�
'                                 �EResultScrnSpec_f
'   2012/11/14�E11/15 Ver0.6 T.Morimoto  �ȉ��̊֐���ǉ��E�C��
'                                 �EFW_DcScreeningSet�AScreening_GetParameter�AFW_DcScreeningStop

Option Explicit

'+
' Name      : ScreeningFlag
' Purpose   : [J]   �[�q�d������̂Ȃ��A�P���ȍ��d���X�N���[�j���O����e�X�g�ŁA������������Ȃ����������A
'                   "Flg_Scrn"�̒l�ŕԂ��B
'             [E]   Test implementation function for simple high-voltage screening. Returns the
'                   "Flg_Scrn" value indicating screening-on/off.
' Arguments : [J]   Test Instances�V�[�g��Arg20/21/22�����́B
'             [E]   Arguments must be specified at cells Arg20,21,22 on Test Instances worksheet.
'                   Arg20   "Condition Name" for relay, illuminator, power supply, pin electronics
'                           settings, which are defined on TestCondition worksheet.
'                   Arg21   Wait time for which high voltage screening is applied.
'                   Arg22   The test label.
' Restrictions  [J] Test Instances�V�[�g����̌Ăяo������B
'               [E] Must be called from TheExec.Flow.
' History       First drafted by TM 2014-Feb-03
'                   - For Kumamoto IMX219 Shinraisei analysis (Koyama-san).
'                   - For masterbook 013/016
'-
Private Function ScreeningFlag() As Double

    On Error GoTo ErrorExit

    Call SiteCheck
    
    '�ϐ���`
    'Arg 20: ����p�����[�^_Opt_�����[ & Set_Voltage_�[�q�ݒ� & Pattern & Wait
    'Arg 21: ����v���d�l���Ɏw�肳���X�N���[�j���O��Wait����
    'Arg 22: �X�N���[�j���O�ƈꏏ�ɍs����DC�����Wait����
    Dim strSetCondition As String       'Arg20: [Test Condition]'s condition name excluding screening wait.
    Dim dblScreeningWait As Double      'Arg21: The wait time for screening specified on the specification sheet.
    Dim testLabelName As String             'Arg22: The test label name of the test.
    
    Dim site As Long
    Dim tmpResult(nSite) As Double

    If Flg_Scrn = 0 Then
        TheResult.Add "IDDBI_HSN", tmpResult
    End If

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmpResult(site) = Flg_Scrn And (Flg_Tenken = 0)
        End If
    Next site

    '�ϐ���荞��
    If Not Screening_GetParameterFlag( _
                strSetCondition, _
                dblScreeningWait, _
                testLabelName) Then
                MsgBox "The Number of ScreeningFlag's arguments is invalid!"
                Call DisableAllTest 'EeeJob�֐�
                Exit Function
    End If

    If Flg_Scrn = 1 And Flg_Tenken = 0 Then

        '========== Set Condition ===============================
        Call TheCondition.SetCondition(strSetCondition)
        Call TheHdw.WAIT(dblScreeningWait)
        Call TheResult.Add(testLabelName, tmpResult)
    Else
        Call TheResult.Add(testLabelName, tmpResult)
        Exit Function
    End If
    
    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function

Private Function Screening_GetParameterFlag( _
    ByRef strSetCondition As String, _
    ByRef dblScreeningWait As Double, _
    ByRef testLabelName As String _
    ) As Boolean
    
    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Screening_GetParameterFlag = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strSetCondition = ArgArr(0)
    dblScreeningWait = ArgArr(1)
    testLabelName = ArgArr(2)
On Error GoTo 0

    Screening_GetParameterFlag = True
    Exit Function
    
ErrHandler:

    Screening_GetParameterFlag = False
    Exit Function

End Function


'��
'��������������������������������������
'IMX145_Scrn_VBA�}�N�� :Start
'��������������������������������������
'��

'���e:
'   Flg�̒l���i�[����B
'
'�p�����[�^:
'
'���ӎ���:
'
'
Public Function ResultScrnFlg_f() As Double

    On Error GoTo ErrorExit
        
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult
    
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = Flg_Scrn
            End If
        Next site
    Else
    End If
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function

'���e:
'   Wait�̒l���i�[����B
'
'�p�����[�^:
'
'���ӎ���:
'
'
Public Function ResultScrnWait_f() As Double

    On Error GoTo ErrorExit

    Call SiteCheck

    Dim site As Long

    Dim retResult(nSite) As Double
    Erase retResult

    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
        '�p�����[�^�̎擾
        '�z�萔��菬������΃G���[�R�[�h
        Dim ArgArr() As String
        If Not EeeAutoGetArgument(ArgArr, 1) Then
            Err.Raise 9999, "ResultScrnWait_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        End If

        'Wait���Ԏ擾
        Dim tmpValue1 As Double
        tmpValue1 = ArgArr(0) '2012/11/15 175Debug Arikawa CDbl Delete

        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = tmpValue1
            End If
        Next site
    Else
    End If

    '�W���b�W
    Call test(retResult)

    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function

ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

Private Function FW_DcScreeningSet() As Double

    On Error GoTo ErrorExit

    Call SiteCheck
    
    '�ϐ���`
    'Arg 20: ����p�����[�^_Opt_�����[ & Set_Voltage_�[�q�ݒ� & Pattern & Wait
    'Arg 21: ����v���d�l���Ɏw�肳���X�N���[�j���O��Wait����
    'Arg 22: �X�N���[�j���O�ƈꏏ�ɍs����DC�����Wait����
    Dim strSetCondition As String       'Arg20: [Test Condition]'s condition name excluding screening wait.
    Dim dblScreeningWait As Double      'Arg21: The wait time for screening specified on the specification sheet.
    Dim dblMeasurementWait As Double    'Arg22: The wait time for DC measurement (V125, VBGR... etc).
    
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then

        '�ϐ���荞��
        If Not Screening_GetParameter( _
                    strSetCondition, _
                    dblScreeningWait, _
                    dblMeasurementWait) Then
                    MsgBox "The Number of FW_DcSetScreening's arguments is invalid!"
                    Call DisableAllTest 'EeeJob�֐�
                    Exit Function
        End If
            
        '========== Set Condition ===============================
        Call TheCondition.SetCondition(strSetCondition)
        If dblScreeningWait > dblMeasurementWait Then Call TheHdw.WAIT(dblScreeningWait - dblMeasurementWait)
    Else
        Exit Function
    End If
        

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function

Private Function Screening_GetParameter( _
    ByRef strSetCondition As String, _
    ByRef dblScreeningWait As Double, _
    ByRef dblMeasurementWait As Double _
    ) As Boolean
    
    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Screening_GetParameter = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strSetCondition = ArgArr(0)
    dblScreeningWait = ArgArr(1)
    dblMeasurementWait = ArgArr(2)
On Error GoTo 0

    Screening_GetParameter = True
    Exit Function
    
ErrHandler:

    Screening_GetParameter = False
    Exit Function

End Function

Private Function FW_DcScreeningMeasure() As Double

    On Error GoTo ErrorExit

    '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
    Call SiteCheck
        
    '�萔��`
    Const PARAM_DELIMITER As String = ","
    
    '�ϐ���`
    Dim strSetDCMeasure As String       'Arg20�@DC�V�i���I
    Dim strTestLabelNames As String     'Arg21 DC Measure���s�������ʂ̃��x����(�����J���}�̉\������)
    Dim strDummyTestResult As String    'Arg22 DC Measure���s��Ȃ������ꍇ�̃_�~�[�̒l(�����J���}�̉\������)
    Dim strTestLabels() As String
    Dim strDummyValues() As String
    Dim dblDummyValues(nSite) As Double
    Dim i As Long
    Dim site As Long
    
    '�ϐ���荞��
    If Not Screening_GetMeasure( _
                strSetDCMeasure, _
                strTestLabelNames, _
                strDummyTestResult) Then
                MsgBox "The Number of FW_DcScreeningMeasure's arguments is invalid!"
                Call DisableAllTest 'EeeJob�֐�
                Exit Function
    End If
    
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
                        
        '========== DC�V�i���I�V�[�g���s ===============================
        Call TheDcTest.SetScenario(strSetDCMeasure)
        TheDcTest.Execute
        
    Else
        '�e�X�g���x������
        strTestLabels = Split(strTestLabelNames, PARAM_DELIMITER)
        '�_�~�[����l����
        strDummyValues = Split(strDummyTestResult, PARAM_DELIMITER)
        '--Error�����F�e�X�g���x�����ƁA�_�~�[�p�̓����l���s��v�̏ꍇ
        If UBound(strTestLabels) <> UBound(strDummyValues) Then
            Call MsgBox("The number of test labels and dummy values do not match. Check <parameter/equation> column on your specification sheet.")
            GoTo ErrorExit
        End If
        
        For i = 0 To UBound(strDummyValues)
            If IsNumeric(strDummyValues(i)) Then
                For site = 0 To nSite
                    dblDummyValues(site) = CDbl(strDummyValues(i))
                Next site
                Call TheResult.Add(strTestLabels(i), dblDummyValues)
            Else
                Call MsgBox("The dummy return value must be numeric.")
                GoTo ErrorExit
            End If
        Next i
        Exit Function
    End If

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function


Private Function FW_DcScreeningStop() As Double

    On Error GoTo ErrorExit

    '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
    Call SiteCheck
    
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
        Call PowerDown4ApmuUnderShoot
    Else
        Exit Function
    End If
        
    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function

Private Function Screening_GetMeasure( _
    ByRef strSetDCMeasure As String, _
    ByRef strTestLabelNames As String, _
    ByRef strDummyTestResult As String _
    ) As Boolean

    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Screening_GetMeasure = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strSetDCMeasure = ArgArr(0)
    strTestLabelNames = ArgArr(1)
    strDummyTestResult = ArgArr(2)
On Error GoTo 0

    Screening_GetMeasure = True
    Exit Function
    
ErrHandler:

    Screening_GetMeasure = False
    Exit Function

End Function

'��
'��������������������������������������
'SCR TOPT�p�@FW_SetConditionMacro:
'��������������������������������������
'��
'

'���e:
'   SetCondition���s��
'
'�p�����[�^:
'    [Arg0]      In Condition Name
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_SetDcScreening_topt(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    '========= TestCondition Call ======================
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
    
        If Parameter.ArgParameterCount() <> 1 Then
            Err.Raise 9999, "FW_SetDcScreening_topt", "The number of FW_SetDcScreening_topt's arguments is invalid." & " @ " & Parameter.ConditionName
        End If
                        
        '========== Set Condition ===============================
        Call TheCondition.SetCondition(Parameter.Arg(0))
        
    Else
        Exit Sub
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'��
'��������������������������������������
'SCR TOPT�p�@FW_MeasureMacro:
'��������������������������������������
'��
'

'���e:
'   Measure���s��
'
'�p�����[�^:
'    [Arg0]      In DC Test Scenario Name
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_DcMeasure_topt(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_DcMeasure_topt", "The number of FW_DcMeasure_topt's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '========= TestCondition Call ======================
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
                        
        '========== DC�V�i���I�V�[�g���s ===============================
        TheDcTest.SetScenario (Parameter.Arg(0))
        TheDcTest.Execute
        
    Else
        Exit Sub
    End If
    '========= TestCondition Call ======================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub


