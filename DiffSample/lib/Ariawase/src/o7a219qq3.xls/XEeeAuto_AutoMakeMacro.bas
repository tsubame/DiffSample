Attribute VB_Name = "XEeeAuto_AutoMakeMacro"
Option Explicit
Public Function DisconnectAllDevicePins(Optional ByVal site As Long = ALL_SITE) As Long
    Call DisconnectPins("P_VDDL,P_HAN", site)
    Call DisconnectPins("ALL_APMU_PINS", site)
    Call DisconnectPins("ALL_IO_PINS", site)
    DisconnectAllDevicePins = TL_SUCCESS
End Function
'======== �������V�X�e������������֐����o�͂����Module =======
'DisconnectAllDevicePins�Ȃǂ��������������֐�
Public Sub TEMP_ConditionMacro(ByVal tmpPatName As String, ByVal timeSetName As String)
    
    Const PAT_START_LABEL As String = "START"
            
    Call StopPattern_Halt 'EeeJob�֐�
    Call SetTimeOut 'EeeJob�֐�
    
    '�������W�X�^�Ή����[�`�� Start
    '���W�X�^�ݒ蕔Only(keep_alive)�FPatRun
    '���W�X�^�ݒ�+�쓮��:PatSet
    Dim tmpPatGroupName() As String
    Dim i As Integer
    tmpPatGroupName = Split(tmpPatName, ",")
    
    PatCheckCounter = 0
    
    For i = 0 To UBound(tmpPatGroupName)
        If i < UBound(tmpPatGroupName) Then
            With TheHdw.Digital
                Call .Timing.Load(timeSetName)
                Call .Patterns.pat(tmpPatGroupName(i)).Run(PAT_START_LABEL)
            End With
            If TheExec.RunOptions.AutoAcquire = True Then
                Dim iStatus As Long
                If TheHdw.Digital.Patgen.IsRunningAnySite = True Then      'True:Still Running
                    iStatus = 0
                ElseIf TheHdw.Digital.Patgen.IsRunningAnySite = False Then 'False:haltend or keepalive
                    iStatus = 1
                End If
                
                While (iStatus <> 1)
                    If PatCheckCounter < 999 Then
                        TheHdw.TOPT.Recall
                        PatCheckCounter = PatCheckCounter + 1
                        Call WaitSet(10 * mS)
                        Exit Sub
                    End If
                Wend
            End If
        ElseIf i = UBound(tmpPatGroupName) Then
            With TheHdw.Digital
                Call .Timing.Load(timeSetName)
                Call .Patterns.pat(tmpPatGroupName(i)).Start(PAT_START_LABEL)
            End With
        End If
    Next i
    '�������W�X�^�Ή����[�`�� End

        With TheHdw.Digital.Patgen
        Call .FlagWait(cpuA, 0)
        Call SetFVMI_APMU("P_TVMON", 2.522 * V, 50 * mA)
        Call SetFVMI_PPMU("Ph_TVCDSIN", 2.8 * V, ppmu2mA)
        Call .Continue(0, cpuA)
    End With

End Sub
