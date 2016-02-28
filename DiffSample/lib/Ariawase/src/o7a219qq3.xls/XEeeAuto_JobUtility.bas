Attribute VB_Name = "XEeeAuto_JobUtility"
'�T�v:
'   Job���ł݂�Ȃ��g���������Ȋ֐��Q
'
'�ړI:
'
'
'�쐬��:
'   2012/02/03 Ver0.1 D.Maruyama
'   2012/12/26 Ver0.2 H.Arikawa Break��ҏW�B
'   2013/02/27 Ver0.3 H.Arikawa OutPutImageCap��ҏW�B
'   2013/08/23 Ver0.4 H.Arikawa EeeAutoRetryCapture��ǉ��B
'   2013/09/02 Ver0.5 H.Arikawa OutPutImageCap���C���B
'   2013/09/06 Ver0.6 H.Arikawa OutPutImageCap���C���B

Option Explicit

'////////// DEBUG TOOL LIST /////////////////////////////////////////
'OutPutImage    :Capture Image Output for Binary File
'Break          :Test Stop
'StartTime      :Time Mesuer Start
'StopTime       :Time Mesuer Stop
'////////////////////////////////////////////////////////////////////

Private Tms As Double

Public Sub OutPutImageCap(plane As CImgPlane, fileName As String, Optional ByVal tmpPMDName As String = "plane.BasePMD")

    Dim site As Long
    Dim myTime As Variant
    Dim WriteFileName As String
    Dim ReduceMag As Integer
    Dim TypeName As String
        
    ReduceMag = 1
    
    If Flg_Debug = 1 Then
        TheExec.Datalog.WriteComment "***** OutPutImageCap :" & fileName & ""
    End If
    
    'Get TypeName
    TypeName = Mid(CurrentJobName, 4, 3)
    
    'Get TypeName For TENKEN Mode
    If UCase(TypeName) = "KEN" Then
        Dim wkshtObj_IA As Object
        Set wkshtObj_IA = ThisWorkbook.Sheets("Job List")
        TypeName = Mid(wkshtObj_IA.Cells(5, 2), 4, 3)
    End If
    
    'PMD Name Decision
    Dim PMDName As Variant
    If tmpPMDName = "plane.BasePMD" Then
        PMDName = plane.BasePMD.Name
    Else
        PMDName = tmpPMDName
    End If
        
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            myTime = Now
            
            '=== For Address Minus case ===
            If ChipAdr_x(site) < 1 Or ChipAdr_y(site) < 1 Then
                ChipAdr_x(site) = 1
                ChipAdr_y(site) = 1
            End If
            '==============================
            
            WriteFileName = "Z:\" & "imx" & TypeName & "\" & DeviceType & "_" & LotName & "-" & Format(CStr(WaferNo), "00") _
                            & Format(CStr(DeviceNumber_site(site)), "0000") & "-" & CStr(ChipAdr_x(site)) _
                            & "-" & CStr(ChipAdr_y(site)) & "-" & fileName & "-" & ReduceMag & "-" _
                            & Format(myTime, "yyyymmddHHMMSS") & ".stb"
            
            Call plane.SetPMD(PMDName)
            Call plane.WriteFile(site, WriteFileName, idpFileBinary, idpColorFlat)
        End If
    Next site

End Sub

Public Sub Break()

    TheExec.Flow.EnableWord("dc") = False
    TheExec.Flow.EnableWord("image") = False
    TheExec.Flow.EnableWord("shiroten") = False
    TheExec.Flow.EnableWord("margin") = False
    TheExec.Flow.EnableWord("grade") = True
    TheExec.Flow.EnableWord("ngCap1") = False
    TheExec.Flow.EnableWord("ngCap2") = False
    TheExec.Flow.EnableWord("ngCap3") = False
    TheExec.Flow.EnableWord("ngCap4") = False
    TheExec.Flow.EnableWord("ngCap5") = False
    
End Sub

Public Sub StartTime(Optional InputTime As Double = 0)
        
    '--- Time Measure  START ! --------
    Tms = InputTime
    Tms = TheExec.timer(Tms)
    '----------------------------------
    
End Sub

Public Sub stopTime(Optional ByRef RtnTime As Double)
    
    '------- Time Measure  END ! ------------------
    Tms = TheExec.timer(Tms)
    '----------------------------------------------------
        
    '---- OutPut Time -------------------------
    TheExec.Datalog.WriteComment "=== Time Measure ==="
    TheExec.Datalog.WriteComment "Time = " & Format$(Tms, "0.##0") & "[sec]"
    '------------------------------------------
    RtnTime = Tms
    
End Sub

'�����ݒ肩���蒼���Ď�荞�݂��Ď��s����ꍇ�̊֐�(�b��)�@2013/08/23 H.Arikawa
'�b��̗��R�FCon1,Acq1�݂̂̏ꍇ�͂������AAcq2�Ȃǂ�����ꍇ���܂������Ȃ��Ȃ�B
'�P�v�I�ɂ́uImageAQTBL�v���������Ď擾����B
Public Sub EeeAutoRetryCapture(ByVal InstanceName As String)

    Dim ConditionName As String
    Dim AcquireName As String
    ConditionName = InstanceName & "_Con1"
    AcquireName = InstanceName & "_Acq1"
    
    Call TheImageTest.RetryAcquire(ConditionName, "FWSetCondition")
    Call TheImageTest.RetryAcquire(AcquireName, "FWImageAcquire")
    Call TheImageTest.RetryAcquire(AcquireName, "FWPostImageAcquire")

End Sub
