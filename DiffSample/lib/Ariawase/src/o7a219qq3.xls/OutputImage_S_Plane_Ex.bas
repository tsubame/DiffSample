Attribute VB_Name = "OutputImage_S_Plane_Ex"
Option Explicit

Public Flg_OutputImageCheck As Boolean

Public Sub OutPutImageCap_S_Plane_Ex(plane As CImgPlane, fileName As String, Optional ByVal Flg_OutputCondition As Boolean = False, Optional ByVal OutputNo As Integer = 2, Optional ByVal site As Long = ALL_SITE, Optional ByVal gDefaultCaptureCategory As String = "imx", Optional ByVal tmpPMDName As String = "plane.BasePMD", Optional ByVal gCaptureDirectory As String = gDefaultCaptureDir)

    Dim myTime As Variant
    Dim WriteFileName(nSite) As String
    Dim ReduceMag As Integer
    Dim TypeName As String
    Dim i As Integer
    Dim Directory As String
    Dim PMDName As Variant
    Dim NowPlane As String

    If Flg_OutputCondition = False Then Exit Sub
    
    If Flg_AutoMode = True Then
    
        If 1 > OutputNo Or OutputNo > 3 Then
            MsgBox "OutputNo is Wrong!! Use Only 1-3 !! "
            Flg_OutputImageCheck = True
            Exit Sub
        End If
        
        If Chip_f < OutputNo Then Exit Sub
        
    End If
    
    ReduceMag = 1
    
    'Get TypeName
    TypeName = Mid(TheExec.CurrentJob, 4, 3)
    
    'Get TypeName For TENKEN Mode
    If UCase(TypeName) = "KEN" Then
        Dim wkshtObj_IA As Object
        Set wkshtObj_IA = ThisWorkbook.Sheets("Job List")
        TypeName = Mid(wkshtObj_IA.Cells(5, 2), 4, 3)
    End If
    
    'Directory Decision
    Directory = Left(gCaptureDirectory, 1)

    'PMD Name Decision
    If tmpPMDName = "plane.BasePMD" Then
        PMDName = plane.BasePMD.Name
    Else
        PMDName = tmpPMDName
    End If
    
    'FileName Check
    If InStr(fileName, "-") > 0 Then
        MsgBox "FileName is Wrong!! Don't use - !! "
        Flg_OutputImageCheck = True
        Exit Sub
    End If
    
    If Flg_AutoMode = True Or Flg_LoopMode = True Then

        If Directory = "F" Or Directory = "G" Or Directory = "H" Or Directory = "Q" Or Directory = "Y" Or Directory = "" Then
            TheExec.Datalog.WriteComment "OutputImagePlace NG!!  site: " & site
            MsgBox "OutputImagePlace NG!! Don't set F/G/H/Q/Y Drive!!"
            Flg_OutputImageCheck = True
            Exit Sub
        End If

        If site = ALL_SITE Then
        
            For site = 0 To nSite
            
                myTime = Now
            
                '=== For Address Minus case ===
                If ChipAdr_x(site) < 1 Or ChipAdr_y(site) < 1 Then
                    ChipAdr_x(site) = 1
                    ChipAdr_y(site) = 1
                End If
                '==============================
            
                WriteFileName(site) = gCaptureDirectory & DeviceType & "_" & LotName & "-" & Format(CStr(WaferNo), "00") _
                            & Format(CStr(DeviceNumber_site(site)), "0000") & "-" & CStr(ChipAdr_x(site)) _
                            & "-" & CStr(ChipAdr_y(site)) & "-" & fileName & "-" & ReduceMag & "-" & site & "-" _
                            & Format(myTime, "yyyymmddHHMMSS") & ".stb"
                
            Next site
            
            NowPlane = plane.Name
            Call plane.SetPMD(PMDName)
            Call TheHdw.IDP.WriteFileEx(-1, NowPlane, idpColorFlat, WriteFileName, idpFileBinary, IdpTesterPC, idpWriteFileCurrPMD, , idpRWFileExProcessNonBlocking)
            
        Else
        
            If TheExec.sites.site(site).Active Then
                myTime = Now
            
                '=== For Address Minus case ===
                If ChipAdr_x(site) < 1 Or ChipAdr_y(site) < 1 Then
                    ChipAdr_x(site) = 1
                    ChipAdr_y(site) = 1
                End If
                '==============================
            
                WriteFileName(site) = gCaptureDirectory & DeviceType & "_" & LotName & "-" & Format(CStr(WaferNo), "00") _
                            & Format(CStr(DeviceNumber_site(site)), "0000") & "-" & CStr(ChipAdr_x(site)) _
                            & "-" & CStr(ChipAdr_y(site)) & "-" & fileName & "-" & ReduceMag & "-" & site & "-" _
                            & Format(myTime, "yyyymmddHHMMSS") & ".stb"
            
                NowPlane = plane.Name
                Call plane.SetPMD(PMDName)
                Call TheHdw.IDP.WriteFileEx(site, NowPlane, idpColorFlat, WriteFileName(site), idpFileBinary, IdpTesterPC, idpWriteFileCurrPMD, , idpRWFileExProcessNonBlocking)
            End If
        End If
        
    Else
    
        If Directory = "F" Or Directory = "G" Or Directory = "H" Or Directory = "Q" Or Directory = "Y" Or Directory = "" Then
            TheExec.Datalog.WriteComment "OutputImagePlace NG!!  site: " & site
            MsgBox "OutputImagePlace NG!! Don't set F/G/H/Q/Y Drive!!"
            Break
            Exit Sub
        Else
            For i = 0 To 10
                TheExec.Datalog.WriteComment "***** OutPutImageCap :" & fileName & " at " & Directory & " Drive!!"
            Next i
        End If
        
        If site = ALL_SITE Then
    
            For site = 0 To nSite
                    myTime = Now
                
                    '=== For Address Minus case ===
                    If ChipAdr_x(site) < 1 Or ChipAdr_y(site) < 1 Then
                        ChipAdr_x(site) = 1
                        ChipAdr_y(site) = 1
                    End If
                    '==============================
                
                    WriteFileName(site) = gCaptureDirectory & "Debug" & gDefaultCaptureCategory & TypeName & "_" & LotName & "-" & Format(CStr(WaferNo), "00") _
                                & Format(CStr(DeviceNumber_site(site)), "0000") & "-" & CStr(ChipAdr_x(site)) _
                                & "-" & CStr(ChipAdr_y(site)) & "-" & fileName & "-" & ReduceMag & "-" & site & "-" _
                                & Format(myTime, "yyyymmddHHMMSS") & ".stb"
                
            Next site
            
            NowPlane = plane.Name
            Call plane.SetPMD(PMDName)
            Call TheHdw.IDP.WriteFileEx(-1, NowPlane, idpColorFlat, WriteFileName, idpFileBinary, IdpTesterPC, idpWriteFileCurrPMD, , idpRWFileExProcessNonBlocking)
        
        Else

            If TheExec.sites.site(site).Active Then
                myTime = Now
            
                '=== For Address Minus case ===
                If ChipAdr_x(site) < 1 Or ChipAdr_y(site) < 1 Then
                    ChipAdr_x(site) = 1
                    ChipAdr_y(site) = 1
                End If
                '==============================
            
                WriteFileName(site) = gCaptureDirectory & "Debug" & gDefaultCaptureCategory & TypeName & "_" & LotName & "-" & Format(CStr(WaferNo), "00") _
                            & Format(CStr(DeviceNumber_site(site)), "0000") & "-" & CStr(ChipAdr_x(site)) _
                            & "-" & CStr(ChipAdr_y(site)) & "-" & fileName & "-" & ReduceMag & "-" & site & "-" _
                            & Format(myTime, "yyyymmddHHMMSS") & ".stb"
            
                NowPlane = plane.Name
                Call plane.SetPMD(PMDName)
                Call TheHdw.IDP.WriteFileEx(site, NowPlane, idpColorFlat, WriteFileName(site), idpFileBinary, IdpTesterPC, idpWriteFileCurrPMD, , idpRWFileExProcessNonBlocking)
            End If
        
        End If
    End If

End Sub

