Attribute VB_Name = "XEeeAuto_ImageAqcuireTest"
Option Explicit

Public Sub ReadDebugImage(ByRef acqParam As CAcquireFunctionInfo, ByRef pParamPlane As CParamPlane)
    
    If acqParam.ArgParameterCount < 22 Then
        Exit Sub
    End If
 
    Call pParamPlane.plane.SetPMD(acqParam.Arg(21))
    
    Dim site As Long
    
    For site = 0 To TheExec.sites.ExistingCount - 1
        Call pParamPlane.plane.ReadFile(site, acqParam.Arg(20) & "_" & CStr(site) & ".stb", idpFileBinary)
    Next site
    
End Sub
