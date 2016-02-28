Attribute VB_Name = "basCheckAPMU"
Option Explicit

' This function calls the APMU_CheckFailSafe_f in CheckAPMUFailSafe.dll
' return value
'   True    : not detect APMU Fail Safe
'   false   : detect APMU Fail Safe

Public Function APMU_CheckFailSafe_f() As Boolean

   On Error GoTo Err_FAILSAFE

    Dim objChkAPMU As Object
    Dim blnRetVal As Boolean
    Dim lngRunMode As Long

    lngRunMode = TheExec.RunMode

    TheExec.RunMode = runModeDebug

    Set objChkAPMU = CreateObject("CheckAPMUFailSafe.clsMain")
    
    blnRetVal = objChkAPMU.APMU_CheckFailSafe_f

    If blnRetVal = False Then
        Application.Visible = True
        Call MsgBox("Detect APMU Fail Safe," & vbCrLf & "Please run the Quick Check for APMU.", vbCritical, "APMU_Fail_Safe")
        Set objChkAPMU = Nothing
        APMU_CheckFailSafe_f = blnRetVal
        TheExec.RunMode = lngRunMode
        Exit Function
    
    End If
    
    TheExec.RunMode = lngRunMode

    Set objChkAPMU = Nothing
    
    APMU_CheckFailSafe_f = blnRetVal
    
    Exit Function

Err_FAILSAFE:

    Set objChkAPMU = Nothing

    Application.Visible = True

    If Err.Number = 13 Or Err.Number = -2147024770 Then
        Call MsgBox("APMU_CheckFailSafe_f calling", vbCritical, "CheckAPMUFailSafe")
    ElseIf Err.Number = 429 Then
        Call MsgBox("CheckAPMUFailSafe.Dll not found or not Registerd", vbCritical, "CheckAPMUFailSafe")
    Else
        Call MsgBox(Err.Description & "[" & Err.Number & "]<" & Err.LastDllError & ">", vbCritical, "CheckAPMUFailSafe")
    End If

    APMU_CheckFailSafe_f = False
    TheExec.RunMode = lngRunMode

End Function
