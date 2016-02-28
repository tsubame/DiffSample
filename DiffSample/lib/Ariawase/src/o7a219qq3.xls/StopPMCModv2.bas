Attribute VB_Name = "StopPMCModv2"
Option Explicit

Public ProberTemp_ref As Double
Private PTEMPLIMIT As Double
#Const EEE_AUTO_JOB_LOCATE = 2      '1:長崎200mm,2:長崎300mm,3:熊本
Function Flg_StopPMC(blnFlg_StopPMC As Boolean, StopPMC_Comment As String)

    '### プローバ温度　FailSafe ######################################
    If ProberTemp_ref = 0 Then Call GetProberTemp
    
    If ProberTemp_ref >= 20 And ProberTemp_ref <= 30 Then
        PTEMPLIMIT = 5
    Else
        PTEMPLIMIT = 1
    End If
    
    If TenkenTemp >= (ProberTemp_ref - PTEMPLIMIT) And TenkenTemp <= (ProberTemp_ref + PTEMPLIMIT) Then
       blnFlg_StopPMC = False
    Else
       MsgBox "Prober temperature is wrong.Please check .", vbExclamation
       blnFlg_StopPMC = True
       StopPMC_Comment = "Error!! Prober temperature is not " & ProberTemp_ref & "deg .Check prober"
    End If
    '################################################################

    '### 光源Lux確認　FailSafe ######################################
    Dim IllumErr As String
    If Flg_Opt_Judgment = True Then
       IllumErr = "TENKEN Error! NowOpt: " & ArmOpt & " : " & ArmLux & " Lux  Please Check!"
       MsgBox IllumErr, vbExclamation, "<<Illuminator Error>>"
       blnFlg_StopPMC = True
       StopPMC_Comment = "Error!! NSIS-Vw is Please Check  "
    End If
    '################################################################

    '### APMU FailSafe ##############################################
    If APMU_CheckFailSafe_f = False Then
       blnFlg_StopPMC = True
    End If
    '################################################################
    
    '### Monitor ReadResponseTime ###################################
    If MonitorRRT = False Then
        blnFlg_StopPMC = True
    End If
    '################################################################
    
    '### PALSエラーフラグ発生 #######################################
    If Flg_StopPMC_PALS = True Then
       blnFlg_StopPMC = True
    End If
    '################################################################
    
    '### CSV-file FailSafe ##########################################
    If Flg_CsvFileFailSafe = False Then
       blnFlg_StopPMC = True
    End If
    '################################################################

    '### ↓OTPタイプの場合はFailSafe処理の結果で判定 ################
    If OTPBWC_ERR <> 0 And Flg_Tenken = 0 Then
        blnFlg_StopPMC = True
        StopPMC_Comment = "Error!! OTPBWC  "
    End If
    '################################################################
    
    '### HashCodeCheck Result #######################################
    If Flg_HashCheckResult = False Then
       blnFlg_StopPMC = True
    End If
    '################################################################
    
    '### ProbeCard DataSave Error ###################################
    'Kumamoto Only
    #If EEE_AUTO_JOB_LOCATE = 3 Then
        If Flg_StopPMC_Contact = True Then
            blnFlg_StopPMC = True
        End If
    #End If
    '################################################################
    
    '### BlowLotNameCheck ###########################################
    If blnFlg_BlowCheck = True Then
        blnFlg_StopPMC = True
        StopPMC_Comment = "Error Occured@MakeBlowData_Lot. LotName has invalid string"
    End If
    '################################################################

End Function

Public Sub GetProberTemp()

    Dim wkshtObj As Object
    
    Const ProductionIFSheetName = "Production IF"

    '======= WorkSheet Select =============
    Set wkshtObj = ThisWorkbook.Sheets(ProductionIFSheetName)
         
    '======= WorkSheet ErrorProcess =======
    If IsEmpty(wkshtObj) Then
        MsgBox "Non Production IF WorkSheet!!"
        Exit Sub
    End If

    '======= Get Device Type ==============
    ProberTemp_ref = wkshtObj.Cells(18, 2)
End Sub
