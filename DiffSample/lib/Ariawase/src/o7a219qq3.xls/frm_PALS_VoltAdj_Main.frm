VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_PALS_VoltAdj_Main 
   Caption         =   "PALS - Auto Volt Adjust"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   OleObjectBlob   =   "frm_PALS_VoltAdj_Main.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frm_PALS_VoltAdj_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
'##########################################################
'フォームの×ボタンを消す処理
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000

' ウィンドウに関する情報を返す
Private Declare Function GetWindowLong Lib "USER32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' ウィンドウの属性を変更
Private Declare Function SetWindowLong Lib "USER32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' Activeなウィンドウのハンドルを取得
Private Declare Function GetActiveWindow Lib "USER32.dll" () As Long
' メニューバーを再描画
Private Declare Function DrawMenuBar Lib "USER32.dll" (ByVal hWnd As Long) As Long

Dim msgtext As String
Dim ANS As Integer


Private Sub cb_call_Click()
    Call Get_KeepPoint
End Sub

Private Sub cb_keep_Click()
    Call Save_KeepPoint
End Sub

Private Sub cb_savecsv_Click()
    If MsgBox("If You Push [OK],PALS Save to CSV...", vbOKCancel + vbQuestion, VOLTTOOLNAME) = vbCancel Then
        Exit Sub
    End If
    
    Call sub_OutPutCsv(ReadSheetName, Power_Supply_VoltageoffsetFileName)

End Sub

Private Sub ConfirmModeCheck_Click()
    If ConfirmModeCheck.Value = True Then Flow3.enabled = False
    If ConfirmModeCheck.Value = False And Onlyvoltagecheck.Value = False Then Flow3.enabled = True
End Sub



Private Sub Onlyvoltagecheck_Click()
    If Onlyvoltagecheck.Value = True Then
            Flow1.enabled = False
            Flow3.enabled = False
            ConfirmModeCheck.Value = True
            Call ConfirmModeCheck_Click
    Else
            Flow1.enabled = True
            Flow3.enabled = True
            ConfirmModeCheck.Value = False
            Call ConfirmModeCheck_Click
    End If
End Sub
Private Sub Goyaman_Click()
    Dim Goyamanpos As Integer
    
    If Goyaman.Left = 216 Then
        For Goyamanpos = 216 To 252 Step 4
            Goyaman.Left = Goyamanpos
            Call mSecSleep(100)
            DoEvents
        Next Goyamanpos
        lbl_pininfo.Visible = True
    Else
        lbl_pininfo.Visible = False
        For Goyamanpos = 252 To 216 Step -4
            Goyaman.Left = Goyamanpos
            Call mSecSleep(100)
            DoEvents
        Next Goyamanpos
    End If
    
End Sub




Public Sub PinNamePrint_Change()
'On change PinName
    Dim index_site As Integer
    Dim index_pin As Integer

    'For Printing Pin# ================
    If frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex <> -1 And frm_PALS_VoltAdj_Main.SitePrint.ListIndex <> -1 Then
        index_site = frm_PALS_VoltAdj_Main.SitePrint.ListIndex
        index_pin = frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex
        frm_PALS_VoltAdj_Main.PinNumPrint = PowerPinNum(index_site, index_pin)
        lbl_pininfo.Caption = "RSC:" & Pin_Resource(index_pin) + vbCrLf & "Chan:" & Pin_Chans(index_site, index_pin)
        DoEvents
    End If
    '==================================

End Sub


Private Sub SitePrint_Change()
'On change Site#
    Dim index_site As Integer
    Dim index_pin As Integer

    'For Printing Pin# ================
    If frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex <> -1 And frm_PALS_VoltAdj_Main.SitePrint.ListIndex <> -1 Then
        index_site = frm_PALS_VoltAdj_Main.SitePrint.ListIndex
        index_pin = frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex
        frm_PALS_VoltAdj_Main.PinNumPrint = PowerPinNum(index_site, index_pin)
        lbl_pininfo.Caption = "RSC:" & Pin_Resource(index_pin) + vbCrLf & "Chan:" & Pin_Chans(index_site, index_pin)
    End If
    '==================================

End Sub

Private Sub DataLogCheck_Click()

    'Select Datalog Save
'    If DataLogCheck.value = True Then
'        FileAddress = Application.GetSaveAsFilename(InitialFileName:="VoltCheckDataLog.txt", _
'                                                    fileFilter:="textfile(*.txt),*.txt", _
'                                                    title:="DetaLog Save")
'        If FileAddress = False Then DataLogCheck.value = False
'    End If
End Sub


Public Sub RunButton_Click()
'VoltageCheck Start

    '==================================
    Dim start_Site As Integer
    Dim start_Condition As Integer
    Dim start_PinName As Integer
    Dim end_Site As Integer
    Dim end_Condition As Integer
    Dim end_PinName As Integer
    
    Dim ProgressCount As Integer
    Dim ProgressBar_MaxValue As Integer
    Dim ProgressBar_NowValue As Integer
    Dim Progress_Pin As Integer
    Dim Progress_site As Integer
    Dim Progress_Condition As Integer
    '==================================

    'Button Visiable
    frm_PALS_VoltAdj_Main.ClearButton.enabled = False
    frm_PALS_VoltAdj_Main.BreakButton.Visible = True
    frm_PALS_VoltAdj_Main.RunButton.Visible = False
    frm_PALS_VoltAdj_Main.ClearButton.Visible = False
    frm_PALS_VoltAdj_Main.ExitButton.Visible = False
    
    Break_Counter = 0
    Flg_End_voltcheck = False
    
    Call ContinueMode
    If Flg_Break_voltcheck = True Then Call Get_BreakPoint
    
'    Call Output_Offset(OutPutSheetname, 0)
    
    Call Check_Flg
    If Flg_Break_voltcheck = False Then Call SelectPin_check
    
    If Flg_End_voltcheck = True Then
        Call Clear_Navi
        Exit Sub
    End If
    
    'Defult Config ====================
    start_Site = 0
    start_Condition = 0
    start_PinName = 0
    end_Site = nSite
    end_Condition = Condition_Number - 1
    end_PinName = pin_number - 1
    '==================================
    
    'On Lock Flag =====================
    If Flg_SiteLock = -1 Then
        start_Site = SitePrint.ListIndex
        end_Site = SitePrint.ListIndex
    End If
    If Flg_ConditionLock = -1 Then
        start_Condition = CONDITIONPrint.ListIndex
        end_Condition = CONDITIONPrint.ListIndex
    End If
    If Flg_PinLock = -1 Then
        start_PinName = PinNamePrint.ListIndex
        end_PinName = PinNamePrint.ListIndex
    End If
    '==================================
    
    'Ready Parameter For Progress =====
    Progress_site = nSite + 1
    Progress_Condition = Condition_Number
    Progress_Pin = pin_number
    
    If Flg_SiteLock = -1 Then Progress_site = 1
    If Flg_ConditionLock = -1 Then Progress_Condition = 1
    If Flg_PinLock = -1 Then Progress_Pin = 1
    
    ProgressBar_MaxValue = Progress_site * Progress_Condition * Progress_Pin
    '==================================
    If Flg_DataLog = -1 Then
        FileNo_vc = FreeFile
        Open FileAddress For Append As #FileNo_vc
        If Flg_DataLog = -1 Then Print #FileNo_vc, ""
        If Flg_DataLog = -1 Then Print #FileNo_vc, "###################################"
        If Flg_DataLog = -1 Then Print #FileNo_vc, Date
        If Flg_DataLog = -1 Then Print #FileNo_vc, TIME
        If Flg_DataLog = -1 Then Print #FileNo_vc, "AdjustOffset= " & AdjustOffset & " V"
        If Flg_DataLog = -1 Then Print #FileNo_vc, "###################################"
    End If

    'Auotmatic Process ====================================
    For Now_PinName = start_PinName To end_PinName  'Pin LOOP
        PinNamePrint.ListIndex = Now_PinName

        For Now_Site = start_Site To end_Site       'Site LOOP
            SitePrint.ListIndex = Now_Site
        
            For Now_Condition = start_Condition To end_Condition    'Condition LOOP
                CONDITIONPrint.ListIndex = Now_Condition
                'Flow Clear =======================
                frm_PALS_VoltAdj_Main.ConnectCheck.Visible = False
                frm_PALS_VoltAdj_Main.Applycheck.Visible = False
                frm_PALS_VoltAdj_Main.OffsetCheck.Visible = False
                DoEvents
                '==================================
                
                If CheckRun(Now_PinName) = "ON" Then
                    If Flg_Break_voltcheck = False Or (Flg_Break_voltcheck = True And ProgressCount >= Break_Counter_prev) Then
                        If Flg_DataLog = -1 Then Print #FileNo_vc, ""
                        If Flg_DataLog = -1 Then Print #FileNo_vc, ""
                        If Flg_DataLog = -1 Then Print #FileNo_vc, ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
                        If Flg_DataLog = -1 Then Print #FileNo_vc, ">>> Tester:" & CStr(Sw_Node) & "  Site:" & CStr(Now_Site) & "  Condition:" & CStr(Test_Condition(Now_Condition)) & "  Pin:" & PowerPinName(Now_PinName)
                        Call Check_Pin
                        If Flg_Stop_voltcheck = True Then GoTo BreakPointcheckpin
                        If Flg_MeasSkip = False Then
                            Call Adjust_Volt
                        End If
                    End If
                End If
                
                'For Progress =====================
                ProgressCount = ProgressCount + 1
                ProgressBar_NowValue = Int((ProgressCount / ProgressBar_MaxValue) * frm_PALS_VoltAdj_Main.progress_lbl_full.width)
                If ProgressBar_NowValue <= 0 Then ProgressBar_NowValue = 1
                If ProgressBar_NowValue > frm_PALS_VoltAdj_Main.progress_lbl_full.width Then ProgressBar_NowValue = frm_PALS_VoltAdj_Main.progress_lbl_full.width
                frm_PALS_VoltAdj_Main.progress_lbl.width = ProgressBar_NowValue
                frm_PALS_VoltAdj_Main.ProgressPer = Int(ProgressBar_NowValue / frm_PALS_VoltAdj_Main.progress_lbl_full.width * 100)
                DoEvents
                '==================================
BreakPointcheckpin:
                Break_Counter = ProgressCount
                'On Click Break Button ============
                If Flg_Stop_voltcheck = True And CheckRun(Now_PinName) = "ON" Then
                    msgtext = "You pushed [Break]" + vbCrLf + vbCrLf + _
                              "[OK]Stop Voltcheck" + vbCrLf + _
                              "   (NextTime,You can continue  at after last check)" + vbCrLf + _
                              "[Cancel] Cancel Break & Continue Voltcheck"
                    ANS = MsgBox(msgtext, vbOKCancel + vbQuestion, "Inquire Message")
                    If ANS = vbOK Then
                        Call Save_BreakPoint
                        GoTo BreakPoint
                    ElseIf ANS = vbCancel Then
                        Flg_Stop_voltcheck = False
                    End If
                End If
                '==================================

            Next Now_Condition
        Next Now_Site
    Next Now_PinName
    '======================================================

BreakPoint:
    If Flg_DataLog = -1 Then Close #FileNo_vc
    Call offset_write
    Call Clear_Navi(1)
    Call Get_Power_Condition
End Sub

Public Sub ClearButton_Click()

    Flg_Clear_voltcheck = True
    Call Clear_Navi
    Flg_Clear_voltcheck = False

End Sub

Public Sub BreakButton_Click()

    Flg_Stop_voltcheck = True

End Sub

Public Sub ExitButton_Click()

    If Flag_debug_Voltage <> 1 Then Call EndGPIB_AD7461A
    
    Worksheets(ReadSheetName).Outline.ShowLevels RowLevels:=1
    
    Call Get_Power_Condition
    Call Unload(frm_PALS_VoltAdj_Main)

End Sub

Private Sub UserForm_Activate()
    Dim hWnd As Long
    Dim Wnd_STYLE As Long

    hWnd = GetActiveWindow()
    Wnd_STYLE = GetWindowLong(hWnd, GWL_STYLE)
    Wnd_STYLE = Wnd_STYLE And (Not WS_SYSMENU)
    SetWindowLong hWnd, GWL_STYLE, Wnd_STYLE
    DrawMenuBar hWnd
    Me.Caption = VOLTTOOLNAME & " Ver:" & VOLTTOOLVER

End Sub

Private Sub UserForm_Initialize()
    Call sub_PalsFileCheck(PALS_PARAMFOLDERNAME_VOLT)
    
    FileAddress = PALS_ParamFolder & "\" & PALS_PARAMFOLDERNAME_VOLT & "\" & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & "_#" & CStr(Sw_Node) & "\VoltCheck_" & Left(ThisWorkbook.Name, val(InStr(1, ThisWorkbook.Name, ".xls")) - 1) & _
                            "_#" & CStr(Sw_Node) & "_" & Format(Date, "yyyymmdd") & ".txt"
    DataLogCheck.ControlTipText = FileAddress

    If g_blnUseCSV = False Then Me.cb_savecsv.Visible = False

End Sub
