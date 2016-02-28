VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_PALS 
   Caption         =   "PALS     ParameterAuto-adjustLinkSystem"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   OleObjectBlob   =   "frm_PALS.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frm_PALS"
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


Private Sub cmd_exit_Click()
    Unload frm_PALS
End Sub

Public Sub sub_ShowForm(ByVal strTargetForm As String)
    
'    Set objLoadedJob = Nothing
'    Set objLoadedJob = GetObject(, "excel.application")
    
    Select Case strTargetForm
'>>>2011/6/24 M.IMAMURA Mod.
'        Case PALS_PARAMFOLDERNAME_BIAS: frm_PALS_BiasAdj_Main.Show
'        Case PALS_PARAMFOLDERNAME_LOOP: frm_PALS_LoopAdj_Main.Show
'        Case PALS_PARAMFOLDERNAME_OPT: frm_PALS_OptAdj_Main.Show
'        Case PALS_PARAMFOLDERNAME_TRACE: frm_PALS_TraceAdj_Main.Show
'        Case PALS_PARAMFOLDERNAME_VOLT: frm_PALS_VoltAdj_Main.Show
'        Case PALS_PARAMFOLDERNAME_WAIT: frm_PALS_WaitAdj_Main.Show
'        Case PALS_PARAMFOLDERNAME_WAVE: frm_PALS_WaveAdj_Main.Show
        
        Case PALS_PARAMFOLDERNAME_BIAS: Call Excel.Application.Run("sub_BiasFrmShow")
        Case PALS_PARAMFOLDERNAME_LOOP: Call Excel.Application.Run("sub_LoopFrmShow")
        Case PALS_PARAMFOLDERNAME_OPT: Call Excel.Application.Run("sub_OptFrmShow")
        Case PALS_PARAMFOLDERNAME_TRACE: Call Excel.Application.Run("sub_TraceFrmShow")
        Case PALS_PARAMFOLDERNAME_VOLT: Call Excel.Application.Run("sub_VoltFrmShow")
        Case PALS_PARAMFOLDERNAME_WAIT: Call Excel.Application.Run("sub_WaitFrmShow")
        Case PALS_PARAMFOLDERNAME_WAVE: Call Excel.Application.Run("sub_WaveFrmShow")
'<<<2011/6/24 M.IMAMURA Mod.
        
    End Select
End Sub

Private Sub cmd_BiasRun_Click()
    If FLG_PALS_DISABLE.BiasAdj = True Then
        MsgBox cmd_BiasRun.ControlTipText, vbCritical, PALS_ERRORTITLE
        Exit Sub
    End If
    
    FLG_PALS_RUN.BiasAdj = True
    Me.Hide
    
    If FLG_PALS_DISABLE.BiasAdj = False Then Call sub_ShowForm(PALS_PARAMFOLDERNAME_BIAS)
    
    FLG_PALS_RUN.BiasAdj = False
    Me.Show

End Sub

Private Sub cmd_LoopRun_Click()
    If FLG_PALS_DISABLE.LoopAdj = True Then
        MsgBox cmd_LoopRun.ControlTipText, vbCritical, PALS_ERRORTITLE
        Exit Sub
    End If
    
    'フラグの初期化
    g_blnLoopStop = False
    g_ErrorFlg_PALS = False
    
    FLG_PALS_RUN.LoopAdj = True
    Me.Hide
    
    If FLG_PALS_DISABLE.LoopAdj = False Then Call sub_ShowForm(PALS_PARAMFOLDERNAME_LOOP)
    
    FLG_PALS_RUN.LoopAdj = False
    Me.Show

End Sub

Private Sub cmd_TraceRun_Click()
    If FLG_PALS_DISABLE.TraceAdj = True Then
        MsgBox cmd_TraceRun.ControlTipText, vbCritical, PALS_ERRORTITLE
        Exit Sub
    End If
    
    'フラグの初期化
    g_blnLoopStop = False
    g_ErrorFlg_PALS = False
    
    FLG_PALS_RUN.TraceAdj = True
    Me.Hide

    If FLG_PALS_DISABLE.TraceAdj = False Then Call sub_ShowForm(PALS_PARAMFOLDERNAME_TRACE)

    FLG_PALS_RUN.TraceAdj = False
    Me.Show

End Sub


Private Sub cmd_WaitRun_Click()
    If FLG_PALS_DISABLE.WaitAdj = True Then
        MsgBox cmd_WaitRun.ControlTipText, vbCritical, PALS_ERRORTITLE
        Exit Sub
    End If
    
    'フラグの初期化
    g_blnLoopStop = False
    g_ErrorFlg_PALS = False
    
    FLG_PALS_RUN.WaitAdj = True
    Me.Hide
    
    If FLG_PALS_DISABLE.WaitAdj = False Then Call sub_ShowForm(PALS_PARAMFOLDERNAME_WAIT)
    
    FLG_PALS_RUN.WaitAdj = False
    Me.Show

End Sub

Private Sub cmd_OptRun_Click()
    If FLG_PALS_DISABLE.OptAdj = True Then
        MsgBox cmd_OptRun.ControlTipText, vbCritical, PALS_ERRORTITLE
        Exit Sub
    End If
    
    FLG_PALS_RUN.OptAdj = True
    Me.Hide
    
    If FLG_PALS_DISABLE.OptAdj = False Then Call sub_ShowForm(PALS_PARAMFOLDERNAME_OPT)
    
    FLG_PALS_RUN.OptAdj = False
    Me.Show
End Sub

Private Sub cmd_VoltageRun_Click()
    If FLG_PALS_DISABLE.VoltageAdj = True Then
        MsgBox cmd_VoltageRun.ControlTipText, vbCritical, PALS_ERRORTITLE
        Exit Sub
    End If
    
    FLG_PALS_RUN.VoltageAdj = True
    Me.Hide
    
    If FLG_PALS_DISABLE.VoltageAdj = False Then Call VoltCheck_start
    
    FLG_PALS_RUN.VoltageAdj = False
    Me.Show

End Sub

Private Sub cmd_WaveRun_Click()
    If FLG_PALS_DISABLE.WaveAdj = True Then
        MsgBox cmd_WaveRun.ControlTipText, vbCritical, PALS_ERRORTITLE
        Exit Sub
    End If
    
'    Set objLoadedJob = Nothing
'    Set objLoadedJob = GetObject(, "excel.application")

'>>> 2011/6/24 M.IMAMURA Mod.
'    If Check_OSCGPIB = False Then
    Call Excel.Application.Run("Check_OSCGPIB")
    If FLG_PALS_DISABLE.WaveAdj = True Then
        Exit Sub
    End If
'<<< 2011/6/24 M.IMAMURA Mod.
    
    FLG_PALS_RUN.WaveAdj = True
    Me.Hide
    
    If FLG_PALS_DISABLE.WaveAdj = False Then Call sub_ShowForm(PALS_PARAMFOLDERNAME_WAVE)
    
    FLG_PALS_RUN.WaveAdj = False
    Me.Show

End Sub


Private Sub UserForm_Activate()
    Dim hWnd As Long
    Dim Wnd_STYLE As Long
    
    hWnd = GetActiveWindow()
    Wnd_STYLE = GetWindowLong(hWnd, GWL_STYLE)
    Wnd_STYLE = Wnd_STYLE And (Not WS_SYSMENU)
    SetWindowLong hWnd, GWL_STYLE, Wnd_STYLE
    DrawMenuBar hWnd
    Me.Caption = PALSNAME & " Ver:" & PALSVER

    Dim intImg_S_Height As Integer
    Dim intImg_S_HeightPer As Integer

    If blnRunPals = False Then
        cmd_exit.Visible = False
        If blnPALS_ANI = True Then
            intImg_S_Height = Img_S.height
            Img_S.height = 1
            Img_S.Visible = True
            Pause 1
            For intImg_S_HeightPer = 10 To 100 Step 10
                Pause 0.2
                Img_S.height = Int(intImg_S_Height * 0.01 * intImg_S_HeightPer)
    '            DoEvents
            Next intImg_S_HeightPer
        Else
            Img_S.Visible = True
        End If
        
        Pause 1.5
    
    End If
    
    Img_PAL.Visible = False
    Img_S.Visible = False

    frame_runmenu.Visible = True
    cmd_exit.Visible = True
    blnRunPals = True
    Call sub_PalsAliveCheck

End Sub

Private Sub UserForm_Initialize()
    
On Error GoTo errPALSUserForm_Initialize
    '>>>2011/9/20 M.Imamura g_RunAutoFlg_PALS add. 念のため
    'sub_RunLoopAuto/sub_RunOptAutoでもやってるけどね
    g_RunAutoFlg_PALS = False
    '<<<2011/9/20 M.Imamura g_RunAutoFlg_PALS add.
    Call sub_PalsAliveCheck

Exit Sub

errPALSUserForm_Initialize:
    Call sub_errPALS("frm_PALS not initialize at 'UserForm_Initialize'", "0-1-01-0-01")

End Sub
Private Sub UserForm_Terminate()
    Set PALS = Nothing
End Sub

Private Sub sub_PalsAliveCheck()

    '>>>2011/06/14 M.IMAMURA RGB->VBA.RGB Mod.
    Me.cmd_OptRun.BackColor = VBA.RGB(255, 100, 100)
    Me.cmd_BiasRun.BackColor = VBA.RGB(255, 100, 100)
    Me.cmd_TraceRun.BackColor = VBA.RGB(255, 100, 100)
    Me.cmd_LoopRun.BackColor = VBA.RGB(255, 100, 100)
    Me.cmd_WaitRun.BackColor = VBA.RGB(255, 100, 100)
    Me.cmd_VoltageRun.BackColor = VBA.RGB(255, 100, 100)
    Me.cmd_WaveRun.BackColor = VBA.RGB(255, 100, 100)
    '<<<2011/06/14 M.IMAMURA RGB->VBA.RGB Mod.

    If FLG_PALS_DISABLE.OptAdj = False Then
        Me.cmd_OptRun.BackColor = vbCyan
    End If
    
    If FLG_PALS_DISABLE.VoltageAdj = False Then
        Me.cmd_VoltageRun.BackColor = vbCyan
    End If
    
    If FLG_PALS_DISABLE.TraceAdj = False Then
        Me.cmd_TraceRun.BackColor = vbCyan
    End If

    If FLG_PALS_DISABLE.LoopAdj = False Then
        Me.cmd_LoopRun.BackColor = vbCyan
    End If
    
    If FLG_PALS_DISABLE.WaitAdj = False Then
        Me.cmd_WaitRun.BackColor = vbCyan
    End If
    
    If FLG_PALS_DISABLE.WaveAdj = False Then
        Me.cmd_WaveRun.BackColor = vbCyan
    End If
    
    If FLG_PALS_DISABLE.BiasAdj = False Then
        Me.cmd_BiasRun.BackColor = vbCyan
    End If


End Sub
