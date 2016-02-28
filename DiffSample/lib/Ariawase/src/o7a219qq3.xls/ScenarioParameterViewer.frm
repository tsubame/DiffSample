VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScenarioParameterViewer 
   Caption         =   "ScenarioParameterViewer"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4875
   OleObjectBlob   =   "ScenarioParameterViewer.frx":0000
End
Attribute VB_Name = "ScenarioParameterViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








'概要:
'   パラメータ表示用フォーム
'
'目的:
'   ライターでダンプされた情報を表示する
'
'作成者:
'   0145184306
'
Option Explicit

Private m_Active As Boolean
Private m_EndStatus As Boolean

Public Sub Display()
'内容:
'   ライターでダンプされた情報を表示する。
'
'備考:
'
'
    Show vbModeless
    m_Active = True
    m_EndStatus = False
    While m_Active = True
        DoEvents
    Wend
End Sub

Private Sub btnEnd_Click()
'内容:
'   強制終了ボタン
'   押された場合は、強制終了フラグをTrueにする
'
'備考:
'
'
    m_Active = False
    m_EndStatus = True
    Me.ScenarioParamView.Value = ""

End Sub

Private Sub btnContinue_Click()
'内容:
'   OKボタン
'
'備考:
'
'
    m_Active = False
    m_EndStatus = False
    Me.ScenarioParamView.Value = ""
    
End Sub


Private Sub QuitEnable_Change()
'内容:
'   強制終了ボタンのON/OFFを切り替える
'
'備考:
'
'
    If QuitEnable = True Then
        btnEnd.enabled = True
    Else
        btnEnd.enabled = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'内容:
'   "X"ボタンが押された時の動作
'
'備考:
'
'
    If CloseMode = vbFormControlMenu Then
        btnContinue_Click
        Cancel = True
    End If
    Me.ScenarioParamView.Value = ""

End Sub

Property Get EndStatus() As Boolean
'内容:
'   強制終了フラグプロパティ
'
'戻り値:
'   強制終了フラグ(Boolean型)
'
'備考:
'
'
    EndStatus = m_EndStatus
End Property

Property Let EndStatus(pStatus As Boolean)
'内容:
'   強制終了フラグプロパティ
'
'引数:
'   強制終了フラグ(Boolean型)
'
'備考:
'
'
    m_EndStatus = pStatus
End Property

Private Sub UserForm_Initialize()
'内容:
'   コンストラクタ
'
'備考:
'
'
    With Me
        .ScenarioParamView.Text = ""
        .ScenarioParamView.Locked = True
        .QuitEnable = False
        btnEnd.enabled = False
    End With
    m_Active = True
    m_EndStatus = False
End Sub

Private Sub UserForm_Terminate()
'内容:
'   デストラクタ
'
'備考:
'
'
    With Me
        .ScenarioParamView.Text = ""
        .QuitEnable = False
        btnEnd.enabled = False
    End With
    m_Active = False
    m_EndStatus = False
End Sub



