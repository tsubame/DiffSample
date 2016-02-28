Attribute VB_Name = "XLibImpExperimentUtility"
'概要:
'   ImSceExperimentControllerのUtility
'
'目的:
'   ImSceExperimentControllerの初期化/破棄のユーテリティを定義する
'
'作成者:
'   0145184306
'
Option Explicit

Public Const RESTART_ERROR_NUMBER As Long = 8000

Private mEnableExperimentMode As Boolean

Public Property Get EnableExperimentMode() As Boolean
'内容:
'   実験機能の状態取得
'
'備考:
'
'
    EnableExperimentMode = mEnableExperimentMode
End Property

Public Property Let EnableExperimentMode(ByVal pEnable As Boolean)
'内容:
'   実験機能の状態設定
'
'備考:
'
'
    mEnableExperimentMode = pEnable
End Property

Public Function GetSubParamLabel(ByVal pPath As String, ByVal pCurPath As String) As String
'内容:
'   パラメータクラスのメンバ変数のラベルを取得する
'
'[pPath]       IN String型:     メンバ変数の絶対パス
'[pCurPath]    IN String型:     パラメータクラスのパス
'
'備考:
'
'
    Dim myLabel As String
    myLabel = Mid$(pPath, Len(pCurPath) + 1)
    If myLabel Like "\*" Then
        myLabel = Mid$(myLabel, 2)
    End If
    Dim myIndex As Long
    myIndex = InStr(myLabel, "\")
    If myIndex > 0 Then
        myLabel = Left$(myLabel, myIndex - 1)
    End If
    GetSubParamLabel = myLabel
End Function

Public Function GetSubParamIndex(ByVal pPath As String, ByVal pCurPath As String) As Long
    GetSubParamIndex = CLng(Mid$(Strings.Left$(pPath, InStr(Len(pCurPath) + 1, pPath, ")") - 1), InStr(Len(pCurPath) + 1, pPath, "(") + 1))
End Function
