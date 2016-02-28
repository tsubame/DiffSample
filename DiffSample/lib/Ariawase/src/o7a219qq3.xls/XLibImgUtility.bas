Attribute VB_Name = "XLibImgUtility"
'概要:
'   TheIDPのユーティリティ
'
'目的:
'   TheIDP:CImgIDPの初期化/破棄のユーテリティを定義する
'
'作成者:
'   a_oshima

Option Explicit

Public Type T_PIXINFO
    x As Long
    y As Long
    Value As Double
End Type

Public TheIDP As CImgIDP

Public Const EEE_COLOR_FLAT As String = "EEE_COLOR_FLAT"
Public Const EEE_COLOR_ALL As String = "EEE_COLOR_ALL"

Private mSaveFileName As String

Public Sub SetLogModeTheIDP(ByVal pEnableLoggingTheIDP As Boolean, Optional saveFileName As String)
    
    mSaveFileName = saveFileName
    TheIDP.SaveMode = pEnableLoggingTheIDP
    TheIDP.saveFileName = saveFileName

End Sub

Public Function TempPMD(ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long) As CImgPmdInfo
    Set TempPMD = New CImgPmdInfo
    Call TempPMD.Create("", pX, pY, pWidth, pHeight)
End Function

Public Sub CreateTheIDPIfNothing()
'内容:
'   As Newの代替として初回に呼ばせるインスタンス生成処理
'
'パラメータ:
'   なし
'
'注意事項:
'
    On Error GoTo ErrHandler
    If TheIDP Is Nothing Then
        Set TheIDP = New CImgIDP
    End If
    Exit Sub
ErrHandler:
    Set TheIDP = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub DestroyTheIDP()
    Set TheIDP = Nothing
End Sub

Public Function RunAtJobEnd() As Long
    If Not TheIDP Is Nothing Then
        TheIDP.PlaneBank.IsOverwriteMode = False
        TheIDP.PlaneList.Clear
        TheIDP.saveFileName = ""
        TheIDP.SaveMode = False
    End If
End Function

