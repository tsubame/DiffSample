Attribute VB_Name = "XLibIGXLreference"
Option Explicit

' IG-XL が稼動状態かを調査する関数（コンパイルエラー防止策） '2011.9.1
' テラダインからの情報がベース
Public Function IsIGXLrunning() As Boolean

    Const TargetCaption As String = "Teradyne IG-XL DataTool"  'IG-XL稼動時はこの文字列を含む
    
    Dim ExcelCaptionName As String
    ExcelCaptionName = Application.Caption
    
    If InStr(ExcelCaptionName, TargetCaption) > 0 Then
'        MsgBox "Caption情報－IGXL datatool 稼働中です"
        IsIGXLrunning = True
    Else
'        MsgBox "Caption情報－IGXL datatool は稼動していません"
        IsIGXLrunning = False
    
    End If

End Function

Public Sub Close_EeeJOB()

    If IsIGXLrunning Then
        Call TheIDP_Destory
#If ITS <> 0 Then
        Call XLibImpUIControllerUtility.DestroyImpUIController
#End If
    End If
End Sub
