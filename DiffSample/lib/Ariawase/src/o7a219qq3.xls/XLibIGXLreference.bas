Attribute VB_Name = "XLibIGXLreference"
Option Explicit

' IG-XL ���ғ���Ԃ��𒲍�����֐��i�R���p�C���G���[�h�~��j '2011.9.1
' �e���_�C������̏�񂪃x�[�X
Public Function IsIGXLrunning() As Boolean

    Const TargetCaption As String = "Teradyne IG-XL DataTool"  'IG-XL�ғ����͂��̕�������܂�
    
    Dim ExcelCaptionName As String
    ExcelCaptionName = Application.Caption
    
    If InStr(ExcelCaptionName, TargetCaption) > 0 Then
'        MsgBox "Caption���|IGXL datatool �ғ����ł�"
        IsIGXLrunning = True
    Else
'        MsgBox "Caption���|IGXL datatool �͉ғ����Ă��܂���"
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
