Attribute VB_Name = "XLibImgUtility"
'�T�v:
'   TheIDP�̃��[�e�B���e�B
'
'�ړI:
'   TheIDP:CImgIDP�̏�����/�j���̃��[�e���e�B���`����
'
'�쐬��:
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
'���e:
'   As New�̑�ւƂ��ď���ɌĂ΂���C���X�^���X��������
'
'�p�����[�^:
'   �Ȃ�
'
'���ӎ���:
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

