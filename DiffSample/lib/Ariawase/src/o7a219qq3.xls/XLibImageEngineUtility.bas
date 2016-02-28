Attribute VB_Name = "XLibImageEngineUtility"
'�T�v:
'   TheImageTest�̃��[�e�B���e�B
'
'�ړI:
'   TheImageTest:CImageEngine�̏�����/�j���̃��[�e���e�B���`����
'
'�쐬��:
'   a_oshima

Option Explicit

Public TheImageTest As CImageEngine

Public Sub CreateTheImageTestIfNothing()
'���e:
'   As New�̑�ւƂ��ď���ɌĂ΂���C���X�^���X��������
'
'�p�����[�^:
'   �Ȃ�
'
'���ӎ���:
'
    On Error GoTo ErrHandler
    If TheImageTest Is Nothing Then
        '### TheImageTest�̏����� ###################
        Set TheImageTest = New CImageEngine
        With TheImageTest
            .Initialize GetActionLoggerInstance, GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeTestInstances)
            If .CreateScenario = TL_ERROR Then
                TheError.Raise 9999, "XLibImageEngineUtility.CreateTheImageTestIfNothing", "CreateScenario returned TL_ERROR"
            End If
        End With
    End If
    Exit Sub
ErrHandler:
    Set TheImageTest = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub DestroyTheImageTest()
    Set TheImageTest = Nothing
End Sub

Public Function RunAtJobEnd() As Long

End Function

Public Sub EnableInterceptor(pFlag As Boolean, pLogger As CActionLogger)
    Call TheImageTest.EnableInterceptor(pFlag, pLogger)
End Sub
