Attribute VB_Name = "XLibImgPmdSheet"
'�T�v:
'   ###���̃��W���[���̖����Ȃǂ��P�`�Q�s�ł킩��悤�ɋL�q���Ă�������###
'
'�ړI:
'   ###�ړI���̏ڍׂ��L�q���Ă�������###
'
'�쐬��:
'   0145184004
'
Option Explicit

Private m_PmdSht As CImgPmdSheet

Public Sub Initialize(ByVal pShtName As String)
'���e:
'   �w�肵���V�[�g�̋@�\��L���ɂ���B
'
'[pShtName]    IN   String:         �Ώۂ̃V�[�g��
'
'���l:
'
    Set m_PmdSht = New CImgPmdSheet
    Set m_PmdSht.targetSheet = Worksheets(pShtName)
End Sub

Public Sub CreatePMD(ByVal pShtName As String)
'���e:
'   �w�肵���V�[�g�̃f�[�^�ɏ]���āAPMD���쐬����B
'
'[pShtName]    IN   String:         �Ώۂ̃V�[�g��
'
'���l:
'
    Call Initialize(pShtName)
    Application.StatusBar = "Creating Base PMD..."
    Call m_PmdSht.CreatePMD
    Application.StatusBar = False
End Sub

Public Sub AddPmdSheet()
'���e:
'   PMD�V�[�g��Job�ɒǉ�����B
'
'���l:
'
        
    Dim shtEnd As Worksheet
    
    With Worksheets
        Set shtEnd = .Item(.Count)
    End With
    
    Call ShtPMD.Copy(, shtEnd)
    
End Sub

Public Sub CreatePMDIfNothing()
    On Error GoTo ErrHandler
    If TheIDP.PlaneManagerCount = 0 Then
        Call CreatePMD(GetWkShtReaderManagerInstance.GetActiveSheetName(shtTypePMDDefinition))
    End If
    Exit Sub
ErrHandler:
    DestroyPMDSheet
    DestroyTheIDP
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub DestroyPMDSheet()
    Set m_PmdSht = Nothing
End Sub
