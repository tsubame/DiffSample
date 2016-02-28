Attribute VB_Name = "XLibOutputInfoMod"
Option Explicit

'{
'���̏o�͂��s���@�\��L����֐��ނ́A���̃��W���[���ɏW�߂���
'}

Private dbgInfHeaderLength As Long  '�f�o�b�N���o�̓w�b�_�̑��������i�[�p

Public Sub WriteDefectInfoHeader(ByVal testName As String, ByVal SiteNumber As Integer)
'���e:
'   �_���׏��̃w�b�_���f�[�^���O�֏o�͂���
'
'[testName]   In  �\��������e�X�g����
'[siteNumber] In  �\��������T�C�g�ԍ�
'
    Dim infMsg As String

    infMsg = "***** " & testName & " DEFECT ADDRESS & DATA (SITE:" & SiteNumber & ") *****"
    Call WriteComment(infMsg)

End Sub

Public Sub WriteDebugInfoHeader(ByVal infoMessage As String, ByVal SiteNumber As Integer)
'���e:
'   �f�o�b�N���̃w�b�_���f�[�^���O�֏o�͂���
'
'[infoMessage] In  �\��������f�o�b�N��񖼏�
'[siteNumber]  In  �\��������T�C�g�ԍ�
'
'���ӎ���
'dbgInfHeaderLength�ϐ��ɏo�͕��������i�[����
'
    Dim infoMsg As String

    infoMsg = "***** " & infoMessage & " (SITE:" & SiteNumber & ") *****"
    
    dbgInfHeaderLength = Len(infoMsg)
    
    Call WriteComment(infoMsg)

End Sub

Public Sub WriteDebugInfoFooter()
'���e:
'   �f�o�b�N���̃t�b�^���f�[�^���O�֏o�͂���
'
'���ӎ���:
'  WriteDebugInfoHeader�T�u���[�`���ƃy�A�Ŏg�p����
'  dbgInfHeaderLength�ϐ��̒l���g�p����
'
    Dim msgCounter As Long
    Dim outputFooter As String
        
    For msgCounter = 1 To dbgInfHeaderLength Step 1
        outputFooter = outputFooter & "*"
    Next msgCounter
        
    Call WriteComment(outputFooter)

End Sub

Public Sub WriteComment(ByVal outPutMsg As String, Optional ByVal outPutFileName As String = "")
'���e:
'   �R�����g�����f�[�^���O�֏o�͂���
'
'[outPutMsg]       In  �f�[�^���O�֏o�͂��郁�b�Z�[�W
'[outPutFileName]  In  �t�@�C���֏o�͂��鎞�̃t�@�C����
'
'���ӎ���:
'   outPutFileName�̓I�v�V�����B
'   �w�肪�Ȃ��ꍇ�t�@�C���ւ̏��o�͍͂s���Ȃ�
'
    Call mf_OutPutComment(outPutMsg)
    
    If outPutFileName <> "" Then
        Call mf_AppendTxtFile(outPutFileName, outPutMsg)
    End If

End Sub

Private Sub mf_OutPutComment(ByVal outPutMsg As String)
'���e:
'   �����f�[�^���OWindow�֏o�͂���
'
'[OutPutMsg] In  �\�������郁�b�Z�[�W
'
    TheExec.Datalog.WriteComment outPutMsg

End Sub

Private Function mf_AppendTxtFile(ByVal appendFileName As String, outPutMsg As String) As Boolean
'���e:
'   �����w�肳�ꂽ�e�L�X�g�t�@�C���֒ǋL����
'
'[appendFileName] In  ����ǋL����t�@�C����
'[outPutMsg]      In  �ǋL���郁�b�Z�[�W
'
'�߂�l�F
'   ���s���ʃX�e�[�^�X
'    �G���[�Ȃ��FTrue
'    �G���[����FFalse
'
    Dim fileNum As Integer
    Dim errFunctionName As String
    
    On Error GoTo OUT_PUT_LOG_ERR
    errFunctionName = mf_AppendTxtFile
    
    fileNum = FreeFile
    Open appendFileName For Append As fileNum
    Print #fileNum, outPutMsg
    Close fileNum
    
    mf_AppendTxtFile = True
    
    Exit Function

OUT_PUT_LOG_ERR:
    Call MsgBox(appendFileName & " Output File Error", vbFalse Or vbCritical, "@" & errFunctionName)
    mf_AppendTxtFile = False
'   Stop

End Function
