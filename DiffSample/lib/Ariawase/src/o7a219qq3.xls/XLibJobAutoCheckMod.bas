Attribute VB_Name = "XLibJobAutoCheckMod"
Option Explicit
'{
'���؎������؃��[�h����pFLAG  0:NOT�������؁A1:�������� (�Ƃ肠�����A�R�R�ɓ���Ă��邪��Ō�����)
Public Flg_JobAutoCheck As Integer

Private Const IDP_EXTENSION_NAME = ".idp" '�摜�t�@�C���̊g���q

'�O���t�@�C���Ǎ��惉�x�����̒�`�p
Private Const JOBAUTOCHK_SHEETNAME = "JobAutoCheck"           '���[�N�V�[�g�̖��O
Private Const LABEL_IDP_LOAD_PATH = "IDP_LOAD_PATH"           '�摜�t�@�C���̃��x��
Private Const LABEL_SNAPSHOT_FILENAME = "SNAPSHOT_FILE_NAME"  '�X�i�b�v�V���b�g�̃��x��
Private Const LABEL_COLUMN = 1                                '���x���L���s
Private Const VALUE_COLUMN = 2                                '�l�L���s

'�X�i�b�v�V���b�g�擾�p�ϐ�
Private sampleNumber As Long
Private snapFileName As String        '�e�X�^���X�i�b�v�V���b�g�ۑ��p�t�@�C����

'�O���t�@�C���̕ۑ�����p�ϐ�
Private idpLoadPath As String          '��荞�݉摜�AADCEOEF�̕ۑ���
Private snapShotLogSavePath As String  '�X�i�b�v�V���b�g���O���t�@�C���ɏo�͂���Ƃ��̏o��

'��荞�݉摜�̕ۑ����Path�������J
Public Function GetIdpPath() As String
    GetIdpPath = idpLoadPath
End Function

'���[�U�[�����͂����T���v���ԍ��l�����J
Public Function GetSampleNumber() As Long
    GetSampleNumber = sampleNumber
End Function

'JobAutoCheck�X�^�[�g�ݒ�
Public Sub InitJobAutoCheck(ByVal MinChipNumber As Long, ByVal Max_ChipNumber As Long)
    
'�e�X�^���[�h���擾���āA���@���V�~�����[�^�����m�F����B
'testModeOffline:�V�~�����[�^�AtestModeOnline�F�e�X�^�[���@
    If TheExec.TesterMode = testModeOffline Then

        Flg_JobAutoCheck = 1 '�b��{�� ����͂��Ƃŉ��Ƃ�������
            
        '���؂Ɏg�p����O���t�@�C���̒�`���V�[�g����擾
        Call mf_GetFileLoadPath
            
        Dim inputChipNumber As Variant
        inputChipNumber = 0
        '�C���v�b�g�{�b�N�X�ɁA�g�p����_�~�[�摜�̃T���v���ԍ�����́B
        Do
            inputChipNumber = InputBox("Enter CHIP number (" & MinChipNumber & "-" & Max_ChipNumber & "): ", "CHIP Number Input")
        Loop While (inputChipNumber <= 0 Or inputChipNumber > Max_ChipNumber Or inputChipNumber < MinChipNumber)
        
        sampleNumber = CLng(inputChipNumber)
                         
        If mf_Get_JAC_SheetVal("TESTER_SNAPSHOT_SAVE") Then
            Call mf_deleteFile(snapFileName)    '�X�i�b�v�V���b�g�ۑ��p�t�@�C�������|��
        End If
    Else
        Flg_JobAutoCheck = 0 '�b��{�� ����͂��Ƃŉ��Ƃ�������
    End If

End Sub

'Measure���ɂȂɂ���肽�����p(���́A�e�X�^Snap�̂�)
Public Sub CheckMeasureStatus(Optional idLabel As String = "")
    DoEvents
    Call mf_saveSnapShot(idLabel)  '�e�X�^�X�i�b�v�V���b�g���擾������
End Sub

'�_�~�[�摜�̃t�@�C�������t���p�X�w��ŕԂ��܂�
Public Sub GetIdpFileName(ByVal siteNum As Long, ByVal TestLabel As String, ByRef IdpFilePath As String)
    IdpFilePath = idpLoadPath & sampleNumber & "\" & TestLabel & "_" & siteNum & IDP_EXTENSION_NAME
End Sub

'�C���[�W�v���[���̃f�[�^���t�@�C������ǂݍ��݂܂�
Public Sub ReadIdpFile(InputPlaneName As String, _
    BasePmdName As String, _
    IdpFileName As String, _
    Optional InputSiteNumber As Long = ALL_SITE, _
    Optional IdpFileType As IdpFileFormat = idpFileBinary)

    Dim idpLogMsg As String

    TheHdw.IDP.SetPMD InputPlaneName, BasePmdName
    TheHdw.IDP.ReadFile InputSiteNumber, InputPlaneName, idpColorFlat, IdpFileName, IdpFileType
    TheHdw.IDP.SetPMD InputPlaneName, BasePmdName

    '�摜�t�@�C���ǂݍ��ݎ��̃��O�o�͗p
    #If IDP_READ_LOG = 1 Then
        idpLogMsg = "IDP_READ," & "Instances=" & _
        TheExec.DataManager.InstanceName & _
        ",Plane=" & InputPlaneName & _
        ",Site=" & InputSiteNumber & _
        ",File=" & IdpFileName
        Call WriteComment(idpLogMsg)
    #End If

End Sub

'�f�[�^���Owindow�ɏo�͂���ADebug�����K���Format�i������r�p�j�ŏo�͂��邽�߂̃T�u���[�`�� ���ʗp
Public Sub OutputDebugInfo(ByVal testCategory As String, _
    ByVal testName As String, _
    ByVal SiteNumber As Long, _
    ByVal OutputValue As Double, _
    Optional UnitLabel As String = "")

    Dim outPutMsg As String

    outPutMsg = "#" & testCategory & ":" & testName & ":" & SiteNumber & ":" & " = " & OutputValue & "" & UnitLabel

    Call WriteComment(outPutMsg)

End Sub

'�f�[�^���Owindow�ɁA�����̐ݒ��Ԃ��o�͂��邽�߂̃T�u���[�`��
Public Sub OutputOptsetInfo(ByVal category As String, ByVal TestInstanceName As String, ByVal CommandListIdentifier As String)
    
    Dim outPutMsg As String
    
    outPutMsg = "#" & category & ":" & TestInstanceName & ":" & " = " & CommandListIdentifier
    
    Call WriteComment(outPutMsg)

End Sub

'�������ؗpAPMU SnapShot���s�p
Private Sub mf_saveSnapShot(Optional idLabel As String = "")
    
    If mf_Get_JAC_SheetVal("TESTER_SNAPSHOT_SAVE") Then
        '�e�X�^�X�i�b�v�V���b�g���擾���A���ʂ��t�@�C���ɕۑ�����
        Call TheSnapshot.GetSnapshot(idLabel)
    Else
'        MsgBox ("�X�i�b�v�V���b�g��ۑ�����ƌ���ꂽ��" & "TESTER_SNAPSHOT_SAVE" & "��1�ł͂Ȃ��̂ŉ������Ȃ�")
    End If

End Sub

'JobAutoCheck�V�[�g�̎w�胉�x���ׂ̗̒l���擾���Ă���
Private Function mf_Get_JAC_SheetVal(ByVal LabelName As String) As Long
    mf_Get_JAC_SheetVal = mf_GetCellValue(JOBAUTOCHK_SHEETNAME, mf_SearchWordRow(LabelName), VALUE_COLUMN)
End Function

'�Ώۃ��[�N�V�[�g�̑��݊m�F
Private Function mf_ChkJacSheet() As Integer

    Dim sheetChkFlg As Boolean
    Dim sheetCnt As Long

    For sheetCnt = 1 To ThisWorkbook.Worksheets.Count
        ' �V�[�g������������t���O�ݒ�
        If ThisWorkbook.Worksheets(sheetCnt).Name = JOBAUTOCHK_SHEETNAME Then
            sheetChkFlg = True
            mf_ChkJacSheet = 0
        End If
    Next
    
    '�V�[�g��������Ȃ������ꍇ�̏���
    If sheetChkFlg = False Then
        MsgBox JOBAUTOCHK_SHEETNAME & " ���[�N�V�[�g�����݂��܂���", vbCritical, "SHEET CHECK ERROR"
        mf_ChkJacSheet = 1
        Stop
    End If

End Function

'�w��L�[���[�h�����݂���Cell�̍s�ԍ����擾
Private Function mf_SearchWordRow(ByVal searchWord As String) As Long
    
    Dim tmpRange As Range
    
    With Worksheets(JOBAUTOCHK_SHEETNAME).Columns(LABEL_COLUMN)
        
        Set tmpRange = .Find(searchWord)
        
        If Not tmpRange Is Nothing Then
           'MsgBox "���������L�[���[�h�����݂���̂�" & tmpRange.Cells.Row & "�s�ڂł�"
            mf_SearchWordRow = tmpRange.Cells.Row
        Else
            MsgBox JOBAUTOCHK_SHEETNAME & "�V�[�g��A���" & searchWord & "�����݂��܂���", vbCritical, "SHEET SEARCH ERROR"
            Stop
        End If
        
    End With
    
    Set tmpRange = Nothing

End Function

'�w��V�[�g�̎w��CELL�̒l�����
Private Function mf_GetCellValue(ByVal workSheetName As String, ByVal cellsColumn As Long, ByVal cellsRow As Long) As Variant
    mf_GetCellValue = Worksheets(workSheetName).Cells(cellsColumn, cellsRow)
End Function

'JobAutoCheck���[�N�V�[�g�̃t�@�C���̕ۑ���l���擾���ϐ��Ɋi�[
Private Sub mf_GetFileLoadPath()

    Dim idpRow As Long
    Dim snapRow As Long
    
    Call mf_ChkJacSheet
    
    idpRow = mf_SearchWordRow(LABEL_IDP_LOAD_PATH)     '�_�~�[�摜�̕ۑ���̒�`�s
    snapRow = mf_SearchWordRow(LABEL_SNAPSHOT_FILENAME) '�X�i�b�v�V���b�g�ۑ��t�@�C���̒�`�s
        
    idpLoadPath = mf_GetCellValue(JOBAUTOCHK_SHEETNAME, idpRow, VALUE_COLUMN)
    snapFileName = mf_GetCellValue(JOBAUTOCHK_SHEETNAME, snapRow, VALUE_COLUMN)

End Sub

'�X�i�b�v�V���b�g�̃��O�����|������
Private Sub mf_deleteFile(ByVal delFileName As String)
    
    On Error GoTo FileDelErr
    
    Call Kill(delFileName)
    Exit Sub

FileDelErr:
'    MsgBox "�w�肳�ꂽ�t�@�C���͂Ȃ������̂ł��|�����Ȃ��čς�"

End Sub

'�X�i�b�v�V���b�g�p�̃��O���t�@�C���ɏo�͂���B
Private Sub mf_OutPutLog(ByVal LogFileName As String, outPutMessage As String)
    Dim fp As Integer
    On Error GoTo OUT_PUT_LOG_ERR
    
    fp = FreeFile
    Open LogFileName For Append As fp
    Print #fp, outPutMessage
    Close fp
    
    Exit Sub

OUT_PUT_LOG_ERR:
    Call MsgBox(LogFileName & " MsgOutPut Error", vbFalse Or vbCritical, "@mf_OutPutLog")
    Stop

End Sub

'�X�i�b�v�V���b�g�̕ۑ����������
Public Function GetSnapFilename() As String
    
'    If snapFileName = "" Then
        Call mf_GetFileLoadPath
'    End If
    
    GetSnapFilename = snapFileName

End Function

'�X�i�b�v�V���b�g�擾�t���O�̊m�F
Public Function IsSnapshotOn() As Boolean

    If mf_Get_JAC_SheetVal("TESTER_SNAPSHOT_SAVE") <> 0 Then
        IsSnapshotOn = True
    Else
        IsSnapshotOn = False
    End If

End Function

'}

