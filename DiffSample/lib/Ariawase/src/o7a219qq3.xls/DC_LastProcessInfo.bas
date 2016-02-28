Attribute VB_Name = "DC_LastProcessInfo"
'�T�v:
'�@�O�H���ō쐬���ꂽ�s�Ǐ��t�@�C����ǂݍ���
'
'�ړI:
'�@�O�H�������ɂ�NG�ɂȂ����`�b�v�������s�ǂƂ���B
'
'�쐬��:
'   2012/10/16 Ver0.1 A.Hamaya
'   2013/01/16 Ver0.2 A.Hamaya
'   2013/01/24 Ver0.3 H.Arikawa �����z�肵��������ǉ��B
'
'
'�g�p����ɂ́A���L�R�[�h��dc_setup�֋L��
'
'    '### ������ ###
'    If Flg_AutoMode = True Then
'        If CInt(DeviceNumber_site(0)) = 1 Then  '�f�o�C�XNo.��1�̎�
'            Call Init_LastProcessInfoFILE
'        End If
'    End If
'

Option Explicit

#Const EEE_AUTO_JOB_LOCATE = 2      '1:����200mm,2:����300mm,3:�F�{

'### �֐���` ###
Public USonic(nSite) As Double
Public Wasavi(nSite) As Double
Public PadClo(nSite) As Double
Public Fmura(nSite) As Double
Public Proces(nSite) As Double

'### �e�H���̃i���o�[��` ###
Private Const USonicNum As Integer = 1
Private Const WasaviNum As Integer = 2
Private Const PadCloNum As Integer = 3
Private Const FmuraNum As Integer = 4

Private NowWaferID As String
Private NGchipCNT As Integer            'NG�`�b�v�̌�
Private NGdataCNT As Integer            '�f�[�^�̌�
Private NGChipNo() As String            'NG�`�b�vNo.�i�[�p
Private NGChipData() As String          'NG�`�b�v�f�[�^�i�[�p
Private FileStatus() As String          '�t�@�C����Ԋi�[�p exist/not-exist

Private flg_NoFILE As Boolean           '�G���h�t�@�C�������������ꍇ�ɗ��t���O
Private flg_NoEndFILE As Boolean        '�G���h�t�@�C�������������ꍇ�ɗ��t���O

Private LastProcessInfoFILE As String       '�t�@�C���{��
Private LastProcessInfoFILE_END As String   '�t�@�C���̃G���h�t�@�C��

Private Const LastProcessInfo_FilePATH_K As String = "f:\job\failchipdetection\"        '�O�H���s�Ǐ��t�@�C���̃p�X(�F�{)
Private Const LastProcessInfo_FilePATH_N As String = "f:\job\failchipdetection\"        '�O�H���s�Ǐ��t�@�C���̃p�X(����)�@�\��

Public Function ultrasonic_f() As Double

    Dim site As Long
    Call SiteCheck

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            USonic(site) = Get_NgChipInfo(USonicNum, site)
        End If
    Next site
    
    Call test(USonic)

End Function

Public Function wasavi_f() As Double

    Dim site As Long
    Call SiteCheck

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Wasavi(site) = Get_NgChipInfo(WasaviNum, site)
        End If
    Next site

    Call test(Wasavi)

End Function

Public Function padclosing_f() As Double

    Dim site As Long
    Call SiteCheck

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            PadClo(site) = Get_NgChipInfo(PadCloNum, site)
        End If
    Next site

    Call test(PadClo)

End Function

Public Function fmura_f() As Double

    Dim site As Long
    Call SiteCheck

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Fmura(site) = Get_NgChipInfo(FmuraNum, site)
        End If
    Next site

    Call test(Fmura)

End Function

'###�@�����s�ǃ`�b�v���������ꍇ�A���̊֐���Fail�ɂ���B ###

Public Function processng_f() As Double

    Dim site As Long
    Call SiteCheck

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then

            If USonic(site) = 1 Or Wasavi(site) = 1 Or PadClo(site) = 1 Or Fmura(site) = 1 Then
                Proces(site) = 1
            Else
                Proces(site) = 0
            End If
        
        End If
    Next site

    Call test(Proces)

End Function

'���݂̃E�F�[�nID���擾����

Private Sub Get_waferID()
    
    Dim typeFolder As String
    
    '### Production IF�V�[�g�ǂݍ��� ###
    Dim wkshtObj As Object
    Set wkshtObj = ThisWorkbook.Sheets("Production IF")
    '======= WorkSheet ErrorProcess ========
    If wkshtObj Is Nothing Then
        MsgBox "Not Find Sheet : " & " Production IF"
        Exit Sub
    End If

    '### Production IF�V�[�g����WaferID���擾 ###
'    NowWaferID = wkshtObj.Cells(WaferNo + 2, 10)
    NowWaferID = "ESD105706-08"         'DEBUG!!!!!!!
    
    typeFolder = Mid(NowWaferID, 3, 4)  'ex)M105
    
    NowWaferID = typeFolder + "\" + NowWaferID      'ex)M105\29M105001-01
    
End Sub

Private Function Open_File() As Boolean

    Dim FileNo As Integer                   '�t�@�C���i���o�[
    Dim strText As String                   '�ǂݍ��񂾓��e���i�[���܂��B
    Dim i, j As Integer
    
    Dim fileData, fileData2 As Variant      '�t�@�C������ǂݍ���NG�`�b�v�f�[�^�i�[�p
    
    Call Get_waferID    '�E�F�[�nID�擾
    
    #If EEE_AUTO_JOB_LOCATE = 1 Or EEE_AUTO_JOB_LOCATE = 2 Then
        LastProcessInfoFILE = LastProcessInfo_FilePATH_N & NowWaferID & ".txt"
        LastProcessInfoFILE_END = LastProcessInfo_FilePATH_N & NowWaferID & ".txt.END"
    #ElseIf EEE_AUTO_JOB_LOCATE = 3 Then
        LastProcessInfoFILE = LastProcessInfo_FilePATH_K & NowWaferID & ".txt"
        LastProcessInfoFILE_END = LastProcessInfo_FilePATH_K & NowWaferID & ".txt.END"
    #End If

    '### �Ώۂ̃t�@�C����������Δ����� ###
    flg_NoFILE = False
    If Dir(LastProcessInfoFILE) = "" Then
        Open_File = False
        flg_NoFILE = True
        Exit Function
    End If

    '### �Ώۂ̃G���h�t�@�C����������Δ����� ###
    flg_NoEndFILE = False
    If Dir(LastProcessInfoFILE_END) = "" Then
        Open_File = False
        flg_NoEndFILE = True
        Exit Function
    End If

    '### �t�@�C�����J�� ###
    FileNo = FreeFile
    Open LastProcessInfoFILE For Input As #FileNo
    On Error GoTo CloseFile

    '### �t�@�C������NG�`�b�v���^�f�[�^�����擾���� ###
    NGchipCNT = 0
    NGdataCNT = 0
    Do Until EOF(FileNo)
        Line Input #FileNo, strText
        fileData = Split(strText, ":")
        fileData2 = Split(fileData(1), ",")
        '=== 2�s�ڂ���f�[�^�����擾���� ===
        If NGchipCNT = 1 Then
            For j = 0 To UBound(fileData2)
                NGdataCNT = NGdataCNT + 1               '�f�[�^��
            Next j
        End If
        NGchipCNT = NGchipCNT + 1
    Loop
    NGchipCNT = NGchipCNT - 2                           'NG�`�b�v�̌�

    '### �t�@�C������� ###
    Close #FileNo

    Open_File = True


FILE_end:

Exit Function

CloseFile:

    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    Close #FileNo
    flg_NoEndFILE = True
'    MsgBox ("File Open Error! Please Check a File")
    GoTo FILE_end

End Function

'�e�L�X�g�t�@�C������f�[�^�ǂݍ��݁A�z��ϐ��֏�������

Public Function Init_LastProcessInfoFILE() As Boolean

    Dim FileNo As Integer                   '�t�@�C���i���o�[
    Dim strText As String                   '�ǂݍ��񂾓��e���i�[���܂��B
    Dim i, j As Integer
    
    Dim fileData, fileData2 As Variant      '�t�@�C������ǂݍ���NG�`�b�v�f�[�^�i�[�p
    
    '### File Search&Open ###
    If Open_File = False Then
        Exit Function
    End If
    '#################
    
    '--- �ϐ��錾 ---
    ReDim NGChipNo(NGchipCNT)                   'NG�`�b�vNo.�i�[�p
    ReDim NGChipData(NGchipCNT, NGdataCNT)      'NG�`�b�v�f�[�^�i�[�p
    ReDim FileStatus(NGdataCNT)                 '�t�@�C����Ԋi�[�p exist/not-exist
    '----------------

    '### �t�@�C�����J�� ###
    FileNo = FreeFile
    Open LastProcessInfoFILE For Input As #FileNo
    On Error GoTo CloseFile

    '### �t�@�C������NG�`�b�vNo.�ƃf�[�^���擾���� ###
    NGchipCNT = 0
    Do Until EOF(FileNo)
        Line Input #FileNo, strText
        fileData = Split(strText, ":")
        fileData2 = Split(fileData(1), ",")
        '=== 2�s�ڂ̃t�@�C����ԓǂݍ��� ===
        If NGchipCNT = 1 Then
            If fileData(0) <> "File" Then
                GoTo CloseFile
            End If
            For j = 0 To UBound(fileData2)
                If fileData2(j) = "exist" Or fileData2(j) = "not-exist" Or fileData2(j) = "" Then
                    FileStatus(j) = fileData2(j)                    '�t�@�C����Ԏ擾
                Else
                    GoTo CloseFile
                End If
            Next j
        End If
        '=== 3�s�ڂ���̃f�[�^�ǂݍ��� ===
        If NGchipCNT > 1 Then
            If CInt(fileData(0)) > 0 Then
                NGChipNo(NGchipCNT - 2) = fileData(0)                   'NG�`�b�vNo.
            Else
                GoTo CloseFile
            End If
            For j = 0 To UBound(fileData2)
                If fileData2(j) = "0" Or fileData2(j) = "1" Or fileData2(j) = "-1" Or fileData2(j) = "" Then
                    NGChipData(NGchipCNT - 2, j) = fileData2(j)     'NG�`�b�v�f�[�^
                Else
                    GoTo CloseFile
                End If
            Next j
        End If
        NGchipCNT = NGchipCNT + 1
    Loop
    NGchipCNT = NGchipCNT - 2
    
    '### �t�@�C������� ###
    Close #FileNo

FILE_end:

Exit Function

CloseFile:

    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    Close #FileNo
    flg_NoEndFILE = True
'    MsgBox ("File Open Error! Please Check a File")
    GoTo FILE_end
    
End Function

Public Function Get_NgChipInfo(failNUM As Integer, site As Long) As Integer

    Dim i, j As Integer

    Dim flg_Not_Exist As Boolean

    If Flg_AutoMode = True Then
    
        '### �Ώۂ̃G���h�t�@�C�������������ꍇ ###
        If flg_NoFILE = True Then
            Exit Function
        End If
        
        '### �Ώۂ̃G���h�t�@�C�������������ꍇ ###
        If flg_NoEndFILE = True Then
            Get_NgChipInfo = -1
            Exit Function
        End If

        '### �t�@�C����Ԃ̊m�F ###
        If FileStatus(failNUM - 1) = "not-exist" Then
            Get_NgChipInfo = -1
            Exit Function
        End If

        '### �e�H���̏����e�X�g���ʂƂ��ĕԂ� ###
        For i = 0 To NGchipCNT - 1
            If CInt(DeviceNumber_site(site)) = CInt(NGChipNo(i)) Then
                If NGChipData(i, failNUM - 1) <> "" Then
                    Get_NgChipInfo = CInt(NGChipData(i, failNUM - 1))     'NG�`�b�v�f�[�^�擾
                End If
                Exit For
            End If
        Next i
        
    End If
    
End Function

