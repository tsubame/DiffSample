Attribute VB_Name = "RegisterVersionManagementMod"
Option Explicit
'   ///Version 1.1///
'
'   Update history
'Ver1.1 2013/10/9 H.Arikawa HashCode�G���[����StopPMC�Ŏ~�߂�悤�ɏ����ǉ��B

'================================================================================
' For Hash Code Definition
'================================================================================
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
                            (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
                             ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
                            (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" _
                            (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, _
                             ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
                            (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" _
                            (ByVal hHash As Long, pbData As Any, ByVal cbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
                            (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, ByRef pcbData As Long, _
                             ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL   As Long = 1
Private Const PROV_RSA_AES    As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Private Const HP_HASHVAL      As Long = 2
Private Const HP_HASHSIZE     As Long = 4

Private Const ALG_TYPE_ANY    As Long = 0
Private Const ALG_CLASS_HASH  As Long = 32768

Private Const ALG_SID_MD2     As Long = 1
Private Const ALG_SID_MD4     As Long = 2
Private Const ALG_SID_MD5     As Long = 3
Private Const ALG_SID_SHA     As Long = 4
Private Const ALG_SID_SHA_256 As Long = 12
Private Const ALG_SID_SHA_512 As Long = 14

Private Const CALG_MD2        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
Private Const CALG_MD4        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
Private Const CALG_MD5        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
Private Const CALG_SHA        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
Private Const CALG_SHA_256    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)
Private Const CALG_SHA_512    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512)

'================================================================================
' Module Definition
'================================================================================
Private Const RVVM_LARGE_FILE_SIZE_HASHCODE As String = "This File Size is too large."
Private Const RVVM_FILE_NOT_FOUND_HASHCODE As String = "This File is not found."
Private Const RVVM_DONT_MANAGEMENT_VERSION As String = "This File is not out of version management!"

Private Type HashCode_information
    PatName As String
    filePath As String
    HashCode As String
End Type

Private Enum RegisterVersionManagementModState
    
    RVMM_STATE_UNINITIALIZED = 0
    RVMM_STATE_INITIALIZED = 1
    RVMM_STATE_HASE_CREATED = 2

End Enum

'================================================================================
' Module Variables
'================================================================================
Private HashCode_Data() As HashCode_information '���̃n�b�V���R�[�h�̏��
Private m_RvmmState As RegisterVersionManagementModState
Public Flg_HashCheckResult As Boolean           '�n�b�V���R�[�h�`�F�b�N���ʃt���O

Public Function myState() As Long
    myState = m_RvmmState
End Function

'================================================================================
' Public Functions
'================================================================================
Public Sub RVMM_Initialize()

    Erase HashCode_Data
    
    '��ԑJ��
    m_RvmmState = RVMM_STATE_INITIALIZED
    
End Sub


Public Sub RVMM_CreateRegisterHashCode()
        
        
'    HASHCODE�V�[�g���Ȃ������瓮���Ȃ�
    If Not IsHashCodeFunctionEnable() Then
        Call MsgBox("HashCode Sheet is not found! This function is disable!", vbCritical, "RVMM_CreateRegisterHashCode")
        Exit Sub
    End If
    
    '�ϐ��錾
    Dim VerX As Long
    Dim HashCode_Data_BeforeVersion() As HashCode_information '��O��Version�̃n�b�V���R�[�h�̏��
        
        
'    �������`�F�b�N
    If m_RvmmState = RVMM_STATE_UNINITIALIZED Then
        Call MsgBox("Register Version Management Mod is not be initialized!", vbCritical, "RVMM_CreateRegisterHashCode")
        Exit Sub
    End If
    
    'LoadPat���������̂Ɋւ��Ă͍Ō�ɏ��������Ȃ��̂ŕۑ����Ă���
    Dim ArrayMax As Long
    ArrayMax = GetUBoundHashCode_Data(HashCode_Data)

'    LoadPat���Ă΂�Ă��邱�Ƃ��O��
'    PatGrp�V�[�g��TestInstance��������W������
    Call GetHashCode_information
    

'    �p�X��S�����ׂ�Version���S���ꏏ�łȂ��ƃG���[�Ƃ���
    If Not IsSamePattenVersion(HashCode_Data, VerX) Then
        Call MsgBox("The Register Versions is not same!", vbCritical, "RVMM_CreateRegisterHashCode")
        GoTo ErrorExit
    End If
    
    
'    HashCode_Data����Ƃ���
'    REGVER�t�H���_��S���݂Ĥ�ŐV��REGVER�t�H���_�łȂ���΃G���[�Ƃ���
    If Not IsLatestRegversion(VerX) Then
        Call MsgBox("Pat files are not the latest!", vbCritical, "RVMM_CreateRegisterHashCode")
        GoTo ErrorExit
    End If
    
'    ����Version��HASHCODE���쐬����
    Call CreateHashCode_impl(HashCode_Data)
    
'  �t�@�C�������݂��Ȃ��ꍇ�͂����ň����|����
    Dim strNotFoundFile As String
    If Not IsAllFileHashCodeCreated(HashCode_Data, strNotFoundFile) Then
        Call MsgBox("Pat file " & strNotFoundFile & " is not found!", vbCritical, "RVMM_CreateRegisterHashCode")
        GoTo ErrorExit
    End If
    
'    ����Version��1�łȂ��Ȃ�A�O�̃o�[�W�����̃p�X���쐬���A�n�b�V���R�[�h���쐬����
    If VerX <> 1 Then
        Call ConvertBeforeVersionPath(HashCode_Data, HashCode_Data_BeforeVersion)
        Call CreateHashCode_impl(HashCode_Data_BeforeVersion)
    End If

    
'    ����Version��CreateHash�ōX�V����Ȃ��ꍇ�̓G���[���o���
    If Not IsUpdateRegisterVersion() Then
        Call MsgBox("All HashCode is same. Please check if you move pat files to ""RegVerX"" folder!", vbCritical, "RVMM_CreateRegisterHashCode")
        GoTo ErrorExit
    End If
        
'    �V�[�g�ɋL�ڂ�����O�Ƀ\�[�g���s���BVer��1�̏ꍇ�̓\�[�g���Ȃ�
    If VerX <> 1 Then
        '����Version����ɑO��Version�̃p�X����ёւ�
         Call SortHashCodeInformation(HashCode_Data, HashCode_Data_BeforeVersion)
    End If
      
'    ���̃o�[�W������HASHCODE�Ƥ�O�̃o�[�W������HASHCODE��HASHCODE�V�[�g�ɏ���
'    �V�[�g�̃N���A���s��
    Call WrtieHashCode(VerX <> 1, HashCode_Data, HashCode_Data_BeforeVersion)
    
'    ����Version�ƑO��Version��HASH�R�[�h�ɈႢ���������炵�邵������
    Call CheckHashCoceWithBeforeVersion
      
'    CreatHash����RegVer���O���t�@�C���ɗ������̂���
    Call WriteCreateHashCodeRecord(VerX)
        
'    ��n��
    Erase HashCode_Data_BeforeVersion
    Call RecoverHashCodeData(HashCode_Data, ArrayMax)

    
    '��ԑJ��
    m_RvmmState = RVMM_STATE_HASE_CREATED
    
    Call MsgBox("RVMM_CreateRegisterHashCode is succeeded !", , "Congratulation")
    
    Exit Sub
    
ErrorExit:
'    ��n��
    Erase HashCode_Data_BeforeVersion
    Call RecoverHashCodeData(HashCode_Data, ArrayMax)
         
End Sub

Public Function RVMM_GetRegisterVersion() As Long
    
    Dim VerX As Long
    
'    HASHCODE�V�[�g���Ȃ������瓮���Ȃ�
    If Not IsHashCodeFunctionEnable() Then
        Call OutpuMessageToIgxlDataLog("HashCode Sheet is not found! This function is disable!")
        RVMM_GetRegisterVersion = 0
        Flg_HashCheckResult = False
        Exit Function
    End If
    
'    �������`�F�b�N
    If m_RvmmState = RVMM_STATE_UNINITIALIZED Then
        Call OutpuMessageToIgxlDataLog("Register Version Management Mod is not be initialized!")
        RVMM_GetRegisterVersion = -1
        Flg_HashCheckResult = False
        Exit Function
    End If
    
'    LoadPat�Œǉ����ꂽ���͏��������Ȃ��̂Œl��ێ�����
    Dim ArrayMax As Long
    ArrayMax = GetUBoundHashCode_Data(HashCode_Data)
    
'    LoadPat���Ă΂�Ă��邱�Ƃ��O��
'    PatGrp�V�[�g��TestInstance��������W������
    Call GetHashCode_information

'    �p�X��S�����ׂ�Version���S���ꏏ�łȂ��ƃG���[�Ƃ���
    If Not IsSamePattenVersion(HashCode_Data, VerX) Then
        Call OutpuMessageToIgxlDataLog("The Register Versions is not same!")
        RVMM_GetRegisterVersion = -2
        Call DisableAllTest '�e�X�g�̒�~(EeeJob�֐�)
        GoTo ErrorExit
    End If
    
'    �O���t�@�C���̗������炷�ׂĂ�RegVer�t�H���_���S��HASHCODE�ϊ����ꂽ���`�F�b�N����
    If Not IsAllRegVerHashCreated Then
        Call OutpuMessageToIgxlDataLog("All ""RegVerX"" folder is not created hashcode!")
        RVMM_GetRegisterVersion = -3
        Call DisableAllTest '�e�X�g�̒�~(EeeJob�֐�)
        GoTo ErrorExit
    End If
    
'    �p�X�ƃp�^��������n�b�V���R�[�h�𐶐����ĤHASHCODE�V�[�g��Hash�R�[�h�Ɣ�r������v���Ȃ��ƃG���[
'    ����Version��HASHCODE���쐬����
    Call CreateHashCode_impl(HashCode_Data)

'   �p�^���t�@�C�����Ȃ������ꍇ�̏���
    Dim strNotFoundFile As String
    If Not IsAllFileHashCodeCreated(HashCode_Data, strNotFoundFile) Then
        Call OutpuMessageToIgxlDataLog("Pat file " & strNotFoundFile & " is not found!")
        RVMM_GetRegisterVersion = -4
        Call DisableAllTest '�e�X�g�̒�~(EeeJob�֐�)
        GoTo ErrorExit
    End If
    
'    �V�[�g�Ɣ�r����
    If Not IsEqaulToHashCode(HashCode_Data, strNotFoundFile) Then
        If (Len(strNotFoundFile) = 0) Then
            Call OutpuMessageToIgxlDataLog("HashCode Sheet is empty!")
        Else
            Call OutpuMessageToIgxlDataLog("Pat file " & strNotFoundFile & "'s hashcode is mismatch!")
        End If
        RVMM_GetRegisterVersion = -5
        Call DisableAllTest '�e�X�g�̒�~(EeeJob�֐�)
        GoTo ErrorExit
    End If
  
    '�Ԃ�l�Z�b�g
    RVMM_GetRegisterVersion = VerX
    
'    ��n��
    Call RecoverHashCodeData(HashCode_Data, ArrayMax)
    
    '��ԑJ��
    m_RvmmState = RVMM_STATE_HASE_CREATED
    
    Exit Function
    
ErrorExit:
'    ��n��
    Call RecoverHashCodeData(HashCode_Data, ArrayMax)
    First_Exec = 0
    Flg_HashCheckResult = False
End Function

Public Sub RVMM_LoadPat(ByVal PatPath As String)

'    ���삳��
'
'    Pass����PatName������ǂݎ���Ĥ������\���̂Ƀp�^�[���̖��O�Ƃ��Ĉ���
'    �p�^�[���̃��[�h������ (HashCode_information)

'    ' �p�^�[����ǂ݂���(FULL�p�X���������A�t�@�C�������������E�E�E)
#If 1 Then
    Call TheHdw.Digital.Patterns.pat(PatPath).Load
#End If

'    HASHCODE�V�[�g���Ȃ������炱��ȏ�͓����Ȃ�
    If Not IsHashCodeFunctionEnable() Then
        Exit Sub
    End If

'    �������`�F�b�N
    If m_RvmmState <> RVMM_STATE_INITIALIZED Then
        Call MsgBox("Register Version Management Mod is not be initialized!", vbCritical, "RVMM_LoadPat")
        Exit Sub
    End If

    'Path����PatName��ǂݎ��ALoad�ł������Ƃ���"PatPass"�̓t�@�C���p�X���ƍl���Ă悢
    Dim i As Integer
    Dim j As Integer
    i = InStrRev(PatPath, "\") + 1
    j = InStrRev(UCase(PatPath), UCase(".pat"))
        
    Dim tempstr As String
    tempstr = Mid(PatPath, i, j - i)


    '�\���̂Ƀp�^�[���̃p�X�Ɩ��O���i�[(2��ڈȍ~)
On Error GoTo FIRST_CYCLE
    Dim elem_max As Integer
    elem_max = UBound(HashCode_Data) + 1
On Error GoTo 0
    ReDim Preserve HashCode_Data(elem_max) As HashCode_information
    
    HashCode_Data(elem_max).filePath = PatPath
    HashCode_Data(elem_max).PatName = tempstr
        
    Exit Sub
    
    '�\���̂Ƀp�^�[���̃p�X�Ɩ��O���i�[(����)
FIRST_CYCLE:
    ReDim HashCode_Data(0) As HashCode_information
    HashCode_Data(0).filePath = PatPath
    HashCode_Data(0).PatName = tempstr

End Sub

'================================================================================
' Private Functions
'================================================================================

Private Sub OutpuMessageToIgxlDataLog(ByRef strMsg As String)

#If 1 Then
        Call TheExec.Datalog.WriteComment(strMsg)
#Else
        Debug.Print strMsg
#End If

End Sub


Private Function GetUBoundHashCode_Data(ByRef HashCodeArray() As HashCode_information)

    On Error GoTo FirstArrayAlloc
    GetUBoundHashCode_Data = UBound(HashCodeArray)
    GoTo AllocEnd
FirstArrayAlloc:
    GetUBoundHashCode_Data = -1
AllocEnd:
     On Error GoTo 0

End Function

Private Sub RecoverHashCodeData(ByRef HashCodeArray() As HashCode_information, ByVal lRecoverSize As Long)

    If lRecoverSize = -1 Then
        Erase HashCodeArray
    Else
        ReDim Preserve HashCodeArray(lRecoverSize)
    End If

End Sub

Private Function IsAllFileHashCodeCreated(ByRef HashCodeArray() As HashCode_information, ByRef strNotFoundFile As String) As Boolean

    Dim lBegin As Long, lEnd As Long
    
    lBegin = LBound(HashCodeArray)
    lEnd = UBound(HashCodeArray)
    
    Dim i As Long
    
    For i = lBegin To lEnd
        If StrComp(HashCodeArray(i).HashCode, RVVM_FILE_NOT_FOUND_HASHCODE) = 0 Then
            strNotFoundFile = HashCodeArray(i).PatName
            IsAllFileHashCodeCreated = False
            Exit Function
        End If
    Next i

    IsAllFileHashCodeCreated = True

End Function

Private Function GetHashCode_information() As Double

    '��������������������������������������������������������������������������
    '���e�X�g�C���X�^���X�̖��O�ƃp�X������Ă���
    '��������������������������������������������������������������������������
    Dim ShtTestInstances As Worksheet
    Set ShtTestInstances = ThisWorkbook.Worksheets("Test Instances")

    Dim ArrMax As Long
    Dim j As Long
    Dim i As Long
    
    On Error GoTo FirstArrayAlloc
    ArrMax = UBound(HashCode_Data)
    GoTo AllocEnd
FirstArrayAlloc:
    ArrMax = -1
AllocEnd:
     On Error GoTo 0
   
    j = ArrMax + 1 'GetRegVer�֓����z��̏����l
    i = 5   '�e�X�g�C���X�^���X��Check�J�n�s���w��i�Œ�j
    
    Do Until Len(Trim(ShtTestInstances.Cells(i, 2))) = 0  'Group Name���󔒃Z���i Len(Trim(�󔒃Z��)) �j�܂ōs��ς���Check����
        If ShtTestInstances.Cells(i, 3) = "IG-XL Template" Then
            ReDim Preserve HashCode_Data(j)
            HashCode_Data(j).PatName = ShtTestInstances.Cells(i, 2)
            HashCode_Data(j).filePath = ShtTestInstances.Cells(i, 14)
            j = j + 1
        End If
        i = i + 1
    Loop

    '��������������������������������������������������������������������������
    '��PatGrps�̖��O�ƃp�X������Ă���
    '��������������������������������������������������������������������������

    Dim objsheet_PatGrps As Object
    Dim shtPatGrp As Worksheet
    
    For Each objsheet_PatGrps In Worksheets
        If objsheet_PatGrps.Name = "PatGrps" Then
            Set shtPatGrp = ThisWorkbook.Worksheets("PatGrps")
            i = 4   'PatGrps��Check�J�n�s���w��i�Œ�j
            
            Do Until Len(Trim(shtPatGrp.Cells(i, 2))) = 0  'Group Name���󔒃Z���i Len(Trim(�󔒃Z��)) �j�܂ōs��ς���Check����
                ReDim Preserve HashCode_Data(j)
                HashCode_Data(j).PatName = shtPatGrp.Cells(i, 2)
                HashCode_Data(j).filePath = shtPatGrp.Cells(i, 3)
                j = j + 1
                i = i + 1
            Loop
        End If
    Next

End Function

'HASHCODE�V�[�g���Ȃ������瓮���Ȃ�
'�c������
Private Function IsHashCodeFunctionEnable() As Boolean

    IsHashCodeFunctionEnable = False
    
    Dim shtHashCode As Worksheet
    
    On Error GoTo errLable

    '======= WorkSheet Select ========
    Set shtHashCode = ThisWorkbook.Sheets("HashCode")
    
    IsHashCodeFunctionEnable = True
    
    Set shtHashCode = Nothing
    
    Exit Function
    
errLable:
    IsHashCodeFunctionEnable = False
End Function

'REGVER�t�H���_��S���݂Ĥ�ŐV��REGVER�t�H���_�łȂ���΃G���[�Ƃ���
Private Function IsLatestRegversion(ByVal VersionX As Long) As Boolean
'����(�^)
Dim strRet As String
Dim strChar As String
Dim strOrg As String
Dim PatFolder() As String
Dim PatPath, FolderName
Dim GetFolderName As Integer
Dim FolderNoA As Integer
Dim FolderNoB As Integer
Dim LatestVer As Integer
Dim NameLength As Integer
Dim LatestFolderNo As Integer

    '�g�p����f�B���N�g���w��
    PatPath = ThisWorkbook.Path & "\PAT\"
    'Regver�Ƃ������O�̃t�H���_�̌���(�ŏ��̃t�H���_�̒l������B)
    FolderName = Dir(PatPath & "Regver*", vbDirectory)
    '�����ݒ�
    FolderNoA = 0
    LatestVer = 0
    ReDim PatFolder(0)
    
    '�f�B���N�g������Regver�Ƃ������O�̃t�H���_�[��S��PatFolder�Ɋi�[�B
    'Regver���t�H���_�ł��邱�Ƃ��m�F����B
    Do While FolderName <> ""
        '���݂̃t�H���_�Ɛe�t�H���_�͖����B
        If FolderName <> "." And FolderName <> ".." Then
            If (GetAttr(PatPath & FolderName) And vbDirectory) = vbDirectory Then
                ReDim Preserve PatFolder(FolderNoA)
                PatFolder(FolderNoA) = FolderName
                FolderNoA = FolderNoA + 1
            End If
        End If
        FolderName = Dir '���̃t�H���_����Ԃ�
    Loop
    
    For FolderNoB = 0 To FolderNoA - 1
        strRet = ""
        strOrg = PatFolder(FolderNoB)
        
        '�t�H���_������Version�ƂȂ鐔���𔲂��o���B
        For NameLength = 1 To Len(strOrg)
            strChar = Mid(strOrg, NameLength, 1)
            If IsNumeric(strChar) Then
                strRet = strRet & strChar
            End If
        Next NameLength
        
        If LatestVer < strRet Then
            LatestVer = strRet
        ElseIf LatestVer = strRet Then
            '����ver�͑��݂��Ȃ����߃G���[�Ƃ���B
            MsgBox "����version�����݂��Ă��܂��B"
            Exit Function
        End If
    Next FolderNoB
    
    If LatestVer = VersionX Then
        IsLatestRegversion = True
    End If
    
End Function
    
' �p�X��S�����ׂ�Version���S���ꏏ�łȂ��ƃG���[�Ƃ���
Private Function IsSamePattenVersion(ByRef HashCodeArray() As HashCode_information, ByRef VerX As Long) As Boolean
    '�ԍ�
    Const Offset_RegVer As Long = 6   'RegVerX��X��ǂވׂ�6(RegVer�̕�����)��Offset�Ƃ��ē���
    
    Dim i As Long, lBegin As Long, lEnd As Long
    
    '�z��̑傫���m�F
    On Error GoTo NOT_ARRAY_ALLOC
    lBegin = LBound(HashCodeArray)
    lEnd = UBound(HashCodeArray)
    On Error GoTo 0

    '��������������������������������������������������������������������������
    '��Path����RegVerX��X�̒l���r���ē����Ȃ�Ver��Ԃ��B����Ă�����-1��Ԃ��B
    '��������������������������������������������������������������������������
    Dim RegVer_Posision As Long
    Dim counter As Long
    Dim lArraySize As Long
    lArraySize = 0
    For counter = lBegin To lEnd       '�z�������r����B
        Dim VerNumber As Long
        Dim ArrRegVerX() As String
        
        VerNumber = 0   'X�̕�����(�����l0)

        'Path����RegVer�̈ʒu���擾���܂��B
        RegVer_Posision = InStr(UCase(HashCodeArray(counter).filePath), UCase("RegVer"))

        If RegVer_Posision <> 0 Then
            'RegVerX��X�����������擾���܂�(a�����߂�)
            Do While IsNumeric(Mid(HashCodeArray(counter).filePath, RegVer_Posision + Offset_RegVer + VerNumber, 1))
                VerNumber = VerNumber + 1   '�����؂����������������������玟�̕��������ɍs���B
            Loop
        
            'X�̒l�����߂܂��B
            ReDim Preserve ArrRegVerX(lArraySize)
            ArrRegVerX(lArraySize) = Mid(HashCodeArray(counter).filePath, RegVer_Posision + Offset_RegVer, VerNumber)
            
            'X�̒l���r���܂��B
            If lArraySize <> 0 Then
                If ArrRegVerX(lArraySize) <> ArrRegVerX(lArraySize - 1) Then
                    IsSamePattenVersion = False
                    VerX = -1
                    Exit Function
                End If
                VerX = ArrRegVerX(lArraySize)
            End If
            lArraySize = lArraySize + 1
        End If
    Next

    'RegVerX��X�S�Ă��������̂�IsSamePattenVersion��True��Ԃ��܂��B
    IsSamePattenVersion = True
    
     Exit Function
NOT_ARRAY_ALLOC:
    Err.Raise 9999, "Test"
    
End Function
    
Private Sub ConvertBeforeVersionPath(ByRef NowArray() As HashCode_information, ByRef BeforeArray() As HashCode_information)
    '�ԍ�
    Const Offset_RegVer As Long = 6   'RegVerX��X��ǂވׂ�6(RegVer�̕�����)��Offset�Ƃ��ē���
    
    Dim i As Long, lBegin As Long, lEnd As Long
    
    '�z��̑傫���m�F
    On Error GoTo NOT_ARRAY_ALLOC
    lBegin = LBound(NowArray)
    lEnd = UBound(NowArray)
    On Error GoTo 0

    
    '�z��m��
    ReDim BeforeArray(lEnd)
    
    '��������������������������������������������������������������������������
    '������Version�p�X����O��Version�̃p�X�����o���B
    '��������������������������������������������������������������������������
    Dim RegVer_Posision As Long
    Dim counter As Long

    counter = 0     '������

    For counter = lBegin To lEnd       '�z�������r����B
        Dim VerNumber As Long
        Dim ArrRegVerX() As String
        Dim lX() As Long
        Dim BeforeX() As String
        Dim BeforelX() As Long
        Dim FilePath_No As Long
        Dim FirstPath As String
        Dim LatterPath As String
        Dim FirstFilePath_No As Long
        
       
        VerNumber = 0   'X�̕�����(�����l0)

        'Pass����RegVer�̈ʒu���擾���܂��B
        RegVer_Posision = InStr(UCase(NowArray(counter).filePath), UCase("RegVer"))

        If RegVer_Posision <> 0 Then
            'RegVerX��X�����������擾���܂�(a�����߂�)
            Do While IsNumeric(Mid(NowArray(counter).filePath, RegVer_Posision + Offset_RegVer + VerNumber, 1))
                VerNumber = VerNumber + 1   '�����؂����������������������玟�̕��������ɍs���B
            Loop
        
            'X�̒l�B�����߂܂��B
            ReDim Preserve ArrRegVerX(counter)
            ReDim Preserve lX(counter)
            ReDim Preserve BeforeX(counter)
            ReDim Preserve BeforelX(counter)
            ArrRegVerX(counter) = Mid(NowArray(counter).filePath, RegVer_Posision + Offset_RegVer, VerNumber)
            lX(counter) = val(ArrRegVerX(counter))   '������𐔒l�ɕϊ�
            
            'X�̒l��-1���܂��B�iVer��������܂��B�j
            BeforelX(counter) = lX(counter) - 1
            BeforeX(counter) = str(BeforelX(counter))
            
            'Path����������ׂɑS�̂̕������𐔂��܂��B
            FilePath_No = Len(NowArray(counter).filePath)
            
            'RegVer�܂ł�Path�ƕ��������擾���܂��B
            FirstPath = Left(NowArray(counter).filePath, RegVer_Posision + Offset_RegVer - 1)
            FirstFilePath_No = Len(FirstPath)
            
            'RegVerX�ȍ~��Path���擾���܂��B
            LatterPath = Right(NowArray(counter).filePath, FilePath_No - FirstFilePath_No - VerNumber)
            
            '�O��Version�̃p�X�����o���܂��B
            BeforeArray(counter).filePath = FirstPath & LTrim(BeforeX(counter)) & LatterPath
            BeforeArray(counter).PatName = NowArray(counter).PatName
        Else
            BeforeArray(counter).filePath = RVVM_DONT_MANAGEMENT_VERSION
            BeforeArray(counter).PatName = NowArray(counter).PatName
        End If
    Next
    
     Exit Sub
NOT_ARRAY_ALLOC:
    Err.Raise 9999, "Test"
End Sub

'����Version��HASHCODE���쐬����
'�O��Version��HASHCODE���쐬����
Private Sub CreateHashCode_impl(ByRef HashCodeArray() As HashCode_information)
    '�ێR
    Const MAX_SIZE As Long = 10# * 1024 * 1024
   
    Dim i As Long, lBegin As Long, lEnd As Long
    
    '�z��̑傫���m�F
    On Error GoTo NOT_ARRAY_ALLOC
    lBegin = LBound(HashCodeArray)
    lEnd = UBound(HashCodeArray)
    On Error GoTo 0
    
    For i = lBegin To lEnd
        
        '�t�@�C���̃T�C�Y���擾
        Dim File_Size As Long
    
        If StrComp(HashCodeArray(i).filePath, RVVM_DONT_MANAGEMENT_VERSION) = 0 Then '�o�[�W�����Ǘ����Ȃ��t�@�C���̏ꍇ
            File_Size = -2 '�o�[�W�����Ǘ����Ȃ��Ƃ���-1�Ƃ���
        ElseIf Dir(HashCodeArray(i).filePath) <> "" Then '�t�@�C���̑��݊m�F
            File_Size = FileLen(HashCodeArray(i).filePath)
        Else
            File_Size = -1  '�t�@�C���̑��݂��Ȃ��Ƃ���-1�Ƃ���
        End If

        If (File_Size = 0) Then
            HashCodeArray(i).HashCode = ""
        ElseIf (File_Size = -1) Then
            HashCodeArray(i).HashCode = RVVM_FILE_NOT_FOUND_HASHCODE
        ElseIf (File_Size = -2) Then
            HashCodeArray(i).HashCode = RVVM_DONT_MANAGEMENT_VERSION
        ElseIf (File_Size > MAX_SIZE) Then
            HashCodeArray(i).HashCode = RVVM_LARGE_FILE_SIZE_HASHCODE
        Else
            HashCodeArray(i).HashCode = CreateHashFile(HashCodeArray(i).filePath, CALG_MD5)
        End If
'        HashCodeArray(i).HashCode = CreateMD5HashString(HashCodeArray(i).FilePath)
    Next i
    
    Exit Sub
    
NOT_ARRAY_ALLOC:
    Err.Raise 9999, "CreateHashCode_impl", "Memory Allocation Error!"
End Sub
    
'����Version��CreateHash�ōX�V����Ȃ��ꍇ�̓G���[���o���
Private Function IsUpdateRegisterVersion() As Boolean
    '�ێR
    
    IsUpdateRegisterVersion = False
    
    '�V�[�g�̎擾
    Dim shtHashCode As Worksheet
    Set shtHashCode = ThisWorkbook.Worksheets("HashCode")
    
    'HashCode_Data�z��̑傫���m�F
    Dim lBegin As Long, lEnd As Long
    On Error GoTo NOT_ARRAY_ALLOC
    lBegin = LBound(HashCode_Data)
    lEnd = UBound(HashCode_Data)
    On Error GoTo 0
    
    
    '===����j�Ƃ��Ă̓f�o�b�O���ǉ��̂��Ƃ��l���āAHashCode_Data����ɃV�[�g���m�F���Ă���====
    '�܂��V�[�g������̎擾�A����ۂ̏ꍇ��VersionUp���Ȃ����Ƃ��Ă����ʂ���
    Dim i As Long, j As Long
    Dim shtDataArray() As HashCode_information
    
    i = 4
    j = 0
    If Len(Trim(shtHashCode.Cells(i, 2))) = 0 Then
        Set shtHashCode = Nothing '�V�[�g�̊J��
        IsUpdateRegisterVersion = True
        Exit Function
    End If
    
    Do Until Len(Trim(shtHashCode.Cells(i, 2))) = 0  'Group Name���󔒃Z���i Len(Trim(�󔒃Z��)) �j�܂ōs��ς���Check����
        ReDim Preserve shtDataArray(j)
        With shtDataArray(j)
            .PatName = shtHashCode.Cells(i, 2)
            .HashCode = shtHashCode.Cells(i, 3)
        End With
        j = j + 1
        i = i + 1
    Loop
    
    Set shtHashCode = Nothing '�V�[�g�̊J��
    
    'shtDataArray�z��̑傫���m�F
    Dim lBeginSht As Long, lEndSht As Long
    On Error GoTo NOT_ARRAY_ALLOC
    lBeginSht = LBound(shtDataArray)
    lEndSht = UBound(shtDataArray)
    On Error GoTo 0
    
    
    '����HashCode_Data����ɃV�[�g�̏����m�F���Ă���
    Dim IsMatch As Boolean
    j = 0
    For i = lBegin To lEnd
        IsMatch = False '��������False�ɂ���
        For j = lBeginSht To lEndSht
            If (StrComp(HashCode_Data(i).PatName, shtDataArray(j).PatName) = 0) Then
                If (StrComp(HashCode_Data(i).HashCode, RVVM_LARGE_FILE_SIZE_HASHCODE) = 0) And _
                    (StrComp(shtDataArray(i).HashCode, RVVM_LARGE_FILE_SIZE_HASHCODE) = 0) Then
                    IsMatch = False 'PatName����v���ăt�@�C���T�C�Y���傫���ꍇFalse�Ƃ���
                    Exit For
                ElseIf (StrComp(HashCode_Data(i).HashCode, shtDataArray(j).HashCode) = 0) Then
                    IsMatch = True 'PatName�ƃn�b�V���R�[�h����v����ꍇTrue�Ƃ���
                    Exit For
                Else
                    IsMatch = False 'PatName����v���ăn�b�V���R�[�h���قȂ�ꍇFalse�Ƃ���
                    Exit For
                End If
            End If
            
        Next j
        
        '������IsMathch=False�̏ꍇ��
        '�EHashCode_Data�̏�񂪃V�[�g�ɂ݂���Ȃ�����
        '�EHashCode_Data�̏�񂪃V�[�g�ƈقȂ��Ă���
        '���Ƃ��Ӗ�����̂Ńo�[�W�����A�b�v���Ȃ��ꂽ�Ƃ��Ĕ����Ă悢
        If Not IsMatch Then
            Erase shtDataArray
            IsUpdateRegisterVersion = True
            Exit Function
        End If
    Next i
    
    IsUpdateRegisterVersion = False
    
    Exit Function
    
NOT_ARRAY_ALLOC:
    Err.Raise 9999, "IsUpdateRegisterVersion", "Memory Allocation Error!"
    
End Function

        
'���̃o�[�W������HASHCODE�Ƥ�O�̃o�[�W������HASHCODE��HASHCODE�V�[�g�ɏ���
'���Ń\�[�g���邱��
Private Sub WrtieHashCode(ByVal IsWriteBeforeVersion As Boolean, _
        ByRef NowArray() As HashCode_information, ByRef BeforeArray() As HashCode_information)
    
    '�ێR
    '�G���[�����͂ǂ��܂ł��悤���H�ЂƂ܂��������Ȃ��ł���
    
    Const COLUMN_PAT_NAME As Long = 2
    Const COLUMN_PAT_NOW_HASH As Long = 3
    Const COLUMN_PAT_BEFORE_HASH As Long = 4
    
    Const ROW_START As Long = 4
    
    Dim shtHashCode As Worksheet
    Set shtHashCode = ThisWorkbook.Worksheets("HashCode")
    
    '��������̂̃N���A
    Call ClearWorkShet(shtHashCode)
        
    '�v�f���̎擾�A���ёւ���Ȃ̂ŁA����������Ă���͂�
    Dim i As Long, lBegin As Long, lEnd As Long
    lBegin = LBound(NowArray)
    lEnd = UBound(NowArray)
    
    '���C�g����
    If IsWriteBeforeVersion Then
        For i = lBegin To lEnd
            With shtHashCode
                .Cells(i + ROW_START, COLUMN_PAT_NAME) = NowArray(i).PatName
                .Cells(i + ROW_START, COLUMN_PAT_NOW_HASH) = NowArray(i).HashCode
                .Cells(i + ROW_START, COLUMN_PAT_BEFORE_HASH) = BeforeArray(i).HashCode
            End With
        Next i
    Else
        For i = lBegin To lEnd
            With shtHashCode
                .Cells(i + ROW_START, COLUMN_PAT_NAME) = NowArray(i).PatName
                .Cells(i + ROW_START, COLUMN_PAT_NOW_HASH) = NowArray(i).HashCode
            End With
        Next i
    End If
    
    Set shtHashCode = Nothing '�V�[�g�̊J��
    
End Sub
    
'����Version�ƑO��Version��HASH�R�[�h�ɈႢ���������炵�邵������
Private Sub CheckHashCoceWithBeforeVersion()

    '�c������
    Const COLUMN_PAT_NAME As Long = 2
    Const COLUMN_PAT_NOW_HASH As Long = 3
    Const COLUMN_PAT_BEFORE_HASH As Long = 4
    Const COLUMN_PAT_DIFF As Long = 5
    Const ROW_START As Long = 4
    
    Dim shtHashCode As Worksheet
    Set shtHashCode = ThisWorkbook.Worksheets("HashCode")
    
    Dim i As Long
    Dim strNowHash As String
    Dim strBeforeHash As String
    
    Do Until Len(Trim(shtHashCode.Cells(i + ROW_START, COLUMN_PAT_NAME))) = 0 'Group Name���󔒃Z���i Len(Trim(�󔒃Z��)) �j�܂ōs��ς���Check����
        strNowHash = shtHashCode.Cells(i + ROW_START, COLUMN_PAT_NOW_HASH)
        strBeforeHash = shtHashCode.Cells(i + ROW_START, COLUMN_PAT_BEFORE_HASH)
        
        If StrComp(RVVM_LARGE_FILE_SIZE_HASHCODE, strNowHash) = 0 Or StrComp(RVVM_LARGE_FILE_SIZE_HASHCODE, strBeforeHash) = 0 Then
            shtHashCode.Cells(i + ROW_START, COLUMN_PAT_DIFF) = "*"
        ElseIf StrComp(RVVM_DONT_MANAGEMENT_VERSION, strBeforeHash) = 0 Then
            shtHashCode.Cells(i + ROW_START, COLUMN_PAT_DIFF) = "+"
        ElseIf StrComp(strNowHash, strBeforeHash) <> 0 Then
            shtHashCode.Cells(i + ROW_START, COLUMN_PAT_DIFF) = "O"
        End If
        i = i + 1
    Loop

End Sub
  
Private Sub WriteCreateHashCodeRecord(ByRef VerX As Long)
    '�ԍ�
    '��������������������������������������������������������������������������
    '��CreatHash����RegVer���O���t�@�C���ɗ������̂����B
    '��������������������������������������������������������������������������

    Dim BookName As String
    Dim TypeName As String
    Dim DerivationName As String
    Dim RecordFullPathName As String
    Dim intFileNum As Integer
    Dim RecordDate As Date

    Const StartTypeNamePosition As String = 4   'Job����Type�������Ă���J�n�ʒu(Ex.p7e104lq4�Ȃ�4�����ڂ���)
    Const TypeNameNumber As String = 3          'Type�̕������i083,104���j
    Const StartDerivationNamePosition As String = 7   'Job���̔h���������Ă���J�n�ʒu(Ex.p7e104lq4�Ȃ�7�����ڂ���)
    Const DerivationNameNumber As String = 2          '�h���̕������ilq,cq,aq���j
    

    'Type���擾(Excel��JobName����擾)�@p7e104lq4
    BookName = ThisWorkbook.Name
    TypeName = Mid(BookName, StartTypeNamePosition, TypeNameNumber)
    
    If Not IsNumeric(TypeName) Then
        Call MsgBox("Pleas Check Job TypeName (ExcelFileName)")     '�����؂����������������Ŗ���������G���[���b�Z�[�W��\���B
        Err.Raise 9999, "Test"
    End If
    
    '�h�����擾
    DerivationName = Mid(BookName, StartDerivationNamePosition, DerivationNameNumber)

    '���t���擾
    RecordDate = Date

    '���������邩�ǂ����m�F(�܂��AJob��FilePath���擾)
    RecordFullPathName = ThisWorkbook.Path & "\PAT\" & "HashCodeRecorde" & "_" & TypeName & "_" & DerivationName & ".txt"

    '�����̍쐬
    If Dir(RecordFullPathName) = "" Then   '������ΐV�K�쐬
        intFileNum = FreeFile
        Open RecordFullPathName For Output As intFileNum
        Print #intFileNum, RecordDate & " " & "RegVer" & LTrim(str(VerX))
        Close #intFileNum
    Else                                    '���ɗ���������Ώ㏑��
        intFileNum = FreeFile
        Open RecordFullPathName For Append As intFileNum
        Print #intFileNum, RecordDate & " " & "RegVer" & LTrim(str(VerX))
        Close #intFileNum
    End If
    
End Sub
    
Private Function IsAllRegVerHashCreated() As Boolean
    '�ԍ�
Dim PatPath As String
Dim FolderName As String
Dim FolderNo As Integer

    '��������������������������������������������������������������������������
    '���O���t�@�C���̗������炷�ׂĂ�RegVer�t�H���_���S��HASHCODE�ϊ����ꂽ���`�F�b�N����
    '��������������������������������������������������������������������������


    '�����������t�H���_����RegVer���擾���遡��������
    '�g�p����f�B���N�g���w��
    PatPath = ThisWorkbook.Path & "\PAT\"
    'Regver�Ƃ������O�̃t�H���_�̌���(�ŏ��̃t�H���_)
    FolderName = Dir(PatPath & "Regver*", vbDirectory)
    '�����ݒ�
    FolderNo = 0
    ReDim ExistenceFolder(0)
    
    '  �f�B���N�g������Regver�Ƃ������O�̃t�H���_�[��S��ExistenceFolder�Ɋi�[�B
    'Regver���t�H���_�ł��邱�Ƃ��m�F����B
    Do While FolderName <> ""
        '���݂̃t�H���_�Ɛe�t�H���_�͖����B
        If FolderName <> "." And FolderName <> ".." Then
            If (GetAttr(PatPath & FolderName) And vbDirectory) = vbDirectory Then
                'ExistenceFolder�̃T�C�Y�ɍ��킹�Ĕz��̑傫����ς���B
                ReDim Preserve ExistenceFolder(FolderNo)
                ExistenceFolder(FolderNo) = FolderName
                FolderNo = FolderNo + 1
            End If
        End If
        FolderName = Dir '���̃t�H���_����Ԃ�
    Loop
    
    
    '�����������������L�邩���m�F���遡��������
    Dim BookName As String
    Dim TypeName As String
    Dim DerivationName As String
    Dim RecordFullPathName As String
    Dim intFileNum As Integer

    Const StartTypeNamePosition As String = 4   'Job����Type�������Ă���J�n�ʒu(Ex.p7e104lq4�Ȃ�4�����ڂ���)
    Const TypeNameNumber As String = 3          'Type�̕������i083,104���j
    Const StartDerivationNamePosition As String = 7   'Job���̔h���������Ă���J�n�ʒu(Ex.p7e104lq4�Ȃ�7�����ڂ���)
    Const DerivationNameNumber As String = 2          '�h���̕������ilq,cq,aq���j

    'Type���擾(Excel��JobName����擾)�@p7e104lq4
    BookName = ThisWorkbook.Name
    TypeName = Mid(BookName, StartTypeNamePosition, TypeNameNumber)
    
    If Not IsNumeric(TypeName) Then
        Call MsgBox("Pleas Check Job TypeName (ExcelFileName)")     '�����؂����������������Ŗ���������G���[���b�Z�[�W��\���B
        Err.Raise 9999, "Test"
    End If
    
    '�h�����擾
    DerivationName = Mid(BookName, StartDerivationNamePosition, DerivationNameNumber)

    '���������邩�ǂ����m�F(�܂��AJob��FilePath����擾)
    RecordFullPathName = ThisWorkbook.Path & "\PAT\" & "HashCodeRecorde" & "_" & TypeName & "_" & DerivationName & ".txt"
    
    
    If Dir(RecordFullPathName) <> "" Then   '����������ꍇ�̂ݔ�r����
        '����������txt����RegVer���擾���遡��������
        
        Const Offset_RegVer As Long = 6   'RegVerX��X��ǂވׂ�6(RegVer�̕�����)��Offset�Ƃ��ē���
    
        Dim LineDate As String
        Dim VerNumber As Long
        Dim RecodeDate() As String
        Dim lCounter As Long
        Dim RegVer_Posision As Long

        lCounter = 0 '������
        
        intFileNum = FreeFile
        
        Open RecordFullPathName For Input As intFileNum
        While Not EOF(intFileNum)
            Line Input #intFileNum, LineDate
                '�ǂݍ��񂾃f�[�^����RegVer�̈ʒu���擾���܂��B
                RegVer_Posision = InStr(1, LineDate, "RegVer", vbBinaryCompare)
        
                'RegVerX��X�����������擾���܂�(a�����߂�)
                VerNumber = 0   'X�̕�����(�����l0)
                Do While IsNumeric(Mid(LineDate, RegVer_Posision + Offset_RegVer + VerNumber, 1))
                    VerNumber = VerNumber + 1   '�����؂����������������������玟�̕��������ɍs���B
                Loop
            
                '�ǂݍ��񂾃f�[�^��RegVerX�����߂܂��B
                ReDim Preserve RecodeDate(lCounter)
                RecodeDate(lCounter) = Mid(LineDate, RegVer_Posision, VerNumber + Offset_RegVer)
                lCounter = lCounter + 1
        Wend
        
        Close intFileNum
        
        '�����������t�H���_�ƃe�L�X�g��RegVer���r���遡��������
        Dim Loop1 As Long
        Dim Loop2 As Long
        Dim IsFolderCheck As Boolean
        
        FolderNo = FolderNo - 1
        lCounter = lCounter - 1
        
        For Loop1 = 0 To FolderNo
            IsFolderCheck = False
            For Loop2 = 0 To lCounter
                If StrComp(UCase(ExistenceFolder(Loop1)), UCase(RecodeDate(Loop2))) = 0 Then
                    IsFolderCheck = True
                    Exit For
                End If
            Next
            If IsFolderCheck = False Then
                IsAllRegVerHashCreated = False
                Exit Function
            End If
        Next
    Else
        IsAllRegVerHashCreated = False
        Exit Function
    End If
    IsAllRegVerHashCreated = True
End Function

'�V�[�g�Ɣ�r����
Private Function IsEqaulToHashCode(ByRef HashCodeArray() As HashCode_information, ByRef strPatName As String) As Boolean
    '�c������
    
    'HashCode_Data�z��̑傫���m�F
    Dim lBegin As Long, lEnd As Long
    On Error GoTo NOT_ARRAY_ALLOC
    lBegin = LBound(HashCodeArray)
    lEnd = UBound(HashCodeArray)
    On Error GoTo 0
    
    '���[�N�V�[�g�̎擾
    Dim shtHashCode As Worksheet
    Set shtHashCode = ThisWorkbook.Worksheets("HashCode")
     
    
    '===�ǂݍ��܂ꂽ�p�^���t�@�C����HashCode�V�[�g�̃n�b�V���R�[�h�Ɠ��������Ƃ��m�F������====
    '===HashCode�V�[�g�ɗ]���Ȃ��̂������Ă��邱�ƂɊւ��Ă�OK�Ƃ���====
    '�܂��V�[�g������̎擾�A����ۂ̏ꍇ��VersionUp���Ȃ����Ƃ��Ă����ʂ���
    Dim i As Long, j As Long
    Dim shtDataArray() As HashCode_information
    i = 4
    j = 0
    If Len(Trim(shtHashCode.Cells(i, 2))) = 0 Then
        Set shtHashCode = Nothing '�V�[�g�̊J��
        IsEqaulToHashCode = False
        strPatName = ""
        Exit Function
    End If
    
    Do Until Len(Trim(shtHashCode.Cells(i, 2))) = 0  'Group Name���󔒃Z���i Len(Trim(�󔒃Z��)) �j�܂ōs��ς���Check����
        ReDim Preserve shtDataArray(j)
        With shtDataArray(j)
            .PatName = shtHashCode.Cells(i, 2)
            .HashCode = shtHashCode.Cells(i, 3)
        End With
        j = j + 1
        i = i + 1
    Loop
    
     
    'shtDataArray�z��̑傫���m�F
    Dim lBeginSht As Long, lEndSht As Long
    On Error GoTo NOT_ARRAY_ALLOC
    lBeginSht = LBound(shtDataArray)
    lEndSht = UBound(shtDataArray)
    On Error GoTo 0
     
     
     '����HashCode_Data����ɃV�[�g�̏����m�F���Ă���
    Dim IsMatch As Boolean
    j = 0
    For i = lBegin To lEnd
        IsMatch = False '��������False�ɂ���
        For j = lBeginSht To lEndSht
            If (StrComp(HashCodeArray(i).PatName, shtDataArray(j).PatName) = 0) Then
                If (StrComp(HashCodeArray(i).HashCode, shtDataArray(j).HashCode) = 0) Then
                    IsMatch = True 'PatName�ƃn�b�V���R�[�h����v����ꍇTrue�Ƃ���
                    Exit For
                Else
                    IsMatch = False 'PatName����v���Ăƃn�b�V���R�[�h���قȂ�ꍇFalse�Ƃ���
                    Exit For
                End If
            End If
            
        Next j
        
        '������IsMathch=False�̏ꍇ��
        '�EHashCode_Data�̏�񂪃V�[�g�ɂ݂���Ȃ�����
        '�EHashCode_Data�̏�񂪃V�[�g�ƈقȂ��Ă���
        '���Ƃ��Ӗ�����̂Ŕ����Ă悢
        If Not IsMatch Then
            Erase shtDataArray
            IsEqaulToHashCode = False
            strPatName = HashCodeArray(i).PatName
            Exit Function
        End If
    Next i
   
   '�����܂ł����犮�S��v���v���[���g
   IsEqaulToHashCode = True
   
   Exit Function
   
NOT_ARRAY_ALLOC:
    Err.Raise 9999, "IsEqaulToHashCode", "Memory Allocation Error!"
    
End Function

Private Sub SortHashCodeInformation(ByRef NowArray() As HashCode_information, ByRef BeforeArray() As HashCode_information)
'����(�^)
    Dim BuffHensu() As HashCode_information    '�l���X���b�v���邽�߂̍�ƈ�
  
    Dim lngBaseNumber As Long        '�����̗v�f�ԍ����i�[����ϐ�
    Dim iLoopNow  As Long            '���[�v�J�E���^(��ver��HushCode_data)
    Dim iLoopBefore  As Long         '���[�v�J�E���^(�Over��HushCode_data)
    Dim lngEnd As Long
    
    
    lngEnd = UBound(HashCode_Data)   '���[�v�J�E���^(��ver��HushCode_data)�̏I����
    ReDim BuffHensu(UBound(HashCode_Data))

    '��PatName�ƑOPatName�������ł���΁A�ꎞ�I��BuffHensu�Ɉړ��B
    For iLoopNow = 0 To lngEnd
        For iLoopBefore = 0 To lngEnd
            If NowArray(iLoopNow).PatName = BeforeArray(iLoopBefore).PatName Then
                BuffHensu(iLoopNow).filePath = BeforeArray(iLoopBefore).filePath
                BuffHensu(iLoopNow).HashCode = BeforeArray(iLoopBefore).HashCode
                BuffHensu(iLoopNow).PatName = BeforeArray(iLoopBefore).PatName
                Exit For
            End If
        Next iLoopBefore
    Next iLoopNow

    'BuffHensu�ɂ���f�[�^��Oversion�̃f�[�^�Ƃ��Ĉړ��B
    For iLoopNow = 0 To lngEnd
    BeforeArray(iLoopNow) = BuffHensu(iLoopNow)
    Next iLoopNow

End Sub

'================================================================================
' Function Level2
'================================================================================
Private Sub ClearWorkShet(ByRef sht As Worksheet)
'���e:
'   �O�o�͂�������N���A����
'
'[CndChk_wkst]    IN   Worksheet:    �������[�N�V�[�g
'
'���ӎ���:
'
    Const COLUMN_PAT_NAME As Long = 2
    Const ROW_START As Long = 4
    
    '�`�������
    Application.ScreenUpdating = False
        
    '�Ō�̃Z�����擾
    Dim rgLast As Range
    Set rgLast = sht.Cells.SpecialCells(xlCellTypeLastCell)

    '�Ώۗ̈��I�����āA����Ƃ�������̂��N���A
    With sht.Range(sht.Cells(ROW_START, COLUMN_PAT_NAME), rgLast)
        .ClearContents
        .Interior.ColorIndex = 0
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
            
    '�`������ǂ�
    Application.ScreenUpdating = True
    
    '��n��
    Set rgLast = Nothing
    
End Sub

'================================================================================
' Hash code Produce Functions
'================================================================================
Private Function CreateHashFile(ByVal strFileName As String, ByVal lngAlgID As Long) As String
    Dim abytData() As Byte
    Dim intFile As Integer
    Dim lngError As Long
    On Error Resume Next
        If Len(Dir(strFileName)) > 0 Then
            intFile = FreeFile
            Open strFileName For Binary Access Read Shared As #intFile
            abytData() = InputB(LOF(intFile), #intFile)
            Close #intFile
        End If
        lngError = Err.Number
    On Error GoTo 0
    If lngError = 0 Then CreateHashFile = CreateHashFromBinary(abytData(), lngAlgID) _
                    Else CreateHashFile = ""
End Function

   
' Create Hash
Private Static Function CreateHashFromBinary(abytData() As Byte, ByVal lngAlgID As Long) As String
    Dim hProv As Long, hHash As Long
    Dim abytHash(0 To 63) As Byte
    Dim lngLength As Long
    Dim lngResult As Long
    Dim strHash As String
    Dim i As Long
    strHash = ""
    If CryptAcquireContext(hProv, vbNullString, vbNullString, _
                           IIf(lngAlgID >= CALG_SHA_256, PROV_RSA_AES, PROV_RSA_FULL), _
                           CRYPT_VERIFYCONTEXT) <> 0& Then
        If CryptCreateHash(hProv, lngAlgID, 0&, 0&, hHash) <> 0& Then
            lngLength = UBound(abytData()) - LBound(abytData()) + 1
            If lngLength > 0 Then lngResult = CryptHashData(hHash, abytData(LBound(abytData())), lngLength, 0&) _
                             Else lngResult = CryptHashData(hHash, ByVal 0&, 0&, 0&)
            If lngResult <> 0& Then
                lngLength = UBound(abytHash()) - LBound(abytHash()) + 1
                If CryptGetHashParam(hHash, HP_HASHVAL, abytHash(LBound(abytHash())), lngLength, 0&) <> 0& Then
                    For i = 0 To lngLength - 1
                        strHash = strHash & Right$("0" & Hex$(abytHash(LBound(abytHash()) + i)), 2)
                    Next
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0&
    End If
    CreateHashFromBinary = LCase$(strHash)
End Function

' MD5
Public Static Function CreateMD5Hash(abytData() As Byte) As String
    CreateMD5Hash = CreateHashFromBinary(abytData(), CALG_MD5)
End Function

Public Static Function CreateMD5HashString(ByVal strData As String) As String
    CreateMD5HashString = CreateHashString(strData, CALG_MD5)
End Function
' Create Hash From String(Shift_JIS)
Private Static Function CreateHashString(ByVal strData As String, ByVal lngAlgID As Long) As String
    CreateHashString = CreateHashFromBinary(StrConv(strData, vbFromUnicode), lngAlgID)
End Function

