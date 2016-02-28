VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgKernelManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   Kernel���̊Ǘ��p
'�ړI:
'   Kernel�̓o�^���₢���킹�p�B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��30��

Option Explicit

Private Const ERR_NUMBER = 9999                         ' Error�ԍ�
Private Const CLASS_NAME = "CImgKernelManager"   ' Class����

Private Const SHT_LABEL_KERNEL_NAME As String = "Kernel Name"
Private Const SHT_LABEL_KERNEL_WIDTH As String = "Width"
Private Const SHT_LABEL_KERNEL_HEIGHT As String = "Height"
Private Const SHT_LABEL_KERNEL_ANCHOR_X As String = "X Anchor"
Private Const SHT_LABEL_KERNEL_ANCHOR_Y As String = "Y Anchor"
Private Const SHT_LABEL_KERNEL_SHIFTR As String = "ShiftR"
Private Const SHT_LABEL_KERNEL_TYPE As String = "Type"
Private Const SHT_LABEL_KERNEL_COMMENT As String = "Comment"
Private Const SHT_LABEL_VAL As String = "Val "
Private Const SHT_LABEL_VAL_ARG_START As Integer = 1
Private Const SHT_LABEL_VAL_ARG_END As Integer = 64

Private mKernelInfos As Collection
Private mTmpKernelInfos As Collection   '�J�[�l�����
Private mKernelErr As String                'EeeJOB�`�F�b�N�ł̃G���[�L�^�p
Private mFlagKernelError As Boolean         'IG-XL�G���[�������������ǂ����̊m�F�p�BEeeJOB��IG-XL�ŏ��̍��������h�~�B

Public Sub Init()
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��10��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'
'

    Call RegKernelShtInfo

End Sub

Public Property Get IsErrIGXL() As Boolean
'���e:
'   TheIDP�����������ׂ����ۂ��̔��f�ޗ��ƂȂ���
'   ���̃t���O��True�̂Ƃ���TheIDP.Removeresouces����K�v������܂�
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��10��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Boolean
'
'���ӎ���:
'
'
    IsErrIGXL = mFlagKernelError

End Property

Public Sub CreateKernel( _
    pKernelName As String, ByVal pKernelType As IdpKernelType, _
    ByVal pNumCol As Long, ByVal pNumRow As Long, ByVal pShiftR As Long, ByVal pKernelVal As String, Optional ByVal pComment As String = "" _
)
'���e:
'   �J�[�l����`��ێ����܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��30��
'�p�����[�^:
'   [pKernelName]   In  1):�J�[�l����
'   [pKernelType]   In  2):�J�[�l���^�C�v
'   [pNumCol]       In  3):�������̐�
'   [pNumRow]       In  4):�c�����̐�
'   [pShiftR]       In  5):�E�V�t�gbit��
'   [pKernelVal]    In  6):�J�[�l����`�̕�����
'�߂�l:
'
'���ӎ���:
'

    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.CreateKernel(pKernelName, pKernelType, pNumCol, pNumRow, pShiftR, pKernelVal)

    '�e�[�u���Ƀf�[�^�ǉ�
    Dim newKernel As CImgKernelInfo
    Set newKernel = New CImgKernelInfo
    Call newKernel.Create(pKernelName, pKernelType, pNumCol, pNumRow, pShiftR, pKernelVal, pComment)
    Call mKernelInfos.Add(newKernel, pKernelName)
    
    Exit Sub
    
ERROR_IGXL:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "CreateKernel Error")


End Sub

Public Function IsExist(pKernelName As String) As Boolean
'���e:
'   �Ώۂ�Kernel�̗L����Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��30��
'�p�����[�^:
'   [pKernelName]  In  1):
'�߂�l:
'   Boolean
'
'���ӎ���:
'
    On Error GoTo NOT_FOUND
    Call mKernelInfos.Item(pKernelName)
    IsExist = True
    Exit Function
NOT_FOUND:
    IsExist = False
End Function

Public Function Kernel(pKernelName As String) As CImgKernelInfo
'���e:
'   �Ώۂ�Kernel�̃C���X�^���X��Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N12��2��
'�p�����[�^:
'   [pKernelName]  In  1):Kernel��
'�߂�l:
'   CImgKernelInfo
'
'���ӎ���:
'   CImgIDP�ɂ����l�̃��\�b�hKernel������܂��B
'   ���o�[�W�����Ƃ̌݊��m�ۂ̂��߂Ɏc���Ă܂��B

    On Error GoTo NOT_FOUND
    Set Kernel = mKernelInfos.Item(pKernelName)
    Exit Function
NOT_FOUND:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "This Kernel Name is not found!" & vbCrLf & "Kernel Name:" & pKernelName)
End Function

Public Property Get Count() As Long
'���e:
'   ���̃N���X������Kernel�̐���Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N12��2��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Long
'
'���ӎ���:

    Count = mKernelInfos.Count

End Property

Private Sub RegKernelShtInfo()
'���e:
'   kernel�V�[�g������΁A�������ēo�^����
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��3��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'

    '#### �O���[�v ####
    Dim pGroupErrCnt As Integer
    Dim pGroupStart As Integer
    Dim pGroupEnd As Integer
    Dim pGroupInfo As Collection
    Dim pTmp As Variant
    
    Set pGroupInfo = New Collection
    
    '#### �s�̃f�[�^�m�F�J�E���^ ####
    Dim intStartRow As Integer
    Dim intDataCnt As Integer
    Dim intGroupRowCnt As Integer
    
    Dim pAllKernelErrCnt As Long
    Dim pKernelErrCnt As Long
    
    '#### Kernel�̃p�����[�^ ####
    Dim pName As String
    Dim pWidth As Integer
    Dim pHeight As Integer
    Dim pShiftR As Integer
    Dim pKernelType As IdpKernelType
    Dim pData As String
    Dim pComment As String
    
    Dim pNameForChk As Collection
    
    Dim pNowHeight As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim pStr As String
    
    '#####  SheetReader�𗘗p���āAKernel�V�[�g��ǂݍ���  #####
    On Error GoTo NOTHING_SHT       '�V�[�g�ǂݍ��݃G���[�̏ꍇ(�V�[�g�������ꍇ�������Ȃ�)
    Dim clsWrkShtRdr As CWorkSheetReader
    Set clsWrkShtRdr = GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeKernel)
    On Error GoTo 0
    
    Dim IFileStream As IFileStream
    Dim IParamReader As IParameterReader

    With clsWrkShtRdr
        Set IFileStream = .AsIFileStream
        Set IParamReader = .AsIParameterReader
    End With
    
    mKernelErr = "Image Kernel Definition Sheet Error!" & vbCrLf
    
    '#####  Kernel�V�[�g�̓ǂݍ���  #####
    Set pNameForChk = New Collection
    Do While Not IFileStream.IsEOR

        '#####  �J�[�l�����̂��擾  #####
        If IParamReader.ReadAsString(SHT_LABEL_KERNEL_NAME) <> "" Then
            '##### �O�s�܂ł̒�`�I�[����(�����ƒ�`���ꂽ���̂ݓo�^�B�G���[�͍Ō�ɂ܂Ƃ߂ďo��) #####
            If pName <> "" Then         '�ŏ��̒�`�ȊO
                If pNowHeight <> pHeight Then   '�f�[�^�����`�F�b�N
                    pKernelErrCnt = pKernelErrCnt + 1
                End If
                If pKernelErrCnt = 0 Then
                    pData = LTrim(pData)        '�ŏ��̃X�y�[�X�����폜
                    Call StockKernelDefinition(pName, pKernelType, pWidth, pHeight, pShiftR, pData, pComment)
                Else
                    Call StockErr(pName)
                End If
            End If
            
            '#####  �J�[�l���p�����[�^���擾  #####
            pName = IParamReader.ReadAsString(SHT_LABEL_KERNEL_NAME)
            pWidth = IParamReader.ReadAsInteger(SHT_LABEL_KERNEL_WIDTH)
            pHeight = IParamReader.ReadAsInteger(SHT_LABEL_KERNEL_HEIGHT)
            pShiftR = IParamReader.ReadAsInteger(SHT_LABEL_KERNEL_SHIFTR)
            pKernelType = CIdpKernel(IParamReader.ReadAsString(SHT_LABEL_KERNEL_TYPE))
            pComment = IParamReader.ReadAsString(SHT_LABEL_KERNEL_COMMENT)
            
            pData = ""
            pAllKernelErrCnt = pAllKernelErrCnt + pKernelErrCnt
            pKernelErrCnt = 0
            pNowHeight = 0
            
            If IsKey(pName, pNameForChk) = True Then
                pKernelErrCnt = 1
            Else
                Call pNameForChk.Add(pName, pName)
            End If
            If ChkSize(pWidth) = False Or _
               ChkSize(pHeight) = False Or _
               ChkShiftR(pShiftR) = False Or _
               pKernelType = -1 Then
                pKernelErrCnt = 1
            End If
        
        End If
        
        '#####  �J�[�l����`�f�[�^���擾  #####
        For i = SHT_LABEL_VAL_ARG_START To SHT_LABEL_VAL_ARG_END
            pStr = IParamReader.ReadAsString(SHT_LABEL_VAL & i)
            If pStr = "" Then
                If i <= pWidth Then
                    '��`�f�[�^�̕���Width�ݒ�ȉ�
                    pKernelErrCnt = pKernelErrCnt + 1
                End If
            Else
                If i > pWidth Then
                    '��`�f�[�^�̕���Width�ݒ�����傫��
                    pKernelErrCnt = pKernelErrCnt + 1
                Else
                    pData = pData & " " & pStr
                End If
            End If
        Next i
        
        pNowHeight = pNowHeight + 1

        IFileStream.MoveNext
    Loop

    '##### �O�s�܂ł̒�`�I�[����(�����ƒ�`���ꂽ���̂ݓo�^�B�G���[�͍Ō�ɂ܂Ƃ߂ďo��) #####
    If pName <> "" Then         '�ŏ��̒�`�ȊO
        If pNowHeight <> pHeight Then   '�f�[�^�����`�F�b�N
            pKernelErrCnt = pKernelErrCnt + 1
        End If
        If pKernelErrCnt = 0 Then
            pData = LTrim(pData)        '�ŏ��̃X�y�[�X�����폜
            Call StockKernelDefinition(pName, pKernelType, pWidth, pHeight, pShiftR, pData, pComment)
        Else
            Call StockErr(pName)
        End If
    End If
    pAllKernelErrCnt = pAllKernelErrCnt + pKernelErrCnt
    
    '�V�[�g�̐��`
    Call ControlShtFormatKernel
    
    Set clsWrkShtRdr = Nothing
    Set IFileStream = Nothing
    Set IParamReader = Nothing
    Set pGroupInfo = Nothing
    
    If pAllKernelErrCnt > 0 Then
'        MsgBox mKernelErr
        Set mTmpKernelInfos = New Collection
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME, mKernelErr)
    End If
        Call TriggerRegDifinition
    Exit Sub

NOTHING_SHT:
    
End Sub

Private Sub StockErr(ByVal pName As String)
'���e:
'   �G���[�����߂Ă���
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��4��
'�p�����[�^:
'   [pName] In/Out  1):
'�߂�l:
'
'���ӎ���:
'
'
    mKernelErr = mKernelErr & vbCrLf & "Definition Name: " & pName

End Sub

Private Sub StockKernelDefinition( _
    pKernelName As String, ByVal pKernelType As IdpKernelType, _
    ByVal pNumCol As Long, ByVal pNumRow As Long, ByVal pShiftR As Long, ByVal pKernelVal As String, Optional ByVal pComment As String = "" _
)
'�V�[�g��̒�`����~���Ă����܂��B

    '�e�[�u���Ƀf�[�^�ǉ�
    Dim newKernel As CImgKernelInfo
    Set newKernel = New CImgKernelInfo
    Call newKernel.Create(pKernelName, pKernelType, pNumCol, pNumRow, pShiftR, pKernelVal, pComment)
    Call mTmpKernelInfos.Add(newKernel, pKernelName)

End Sub

Private Sub TriggerRegDifinition()
'���e:
'   �~�����V�[�g��̒�`��IG-XL�ɓo�^���܂��B
'   �S���̃f�[�^�����������Ƃ��m�F���āAIG-XL�ɓo�^���܂��B
'   ���p��TheIDP�̔j��������܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��10��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'
'


    On Error GoTo ERROR_IGXL
    
    Dim pKernel As CImgKernelInfo
    For Each pKernel In mTmpKernelInfos
    
        Call TheHdw.IDP.CreateKernel(pKernel.Name, pKernel.KernelType, pKernel.width, pKernel.height, pKernel.ShiftR, pKernel.Value)
    
    Next pKernel

    'EeeJOB�`�F�b�N���Ȃ��AIG-XL�o�^���Ȃ��B�J�[�l���V�[�g���o�^����
    Set mKernelInfos = mTmpKernelInfos
    Set mTmpKernelInfos = New Collection
    
    Exit Sub

ERROR_IGXL:
    '���̃G���[�̏ꍇ�AEeeJOB��IG-XL�ŏ��ɍ������������Ă�̂�
    'mFlagKernelError��True�ɂ��āA�O������TheIDP.RemoveResouces�ł���悤�ɂ��Ă����B
    Set mTmpKernelInfos = New Collection
    mFlagKernelError = True
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, mKernelErr)

End Sub

Private Function ChkShiftR(ByVal pShiftRbit As Integer) As Boolean
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��3��
'�p�����[�^:
'   [pShiftRbit]    In/Out  1):
'�߂�l:
'   Integer
'
'���ӎ���:
'
'
    If pShiftRbit >= 0 And pShiftRbit <= 16 Then
        ChkShiftR = True
    Else
        ChkShiftR = False
    End If
    
End Function

Private Function CIdpKernel(ByVal pKernelType As String) As IdpKernelType
'���e:
'   ��������idpKernelType�ɕϊ��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��3��
'�p�����[�^:
'   [pKernelType]   In/Out  1):
'�߂�l:
'   IdpKernelType
'
'���ӎ���:
'   �Y�����Ȃ��ꍇ��-1

    pKernelType = UCase(pKernelType)    '�召��������
    
    If pKernelType = "INTEGER" Then
        CIdpKernel = idpKernelInteger
    ElseIf pKernelType = "FLOAT" Then
        CIdpKernel = idpKernelFloat
    Else
        CIdpKernel = -1
    End If

End Function

Private Function ChkSize(pSize As Integer) As Boolean
'���e:
'   �傫����1�`64���`�F�b�N���Ė��Ȃ���΂��̂܂ܕԂ��B�G���[�Ȃ�-1��Ԃ��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N3��4��
'�p�����[�^:
'   [pSize] In/Out  1):
'�߂�l:
'   Integer
'
'���ӎ���:
'
'
    If pSize >= 1 And pSize <= 64 Then
        ChkSize = True
    Else
        ChkSize = False
    End If

End Function

Private Function IsKey(ByVal pKey As String, ByRef pObj As Collection) As Boolean
'���e:
'   �Y��Collection�I�u�W�F�N�g�ɃL�[�����݂��邩���ׂ�
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��22��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Boolean :True���ł�add�ς݁B���݂��܂��B       False�܂�����
'
'���ӎ���:
'
    On Error GoTo ALREADY_REG
    Call pObj.Item(pKey)
    IsKey = True
    Exit Function

ALREADY_REG:
    IsKey = False

End Function

Private Sub Class_Initialize()
'���e:
'   �R���X�g���N�^
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��30��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'
'
    Set mKernelInfos = New Collection
    Set mTmpKernelInfos = New Collection

End Sub

Private Sub Class_Terminate()
'���e:
'   �f�X�g���N�^
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��30��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'
'
    Set mKernelInfos = Nothing
    Set mTmpKernelInfos = Nothing

End Sub