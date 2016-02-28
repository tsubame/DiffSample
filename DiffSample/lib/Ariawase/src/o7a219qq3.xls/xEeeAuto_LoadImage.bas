Attribute VB_Name = "xEeeAuto_LoadImage"
'�T�v:
'   �t�@�C�����[�h�̃e�X�g�C���X�^���X
'
'�ړI:
'   �v�����o���N�ɂȂ��Ƃ������t�@�C�������[�h����
'
'�쐬��:
'   2012/03/09 Ver0.1 D.Maruyama
'   2012/03/16 Ver0.2 D.Maruyama  145�ɂ��킹�ēǂݍ��݊g���q���Ƃ肠����".idp"�ɕύX
'                                 �t�@�C�����[�h���Ƀ��b�Z�[�W���o���悤�ɕύX
'   2012/11/12 Ver0.3 H.Arikawa   ������LoadRefImage�̊֐���ǉ��B
'   2013/01/24 Ver0.4 H.Arikawa   LoadRefImage�̃��C���f�o�b�O���ʔ��f(�R���p�C���m�F)�B

Option Explicit

Private Const EEE_AUTO_LOADIMAGE_ARGS As Long = 5

Private Function EeeAutoLoadImage_f() As Double

    Dim site As Long
    
    Dim strPlaneBankName As String
    Dim strPlaneGroup As String
    Dim eBitDepth As IdpBitDepth
    Dim strDstZone As String
    Dim strFilePath As String
    
    On Error GoTo ErrorHandler
    
    'TestInstanse���p�����[�^���擾
    If Not LoadImage_GetParameter( _
        strPlaneBankName, _
        strPlaneGroup, _
        eBitDepth, _
        strDstZone, _
        strFilePath) Then
            Err.Raise 9999, "EeeAutoLoadImage_f", "Invalid argment count!"
    End If

    '�v�����o���N�ɂ���ꍇ�͔�����
    If TheIDP.PlaneBank.isExisting(strPlaneBankName) Then
        Exit Function
    End If

    '�t�@�C���t���p�X�̐���
    Dim strFileFullPath As String
    strFileFullPath = strFilePath & "\" & strPlaneBankName & ".idp"
    
    '�v�����m�ہA�̈�ݒ�
    Dim dstPlane As CImgPlane
    Call GetFreePlane(dstPlane, strPlaneGroup, eBitDepth, False, strPlaneBankName)
    Call dstPlane.SetPMD(strDstZone)
    
    
    '�S�T�C�g�Ƀt�@�C�����[�h
    For site = 0 To nSite
         TheExec.Datalog.WriteComment "Image Data " & strFileFullPath & " is loading for Site" & CStr(site)
        Call dstPlane.ReadFile(site, strFileFullPath)
    Next site
    
    'PlaneBank�ɓo�^
    Call TheIDP.PlaneBank.Add(strPlaneBankName, dstPlane, True, True)
    
    EeeAutoLoadImage_f = TL_SUCCESS
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    EeeAutoLoadImage_f = TL_ERROR
    
End Function

Private Function LoadImage_GetParameter( _
    ByRef strPlaneBankName As String, _
    ByRef strPlaneGroup As String, _
    ByRef eBitDepth As IdpBitDepth, _
    ByRef strDstZone As String, _
    ByRef strFilePath As String _
    ) As Boolean

    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_LOADIMAGE_ARGS) Then
        LoadImage_GetParameter = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strPlaneBankName = ArgArr(0)
    strPlaneGroup = ArgArr(1)
    eBitDepth = CovertIdpDepth(ArgArr(2))
    strDstZone = ArgArr(3)
    strFilePath = ArgArr(4)
On Error GoTo 0

    LoadImage_GetParameter = True
    Exit Function
    
ErrHandler:

    LoadImage_GetParameter = False
    Exit Function

End Function

Private Function CovertIdpDepth(ByVal strBitDepth As String) As IdpBitDepth

    Select Case strBitDepth
    Case "S16"
        CovertIdpDepth = idpDepthS16
    Case "S32"
        CovertIdpDepth = idpDepthS32
    Case "F32"
        CovertIdpDepth = idpDepthF32
    Case Else
        Err.Raise 9999, "EeeAutoLoadImage_f", "Invalid parameter"
    End Select
    
End Function

Public Sub LoadRefImage()

    Dim NumImageFile As Long
    Dim i As Long
    
    Dim Shts As Worksheet           'For Work Sheet Check
    
    Dim Cell_Position As String     'For Serch Cell Position
    Dim basePoint As Variant        'For Base Point Serch
    Dim EndPoint As Variant         'For End Point Serch
    Dim EPRow As Long               'End Point Row
    
    Dim LoadDataInputPlaneName() As String  'LoadRefImage�V�[�g����擾�����v���[����
    Dim LoadDataInputBasePlane() As String  'LoadRefImage�V�[�g����擾�����x�[�X�v���[����
    Dim LoadDataInputBitDepth() As String   'LoadRefImage�V�[�g����擾����BitDepth
    Dim LoadDataInputPMD() As String        'LoadRefImage�V�[�g����擾����PMD
    Dim LoadDataInputFilePlace() As String  'LoadRefImage�V�[�g����擾�����t�@�C���p�X
    
    Dim site As Long
    
    Set Shts = Sheets("LoadRefImage")
    
    
    '++++ Get Data From LoadRefImage Sheet ++++
    
    '--CountRefImage--
    Cell_Position = "B4"
    
    Set EndPoint = Shts.Range(Cell_Position).End(xlDown)
    EPRow = EndPoint.Row
        
    If EPRow = 65536 Then
        NumImageFile = 0
    Else
        NumImageFile = EPRow - 4
    End If
    '--CountRefImage--
    
    If NumImageFile = 0 Then GoTo SkipLoadRefImage
    
    '--GetData--
    ReDim LoadDataInputPlaneName(NumImageFile - 1)
    ReDim LoadDataInputBasePlane(NumImageFile - 1)
    ReDim LoadDataInputBitDepth(NumImageFile - 1)
    ReDim LoadDataInputPMD(NumImageFile - 1)
    ReDim LoadDataInputFilePlace(NumImageFile - 1)
    
    For i = 0 To NumImageFile - 1
        LoadDataInputPlaneName(i) = Shts.Cells(5 + i, 2)
        LoadDataInputBasePlane(i) = Shts.Cells(5 + i, 3)
        LoadDataInputBitDepth(i) = Shts.Cells(5 + i, 4)
        LoadDataInputPMD(i) = Shts.Cells(5 + i, 5)
        LoadDataInputFilePlace(i) = Shts.Cells(5 + i, 6)
    Next i
    '--GetData--
    
    '++++ Get Data From LoadRefImage Sheet ++++



    '++++ Input Image Process ++++
    For i = 0 To NumImageFile - 1

        Dim BitDepth As IdpBitDepth
        Select Case LoadDataInputBitDepth(i)
            Case "S16"
                BitDepth = idpDepthS16
            Case "S32"
                BitDepth = idpDepthS32
            Case "F32"
                BitDepth = idpDepthF32
            Case Else
                GoTo ErrHandler
        End Select
        
        If TheIDP.PlaneBank.isExisting(LoadDataInputPlaneName(i)) Then
            '========= �v���[���o���N�ɓo�^���ꂽ�摜�̍폜�����s =========
            Call TheIDP.PlaneBank.Delete(LoadDataInputPlaneName(i))
        End If

        '====================== Input RefImage =======================
        Dim refPlane As CImgPlane
        Call GetFreePlane(refPlane, LoadDataInputBasePlane(i), BitDepth, True, "refPlane")
        For site = 0 To nSite
            Call InPutImage(site, refPlane, LoadDataInputPMD(i), LoadDataInputFilePlace(i) & LoadDataInputPlaneName(i) & ".stb")
        Next site
        Call TheIDP.PlaneBank.Add(LoadDataInputPlaneName(i), refPlane, True, True)
        '=============================================================
    Next i

    '++++ Input Image Process ++++

SkipLoadRefImage:
    Exit Sub

ErrHandler:
    Call MsgBox("BitDepth is Wrong", vbOKOnly)
    Call DisableAllTest 'EeeJob�֐�
End Sub

