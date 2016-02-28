VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �v���[�����̃��b�p�[�B(��CTraceablePlane)
'   �f�o�C�X�v���t�@�C���Ȃǃv���[���̐������̏����Z�b�g�ŕێ�����B
'
'   Revision History:
'       Data        Description
'       2011/02/01  TraceablePlane���C���^�[�t�F�[�X������CParamPlane�Ƃ��Ď���
'       2011/02/15  �V�i���I�G���W���g�ݍ��݂ɔ����������C��
'       2011/02/18  Plane,DevInfo��WriteOnce�G���[�������폜(�㏑���\), Dump�̈���pWriter��ByVal=>ByRef
'       2011/02/23  �������،��ʂ��ŐV�\�[�X�擾�BInitialize Byval => ByRef
'       2011/11/01  Read�̎���
'
'�ړI:
'   �f�o�C�X���ACImgPlane�������A�K�v�ɉ����ăf�o�C�X���A�摜�C���X�^���X��Ԃ��B
'
'�쐬��:
'   0145184304
'

Option Explicit

Implements IProcParameter
Implements IProcParamReplaceable

Private Const ERR_NUMBER As Long = 9999                  ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CParamPlane"       ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default�l"Empty"��ێ�����

Private Const DUMPKEYWORD_FOR_PARAM_NAME As String = "Parameter Name"
Private Const DUMPKEYWORD_FOR_CLASS_NAME As String = "Parameter Type"
Private Const DUMPKEYWORD_FOR_COMMENT As String = "Comment"

Private Const PARAMETER_REPLACER_NAME As String = "Name"
Private Const PARAMETER_REPLACER_PLANE As String = "Plane"
Private Const PARAMETER_REPLACER_DEVICEINFO As String = "DeviceInfo"
Private Const PARAMETER_REPLACER_COMMENT As String = "Comment"
Private Const PARAMETER_REPLACER_TYPE As String = "Type"

Private mName As String
Private mCImgPlane As CImgPlane
Private mCDeviceConfigInfo As CDeviceConfigInfo
Private mComment As String

Public Sub Initialize(ByVal paramName As String, ByRef clsImgPlane As CImgPlane, ByRef clsDevInfo As CDeviceConfigInfo, Optional ByVal pComment As String = "")
'���e:
'   ������
'
'[paramName]   IN String�^:     �p�����[�^�����w��
'[clsImgPlane] IN CImgPlane�^:  �i�[����CImagPlane���w��
'[clsDevInfo]  IN CDeviceConfigInfo�^:  �i�[����CDeviceConfigInfo���w��
'
'���l:
'

    mName = paramName
    
    Set mCImgPlane = clsImgPlane
    Set mCDeviceConfigInfo = clsDevInfo
    
    mComment = pComment

End Sub
Private Function IProcParameter_Clone() As IProcParameter
'���e:
'   Clone���쐬����
'
'�߂�l:
'   CParamPlane��Ԃ��i�N���[���j
'
'���l:
'
    
    '##### ���g�������Ă���DeviceInfo�̃N���[������� #####
    Dim objFrameRate As CParamFrameRate
    Dim objAccTime As CParamAccTime
    Dim objGain As CParamGain
    Dim objLSB As CParamLsb
    
    Dim objDevInfo As CDeviceConfigInfo
    
    If Not Me.DeviceConfigInfo Is Nothing Then
        Set objDevInfo = Me.DeviceConfigInfo.AsIProcParameter.Clone
    Else
        Set objDevInfo = Nothing
    End If
    
    '#####  ���g��Clone���쐬���Ė߂�  #####
    Dim cloneobj As CParamPlane
    Set cloneobj = New CParamPlane
    cloneobj.Initialize Me.AsIProcParameter.Name, Me.plane, objDevInfo, Me.AsIProcParameter.Comment
    Set IProcParameter_Clone = cloneobj
    Set cloneobj = Nothing

    Set objDevInfo = Nothing
    Set objFrameRate = Nothing
    Set objAccTime = Nothing
    Set objGain = Nothing
    Set objLSB = Nothing

End Function
Private Sub IProcParameter_Dump(ByRef pWriter As IProcParamWriter)
'���e:
'   Plane��DeviceInfo��Comment��Dump����(pWriter�Ƀf�[�^�������n��)
'
'[pWriter]   IN IProcParamWriter�^:     Dump����Writer���w�肷��
'
'���l:
'

    '#####  ���g�̃N���X�� / �v���[�� / �f�o�C�X��� / �R�����g ��Dump #####
    ' Parameter Block
    pWriter.WriteAsString "", "Parameter " & Me.AsIProcParameter.Name & " {"
    pWriter.WriteAsBlockStart "Dump"
    pWriter.WriteAsString "", "Type: " & CLASS_NAME                             ' �N���X��
    ' Plane Block
    pWriter.WriteAsString "", "Plane {"
    pWriter.WriteAsBlockStart Me.AsIProcParameter.Name
    pWriter.WriteAsPlane Me.AsIProcParameter.Name, Me.plane                     ' �v���[��
    pWriter.WriteAsBlockEnd "ParamPlane"
    pWriter.WriteAsString "", "}"
    ' DeviceConfigInfo Block
    If Not Me.DeviceConfigInfo Is Nothing Then
        Me.DeviceConfigInfo.AsIProcParameter.Dump pWriter                           ' �f�o�C�X���
    End If
    pWriter.WriteAsString "", "Comment: " & Me.AsIProcParameter.Comment         ' �R�����g
    pWriter.WriteAsBlockEnd "Dump"
    pWriter.WriteAsString "", "}"
    
    
End Sub

Private Property Let IProcParameter_Name(ByVal pName As String)
'���e:
'   �p�����[�^����ێ�����
'
'[pName]   IN String�^:     �ێ�����p�����[�^�����w�肷��
'
'���l:
'

    mName = pName

End Property

Private Property Get IProcParameter_Name() As String
'���e:
'   �ێ������p�����[�^����Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^��
'
'���l:
'

    IProcParameter_Name = mName

End Property

Private Property Let IProcParameter_Comment(ByVal pComment As String)
'���e:
'   �R�����g��ێ�����
'
'[pComment]   IN String�^:     �ێ�����R�����g���w�肷��
'
'���l:
'

    mComment = pComment

End Property

Private Property Get IProcParameter_Comment() As String
'���e:
'   �ێ������R�����g��Ԃ�
'
'�߂�l:
'   �ێ������R�����g
'
'���l:
'

    IProcParameter_Comment = mComment

End Property

Public Property Set plane(pPlane As CImgPlane)
'���e:
'   �v���[����ێ�����
'
'[pPlane]   IN CImgPlane�^:     �ێ�����v���[�����w�肷��
'
'���l:
'

        Set mCImgPlane = pPlane

End Property

Public Property Get plane() As CImgPlane
'���e:
'   �ێ������v���[����Ԃ�
'
'�߂�l:
'   �ێ������v���[��
'
'���l:
'

    Set plane = mCImgPlane

End Property


Public Property Set DeviceConfigInfo(pDeviceInfo As CDeviceConfigInfo)
'���e:
'   �f�o�C�X����ێ�����
'
'[pDeviceInfo]   IN CDeviceConfigInfo�^:     �ێ�����f�o�C�X�����w�肷��
'
'���l:
'
        Set mCDeviceConfigInfo = pDeviceInfo

End Property

Public Property Get DeviceConfigInfo() As CDeviceConfigInfo
'���e:
'   �ێ������f�o�C�X����Ԃ�
'
'�߂�l:
'   �ێ������f�o�C�X���
'
'���l:
'

    Set DeviceConfigInfo = mCDeviceConfigInfo

End Property

Public Function PixelLog(ByVal pFlgName As String, Optional ByVal pMaxNum As Long = 0, _
                            Optional ByVal pAddrMode As IdpAddrMode = idpAddrAbsolute) As CParamDefect
'���e:
'   �Y������t���O��Pixel���O��CParamDefect�^�ŕԂ��܂��B
'�p�����[�^:
'   [pFlgName]  In  1):
'   [pMaxNum]   In  2):
'   [pAddrMode] In  3):
'�߂�l:
'   CParamDefect
'
'���ӎ���:
'   CImgPlane��PixelLog��Sub�ł����A�������Function�ł��B


    '-- �ϐ���`�Ȃ� ---------------------------------------
    Dim pSiteCount As Long                       'SiteLoop�J�E���g�p
    Dim pTypePointDefects() As T_PIXINFO         '���׍ő�����̉�f���\����
    Dim pDefectPixel As CParamPixelLog           '1��f���̌��׏��
    Dim pThisSiteDefects As CParamSitePixelLog   '1Site���̓_���׏��
    Dim pAllSiteDefects() As CParamSitePixelLog  '�S�T�C�g�̓_���׏��S������
    Dim pDefects As CParamDefect                 '�S�T�C�g�̓_���׏��S������I�u�W�F�N�g
    Dim pColor As CImgColor                     '
    Dim pPixelCount As Long                      '���׌�
    '--------------------------------------------------------
    
    '-- ���������� ------------------------------------------
    '�S�T�C�g���׏��(���̓T�C�g�̗L�������Ɍ��炸�S�T�C�g�����p��)
    ReDim pAllSiteDefects(TheExec.sites.ExistingCount - 1)
    '--------------------------------------------------------
    
    'SiteLOOP�ŏ������s
    For pSiteCount = 0 To UBound(pAllSiteDefects) Step 1
        '�L���T�C�g�̂ݏ��擾
        If TheExec.sites.site(pSiteCount).Active = True Then
            Call Me.plane.PixelLog(pSiteCount, pFlgName, pTypePointDefects, pMaxNum, pAddrMode)
            '�T�C�g�����׏��쐬
            Set pThisSiteDefects = New CParamSitePixelLog
            Call pThisSiteDefects.Initialize(pSiteCount)
            '���׉�fLOOP
            If UBound(pTypePointDefects) <> -1 Then
                For pPixelCount = 0 To UBound(pTypePointDefects) Step 1
'With�X�e�[�g�����g�g���Ɠ��I�z�񂪃��b�N�����
'                    With pTypePointDefects(pPixelCount)
'                        If .X = 0 Or .Y = 0 Then
'                            Exit For
'                        End If
'                        Set pDefectPixel = New CParamPixelLog
'                        Set pColor = mCImgPlane.ColorByAddress(.X, .Y, pAddrMode)
'                        Call pDefectPixel.Initialize(.X, .Y, .Value, pColor)
'                    End With
                    If pTypePointDefects(pPixelCount).x = 0 Or pTypePointDefects(pPixelCount).y = 0 Then
                        Exit For
                    End If
                    Set pDefectPixel = New CParamPixelLog
                    Set pColor = mCImgPlane.ColorByAddress(pTypePointDefects(pPixelCount).x, pTypePointDefects(pPixelCount).y, pAddrMode)
                    Call pDefectPixel.Initialize(pTypePointDefects(pPixelCount).x, pTypePointDefects(pPixelCount).y, pTypePointDefects(pPixelCount).Value, pColor)
                    Call pThisSiteDefects.Add(pDefectPixel)
                Next
            End If
            '�T�C�g�̌��׏���ݒ�
            Set pAllSiteDefects(pSiteCount) = pThisSiteDefects
        Else
            '�T�C�g�̌��׏���ݒ�
            Set pAllSiteDefects(pSiteCount) = Nothing
        End If
    Next
    
    '�ŏI���ʓo�^����
    '�i�����T�C�g�̒l���擾�����ۂ͎��s���G���[�łȂ�Nothing���߂邱�ƂɂȂ�j
    Set pDefects = New CParamDefect
    For pSiteCount = 0 To UBound(pAllSiteDefects) Step 1
        Call pDefects.Add(pAllSiteDefects(pSiteCount), pSiteCount)
    Next
    Set pDefects.DeviceConfig = mCDeviceConfigInfo
    Set PixelLog = pDefects

End Function

Private Sub Class_Initialize()

    Set mCImgPlane = Nothing
    Set mCDeviceConfigInfo = Nothing

End Sub

Private Sub Class_Terminate()

    mName = INITIAL_EMPTY_VALUE
    Set mCImgPlane = Nothing
    Set mCDeviceConfigInfo = Nothing
    mComment = INITIAL_EMPTY_VALUE

End Sub
Public Function AsIProcParameter() As IProcParameter
    Set AsIProcParameter = Me
End Function

Public Sub IProcParameter_Read(ByRef pReader As IProcParamReader)
'���e:
'   �w��Plane��stb�t�@�C���摜��ǂݍ���
'
'����:
'[pReader]     IN/OUT IProcParamReader�^:   Reader
'
'���l:
'
'

    Call pReader.ReadAsBlockStart("ParamPlane")
    Call pReader.ReadAsPlane(Me.AsIProcParameter.Name, mCImgPlane)
    Call pReader.ReadAsBlockEnd("ParamPlane")
End Sub

Private Sub IProcParamReplaceable_GetContents(ByVal pRoot As String, ByRef pRes As Collection)
'���e:
'   �p�����[�^�̐ݒ���e���擾����
'
'�p�����[�^:
'   [pRoot]     IN  String�^:  ���̃p�����[�^�N���X�̊K�w���
'   [pRes]      IN/OUT  Collection�^:  �p�����[�^�̐ݒ���e���i�[����R���N�V����
'
'���ӎ���:
'
'

    Dim bufSubParam As IProcParamReplaceable
    With pRes
        '   2012-01-27 (0145184931): Name��ҏW�s�ɕύX
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_NAME, mName, "String", False, PARAMETER_REPLACER_NAME)
        Set bufSubParam = mCImgPlane
        If Not bufSubParam Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PLANE, vbNullString, "CImgPlane", False, PARAMETER_REPLACER_PLANE)
            Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_PLANE, pRes)
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PLANE, "Nothing", "CImgPlane", False, PARAMETER_REPLACER_PLANE)
        End If
        Set bufSubParam = mCDeviceConfigInfo
        If Not bufSubParam Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, vbNullString, "CDeviceConfigInfo", False, PARAMETER_REPLACER_DEVICEINFO)
            Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, pRes)
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, "Nothing", "CDeviceConfigInfo", False, PARAMETER_REPLACER_DEVICEINFO)
        End If
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COMMENT, mComment, "String", True, PARAMETER_REPLACER_COMMENT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_TYPE, CLASS_NAME, "String", False, PARAMETER_REPLACER_TYPE)
    End With
End Sub

Private Sub IProcParamReplaceable_ReplaceContent(ByVal pRoot As String, ByRef pParamContents As CImpLocalParamsContent)
'���e:
'   �p�����[�^�̓��e������������
'
'�p�����[�^:
'   [pRoot]             IN  String�^:  ���̃p�����[�^�N���X�̊K�w���
'   [pParamContents]    IN  CImpLocalParamsContent�^:  �����������s���p�����[�^�̓��e
'
'���ӎ���:
'
'

    Dim targetSubParam As IProcParamReplaceable
    Dim targetIndex As Long
    With pParamContents
        Dim subParamLabel As String
        subParamLabel = XLibImpExperimentUtility.GetSubParamLabel(.Path, pRoot)
        Select Case subParamLabel
'   2012-01-27 (0145184931): Name��ҏW�s�ɕύX
'            Case PARAMETER_REPLACER_NAME
'                mName = .value
            Case PARAMETER_REPLACER_COMMENT
                mComment = .Value
        End Select
    End With
End Sub

Private Sub AddParamContents(ByRef pRes As Collection, ByVal pPath As String, ByVal pValue As String, ByVal pTypeName As String, ByVal pEditable As Boolean, ByVal pLabel As String)
'���e:
'   �p�����[�^�̏����R���N�V�����ɒǉ�����
'
'�p�����[�^:
'   [pRes]          IN/OUT  Collection�^:  �p�����[�^�̐ݒ���e��ǉ�����R���N�V����
'   [pPath]         IN  String�^:  �p�����[�^�̊K�w���
'   [pValue]        IN  String�^:  �p�����[�^�̒l
'   [pTypeName]     IN  String�^:  �p�����[�^�̃f�[�^�^
'   [pEditable]     IN  Boolean�^:  �p�����[�^���ҏW�\���ǂ���
'   [pLabel]        IN  String�^:  �p�����[�^��
'
'���ӎ���:
'
'

    Dim myContents As New CImpLocalParamsContent
    Call myContents.Initialize(pPath, pValue, pTypeName, pEditable, pLabel)
    Call pRes.Add(myContents)
End Sub