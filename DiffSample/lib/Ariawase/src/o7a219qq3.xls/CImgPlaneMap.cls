VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �J���[�}�b�v�̏��
'�ړI:
'   �J���[�}�b�v���P�����Ǘ����܂��B
'   �J���[�}�b�v�Ɋւ���String�����AidpColorType�̏��ɉ��߂��܂�
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��22��
'   2011/02/02  CImgPlaneMap��IdpColorAll��IdpColorFlat�͖ʓ|���Ȃ�(CImgPlane������)
'               �����̎w�肪���ꂽ�ꍇ�̓G���[�B
'   2012/02/20  idpColorType��0�Ԃ��珇�Ɋ���U��d�g�݂���ACIdpColorMapper�N���X��
'               �\�ߊ���U���Ă�����idpColorType���g�p����悤�ɕύX�B

Option Explicit

Implements IProcParamReplaceable

Private Const ERR_NUMBER = 9999                         ' Error�ԍ�
Private Const CLASS_NAME = "CImgPlaneMap"           ' Class����

Private Const ORIGIN_X As Integer = 1
Private Const ORIGIN_Y As Integer = 1

Private Const PARAMETER_REPLACER_MAPNAME As String = "MapName"
Private Const PARAMETER_REPLACER_COMMENT As String = "Comment"
Private Const PARAMETER_REPLACER_WIDTH As String = "Width"
Private Const PARAMETER_REPLACER_HEIGHT As String = "Height"
Private Const PARAMETER_REPLACER_COLORLIST As String = "ColorList"
Private Const PARAMETER_REPLACER_IDPCOLORSPMD As String = "IdpColorsPmd"

Private m_MapName As String                     '���̃N���X��Color Map Name
Private m_Comment As String

Private m_Width As Integer                  '���̃J���[�}�b�v�̕�
Private m_Height As Integer                 '���̃J���[�}�b�v�̍���

'For PMD�V�[�g�t�H�[�}�b�g
Private m_ColorNamesPmd As String                        'String PMD�t�H�[�}�b�g�B�J���}�A�J�b�R��؂�B
Private m_IdpColorsPmd As String                   'idpColorType�ԍ� PMD�t�H�[�}�b�g�B�J���}�A�J�b�R��؂�B

'For idp <=> Str�@Change
Private m_IdpStrs As Collection              '���[�U�[��`��idpColorType�����L�[��String������
Private m_StrIdps As Collection              '���[�U�[��`��String�����L�[��idpColorType������

'For Colors
Private m_Colors As Collection              '���[�U�[��`��String�����L�[��CImgColor������

'For Width,Height,retArray
Private m_ArrayColorName() As String
Private m_ArrayIdpColorNumber() As IdpColorType

'For idpColorMapper
Private m_IdpColorMapper As CIdpColorMapper     'String -> idpColorType�ϊ���������

'OK
Public Sub Create(ByVal pMapName As String, ByRef pMap() As String, ByRef pIdpColorMapper As CIdpColorMapper, Optional pComment As String = "")
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��25��
'�p�����[�^:
'   [pMapName]          In  1):Color Map Name
'   [pMap]              In  2):Color Map ��`������(���[�U�[��`�̕�����ԂŔz��^)
'   [pIdpColorMapper]   In  3):�F��->idpColorType�ϊ����
'   [pComment]          In  4):�R�����g
'�߂�l:
'
'���ӎ���:
'

    '�}�b�v���쐬
    m_MapName = pMapName
    m_Comment = pComment

    Set m_IdpColorMapper = pIdpColorMapper
    
    'idpColorType�֕ϊ�
    Call StrToIdpColor(pMap)
    
    'CImgColor,CImgPixel����
    Dim p_Color As Variant
    Dim p_X As Long, p_Y As Long
    Dim p_ColorObj As CImgColor
    Dim p_PixelObj As CImgPixel
    Dim p_PixelObjs As Collection
    
    Set m_Colors = New Collection
    For Each p_Color In m_IdpStrs
        Set p_PixelObjs = New Collection
        For p_Y = ORIGIN_Y To m_Height
            For p_X = ORIGIN_X To m_Width
                If CStr(p_Color) = m_ArrayColorName(p_X, p_Y) Then
                    Set p_PixelObj = New CImgPixel
                    Call p_PixelObj.Init(p_X, p_Y)
                    Call p_PixelObjs.Add(p_PixelObj)
                    Set p_PixelObj = Nothing
                End If
            Next p_X
        Next p_Y
        Set p_ColorObj = New CImgColor
        Call p_ColorObj.Init(Me.Name, CStr(p_Color), CIdpColor(m_StrIdps.Item(CStr(p_Color))), p_PixelObjs)
        Call m_Colors.Add(p_ColorObj, CStr(p_Color))
        Set p_PixelObjs = Nothing
    Next p_Color
    
End Sub

'OK
Public Function Clone() As CImgPlaneMap
'���e:
'   ���̃C���X�^���X�̃N���[����Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��25��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   CImgPlaneMap
'
'���ӎ���:
'

    Dim pClone As CImgPlaneMap
    Dim pStr() As String
    Set pClone = New CImgPlaneMap
    
    With Me
        pStr = m_ArrayColorName
        Call pClone.Create(.Name, pStr, m_IdpColorMapper, .Comment)
    End With
    
    Set Clone = pClone
    
End Function

Public Property Get ColorList() As String()
'���e:
'   �o�^�ς݂�Color Map����String�^�̓��I�z��ŕԂ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N1��25��
'�p�����[�^:
'   [)] In  1):
'�߂�l:
'   String()
'
'���ӎ���:
'
'

    Dim pTmp As Variant
    Dim pArray() As String
    Dim PNum As Long

    ReDim pArray(m_IdpStrs.Count - 1)
    
    For Each pTmp In m_IdpStrs
        pArray(PNum) = pTmp
        PNum = PNum + 1
    Next pTmp

    ColorList = pArray

End Property

'OK
Public Property Get Name() As String
'���e:
'   ���̃C���X�^���X�̖��O(Color Map Definition �ɒ�`���ꂽ���O)���擾
'
'���l:
'
    Name = m_MapName
End Property

'OK
Public Property Get width() As Long
'���e:
'   �}�b�v�̕����擾
'
'���l:
'
    width = m_Width
End Property

'OK
Public Property Get height() As Long
'���e:
'   �}�b�v�̍������擾
'
'���l:
'
    height = m_Height
End Property

'OK
Public Property Get Comment() As String
'���e:
'   �}�b�v�̃R�����g���擾
'
'���l:
'
    Comment = m_Comment
End Property

'OK
Public Property Get IdpColorAddress(ByVal pX As Long, ByVal pY As Long) As IdpColorType
'���e:
'   �w��A�h���X�̐F�����擾
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��22��
'�p�����[�^:
'   [pX]    In  1):     X�A�h���X(1�X�^�[�g)
'   [pY]    In  2):     Y�A�h���X(1�X�^�[�g)
'�߂�l:
'   idpColorType
'
'���ӎ���:
'

    If (pX <= 0) Or (pX > m_Width) Or (pY <= 0) Or (pY > m_Height) Then
        Call TheError.Raise(9999, "Color Map Infomation", "Invalid Parameter")
        Exit Property
    End If
    
    IdpColorAddress = m_ArrayIdpColorNumber(pX, pY)

End Property

'OK
Public Property Get ColorNameAddress(ByVal pX As Long, ByVal pY As Long) As String
'���e:
'   �w��A�h���X�̐F�����擾
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��22��
'�p�����[�^:
'   �w��A�h���X�̐F�����擾
'   [pX]    In  1):     X�A�h���X(1�X�^�[�g)
'   [pY]    In  2):     Y�A�h���X(1�X�^�[�g)
'�߂�l:
'   String
'
'���ӎ���:
'

    If (pX <= 0) Or (pX > m_Width) Or (pY <= 0) Or (pY > m_Height) Then
        Call TheError.Raise(9999, "Color Map Infomation", "Invalid Parameter")
        Exit Property
    End If
    
    ColorNameAddress = m_ArrayColorName(pX, pY)

End Property

'OK
Public Property Get IdpColorsPmd() As String
'���e:
'   �}�b�v��`�̕������Ԃ��i���[�U�[��`idpColorType���j
'
'���l:
'
    IdpColorsPmd = m_IdpColorsPmd
End Property

'OK
Public Property Get ColorNamesPmd() As String
'���e:
'   �}�b�v��`�̕������Ԃ��i���[�U�[��`String���j
'
'���l:
'
    ColorNamesPmd = m_ColorNamesPmd
End Property

Public Property Get ColorIdpType(p_IdpColorType As IdpColorType) As CImgColor
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��1��
'�p�����[�^:
'   [pIdpColor] In  1):
'�߂�l:
'   CImgColor
'
'���ӎ���:
'

    On Error GoTo NOTHING_KEY
    Set ColorIdpType = m_Colors.Item(CStr(ColorName(p_IdpColorType))).Clone
    Exit Property
NOTHING_KEY:
    Call TheError.Raise(9999, "idp", "This idpColorType is no use! Color Map Definition: " & p_IdpColorType)

End Property

Public Property Get color(p_ColorName As String) As CImgColor
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��1��
'�p�����[�^:
'   [pColorName]    In  1):
'�߂�l:
'   CImgColor
'
'���ӎ���:
'

    On Error GoTo NOTHING_KEY
    Set color = m_Colors.Item(p_ColorName)
    Exit Property

NOTHING_KEY:
    Call TheError.Raise(9999, "idp", "This Color name is no use! Color Map Definition: " & p_ColorName)

End Property

'OK
'Private Property Get IdpColorNumber(pColor As Variant) As IdpColorType
''���e:
''   �Y������idpColorType��Ԃ��i���̃N���X�Ɋ���U��ꂽidpColorType���j
''
''���l:
''
'    On Error GoTo NOTHING_KEY
'
'    If pColor = EEE_COLOR_ALL Then
'        IdpColorNumber = idpColorAll
'        Exit Property
'    ElseIf pColor = EEE_COLOR_FLAT Then
'        IdpColorNumber = idpColorFlat
'        Exit Property
'    End If
'
'    IdpColorNumber = m_StrIdps.Item(pColor)
'    Exit Property
'
'NOTHING_KEY:
'    Call TheError.Raise(9999, "idp", "This Color name is no use! Color Map Definition: " & pColor)
'
'End Property

'OK
Private Property Get ColorName(pIdpColorType As IdpColorType) As String
'���e:
'   �Y��idpColorType�̕������Ԃ��i���[�U�[��`String���j
'
'���l:
'
    On Error GoTo NOTHING_KEY
    ColorName = m_IdpStrs.Item(CStr(pIdpColorType))
    Exit Property
NOTHING_KEY:
    Call TheError.Raise(9999, "idp", "This idpColorType is no use! Color Map Definition: " & pIdpColorType)

End Property

'OK
Public Property Get Count() As Integer
'���e:
'   �o�^����Ă���F����Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��19��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Integer
'
'���ӎ���:
'
    Count = m_IdpStrs.Count
End Property

Public Function OddRow() As Collection
'���e:
'   ��s�ɂ���F��S�ĕԂ��܂��B
'   CImgPixel��Collection�ŕԂ��܂��B
'   �ΏہA��Ώۂ͈̔͂ŁA�ǂ���ɂ��Y������ꍇ�A�ΏۂƂ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��4��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Collection
'
'���ӎ���:
'   �J���[�}�b�v��`�ɂ������s�ł��B
'   PMD��`�ɂ������s�Ƃ͈Ӗ����قȂ�ꍇ������܂��B

    Dim p_Color As Variant
    Dim p_Pixel As Variant
    Dim p_ColorCollection As Collection
    
    Set p_ColorCollection = New Collection
    
    For Each p_Color In m_Colors
        For Each p_Pixel In p_Color.Pixels
            If (((p_Pixel.y - 1) Mod 2) + 1) = 1 Then
                Call p_ColorCollection.Add(p_Color)
                Exit For
            End If
        Next p_Pixel
    Next p_Color

    Set OddRow = p_ColorCollection

End Function

Public Function OddColumn() As Collection
'���e:
'   ���ɂ���F��S�ĕԂ��܂��B
'   CImgPixel��Collection�ŕԂ��܂��B
'   �ΏہA��Ώۂ͈̔͂ŁA�ǂ���ɂ��Y������ꍇ�A�ΏۂƂ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��4��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Collection
'
'���ӎ���:
'   �J���[�}�b�v��`�ɂ�������ł��B
'   PMD��`�ɂ�������Ƃ͈Ӗ����قȂ�ꍇ������܂��B

    Dim p_Color As Variant
    Dim p_Pixel As Variant
    Dim p_ColorCollection As Collection
    
    Set p_ColorCollection = New Collection
    
    For Each p_Color In m_Colors
        For Each p_Pixel In p_Color.Pixels
            If (((p_Pixel.x - 1) Mod 2) + 1) = 1 Then
                Call p_ColorCollection.Add(p_Color)
                Exit For
            End If
        Next p_Pixel
    Next p_Color

    Set OddColumn = p_ColorCollection

End Function

Public Function EvenRow() As Collection
'���e:
'   �����s�ɂ���F��S�ĕԂ��܂��B
'   CImgPixel��Collection�ŕԂ��܂��B
'   �ΏہA��Ώۂ͈̔͂ŁA�ǂ���ɂ��Y������ꍇ�A�ΏۂƂ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��4��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Collection
'
'���ӎ���:
'   �J���[�}�b�v��`�ɂ���������s�ł��B
'   PMD��`�ɂ���������s�Ƃ͈Ӗ����قȂ�ꍇ������܂��B

    Dim p_Color As Variant
    Dim p_Pixel As Variant
    Dim p_ColorCollection As Collection
    
    Set p_ColorCollection = New Collection
    
    For Each p_Color In m_Colors
        For Each p_Pixel In p_Color.Pixels
            If (((p_Pixel.y - 1) Mod 2) + 1) = 2 Then
                Call p_ColorCollection.Add(p_Color)
                Exit For
            End If
        Next p_Pixel
    Next p_Color

    Set EvenRow = p_ColorCollection

End Function

Public Function EvenColumn() As Collection
'���e:
'   ������ɂ���F��S�ĕԂ��܂��B
'   CImgPixel��Collection�ŕԂ��܂��B
'   �ΏہA��Ώۂ͈̔͂ŁA�ǂ���ɂ��Y������ꍇ�A�ΏۂƂ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��4��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Collection
'
'���ӎ���:
'   �J���[�}�b�v��`�ɂ����������ł��B
'   PMD��`�ɂ����������Ƃ͈Ӗ����قȂ�ꍇ������܂��B

    Dim p_Color As Variant
    Dim p_Pixel As Variant
    Dim p_ColorCollection As Collection
    
    Set p_ColorCollection = New Collection
    
    For Each p_Color In m_Colors
        For Each p_Pixel In p_Color.Pixels
            If (((p_Pixel.x - 1) Mod 2) + 1) = 2 Then
                Call p_ColorCollection.Add(p_Color)
                Exit For
            End If
        Next p_Pixel
    Next p_Color

    Set EvenColumn = p_ColorCollection

End Function

Private Sub StrToIdpColor(ByRef pMap() As String)
'���e:
'   ���[�U�[����`�����J���[�}�b�v�̃V�[�g����idpColorType�̒�`�ɕϊ����܂��B
'   �J���[�ԍ���CIdpColorMapper�N���X�̏����g���Ċ���U��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��22��
'�p�����[�^:
'   [pMap]   In  1):���[�U�[����`�����}�b�v�̃f�[�^�̈�(Range�I�u�W�F�N�g)
'�߂�l:
'   �Ȃ�
'
'���ӎ���:
'   �ő�W�F�ł�(IG-XL�d�l)

    Dim i As Integer, j As Integer

    '���F�̉������ŊԊu�Ⴂ�G���[�`�F�b�N�p(IG-XL�d�l�AEeeJOB�ŃG���[�o�͂��܂�)
    Dim pHorizontalColorMap(1 To 8) As Integer
    Dim colorCounter As Integer
    colorCounter = 0

    m_Width = UBound(pMap, 1)
    m_Height = UBound(pMap, 2)

    ReDim m_ArrayIdpColorNumber(ORIGIN_X To m_Width, ORIGIN_Y To m_Height)

    m_ArrayColorName = pMap          '���I�z��ɒl�n��

    Set m_StrIdps = New Collection
    Set m_IdpStrs = New Collection

    For i = ORIGIN_Y To m_Height
        
        Erase pHorizontalColorMap
        For j = ORIGIN_X To m_Width
            Dim tmpStr As String
            tmpStr = pMap(j, i)
            Dim tmpIdpColor As Integer
            tmpIdpColor = m_IdpColorMapper.ColorNumber(tmpStr)  'idpColorType�擾

            If Not IsKey(tmpStr, m_StrIdps) Then           'Key(�F��`�̖��O)���쐬����Ă��Ȃ���Γo�^
                colorCounter = colorCounter + 1
                If colorCounter > 8 Then
                    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Name is Too Much!" & vbCrLf & _
                                                                "At most 8 colors" & vbCrLf & _
                                                                "Error Color Map Name: " & m_MapName)
                End If

                pHorizontalColorMap(j) = tmpIdpColor    '�F�̉������Ԋu�`�F�b�N�p

                Call m_StrIdps.Add(tmpIdpColor, tmpStr)
                Call m_IdpStrs.Add(tmpStr, CStr(tmpIdpColor))
            Else
                pHorizontalColorMap(j) = tmpIdpColor    '�F�̉������Ԋu�`�F�b�N�p
            End If
            m_ArrayIdpColorNumber(j, i) = tmpIdpColor
        Next j
        If j < 9 Then
            pHorizontalColorMap(j) = -1            '�F�̉������Ԋu�`�F�b�N�p(�ԕ�)
        End If
        Call CheckIntervalSameColor(pHorizontalColorMap)
    Next i

    'IG-XL PMD Map��񐶐�
    m_IdpColorsPmd = ""
    m_ColorNamesPmd = ""
    
    For i = ORIGIN_Y To m_Height
        m_IdpColorsPmd = m_IdpColorsPmd & "("
        m_ColorNamesPmd = m_ColorNamesPmd & "("
        
        For j = ORIGIN_X To m_Width
            m_IdpColorsPmd = m_IdpColorsPmd & CStr(m_ArrayIdpColorNumber(j, i))
            m_ColorNamesPmd = m_ColorNamesPmd & pMap(j, i)
            
            If j <> m_Width Then
                m_IdpColorsPmd = m_IdpColorsPmd & ","
                m_ColorNamesPmd = m_ColorNamesPmd & ","
            End If
        Next j
        
        m_IdpColorsPmd = m_IdpColorsPmd & ")"
        m_ColorNamesPmd = m_ColorNamesPmd & ")"
    Next i
    
End Sub

'OK
Private Function IsKey(ByVal pKey As String, ByVal pObj As Collection) As Boolean
'���e:
'   �Y��Collection�I�u�W�F�N�g�ɃL�[�����݂��邩���ׂ�
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��22��
'�p�����[�^:
'   [pKey]  In  1):
'   [pObj]  In  2):
'�߂�l:
'   Boolean :True���łɑ��݂��܂�       False�܂�����
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

'OK
Private Function CIdpColor(pInt As Integer) As IdpColorType
    CIdpColor = pInt
End Function

Private Sub CheckIntervalSameColor(pHorizontalColorMap() As Integer)
'���e:
'   ���F�A�������̊Ԋu�`�F�b�N�B
'
'   �z��͉������A�h���X�A�l��IdpColorType
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��22��
'�p�����[�^:
'   [pHorizontalColorMap]  In  1):��������idpColorType���
'�߂�l:
'
'���ӎ���:

Dim i As Integer
Dim pSameColorIndex As Collection
Dim pSameColorAddress(11) As Integer
Dim pSameColorInterval(11) As Integer

    Set pSameColorIndex = New Collection
    Erase pSameColorAddress
    Erase pSameColorInterval
    
    For i = 1 To 8
        If pHorizontalColorMap(i) = -1 Then Exit For       '�ԕ��`�F�b�N

        '�F���Ƃ�0�Ԃ���ԍ�������U��(�Ԋu�`�F�b�N�p)
        On Error Resume Next
        Call pSameColorIndex.Add(pSameColorIndex.Count, CStr(pHorizontalColorMap(i)))
        On Error GoTo 0
        Dim tmpColorIndex As Integer
        tmpColorIndex = pSameColorIndex(CStr(pHorizontalColorMap(i)))
        
        If pSameColorAddress(tmpColorIndex) = 0 Then
            '�Ή��F�ɃA�h���X���
            pSameColorAddress(tmpColorIndex) = i
        Else
            If pSameColorInterval(tmpColorIndex) = 0 Then
                '�Ή��F�ɓ��F�̊Ԋu����
                pSameColorInterval(tmpColorIndex) = i - pSameColorAddress(tmpColorIndex)
                pSameColorAddress(tmpColorIndex) = i
            Else
                '�Ή��F�̊Ԋu�`�F�b�N
                If pSameColorInterval(tmpColorIndex) <> (i - pSameColorAddress(tmpColorIndex)) Then
                    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Map Definition Sheet Error!" & vbCrLf & _
                                                                "Interval of same color is different!" & vbCrLf & vbCrLf & _
                                                                "Color Map Name: " & m_MapName & vbCrLf & _
                                                                "Color Name: " & m_IdpStrs.Item(CStr(pHorizontalColorMap(i))) & vbCrLf & _
                                                                "Error Intervals of number: " & pSameColorInterval(tmpColorIndex) & "," & (i - pSameColorAddress(tmpColorIndex)))
                End If
                pSameColorAddress(tmpColorIndex) = i
            End If
        End If
    Next i

End Sub

''OK
'Private Sub CheckIntervalSameColor(pHorizontalColorMap() As Integer)
''���e:
''   ���F�A�������̊Ԋu�`�F�b�N�B
''
''   �z��͉������A�h���X�A�l��IdpColorType
''�쐬��:
''  tomoyoshi.takase
''�쐬��: 2010�N11��22��
''�p�����[�^:
''   [pHorizontalColorMap]  In  1):��������idpColorType���
''�߂�l:
''
''���ӎ���:
'
'Dim i As Integer
'Dim pSameColorAddress(11) As Integer
'Dim pSameColorInterval(11) As Integer
'
'    Erase pSameColorAddress
'    Erase pSameColorInterval
'
'    For i = 1 To 8
'        If pHorizontalColorMap(i) = -1 Then Exit For       '�ԕ��`�F�b�N
'
'        If pSameColorAddress(pHorizontalColorMap(i)) = 0 Then
'            '�Ή��F�ɃA�h���X���
'            pSameColorAddress(pHorizontalColorMap(i)) = i
'        Else
'            If pSameColorInterval(pHorizontalColorMap(i)) = 0 Then
'                '�Ή��F�ɓ��F�̊Ԋu����
'                pSameColorInterval(pHorizontalColorMap(i)) = i - pSameColorAddress(pHorizontalColorMap(i))
'                pSameColorAddress(pHorizontalColorMap(i)) = i
'            Else
'                '�Ή��F�̊Ԋu�`�F�b�N
'                If pSameColorInterval(pHorizontalColorMap(i)) <> (i - pSameColorAddress(pHorizontalColorMap(i))) Then
'                    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Map Definition Sheet Error!" & vbCrLf & _
'                                                                "Interval of same color is different!" & vbCrLf & vbCrLf & _
'                                                                "Color Map Name: " & m_MapName & vbCrLf & _
'                                                                "Color Name: " & m_IdpStrs.Item(CStr(pHorizontalColorMap(i))) & vbCrLf & _
'                                                                "Error Intervals of number: " & pSameColorInterval(pHorizontalColorMap(i)) & "," & (i - pSameColorAddress(pHorizontalColorMap(i))))
'                End If
'                pSameColorAddress(pHorizontalColorMap(i)) = i
'            End If
'        End If
'    Next i
'
'End Sub


'OK
Private Sub Class_Initialize()
'���e:
'   �R���X�g���N�^
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��26��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'

    Set m_IdpStrs = New Collection
    Set m_StrIdps = New Collection

End Sub

'OK
Private Sub Class_Terminate()
'���e:
'   �f�X�g���N�^
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��26��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'

    Set m_IdpStrs = Nothing
    Set m_StrIdps = Nothing
    Set m_Colors = Nothing
    Set m_IdpColorMapper = Nothing
End Sub

Public Property Get MapName() As String
'���e:
'   ���̃N���X��Color Map Name��߂�
'
'�p�����[�^:
'
'�߂�l:
'   ���̃N���X��Color Map Name��߂�
'
'���ӎ���:
'
    MapName = m_MapName
End Property

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

    With pRes
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_MAPNAME, m_MapName, "String", False, PARAMETER_REPLACER_MAPNAME)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COMMENT, m_Comment, "String", False, PARAMETER_REPLACER_COMMENT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_WIDTH, CStr(m_Width), "Integer", False, PARAMETER_REPLACER_WIDTH)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_HEIGHT, CStr(m_Height), "Integer", False, PARAMETER_REPLACER_HEIGHT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COLORLIST, m_ColorNamesPmd, "String", False, PARAMETER_REPLACER_COLORLIST)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_IDPCOLORSPMD, m_IdpColorsPmd, "String", False, PARAMETER_REPLACER_IDPCOLORSPMD)
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