VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneMapManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ColorMapInfomation�V�[�g�̏����Ǘ�����N���X
'�ړI:
'   ColorMapInfomation�V�[�g�̑S�����Ǘ��ACImgPlaneMap�ɐU�蕪����B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��21��
'

Option Explicit

Private Const ERR_NUMBER = 9999                         ' Error�ԍ�
Private Const CLASS_NAME = "CImgPlaneMapManager"   ' Class����

Private Const SHT_LABEL_COLOR_MAP_NAME As String = "Color Map Name"
Private Const SHT_LABEL_COLOR_MAP As String = "Color Map"
Private Const SHT_LABEL_COLOR As String = "Color"
Private Const SHT_LABEL_COLOR_ARG_START As Integer = 1
Private Const SHT_LABEL_COLOR_ARG_END As Integer = 8
Private Const SHT_LABEL_COMMENT As String = "Comment"

Private mColorMaps As Collection       '
Private mIdpColorMapper As CIdpColorMapper  'idpColorType���蓖�ăN���X��ێ�

'OK
Public Sub Init()
'���e:
'   ColorMap�V�[�g�f�[�^���擾���āACImgPlaneMap�C���X�^���X�Q�𐶐����܂�
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��21��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'

    '#### Color Map �V�[�g��� ####
    Dim pColorMapName As String        'Color Map Name
    Dim pColorMap As CImgPlaneMap       'Color Map Instance
    Dim pComment As String
    Dim pChkNumColPreLine As Integer
    Dim pChkNumColNowLine As Integer
    Dim pMapData() As String
    Dim pTmpMapData(1 To 8, 1 To 9) As String

    Set mColorMaps = New Collection

    '### Color Map���ꎞ�ۑ�
    Dim pMapNames As Collection
    Set pMapNames = New Collection
    Dim pComments As Collection
    Set pComments = New Collection
    Dim pMapDataList As Collection
    Set pMapDataList = New Collection

    '### idpColorType����U��N���X ###
    Set mIdpColorMapper = New CIdpColorMapper
    
    '#### �s�̃f�[�^�m�F�J�E���^ ####
    Dim intDataCnt As Integer
    Dim intGroupRowCnt As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ERROR_PLANEMAP
    '#####  SheetReader�𗘗p���āAColorMapInfo�V�[�g��ǂݍ���  #####
    Dim clsWrkShtRdr As CWorkSheetReader
    Set clsWrkShtRdr = GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeColorMap)
    
    Dim strSheetName As String
    strSheetName = GetWkShtReaderManagerInstance.GetActiveSheetName(eSheetType.shtTypeColorMap)

    Dim IFileStream As IFileStream
    Dim IParamReader As IParameterReader

    With clsWrkShtRdr
        Set IFileStream = .AsIFileStream
        Set IParamReader = .AsIParameterReader
    End With

    Set clsWrkShtRdr = Nothing

    '#####  ColorMapInfo�V�[�g�̏���ǂݏo�����i�[����  #####
    Do While Not IFileStream.IsEOR
    
        '#####  ColorMapTable ���̂��擾���A�z�����CImgPlaneMap�ɓn��  #####
        If IParamReader.ReadAsString(SHT_LABEL_COLOR_MAP_NAME) <> "" Then
            '#### �O�s�܂ł̃e�[�u�������R���N�V�����Ɋi�[ ####
            If Not (pColorMap Is Nothing) Then
                '#### �O�̍s��Color Map Name�I�[���� ####
                If intGroupRowCnt > 8 Then Call ErrorRowOver8(pColorMapName, intGroupRowCnt)
                
                ReDim pMapData(1 To pChkNumColPreLine, 1 To intGroupRowCnt)
                For i = 1 To UBound(pMapData, 2)
                    For j = 1 To UBound(pMapData, 1)
                        pMapData(j, i) = pTmpMapData(j, i)
                    Next j
                Next i

                'Map�����ꎞ�I�ɕێ�����
                Call pMapDataList.Add(pMapData, pColorMapName)
                Call pMapNames.Add(pColorMapName)
                Call pComments.Add(pComment, pColorMapName)

                'idpColorType����U��N���X��Map�f�[�^�ǉ�
                Call mIdpColorMapper.AddMap(pMapData, pColorMapName)

'                Call pColorMap.Create(pColorMapName, pMapData, pComment)
                Call mColorMaps.Add(pColorMap, pColorMapName)
                
                Erase pTmpMapData
                Set pColorMap = Nothing
            End If
            
            '##### ���݂̍s�̃e�[�u�����擾 #####
            intGroupRowCnt = 0
            pChkNumColPreLine = 0
            pChkNumColNowLine = 0
            pColorMapName = ""
            
            Set pColorMap = New CImgPlaneMap
            pColorMapName = IParamReader.ReadAsString(SHT_LABEL_COLOR_MAP_NAME)
            pComment = IParamReader.ReadAsString(SHT_LABEL_COMMENT)
            If IsKey(pColorMapName, mColorMaps) Then       'Color Map Name�d���`�F�b�N
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "This Color Map Name Already is Exist! " & pColorMapName)
            End If
        End If
        
        If intGroupRowCnt > 8 Then Call ErrorRowOver8(pColorMapName, intGroupRowCnt)
        
        Dim pStr As String
        For i = SHT_LABEL_COLOR_ARG_START To SHT_LABEL_COLOR_ARG_END
            pStr = IParamReader.ReadAsString(SHT_LABEL_COLOR & i & "@" & SHT_LABEL_COLOR_MAP)
            If pStr <> "" Then
                pTmpMapData(pChkNumColNowLine + 1, intGroupRowCnt + 1) = pStr
                pChkNumColNowLine = pChkNumColNowLine + 1
            End If
            
            '#### �J�����̐��̃`�F�b�N ####
            If pStr = "" Or i = SHT_LABEL_COLOR_ARG_END Then
            
                '  �J�����̐����O�̍s�ƈႤ               And �P�s�ڂłȂ�
                If pChkNumColPreLine <> pChkNumColNowLine And pChkNumColPreLine <> 0 Then
                    Call ErrorDifferentColumnNum(pColorMapName, pChkNumColPreLine, pChkNumColNowLine)    '�����̏��̐��ɈႢ������B
                End If
                pChkNumColPreLine = pChkNumColNowLine
                pChkNumColNowLine = 0
                Exit For
            End If
        Next i
        
        intDataCnt = intDataCnt + 1
        intGroupRowCnt = intGroupRowCnt + 1
        IFileStream.MoveNext
    Loop

    '##### �Ō�̃J���[�}�b�v��`�̏I�[���� #####
    If Not (pColorMap Is Nothing) Then
        If intGroupRowCnt > 8 Then Call ErrorRowOver8(pColorMapName, intGroupRowCnt)
        
        ReDim pMapData(1 To pChkNumColPreLine, 1 To intGroupRowCnt)
        For i = 1 To UBound(pMapData, 2)
            For j = 1 To UBound(pMapData, 1)
                pMapData(j, i) = pTmpMapData(j, i)
            Next j
        Next i

        'Map�����ꎞ�I�ɕێ�����
        Call pMapDataList.Add(pMapData, pColorMapName)
        Call pMapNames.Add(pColorMapName)
        Call pComments.Add(pComment, pColorMapName)

        'idpColorType����U��N���X��Map�f�[�^�ǉ�
        Call mIdpColorMapper.AddMap(pMapData, pColorMapName)

'        Call pColorMap.Create(pColorMapName, pMapData, pComment)
        Call mColorMaps.Add(pColorMap, pColorMapName)
        
        Erase pTmpMapData
        Set pColorMap = Nothing
    End If
    
    Set IFileStream = Nothing
    Set IParamReader = Nothing

    '##### idpColorType����U��v�Z���s #####
    Call mIdpColorMapper.Execute

    '##### �e�[�u�������C���X�^���X�Ɋi�[ #####
    Dim mapIndex As Long
    For mapIndex = 1 To pMapDataList.Count
        Set pColorMap = mColorMaps(mapIndex)
        pMapData = pMapDataList(mapIndex)
        Call pColorMap.Create(pMapNames(mapIndex), pMapData, mIdpColorMapper, pComments(mapIndex))
    Next mapIndex
    Exit Sub
    
ERROR_PLANEMAP:
    Set mColorMaps = New Collection
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

Private Sub Class_Initialize()
'���e:
'   �R���X�g���N�^
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��25��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'

    Set mColorMaps = New Collection
'    Call Initialize
End Sub

Private Sub Class_Terminate()
'���e:
'   �f�X�g���N�^
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��25��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'

    Set mColorMaps = Nothing
    Set mIdpColorMapper = Nothing
End Sub

'OK
Public Function PlaneMapList() As String()
'���e:
'   �o�^�ς݂�Color Map ���𓮓I�z���String�ŕԂ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��25��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   String
'
'���ӎ���:
'

    Dim pTmp As Variant
    Dim pArray() As String
    Dim PNum As Long

    ReDim pArray(mColorMaps.Count - 1)
    
    For Each pTmp In mColorMaps
        pArray(PNum) = CStr(pTmp.Name)
        PNum = PNum + 1
    Next pTmp

    PlaneMapList = pArray

End Function

'OK
Public Function planeMap(pColorMapName As String) As CImgPlaneMap
'���e:
'   �o�^�ς݂�Color Map��Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��25��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   String
'
'���ӎ���:
'
    On Error GoTo NOTHING_COLOR_MAP
    Set planeMap = mColorMaps.Item(pColorMapName)
    On Error GoTo 0
    Exit Function
    
NOTHING_COLOR_MAP:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Map Name is Nothing!: " & pColorMapName & vbCrLf)
End Function

'OK
Public Property Get Count() As Long
'���e:
'   �o�^�ς݂�Color Map�̐���Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N10��25��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Long
'
'���ӎ���:
'
    Count = mColorMaps.Count
End Property

Public Function IsExist(pColorMapName As String) As Boolean
'���e:
'   �Ώۂ�ColorMap�̗L����Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��30��
'�p�����[�^:
'   [pColorMapName]  In  1):
'�߂�l:
'   Boolean
'
'���ӎ���:
'
    On Error GoTo NOT_FOUND
    Call mColorMaps.Item(pColorMapName)
    IsExist = True
    Exit Function
NOT_FOUND:
    IsExist = False
End Function

Private Sub ErrorRowOver8(pTableName As String, pOverNum As Integer)
'���e:
'   ColorMap�̒�`��
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��11��
'�p�����[�^:
'   [pTableName]    In  1):
'   [pOverNum]      In  2):
'�߂�l:
'
'���ӎ���:

    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Map Row Data is Too Much!" & vbCrLf _
                                                & "Row Data Max 8,But Written: " & pOverNum & vbCrLf _
                                                & "Error Color Map Name: " & pTableName)
End Sub

Private Sub ErrorDifferentColumnNum(pTableName As String, pPreNum As Integer, pNowNum As Integer)
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��11��
'�p�����[�^:
'   [pTableName]    In  1):
'   [pPreNum]       In  2):
'   [pNowNum]       In  3):
'�߂�l:
'
'���ӎ���:

    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Map Name is differnet Collummn Number! " & pTableName & vbCrLf & _
                                                "Number Of Column: " & pPreNum & " and " & pNowNum)
End Sub

Private Function IsKey(ByVal pKey As String, ByVal pObj As Collection) As Boolean
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


