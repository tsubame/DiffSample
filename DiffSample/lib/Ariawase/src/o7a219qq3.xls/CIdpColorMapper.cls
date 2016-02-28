VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIdpColorMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ���ꂼ��̐F����idpColorType������U��
'
'�ړI:
'   �o�^���ꂽMap�������ɁA�g�p����idpColorType�������Ȃ��}����
'   ����Map���Ŕԍ����Փ˂��Ȃ��悤��idpColorType�̊���U����s��
'
'�쐬��:
'   0145184931
'
Option Explicit

Private Const ERR_NUMBER As Long = 9999                 ' �G���[�ԍ���ێ�����
Private Const CLASS_NAME As String = "CIdpColorMapper"  ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty     ' Default�l"Empty"��ێ�����

Private Const ORIGIN_X As Integer = 1
Private Const ORIGIN_Y As Integer = 1

Private mColorNumbers As Collection         '�F�ԍ��̃��X�g
Private mColorNames As Collection           '�F���̃��X�g
Private mReferenceMapList As Collection     '�e�F�����g�p����Map�̔ԍ��̃��X�g (key1: �F��, key2: Map�ԍ�)
Private mMapNames As Collection             'Map���̃��X�g

Public Sub AddMap(ByRef pMapData() As String, ByVal pName As String)
'���e:
'   Map����ǉ�����
'
'[pMapData()]  IN String�^: Map�f�[�^(�F����2�����z��)���w��
'[pName]       IN String�^: Map�����w��
'
'���l:
'   �d������Map����o�^���邱�Ƃ͂ł��Ȃ�
'
    '### Map����o�^ ###
    'Map�ԍ������蓖�Ă�(���ɓ���Map�����o�^����Ă���ꍇ�̓G���[)
    Call mMapNames.Add(pName, pName)

    '### �F���o�^/�F�����g�p����Map�ԍ��o�^ ###
    'Map����F����S�Ď擾���ď���o�^����
    Dim myWidth As Integer
    myWidth = UBound(pMapData, 1)
    Dim myHeight As Integer
    myHeight = UBound(pMapData, 2)
    Dim mapIndexY As Integer
    For mapIndexY = ORIGIN_Y To myHeight
        Dim mapIndexX As Integer
        For mapIndexX = ORIGIN_X To myWidth
            Dim tmpColorName As String
            '���݂̍��W�̐F�����擾
            tmpColorName = pMapData(mapIndexX, mapIndexY)

            '�F�����o�^�̏ꍇ�̂ݒǉ�����
            If KeyExists(mColorNames, tmpColorName) = False Then
                Call mColorNames.Add(tmpColorName, tmpColorName)
                Call mReferenceMapList.Add(New Collection, tmpColorName)
            End If

            '�F�����g�p����Map����o�^
            On Error Resume Next
            Call mReferenceMapList(tmpColorName).Add(pName, pName)
            On Error GoTo 0
        Next mapIndexX
    Next mapIndexY

    Exit Sub
End Sub

Public Sub Execute()
'���e:
'   �o�^���ꂽMap�������ɂ��Ă��ꂼ��̐F����idpColorType������U��
'
'���l:
'   �K�v��Map����S�ēo�^������Ŏ��s���邱��
'
    Set mColorNumbers = New Collection

    '�F�ԍ����ݒ�̐F�����X�g���쐬
    Dim bufUnregistered As Collection
    Set bufUnregistered = New Collection
    Dim tmpStr As Variant
    For Each tmpStr In mColorNames
        Call bufUnregistered.Add(tmpStr, tmpStr)
    Next tmpStr

    '�F�ԍ� (0���珇�Ɋ���U��)
    Dim colorCounter As Integer
    colorCounter = 0

    '�S�Ă̐F�ɐF�ԍ�������U����܂Ń��[�v
    Do While bufUnregistered.Count > 0
        'Map���L�^�pCollection (�F�ԍ��̏d���`�F�b�N�p)
        Dim tmpMapRegister As Collection
        Set tmpMapRegister = New Collection

        Dim tmpColorName As Variant
        For Each tmpColorName In bufUnregistered
            Dim tmpRefMaps As Collection
            '�F��:tmpColorName���g�p����Map���̃��X�g���擾
            Set tmpRefMaps = mReferenceMapList(tmpColorName)

            '����Map���ŕʁX�̐F�ɑ΂��ē����F�ԍ�������U���Ȃ����m�F
            If IsOverlapping(tmpMapRegister, tmpRefMaps) = False Then
                '�F���ɐF�ԍ���o�^����
                Call mColorNumbers.Add(colorCounter, tmpColorName)
                '�����F�����g�p����Map���L�^����
                Call RegisterMaps(tmpMapRegister, tmpRefMaps)

                '�ԍ����ݒ胊�X�g���珜��
                Call bufUnregistered.Remove(tmpColorName)
            End If
        Next tmpColorName
        colorCounter = colorCounter + 1
    Loop
End Sub

Public Property Get ColorNumber(ByVal pName As String) As Integer
'���e:
'   �F���Ɋ���U��ꂽidpColorType��Ԃ�
'
'[pName]        IN String�^:    idpColorType���擾����F��
'
'���l:
'   Execute���s��Ɏg�p���邱��
'
    ColorNumber = mColorNumbers(pName)
End Property

Private Sub Class_Initialize()
    Set mColorNames = New Collection
    Set mReferenceMapList = New Collection
    Set mMapNames = New Collection
End Sub

Private Sub Class_Terminate()
    Set mColorNumbers = Nothing
    Set mColorNames = Nothing
    Set mReferenceMapList = Nothing
    Set mMapNames = Nothing
End Sub

Private Sub RegisterMaps(ByRef pRegister As Collection, ByRef pMaps As Collection)
'���e:
'   ���X�g��Map�����R���N�V�����L�[�Ƃ��ēo�^���Ă���
'
'[pRegister]    IN/OUT Collection�^:    �L�^�p�R���N�V����
'[pMaps]        IN Collection�^:        �L�^����Map���̃��X�g
'
'���l:
'
    Dim bufKey As Variant
    For Each bufKey In pMaps
        Call pRegister.Add(True, bufKey)
    Next bufKey
End Sub

Private Function IsOverlapping(ByRef pRegister As Collection, ByRef pMaps As Collection) As Boolean
'���e:
'   ���X�g��Map���̒��ɃR���N�V�����L�[�Ƃ��ēo�^�ς݂̂��̂����邩�ǂ�����Ԃ�
'
'[pRegister]    IN Collection�^:    �L�^�m�F��R���N�V����
'[pMaps]        IN Collection�^:    �m�F����Map���̃R���N�V����
'
'���l:
'
    Dim bufKey As Variant
    For Each bufKey In pMaps
        If KeyExists(pRegister, bufKey) = True Then
            IsOverlapping = True
            Exit Function
        End If
    Next bufKey
    IsOverlapping = False
End Function

Private Function KeyExists(ByRef pCollection As Collection, ByVal pKey As String) As Boolean
'���e:
'   �R���N�V�����ɓo�^�ς݂̃L�[���ǂ�����Ԃ�
'
'[pCollection]  IN Collection�^:    �m�F����R���N�V����
'[pKey]         IN String�^:        �m�F����L�[
'
'���l:
'
    On Error GoTo ErrHandler
    Call pCollection(pKey)
    KeyExists = True
    Exit Function
ErrHandler:
    KeyExists = False
End Function