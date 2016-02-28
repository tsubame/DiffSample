VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �v���[����ۑ����邽�߂̃N���X
'
'�ړI:
'   �O���[�o���ϐ��̎g�p������邽�߁A�ۑ��p�̃N���X��p�ӁB
'   ���O������CImgPlane���֘A�t����B
'
'�쐬��:
'   0145184004
'
'2011/02/09
'ParameterBank�̎g�p�ɔ������̃N���X�͎g�p���܂���B
'TheIDP����PlaneBank���g�p����Ƃ��́ACImgPlanePublicBank

Option Explicit

Implements IImgPlaneBank

Private m_Names As Collection
Private m_Planes As Collection

Private m_FixedNames As Collection
Private m_FixedPlanes As Collection
Private m_OverWriteMode As Boolean


Private Sub IImgPlaneBank_Add(ByVal pName As String, ByRef pPlane As CImgPlane, Optional ByVal pFixedPlane As Boolean = False, Optional ByVal pReadOnly As Boolean = False)
'���e:
'   �v���[����o�^����B
'   pFixedPlane��True�ɂ����Clear�ňꊇ�������Ȃ��Ȃ�B
'   pReadOnly��True�ɂ���ƁA�摜��ύX����֐�(��:WritePixel, MultiMean, etc)���g���Ȃ��Ȃ�B
'    �����͂ɂ͎g�����Ƃ͉\�BAverage�Ȃǃf�[�^�̎擾���\�B
'
'[pName]       IN   String�^:       ���ʖ�
'[pPlane]      IN   CImgPlane�^:    �o�^����v���[��
'[pFixedPlane] IN   Boolean�^:      �Œ�v���[���I�v�V����
'[pReadOnly]   IN   Boolean�^:      ���[�h�I�����[�I�v�V����
'
'���l:
'   �G���[����
'   1,�w�肵�����O�����ɑ��݂��Ă��鎞
'
'   �T���v���R�[�h
'    Dim src As CImgPlane
'    Set src = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call TheIDP.PlaneBank.Add("SOURCE", src)
'
    
    '�G���[
''    If IImgPlaneBank_isExisting(pName) = True Then
''        Call TheError.Raise(999, "Add", ErrMsgNameAlreadyExists(pName))
''        Exit Sub
''    End If
    If IImgPlaneBank_IsOverwriteMode = False Then
        If IImgPlaneBank_isExisting(pName) = True Then
            Call TheError.Raise(999, "Add", ErrMsgNameAlreadyExists(pName))
            Exit Sub
        End If
    Else
        'Bank�o�^�ɏ㏑���BTOPT Acquire��蒼���p
        If isExisting_(m_FixedNames, pName) = True Then
            Call m_FixedNames.Remove(pName)
            Call m_FixedPlanes.Remove(pName)
        ElseIf isExisting_(m_Names, pName) = True Then
            Call m_Names.Remove(pName)
            Call m_Planes.Remove(pName)
        End If
        '�o�^������Ή������Ȃ�
    End If

    
    pPlane.ReadOnly = pReadOnly
    If pPlane.Comment = "" Then
        pPlane.Comment = pName
    End If
    
    If pFixedPlane = True Then
        Call m_FixedNames.Add(pName, pName)
        Call m_FixedPlanes.Add(pPlane, pName)
    Else
        Call m_Names.Add(pName, pName)
        Call m_Planes.Add(pPlane, pName)
    End If
    
End Sub

Private Property Get IImgPlaneBank_IsOverwriteMode() As Boolean
    IImgPlaneBank_IsOverwriteMode = m_OverWriteMode
End Property

Private Property Let IImgPlaneBank_IsOverwriteMode(ByRef mode As Boolean)
    m_OverWriteMode = mode
End Property

Private Function IImgPlaneBank_Item(ByVal pName As String) As CImgPlane
'���e:
'   �w�肵�����O�̃v���[�����擾����
'
'[pName]        IN      String�^:       ���ʖ�
'
'���l:
'   �G���[����
'   1,�w�肵�����O�����݂��Ȃ���
'
'   �T���v���R�[�h
'    Dim src As CImgPlane
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'
    Set IImgPlaneBank_Item = GetPlane(pName, m_Planes)
    If IImgPlaneBank_Item Is Nothing Then
        Set IImgPlaneBank_Item = GetPlane(pName, m_FixedPlanes)
    End If

    '�G���[
    If IImgPlaneBank_Item Is Nothing Then
        Call TheError.Raise(999, "Item", ErrMsgNameDoesntExist(pName))
    End If
End Function

Private Sub IImgPlaneBank_Delete(ByVal pName As String)
'���e:
'   �w�肵�����O�̃v���[���̓o�^��j������B
'
'[pName]       IN/OUT String�^:     ���ʖ�
'
'���l:
'   �G���[����
'   1,�w�肵�����O�����݂��Ȃ���
'
'   �T���v���R�[�h
'    Call TheIDP.PlaneBank.Delete("SOURCE")
'
    If Not GetPlane(pName, m_Planes) Is Nothing Then
        Call m_Names.Remove(pName)
        Call m_Planes.Remove(pName)
    ElseIf Not GetPlane(pName, m_FixedPlanes) Is Nothing Then
        Call m_FixedNames.Remove(pName)
        Call m_FixedPlanes.Remove(pName)
    Else
        '�G���[
        Call TheError.Raise(999, "Delete", ErrMsgNameDoesntExist(pName))
    End If

End Sub

Private Function IImgPlaneBank_isExisting(ByVal pName As String) As Boolean
'���e:
'   �w�肵�����O�̓o�^�����邩�ǂ����̊m�F�B
'
'[pName]       IN/OUT String�^:     ���ʖ�
'
'���l:
'   �T���v���R�[�h
'    Debug.Print TheIDP.PlaneBank.isExisting("SOURCE")
'

    If (Not isExisting_(m_Names, pName)) And (Not isExisting_(m_FixedNames, pName)) Then
        IImgPlaneBank_isExisting = False
    Else
        IImgPlaneBank_isExisting = True
    End If
    
End Function

Private Function isExisting_(ByRef pPlanes As Collection, ByVal pName As String) As Boolean

    Dim tmp As String
    On Error GoTo NOT_FOUND
    tmp = pPlanes.Item(pName)
    
    isExisting_ = True
    Exit Function
    
NOT_FOUND:
    isExisting_ = False

End Function

Private Sub IImgPlaneBank_Clear()
'���e:
'   �Œ�v���[���ȊO�̓o�^���ꂽ�v���[�������ׂĉ������B
'
'���l:
'   �T���v���R�[�h
'    Call TheIDP.PlaneBank.Clear
'
    Set m_Planes = New Collection
    Set m_Names = New Collection
End Sub

Private Function IImgPlaneBank_List() As String
'���e:
'   �o�^���ꂽ�v���[���̃��X�g���擾
'
'���l:
'   �J���}��؂�̕�����ŏo�́B
'   �o�^��,�v���[����,�c�@�̏�
'
'   �T���v���R�[�h
'    Debug.Print TheIDP.PlaneBank.List
'
    Dim tmp As Variant
    For Each tmp In m_FixedNames
        IImgPlaneBank_List = IImgPlaneBank_List & tmp & "(Fixed)" & "," & m_FixedPlanes(tmp).Name & vbCrLf
    Next tmp
    
    For Each tmp In m_Names
        IImgPlaneBank_List = IImgPlaneBank_List & tmp & "," & m_Planes(tmp).Name & vbCrLf
    Next tmp
    
End Function

Private Property Get IImgPlaneBank_Count() As Long
'���e:
'   �o�^���ꂽ�v���[���̐����擾
'
'���l:
'   �T���v���R�[�h
'    Debug.Print TheIDP.PlaneBank.Count
'
    IImgPlaneBank_Count = m_Planes.Count + m_FixedPlanes.Count
End Property

Private Sub IImgPlaneBank_Display()
'���e:
'   �o�^���ꂽ�v���[���̃��X�g��\��
'
'���l:
'   �T���v���R�[�h
'    Call TheIDP.PlaneBank.Display
'
    Dim disp As New PlaneBankDisplay
    Call disp.Display
End Sub

Private Function GetPlane(ByVal pName As String, ByRef pPlanes As Collection) As CImgPlane

On Error GoTo NOT_FOUND
    Set GetPlane = pPlanes(pName)
    
    Exit Function
    
NOT_FOUND:
    Set GetPlane = Nothing
    
End Function

Private Sub Class_Initialize()

    Set m_Names = New Collection
    Set m_Planes = New Collection
    Set m_FixedNames = New Collection
    Set m_FixedPlanes = New Collection
    
End Sub

Private Sub Class_Terminate()

    Set m_Names = Nothing
    Set m_Planes = Nothing
    Set m_FixedNames = Nothing
    Set m_FixedPlanes = Nothing
    
End Sub

Private Function ErrMsgNameDoesntExist(ByVal pName As String)
    ErrMsgNameDoesntExist = """" & pName & """ doesn't exist."
End Function

Private Function ErrMsgNameAlreadyExists(ByVal pName As String)
    ErrMsgNameAlreadyExists = """" & pName & """ already exists."
End Function

Public Function AsIImgPlaneBank() As IImgPlaneBank
    Set AsIImgPlaneBank = Me
End Function

