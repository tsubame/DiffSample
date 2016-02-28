VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlanePublicBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �v���[����ۑ����邽�߂̃N���X
'�ړI:
'   ParameterBank�̒���ҁB���ۂɕێ����Ă���̂�ParameterBank
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'

Option Explicit

Implements IImgPlaneBank

Private m_ParameterBank As IParameterBank

Public Function AsIImgPlaneBank() As IImgPlaneBank
    Set AsIImgPlaneBank = Me
End Function

Public Sub Init(pParamBank As IParameterBank)
'ParameterBank�̎Q�Ƃ������Ă����K�v������B
'�������Ɋ����\��
    Set m_ParameterBank = pParamBank
End Sub

Private Sub IImgPlaneBank_Add(ByVal pName As String, ByRef pPlane As CImgPlane, Optional ByVal pFixedPlane As Boolean = False, Optional ByVal pReadOnly As Boolean = False)
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   [pName]         In  1):
'   [pPlane]        In  2):
'   [pFixedPlane]   In  3):
'   [pReadOnly]     In  4):
'�߂�l:
'
'���ӎ���:
'PrameterBank���\�b�h�Ƃ̈Ⴂ
'   �󂯂�^��CImgPlane�^
'IProcParameter�^�̕t���i
'   DeviceInfo  --> Nothing
'   Plane       --> CImgPlane

    Dim pParam As CParamPlane
    Set pParam = New CParamPlane
    Set pParam.plane = pPlane
    
    pParam.plane.ReadOnly = pReadOnly
    If pParam.plane.Comment = "" Then
        pParam.plane.Comment = pName
    End If
    
    Dim pProcParameter As IProcParameter
    Set pProcParameter = pParam         '�L���X�g
    pProcParameter.Name = pName
    
    Call m_ParameterBank.Add(pProcParameter, pName, pFixedPlane)

    Set pParam = Nothing
    Set pProcParameter = Nothing

End Sub

Private Property Get IImgPlaneBank_IsOverwriteMode() As Boolean
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Boolean
'
'���ӎ���:
'PrameterBank���\�b�h�Ƃ̈Ⴂ
'
    IImgPlaneBank_IsOverwriteMode = m_ParameterBank.IsOverwriteMode
End Property

Private Property Let IImgPlaneBank_IsOverwriteMode(ByRef mode As Boolean)
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   [mode]  In/Out  1):
'�߂�l:
'
'���ӎ���:
'PrameterBank���\�b�h�Ƃ̈Ⴂ
'   ParameterBank�S�Ăɉe�����܂��B
    m_ParameterBank.IsOverwriteMode = mode
End Property

Private Function IImgPlaneBank_Item(ByVal pName As String) As CImgPlane
'���e:
'   �L�[�ɊY������CImgPlane Obj��Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   [pName] In/Out  1):
'�߂�l:
'   CImgPlane
'
'���ӎ���:
'PrameterBank���\�b�h�Ƃ̈Ⴂ
'
'�p�����[�^�o���N����󂯎���ĉ摜��������摜�I�u�W�F�N�g��Ԃ��B
    Dim pParameter As IProcParameter
    Set pParameter = m_ParameterBank.Item(pName)
    If TypeName(pParameter) = "CParamPlane" Then
        Dim pParamPlane As CParamPlane
        Set pParamPlane = pParameter
        Set IImgPlaneBank_Item = pParamPlane.plane
        Exit Function
    End If
    
    Set IImgPlaneBank_Item = Nothing

End Function

Private Sub IImgPlaneBank_Delete(ByVal pName As String)
'���e:
'   �L�[�̓��e���摜���ǂ������肵�āA�摜�Ȃ�폜�B�������G���[�B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   [pName] In  1):
'�߂�l:
'
'���ӎ���:
'PrameterBank���\�b�h�Ƃ̈Ⴂ
'   �摜�̂ݍ폜�\�ł��B

    Dim pParameter As IProcParameter
    Set pParameter = m_ParameterBank.Item(pName)
    If TypeName(pParameter) = "CParamPlane" Then
        Call m_ParameterBank.Delete(pName)
        Exit Sub
    Else
        Call TheError.Raise(9999, "CImgPlanePublicBank", "This Key can not delete Plane: " & pName)
    End If

End Sub

'OK
Private Function IImgPlaneBank_isExisting(ByVal pName As String) As Boolean
'���e:
'   ParameterBank�ɓo�^����Ă��邩��Ԃ��܂��B�i�摜�̂݁j
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   [pName] In/Out  1):
'�߂�l:
'   Boolean
'
'���ӎ���:
'PrameterBank���\�b�h�Ƃ̈Ⴂ
'   �o�^����Ă���摜�̂ݑΏۂł��B
    
    On Error GoTo NOTHING_OBJ
    
    If m_ParameterBank.IsExist(pName) = True Then
        If TypeName(m_ParameterBank.Item(pName)) = "CParamPlane" Then
            IImgPlaneBank_isExisting = True
            Exit Function
        End If
    Else
        IImgPlaneBank_isExisting = False
    End If
    
NOTHING_OBJ:
    IImgPlaneBank_isExisting = False

End Function

'OK
Private Sub IImgPlaneBank_Clear()
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'PrameterBank���\�b�h�Ƃ̈Ⴂ
'   ����܂���B
'   ���̃��\�b�h�����s���邱�Ƃ́AParameterBank��Clear�����s���邱�ƂƓ����ł��B
'
    Call m_ParameterBank.Clear
End Sub

'OK
Private Function IImgPlaneBank_List() As String
'���e:
'   �摜�̃��X�g��Ԃ��܂��B
'   ParameterBank�ɓo�^����Ă���摜�̃��X�g
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   String
'
'���ӎ���:
'   PrameterBank���\�b�h�Ƃ̈Ⴂ
'   �摜�݂̂̃��X�g��Ԃ��܂��B
'   �Ԃ��Ă���String�^�̏������኱�Ⴂ�܂��B
'   �EParameterBank����Ⴆ��List���̏���
'   �@�L�[�P(Fixed or Non),�I�u�W�F�N�g�̌^���P,�L�[�Q(Fixed or Non),�I�u�W�F�N�g�̌^���Q,�E�E�E
'   �EPlaneBank��List���̏���
'   �@�o�^���P(Fixed or Non),�摜���P & vbcrlf & �o�^���Q(Fixed or Non),�摜���Q & vbcrlf

    Dim listArr()
    Dim PlaneList As Variant
    Dim pParamPlaneObj As CParamPlane
    Dim pKey As String
        
    If IImgPlaneBank_Count = 0 Then
        IImgPlaneBank_List = ""
        Exit Function
    End If
    
    PlaneList = Split(m_ParameterBank.List, ",")
    
    ReDim listArr(UBound(PlaneList) / 2 - 1, 1)
    
    Dim i As Long
    For i = 0 To UBound(listArr, 1)
        listArr(i, 0) = PlaneList(2 * i + 0)
        listArr(i, 1) = PlaneList(2 * i + 1)
    Next i
    
    For i = 0 To UBound(listArr, 1)
        If listArr(i, 1) = "CParamPlane" Then
            pKey = Replace(listArr(i, 0), "(Fixed)", "")
            Set pParamPlaneObj = m_ParameterBank.Item(pKey)
            IImgPlaneBank_List = IImgPlaneBank_List & listArr(i, 0) & "," & pParamPlaneObj.plane.Name & vbCrLf
        End If
    Next i

End Function

'OK
Private Property Get IImgPlaneBank_Count() As Long
'���e:
'   �o�^�摜�̐���Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Long
'
'���ӎ���:
'   PrameterBank���\�b�h�Ƃ̈Ⴂ
'   �摜�݂̂̃J�E���g��Ԃ��܂��B


    Dim listArr()
    Dim PlaneList As Variant

    PlaneList = Split(m_ParameterBank.List, ",")
    
    '�����Ȃ���ubound=-1,lbound=0
    If UBound(PlaneList) = -1 Then
        IImgPlaneBank_Count = 0
        Exit Property
    End If
    
    Dim i As Long
    Dim j As Long
    For i = 0 To UBound(PlaneList) - 1
        If PlaneList(i) = "CParamPlane" Then
            j = j + 1
        End If
    Next i
    
    IImgPlaneBank_Count = j

End Property

'OK
Private Sub IImgPlaneBank_Display()
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'PrameterBank���\�b�h�Ƃ̈Ⴂ
'   ParameterBank�̃��\�b�h�ɁADisplay���\�b�h�͂���܂���B
    
    Dim disp As New PlaneBankDisplay
    Call disp.Display

End Sub

Private Sub Class_Initialize()
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'
    Set m_ParameterBank = Nothing
End Sub

Private Sub Class_Terminate()
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��9��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'
'���ӎ���:
'
    Set m_ParameterBank = Nothing
End Sub
