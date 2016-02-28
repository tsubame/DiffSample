VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPixel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ��f�ЂƂɑ���������̃N���X
'�ړI:
'   ��f�̃A�h���X����Ԃ��܂��BOrigin�͂P
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N1��31��

Option Explicit

Implements IProcParamReplaceable

Private Const PARAMETER_REPLACER_X As String = "x"
Private Const PARAMETER_REPLACER_Y As String = "y"

Private mX As Long
Private mY As Long

Public Sub Init(pX As Long, pY As Long)
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��3��
'�p�����[�^:
'   [pX]    In/Out  1):
'   [pY]    In/Out  2):
'�߂�l:
'
'���ӎ���:
'
'

    
    If mX > 0 Or mY > 0 Then
        Call TheError.Raise(9999, "", "CImgPixel is set already!")
    End If
    
    mX = pX
    mY = pY

End Sub

Public Function Clone() As CImgPixel
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��3��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   CImgPixel
'
'���ӎ���:
'
'

    Dim p_Pixel As CImgPixel
    
    Set p_Pixel = New CImgPixel
    Call p_Pixel.Init(mX, mY)
    Set Clone = p_Pixel

End Function

Public Property Get x() As Long
    x = mX
End Property

Public Property Get y() As Long
    y = mY
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
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_X, CStr(mX), "Long", False, PARAMETER_REPLACER_X)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_Y, CStr(mY), "Long", False, PARAMETER_REPLACER_Y)
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