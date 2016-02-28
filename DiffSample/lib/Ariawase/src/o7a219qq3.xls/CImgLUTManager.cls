VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgLUTManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   LUT���Ǘ�
'�ړI:
'   LUT�̏��Ǘ��p
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��30��

Option Explicit

Private Const ERR_NUMBER = 9999                         ' Error�ԍ�
Private Const CLASS_NAME = "CImgLUTManager"             ' Class����

Private m_LutInfos As Collection

Public Sub CreateIndexLUT( _
    ByVal pLUTName As String, ByVal pInVal1 As Long, ByVal pInVal2 As Long, _
    ByVal pOutVal1 As Long, ByVal pOutVal2 As Long, Optional ByVal pValidBit As Integer = 12 _
)
'���e:
'   LUT�̍쐬
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��30��
'�p�����[�^:
'   [pLUTName]  In  1):LUT��`��
'   [pInVal1]   In  2):InputValue1
'   [pInVal2]   In  3):InputValue2
'   [pOutVal1]  In  4):OutputValue1
'   [pOutVal2]  In  5):OutputValue2
'   [pValidBit] In  6):Bit
'�߂�l:
'
'���ӎ���:
'
    Dim pLut As CImgLUTInfo
    
    'IG-XL LUT�V�[�g�G���[�d�l�Ɠ����̃G���[����
    If pOutVal1 < -32768 Or pOutVal1 > 32767 Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CreateIndexLUT", _
                            pLUTName & "'s OutputValue is Too big or Too Small")
    End If
    If pOutVal2 < -32768 Or pOutVal2 > 32767 Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CreateIndexLUT", _
                            pLUTName & "'s OutputValue is Too big or Too Small")
    End If
    
    If IsExist(pLUTName) Then
        Set pLut = m_LutInfos.Item(pLUTName)
        If pLut.Bit <> pValidBit Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CreateIndexLUT", _
                                pLUTName & "'s LUT bit is different.")
        End If
    Else
        Set pLut = New CImgLUTInfo
        Call pLut.Initialize(pLUTName, pValidBit)
        Call m_LutInfos.Add(pLut, pLUTName)
    End If
    Call pLut.Add(pLUTName, pInVal1, pInVal2, pOutVal1, pOutVal2)
    Set pLut = Nothing

    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.CreateIndexLUT(pLUTName, pInVal1, pInVal2, pOutVal1, pOutVal2, 0, pValidBit)
    
    Exit Sub
    
ERROR_IGXL:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CreateIndexLUT", "CreateIndexLUT")
    
End Sub

'IG-XL�̎d�l�ɍ��킹�āALUT��`�̏㏑�����\�Ƃ��邽�ߔp�~�B
'�\�[�X�R�[�h��ɎU�݂���LUT���́A���[�U�[���Ǘ����Ȃ��Ƃ����Ȃ��B
'Public Sub TerminateLUT()
''���e:
''   LUT��`�̏I�[����
''   LUT�̒�`���I���������Ƃ𖾎����āA��`�I����̓����̍Ē�`�͋����܂���B
''�쐬��:
''  tomoyoshi.takase
''�쐬��: 2010�N11��30��
''�p�����[�^:
''   �Ȃ�
''�߂�l:
''
''���ӎ���:
''
''�T���v���R�[�h
''    Call TheIDP.CreateIndexLUT("ABS", -1024, 0, 1024, 0, 11)
''    Call TheIDP.CreateIndexLUT("ABS", 0, 1024, 0, 1024, 11)
''    Call TheIDP.TerminateLUT
''
''2010/12/03 ���\�b�h���s��CImgLUTManager�Ɉړ�
'
'    If m_tmpLutInfo.Count = 0 Then Exit Sub
'
'    Dim pLut As CImgLUTInfo
'    Set pLut = New CImgLUTInfo
'    Call pLut.Initialize(m_CurrentLutName, m_LatestBitNumber, m_tmpLutInfo)
'
'    On Error GoTo ALREADY_ADD
'    Call m_LutInfos.Add(pLut, m_CurrentLutName)
'
'    Set m_tmpLutInfo = Nothing
'    Set m_tmpLutInfo = New Collection
'
'    Set pLut = Nothing
'
'    Exit Sub
'ALREADY_ADD:
'    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".TerminateLUT", "")
'
'End Sub

Public Function IsExist(pLUTName As String) As Boolean
'���e:
'   �Ώۂ�LUT�̗L����Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��30��
'�p�����[�^:
'   [pLutName]  In  1):
'�߂�l:
'   Boolean
'
'���ӎ���:
'
    On Error GoTo NOT_FOUND
    Call m_LutInfos.Item(pLUTName)
    IsExist = True
    Exit Function
NOT_FOUND:
    IsExist = False
End Function

Public Function LUT(pLUTName As String) As CImgLUTInfo
'���e:
'   �Ώۂ�LUT�̃C���X�^���X��Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N12��2��
'�p�����[�^:
'   [pLUTName]  In  1):LUT��
'�߂�l:
'   CImgLUTInfo
'
'���ӎ���:
'

    On Error GoTo NOT_FOUND
    Set LUT = m_LutInfos.Item(pLUTName)
    Exit Function
NOT_FOUND:
    Call TheError.Raise(9999, "", "This LUT Name is not found!" & vbCrLf & "LUT Name:" & pLUTName)
End Function

'Public Property Get LutCollection() As Collection
''���e:
''   ���̃N���X������Lut��Collection��Ԃ��܂��B
''�쐬��:
''  tomoyoshi.takase
''�쐬��: 2010�N12��2��
''�p�����[�^:
''   �Ȃ�
''�߂�l:
''   Collection
''
''���ӎ���:
'
'    Set LutCollection = m_LutInfos
'
'End Property

Public Property Get Count() As Long
'���e:
'   ���̃N���X������Lut�̐���Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N12��2��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Long
'
'���ӎ���:

    Count = m_LutInfos.Count

End Property

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
    Set m_LutInfos = New Collection

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
    Set m_LutInfos = Nothing

End Sub