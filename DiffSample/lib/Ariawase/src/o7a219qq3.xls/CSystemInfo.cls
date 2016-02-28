VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �V�X�e���Ɋւ����̃Z�b�g��ێ�����N���X
'
'�ړI:
'   �e�X�^�[ID�AWaferID�ȂǃV�X�e���Ɋւ����̃Z�b�g��ێ�����B
'
'�쐬��:
'   0145184306
'

Option Explicit

Private Const ERR_NUMBER = 9999         ' �G���[�ԍ���ێ�����
Private Const CLASS_NAME = "CSystemInfo" ' Class���̂�ێ�����
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default�l"Empty"��ێ�����

Private m_JobName As String
Private m_LotNo As Long
Private m_MapName As String
Private m_NodeNo As Long
Private m_ProductName As String
Private m_SystemName As String
Private m_WaferNo As Long
Private m_ChipNo() As Long
Private m_XAddress() As Long
Private m_YAddress() As Long

Public Property Get chipNo() As Long()
'���e:
'   �ێ������p�����[�^��Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^(Long()�^)
'
'���l:
'
'
    chipNo = m_ChipNo
End Property

Public Property Let chipNo(pChipNo() As Long)
'���e:
'   �p�����[�^��ێ�����
'
'[pChipNo()]    IN Long�^:     �p�����[�^
'
'���l:
'
'
    m_ChipNo = pChipNo
End Property

Public Property Get JobName() As String
'���e:
'   �ێ������p�����[�^��Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^(String�^)
'
'���l:
'
'
    JobName = m_JobName
End Property

Public Property Let JobName(pJobName As String)
'���e:
'   �p�����[�^��ێ�����
'
'[pJobName]    IN String�^:     �p�����[�^
'
'���l:
'
'
    m_JobName = pJobName
End Property

Public Property Get LotNo() As Long
'���e:
'   �ێ������p�����[�^��Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^(Long�^)
'
'���l:
'
'
    LotNo = m_LotNo
End Property

Public Property Let LotNo(pLotNo As Long)
'���e:
'   �p�����[�^��ێ�����
'
'[pLotNo]    IN Long�^:     �p�����[�^
'
'���l:
'
'
    m_LotNo = pLotNo
End Property

Public Property Get NodeNo() As Long
'���e:
'   �ێ������p�����[�^��Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^(Long�^)
'
'���l:
'
'
    NodeNo = m_NodeNo
End Property

Public Property Let NodeNo(pNodeNo As Long)
'���e:
'   �p�����[�^��ێ�����
'
'[pNodeNo]    IN Long�^:     �p�����[�^
'
'���l:
'
'
    m_NodeNo = pNodeNo
End Property

Public Property Get ProductName() As String
'���e:
'   �ێ������p�����[�^��Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^(String�^)
'
'���l:
'
'
    ProductName = m_ProductName
End Property

Public Property Let ProductName(pProductName As String)
'���e:
'   �p�����[�^��ێ�����
'
'[pProductName]    IN String�^:     �p�����[�^
'
'���l:
'
'
    m_ProductName = pProductName
End Property

Public Property Get SystemName() As String
'���e:
'   �ێ������p�����[�^��Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^(String�^)
'
'���l:
'
'
    SystemName = m_SystemName
End Property

Public Property Let SystemName(pSystemName As String)
'���e:
'   �p�����[�^��ێ�����
'
'[pSystemName]    IN String�^:     �p�����[�^
'
'���l:
'
'
    m_SystemName = pSystemName
End Property

Public Property Get WaferMapName() As String
'���e:
'   �ێ������p�����[�^��Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^(Stirng�^)
'
'���l:
'
'
    WaferMapName = m_MapName
End Property

Public Property Let WaferMapName(pMapName As String)
'���e:
'   �p�����[�^��ێ�����
'
'[pMapName]    IN String�^:     �p�����[�^
'
'���l:
'
'
    m_MapName = pMapName
End Property

Public Property Get WaferNo() As Long
'���e:
'   �ێ������p�����[�^��Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^(Long�^)
'
'���l:
'
'
    WaferNo = m_WaferNo
End Property

Public Property Let WaferNo(pWaferNo As Long)
'���e:
'   �p�����[�^��ێ�����
'
'[pWaferNo]    IN Long�^:     �p�����[�^
'
'���l:
'
'
    m_WaferNo = pWaferNo
End Property

Public Property Get xAddress() As Long()
'���e:
'   �ێ������p�����[�^��Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^(Long()�^)
'
'���l:
'
'
    xAddress = m_XAddress
End Property

Public Property Let xAddress(pXAddress() As Long)
'���e:
'   �p�����[�^��ێ�����
'
'[pXAddress()]    IN Long�^:     �p�����[�^
'
'���l:
'
'
    m_XAddress = pXAddress
End Property

Public Property Get yAddress() As Long()
'���e:
'   �ێ������p�����[�^��Ԃ�
'
'�߂�l:
'   �ێ������p�����[�^(Long()�^)
'
'���l:
'
'
    yAddress = m_YAddress
End Property

Public Property Let yAddress(pYAddress() As Long)
'���e:
'   �p�����[�^��ێ�����
'
'[pYAddress()]    IN Long�^:     �p�����[�^
'
'���l:
'
'
    m_YAddress = pYAddress
End Property

Public Function GetChipNoOfSite(pSite As Long) As Long
'���e:
'   �w��T�C�g��X�A�h���X��Ԃ�
'
'����:
'[pSite]       IN Long�^:       �T�C�g�ԍ�
'
'�߂�l:
'   �w��T�C�g��ChipNo(Long�^)
'
'���l:
'
'
    On Error GoTo ErrHundler
    GetChipNoOfSite = m_ChipNo(pSite)
    Exit Function
ErrHundler:
    GetChipNoOfSite = -1
End Function

Public Function GetXAddressOfSite(pSite As Long) As Long
'���e:
'   �w��T�C�g��X�A�h���X��Ԃ�
'
'����:
'[pSite]       IN Long�^:       �T�C�g�ԍ�
'
'�߂�l:
'   �w��T�C�g��X�A�h���X(Long�^)
'
'���l:
'
'
    On Error GoTo ErrHundler
    GetXAddressOfSite = m_XAddress(pSite)
    Exit Function
ErrHundler:
    GetXAddressOfSite = -1
End Function

Public Function GetYAddressOfSite(pSite As Long) As Long
'���e:
'   �w��T�C�g��Y�A�h���X��Ԃ�
'
'����:
'[pSite]       IN Long�^:       �T�C�g�ԍ�
'
'�߂�l:
'   �w��T�C�g��Y�A�h���X(Long�^)
'
'���l:
'
'
    On Error GoTo ErrHundler
    GetYAddressOfSite = m_YAddress(pSite)
    Exit Function
ErrHundler:
    GetYAddressOfSite = -1
End Function

Private Sub Class_Initialize()
'���e:
'   �R���X�g���N�^
'
'���l:
'
'
    m_JobName = INITIAL_EMPTY_VALUE
    m_LotNo = 0
    m_MapName = INITIAL_EMPTY_VALUE
    m_NodeNo = 0
    m_ProductName = INITIAL_EMPTY_VALUE
    m_SystemName = INITIAL_EMPTY_VALUE
    m_WaferNo = 0
    Erase m_ChipNo(), m_XAddress(), m_YAddress()
End Sub

Private Sub Class_Terminate()
'���e:
'   �f�X�g���N�^
'
'���l:
'
'
    m_JobName = INITIAL_EMPTY_VALUE
    m_LotNo = 0
    m_MapName = INITIAL_EMPTY_VALUE
    m_NodeNo = 0
    m_ProductName = INITIAL_EMPTY_VALUE
    m_SystemName = INITIAL_EMPTY_VALUE
    m_WaferNo = 0
    Erase m_ChipNo(), m_XAddress(), m_YAddress()
End Sub