VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_TestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' �N���X��  �F  csLoopTestInfo.cls
' �T�v      �F �e�e�X�g���ڂ̃p�����[�^�̊i�[�A�������s���N���X
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/10�@�V�K�쐬        K.Sumiyashiki
'==========================================================================

'�t���[�e�[�u���f�[�^
Public tname     As String
Public Parameter As String
Public Bin       As Integer
Public OpCode    As String
Public Label     As String

'�e�X�g�C���X�^���X�f�[�^
Public HighLimit As Double
Public LowLimit  As Double
Public Unit      As String
Public arg2      As Integer

'LOOP�����p�f�[�^
Public CapCategory1   As String
Public CapCategory2   As String
Public LoopJudgeLimit As Double

'���ʒ����p�f�[�^
Public OptTarget     As Double
Public OptIdentifier As String
Public OptJudgeLimit As Double

'Wait�����p�f�[�^
Public WaitAdjFlg    As String

Private m_xSite(nSite) As csPALS_TestInfoParams     '�e�T�C�g�̓����l�����i�[����ׂ̃N���X��`


'********************************************************************************************
' ���O: init_csPALS_TestInfo
' ���e: �e�ϐ��̏�����
'�@�@�@ m_xSite�̓T�C�g�����̃C���X�^���X�𐶐�
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub init_csPALS_TestInfo()

On Error GoTo errPALSinit_csPALS_TestInfo

    tname = vbNullString
    Parameter = vbNullString
    Bin = -1
    OpCode = vbNullString
    Label = vbNullString
    
    HighLimit = -1
    LowLimit = -1
    Unit = vbNullString
    
    OptTarget = -1
    OptIdentifier = vbNullString
    OptJudgeLimit = -1

    CapCategory1 = vbNullString
    CapCategory2 = vbNullString
    LoopJudgeLimit = -1

    arg2 = -1

    Dim sitez As Long       '���[�v�J�E���^
    '�T�C�g�����̃C���X�^���X�𐶐�
    For sitez = 0 To nSite
        Set m_xSite(sitez) = New csPALS_TestInfoParams
    Next sitez

Exit Sub

errPALSinit_csPALS_TestInfo:
    Call sub_errPALS("csPALS_TestInfo not initialize at 'init_csPALS_TestInfo'", "0-6-01-0-22")

End Sub


'*********************************************************************************
' �֐���    �F  Class_Initialize
' �T�v      �F  �C���X�^���X�������̏�����
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()
    '�������֐��̌Ăяo��
    Call init_csPALS_TestInfo
End Sub


'*********************************************************************************
' �֐���    �F  Class_Terminate
' �T�v      �F  �C���X�^���X�J�����̏���
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Terminate()
    '�C���X�^���X��������ꂽ�ۂɎ��s�����֐�
    Erase m_xSite
End Sub


'*********************************************************************************
' �v���p�e�B��  �F  Site
' �T�v          �F  �eSite�̓����l�����i�[���Ă���Site�I�u�W�F�N�g��Ԃ��v���p�e�B
' ����          �F  sitez :�T�C�g�ԍ�
' ���l          �F  �Ȃ�
' �X�V����      �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Public Property Get site(ByVal sitez As Long) As csPALS_TestInfoParams
    
On Error GoTo errPALSSiteGet
    
    '������site�v���p�e�B�̒l���擾���鏈��������
    Set site = m_xSite(sitez)

Exit Property

errPALSSiteGet:
    Call sub_errPALS("Get site data error at 'csPALS_TestInfo -> Get Site'", "0-6-02-0-23")

End Property


Public Property Let site(ByVal sitez As Long, ByVal vNewValue As csPALS_TestInfoParams)
    
On Error GoTo errPALSSiteLet

    '������site�v���p�e�B�ɒl���Z�b�g���鏈��������
    Set m_xSite(sitez) = vNewValue

Exit Property

errPALSSiteLet:
    Call sub_errPALS("Input site data error at 'csPALS_TestInfo -> Let Site'", "0-6-03-0-24")

End Property


Public Property Set site(ByVal sitez As Long, ByVal vNewValue As csPALS_TestInfoParams)
    
On Error GoTo errPALSSiteSet

    '������site�v���p�e�B�ɃI�u�W�F�N�g���Z�b�g���鏈��������
    Set m_xSite(sitez) = vNewValue

Exit Property

errPALSSiteSet:
    Call sub_errPALS("Set site data error at 'csPALS_TestInfo -> Set Site'", "0-6-04-0-25")

End Property
