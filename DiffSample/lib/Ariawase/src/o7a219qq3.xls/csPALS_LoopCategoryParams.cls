VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_LoopCategoryParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' �N���X��  �F  csPALS_LoopParamsMenber.cls
' �T�v      �F  ��������̊i�[�E�������s���N���X
'               =>Category�AAverage�񐔁AWait
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/10�@�V�K�쐬        K.Sumiyashiki
'==========================================================================

Public category As String                       '�J�e�S����(ex:ML,OF,etc..)
Public Average  As Integer                      '��荞�݉�
Public WAIT     As Double                       '��荞�ݑO�E�F�C�g
Public mode     As String                       '��荞�݉񐔂̔{���w��ݒ�(�����̏ꍇAuto)

Public TargetTestName As String                 '�ő�o���c�L���ږ�
Public VariationLevel As Double                 '�ő�o���c�L��(3��/�K�i��)
Public VariationSite  As Integer                '�ő�o���c�LSite
Public VariationTrend As enum_DataTrendType     '�o���c�L�X��


'*********************************************************************************
' �֐���    �F  Class_Initialize
' �T�v      �F  �C���X�^���X�������̏�����
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub init_csPALS_LoopCategoryParams()

On Error GoTo errPALSinit_csPALS_LoopCategoryParams

    category = ""
    Average = -1
    WAIT = -1               '0�ȏ�̒l���L�蓾��̂ŁA-1�ŏ������B���̕ϐ������ꂵ��-1�ŏ������B
    mode = ""

    TargetTestName = ""
    VariationLevel = -1
    VariationSite = -1
    VariationTrend = em_trend_None

Exit Sub

errPALSinit_csPALS_LoopCategoryParams:
    Call sub_errPALS("csPALS_LoopCategoryParams not initialize at 'init_csPALS_LoopCategoryParams'", "2-3-01-0-36")

End Sub


'*********************************************************************************
' �֐���    �F  Class_Terminate
' �T�v      �F  �C���X�^���X�J�����̏���
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()
    '�N���X�̃C���X�^���X���������ꂽ�ۂɎ��s�����֐�
    '�������֐��̌Ăяo��
    Call init_csPALS_LoopCategoryParams
End Sub


