VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_TestInfoParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' �N���X��  �F  csLoopPTestInfo.cls
' �T�v      �F  ����f�[�^�̊i�[�A�������s���N���X
'               =>data,Min,Ave,Max,��
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/10�@�V�K�쐬        K.Sumiyashiki
'           �F  Rev2.0      2011/05/16�@�����o�ǉ�      K.Sumiyashiki
'                                       ��Enable, ActiveValueCnt�ǉ�
'           �F  Rev3.0      2011/06/17�@�֐��ǉ�        K.Sumiyashiki
'                                       ��sub_CalcDifferentialSigma�ǉ�
'==========================================================================

'�����l�f�[�^
Private m_Min   As Double               '�ŏ��l
Private m_Ave   As Double               '���ϒl
Private m_Max   As Double               '�ő�l
Private m_Sigma As Double               '�V�O�}
Private m_testdata() As Double          '�����l�f�[�^
'>>>2011/05/12 K.SUMIYASHIKI ADD
Private m_PassOrFail() As Boolean       'Result(PASS/FAIL)���
Private m_ActiveValueCnt As Integer     '���̃e�X�g���ڂ�PASS������
'<<<2011/05/12 K.SUMIYASHIKI ADD


'********************************************************************************************
' ���O: init_csPALS_TestInfoParams
' ���e: �e�ϐ��̏�����
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub init_csPALS_TestInfoParams()
'�N���X������
    m_Min = 0
    m_Ave = 0
    m_Max = 0
    m_Sigma = 0
'>>>2011/06/20 K.SUMIYASHIKI ADD
    m_ActiveValueCnt = 0
'>>>2011/06/20 K.SUMIYASHIKI ADD
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
    Call init_csPALS_TestInfoParams
End Sub


'*********************************************************************************
' �֐���    �F  Class_Terminate
' �T�v      �F  �C���X�^���X�J�����̏���
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Terminate()

    Erase m_testdata

End Sub


'*********************************************************************************
' �v���p�e�B��  �F  Data
' �T�v          �F  �eSite�̓����l����Ԃ��v���p�e�B
' ����          �F  index :�����
' ���l          �F  �Ȃ�
' �X�V����      �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Public Property Get Data(index As Long) As Double
    
'    If g_ErrorFlg_PALS Then
'        Exit Sub
'    End If
    
On Error GoTo errPALSDataGet

    '������data�v���p�e�B�̒l���擾���鏈��������
    If index > g_MaxPalsCount Then
        Call sub_errPALS("Data count can't set over " & g_MaxPalsCount, "0-7-01-5-26")
        Exit Property
    End If
    
    Data = m_testdata(index)

Exit Property

errPALSDataGet:
    Call sub_errPALS("Get measure data error at 'csPALS_TestInfoParams -> Get Data'", "0-7-01-0-27")

End Property

Public Property Let Data(index As Long, ByVal testdata As Double)
    
'    If g_ErrorFlg_PALS Then
'        Exit Property
'    End If
    
On Error GoTo errPALSDataLet
    
    '������data�v���p�e�B�ɒl���Z�b�g���鏈��������
    If index > g_MaxPalsCount Then
        Call sub_errPALS("Data count can't set over  " & g_MaxPalsCount, "0-7-02-5-28")
        Exit Property
    End If
    
    m_testdata(index) = testdata
    
'>>>2011/05/12 K.SUMIYASHIKI ADD
    m_ActiveValueCnt = m_ActiveValueCnt + 1
'<<<2011/05/12 K.SUMIYASHIKI ADD

    '���ρA�ő�A�ŏ��A�Ђ��v�Z
    Call calc_ave(index, testdata)
    Call calc_max(index, testdata)
    Call calc_min(index, testdata)
    Call calc_sigma(index)
    
Exit Property

errPALSDataLet:
    Call sub_errPALS("Input measure data error at 'csPALS_TestInfoParams -> Let Data'", "0-7-02-0-29")
    
End Property


'********************************************************************************************
' ���O: DataCount
' ���e: �ő呪��񐔂�Ԃ��֐�
' ����: �Ȃ�
' �ߒl: �ő呪���
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Function DataCount() As Integer

    DataCount = UBound(m_testdata)

End Function


'*********************************************************************************
' �v���p�e�B��  �F  Ave
' �T�v          �F  �w�荀�ڂ̕��ϒl��Ԃ��v���p�e�B
' ����          �F  �Ȃ�
' ���l          �F  �ǂݎ���p
' �X�V����      �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Public Property Get ave() As Double

    ave = m_Ave

End Property


'*********************************************************************************
' �v���p�e�B��  �F  Max
' �T�v          �F  �w�荀�ڂ̍ő�l��Ԃ��v���p�e�B
' ����          �F  �Ȃ�
' ���l          �F  �ǂݎ���p
' �X�V����      �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Public Property Get max() As Double

    max = m_Max

End Property


'*********************************************************************************
' �v���p�e�B��  �F  Min
' �T�v          �F  �w�荀�ڂ̍ŏ��l��Ԃ��v���p�e�B
' ����          �F  �Ȃ�
' ���l          �F  �ǂݎ���p
' �X�V����      �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Public Property Get Min() As Double

    Min = m_Min

End Property


'*********************************************************************************
' �v���p�e�B��  �F  Sigma
' �T�v          �F  �w�荀�ڂ̃Ђ�Ԃ��v���p�e�B
' ����          �F  �Ȃ�
' ���l          �F  �ǂݎ���p
' �X�V����      �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Public Property Get Sigma() As Double

    Sigma = m_Sigma

End Property


'********************************************************************************************
' ���O: calc_ave
' ���e: �w�荀�ڂ̕��ϒl���v�Z����֐�
' ����: mes_cnt   : �����
'       testdata  : �ŐV�̑���f�[�^
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub calc_ave(ByVal mes_cnt As Long, ByVal testdata As Double)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALScalc_ave

'>>>2011/06/20 K.SUMIYASHIKI UPDATE
'    If mes_cnt = 1 Then
    If m_ActiveValueCnt = 1 Then
        m_Ave = testdata
'<<<2011/06/20 K.SUMIYASHIKI UPDATE
    
    Else
'>>>2011/05/12 K.SUMIYASHIKI CHANGE
'old        m_Ave = (m_Ave * (mes_cnt - 1) + testdata) / mes_cnt
        m_Ave = (m_Ave * (m_ActiveValueCnt - 1) + testdata) / m_ActiveValueCnt
'<<<2011/05/12 K.SUMIYASHIKI CHANGE
    End If

Exit Sub

errPALScalc_ave:
    Call sub_errPALS("Calculate average error at 'csPALS_TestInfoParams -> calc_ave'", "0-7-03-0-30")

End Sub


'********************************************************************************************
' ���O: calc_max
' ���e: �w�荀�ڂ̍ő�l���v�Z����֐�
' ����: mes_cnt   : �����
'       testdata  : �ŐV�̑���f�[�^
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub calc_max(ByVal mes_cnt As Long, ByVal testdata As Double)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALScalc_max

'>>>2011/06/20 K.SUMIYASHIKI UPDATE
'    If mes_cnt = 1 Then
    If m_ActiveValueCnt = 1 Then
        m_Max = testdata
'<<<2011/06/20 K.SUMIYASHIKI UPDATE
    ElseIf testdata > m_Max Then
        m_Max = testdata
    Else
    
    End If

Exit Sub

errPALScalc_max:
    Call sub_errPALS("Calculate max data error at 'csPALS_TestInfoParams -> calc_max'", "0-7-04-0-31")

End Sub


'********************************************************************************************
' ���O: calc_min
' ���e: �w�荀�ڂ̍ŏ��l���v�Z����֐�
' ����: mes_cnt   : �����
'       testdata  : �ŐV�̑���f�[�^
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub calc_min(ByVal mes_cnt As Long, ByVal testdata As Double)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALScalc_min

'>>>2011/06/20 K.SUMIYASHIKI UPDATE
'    If mes_cnt = 1 Then
    If m_ActiveValueCnt = 1 Then
        m_Min = testdata
'<<<2011/06/20 K.SUMIYASHIKI UPDATE
    ElseIf m_Min > testdata Then
        m_Min = testdata
    Else
    
    End If

Exit Sub

errPALScalc_min:
    Call sub_errPALS("Calculate minimum data error at 'csPALS_TestInfoParams -> calc_min'", "0-7-05-0-32")

End Sub


'********************************************************************************************
' ���O: calc_sigma
' ���e: �w�荀�ڂ̃Ђ��v�Z����֐�
' ����: mes_cnt   : �����
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function calc_sigma(ByVal mes_cnt As Long) As Double

    If mes_cnt = 1 Or g_ErrorFlg_PALS Then
        m_Sigma = 0
        Exit Function
    End If
    
On Error GoTo errPALScalc_sigma

    m_Sigma = 0

    Dim sigma_i As Long
'>>>2011/05/12 K.SUMIYASHIKI CHANGE
    For sigma_i = 1 To mes_cnt
        If m_PassOrFail(sigma_i) = True Then
            m_Sigma = m_Sigma + ((m_Ave - m_testdata(sigma_i)) ^ 2)
        End If
    Next sigma_i
    
'old    m_Sigma = Sqr(m_Sigma / (mes_cnt))
    m_Sigma = Sqr(m_Sigma / (m_ActiveValueCnt))
'<<<2011/05/12 K.SUMIYASHIKI CHANGE

Exit Function

errPALScalc_sigma:
    Call sub_errPALS("Calculate sigma data error at 'csPALS_TestInfoParams -> calc_sigma'", "0-7-06-0-33")

End Function


'********************************************************************************************
' ���O: calc_sigma
' ���e: �t�H�[���Ŏw�肵�����[�v�񐔂ɉ����āA�z�񐔂�ύX����֐�
' ����: lngCnt   : �ő呪���
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/18�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Sub sub_ChangeDataDivision(ByVal lngCnt As Long)

    ReDim m_testdata(lngCnt)

'>>>2011/05/12 K.SUMIYASHIKI ADD
    ReDim m_PassOrFail(lngCnt)
'<<<2011/05/12 K.SUMIYASHIKI ADD

End Sub


'*********************************************************************************
' �v���p�e�B��  �F  ActiveValueCnt
' �T�v          �F  PASS�������ڐ���Ԃ��v���p�e�B
' ����          �F  �Ȃ�
' ���l          �F  �ǂݎ���p
' �X�V����      �F  Rev1.0      2011/05/13�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Public Property Get ActiveValueCnt() As Double

    ActiveValueCnt = m_ActiveValueCnt

End Property


'*********************************************************************************
' �v���p�e�B��  �F  Enable
' �T�v          �F  �����l��PASS or FAIL��Ԃ��v���p�e�B
' ����          �F  index :�����
' ���l          �F  �Ȃ�
' �X�V����      �F  Rev1.0      2011/05/16�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Public Property Get Enable(index As Long) As Boolean

On Error GoTo errPALSEnableGet

    '������Enable�v���p�e�B�̒l���擾���鏈��������
    Enable = m_PassOrFail(index)

Exit Property

errPALSEnableGet:
    Call sub_errPALS("Get measure data error at 'csPALS_TestInfoParams -> Get Enable'", "0-7-07-0-34")

End Property

Public Property Let Enable(index As Long, ByVal ChangeData As Boolean)
    
On Error GoTo errPALSEnableLet
    
    '������Enable�v���p�e�B�ɒl���Z�b�g���鏈��������
    m_PassOrFail(index) = ChangeData
    
Exit Property

errPALSEnableLet:
    Call sub_errPALS("Input measure data error at 'csPALS_TestInfoParams -> Let Enable'", "0-7-08-0-35")
    
End Property



'*********************************************************************************
' �v���p�e�B��  �F  Differential_Sigma
' �T�v          �F  �w�荀�ڂ̔����l�̃Ђ�Ԃ��v���p�e�B
' ����          �F  �Ȃ�
' ���l          �F  �ǂݎ���p
' �X�V����      �F  Rev1.0      2011/06/17�@�V�K�쐬   K.Sumiyashiki
'*********************************************************************************
Public Property Get Differential_Sigma(ByVal mes_cnt As Long) As Double

    Differential_Sigma = sub_CalcDifferentialSigma(mes_cnt)

End Property


'********************************************************************************************
' ���O: sub_CalcDifferentialSigma
' ���e: �w�荀�ڂ̔����l�̃Ђ��v�Z����֐�
' ����: mes_cnt   : �����
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2011/06/17�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function sub_CalcDifferentialSigma(ByVal mes_cnt As Long) As Double

    If mes_cnt = 1 Then
        sub_CalcDifferentialSigma = 0
        Exit Function
    End If
    
On Error GoTo errPALSsub_CalcDifferentialSigma

    Dim dblDifferentialData() As Double     '�����l���i�[����z��
    Dim sigma_i As Long                     '���[�v�ϐ�
    Dim lngActiveDataCnt As Long            '�L���Ȕ����f�[�^��
    Dim tmpStandardData As Double           '�����l�擾���ɑO�̃f�[�^���ꎞ�i�[���Ă����ϐ�

    '�L�������f�[�^���̏�����(Redim�̏��������₷���Ȃ�̂ŁA�����l��-1)
    lngActiveDataCnt = -1

    '�f�[�^�����J��Ԃ�
    For sigma_i = 1 To mes_cnt
        '�L���ȃf�[�^�̏ꍇ�̂ݏ��������{
        If m_PassOrFail(sigma_i) = True Then
        
            '�L���f�[�^1��ڂ̏���
            If lngActiveDataCnt = -1 Then
                '�����l���擾����ׂɁA�f�[�^���ꎞ�i�[
                tmpStandardData = m_testdata(sigma_i)
                '�L���Ȕ����f�[�^�����C���N�������g
                lngActiveDataCnt = lngActiveDataCnt + 1
            Else
                '�����f�[�^���i�[����z��̍Ē�`
                ReDim Preserve dblDifferentialData(lngActiveDataCnt)
                '�����f�[�^�擾
                dblDifferentialData(lngActiveDataCnt) = m_testdata(sigma_i) - tmpStandardData
                '���̑���f�[�^�Ƃ̔����l���擾����ׁA�f�[�^���ꎞ�i�[
                tmpStandardData = m_testdata(sigma_i)
                '�L���Ȕ����f�[�^�����C���N�������g
                lngActiveDataCnt = lngActiveDataCnt + 1
            End If
        End If
    Next sigma_i
    
    '�L�������f�[�^��0�̏ꍇ�̃G���[����
    If lngActiveDataCnt < 1 Then
        sub_CalcDifferentialSigma = 0
        Exit Function
    End If
    
    
    Dim dblDifferentialAve As Double
    '�ϐ�������
    dblDifferentialAve = 0
    
    '�L���f�[�^�����J��Ԃ�
    For sigma_i = 0 To lngActiveDataCnt - 1
        '���v�l�擾
        dblDifferentialAve = dblDifferentialAve + dblDifferentialData(sigma_i)
    Next sigma_i
    
    '���ϒl�擾
    dblDifferentialAve = dblDifferentialAve / lngActiveDataCnt
        
        
    Dim dblSumData As Double
    '�ϐ�������
    dblSumData = 0
    
    '�Ђ̌v�Z
    For sigma_i = 0 To lngActiveDataCnt - 1
        dblSumData = dblSumData + ((dblDifferentialAve - dblDifferentialData(sigma_i)) ^ 2)
    Next sigma_i
    
    sub_CalcDifferentialSigma = Sqr(dblSumData / lngActiveDataCnt)

Exit Function

errPALSsub_CalcDifferentialSigma:
    Call sub_errPALS("Calculate differential sigma data error at 'csPALS_TestInfoParams -> sub_CalcDifferentialSigma'", "0-7-06-0-36")
End Function