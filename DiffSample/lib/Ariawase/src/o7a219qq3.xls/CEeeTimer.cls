VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEeeTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   PerformanceCounter�𗘗p�����^�C�}�[�@�\�̒�
'
'�ړI:
'   Private�^�C�}�[�̎���
'
'�쐬��:
'   SLSI����
'
'���ӓ_:
'

Option Explicit

'�@�\�̏���
Private Declare Function QueryPerformanceCounter Lib "kernel32" (pPerformanceCount As Currency) As Boolean
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (pFrequency As Currency) As Boolean

Private Const ERR_NUMBER = 9999                '�G���[���ɓn���G���[�ԍ�
Private Const CLASS_NAME = "CEeeTimer"         '���̃N���X�̖��O

'���ԋL�^�p
Private mStartCount As Currency

'������񏉊�������
Public Sub Initialize()
    mStartCount = 0#
End Sub

'�J�n���Ԃ̎擾
Public Sub StartTimer()
    Call QueryPerformanceCounter(mStartCount)
End Sub

'�r�����Ԃ̎擾
Public Function LapTimer() As Double
    LapTimer = CalculationTime(mStartCount)
End Function

'�I�����Ԃ̎Z�o
Public Function StopTimer() As Double
    StopTimer = CalculationTime(mStartCount)
    mStartCount = 0#
End Function

'���ԎZ�o����
Private Function CalculationTime(ByVal pStartCount As Currency) As Double
    Dim curFrequency As Currency
    Dim EndCount As Currency
    
    If pStartCount <= 0# Then
        Call RaiseError("StartTimer is unexecution")
    Else
        Call QueryPerformanceCounter(EndCount)
        Call QueryPerformanceFrequency(curFrequency)
        If mStartCount > EndCount Then
            Call RaiseError("Time calculation error")
        Else
'            Debug.Print "END=" & EndCount & " Start=" & mStartCount
            CalculationTime = (EndCount - mStartCount) / curFrequency
        End If
    End If

End Function

'�G���[���C�Y�p�֐�
Private Sub RaiseError(ByVal pErrMsg As String)
    If Not (TheError Is Nothing) Then
        '�G���[�}�l�[�W����
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME, pErrMsg)
    Else
        'VB��Err��
        Call Err.Raise(ERR_NUMBER, CLASS_NAME, pErrMsg)
    End If
End Sub
