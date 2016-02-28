VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgLUTInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   LUT�̏��ێ��p
'�ړI:
'   �o�^����LUT�̏��̖₢���킹�p�B
'
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N11��30��
'���ӎ���:
'   LUT�̏��̓��[�v�ōׂ����ݒ肷��(�R�������x)���Ƃ����邽�߁A
'   ���̃N���X�����������ʂɏ����\��������܂��B
'   IMX083�̒�`��Ŋm�F�����ꍇ�A����������͂P�l�ȉ��ł����B
'�Q�l
'   IG-XL��LUT�o�^�d�l
'   �EInputValue1��InputValue2���傫���ƃG���[�B
'   �E����A�h���X�ւ̍Ē�`�͍Ō�̒�`���L��
'   �E�������O��Bit�w��̕ύX�͖��Ȃ��B�ύX�����ꍇ�A�Ō��Bit�w�肪�L���B
'�@�@�@(�R�[�h�w��̏ꍇ�BLUT�V�[�g�̏ꍇ�͕s�\�BEee-JOB�̓G���[�Ƃ���)
'   �EOutputValue ���Z�l�̏����_�ȉ��͐؂�̂āB
'   �@(��Fix�֐������BInt�֐��ł͂Ȃ��B)
'   �E����A�h���X�iInputValue1��InputValue2�������j�ɈႤ�o�͒l��ݒ肵���ꍇ�AOutputValue1���L���B

Option Explicit

Private Const INPUT1 As Integer = 0
Private Const INPUT2 As Integer = 1
Private Const OUTPUT1 As String = 2
Private Const OUTPUT2 As String = 3

Private mName As String                 '����LUT��`�̖��O
Private mValidBit As Integer            '����LUT��`�̗L��bit��
Private mOutputValue() As Long          'LUT�ϊ��l�̕ێ��p�B�z��ԍ���InputValue�ɑ����B
                                        '�z�񐔂�CreateIndexLUT�ōŌ�Ɏw�肵��bit���Ō��肷��B�Ή������炠�ӂꂽ���͖����B(IG-XL�ɍ��킹��)
Private mLutDefinitions As Collection   '��`����Collection

Public Sub Initialize(ByVal pName As String, pBit As Integer)
'���e:
'   LUT�̏���o�^���܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N12��1��
'�p�����[�^:
'   [pName]  In  1):LUT��`��
'   [pBit]   In  2):�o�^���ꂽbit��
'�߂�l:
'
'���ӎ���:

    mName = pName
    mValidBit = pBit

End Sub

Public Sub Add(pName As String, pInputVal1 As Long, pInputVal2 As Long, pOutputVal1 As Long, pOutputVal2 As Long)
'���e:
'   LUT��`�������̃N���X��Collection�ɒǉ����܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2011�N2��3��
'�p�����[�^:
'   [pName]         In  1):LUT��`���B�`�F�b�N���邽�߁B
'   [pInputVal1]    In  2):
'   [pInputVal2]    In  3):
'   [pOutputVal1]   In  4):
'   [pOutputVal2]   In  5):
'�߂�l:
'
'���ӎ���:
'

    Dim pArrayVal(3) As Long    'LUT��`�i�[�p�B���[�U�[��`�ϐ���Collection�ł��Ȃ��̂Ŕz��g���Ă�B
    
    pArrayVal(INPUT1) = pInputVal1
    pArrayVal(INPUT2) = pInputVal2
    pArrayVal(OUTPUT1) = pOutputVal1
    pArrayVal(OUTPUT2) = pOutputVal2

    Call mLutDefinitions.Add(pArrayVal)

End Sub

Public Property Get OutputValue(ByVal pInputValue As Long) As Long
'���e:
'   InputValue�ɑΉ�����OutputValue��Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N12��1��
'�p�����[�^:
'   [pInputValue]    In/Out  1):
'�߂�l:
'   Long
'
'���ӎ���:
'

    If mLutDefinitions.Count <> 0 Then
        Call ExpandData
    End If
    
    On Error GoTo NOT_DEFINITION
    OutputValue = mOutputValue(pInputValue)
    Exit Property
NOT_DEFINITION:
    Call TheError.Raise(9999, "", "This InputValue is not found at LUT." & vbCrLf & "LUT Name:" & mName & vbCrLf & "Error InputValue:" & pInputValue)

End Property

Public Function AllValue() As Long()
'���e:
'   �S�ϊ����𓮓I�z��ŕԂ��܂��B
'   �z��ԍ������͒l�A�Ή��l���o�͒l�ɑ������܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N12��1��
'�p�����[�^:
'
'�߂�l:
'   Long
'
'���ӎ���:
'
    If mLutDefinitions.Count <> 0 Then
        Call ExpandData
    End If
    
    AllValue = mOutputValue

End Function

Public Property Get Bit() As Integer
'���e:
'   ����LUT�̗L��Bit����Ԃ��܂��B
'�쐬��:
'  tomoyoshi.takase
'�쐬��: 2010�N12��1��
'�p�����[�^:
'   �Ȃ�
'�߂�l:
'   Integer
'
'���ӎ���:

    Bit = mValidBit
End Property

Private Sub ExpandData()
'���e:
'   �o�^����LUT��`�f�[�^���A�z����ɓW�J���܂��B

Dim x1 As Long, y1 As Long      '�Q�_�̍��W�̎n�_
Dim x2 As Long, y2 As Long      '�Q�_�̍��W�̏I�_
Dim A As Double                 '�Q�_��ʂ钼���̌X��
Dim B As Double                 '�Q�_��ʂ钼���̐ؕ�
Dim pTmp  As Double             '�ϊ����ʂ̊i�[�p
Dim pValue  As Variant          'Value�Q�̈ꎞ�ێ��p

Dim pLimitData As Long          'LUT�f�[�^�͈̔�

Dim i As Long

    '#### ���[�U�[��`�Ō��bit�w�肩��f�[�^�͈͂����� ####
    pLimitData = 2 ^ mValidBit / 2                    '�v���X�}�C�i�X�͈͂Ȃ̂Ł��Q
    If UBound(mOutputValue) = 0 Then                    '���̃��\�b�h�����������Ƃ��Ȃ��B
        ReDim mOutputValue(-pLimitData To pLimitData - 1)
    End If
    '#### ���[�U�[��`�f�[�^��z��ɓW�J ####
    For Each pValue In mLutDefinitions
        '�n�_���W
        x1 = pValue(INPUT1)
        y1 = pValue(OUTPUT1)
        '�I�_���W
        x2 = pValue(INPUT2)
        y2 = pValue(OUTPUT2)
        
        '#### �Q�_��ʂ钼���̎Z�o ####
        
        If x2 - x1 = 0 Then                 '�[�����΍�
            mOutputValue(x1) = y1           'y1��y2��������ꍇ�Ay1���L��
            
        ElseIf x2 - x1 = 1 Then             'For��1Step���݂Œ�`���邱�Ƃ��������Ȃ̂ŁA�v�Z�����őΉ�
            mOutputValue(x1) = y1
            mOutputValue(x2) = y2
        Else
            '�Q�_��ʂ������(y=ax+b)���Z�o���āAOutput�f�[�^��z��ɓW�J
            A = (y1 - y2) / (x1 - x2)               '�X��
            B = (x1 * y2 - x2 * y1) / (x1 - x2)     '�ؕ�
            
            For i = x1 To x2
                pTmp = A * i + B
                '�����_�ȉ��؎̂�(fix�֐������BInt�֐��Ƃ͈Ⴄ�BIG-XL�d�l���B)
                pTmp = Fix(pTmp)
                
                mOutputValue(i) = pTmp
            Next i
        End If
        
    Next
    
    Set mLutDefinitions = Nothing
    Set mLutDefinitions = New Collection

End Sub

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

    ReDim mOutputValue(0)
    Set mLutDefinitions = New Collection

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

    Set mLutDefinitions = Nothing

End Sub