VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIcul1gPpmuResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@@@IP750EX Support
'�T�v:
'   PinListData�^�̃N���X��ICUL1��Measure���ʗp�Ɋg�������N���X
'
'�ړI:
'   �s���@�\�̒ǉ�
'

Option Explicit

'Tool�Ή���ɃR�����g�O���Ď��������ɂ���B�@2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1               '1CUL1G�{�[�h�̎g�p�@  0�F���g�p�A0�ȊO�F�g�p  <Tester��IP750EX�Ȃ�Default:1�ɂ��Ă���>

#If ICUL1G_USE <> 0 Then

Private m_PinListData As PinListData 'PINLIST�f�[�^Obejct

'�R���X�g���N�^
Private Sub Class_Initialize()
    Set m_PinListData = New PinListData                    'PINLIST�f�[�^Obejct����
    m_PinListData.ResultType = tlResultTypeParametricValue '����l�i�[Mode�ݒ�
End Sub

'�f�X�g���N�^
Private Sub Class_Terminate()
    Set m_PinListData = Nothing
End Sub

'����Ώۃs�����X�g�̌��J
Public Property Get PinList() As String
    If m_PinListData.Pins.Count = 0 Then
        PinList = Empty
    Else
        PinList = PinListData2CsvPinList
    End If
End Property

'PinListData�I�u�W�F�N�g�̌��J
Public Property Get PinListData() As PinListData
    Set PinListData = m_PinListData
End Property

'PerPinResult�̌��J
Public Property Get PerPinResult() As Collection
    If m_PinListData.Pins.Count = 0 Then
        '����f�[�^���o�^
        Set PerPinResult = Nothing
    Else
        Set PerPinResult = PinListData2PerPinResults(m_PinListData)
    End If
End Property

'PinResult�̌��J
Public Property Get PinResult(ByVal pPinName As String) As Double()
    If m_PinListData.Pins.Count = 0 Then
        '����f�[�^���o�^
        Erase PinResult
    Else
        Dim PerPinResult As Collection
        Set PerPinResult = PinListData2PerPinResults(m_PinListData)
        PinResult = PerPinResult.Item(pPinName)
    End If
End Property


'--------- �ȉ� Private --------------------------------------------------

'�s�����X�g�f�[�^����PerPinResult�f�[�^���쐬�i�A�x���[�W�̊��߂��A�����T�C�g�̌��ʒl��0��Ԃ��j
Private Function PinListData2PerPinResults(ByRef pResultData As PinListData) As Collection

    Set PinListData2PerPinResults = New Collection
    
    'CSV�s�����X�g��z��s�����X�g�ɂ���
    Dim RetPinList() As String
    Dim RetPinCount As Long '���g�p
    With TheExec.DataManager
        Call .DecomposePinList(PinListData2CsvPinList, RetPinList, RetPinCount)
    End With
        
    Dim pinCount As Long
    Dim tempResult As Double
    Dim SiteCount As Long
    Dim retResult As Variant
    Dim TargetPin As PinData
    Dim SiteResult() As Double
    Dim SampleCount As Long
    Dim targetSite As Long
    
    '�s�����X�gLOOP
    For pinCount = 0 To UBound(RetPinList) Step 1
        
        '�����Ώۃs���̑I��
        Set TargetPin = pResultData.Pins(RetPinList(pinCount))
        
        '�T�C�g�����ʂ̓��ꕨ�p��
        ReDim SiteResult(TheExec.sites.ExistingCount - 1)
    
        '���莞�̗L���T�C�g�ԍ�LOOP
        For SiteCount = 0 To UBound(pResultData.sites) Step 1
                
                '�^�[�Q�b�g�T�C�g�̃f�[�^���o���i���莞�̗L���T�C�g�̂݁j
                targetSite = pResultData.sites(SiteCount)
                retResult = TargetPin.Value(targetSite)
                
                '�A�x���[�W���葪��f�[�^
                If IsArray(retResult) = True Then
                    '�A�x���[�W���߂�����
                    tempResult = 0#
                    For SampleCount = 0 To UBound(retResult) Step 1
                        'SAMPLE���v�f�[�^�Z�o
                        tempResult = tempResult + retResult(SampleCount)
                    Next SampleCount
                    '���߂�����
                    SiteResult(targetSite) = tempResult / (UBound(retResult) + 1)
                '�A�x���[�W�Ȃ�����f�[�^
                Else
                    SiteResult(targetSite) = CDbl(retResult)
                End If
                
        Next SiteCount
    
        '�s���̃f�[�^�Ƃ��ăT�C�g�z���o�^
        Call PinListData2PerPinResults.Add(SiteResult, RetPinList(pinCount))
    
    Next pinCount

End Function

'PinList�̃f�[�^����CSVPinList���쐬����֐�
Private Function PinListData2CsvPinList() As String
    Dim pinCount As Long
    PinListData2CsvPinList = Empty
    For pinCount = 0 To m_PinListData.Pins.Count - 1 Step 1
        If pinCount = 0 Then
            PinListData2CsvPinList = m_PinListData.Pins(pinCount).Name
        Else
            PinListData2CsvPinList = PinListData2CsvPinList & "," & m_PinListData.Pins(pinCount).Name
        End If
    Next pinCount
End Function

#End If
