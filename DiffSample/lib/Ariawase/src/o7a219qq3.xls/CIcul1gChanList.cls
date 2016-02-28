VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIcul1gChanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@@@IP750EX Support
'�T�v:
'   ICUL1G�̃`�����l�����X�g�Ǘ��p�N���X
'
'�ړI:
'   ICUL1G���\�[�X�݂̂̃`�����l���}�b�v���
'

Option Explicit

'Tool�Ή���ɃR�����g�O���Ď��������ɂ���B�@2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1               '1CUL1G�{�[�h�̎g�p�@  0�F���g�p�A0�ȊO�F�g�p  <Tester��IP750EX�Ȃ�Default:1�ɂ��Ă���>

#If ICUL1G_USE <> 0 Then

Private m_List As Collection '�`�����l�����X�g

'�R���X�g���N�^
Private Sub Class_Initialize()
End Sub

'�f�X�g���N�^
Private Sub Class_Terminate()
    Set m_List = Nothing
End Sub

'�w��̃s�����AICUL1G�{�[�h�ɏ�������s�������m�F����֐�
Public Function IsICUL1GPins(ByVal pPinList As String, Optional ByVal pType As Icul1gPinType = icul1gPinAny) As Boolean
    'TERADYNE-API�𗘗p
    IsICUL1GPins = TheHdw.ICUL1G.IsICUL1GPin(pPinList, pType)
End Function

'���ׂẴ`�����l�����X�g�̌��J
Public Property Get AllInfo() As Collection
    Call MakeList
    If m_List.Count <> 0 Then
        Set AllInfo = m_List
    Else
        Set AllInfo = Nothing
    End If
End Property

'�w��^�C�v�̃`�����l�����X�g���J
Public Property Get AtChanType(ByVal pTargetType As Icul1gChannelType) As Collection
    Call MakeList
    If m_List.Count <> 0 Then
        Set AtChanType = MakeTypeOfList(pTargetType)
    Else
        Set AtChanType = Nothing
    End If
End Property

'�w��s���̃`�����l�����X�g���J
Public Property Get AtPin(ByVal pPinName As String) As Collection
    If CheckIsSinglePin(pPinName) = True And Me.IsICUL1GPins(pPinName) = True Then
        Call MakeList
        Set AtPin = New Collection
        Dim Retinfo As CIcul1gChanInfo
        For Each Retinfo In m_List
            If StrComp(Retinfo.PinName, pPinName, vbTextCompare) = 0 Then
                Call AtPin.Add(Retinfo)
            End If
        Next
    Else
        Set AtPin = Nothing
    End If
End Property

'�w��s�����X�g�̃`�����l�����X�g���J
Public Property Get AtPinList(ByVal pCsvPinList As String) As Collection
    If Me.IsICUL1GPins(pCsvPinList) = True Then
        Dim PinList() As String
        Dim pinCount As Long
        Call TheExec.DataManager.DecomposePinList(pCsvPinList, PinList, pinCount)
        Set AtPinList = New Collection
        Dim itemCount As Long
        Dim Retinfo As CIcul1gChanInfo
        Dim ChanList As Collection
        For itemCount = 0 To UBound(PinList) Step 1
            Set ChanList = Me.AtPin(PinList(itemCount))
            If Not ChanList Is Nothing Then
                For Each Retinfo In ChanList
                    Call AtPinList.Add(Retinfo)
                Next
            End If
        Next itemCount
    Else
        Set AtPinList = Nothing
    End If
End Property

'�ЂƂ̃s���̃`�����l������Ԃ��֐�
Public Property Get AtSiteAndPin(ByVal pPinName As String, ByVal pSiteNumber As Long) As CIcul1gChanInfo
    If CheckIsSinglePin(pPinName) = True And Me.IsICUL1GPins(pPinName) = True And CheckSiteNumber(pSiteNumber) = True Then
        Dim Retinfo As Collection
        Set Retinfo = Me.AtSiteAndPinList(pPinName, pSiteNumber)
        Set AtSiteAndPin = Retinfo.Item(1)
    Else
        Set AtSiteAndPin = Nothing
    End If
End Property

'�w��T�C�g�̃`�����l�����X�g��Ԃ��֐�
Public Property Get AtSiteAndPinList(ByVal pPinList As String, ByVal pSiteNumber As Long) As Collection
    If Me.IsICUL1GPins(pPinList) = True And CheckSiteNumber(pSiteNumber) = True Then
        Set AtSiteAndPinList = New Collection
        Dim PinList() As String
        Dim pinCount As Long
        Call TheExec.DataManager.DecomposePinList(pPinList, PinList, pinCount)
        Dim AllSiteList As Collection
        Dim itemCount As Long
        Dim RetChanInfo As CIcul1gChanInfo
        For itemCount = 0 To UBound(PinList) Step 1
            Set AllSiteList = Me.AtPin(PinList(itemCount))
            For Each RetChanInfo In AllSiteList
                If RetChanInfo.SiteNumber = pSiteNumber Then
                    Call AtSiteAndPinList.Add(RetChanInfo)
                End If
            Next
        Next itemCount
    Else
        Set AtSiteAndPinList = Nothing
    End If
End Property


'============== �ȉ� Private �@�\�Q ==============

'�T�C�g�ԍ��̊m�F(TRUE:OK)
Private Function CheckSiteNumber(ByVal pNumber As Long) As Boolean
    If pNumber < 0 Or TheExec.sites.ExistingCount - 1 < pNumber Then
        CheckSiteNumber = False
    Else
        CheckSiteNumber = True
    End If
End Function

'�s�����X�g�ɕ����s������`����Ă��邩�ۂ����m�F����֐�(TRUE:OK)
Private Function CheckIsSinglePin(ByVal pPinName As String) As Boolean
    Dim RetPins() As String
    Dim pinCount As Long
    Call TheExec.DataManager.DecomposePinList(pPinName, RetPins, pinCount)
    If pinCount <> 1 Then
        CheckIsSinglePin = False
    Else
        CheckIsSinglePin = True
    End If
End Function

'�w��^�C�v�̃`�����l�����X�g���쐬���ĕԂ��֐�
Private Function MakeTypeOfList(ByVal pType As Icul1gChannelType) As Collection
    Set MakeTypeOfList = New Collection
    Dim Retinfo As CIcul1gChanInfo
    For Each Retinfo In m_List
        If Retinfo.Icul1gChanType = pType Then
            Call MakeTypeOfList.Add(Retinfo)
        End If
    Next
    If MakeTypeOfList.Count = 0 Then
        Set MakeTypeOfList = Nothing
    End If
End Function

'�s���ƃ`�����l���̃��X�g�쐬�֐�
Private Sub MakeList()

    Dim RetChannelList() As Long
    Dim RetPinList() As String
    Dim RetSiteCount As Long

    '�����Ώۂ̃`�����l���^�C�v���X�g���쐬
    Dim TargetBoardType As Collection
    Set TargetBoardType = New Collection
    With TargetBoardType
        Call .Add(chICUL1Gclk_high)
        Call .Add(chICUL1Gclk_low)
        Call .Add(chICUL1Gdata_high)
        Call .Add(chICUL1Gdata_low)
    End With
    
    '�`�����l�����쐬����
    Set m_List = New Collection
    Dim TargetType As Variant
    Dim RetStatus As Boolean
    For Each TargetType In TargetBoardType
        RetStatus = GetChannelMapInformations(TargetType, RetChannelList, RetPinList, RetSiteCount)
        If RetStatus = True Then
            Call MakeChannelIInformations(TargetType, RetChannelList, RetPinList, m_List, RetSiteCount)
        End If
    Next

End Sub

'�`�����l���������֐�
Private Sub MakeChannelIInformations(ByVal pType As Icul1gChannelType, _
ByRef pTargetChanList() As Long, ByRef pTargetPinList() As String, _
ByRef pChannelList As Collection, ByVal pSiteCount As Long)
    
    Dim pinCount As Long
    Dim ChanCount As Long
    Dim StartCount As Long
    Dim CreateCount As Long
    Dim SiteNumber As Long
    Dim ChannelInfo As CIcul1gChanInfo
                
    StartCount = 0
    CreateCount = 0
        
    For pinCount = 0 To UBound(pTargetPinList) Step 1
        SiteNumber = 0
        For ChanCount = StartCount To StartCount + pSiteCount - 1 Step 1
            Set ChannelInfo = New CIcul1gChanInfo
            With ChannelInfo
                .chanType = pType
                .ChanNumber = pTargetChanList(ChanCount)
                .PinName = pTargetPinList(pinCount)
                .SiteNumber = SiteNumber
            End With
            Call pChannelList.Add(ChannelInfo)
            CreateCount = CreateCount + 1
            SiteNumber = SiteNumber + 1
        Next ChanCount
        StartCount = CreateCount
    Next pinCount

End Sub

'�w�胊�\�[�X�̃`�����l���}�b�v�����擾����֐�
Private Function GetChannelMapInformations(ByVal pChannelType As chtype, _
ByRef pChannelList() As Long, _
ByRef pPinList() As String, _
ByRef pSiteCount As Long) As Boolean
    
    Const ALL_SITE As Long = -1
    
    '�w�胊�\�[�X���g�p���Ă���PIN�����擾
    Dim RetPinCnt As Long
    Call TheExec.DataManager.GetPinNames(pPinList, pChannelType, RetPinCnt)
                                                   
    '�w�肳�ꂽ���\�[�X�����݂��Ȃ��Ƃ���False��Ԃ��ďI��
    If RetPinCnt = 0 Then
        GetChannelMapInformations = False
        Exit Function
    End If
                                                
    'CSV�s�����X�g�����
    Dim CsvPinList As String
    CsvPinList = MakeCsvPinList(pPinList)
                    
    '�w�胊�\�[�X�̂��ׂĂ�PIN�̃`�����l���ԍ����擾
    Dim RetChanCnt As Long
    Dim RetMsg As String
    Call TheExec.DataManager.GetChanList(CsvPinList, ALL_SITE, pChannelType, pChannelList, RetChanCnt, pSiteCount, RetMsg)
    GetChannelMapInformations = True

End Function

'�z��`���̃s�����X�g��CSV�`��������ɕϊ�����֐�
Private Function MakeCsvPinList(ByRef pPinList() As String) As String
    MakeCsvPinList = pPinList(0)
    Dim itemCount As Long
    For itemCount = 1 To UBound(pPinList) Step 1
        MakeCsvPinList = MakeCsvPinList & "," & pPinList(itemCount)
    Next itemCount
End Function


#End If
