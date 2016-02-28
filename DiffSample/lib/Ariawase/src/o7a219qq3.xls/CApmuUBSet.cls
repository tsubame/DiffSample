VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CApmuUBSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   APMU��UtilityBit�ݒ�p�N���X
'
'�ړI:
'   Utility Bit�����\�ɒ�`���ꂽ�����g�p����
'   Utility Bit��ݒ肷��
'
'���ӎ���:
'   APMU Utility Bit�����\�̃��[�N�V�[�g�Ƒg�ݍ��킹�Ďg�p���܂��B
'
'�쐬��:
'   SLSI����
'
'Code Checked
'

Option Explicit

Private m_UtyBitSetWs As Object        ' APMU UtilityBit�����\�V�[�gOBJ
Private m_LastConditionName As String  ' �ŏI�ݒ�������i�[�p�ϐ�
Private m_UtyBitStatus() As Long       ' UB�ݒ��ԋL�^�p
Private Const MAX_BIT_CH = 255         ' UBCH�̍ő�l

'#Pass-Ok
Public Property Get BoardNumber() As Long
'���e:
'   APMU�{�[�hNO�̐ݒ�l���擾����
'
'�p�����[�^:
'
'�߂�l:
'   �V�[�g�ɒ�`����Ă���APMU�{�[�h�ԍ�
'
'���ӎ���:
'
    
    '�{�[�h�ԍ��̒�`��񂪑��݂��Ȃ��ꍇ��
    '���[�N�V�[�g�̃f�[�^��ǂݍ��݂܂�
    If m_UtyBitSetWs.ApmuBoardNum = -1 Then
        Call Me.LoadCondition
    End If

    BoardNumber = m_UtyBitSetWs.ApmuBoardNum

End Property

#If UB_SNAPSHOT_USE <> 0 Then
'#Pass-Ok UB�̐ݒ�s������肪��������܂ł͕���
Public Sub ClearUBInformation()
'���e:
'   Bit�ݒ������������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    ReDim m_UtyBitStatus(MAX_BIT_CH)

End Sub

'#Pass-Ok UB�̐ݒ�s������肪��������܂ł͕���
Public Property Get UBSnapshot(ByVal ChannelNumber As Long) As Long
'���e:
'   �w��CH�ԍ��̌��݂̐ݒ�l���擾����
'
'�p�����[�^:
'   [ChannelNumber]  In  Long�^:    ��Ԃ��擾����UB��CH�ԍ�
'
'�߂�l:
'   Bit�ݒ�l
'
'���ӎ���:
'   Utility Bits Display�ATERADYNE-API�ɂ��ݒ�l��
'   �擾���ʂɔ��f����܂���
'

    UBSnapshot = m_UtyBitStatus(ChannelNumber)

End Property

'#Pass-Ok UB�̐ݒ�s������肪��������܂ł͕���
Public Property Get UBSnapshotInHex() As String
'���e:
'   0-255CH �eCH�̌��݂̐ݒ�l��Hex�`���Ŏ擾����
'
'�p�����[�^:
'
'�߂�l:
'   HEX�`����0-255ch��Bit�ݒ�l
'
'���ӎ���:
'   Utility Bits Display�ATERADYNE-API���g�p���čs�����ݒ�l��
'   �擾���ʂɔ��f����܂���
'
    
    UBSnapshotInHex = MakeHexMsg(m_UtyBitStatus)

End Property
#End If

'#Pass-Ok
Public Property Set UBSetSht(ByVal UtyBitConditionSht As Worksheet)
'���e:
'   UtilityBit�ݒ�����f�[�^�̂��郏�[�N�V�[�g���w�肷��
'
'�p�����[�^:
'   [UtyBitConditionSht]    In   UtilityBit�����ݒ胏�[�N�V�[�g
'
'�߂�l:
'
'���ӎ���:
'
    
    Set m_UtyBitSetWs = UtyBitConditionSht

End Property

'#Pass-Ok
Public Property Get UBSetShtName() As String
'���e:
'  UtilityBit�ݒ�������[�N�V�[�g��
'
'�p�����[�^:
'
'�߂�l:
'  UtilityBit�ݒ�������[�N�V�[�g��
'
'���ӎ���:
'
    
    UBSetShtName = m_UtyBitSetWs.Name

End Property

'#Pass-Ok
Public Sub LoadCondition()
'���e:
'   �����[�ݒ�����f�[�^�̏��������s���܂��B
'  �i���[�N�V�[�g�ɒ�`���ꂽ�f�[�^��ǂݍ��݂܂��B�j
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'   ���[�N�V�[�g�C�x���g�����s�ł��Ȃ���(��:�u���[�N�|�C���g�ɂ�钆�f��)
'   �ɂă����[�����\�̒l��ύX�����ꍇ�͕ύX�������e�����f����Ȃ��̂Ť
'   �����[�ݒ�@�\�̎g�p�O�ɕK�����̖��߂����s��������\�̒l�̍ēǂݍ��݂�
'   �s���Ă��������
'
'   �������[�����\���[�N�V�[�g���̃f�[�^�̓ǂݍ��݋@�\�����s���Ă��܂��B
'
    
    'APMU Utility Bit setting�̃��[�N�V�[�g���v���p�e�B[UBSetSht]
    '�ɐݒ肳��Ă��Ȃ��Ƃ��ɂ̓G���[
    If m_UtyBitSetWs Is Nothing Then
        Call OutputErrMsg("APMU Utility Bit setting worksheet is nothing" & vbCrLf _
        & "    Please set UBSetSht")
        Exit Sub
    End If
    
    Call m_UtyBitSetWs.Initialize
    
End Sub

'#Pass-Ok
Public Sub SetUB(ByVal AssignName As String, ByVal SetValue As Long)
'���e:
'   AssignName�ɑΉ�����UB���w��l�ɐݒ�ɂ���
'
'�p�����[�^:
'   [AssignName]    In     AssignName
'   [Status]        In     �ݒ�l
'
'�߂�l:
'
'���ӎ���:
'   UB�����\�ɒ�`����Ă��Ȃ����O���w�肵���ꍇ�̓G���[�ƂȂ�܂��B
'   ����Bit����x�ɐݒ肷��ۂ́A[AssignName]���J���}��؂�Ŏw�肵�܂��B
'       ��:  UbObj.SetUB("Name1,Name2,Name3",1)
'

    'APMU Utility Bit setting�̃��[�N�V�[�g���v���p�e�B[UBSetSht]
    '�ɐݒ肳��Ă��Ȃ��Ƃ��ɂ̓G���[
    If m_UtyBitSetWs Is Nothing Then
        Call OutputErrMsg("APMU Utility Bit setting worksheet is nothing" & vbCrLf _
        & "Please set UBSetSht")
        Exit Sub
    End If

    '�A�T�C������UB_CH�̒�`��񂪑��݂��Ȃ��ꍇ��
    '���[�N�V�[�g�̃f�[�^��ǂݍ��݂܂�
    If m_UtyBitSetWs.UBChanMap Is Nothing Then
        Call Me.LoadCondition
    End If
        
    Dim assignNameArr() As String
    Dim tmpAssignName As Variant
        
    '�J���}��؂蕶�����z��`���ɕϊ�
    Call StrToArray(AssignName, assignNameArr)
    
    'UB���V���A���ɐݒ肵�Ă���
    For Each tmpAssignName In assignNameArr
        With TheHdw.APMU.board(m_UtyBitSetWs.ApmuBoardNum)
            On Error GoTo UB_SET_ERR
            .UtilityBit(m_UtyBitSetWs.UBChanMap(CStr(tmpAssignName))) = SetValue
            m_UtyBitStatus(m_UtyBitSetWs.UBChanMap(CStr(tmpAssignName))) = SetValue 'UtyBit�̐ݒ��Ԏ擾�p
            On Error GoTo 0
        End With
    Next tmpAssignName
    
    Exit Sub
    
UB_SET_ERR:
    
    Call OutputErrMsg(tmpAssignName & " Assign UB_CH# is Not Found")
    
End Sub

Public Sub SetUBCondition(ByVal pCsvLabelList As String)
'���e:
'   �w�肵���R���f�B�V�������Œ�`���ꂽ�l��UB���ꊇ�ݒ肷��
'
'�p�����[�^:
'   [pCsvLabelList]  In  String�^:    ��������
'
'�߂�l:
'
'���ӎ���:
'   �����[�����\�ɒ�`����Ă��Ȃ��R���f�B�V�������̂�
'   �w�肵���ꍇ�̓G���[�ƂȂ�܂��B
'
'   �����\�Œl���󗓁i���ݒ�j��CH�́A����̒l�̂܂܂ƂȂ�܂��B
'
'   �����̏������ꊇ�w�肷��ۂɂ́A�J���}��؂�Ŏw�肵�܂��B
'   ��:  UbObj.SetUBCondition("Name1,Name2,Name3")
                       
                       
    '�p�����[�^�V�[�g�̑��݊m�F
    If m_UtyBitSetWs Is Nothing Then
        Call OutputErrMsg("APMU Utility Bit setting worksheet is nothing" & vbCrLf _
        & "Please set UBSetSht")
        Exit Sub
    End If
                       
    '�p�����[�^�L���b�V�����݊m�F
    If m_UtyBitSetWs.UBCondition Is Nothing Then
        Call Me.LoadCondition
    End If
                
    '�p�����[�^�L���b�V���Q�Ǝ��o��
    Dim UBParamsCache As Collection
    Set UBParamsCache = m_UtyBitSetWs.UBCondition
                
    '���x�����X�g��z��ɓW�J
    If Len(pCsvLabelList) = 0 Then
        '���X�g�󕶎���`�F�b�N
        Call OutputErrMsg("Label name is empty")
    End If
    Dim LabelList() As String
    LabelList = Split(pCsvLabelList, ",", , vbTextCompare)
    Dim LabelCount As Long
    For LabelCount = 0 To UBound(LabelList) Step 1
        If Len(LabelList(LabelCount)) = 0 Then
            '���X�g���̋󕶎���`�F�b�N
            Call OutputErrMsg("The empty label is included")
        End If
    Next
                                    
    '�}�[�W���x�������(�z����̕������S����)
    Dim Mergelabel As String
    Mergelabel = Join(LabelList, Empty)
    
    '�}�[�W���x���Őݒ�f�[�^���擾
    Dim MergeParams As Collection
    On Error Resume Next
    Set MergeParams = UBParamsCache(Mergelabel)
    On Error GoTo 0
    
    '�}�[�W�p�����[�^�L���b�V�����ݔ���
    If MergeParams Is Nothing Then
        '�}�[�W�p�����[�^�쐬����
        Set MergeParams = MakeMergeParams(LabelList, UBParamsCache)
        '�L���b�V���Ƀ}�[�W�p�����[�^�o�^
        Call UBParamsCache.Add(MergeParams, Mergelabel)
    End If
    
    'UB�ݒ蓮��
    Dim UBParam As Variant
    With TheHdw.APMU.board(m_UtyBitSetWs.ApmuBoardNum)
        On Error GoTo UB_SET_ERR
        For Each UBParam In MergeParams
            .UtilityBit(CLng(UBParam(0))) = CLng(UBParam(1))
            m_UtyBitStatus(CLng(UBParam(0))) = CLng(UBParam(1))
        Next
        On Error GoTo 0
    End With
    
    '�ŏI�ݒ薼�̊i�[(���������o�����͖���)
    m_LastConditionName = Mergelabel
    
    Exit Sub
    
'CH�ݒ菈���G���[�n���h��
UB_SET_ERR:
    Call OutputErrMsg("Could not be set CH=" & UBParam(0))

End Sub

'#Pass-Ok
Public Sub ResetUtilityBits()
'���e:
'   �r�b�g�̐ݒ�l�̃��Z�b�g(�S�r�b�g��OFF�ɂ���)
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    
    Dim apmuCh As Long
        
    '�{�[�h�ԍ��̒�`��񂪑��݂��Ȃ��ꍇ��
    '���[�N�V�[�g�̃f�[�^��ǂݍ��݂܂�
    If m_UtyBitSetWs.ApmuBoardNum = -1 Then
        Call Me.LoadCondition
    End If
        
    '���ׂĂ�APMU CH��Bit�ݒ�l��0�ɐݒ肷��
    For apmuCh = 0 To MAX_BIT_CH Step 1
        With TheHdw.APMU.board(m_UtyBitSetWs.ApmuBoardNum)
            .UtilityBit(apmuCh) = 0
        End With
        m_UtyBitStatus(apmuCh) = 0
    Next apmuCh

End Sub

'#Pass-Ok
Private Sub Class_Initialize()
    
    Set m_UtyBitSetWs = Nothing
    m_LastConditionName = ""
    ReDim m_UtyBitStatus(MAX_BIT_CH)

End Sub

'#Pass-Ok
Private Sub StrToArray(ByVal astrInputStr As String, ByRef alngOutputArray() As String)
'�J���}��؂蕶�����z��`���ɕϊ�����  CUBUtilBitCtrl���R�s�[
    
    Dim ret As Long
    Dim i As Long

    Erase alngOutputArray()

    Do
        ret = InStr(1, astrInputStr, ",")

        If ret = 0 Then
            ReDim Preserve alngOutputArray(i)
            alngOutputArray(i) = astrInputStr
            Exit Do
        End If
        ReDim Preserve alngOutputArray(i)
        alngOutputArray(i) = Left(astrInputStr, ret - 1)
        astrInputStr = Right(astrInputStr, Len(astrInputStr) - ret)
        i = i + 1
    Loop

End Sub

'#Pass-Ok
Private Function MakeHexMsg(bitArr() As Long) As String
'�z��Ɋi�[����Ă���Bit�l���Hex�o�͌`���̃��b�Z�[�W���쐬

    Dim lpCnt As Long
    Dim strBin As String
    Dim chkCnt As Long
    Dim dispMsg As String

    chkCnt = UBound(bitArr) - 3
    strBin = ""
    dispMsg = ""
        
    For lpCnt = UBound(bitArr) To 0 Step -1
        strBin = strBin & bitArr(lpCnt)
        If lpCnt = chkCnt Then
            dispMsg = dispMsg & Hex$(BinToDec(strBin))
            strBin = ""
            chkCnt = chkCnt - 4
        End If
    Next lpCnt

    MakeHexMsg = dispMsg

End Function

'#Pass-Ok
Private Function BinToDec(BinValue As String) As Variant
'Bit�l�������10�i���ɕϊ�����
    
    Dim lpCnt As Long
    Dim StrVal As String
    Dim Cursor As Long
  
    '���l�łȂ����̂������Ă�����ϊ����Ȃ�
    If IsNumeric(BinValue) = False Then
        BinToDec = Empty
        Exit Function
    End If
  
    Cursor = 0
    
    For lpCnt = Len(BinValue) To 1 Step -1
        Select Case Mid$(BinValue, lpCnt, 1)
            '0�̂Ƃ�
            Case "0"
            '1�̂Ƃ�
            Case "1"
                BinToDec = BinToDec + (2 ^ Cursor)
            '0/1�ȊO�̐��l���w�肳�ꂽ�ꍇ
            Case Else
                BinToDec = Empty
                Exit Function
        End Select
        Cursor = Cursor + 1
    Next lpCnt

End Function

'�}�[�W�p�����[�^�쐬�֐�
Private Function MakeMergeParams(ByRef pLabelList() As String, ByRef pParams As Collection) As Collection
    Dim LabelCount As Long
    Dim UBParams As Collection
    Dim UBParam As Variant
    Set MakeMergeParams = New Collection
    For LabelCount = 0 To UBound(pLabelList) Step 1
        On Error GoTo PARAM_READ_ERROR
        Set UBParams = pParams.Item(pLabelList(LabelCount))
        On Error GoTo 0
        For Each UBParam In UBParams
            On Error Resume Next
            Call MakeMergeParams.Remove(CStr(UBParam(0)))
            On Error GoTo 0
            Call MakeMergeParams.Add(UBParam, CStr(UBParam(0)))
        Next UBParam
    Next LabelCount
    Exit Function

'�p�����[�^����`�G���[
PARAM_READ_ERROR:
    Set MakeMergeParams = Nothing
    Call OutputErrMsg(pLabelList(LabelCount) & " Condition is Not Found")
End Function