Attribute VB_Name = "XLibVISUtility"
'�T�v:
'   VIS�N���X�Ŏg�p���郆�[�e�B���e�B�֐��Q
'
'�ړI:
'   �eVIS�N���X���ʂɎg�p���郆�[�e�B���e�B�֐����܂Ƃ߂�
'
'�쐬��:
'   SLSI����
'
'   XlibSTD_CommonDCMod_V01���́A���ʏ�����؂�o��
'   �G���[���b�Z�[�W�\���p�T�u���[�`���ǉ�
'
'
'Code Checked
'Comment Checked
'

Option Explicit

Private Const ALL_SITE = -1

'#Pass
Public Function ConvertVariableToArray(ByRef DstVar As Variant) As Boolean
'���e:
'   �ϐ���v�f��������Site���̔z��ϐ��ɕϊ�����
'
'�p�����[�^:
'    [DstVar]    In   �ϊ��Ώەϐ�
'    [DstVar]    Out  �ϊ���z��ϐ�
'
'�߂�l:
'   �X�e�[�^�X�i����I��=True�j
'
'���ӎ���:
'
    Dim VarArray() As Double
    Dim site As Long

    If IsArray(DstVar) Then
        If UBound(DstVar) <> CountExistSite Then
            ConvertVariableToArray = False
        Else
            ConvertVariableToArray = True
        End If
    Else
        ReDim VarArray(CountExistSite)

        For site = 0 To UBound(VarArray)
            VarArray(site) = DstVar
        Next site
        DstVar = VarArray
        ConvertVariableToArray = True
    End If

End Function

'#Pass
Public Sub GetChanList(ByVal PinList As String, ByVal site As Long, ByVal chanType As chtype, ByRef retChannels() As Long)
'���e:
'   �Ώۃs����Channel�ԍ����擾����
'
'�p�����[�^:
'    [PinList]        In    �Ώۃs�����X�g
'    [site]           In    Site�ԍ�
'    [chanType]       In    ChannelType
'    [retChannels()]  Out   �擾����Channel�ԍ�
'
'�߂�l:
'
'���ӎ���:
'
    Dim ChanNum As Long
    Dim siteNum As Long
    Dim errMsg As String

    Call TheExec.DataManager.GetChanList(PinList, site, chanType, retChannels, ChanNum, siteNum, errMsg)
    If errMsg <> "" Then
        Call OutputErrMsg(errMsg & " (at GetChanList)")
    End If
End Sub

'#Pass
Public Sub GetActiveChanList(ByVal PinList As String, ByVal chanType As chtype, ByRef retChannels() As Long)
'���e:
'   �I������Ă���Site�̑Ώۃs����Channel�ԍ����擾����
'
'�p�����[�^:
'    [PinList]         In    �Ώۃs�����X�g
'    [chanType]        In    ChannelType
'    [retChannels()]   Out   �擾����Channel�ԍ�
'
'�߂�l:
'
'���ӎ���:
'
    Dim ChanNum As Long
    Dim siteNum As Long
    Dim errMsg As String

    Call TheExec.DataManager.GetChanListForSelectedSites(PinList, chanType, retChannels, ChanNum, siteNum, errMsg)
    If errMsg <> "" Then
        Call OutputErrMsg(errMsg & " (at GetActiveChanList)")
    End If
End Sub

'#Pass
Public Function CountExistSite() As Long
'���e:
'   ���݂���Site�����擾����
'
'�p�����[�^:
'
'�߂�l:
'   ����Site��
'
'���ӎ���:
'
    CountExistSite = TheExec.sites.ExistingCount - 1

End Function

'#Pass
Public Function CountActiveSite() As Long
'���e:
'   Active Site�����擾����
'
'�p�����[�^:
'
'�߂�l:
'   ActiveSite��
'
'���ӎ���:
'   �V���A��LOOP���͖߂�l=1�ƂȂ�
'
    With TheExec.sites
        If .InSerialLoop Then
            CountActiveSite = 1
        Else
            CountActiveSite = .ActiveCount
        End If
    End With
    
End Function

'#Pass
Public Function IsActiveSite(ByVal site As Long) As Boolean
'���e:
'   Site��Active��Ԃł��邩�m�F����
'
'�p�����[�^:
'    [site]     In    �m�FSite�ԍ�
'
'�߂�l:
'   �m�F����(Active���=True)
'
'���ӎ���:
'
    IsActiveSite = TheExec.sites.site(site).Selected

End Function

'#Pass
Public Function GetChanType(ByVal PinList As String) As chtype
'���e:
'   �w��s����ChannelType���擾����
'
'�p�����[�^:
'    [PinList]    In    �m�F�Ώۃs��
'
'�߂�l:
'   �m�F�Ώۃs����ChannelType
'
'���ӎ���:
'   �m�F�Ώۃs�����X�g�ɈقȂ�ChannelType��
'   Pin���w�肳�ꂽ�ꍇ�́A�߂�l��chUnk�ƂȂ�
'
    Dim chanType As chtype
    Dim pinArr() As String
    Dim pinNum As Long
    Dim i As Long

    With TheExec.DataManager
        chanType = .chanType(PinList)

        If chanType = chUnk Then
            On Error GoTo INVALID_PINLIST
            Call .DecomposePinList(PinList, pinArr, pinNum)
            On Error GoTo -1

            chanType = .chanType(pinArr(0))
            For i = 0 To pinNum - 1
                If chanType <> .chanType(pinArr(i)) Then
                    chanType = chUnk
                    Exit For
                End If
            Next i
        End If
    End With

    GetChanType = chanType
    Exit Function

INVALID_PINLIST:
    GetChanType = chUnk

End Function

'#Pass
Public Sub SeparatePinList(ByVal PinList As String, ByRef retPinNames() As String)
'���e:
'   �s�����X�g�̃s������z��`���ɕ���
'
'�p�����[�^:
'    [PinList]          In    �����Ώۃs��
'    [retPinNames()]    Out   ������s��
'
'�߂�l:
'
'���ӎ���:
'
    Dim pinNum As Long
    Call TheExec.DataManager.DecomposePinList(PinList, retPinNames, pinNum)
    
End Sub

'#Pass
Public Function CreateEmpty2DArray(ByVal Dim1 As Long, ByVal Dim2 As Long) As Variant
'���e:
'   �w��T�C�Y�̓񎟌��z��ϐ����쐬
'
'�p�����[�^:
'    [Dim1]    In   �z�񎟌�1�̐�
'    [Dim2]    In   �z�񎟌�2�̐�
'
'�߂�l:
'   Dim1�~Dim2�̃T�C�Y��2�����z��
'
'���ӎ���:
'   �z��̒l��0
'
    Dim ret2DArr() As Variant
    Dim tmp() As Double
    Dim i As Long
    
    ReDim ret2DArr(Dim1)
    ReDim tmp(Dim2)
    
    For i = 0 To UBound(ret2DArr)
        ret2DArr(i) = tmp
    Next i
    
    CreateEmpty2DArray = ret2DArr
    
End Function

'#Pass
Public Function IsValidSite(ByVal site As Long) As Boolean
'���e:
'   �T�C�g���L���ȃT�C�g���m�F
'
'�p�����[�^:
'    [site]    In   �m�F����T�C�g�ԍ�
'
'�߂�l:
'   �m�F����(�L��Site = True)
'
'���ӎ���:
'
    If site = ALL_SITE Then
        IsValidSite = True
    ElseIf 0 <= site And site <= CountExistSite Then
        IsValidSite = True
    Else
        IsValidSite = False
    End If

End Function

'#Pass
Public Function CreateLimit(ByVal dstVal As Variant, ByVal loLim As Double, ByVal hiLim As Double) As Variant
'���e:
'   Limit�l������l�Ɖ����l��萶��
'
'�p�����[�^:
'    [dstVal]    In   �ݒ�l
'    [loLim]     In   �����l
'    [hiLim]     In   ����l
'
'�߂�l:
'   Limit�l
'
'���ӎ���:
'
    Dim i As Long

    If IsArray(dstVal) Then
        For i = 0 To UBound(dstVal)
            If dstVal(i) < loLim Then dstVal(i) = loLim
            If dstVal(i) > hiLim Then dstVal(i) = hiLim
        Next i
    Else
        If dstVal < loLim Then dstVal = loLim
        If dstVal > hiLim Then dstVal = hiLim
    End If

    CreateLimit = dstVal

End Function

'#Pass
Public Function ReadMultiResult(ByVal PinName As String, ByRef retResult() As Double, ByRef Results As Collection) As Boolean
'���e:
'    �s�������L�[�ɃR���N�V�����̗v�f�����o��
'
'�p�����[�^:
'    [PinName]        In   �R���N�V�����̃L�[�ƂȂ�s����
'    [retResult()]    Out  �R���N�V����������o�����l
'    [results]        In   �v�f�����o���R���N�V����
'
'�߂�l:
'   �X�e�[�^�X�i����I��=True�j
'
'���ӎ���:
'
    Dim site As Long
    Dim result As Variant

    On Error GoTo NOT_FOUND
    result = Results(PinName)
    On Error GoTo 0
    
    For site = 0 To CountExistSite
        retResult(site) = result(site)
    Next site

    ReadMultiResult = True
    Exit Function
    
NOT_FOUND:
    ReadMultiResult = False
    
End Function

'#Pass
Public Function IsGangPinlist(ByVal PinList As String, ByVal chtype As chtype) As Boolean
'���e:
'    '�w�肳�ꂽ�s�����X�g�ɃM�����O�s�����܂܂�Ă��邩�m�F
'
'�p�����[�^:
'    [PinName]       In   �m�F���s��PinList
'    [chtype]        In   �ΏۂƂȂ�{�[�h��ChannelType
'
'�߂�l:
'   �m�F���ʁi�M�����O�s�����܂܂�Ă���=True�j
'
'���ӎ���:
'
    Dim pinNames() As String
    Dim Channels() As Long
    
    Call GetChanList(PinList, ALL_SITE, chtype, Channels)
    Call SeparatePinList(PinList, pinNames)
    
    If (UBound(pinNames) + 1) * (CountExistSite + 1) <> UBound(Channels) + 1 Then
        IsGangPinlist = True  '�M�����O�s��������
    Else
        IsGangPinlist = False '�M�����O�s���͂Ȃ�
    End If

End Function

Public Function IsGangMultiPinlist(ByVal PinList As String) As Boolean
'���e:
'   �w�肳�ꂽ�s�����X�g��PinGp�����ׂ�GANG�ڑ��p��PinGp���m�F����
'
'�p�����[�^:
'   [PinList]   In  �m�F���s��PinList
'
'�߂�l:
'   �m�F���ʁi���ׂ�GANG�ڑ��p��PinGp�ł���=True�j
'
'���ӎ���:
'
    Dim pinListArr() As String
    '�s���O���[�v��W�J�����ɃJ���}��؂�`���̔z��ɕϊ�
    Call ConvertStrPinListToArrayPinList(PinList, pinListArr)

    Dim tmpPinGp As Variant
    For Each tmpPinGp In pinListArr
        If IsGangPinlist(tmpPinGp, GetChanType(tmpPinGp)) <> True Then
            IsGangMultiPinlist = False
'            Call MsgBox(tmpPinGp & " ��GANG�ڑ��p��PinGp�ł͂���܂���")
            Exit Function
        End If
    Next tmpPinGp

    IsGangMultiPinlist = True

End Function

Public Sub ConvertStrPinListToArrayPinList(ByVal StrPinList As String, ByRef ArrayPinList() As String)
'���e:
'   �J���}��؂蕶����`���̃s�����X�g���A�z��`���̃s�����X�g�ɕϊ�
'
'�p�����[�^:
'   [StrPinList]    In   �ϊ��Ώۂ̕�����s�����X�g
'   [ArrayPinList]  Out  �ϊ���̔z��`���s�����X�g
'
'�߂�l:
'
'���ӎ���:
'   PinList��PinGp���w�肵���Ƃ��ɂ́APinGp�̃����o�[�͓W�J���܂���
'   �`���ϊ��݂̂ƂȂ�܂��B
'
    Dim ret As Long
    Dim i As Long
    
    Erase ArrayPinList()

    Do
        ret = InStr(1, StrPinList, ",")
        If ret = 0 Then
            ReDim Preserve ArrayPinList(i)
            ArrayPinList(i) = StrPinList
            Exit Do
        End If
        ReDim Preserve ArrayPinList(i)
        ArrayPinList(i) = Left(StrPinList, ret - 1)
        StrPinList = Right(StrPinList, Len(StrPinList) - ret)
        i = i + 1
    Loop

End Sub
