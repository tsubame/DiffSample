Attribute VB_Name = "XLibVISErrorCheck"
'�T�v:
'   VIS�N���X�Ŏg�p����`�F�b�N�p�̊֐��Q
'
'�ړI:
'   �eVIS�N���X���ʂɎg�p����`�F�b�N�p�֐����܂Ƃ߂�
'
'�쐬��:
'   SLSI����
'
'   XlibSTD_CommonDCMod_V01���́ADC�֘A�G���[�`�F�b�N
'   ������؂�o���܂Ƃ߂����́B
'
'
'Code Checked
'Comment Checked
'

Option Explicit

Private Const ALL_SITE = -1

'#Pass
Public Function CheckPinList(ByVal PinList As String, ByVal chanType As chtype, ByVal FunctionName As String) As Boolean
'���e:
'    �w��s�����Ώ�ChannelType�Ƃ��Ē�`����Ă��邩���m�F
'
'�p�����[�^:
'    [PinList]        In  �m�F�Ώۃs�����X�g�B
'    [chanType]       In  �m�F�`�����l���^�C�v
'    [FunctionName]   In  �Ăяo�����֐���
'
'�߂�l:
'   �m�F����
'
'���ӎ���:
'
    If GetChanType(PinList) <> chanType Then
        Call OutputErrMsg(PinList & " is Invalid Channel Type at " & FunctionName & "().")
        CheckPinList = False
    Else
        CheckPinList = True
    End If
    
End Function

'#Pass
Public Function CheckSinglePins(ByVal PinList As String, ByVal chanType As chtype, ByVal FunctionName As String) As Boolean
'���e:
'    �w��s��������ChannelType�̕����m�F
'
'�p�����[�^:
'    [PinList]        In  �m�F�Ώۃs�����X�g�B
'    [chanType]       In  �m�F�`�����l���^�C�v
'    [FunctionName]   In  �Ăяo�����֐���
'
'�߂�l:
'   �m�F����
'
'���ӎ���:
'
    Dim Channels() As Long
    Dim ChanNum As Long
    Dim siteNum As Long
    Dim errMsg As String
    Call TheExec.DataManager.GetChanList(PinList, ALL_SITE, chanType, Channels, ChanNum, siteNum, errMsg)
    
    If ChanNum <> siteNum Then
        Call OutputErrMsg("Don't Support Multi Pins at " & FunctionName & "().")
        CheckSinglePins = False
    Else
        CheckSinglePins = True
    End If
    
End Function

'#Pass
Public Function CheckForceVariantValue(ByVal ForceVal As Variant, ByVal loLim As Double, ByVal hiLim As Double, ByVal FunctionName As String) As Boolean
'���e:
'    Force�l�������l�Ə���l�̊Ԃ̒l�ł��邩���m�F
'
'�p�����[�^:
'    [ForceVal]       In  Force�l
'    [loLim]          In  �����l
'    [hiLim]          In  ����l
'    [FunctionName]   In  �Ăяo�����֐���
'
'�߂�l:
'   �m�F����
'
'���ӎ���:
'
    Dim site As Long
    
    If IsArray(ForceVal) Then
        If UBound(ForceVal) <> CountExistSite Then
            Call OutputErrMsg("ForceVal is Invalid Site Array at " & FunctionName & "().")
            CheckForceVariantValue = False
            Exit Function
        End If
        
        For site = 0 To CountExistSite
            If (ForceVal(site) < loLim Or hiLim < ForceVal(site)) Then
                Call OutputErrMsg("ForceVal(= " & ForceVal(site) & ") must be between " & loLim & " and " & hiLim & " at " & FunctionName & "().")
                CheckForceVariantValue = False
                Exit Function
            End If
        Next site
        
    Else
        If (ForceVal < loLim Or hiLim < ForceVal) Then
            Call OutputErrMsg("ForceVal(= " & ForceVal & ") must be between " & loLim & " and " & hiLim & " at " & FunctionName & "().")
            CheckForceVariantValue = False
            Exit Function
        End If
    End If
    
    CheckForceVariantValue = True
    
End Function

'#Pass
Public Function CheckClampValue(ByVal clampVal As Double, ByVal loLim As Double, ByVal hiLim As Double, ByVal FunctionName As String) As Boolean
'���e:
'    Clamp�l���A�����l�Ə���l�̊Ԃ̒l�ł��邩���m�F
'
'�p�����[�^:
'    [clampVal]       In  Clamp�l
'    [loLim]          In  �����l
'    [hiLim]          In  ����l
'    [FunctionName]   In  �G���[���b�Z�[�W�ɕ\������Ăяo�����֐���
'
'�߂�l:
'   �m�F����
'
'���ӎ���:
'
    If (clampVal < loLim Or hiLim < clampVal) Then
        Call OutputErrMsg("ClampVal(= " & clampVal & ") must be between " & loLim & " and " & hiLim & " at " & FunctionName & "().")
        CheckClampValue = False
    Else
        CheckClampValue = True
    End If

End Function

'#Pass
Public Function IsExistSite(ByVal site As Long, ByVal FunctionName As String) As Boolean
'���e:
'    �w��ԍ���Site�����݂��邩�m�F
'
'�p�����[�^:
'    [site]           In  �m�FSite�ԍ�
'    [FunctionName]   In  �G���[���b�Z�[�W�ɕ\������Ăяo�����֐���
'
'�߂�l:
'   �m�F����
'
'���ӎ���:
'
    If site <> ALL_SITE And (site < 0 Or CountExistSite < site) Then
        Call OutputErrMsg("Site(= " & site & ") must be -1 or between 0 and " & CountExistSite & " at " & FunctionName & "().")
        IsExistSite = False
    Else
        IsExistSite = True
    End If

End Function

'#Pass
Public Function CheckResultArray(ByRef retResult() As Double, ByVal FunctionName As String) As Boolean
'���e:
'    ���ʊi�[�p�̔z��ϐ��̗v�f�������݂���Site���ƍ����Ă��邩�m�F
'
'�p�����[�^:
'    [retResult()]      In  ���ʊi�[�p�̔z��ϐ�
'    [FunctionName]   In  �G���[���b�Z�[�W�ɕ\������Ăяo�����֐���
'
'�߂�l:
'   �m�F����
'
'���ӎ���:
'
    If UBound(retResult) <> CountExistSite Then
        Call OutputErrMsg("Elements of retResult() is Different from Number of Site at " & FunctionName & "().")
        CheckResultArray = False
    Else
        CheckResultArray = True
    End If

End Function

'#Pass
Public Function CheckAvgNum(ByVal avgNum As Long, ByVal FunctionName As String) As Boolean
'���e:
'    ���[�^���[�h���̃A�x���[�W�񐔒l��1�����łȂ����Ƃ��m�F
'
'�p�����[�^:
'    [avgNum]         In  �A�x���[�W�񐔒l
'    [FunctionName]   In  �G���[���b�Z�[�W�ɕ\������Ăяo�����֐���
'
'�߂�l:
'   �m�F����
'
'���ӎ���:
'

    If avgNum < 1 Then
        Call OutputErrMsg("AvgNum must be 1 or More at " & FunctionName & "().")
        CheckAvgNum = False
    Else
        CheckAvgNum = True
    End If
    
End Function

Public Function CheckFailSiteExists(ByVal FunctionName As String) As Boolean
'���e:
'    ���݂���T�C�g��FAIL�T�C�g�����邩�ǂ������m�F����
'
'�p�����[�^:
'    [FunctionName]   In  �G���[���b�Z�[�W�ɕ\������Ăяo�����֐���
'
'�߂�l:
'   �m�F���ʁiTrue�FFAIL�T�C�g�����݂���j
'
'���ӎ���:
'

    With TheExec.sites
        If .ExistingCount <> .ActiveCount Then
            CheckFailSiteExists = True
            Call MsgBox("ExistingSites=" & .ExistingCount & _
            ", ActiveSites=" & .ActiveCount & _
            " FAIL site exists.  at " & FunctionName & "()", vbCritical, "CheckFailSiteExists")
            Exit Function
        End If
    End With

    CheckFailSiteExists = False
    
End Function

