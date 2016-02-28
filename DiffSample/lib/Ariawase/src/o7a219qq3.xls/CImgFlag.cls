VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �t���O�̊Ǘ�
'
'�ړI:
'   �t���O�𖼑O�ŊǗ�����B
'
'�쐬��:
'   0145184004
'
' 2009/06/09 D.Maruyama �t���O�̃N���A���ŏ����ɂ��邽�߂ɑ啝�ύX
' 2010/11/12 t.takase   EeeJOB V3.0 idpColorType�̎w���C�ӂ̕����ŉ\�ɂ��邽�ߕύX

Option Explicit

Implements IImgFlag

'�r�b�g�̏��
Private Enum FlagUsingState
    Uninitialized   '������ԁA�@�Ȃɂ������Ă��邩�s��
    INITIALIZED     '�������ς�(�N���A�ς�)
    Used            '�g�p����Ă���
    NotUsed         '�g�p���I����ĊJ�����ꂽ
End Enum

Private Type sFlagState
    Count As Long
    flagName(1 To 16) As String
    FlagState(1 To 16) As FlagUsingState
End Type

Private m_FlgPlane As CImgPlane
Private m_FlgState As sFlagState
Private m_Group As String

'#FlagExpansion
Private Property Get IImgFlag_Count() As Long
    IImgFlag_Count = FlgCount_
End Property

'#FlagExpansion
Private Property Get IImgFlag_Group() As String
    IImgFlag_Group = m_Group
End Property

Private Property Let IImgFlag_Group(ByVal RHS As String)
    m_Group = RHS
End Property

Private Property Set IImgFlag_FlgPlane(ByRef pPlane As CImgPlane)
    If Not m_FlgPlane Is Nothing Then Exit Property
    
    Set m_FlgPlane = pPlane
    If m_FlgPlane.Comment = "" Then
        m_FlgPlane.Comment = "Shared Flag Plane"
    End If
End Property

Private Property Get IImgFlag_FlgPlane() As CImgPlane
    Set IImgFlag_FlgPlane = m_FlgPlane
End Property

Private Sub IImgFlag_Clear()
'���e:
'   �o�^�����t���O�̈ꊇ�N���A
    
    Call IImgFlag_ResetFlagState
    Set m_FlgPlane = Nothing

End Sub

Private Property Get IImgFlag_FlagBit(ByVal pName As String) As Long
'���e:
'   �w�肵�����O�̃t���O�ʒu���擾
'
'[pName]       IN String�^:     ���O�w��
'
'���l:
'   �w�肵�����O�̃t���O�����݂��Ȃ��ꍇ0��Ԃ��B
'

    Dim i As Long
    
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    For i = minBit To maxBit
        If pName = m_FlgState.flagName(i) Then
            IImgFlag_FlagBit = i
            Exit Property
        End If
    Next i
    
    IImgFlag_FlagBit = 0

End Property

'#FlagExpansion
Private Function IImgFlag_SetFlagBit(ByVal pName As String) As Long
'���e:
'   �w�肵�����O�̃t���O���Z�b�g����B
'   �t���O���Z�b�g���ꂽ�ꍇ�A���̃r�b�g��0�N���A���t���O�ԍ���Ԃ��B
'   �����̖��O�̏ꍇ�͂��̔ԍ����A�t���O�����܂��Ă���ꍇ��0��Ԃ��B
'
'[pName]       IN String�^:     ���O�w��
'
'���l:
'   �t���O�̋󂫂��Ȃ��ꍇ0��Ԃ��B
'

    IImgFlag_SetFlagBit = IImgFlag_FlagBit(pName)
    If IImgFlag_SetFlagBit <> 0 Then Exit Function
    
    Dim oldPmd As CImgPmdInfo
    

    '����0�̂Ƃ��̓t���O�v�������Ȃ��̂Ŋm�ۂ��ɂ���
    If FlgCount_ = 0 And m_FlgPlane Is Nothing Then
    
'2009/09/03 D.Maruayama �[���N���A�͈�񂾂��ɂ���
'        Set m_FlgPlane = TheIDP.PlaneManager(m_Group).GetFreePlane(idpDepthS16, True)
        Set m_FlgPlane = TheIDP.PlaneManager(m_Group).GetFreePlane(idpDepthS16)
        m_FlgPlane.Comment = "Shared Flag Plane"
        Call m_FlgPlane.SetPMD(m_FlgPlane.BasePMD).WritePixel(0)
        Call InitializeFlagState
                
    End If
    
    '�������ς݁i��x���g���Ă��Ȃ��j�t���O�r�b�g���Ƃ��Ă���
    IImgFlag_SetFlagBit = GetInitializedFlag
    
    
    '�Ƃ��Ă���Ȃ�������g�p�ς݃t���O���N���A���āA�������ς݃r�b�g���Ƃ��Ă���
    If IImgFlag_SetFlagBit = 0 Then
    
        Set oldPmd = m_FlgPlane.CurrentPMD
        
        Call RecoverFlagPlane(m_FlgPlane)
        
        If oldPmd.Name = m_FlgPlane.Manager.VariablePMDName Then
            Call m_FlgPlane.SetPMD(oldPmd)
        Else
            Call m_FlgPlane.SetPMD(oldPmd.Name)
        End If
        
        IImgFlag_SetFlagBit = GetInitializedFlag
        
    End If
    
    
    '�����܂łɂƂ��Ă��ꂽ���Ԃ�ύX����,���߂Ȃ�G���[��ʒm
    If IImgFlag_SetFlagBit <> 0 Then
        Call SetUsedState(IImgFlag_SetFlagBit, pName)
        Exit Function
    End If
        
    Call TheError.Raise(999, "SetFlagBit", "There's no bit for the flag")
End Function

'#FlagExpansion
Private Sub IImgFlag_RemoveFlagBit(ByVal pName As String)
'���e:
'   �w�肵�����O�̃t���O���폜����
'
'[pName]       IN String�^:     ���O�w��
'

    Dim i As Long
        
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    For i = minBit To maxBit
        If m_FlgState.flagName(i) = pName Then
            Call SetNotUsedState(i)
'���p��0�ɂȂ��Ă������J���͂��Ȃ�
'            If FlgCount_ = 0 Then
'                Set m_FlgPlane = Nothing
'            End If
            Exit Sub
        End If
    Next i

End Sub

'#FlagExpansion
Private Function GetFlgBit(ByVal pName As String) As Long
    '���g�̒��Ŋm�F
    Call OtherGroupGetFlag(pName)
    
    GetFlgBit = IImgFlag_FlagBit(pName)
    '���҂̃t���O�v���[���ɖ₢���킹
    If GetFlgBit = 0 Then
        GetFlgBit = m_FlgPlane.Manager.GetSharedFlagPlane(pName).FlagBit(pName)
    End If
    '�ǂ��ɂ�����
    If GetFlgBit = 0 Then Call Err.Raise(999, "", pName & " is not exist")
End Function



'''' 0000130343
Private Function OtherGroupGetFlag(ByVal pFlgName As String)

    Dim i As Integer
    Dim sPlane As CImgPlane
    Dim resPlane As CImgPlane
    Dim resBit As Integer
    Dim dstPlane As CImgPlane
    Dim dstBit As Integer

    ' Debug matsu
    Dim resZone As String
    Dim dstZone As String

    '' pFlgName ��NULL�̏ꍇ�́A���̏����͕s�v
    If pFlgName = "" Then Exit Function

    '' �����t���O�����݂��Ă��邱��
    If IImgFlag_FlagBit(pFlgName) <> 0 Then
        Exit Function
    Else
    '' �t���O�����݂��Ă��Ȃ��ꍇ�́A�܂��A�ق��̃v���[���O���[�v�̃t���O���`�F�b�N����
        For i = 1 To TheIDP.PlaneManagerCount
            If TheIDP.PlaneManager(i).GetSharedFlagPlane(pFlgName).FlagBit(pFlgName) <> 0 Then
                '' ���ʂ̃t���O�r�b�g���m�ۂ���B
                Call TheIDP.PlaneManager(m_Group).GetSharedFlagPlane(pFlgName).SetFlagBit(pFlgName)
'                Call IImgFlag_SetFlagBit(pFlgName)
                '' �R�s�[�O��̃t���O�v���[�����w�肷��B
                Set dstPlane = TheIDP.PlaneManager(m_Group).GetSharedFlagPlane(pFlgName).FlgPlane
                Set resPlane = TheIDP.PlaneManager(i).GetSharedFlagPlane(pFlgName).FlgPlane
                '' ZONE���̕ۊ�
                resZone = resPlane.CurrentPMD.Name
                dstZone = dstPlane.CurrentPMD.Name
                
                '' ZONE��Bit�Y���Ȃǂ̃`�F�b�N������̂ł���΂����ɋL�q����B
                dstBit = TheIDP.PlaneManager(m_Group).GetSharedFlagPlane(pFlgName).FlagBit(pFlgName)
                resBit = TheIDP.PlaneManager(i).GetSharedFlagPlane(pFlgName).FlagBit(pFlgName)
                '' �Ώۂ�Bit�̂݃R�s�[����B
                Call dstPlane.SetPMD(dstPlane.BasePMD.Name)
                Call resPlane.SetPMD(resPlane.BasePMD.Name)
                Call dstPlane.LOr(resPlane, resPlane, EEE_COLOR_FLAT, , , dstBit, resBit, resBit)
                
                '' ZONE����FULL������Ƃ̏�Ԃɖ߂��B
                Call dstPlane.SetPMD(dstZone)
                Call resPlane.SetPMD(resZone)
                
                Exit Function
            End If
        Next i
    End If
    '' �����ɂ���̂́A�{���ɂǂ��ł��t���O�𐶐����Ă��Ȃ��P�[�X�̂�

End Function



'#FlagExpansion
Private Sub IImgFlag_LNot(ByVal pDstName As String, ByVal pSrcName As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'���e:
'   pSrcName�̃r�b�g��Not���Z�̌��ʂ�pDstName�̃r�b�g�ɓ����B
'   pDstName���o�^����Ă��Ȃ��ꍇ�A�V���ɓo�^����B
'   ���ɓo�^����Ă���ꍇ�́A���̃r�b�g�ɓ����B
'
'[pSrcName]     IN String�^:        �f�[�^���̖��O
'[pDstName]     IN String�^:        ���ʂ̖��O
'[pColor]       IN IdpColorType�^:  �F�w��
'
'���l:
'   pSrcName��pDstName������ł��\�B
'   ���O��SetPMD�Ŏw�肳�ꂽ�]�[�����Ώ�
'
    '�G���[����
    If m_FlgPlane Is Nothing Then
        Call TheError.Raise(999, "LNot", "There's no flag.")
    End If

On Error GoTo NO_FLAG
    Dim SrcFlagPlane As CImgPlane
    Set SrcFlagPlane = GetInputFlagPlane(pSrcName)
    With m_FlgPlane
        If IImgFlag_FlagBit(pDstName) = 0 Then Call IImgFlag_SetFlagBit(pDstName)
        Call .LNot(SrcFlagPlane, pColor, , GetFlgBit(pDstName), GetFlgBit(pSrcName))
    End With

    Exit Sub
    
NO_FLAG:
    Call TheError.Raise(999, "LNot", Err.Description)
End Sub

'#FlagExpansion
Private Sub IImgFlag_LOr(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'���e:
'   pSrcName1�̃r�b�g��pSrcName2�̃r�b�g��Or���Z�̌��ʂ�pDstName�̃r�b�g�ɓ����B
'   pDstName���o�^����Ă��Ȃ��ꍇ�A�V���ɓo�^����B
'   ���ɓo�^����Ă���ꍇ�́A���̃r�b�g�ɓ����B
'
'[pSrcName1]    IN String�^:        �f�[�^���̖��O1
'[pSrcName2]    IN String�^:        �f�[�^���̖��O2
'[pDstName]     IN String�^:        ���ʂ̖��O
'[pColor]       IN IdpColorType�^:  �F�w��
'
'���l:
'   pSrcName1��pSrcName2��pDstName������ł��\�B
'   ���O��SetPMD�Ŏw�肳�ꂽ�]�[�����Ώ�
'
    '�G���[����
    If m_FlgPlane Is Nothing Then
        Call TheError.Raise(999, "LOr", "There's no flag.")
    End If
    
On Error GoTo NO_FLAG
    Dim SrcFlagPlane1 As CImgPlane
    Set SrcFlagPlane1 = GetInputFlagPlane(pSrcName1)
    Dim SrcFlagPlane2 As CImgPlane
    Set SrcFlagPlane2 = GetInputFlagPlane(pSrcName2)
    With m_FlgPlane
        If IImgFlag_FlagBit(pDstName) = 0 Then Call IImgFlag_SetFlagBit(pDstName)
        Call .LOr(SrcFlagPlane1, SrcFlagPlane2, pColor, , , GetFlgBit(pDstName), GetFlgBit(pSrcName1), GetFlgBit(pSrcName2))
    End With

    Exit Sub
    
NO_FLAG:
    Call TheError.Raise(999, "LOr", Err.Description)
End Sub

'#FlagExpansion
Private Sub IImgFlag_LAnd(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'���e:
'   pSrcName1�̃r�b�g��pSrcName2�̃r�b�g��And���Z�̌��ʂ�pDstName�̃r�b�g�ɓ����B
'   pDstName���o�^����Ă��Ȃ��ꍇ�A�V���ɓo�^����B
'   ���ɓo�^����Ă���ꍇ�́A���̃r�b�g�ɓ����B
'
'[pSrcName1]    IN String�^:        �f�[�^���̖��O1
'[pSrcName2]    IN String�^:        �f�[�^���̖��O2
'[pDstName]     IN String�^:        ���ʂ̖��O
'[pColor]       IN IdpColorType�^:  �F�w��
'
'���l:
'   pSrcName1��pSrcName2��pDstName������ł��\�B
'   ���O��SetPMD�Ŏw�肳�ꂽ�]�[�����Ώ�
'
    '�G���[����
    If m_FlgPlane Is Nothing Then
        Call TheError.Raise(999, "LAnd", "There's no flag.")
    End If
    
On Error GoTo NO_FLAG
    Dim SrcFlagPlane1 As CImgPlane
    Set SrcFlagPlane1 = GetInputFlagPlane(pSrcName1)
    Dim SrcFlagPlane2 As CImgPlane
    Set SrcFlagPlane2 = GetInputFlagPlane(pSrcName2)
    With m_FlgPlane
        If IImgFlag_FlagBit(pDstName) = 0 Then Call IImgFlag_SetFlagBit(pDstName)
        Call .LAnd(SrcFlagPlane1, SrcFlagPlane2, pColor, , , GetFlgBit(pDstName), GetFlgBit(pSrcName1), GetFlgBit(pSrcName2))
    End With

    Exit Sub
    
NO_FLAG:
    Call TheError.Raise(999, "LAnd", Err.Description)
End Sub

'#FlagExpansion
Private Sub IImgFlag_LXor(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'���e:
'   pSrcName1�̃r�b�g��pSrcName2�̃r�b�g��Or���Z�̌��ʂ�pDstName�̃r�b�g�ɓ����B
'   pDstName���o�^����Ă��Ȃ��ꍇ�A�V���ɓo�^����B
'   ���ɓo�^����Ă���ꍇ�́A���̃r�b�g�ɓ����B
'
'[pSrcName1]    IN String�^:        �f�[�^���̖��O1
'[pSrcName2]    IN String�^:        �f�[�^���̖��O2
'[pDstName]     IN String�^:        ���ʂ̖��O
'[pColor]       IN IdpColorType�^:  �F�w��
'
'���l:
'   pSrcName1��pSrcName2��pDstName������ł��\�B
'   ���O��SetPMD�Ŏw�肳�ꂽ�]�[�����Ώ�
'
    '�G���[����
    If m_FlgPlane Is Nothing Then
        Call TheError.Raise(999, "LXor", "There's no flag.")
    End If
    
On Error GoTo NO_FLAG
    Dim SrcFlagPlane1 As CImgPlane
    Set SrcFlagPlane1 = GetInputFlagPlane(pSrcName1)
    Dim SrcFlagPlane2 As CImgPlane
    Set SrcFlagPlane2 = GetInputFlagPlane(pSrcName2)
    With m_FlgPlane
        If IImgFlag_FlagBit(pDstName) = 0 Then Call IImgFlag_SetFlagBit(pDstName)
        Call .LXor(SrcFlagPlane1, SrcFlagPlane2, pColor, , , GetFlgBit(pDstName), GetFlgBit(pSrcName1), GetFlgBit(pSrcName2))
    End With

    Exit Sub
    
NO_FLAG:
    Call TheError.Raise(999, "LXor", Err.Description)
End Sub

'#FlagExpansion
Private Function IImgFlag_SetPMD(ByVal pPmd As Variant) As IImgFlag
'���e:
'   �w�肵���]�[���ɐݒ�
'
'[pPmd]         IN Variant�^:       �]�[���w��
'
'���l:
'   pSrcName1��pSrcName2��pDstName������ł��\�B
'
    '�v���[���m�ۏ���
    If m_FlgPlane Is Nothing Then
        Set m_FlgPlane = TheIDP.PlaneManager(m_Group).GetFreePlane(idpDepthS16)
        m_FlgPlane.Comment = "Shared Flag Plane"
        Call m_FlgPlane.SetPMD(m_FlgPlane.BasePMD).WritePixel(0)
        Call InitializeFlagState
    End If
    
    Call m_FlgPlane.SetPMD(pPmd)
    Set IImgFlag_SetPMD = Me
End Function

'#FlagExpansion
Private Function IImgFlag_SetCustomPMD(ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long) As IImgFlag
'���e:
'   �w�肵���]�[���ɐݒ�
'
'[pPmd]         IN Variant�^:       �]�[���w��
'
'���l:
'   pSrcName1��pSrcName2��pDstName������ł��\�B
'
    '�v���[���m�ۏ���
    If m_FlgPlane Is Nothing Then
        Set m_FlgPlane = TheIDP.PlaneManager(m_Group).GetFreePlane(idpDepthS16)
        m_FlgPlane.Comment = "Shared Flag Plane"
        Call m_FlgPlane.SetPMD(m_FlgPlane.BasePMD).WritePixel(0)
        Call InitializeFlagState
    End If
    
    Call m_FlgPlane.SetCustomPMD(pX, pY, pWidth, pHeight)
    Set IImgFlag_SetCustomPMD = Me
    
End Function

Private Property Get IImgFlag_Name() As String
'���e:
'   �t���O�p�ɐݒ肳��Ă���v���[���̖��O�擾
    If m_FlgPlane Is Nothing Then
        IImgFlag_Name = ""
    Else
        IImgFlag_Name = m_FlgPlane.Name
    End If
End Property

Private Property Get IImgFlag_FlagBitList() As String
'���e:
'   ���ݐݒ肳��Ă���t���O�ԍ��ƁA�t����ꂽ���O�̃��X�g���擾

    Dim i As Long
    
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    For i = minBit To maxBit
        If m_FlgState.flagName(i) <> "" Then
            IImgFlag_FlagBitList = IImgFlag_FlagBitList & Format$(i, "##") & ":" & m_FlgState.flagName(i) & vbCrLf
        End If
    Next i
    
End Property


Private Sub IImgFlag_ResetFlagState()

    Dim i As Long
        
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    With m_FlgState
        For i = minBit To maxBit
            .FlagState(i) = Uninitialized
            .flagName(i) = ""
        Next i
    End With
    
    m_FlgState.Count = 0
    
End Sub

Private Sub InitializeFlagState()

    Dim i As Long
        
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    With m_FlgState
        For i = minBit To maxBit
            .FlagState(i) = INITIALIZED
        Next i
    End With
    
End Sub

Private Function GetInitializedFlag() As Long

    Dim i As Long
    
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
        
    With m_FlgState
        For i = minBit To maxBit
            If .FlagState(i) = INITIALIZED Then
                GetInitializedFlag = i
                Exit Function
            End If
        Next i
    End With
    
    GetInitializedFlag = 0

End Function

Private Sub RecoverFlagPlane(ByRef plane As CImgPlane)
    
    Dim i As Long
    Dim mask As Long
    
    mask = &HFFFF
    
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    With m_FlgState
        For i = minBit To maxBit
            If .FlagState(i) = NotUsed Then
                mask = mask - 2 ^ (i - 1)
                .FlagState(i) = INITIALIZED
            End If
        Next i
    End With
    
    Call m_FlgPlane.SetPMD(m_FlgPlane.BasePMD).WritePixel(0, , , , mask)
    
    
End Sub

Private Sub SetUsedState(ByVal pBit As Long, ByVal pName As String)
    
    On Error GoTo ErrorExit
    With m_FlgState
        .flagName(pBit) = pName
        .FlagState(pBit) = Used
        .Count = .Count + 1
    End With
    Exit Sub
    
ErrorExit:
    Call TheError.Raise(999, "SetUsedState", "Internal Error")
End Sub

Private Sub SetNotUsedState(ByVal pBit As Long)
    
    On Error GoTo ErrorExit
    With m_FlgState
        .flagName(pBit) = ""
        .FlagState(pBit) = NotUsed
        .Count = .Count - 1
    End With
    Exit Sub
    
ErrorExit:
    Call TheError.Raise(999, "SetNotUsedState", "Internal Error")
End Sub


Private Property Get FlgCount_() As Long
    FlgCount_ = m_FlgState.Count
End Property

Private Property Get MinIndex_() As Long
    MinIndex_ = LBound(m_FlgState.flagName)
End Property

Private Property Get MaxIndex_() As Long
    MaxIndex_ = UBound(m_FlgState.flagName)
End Property

Private Sub Class_Initialize()
'init

End Sub

Private Sub Class_Terminate()
'term
    Set m_FlgPlane = Nothing

End Sub

Public Function AsIImgFlag() As IImgFlag
    Set AsIImgFlag = Me
End Function

'#FlagExpansion
'�t���O�������Ƀt���O�v���[�������o���A���ʏo�̓v���[��(����)�Ɠ���PMD�ɐݒ肵�ĕԂ��֐�
Private Function GetInputFlagPlane(ByVal pFlagName As String) As CImgPlane
    '�t���O�v���[���擾
    Dim flag As IImgFlag
    Set flag = m_FlgPlane.Manager.GetSharedFlagPlane(pFlagName)
    'PMD�ݒ�
    If m_FlgPlane.CurrentPMD.Name = m_FlgPlane.Manager.VariablePMDName Then
        flag.SetPMD m_FlgPlane.CurrentPMD
    Else
        flag.SetPMD m_FlgPlane.CurrentPMD.Name
    End If
    Set GetInputFlagPlane = flag.FlgPlane
End Function

