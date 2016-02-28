VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgFlagPlanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �����t���O�v���[���̊Ǘ��N���X
'
'�ړI:
'   �t���O�v���[���𕡐����p�\�ɂ���

Option Explicit

Private m_FlagPlanes As Collection
Private Const MAX_FLAG_BIT As Long = 16

'#FlagExpansion
'������ #Pass
Public Sub Initialize(ByVal pGroupName As String)
    Set m_FlagPlanes = New Collection
    Dim FlagPlane As IImgFlag
    Set FlagPlane = New CImgFlag
    FlagPlane.group = pGroupName
    m_FlagPlanes.Add FlagPlane
End Sub

'#FlagExpansion
'�C���^�[�Z�v�^�[�L���� #Pass
Public Sub EnableInterceptor()
    
    Dim FlagPlane As IImgFlag
    Dim Interceptor As CImgFlagInterceptor
    Dim TempPlanes As Collection
    Set TempPlanes = New Collection
    
    '�^����ɂ�錻��m�F�Ɗ��荞�ݐݒ�
    For Each FlagPlane In m_FlagPlanes
        If TypeName(FlagPlane) = "CImgFlagInterceptor" Then
            '���荞�ݍς݁A���̂܂ܓo�^
            TempPlanes.Add FlagPlane
        Else
            '�C���^�[�Z�v�^�[�����Ɗ��荞�ݓo�^
            Set Interceptor = New CImgFlagInterceptor
            Set Interceptor.TrueInstance = FlagPlane
            TempPlanes.Add Interceptor.AsIImgFlag
        End If
    Next
    
    '�u������
    Set m_FlagPlanes = TempPlanes
    
End Sub

'#FlagExpansion
'�C���^�[�Z�v�^�[������ #Pass
Public Sub DisableInterceptor()
    
    Dim FlagPlane As IImgFlag
    Dim Interceptor As CImgFlagInterceptor
    Dim TempPlanes As Collection
    Set TempPlanes = New Collection
    
    '�^����ɂ�錻��m�F�Ɗ��荞�ݐݒ�
    For Each FlagPlane In m_FlagPlanes
        If TypeName(FlagPlane) = "CImgFlag" Then
            '���荞�ݖ����A���̂܂ܓo�^
            TempPlanes.Add FlagPlane
        Else
            '�C���^�[�Z�v�^�[���O���Ɠo�^
            Set Interceptor = FlagPlane '�^�ϊ�
            TempPlanes.Add Interceptor.TrueInstance
        End If
    Next
    
    '�u������
    Set m_FlagPlanes = TempPlanes
    
End Sub

'#FlagExpansion
'���ׂẴt���O�v���[���̋����J�� #Pass
Public Sub ClearAllFlag()
    '�v���[���J��
    Dim FlagPlane As IImgFlag
    For Each FlagPlane In m_FlagPlanes
        FlagPlane.Clear
    Next
    '�t���O�C���X�^���X��1�����̏�Ԃɂ���
    If m_FlagPlanes.Count >= 2 Then
        Dim MaxCount As Long
        Dim ItemIndex As Long
        MaxCount = m_FlagPlanes.Count
        For ItemIndex = 1 To MaxCount - 1
            m_FlagPlanes.Remove 2
        Next
    End If
End Sub

'#FlagExpansion
'���݂���t���O�v���[�����ׂĂ�Ԃ� #Pass
Public Function FlagPlanes() As Collection
    Set FlagPlanes = m_FlagPlanes
End Function

'#FlagExpansion
'�w��t���O�����݂���t���O�v���[����Ԃ� #Pass
Public Function FlagPlane(ByVal pFlagName As String) As IImgFlag
    
    Dim TargetPlane As IImgFlag
    
    '1:�w��̃t���O�����o�^����Ă���v���[������������
    For Each TargetPlane In m_FlagPlanes
        If TargetPlane.FlagBit(pFlagName) <> 0 Then
            Set FlagPlane = TargetPlane
            Exit Function '���łɓo�^����Ă�����̂𗘗p
        End If
    Next

    '2:�t���O�̋󂫂�����v���[������������
    For Each TargetPlane In m_FlagPlanes
        If TargetPlane.Count < MAX_FLAG_BIT Then
            Set FlagPlane = TargetPlane '�o�^����Ă���v���[����Ԃ�
            Exit Function '�󂫂̑��݂�����̂�Ԃ�
        End If
    Next
    
    '3:�󂫂����݂��Ȃ��̂ŐV�K�쐬�������Ԃ�
    Dim ChkFlagPlane As IImgFlag
    Set ChkFlagPlane = m_FlagPlanes.Item(1)
    Dim NewFlagPlane As IImgFlag
    
    '���݂̃C���^�[�Z�v�^�[��Ԃ��m�F����
    If TypeName(ChkFlagPlane) = "CImgFlag" Then
        '�C���^�[�Z�v�^�[����
        Set NewFlagPlane = New CImgFlag
        NewFlagPlane.group = ChkFlagPlane.group
        m_FlagPlanes.Add NewFlagPlane
        Set FlagPlane = NewFlagPlane
    Else
        '�C���^�[�Z�v�^�[�L��
        Dim Interceptor As CImgFlagInterceptor
        Set Interceptor = New CImgFlagInterceptor
        Set NewFlagPlane = New CImgFlag
        NewFlagPlane.group = ChkFlagPlane.group
        Set Interceptor.TrueInstance = NewFlagPlane
        m_FlagPlanes.Add Interceptor.AsIImgFlag
        Set FlagPlane = Interceptor.AsIImgFlag
    End If

End Function

'#FlagExpansion
'�t���O���p��0�̃t���O�v���[�����J������֐�
Public Function ReleaseUnusedFlagPlane() As Long
    Dim TargetPlane As IImgFlag
    Dim Count As Long
    For Each TargetPlane In m_FlagPlanes
        If TargetPlane.Count = 0 And (Not TargetPlane.FlgPlane Is Nothing) Then
            TargetPlane.Clear
            Count = Count + 1
        End If
    Next
    ReleaseUnusedFlagPlane = Count
End Function