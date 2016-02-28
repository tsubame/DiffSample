VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CScenarioFormDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �t�H�[���ɑ΂���p�����[�^�_���v�N���X
'
'�ړI:
'   �t�H�[���ɏ����o�͂��郉�C�^�[
'   �摜�̏ꍇ�̓C���[�W�r���[���[�ŏo�͂���
'
'�쐬��:
'   0145184306
'
Option Explicit

Implements IProcParamWriter
Implements IParameterViewer

Private Const ERR_NUMBER = 9999         ' �G���[�ԍ���ێ�����
Private Const CLASS_NAME = "CScenarioFormDebugger" ' Class���̂�ێ�����
Private Const INDENT_SPACE = vbTab

Private m_ParamForm As ScenarioParameterViewer
Private m_ViewPlane As Collection
Private m_IGXLDebugDisplay As Object
'For Debug
Private Const DUMPKEYWORD_FOR_PARAM_NAME As String = "Parameter Name"
Private Const DUMPKEYWORD_FOR_CLASS_NAME As String = "Parameter Type"
Private Const DUMPKEYWORD_FOR_COMMENT As String = "Comment"
Private Const DUMPKEYWORD_FOR_NULL As String = ""
Private Const TAB_WIDTH = 4
Private m_IndentLevel As Long

Public Function AsIProcParamWriter() As IProcParamWriter
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IProcParamWriter�^�I�u�W�F�N�g
'
'���l:
'
    Set AsIProcParamWriter = Me
End Function

Public Function AsIParameterViewer() As IParameterViewer
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IParameterViewer�^�I�u�W�F�N�g
'
'���l:
'
    Set AsIParameterViewer = Me
End Function

Private Sub IParameterViewer_View()
'���e:
'   �_���v���������t�H�[���ɕ\������
'   �摜�f�[�^������ꍇ�̓C���[�W�f�[�^�r���[���[�ŕ\������
'
'���l:
'
'
    If m_ParamForm.EndStatus = True Then Exit Sub
    
    Dim vPlane As CImgPlane
    'Plane�̕\��
    If m_ViewPlane.Count > 0 Then
        For Each vPlane In m_ViewPlane
'            vPlane.View
            With theidv
                Call .OpenForm
                .PlaneNameGreen = vPlane.Name
                .PMDName = vPlane.CurrentPmdName
                Call .Refresh
            End With
        Next vPlane
    End If
    m_IGXLDebugDisplay.Show
    'Form�̕\��
    Call m_ParamForm.Display
    
    Call m_ParamForm.Hide
    Set m_ViewPlane = New Collection
    
End Sub

Private Sub IProcParamWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
'���e:
'   Boolean�^�����_���v����B
'
'����:
'[dataName]  �@IN String�^:     �f�[�^��
'[blnData]�@�@ IN Boolean�^:     �f�[�^
'
'���l:
'
'
    With Me.AsIProcParamWriter
        .WriteAsString dataName, blnData
    End With
End Sub

Private Sub IProcParamWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
'���e:
'   Double�^�����_���v����B
'
'����:
'[dataName]  �@IN String�^:     �f�[�^��
'[dblData]�@�@ IN Double�^:     �f�[�^
'
'���l:
'
'
    With Me.AsIProcParamWriter
        .WriteAsString dataName, dblData
    End With
End Sub

Private Sub IProcParamWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
'���e:
'   Integer�^�����_���v����B
'
'����:
'[dataName]  �@IN String�^:     �f�[�^��
'[intData]�@�@ IN Integer�^:     �f�[�^
'
'���l:
'
'
    With Me.AsIProcParamWriter
        .WriteAsString dataName, intData
    End With
End Sub

Private Sub IProcParamWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
'���e:
'   Long�^�����_���v����B
'
'����:
'[dataName]  �@IN String�^:     �f�[�^��
'[lngData]�@�@ IN Long�^:     �f�[�^
'
'���l:
'
'
    With Me.AsIProcParamWriter
        .WriteAsString dataName, lngData
    End With
End Sub

Private Sub IProcParamWriter_WriteAsPlane(ByVal pParamName As String, ByRef pImgPlane As CImgPlane)
'���e:
'   �C���[�W�v���[���̉摜�Ə����_���v����B
'
'[pParamName]  IN String�^:     �w�b�_�[���x��
'[pParamPlane] IN CImgPlane�^:  �C���[�W�v���[��
'
'���l:
'
'
    If Not pImgPlane Is Nothing Then
        With Me.AsIProcParamWriter
            .WriteAsString "Plane Name", pImgPlane.Name
            .WriteAsString "Plane Group", pImgPlane.planeGroup
            .WriteAsString "Bit Depth", idpDepthToStr(pImgPlane.BitDepth)
            .WriteAsString "Color Map Name", pImgPlane.planeMap.Name
            .WriteAsString "Color Map Definition", pImgPlane.planeMap.ColorNamesPmd
            .WriteAsString "Current PMD", pImgPlane.CurrentPmdName
            .WriteAsString "Read Only", pImgPlane.ReadOnly
            .WriteAsString "Plane Comment", pImgPlane.Comment
        End With
    
        m_ViewPlane.Add pImgPlane
    Else
        Me.AsIProcParamWriter.WriteAsString INDENT_SPACE & "Plane", "Nothing"
    End If

End Sub

Private Sub IProcParamWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
'���e:
'   �e�L�X�g�����t�H�[���Ƀ_���v����B
'
'[dataName]    IN String�^:     �f�[�^��
'[strData]     IN String�^:     �f�[�^
'
'���l:
'
'
    With m_ParamForm.ScenarioParamView
        Select Case dataName
            Case DUMPKEYWORD_FOR_PARAM_NAME
                    If .Text = "" Then
                        .Text = "[" & Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & "]" & " = " & strData
                    Else
                        .Text = .Text & vbCrLf & vbCrLf & "[" & Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & "]" & " = " & strData
                    End If
            Case DUMPKEYWORD_FOR_CLASS_NAME, DUMPKEYWORD_FOR_COMMENT
                    If .Text = "" Then
                        .Text = "[" & Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & "]" & " = " & strData
                    Else
                        .Text = .Text & vbCrLf & "[" & Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & "]" & " = " & strData
                    End If
            Case DUMPKEYWORD_FOR_NULL
                    If .Text = "" Then
                        .Text = Strings.Space$(m_IndentLevel * TAB_WIDTH) & strData
                    Else
                        .Text = .Text & vbCrLf & Strings.Space$(m_IndentLevel * TAB_WIDTH) & strData
                    End If
            Case Else
                    If .Text = "" Then
                        .Text = Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & " = " & strData
                    Else
                        .Text = .Text & vbCrLf & Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & " = " & strData
                    End If
        End Select
    End With
End Sub

Private Property Get IParameterViewer_EndStatus() As Boolean
'���e:
'   �t�H�[���̏I���X�e�[�^�X���擾����
'
'�߂�l:
'   �����I���t���O(Boolean�^)
'
'���l:
'
'
    IParameterViewer_EndStatus = m_ParamForm.EndStatus
End Property
Private Property Let IParameterViewer_EndStatus(ByVal pStatus As Boolean)
'���e:
'   �t�H�[���̏I���X�e�[�^�X���擾����
'
'�߂�l:
'   �����I���t���O(Boolean�^)
'
'���l:
'
'
    m_ParamForm.EndStatus = pStatus
End Property
Private Property Get IProcParamWriter_MyName() As String
'���e:
'   �����̖��O���擾����
'
'�߂�l:
'   �N���X��(String�^)
'
'���l:
'
'
    IProcParamWriter_MyName = CLASS_NAME
End Property

Private Property Get IParameterViewer_MyName() As String
'���e:
'   �����̖��O���擾����
'
'�߂�l:
'   �N���X��(String�^)
'
'���l:
'
'
    IParameterViewer_MyName = CLASS_NAME
End Property

Private Sub Class_Initialize()
    Set m_ParamForm = ScenarioParameterViewer
    Set m_ViewPlane = New Collection
    Set m_IGXLDebugDisplay = CreateObject("CMU.IGXLDebugDisplay")
    m_IndentLevel = 0
End Sub

Private Sub Class_Terminate()
'    Unload m_ParamForm
    Set m_ParamForm = Nothing
    Set m_ViewPlane = Nothing
    Set m_IGXLDebugDisplay = Nothing
End Sub

Private Sub setHeader(ByVal pParamName As String)
    With m_ParamForm.ScenarioParamView
        If .Text = "" Then
            .Text = pParamName
        Else
            .Text = .Text & vbCrLf & pParamName
        End If
    End With
End Sub

Private Function composeString(ByRef strList() As String) As String

    Dim dataIndex As Long
    For dataIndex = 0 To UBound(strList)
        If dataIndex = 0 Then
            composeString = strList(dataIndex)
        Else
            composeString = composeString & "," & strList(dataIndex)
        End If
    Next dataIndex
End Function

Private Function idpDepthToStr(ByVal idpDepth As IdpBitDepth) As String
    Select Case idpDepth
        Case idpDepthF32
            idpDepthToStr = "idpDepthF32"
        Case idpDepthS16
            idpDepthToStr = "idpDepthS16"
        Case idpDepthS32
            idpDepthToStr = "idpDepthS32"
        Case Else
            idpDepthToStr = "-"
    End Select
End Function

Private Sub IProcParamWriter_WriteAsBlockEnd(ByVal pParamName As String)
    m_IndentLevel = m_IndentLevel - 1
End Sub

Private Sub IProcParamWriter_WriteAsBlockStart(ByVal pParamName As String)
    m_IndentLevel = m_IndentLevel + 1
End Sub

Public Sub IProcParamWriter_ReturnToRootBlock()
    m_IndentLevel = 0
End Sub