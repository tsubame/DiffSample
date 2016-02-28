VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataSheetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�T�v:
'   �R���g���[���V�[�g�iEx�DJob List�j�̃}�l�[�W���N���X
'
'�ړI:
'   �R���g���[���V�[�g��ɒ�`���ꂽ�A�N�e�B�u�V�[�g���擾����
'
'�쐬��:
'   0145206097
'
Option Explicit

Const DATATOOL_NAME = "B1"

Dim mCtrlSheet As Worksheet
Dim mDataSheet As Collection

Private Sub Class_Initialize()
    Set mCtrlSheet = Nothing
    Set mDataSheet = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCtrlSheet = Nothing
    Set mDataSheet = Nothing
End Sub

Public Sub Initialize(ByVal ctrlSheet As String)
'���e:
'   �R���g���[���V�[�g�̎w��
'
'[ctrlSheet]   IN String�^:     �R���g���[���V�[�g��
'
'���l:
'   �R���g���[���V�[�g����ȏ゠��ꍇ�̓G���[�����������s��
'
    Dim wsheet As Collection
    Set wsheet = collectSheet(ctrlSheet)
    If wsheet Is Nothing Then
        Set mCtrlSheet = Nothing
    ElseIf wsheet.Count = 1 Then
        Set mCtrlSheet = wsheet.Item(1)
    Else
        Err.Raise 9999, "CDataSheetManager.Initialize", "Find More Than One Control Sheet [" & ctrlSheet & "] !"
    End If
End Sub

Public Function GetActiveDataSht(ByVal DataSheet As String) As Worksheet
'���e:
'   �擾������Data Tool�̎�ނ���A�N�e�B�u�ȃV�[�g����肵�I�u�W�F�N�g��Ԃ�
'
'[dataSheet]   IN String�^:     �擾������Data Tool��
'
'���l:
'
    Set GetActiveDataSht = Nothing
    '### Book���̑ΏۂƂȂ�DataTool�V�[�g�̃I�u�W�F�N�g���擾 #################
    Set mDataSheet = collectSheet(DataSheet)
    '### Book���ɑΏۂƂȂ�DataTool�V�[�g�����݂��Ȃ��Ƃ��̓G���[�������� #####
    If mDataSheet Is Nothing Then
        Exit Function
'        Err.Raise 9999, "CDataSheetManager.GetActiveDataSht", "Can Not Find The Data Tool Sheet [" & datasheet & "] !"
    End If

    Dim currSheet As Worksheet
    Select Case DataSheet
        '### �ΏۂƂȂ�DataTool�V�[�g��Channel Map�̏ꍇ#######################
        Case CHAN_MAP_TOOL:
            For Each currSheet In mDataSheet
                If currSheet.Name = GetCurrentChanMap Then
                    Set GetActiveDataSht = currSheet
                End If
            Next currSheet
        Case Else
        '### �R���g���[���V�[�g�����݂��Ȃ��ꍇ ###############################
        If mCtrlSheet Is Nothing Then
            If mDataSheet.Count = 1 Then
                Set GetActiveDataSht = mDataSheet.Item(1)
            Else
                '<<< �Ώۂ�DataTool�V�[�g����������ꍇ�̓G���[��������
'                Err.Raise 9999, "CDataSheetManager.GetActiveDataSht", "Find More Then One Data Tool Sheet [" & datasheet & "] !"
            End If

        '### �R���g���[���V�[�g�����݂���ꍇ #################################
        Else
            '<<< �e�R���g���[���V�[�g�p�̃��[�_�[�I�u�W�F�N�g�𐶐�
            Dim shReader As IParameterReader
            Select Case mCtrlSheet.Range(DATATOOL_NAME).Value
                Case JOB_LIST_TOOL:
                    Dim crtlSheet As New CJobListSheetReader
                    crtlSheet.Initialize mCtrlSheet.Name
                    crtlSheet.AsIFileStream.SetLocation GetCurrentJobName
                    Set shReader = crtlSheet
                Case SHEET_MANAGER_TOOL:
                    Dim manageSheet As New CManagerSheetReader
                    manageSheet.Initialize mCtrlSheet.Name
                    manageSheet.AsIFileStream.SetLocation GetCurrentJobName
                    Set shReader = manageSheet
                Case Else
                    Err.Raise 9999, "CDataSheetManager.GetActiveDataSht", "Does Not Support The Contorl Sheet [" & mCtrlSheet.Range(DATATOOL_NAME).Value & "] !"
            End Select
            '<<< �R���g���[���V�[�g����Ώۂ�DataTool�V�[�g�����擾
            Dim paramName As CParamString
            Set paramName = CreateCParamString
            Dim targetSheet As String
            With paramName.AsIParameter
                .Name = DataSheet
                .Read shReader
                targetSheet = .AsString
            End With
            If targetSheet <> "" Then
                '<<< Book����DataTool�V�[�g���ƃR���g���[���V�[�g���DataTool�V�[�g�����r
                For Each currSheet In mDataSheet
                    If currSheet.Name = targetSheet Then
                        Set GetActiveDataSht = currSheet
                    End If
                Next currSheet
                '<<< �s��v�̏ꍇ�̓G���[��������
'                If GetActiveDataSht Is Nothing Then
''                    Err.Raise 9999, "CDataSheetManager.GetActiveDataSht", "Can Not Find The Sheet Name [" & targetSheet & "] In Data Tool Sheet [" & datasheet & "] !"
'                End If
            End If
        End If
    End Select
End Function

Public Property Get ManagerSheet() As Worksheet
    Set ManagerSheet = mCtrlSheet
End Property

Private Function collectSheet(ByVal DataSheet As String) As Collection
    Dim currSheet As Worksheet
    Set collectSheet = New Collection
    For Each currSheet In Worksheets
        If currSheet.Range(DATATOOL_NAME).Value = DataSheet Then
            collectSheet.Add currSheet
        End If
    Next currSheet
    If collectSheet.Count = 0 Then
        Set collectSheet = Nothing
    End If
End Function