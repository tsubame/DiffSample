Attribute VB_Name = "XLibJobUtility"
'�T�v:
'   JOB�⃉�C�u�����ŋ��ʂɎg�p���郆�[�e�B���e�B�S
'
'�ړI:
'
'
'�쐬��:
'   SLSI����
'
'

Option Explicit

'#Pass
Public Sub OutputErrMsg(ByVal Msg As String)
'���e:
'   �G���[�����b�Z�[�W�Ƌ��ɏo��
'
'�p�����[�^:
'    [Msg]      In   �G���[���e���b�Z�[�W
'
'�߂�l:
'
'���ӎ���:
'
    Const NEW_JOB_ERR_NUMBER = 999

    Msg = "Error Message: " & vbCrLf & "    " & Msg & vbCrLf
    Msg = Msg & "Test Instance Name: " & vbCrLf & "    " & TheExec.DataManager.InstanceName

'    Call MsgBox(Msg, vbExclamation Or vbOKOnly, "Error")
    Call Err.Raise(NEW_JOB_ERR_NUMBER, "OutputErrMsg", Msg)

End Sub

Public Function CompareDblData(ByVal dblValue1 As Double, ByVal dblValue2 As Double, ByVal DIGIT As Long) As Boolean
'���e:
'   2��Double�^�f�[�^���w��L�������ȍ~��؂�̂ĂĔ�r���肷��
'
'�p�����[�^:
'    [dblValue1]    In   ��r�Ώۃf�[�^1
'    [dblValue2]    In   ��r�Ώۃf�[�^2
'    [digit]        In   �L������
'
'�߂�l:
'   ��r���茋��
'
'���ӎ���:
'
    With WorksheetFunction
        CompareDblData = (.RoundDown(dblValue1, DIGIT) = .RoundDown(dblValue2, DIGIT))
    End With
End Function

Public Function RoundDownDblData(ByVal dblValue As Double, ByVal DIGIT As Long) As Double
'���e:
'   Double�^�f�[�^�̎w��L�������ȍ~��؂�̂Ă郏�[�N�V�[�g�֐��̃��b�p�[
'
'�p�����[�^:
'    [dblValue]     In   �Ώۃf�[�^
'    [digit]        In   �L������
'
'�߂�l:
'   �w�肳�ꂽ������Double�^�f�[�^
'
'���ӎ���:
'
    RoundDownDblData = WorksheetFunction.RoundDown(dblValue, DIGIT)
End Function

Public Function GetJobRootPath() As String
'���e:
'   TheExec.Rootpath�̃��b�p�[�֐�
'
'�߂�l:
'   IG-XL�̌��݃��[�h����Ă���o�[�W�����̃C���X�g�[���t�H���_�̐�΃p�X
'
'���ӎ���:
'
    GetJobRootPath = TheExec.Rootpath
End Function

Public Function GetCurrentJobName() As String
'���e:
'   TheExec.CurrentJob�̃��b�p�[�֐�
'
'�߂�l:
'   �A�N�e�B�u��JOB���̎擾
'
'���ӎ���:
'
    GetCurrentJobName = TheExec.CurrentJob
End Function

Public Function GetCurrentChanMap() As String
'���e:
'   TheExec.CurrentChanMap�̃��b�p�[�֐�
'
'�߂�l:
'   �A�N�e�B�u�ȃ`�����l���}�b�v���̎擾
'
'���ӎ���:
'
    GetCurrentChanMap = TheExec.CurrentChanMap
End Function

Public Function IsJobValid() As Boolean
'���e:
'   TheExec.JobIsValid�̃��b�p�[�֐�
'
'�߂�l:
'   �o���f�[�V���������������s���ꂽ���ǂ���
'
'���ӎ���:
'
    IsJobValid = TheExec.JobIsValid
End Function

Public Sub CreateListBox(ByVal selCell As Range, ByRef dataList As Collection)
'���e:
'   �G�N�Z�����[�N�V�[�g�̔C�ӂ̃Z���Ƀ��X�g�{�b�N�X�t�H�[����ݒ肷��
'
'�p�����[�^:
'   [wsSheet]      In   �Ώۃ��[�N�V�[�g�I�u�W�F�N�g
'   [selCell]      In   �ΏۃZ���I�u�W�F�N�g
'   [listBoxData]  In   ���X�g�ɐݒ肷��f�[�^�R���N�V����
'
'���ӎ���:
'   ���X�g�{�b�N�X�̏����I���p�����[�^�͈ȉ��̒ʂ�
'   �@�ΏۃZ���Ɋ��Ƀp�����[�^�����͂���Ă���ꍇ
'    �E���X�g�ɑ��݂���p�����[�^�̏ꍇ�͂��̃p�����[�^�������I���p�����[�^�Ƃ��ĕ\������
'    �E���X�g�ɑ��݂��Ȃ��ꍇ�͋󔒁iListIndex=-1�j�������I������
'   �A�ΏۃZ���Ƀp�����[�^�����͂���Ă��Ȃ��ꍇ�͋󔒁iListIndex=-1�j�������I������
'
    '### �Â����X�g�{�b�N�X�̍폜 #########################
    Const myDropName = "DropDownList"
    On Error Resume Next
    selCell.parent.DropDowns(myDropName).Delete
    On Error GoTo 0
    '### �f�[�^���X�g�����݂��Ȃ��ꍇ��EXIT ###############
    Dim listBoxData As Collection
    Set listBoxData = dataList
    If listBoxData Is Nothing Then Exit Sub
    '### ���X�g�̏����\���f�[�^�̏��� #####################
    Dim dataIndex As Long
    Dim currData As Variant
    Dim IsContain As Boolean
    If IsEmpty(selCell) Then
        dataIndex = 0
    Else
        For Each currData In listBoxData
            If selCell.Value = currData Then
                dataIndex = dataIndex + 1
                IsContain = True
                Exit For
            End If
            dataIndex = dataIndex + 1
        Next currData
        If Not IsContain Then
            dataIndex = 0
        End If
    End If
    '### �E�B���h�E�T�C�Y�̒��� ###########################
    Dim currZoom As Double
    currZoom = ActiveWindow.Zoom
    Application.ScreenUpdating = False
    ActiveWindow.Zoom = 100
    '### ���X�g�{�b�N�X�̐ݒ� #############################
    With selCell.parent.DropDowns
        '���X�g�{�b�N�X������12�̃I�t�Z�b�g�̓h���b�v�{�^���̉�����
        .Add(selCell.Left, selCell.Top, selCell.width + 12, selCell.height).Name = myDropName
        For Each currData In listBoxData
            .AddItem (currData)
        Next currData
        .OnAction = "'selectData " & Chr(34) & myDropName & Chr(34) & "'"
        .ListIndex = dataIndex
    End With
    '### �E�B���h�E�T�C�Y�̍Ē��� #########################
    ActiveWindow.Zoom = currZoom
    Application.ScreenUpdating = True
End Sub

Public Sub SelectData(ByVal dropName As String)
    ActiveCell.Value = ActiveSheet.DropDowns(dropName).List(ActiveSheet.DropDowns(dropName).ListIndex)
End Sub

Public Sub RunAtValidationStart()
'���e:
'   Validation�X�^�[�g���Ɏ��s����֐�
'
'���ӎ���:
'   IG-XL�o�[�W������3.40.10JDXX�̏ꍇ�A
'   IG-XL��OnVaridationStart���̂��̂��Ă΂�Ȃ�
'   Validation�J�n���ɕK�����s���������̂�
'   ���̎�i���l����K�v������B
'

    '### Job���s����Validation�����s�����ꍇ ##############
    If TheExec.Flow.IsRunning Then
        MsgBox "CAUTION:" & vbCrLf & _
                "Validation starts when job is running." & vbCrLf & _
                "Please stop job running, and validate the job again.", vbExclamation, _
                "Eee-Job : IG-XL Event Handler"
    End If
    '######################################################

End Sub
