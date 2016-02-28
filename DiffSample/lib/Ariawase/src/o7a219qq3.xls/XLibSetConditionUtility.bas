Attribute VB_Name = "XLibSetConditionUtility"
'�T�v:
'   ��������ݒ�@�\��
'
'�ړI:
'   ���[�N�V�[�g�̃f�[�^���g�p������������ݒ�̎���
'   ��������ݒ�@�\�̃C�j�V�����C�Y�֌W�̋L�q
'
'�쐬��:
'   SLSI ����
'   tomoyoshi.takase
'
'���ӓ_:
'   ����̗��p���ɁAInitialize�̎��s�����p�ł��B
'   TheError�Ƃ��Č��J����Ă���G���[�}�l�[�W��Object���K�v�ł�
'   �ύX����
'�@ 2010/03/08�A�A�h�C���łō쐬����Ă������̂��A���W���[���œ��삷��悤�ɕύX�B
'�@ AddWorkSheet���\�b�h�̓A�h�C���łȂ��Ȃ����̂Ŕp�~
'

Option Explicit

'�G���[���̏���`
Private Const ERR_NUMBER = 9999                              '�G���[���ɓn���G���[�ԍ�
Private Const CLASS_NAME = "XLibSetConditionUtilities" '���̃N���X�̖��O

'���J�@�\ Object
Public TheCondition As CTestConditionManager

Private mSaveFileName As String

'���JEee-JOB���[�N�V�[�g�̒�`�p
Enum EEE_CONDITION_WORKSHEET
    TestCondition_EeeJobSheet = 0
End Enum

Public Sub CreateTheConditionIfNothing()
'���e:
'   As New�̑�ւƂ��ď���ɌĂ΂���C���X�^���X��������
'
'�p�����[�^:
'   �Ȃ�
'
'���ӎ���:
'
    On Error GoTo ErrHandler
    If TheCondition Is Nothing Then
        '######## TestCondition Block ########
        Call CreateTestCondition(ThisWorkbook)
        '��������\�̏����ݒ�
        TheCondition.TestConditionSheet = GetWkShtReaderManagerInstance.GetActiveSheetName(eSheetType.shtTypeTestCond)        '�����\���[�N�V�[�g����ݒ肷��
        Call TheCondition.LoadCondition
    End If
    Exit Sub
ErrHandler:
    Set TheCondition = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Public Sub SetLogModeTheCondition(ByVal pEnableLoggingTheCondition As Boolean, Optional saveFileName As String = "EeeJOBLogSetCondition.csv")
    Call TheCondition.LoadCondition
    TheCondition.CanHistoryRecord = pEnableLoggingTheCondition
    mSaveFileName = saveFileName
End Sub


'���C�u�����̏�����
Public Sub CreateTestCondition(ByVal pJobWorkBook As Workbook)
'���e:
'   EeeTestCondition�S�̂̏�����
'
'�p�����[�^:
'   [pTheErrorObj]  In  Object�^:     �G���[�Ǘ��@�\��Object
'   [pJobWorkBook]  In  Workbook�^:   JOB��Workbook
'
'�߂�l:
'
'���ӎ���:
'
    '��������}�l�[�W���̏�����
    Call InitTestConditionManager(pJobWorkBook.Name)
    
End Sub

Public Sub DestroyTestCondition()
'���e:
'   EeeTestConditionAddIn�S�̂̏I������
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Set TheCondition = Nothing

End Sub

Public Function GetTestConditionInstance() As CTestConditionManager
'���e:
'   ��������ݒ�@�\Object�̒�
'
'�p�����[�^:
'
'�߂�l:
'   ��������ݒ�@�\Object
'
'���ӎ���:
'   Initialize�������s�̏�ԂŁA�{���߂����s����ƃG���[�ƂȂ�܂��B
'
'
    If TheCondition Is Nothing Then
        Call TheError.Raise(9999, CLASS_NAME, "Initialization is unexecution." & " @EeeTestConditionAddIn")
        'Call CreateTestCondition(ThisWorkbook)
    Else
        Set GetTestConditionInstance = TheCondition
    End If
End Function

Public Function RunAtJobEnd() As Long
    If Not (TheCondition Is Nothing) Then
        If TheCondition.CanHistoryRecord Then
            Call TheCondition.SaveHistoryLog(mSaveFileName)
            Call TheCondition.ClearExecHistory
        End If
        TheCondition.CanHistoryRecord = False
    End If
End Function


'------------------------------------------------------------------------------------------------
'�ȉ� Private

'�C���X�^���X�̐����Ə������̏���
Private Sub InitTestConditionManager(ByVal pJobWorkbookName As String)
'���e:
'
'�쐬��:
'  tomoyoshi.takase
'�p�����[�^:
'   [pJobWorkbookName]  In  1):�Ώۃ��[�N�u�b�N��
'�߂�l:
'
'���ӎ���:
'

    Set TheCondition = New CTestConditionManager
    Call TheCondition.Initialize
    TheCondition.JobWorkbookName = pJobWorkbookName
End Sub

Public Sub ChangeDefaultSettingTheCondition()
'���e:
'   �I�����Ă���O���[�v��Default�ɕύX����悤�ɗv������
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Call TheCondition.ChangeDefaultSetting

End Sub

