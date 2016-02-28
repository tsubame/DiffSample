Attribute VB_Name = "XEeeAuto_TestConditionModifyer"
'�T�v:
'   TestCondition����Ă΂��}�N���W
'
'�ړI:
'   TestCondition�V�[�g�������ŏȗ����邽�߂̋@�\
'
'�쐬��:
'   2012/03/24 Ver0.1 D.Maruyama    Draft
'   2012/04/19 Ver0.2 D.Maruyama    �R�[�h�����ꂢ�ɂ���
'   2013/02/25 Ver0.3 H.Arikawa     �������p�̒�`�E����(�����̓N���X)��ǉ�

Option Explicit

'DCHAN�̏����`�F�b�N�͎��Ԃ�������̂ŁADCHAN�ɐݒ肵�Ȃ��ꍇ�͌��ɍs���Ȃ����߂̒�`
Public Const EEE_AUTO_TESTER_CHECKER_WITHOUT_DCHAN As String = "Without Dchan"
Public Const EEE_AUTO_TESTER_CHECKER_WITH_DCHAN As String = "Dchan"

'�����ݒ�̎��
Public Enum eTestCnditionCheck
    TCC_TESTER_CHANNELS
    TCC_SETVOLTAGE
    TCC_ILLUMINATOR
    TCC_ILLUMINATOR_ESCAPSE
    TCC_ILLUMINATOR_MODZ1
    TCC_ILLUMINATOR_MODZ2
    TCC_APMU_UB
End Enum

'TestCondition�V�[�g����̒�`
Private Const CURRENT_SETTING As String = "C2"
Private Const ARG_MAX As Long = 10
Private Const TESTCONDITION_ITEM_START As String = "B5"
Private Const DEFAULT_ENABLES_START As String = "N5"
Private Const DEFAULT_LABEL As String = "Default"

'CEeeAuto_TestConditionItem���i�[����R���N�V����
Private m_colTestItems As Collection

'�e�X�g�ɉ����ă`�F�b�N���s���N���X�Q
Private m_IllminatorChecker As IEeeAuto_TestConditionChecker
Private m_TesterChannelChecker As IEeeAuto_TestConditionChecker
Private m_ApmuUbChecker As IEeeAuto_TestConditionChecker
Private m_IllmEscapeChecker As IEeeAuto_TestConditionChecker
Private m_IllmOptModZSet1Checker As IEeeAuto_TestConditionChecker
Private m_IllmOptModZSet2Checker As IEeeAuto_TestConditionChecker


'�ŏ��̐ݒ�͕K���s�킹�邽�߂̃t���O
Private m_IsFirstSetIlluminator As Boolean
Private m_IsFirstSetUB As Boolean
Private m_IsFirstSetVoltatge As Boolean
Private m_IsFirstSetIllmEscape As Boolean
Private m_IsFirstSetIllmOptModZSet1 As Boolean
Private m_IsFirstSetIllmOptModZSet2 As Boolean


'SetForceEnableTestCondition�Őݒ肳�ꂽ�R���f�B�V�������͊o���Ă���
'�`�F�b�N�͍s��Ȃ�
Private m_colForceEnableCondition As Collection

'���̃��W���[���̏�Ԓ�`
Private Enum TCCM_STATUS
    UNKNWON = 0
    INITIALIZED = 1
    CHECKED_BEFORE = 2
End Enum

'��ԕێ��ϐ�
Private m_State As TCCM_STATUS

'�����ȗ��@�\�̏�����
Public Sub InitializeAutoConditionModify()
    
    Dim lColumn As Long
    Dim curColumn As Long
    Dim mySht As Worksheet
    
    m_State = UNKNWON
    
    On Error GoTo ErrorHandler:
    
    'AutoModifyMode�ł��邱�Ƃ̒ʒm
    TheExec.Datalog.WriteComment "-----TestCondition Sheet Auto Modify Mode!! ---"
    
    '�����o�̏�����
    Set m_colTestItems = Nothing
    Set m_colTestItems = New Collection
    Set m_colForceEnableCondition = Nothing
    Set m_colForceEnableCondition = New Collection
    
    '�V�[�g�̎擾
    Set mySht = ThisWorkbook.Worksheets(TheCondition.TestConditionSheet)
    
    'Default�łȂ��ꍇ�͓����Ȃ�
    If (Not IsExecuteSetDefault(mySht)) Then
        Err.Raise 9999, "InitializeAutoConditionModify", "Execute list is not Default!"
        Exit Sub
    End If
        
    'TestCondition�V�[�g����̓ǂݍ���
    Call ReadTestCondition(mySht)
    
    '�I�u�W�F�N�g�̍\�z
    Set m_IllminatorChecker = New CEeeAuto_IlluminatorChecker
    Set m_ApmuUbChecker = New CEeeAuto_ApmuUBChecker
    Set m_TesterChannelChecker = New CEeeAuto_TesterChannelChecker
    Set m_IllmEscapeChecker = New CEeeAuto_IllumEscapeChecker
    Set m_IllmOptModZSet1Checker = New CEeeAuto_IllumModeZSet1Checker
    Set m_IllmOptModZSet2Checker = New CEeeAuto_IllumModeZSet2Checker

    '�ŏ��̃R�}���h�ł��邩�����ϐ���������
    m_IsFirstSetIlluminator = True
    m_IsFirstSetUB = True
    m_IsFirstSetVoltatge = True
    m_IsFirstSetIllmEscape = True
    m_IsFirstSetIllmOptModZSet1 = True
    m_IsFirstSetIllmOptModZSet2 = True
    
    '�J��
    Set mySht = Nothing
    
    '��ԑJ��
    m_State = INITIALIZED

    Exit Sub
    
ErrorHandler:
    Set mySht = Nothing
    m_State = UNKNWON
    
End Sub

'�����ȗ��@�\�̏I������
Public Sub UninitializeAutoConditionModify()
    
    If m_State = INITIALIZED Then
        '���ʂ���������
        Call ModifyTestCondtitionSheet
    End If
    
    '�I�u�W�F�N�g���J��
    Set m_IllmOptModZSet2Checker = Nothing
    Set m_IllmOptModZSet1Checker = Nothing
    Set m_IllmEscapeChecker = Nothing
    Set m_TesterChannelChecker = Nothing
    Set m_ApmuUbChecker = Nothing
    Set m_IllminatorChecker = Nothing
    
    Set m_colForceEnableCondition = Nothing
    Set m_colTestItems = Nothing
    
    '�O�̂���
    m_IsFirstSetIlluminator = True
    m_IsFirstSetUB = True
    m_IsFirstSetVoltatge = True
    m_IsFirstSetIllmEscape = True
    m_IsFirstSetIllmOptModZSet1 = True
    m_IsFirstSetIllmOptModZSet2 = True
    
    '��ԑJ��
    m_State = UNKNWON
    
End Sub

'�����ݒ�O�̃R���f�B�V�����擾
Public Sub CheckBeforeTestCondition(ByVal eMode As eTestCnditionCheck, ByRef pInfo As CSetFunctionInfo)
    
    '�������ς݂łȂ��ꍇ�͂����ɔ�����
    If m_State <> INITIALIZED Then
        Exit Sub
    End If
    
    TheHdw.StartStopwatch
    'ForceEnable���`�F�b�N
    If IsForceEnableTestCondition(pInfo.ConditionName) Then
        Exit Sub
    End If
    Dim sTime As Single
        
    '���[�h�ɂ���ČĂԃA�C�e�������߂�
    Dim strTemp As String
    Select Case eMode
        Case TCC_ILLUMINATOR
            If m_IsFirstSetIlluminator Then
                m_IsFirstSetIlluminator = False
                Exit Sub
            End If
            m_IllminatorChecker.CheckBeforeCondition
        Case TCC_APMU_UB
            If m_IsFirstSetUB Then
                m_IsFirstSetUB = False
                Exit Sub
            End If
            m_ApmuUbChecker.CheckBeforeCondition
        Case TCC_SETVOLTAGE
            If m_IsFirstSetVoltatge Then
                m_IsFirstSetVoltatge = False
                Exit Sub
            End If
            m_TesterChannelChecker.SetOperationMode (EEE_AUTO_TESTER_CHECKER_WITH_DCHAN)
            m_TesterChannelChecker.CheckBeforeCondition
        Case TCC_TESTER_CHANNELS
            If IsSetDigitalChannnel(pInfo) Then
               m_TesterChannelChecker.SetOperationMode (EEE_AUTO_TESTER_CHECKER_WITH_DCHAN)
            Else
                m_TesterChannelChecker.SetOperationMode (EEE_AUTO_TESTER_CHECKER_WITHOUT_DCHAN)
            End If
            m_TesterChannelChecker.CheckBeforeCondition
            
        Case TCC_ILLUMINATOR_ESCAPSE
            If m_IsFirstSetIllmEscape Then
                m_IsFirstSetIllmEscape = False
                Exit Sub
            End If
            m_IllmEscapeChecker.SetEndPosition GetFirstOptSetSameCategory(pInfo)
            m_IllmEscapeChecker.CheckBeforeCondition
            
        Case TCC_ILLUMINATOR_MODZ1
            If m_IsFirstSetIllmOptModZSet1 Then
                m_IsFirstSetIllmOptModZSet1 = False
                Exit Sub
            End If
            m_IllmOptModZSet1Checker.SetEndPosition pInfo.Arg(1)
            m_IllmOptModZSet1Checker.CheckBeforeCondition
        
        Case TCC_ILLUMINATOR_MODZ2
            If m_IsFirstSetIllmOptModZSet2 Then
                m_IsFirstSetIllmOptModZSet2 = False
                Exit Sub
            End If
            m_IllmOptModZSet2Checker.SetEndPosition pInfo.Arg(1)
            m_IllmOptModZSet2Checker.CheckBeforeCondition
            
    End Select
    
    sTime = TheHdw.ReadStopwatch
    Select Case eMode
        Case TCC_ILLUMINATOR
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumCheck Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_APMU_UB
            TheExec.Datalog.WriteComment pInfo.ConditionName & " APMU_UB Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_SETVOLTAGE
            TheExec.Datalog.WriteComment pInfo.ConditionName & " SetVoltage Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_TESTER_CHANNELS
            TheExec.Datalog.WriteComment pInfo.ConditionName & " Tester_Channnel Before " & pInfo.FunctionName & " " & CStr(sTime * 1000) & " " & pInfo.Arg(0)
        Case TCC_ILLUMINATOR_ESCAPSE
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumEscapse Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_ILLUMINATOR_MODZ1
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumModZSet1 Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_ILLUMINATOR_MODZ2
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumModZSet2 Before " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
    End Select
    
    '��ԑJ��
    m_State = CHECKED_BEFORE
    
    Exit Sub
    
ErrorHandler:
    m_State = UNKNWON

End Sub

'�����ݒ��̃R���f�B�V�����擾�A�Ӗ��̂�������ݒ肩�m�F
Public Sub CheckAfterTestCondition(ByVal eMode As eTestCnditionCheck, ByRef pInfo As CSetFunctionInfo)
    
    '�����ݒ�O�̏�Ԃ��擾���Ă��Ȃ��ꍇ�͂����ɔ�����
    If m_State <> CHECKED_BEFORE Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler:
    
    TheHdw.StartStopwatch
    
     'Before������Ȃ���Ύ��s�ł��Ȃ��̂ŁA�����ł͍s��Ȃ�
'    'ForceEnable���`�F�b�N
'    If IsForceEnableTestCondition(pInfo.ConditionName) Then
'        Exit Sub
'    End If
    
    Dim sTime As Single
    Dim strTemp As String
    Dim IsValid As Boolean
    IsValid = True
    
    '���[�h�ɂ���ČĂԃA�C�e�������߂�
    Select Case eMode
        Case TCC_ILLUMINATOR
            IsValid = m_IllminatorChecker.CheckAfterCondition
        Case TCC_APMU_UB
            IsValid = m_ApmuUbChecker.CheckAfterCondition
        Case TCC_SETVOLTAGE
            IsValid = m_TesterChannelChecker.CheckAfterCondition
        Case TCC_TESTER_CHANNELS
            IsValid = m_TesterChannelChecker.CheckAfterCondition
    End Select
    
    '���ʎq�̍쐬
    Dim strIdenfier As String
    strIdenfier = GetTestConditionIdenfier(pInfo)
    
    'Item��T��
    Dim IsFound As Boolean
    Dim obj As CEeeAuto_TestConditionItem
    IsFound = False
    For Each obj In m_colTestItems
        If (obj.GetTestConditionIdenfier = strIdenfier) Then
            IsFound = True
            Exit For
        End If
    Next
    
    If IsFound Then
        Call obj.SetValideCodition(IsValid)
    End If
  
    sTime = TheHdw.ReadStopwatch
    Select Case eMode
        Case TCC_ILLUMINATOR
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumCheck After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_APMU_UB
            TheExec.Datalog.WriteComment pInfo.ConditionName & " APMU_UB After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_SETVOLTAGE
            TheExec.Datalog.WriteComment pInfo.ConditionName & " SetVoltage After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_TESTER_CHANNELS
            TheExec.Datalog.WriteComment pInfo.ConditionName & " Tester_Channnel After " & pInfo.FunctionName & " " & CStr(sTime * 1000) & " " & pInfo.Arg(0)
    End Select

    '��ԑJ��
    m_State = INITIALIZED
    
    Exit Sub
    
ErrorHandler:
    m_State = UNKNWON
  
End Sub

'���e:
'   �����ݒ�}�N�����ɂ��̏����ݒ肪�ȗ����邩�Ԃ�
'   �Ώۊ֐���
'       OptEscape
'       OptModOrModZ1
'       OptModOrModZ2
'
'���l:
'�@�@OptEscape, OptModOrModZ1, OptModOrModZ2�͌Ăԏu�Ԃɏȗ����ׂ����ǂ������܂�
'�@�@�ȗ��\�Ȃ̂Ƀ��\�b�h���s����ƁA���̏����ȗ��I�u�W�F�N�g�����������삵�Ȃ��B
'�@�@���Ƃ��Αޔ�s�v�Ȃ̂�OptEscape�����s����ƁA�ŏI�I�ȍs���悪�����ꍇ�ł�
'�@�@OptModOrModZ2���ȗ��s�Ɣ��f���Ă��܂��B��̗�Ŏ�����
' �@�@���ݒn�@DOWN, �ŏI�ړI�n�@DOWN���Ƃ����
'�@�@  OptEscape�őޔ��ʒu(UP)�ֈړ�
'�@�@  OptModOrModZ1�őޔ�������Ȃ������̈ړ�
'�@�@  OptModOrModZ2�ōŏI�ʒu(DOWN)�Ɉړ�
'�@�@OptModOrModZ2��UP��DOWN�ֈʒu�ύX������̂ŁA�ȗ��s�Ɣ��f����B
'�@�@�Ȃ̂�OptEscape�͏ȗ��\�Ɣ��f�����ꍇ�A���s��Skip����K�v������B
Public Function IsValidTestCondition(ByVal eMode As eTestCnditionCheck, ByRef pInfo As CSetFunctionInfo) As Boolean

    '�����ݒ�O�̏�Ԃ��擾���Ă��Ȃ��ꍇ�͂����ɔ�����
    If m_State <> CHECKED_BEFORE Then
        IsValidTestCondition = True
        Exit Function
    End If
    
    On Error GoTo ErrorHandler:
    
    Dim sTime As Single
    Dim IsValid As Boolean
    IsValid = True
    
    '���[�h�ɂ���ČĂԃA�C�e�������߂�
    Select Case eMode
        Case TCC_ILLUMINATOR_ESCAPSE
            IsValid = m_IllmEscapeChecker.CheckAfterCondition
        Case TCC_ILLUMINATOR_MODZ1
            IsValid = m_IllmOptModZSet1Checker.CheckAfterCondition
        Case TCC_ILLUMINATOR_MODZ2
            IsValid = m_IllmOptModZSet2Checker.CheckAfterCondition
    End Select
   
    '���ʎq�̍쐬
    Dim strIdenfier As String
    strIdenfier = GetTestConditionIdenfier(pInfo)
    
    'Item��T��
    Dim IsFound As Boolean
    Dim obj As CEeeAuto_TestConditionItem
    IsFound = False
    For Each obj In m_colTestItems
        If (obj.GetTestConditionIdenfier = strIdenfier) Then
            IsFound = True
            Exit For
        End If
    Next
    
    If IsFound Then
        Call obj.SetValideCodition(IsValid)
    End If
  
    sTime = TheHdw.ReadStopwatch
    Select Case eMode
        Case TCC_ILLUMINATOR_ESCAPSE
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumEscapse After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_ILLUMINATOR_MODZ1
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumModZSet1 After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
        Case TCC_ILLUMINATOR_MODZ2
            TheExec.Datalog.WriteComment pInfo.ConditionName & " IllumModZSet2 After " & pInfo.Arg(0) & " " & CStr(sTime * 1000)
    End Select

    '��ԑJ��
    m_State = INITIALIZED
    
    IsValidTestCondition = IsValid

    Exit Function
    
ErrorHandler:
    m_State = UNKNWON

End Function

Private Sub ModifyTestCondtitionSheet()
    
On Error GoTo ErrorHandler:
    
    '�V�[�g�̎擾
    Dim mySht As Worksheet
    Set mySht = ThisWorkbook.Worksheets(TheCondition.TestConditionSheet)

    '�f�t�H���g�J�����̈ʒu�擾
    Dim lDefaultColumn As Long
    lDefaultColumn = mySht.Range(DEFAULT_ENABLES_START).Column
    
    '�l����������擾����
    Dim myrange As Range
    Dim lRefRow As Long, lRefColumn As Long
    lRefRow = mySht.Range(TESTCONDITION_ITEM_START).Row
    lRefColumn = mySht.Range(TESTCONDITION_ITEM_START).Column
    Set myrange = mySht.Range(mySht.Cells(lRefRow, lDefaultColumn), mySht.Cells(lRefRow + m_colTestItems.Count - 1, lDefaultColumn))
 
    Dim aryTemp() As Variant
    aryTemp = myrange.Value
    
    'false�̂Ƃ���̂ݏ㏑��
    Dim obj As CEeeAuto_TestConditionItem
    Dim i As Long
    i = 1
    For Each obj In m_colTestItems
        If obj.IsEnable Then
            aryTemp(i, 1) = "enable"
        Else
            aryTemp(i, 1) = "disable"
        End If
        i = i + 1
    Next
   
    '�l��߂�
    myrange.Value = aryTemp
    Erase aryTemp
    
    
    Set mySht = Nothing
    
    Exit Sub
ErrorHandler:
    Set mySht = Nothing


End Sub


Public Sub SetForceEnableTestCondition(ByVal strCondition As String)

    '�����ݒ������TRUE
    Dim obj As CEeeAuto_TestConditionItem
    For Each obj In m_colTestItems
        If (obj.ConditionName = strCondition) Then
            Call obj.SetValideCodition(True)
        End If
    Next

    '�L�[���d�Ȃ�ƃG���[�ɂȂ邪�A�d�Ȃ��Ă���Ƃ������Ƃ�
    '���ł�ForceEnable������Ă���̂ŉ��߂Ēǉ�����K�v���Ȃ�
On Error Resume Next
    Call m_colForceEnableCondition.Add(strCondition, strCondition)
On Error GoTo 0

End Sub

Private Function IsExecuteSetDefault(ByRef mySht As Worksheet) As Boolean

    If mySht Is Nothing Then
        IsExecuteSetDefault = False
        Exit Function
    End If
        
    '���݂̐ݒ���擾
    Dim strCurSetting As String
    strCurSetting = mySht.Range(CURRENT_SETTING)
     
    'Default�ݒ�ȊO�̓G���[
    If strCurSetting <> DEFAULT_LABEL Then
        IsExecuteSetDefault = False
    End If
    
    IsExecuteSetDefault = True
 
End Function

Private Sub ReadTestCondition(ByRef mySht As Worksheet)

    Const ARG_START As Long = 3

    '�G���[�`�F�b�N
    If mySht Is Nothing Then
        Exit Sub
    End If
    
    'TestCondition�̑S����z��Ɋi�[
    Dim aryTestConditions As Variant
    aryTestConditions = mySht.Range(mySht.Range(TESTCONDITION_ITEM_START), _
                        mySht.Cells.SpecialCells(xlCellTypeLastCell))
    
    '�f�t�H���g�J�����̔z��ł̈ʒu�擾
    Dim lDefaultColumn As Long
    lDefaultColumn = mySht.Range(DEFAULT_ENABLES_START).Column - mySht.Range(TESTCONDITION_ITEM_START).Column + 1
    
    '�z��̐�΍��W�ɑ΂���I�t�Z�b�g�s���擾
    Dim lOffsetRow As Long
    lOffsetRow = mySht.Range(TESTCONDITION_ITEM_START).Row
    
    
    'TestCondition�̎擾 TheCondition����擾���Ȃ��͍̂s�ԍ����킩��Ȃ�����
    Dim i As Long, j As Long
    Dim tempItem As CEeeAuto_TestConditionItem
    Dim strConditionName As String
    Dim strFuncName As String
    Dim lArgCount As Long
    Dim aryArg(9) As Variant
    Dim lRow As Long
    Dim IsEnable As Boolean
    For i = 1 To UBound(aryTestConditions, 1)
    
        '�󔒃Z���������ꍇ�͔�����
        If (IsEmpty(aryTestConditions(i, 1))) Then
            Exit For
        End If
        
        '�p�����[�^�̓ǂݍ���
        strConditionName = aryTestConditions(i, 1)
        strFuncName = aryTestConditions(i, 2)
        If (aryTestConditions(i, lDefaultColumn) = "enable") Then
            IsEnable = True
        Else
            IsEnable = False
        End If
        j = ARG_START
        While ((Not IsEmpty(aryTestConditions(i, j))) And (aryTestConditions(i, j) <> "#EOP") Or j >= ARG_START + ARG_MAX)
            aryArg(j - ARG_START) = aryTestConditions(i, j)
            j = j + 1
        Wend
        lArgCount = j - ARG_START
        lRow = lOffsetRow + i - 1
        
        '�I�u�W�F�N�g�𐶐��A�R���N�V�����ɒǉ�
        Set tempItem = New CEeeAuto_TestConditionItem
        Call tempItem.SetParams(strConditionName, strFuncName, lArgCount, aryArg, lRow, IsEnable)
        m_colTestItems.Add tempItem
        Set tempItem = Nothing
        
    Next i

End Sub

'���ʎq�̍쐬�@���̃��W���[����p
Private Function GetTestConditionIdenfier(ByRef pInfo As CSetFunctionInfo)

    Dim aryArg(ARG_MAX - 1) As Variant
    
    Dim i As Long
    With pInfo
        For i = 0 To .ArgParameterCount - 1
            aryArg(i) = .Arg(i)
        Next
        GetTestConditionIdenfier = GetTestConditionIdenfier_impl(.ConditionName, .FunctionName, .ArgParameterCount, aryArg)
    End With
    
End Function
 
'���ʎq�̍쐬�����ʂł�点�������߁A�ʊ֐��ŊO�ɏo��
Public Function GetTestConditionIdenfier_impl(ByVal strCndName As String, ByVal strFunName As String, ByVal lCount As Long, ByRef aryArg() As Variant) As String

    Dim strIdenfier As String
    Dim i As Long
    
    strIdenfier = strCndName & "_" & strFunName
    
    For i = 0 To lCount - 1
        strIdenfier = strIdenfier & "_" & CStr(aryArg(i))
    Next i

    GetTestConditionIdenfier_impl = strIdenfier

End Function

'�����ȗ��@�\�ΏۊO�̃R���f�B�V�������m�F����
Private Function IsForceEnableTestCondition(ByVal strConditionName As String) As Boolean

    If m_colForceEnableCondition.Count = 0 Then
        IsForceEnableTestCondition = False
    End If
        
    Dim obj As Variant
    
    For Each obj In m_colForceEnableCondition
        If obj = strConditionName Then
            IsForceEnableTestCondition = True
            Exit Function
        End If
    Next obj
    
    IsForceEnableTestCondition = False
    
    Exit Function
    
End Function

'�f�W�^���s���͎�荞�݂��x�����߁ADCHAN���܂܂Ȃ������ݒ�̏ꍇ�݂͂ɂ��������Ȃ�
'���̂��ߊ֐����ɂ���Ė₢���킹���s���ADCHAN���g�����m�F���A�g�p���Ȃ��Ȃ猩�ɍs���Ȃ�
Private Function IsSetDigitalChannnel(ByRef pInfo As CSetFunctionInfo) As Boolean
       
    Dim chanType As chtype

    Select Case pInfo.FunctionName
        Case "FW_DisconnectPins"
            chanType = GetChanType(pInfo.Arg(0))
            If (chanType = chAPMU) Or (chanType = chDPS) Then
                IsSetDigitalChannnel = False
                Exit Function
            End If
        Case "FW_SetFVMI"
            chanType = GetChanType(pInfo.Arg(0))
            If (chanType = chAPMU) Or (chanType = chDPS) Then
                IsSetDigitalChannnel = False
                Exit Function
            End If
    End Select
    
    IsSetDigitalChannnel = True

End Function

'�`�����l���^�C�v�̌���
Private Function GetChanType(ByVal PinList As String) As chtype

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

'�{���͂���Ȃ��Ƃ������Ȃ����ǁAESCAPE�ōŏI�I�ɂǂ��Ɉړ����邩�����œn��Ȃ��̂�
'�����Œ��ׂ�
Private Function GetFirstOptSetSameCategory(ByRef pInfo As CSetFunctionInfo) As String
    
    GetFirstOptSetSameCategory = ""
    '�s�ԍ��s�v�̂��߁A���ƂȂ���TheCondition�̏����g��
    Dim myCol As Collection
    Set myCol = TheCondition.GetCloneConditionInfo(pInfo.ConditionName)
        
    '�w�肵���R���f�B�V�����O���[�v�̒��ōŏ��Ɍ�������
    'FW_OptModOrModZ1�̈����𗘗p����B������Ă΂�Ă������Ȃ��듮�삷��
    Dim obj As CSetFunctionInfo
    
    For Each obj In myCol
        If obj.FunctionName = "FW_OptModOrModZ1" Or _
           obj.FunctionName = "FW_OptModOrModZ2" Or _
           obj.FunctionName = "FW_OptEscape" Or _
           obj.FunctionName = "FW_OptSet" Or _
           obj.FunctionName = "FW_OptSet_Test" Then
            GetFirstOptSetSameCategory = obj.Arg(1)
            Exit For
        End If
    Next obj

    Set myCol = Nothing
End Function

