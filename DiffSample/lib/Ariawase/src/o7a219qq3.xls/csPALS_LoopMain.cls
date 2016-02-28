VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_LoopMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' �N���X��  �F  csPALS_LoopMain.cls
' �T�v      �F  LOOP���莞�Ɏg�p����������N���X
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/16�@�V�K�쐬        K.Sumiyashiki
'==========================================================================

Private xLoopCategory() As csPALS_LoopCategoryParams        'csPALS_LoopCategoryParams�N���X�̒�`

Public CategoryInfoList As New Collection                   '������(OF,ML,etc..)�ƃC���f�b�N�X��R�t����ׂ̃R���N�V����
Public CategoryKeyList  As New Collection                   '�L�[�̏d���`�F�b�N�ׂ̈̃R���N�V����


'*********************************************************************************
' �֐���    �F  Class_Initialize
' �T�v      �F  �C���X�^���X�������̏�����
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()

    '�R�t���p�R���N�V����������
    Call sub_InitCollection(CategoryInfoList)
    Call sub_InitCollection(CategoryKeyList)
    
    'TestCondition�̃f�[�^�Ǎ�
    Call sub_ReadCategoryInfo
    
    'TestCondition�̃f�[�^�ɕs�����Ȃ����`�F�b�N
    Call sub_CheckCategoryInfo

End Sub


'*********************************************************************************
' �֐���    �F  Class_Terminate
' �T�v      �F  �C���X�^���X�J�����̏���
' ����      �F  �Ȃ�
' ���l      �F  �Ȃ�
' �X�V����  �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Terminate()

    '�R�t���p�R���N�V����������
    Call sub_InitCollection(CategoryInfoList)
    Call sub_InitCollection(CategoryKeyList)

End Sub


'*********************************************************************************
' �v���p�e�B��  �F  LoopCategory
' �T�v          �F  �J�e�S���ԍ��ɑΉ�����LoopCategory�I�u�W�F�N�g��Ԃ��v���p�e�B
' ����          �F  index :�J�e�S�����������C���f�b�N�X
' ���l          �F  �ǂݎ���p
' �X�V����      �F  Rev1.0      2010/08/18�@�V�K�쐬        K.Sumiyashiki
'*********************************************************************************
Public Property Get LoopCategory(ByVal index As Integer) As csPALS_LoopCategoryParams

    Set LoopCategory = xLoopCategory(index)

End Property


'********************************************************************************************
' ���O: sub_ReadCategoryInfo
' ���e: TestCondition�̃p�����[�^��csLoopParams�N���X�Œ�`���ꂽLoopParams�z��Ɋi�[�B
'       LoopParams�͏��������̔z�񂪓��I�ɍ쐬�����B
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/17�@�V�K�쐬   K.Sumiyashiki
' �X�V�����F Rev1.1      2011/06/02�@�����ǉ�   K.Sumiyashiki
'                                    ��Eee-JOB�Ή�
'********************************************************************************************
Private Sub sub_ReadCategoryInfo()

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_ReadCategoryInfo

    Dim TC_y As Long            'Y�����̃Z���ʒu���i�[����ϐ�
    Dim strbuf As String        '���������i�[����ϐ��@��)"OF"�A"HL"��
    
    'TestCondition�V�[�g�̑��݃`�F�b�N
    If Not sub_SheetNameCheck(TESTCONDITION) Then

'>>>2011/06/02 K.SUMIYASHIKI ADD
        If sub_SheetNameCheck(ACQTBLSHTNAME) Then
Jump_Set_ACQTBL:
            '��������LoopParams���Ē�`
            ReDim xLoopCategory(1)
            
            'csLoopParams�N���X�̃C���X�^���X����
            Set xLoopCategory(1) = New csPALS_LoopCategoryParams
            
            strbuf = "DummyData"
            
            '�J�e�S���������
            xLoopCategory(1).category = strbuf
            
            '���������L�[�ɁALoopParams�̔z��ԍ����R���N�V�����ɒǉ�
            CategoryInfoList.Add Item:=1, key:=strbuf
            
            '�L�[�d���`�F�b�N�ׂ̈ɁA��������l�ƃL�[�ɂ����R���N�V�����쐬
            CategoryKeyList.Add Item:=strbuf, key:=strbuf
        
            'Wait��Average�f�[�^�̓���
            With xLoopCategory(CategoryInfoList(strbuf))
                .Average = 1
                .WAIT = 1
            End With
    
            Exit Sub
'<<<2011/06/02 K.SUMIYASHIKI ADD
        Else
            Call sub_errPALS("TestCondition not found at 'sub_ReadCategoryInfo'", "2-4-01-8-37")
            Exit Sub
        End If
    End If
    
    With Worksheets(TESTCONDITION)
    
        'TestCondition�V�[�g��Y�����̃X�^�[�g�Z���ʒu��ݒ�
        TC_y = TC_START_Y
        
        'Condition Name�̍Ō�܂ŌJ��Ԃ�
        Do While Len(.Cells(TC_y, TC_CONDINAME_X).Value)
            
            'Sw_Node����v����s�̂ݏ���
            If Sw_Node = .Cells(TC_y, TC_SWNODE_X).Value Then
            
                'strBuf�ɏ���������
                strbuf = .Cells(TC_y, TC_CONDINAME_X).Value
                
                '�R���N�V�����Ƀf�[�^���Ȃ��ꍇ(1���ږ�)�̏���
                If CategoryInfoList.Count = 0 Then
                    'csLoopParams�N���X�̃C���X�^���X�𐶐����A�f�[�^��ǉ�
                    Call sub_Add_CategoryInfo(strbuf, TC_y)
                Else
                    '���ɏ������R���N�V�����ɒǉ�����Ă��邩�̃`�F�b�N
                    If key_check(CategoryKeyList, strbuf) Then
                        '�ǉ�����Ă����ꍇ�A�f�[�^��ǉ�
                        Call sub_InputTestCondition(strbuf, TC_y)
                    Else
                        'csLoopParams�N���X�̃C���X�^���X�𐶐����A�f�[�^��ǉ�
                        Call sub_Add_CategoryInfo(strbuf, TC_y)
                    End If
                End If
            End If
        '���̍s��
        TC_y = TC_y + 1
        Loop
    End With

    If CategoryInfoList.Count = 0 Then
        If sub_SheetNameCheck(ACQTBLSHTNAME) Then
            GoTo Jump_Set_ACQTBL
        End If
        Call sub_errPALS("NotFound Info@Sheet[" & TESTCONDITION & "], Node[" & Sw_Node & "] at 'sub_ReadCategoryInfo'", "2-4-01-5-38")
    End If

Exit Sub

errPALSsub_ReadCategoryInfo:
    Call sub_errPALS("Read IG-XL sheet error at 'sub_ReadCategoryInfo'", "2-4-01-0-39")

End Sub


'********************************************************************************************
' ���O: key_check
' ���e: �����l���R���N�V�����ɑ��݂��Ă��邩�`�F�b�N����
' ����: col:��������R���N�V����
'       query:�����Ώۂ̕�����
' �ߒl: �����l����=>True
'       �����l�Ȃ�=>False
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/17�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Function key_check(ByRef col As Collection, ByRef query As String) As Boolean

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSkey_check

    Dim Item As Variant
    '�R���N�V�����̍Ō�܂ŌJ��Ԃ�
    For Each Item In col
        '�����l���������ꍇ�ATrue��Ԃ��֐��𔲂���
        If Item = query Then
            key_check = True
            Exit Function
        End If
    Next
    '�����l�������ꍇ�AFalse��Ԃ�
    key_check = False

Exit Function

errPALSkey_check:
    Call sub_errPALS("key check error at 'csPALS_LoopMain -> key_check'", "2-4-02-0-40")
    
End Function


'********************************************************************************************
' ���O: sub_add_CategoryInfo
' ���e: csLoopParams�N���X�̃C���X�^���X�𐶐����A�f�[�^��ǉ�����B
'       �C���X�^���X�͏����������A���I�ɍ����B
'       CategoryInfoList�R���N�V�����́A�L�[���������ɁALoopParams�̔z��ԍ���Item�ɒǉ������B
'       CategoryKeyList�R���N�V�����́A�L�[�����ׂ̈ɍ쐬���Ă���B
' ����: strBuf:������
'       TC_y:�Z����Y�����̈ʒu
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/17�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_Add_CategoryInfo(ByRef strbuf As String, ByVal TC_y As Long)
    
    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALSsub_Add_CategoryInfo

    'LoopParams�̔z��ԍ��̐ݒ�
    Dim intItemCnt As Integer
    
    If CategoryInfoList.Count = 0 Then
        '�z�񂪖����ꍇ�A�z��ԍ���1�ɐݒ�
        '�R�t���p�̃R���N�V�����̃L�[��0��ݒ�o���Ȃ��ׁA1����J�n
        intItemCnt = 1
    Else
        '�z�񂪑��݂���ꍇ�A�z�񐔂��C���N�������g
        intItemCnt = UBound(xLoopCategory) + 1
    End If

    '��������LoopParams���Ē�`
    ReDim Preserve xLoopCategory(intItemCnt)
    
    'csLoopParams�N���X�̃C���X�^���X����
    Set xLoopCategory(intItemCnt) = New csPALS_LoopCategoryParams
    
    '�J�e�S���������
    xLoopCategory(intItemCnt).category = strbuf
    
    '���������L�[�ɁALoopParams�̔z��ԍ����R���N�V�����ɒǉ�
    CategoryInfoList.Add Item:=intItemCnt, key:=strbuf
    
    '�L�[�d���`�F�b�N�ׂ̈ɁA��������l�ƃL�[�ɂ����R���N�V�����쐬
    CategoryKeyList.Add Item:=strbuf, key:=strbuf
    
'    '�����������Z
    
    '�f�[�^����
    Call sub_InputTestCondition(strbuf, TC_y)

Exit Sub

errPALSsub_Add_CategoryInfo:
    Call sub_errPALS("Add CategoryInfo error at 'sub_Add_CategoryInfo'", "2-4-03-0-41")

End Sub


'********************************************************************************************
' ���O: sub_InputTestCondition
' ���e: TestCondition�V�[�g�̃f�[�^�擾
'       �f�[�^�̎�ނɂ���āA����ϐ��𕪂��Ă���
' ����: strBuf:������
'       TC_y:�Z����Y�����̈ʒu
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/17�@�V�K�쐬   K.Sumiyashiki
'            Rev1.1      2011/05/16�@�G���[�R�����g�C��   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_InputTestCondition(ByVal strbuf As String, ByVal TC_y As Long)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_InputTestCondition

    Dim strTmp As String        '�Z���f�[�^���ꎞ�ۑ�
    
    With xLoopCategory(CategoryInfoList(strbuf))
                
        '�Z���f�[�^���ꎞ�i�[
        strTmp = Worksheets(TESTCONDITION).Cells(TC_y, TC_PROCEDURENAME_X).Value
        
        Select Case strTmp
            'Wait�f�[�^�̏ꍇ
            Case SET_WAIT
                If .WAIT = -1 Then
                    '�f�[�^�������l�̏ꍇ�A�f�[�^����
                    .WAIT = Worksheets(TESTCONDITION).Cells(TC_y, TC_ARG1_X).Value
                Else
                    '�����f�[�^����������ꍇ�A�G���[�R�����g�\��
'>>>2011/5/12 K.SUMIYASHIKI CHANGE
'old                    Call sub_errPALS("Error : Wait data overlaps! of 'TestCondition sheet'")
                    Call sub_errPALS("Wait data overlaps! of 'TestCondition sheet'" & vbCrLf _
                                    & "Condition Name -> " & .category & vbCrLf _
                                    & "Error Row -> " & TC_y, "2-4-04-2-42")
'<<<2011/5/12 K.SUMIYASHIKI CHANGE
                End If
            'Average�f�[�^�̏ꍇ
            Case SET_AVERAGE
                If .Average = -1 Then
                    '�f�[�^�������l�̏ꍇ�A�f�[�^����
                    .Average = Worksheets(TESTCONDITION).Cells(TC_y, TC_ARG1_X).Value
                Else
'>>>2011/5/12 K.SUMIYASHIKI CHANGE
                    '�����f�[�^����������ꍇ�A�G���[�R�����g�\��
'old                    Call sub_errPALS("Error : Average data overlaps! of 'TestCondition sheet'")
                    Call sub_errPALS("Average data overlaps! of 'TestCondition sheet'" & vbCrLf _
                                    & "Condition Name -> " & .category & vbCrLf _
                                    & "Error Row -> " & TC_y, "2-4-04-2-43")
                End If
'<<<2011/5/12 K.SUMIYASHIKI CHANGE
            'Mode�f�[�^�̏ꍇ
            Case ACQUIRE_MODE
                If Len(.mode) = 0 Then
                    '�f�[�^�������l�̏ꍇ�A�f�[�^����
                    .mode = Worksheets(TESTCONDITION).Cells(TC_y, TC_ARG1_X).Value
                Else
'>>>2011/5/12 K.SUMIYASHIKI CHANGE
                    '�����f�[�^����������ꍇ�A�G���[�R�����g�\��
'old                    Call sub_errPALS("Error : Mode data overlaps! of 'TestCondition sheet'")
                    Call sub_errPALS("Mode data overlaps! of 'TestCondition sheet'" & vbCrLf _
                                    & "Condition Name -> " & .category & vbCrLf _
                                    & "Error Row -> " & TC_y, "2-4-04-2-44")
'<<<2011/5/12 K.SUMIYASHIKI CHANGE
                End If
            Case Else
                '��̕�����ȊO�̏ꍇ�̓X��
                'Eee-JOB�ő��̕�������g�p���Ă����
        End Select
    End With

Exit Sub

errPALSsub_InputTestCondition:
    Call sub_errPALS("Input TestCondition Data error at 'sub_InputTestCondition'", "2-4-04-0-45")

End Sub


'********************************************************************************************
' ���O: ChangeLoopParams
' ���e: �����l�̃o���c�L���������ۂɁATestCondition�V�[�g�̃p�����[�^�ύX�����{����֐�
' ����: strCategory    :�J�e�S����
'       strWaitOrAve   :Wait or Average������������
'       strValue       :�ύX�������l
' �ߒl: True           :�ύX���G���[�Ȃ�
'       False          :�ύX���G���[����
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/17�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Function ChangeLoopParams(ByVal strCategory As String, ByVal strWaitOrAve As String, ByVal strValue As String) As Boolean

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSChangeLoopParams

    Dim TC_y As Long                    '�e�X�g�R���f�B�V������Y�����̃Z���ʒu
    Dim Tmp_strWaitOrAve As String
    
    Select Case strWaitOrAve
    Case "Wait"
        Tmp_strWaitOrAve = SET_WAIT
    Case "Average"
        Tmp_strWaitOrAve = SET_AVERAGE
    Case Else
        MsgBox "Error : Program argument error at 'ChangeLoopParams'" & vbCrLf & "ErrCode.2-4-05-4-47", vbExclamation, PALS_ERRORTITLE
        ChangeLoopParams = False
        Exit Function
    End Select
    
    '�e�X�g�R���f�B�V�����V�[�g
    With Worksheets(TESTCONDITION)
    
        'TestCondition�V�[�g��Y�����̃X�^�[�g�Z���ʒu��ݒ�
        TC_y = TC_START_Y
        
        'Condition Name�̍Ō�܂ŌJ��Ԃ�
        Do While Len(.Cells(TC_y, TC_CONDINAME_X).Value) > 0
            
            'Condition Name��Procedure Name����v����ꍇ�A�Z���̒l���㏑��
            If (.Cells(TC_y, TC_CONDINAME_X).Value = strCategory) And (.Cells(TC_y, TC_PROCEDURENAME_X).Value = Tmp_strWaitOrAve) _
                And (.Cells(TC_y, TC_SWNODE_X).Value = Sw_Node) Then
                
                '�Z���̒l���㏑��
                .Cells(TC_y, TC_ARG1_X).Value = val(strValue)
                
                '�t���O��True�ɕύX
                ChangeLoopParams = True
                
                '�֐��𔲂���
                Exit Function
            End If
        '���̍s��
        TC_y = TC_y + 1
        Loop
    End With
    
    '��v���鍀�ڂ����������ꍇ�A�G���[���b�Z�[�W
    MsgBox "Error : '" & strCategory & "' not found at 'TestCondition sheet'." & vbCrLf & "         I can't change parameter." & vbCrLf & "ErrCode.2-4-05-4-48", vbExclamation, PALS_ERRORTITLE

Exit Function

errPALSChangeLoopParams:
    Call sub_errPALS("Change LoopParameter error at 'ChangeLoopParams'", "2-4-05-0-46")

End Function


'********************************************************************************************
' ���O: CategoryCount
' ���e: �J�e�S������Ԃ��֐�
' ����: �Ȃ�
' �ߒl: �J�e�S����
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/20�@�V�K�쐬   K.Sumiyashiki
'********************************************************************************************
Public Function CategoryCount() As Integer

On Error GoTo errPALSCategoryCount

    CategoryCount = UBound(xLoopCategory)

Exit Function

errPALSCategoryCount:
    Call sub_errPALS("Get category count error at 'csPALS_LoopMain -> CategoryCount'", "2-4-06-0-49")

End Function


'********************************************************************************************
' ���O: sub_CheckCategoryInfo
' ���e: TestCondition�V�[�g�ɋL�ڂ���Ă���p�����[�^�ɕs�����Ȃ����`�F�b�N����֐�
' ����: �Ȃ�
' �ߒl: �Ȃ�
' ���l    �F �Ȃ�
' �X�V�����F Rev1.0      2010/08/17�@�V�K�쐬   K.Sumiyashiki
'            Rev1.1      2011/03/21�@�����ύX   K.Sumiyashiki
'                                    �˃A�x���[�W�񐔃`�F�b�N������
'********************************************************************************************
Private Sub sub_CheckCategoryInfo()

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALSsub_CheckCategoryInfo

    Dim i As Long       '�J�e�S���[�����������[�v�J�E���^

    '�S�J�e�S�����J��Ԃ�
    For i = 1 To UBound(xLoopCategory)
        With xLoopCategory(i)
            '�A�x���[�W�񐔂�512��ȏ�łȂ����`�F�b�N
'>>>2011/03/21 K.SUMIYASHIKI COMMENT OUT
'            If .Average > 511 Then
'                Call sub_errPALS("Error : " & .Category & " average count is over 512.")
'<<<2011/03/21 K.SUMIYASHIKI COMMENT OUT
            If .Average < 1 Then
                Call sub_errPALS(.category & " average count is under 0.", "2-4-07-5-50")
            End If
            
            '�A�x���[�W or Wait�ݒ肪�V�[�g�ɋL�q����Ă��邩�`�F�b�N
            If .Average = -1 Or .WAIT = -1 Then
                Call sub_errPALS(.category & " Averge or Wait is not set.", "2-4-07-5-51")
            End If
            
            '���[�h�Ɋi�[����Ă���ϐ��������񂩂ǂ����`�F�b�N
            'IsNumeric�֐��˕�����:False�A���l:True
            If IsNumeric(.mode) Then
                '�ݒ肳��Ă���Average�񐔂����[�h�Ŏw�肳�ꂽ���l�ȉ����`�F�b�N
                If (.Average <> -1) And (.Average < .mode) Then
                    Call sub_errPALS(.category & "  average < Set Mode count.", "2-4-07-5-52")
                End If
                
                '�ݒ肳��Ă���Average�񐔂����[�h�Ŏw�肳�ꂽ���l�̔{���ɂȂ��Ă��邩�`�F�b�N
                If (.Average <> -1) And (.Average Mod .mode) Then
                    Call sub_errPALS(.category & "  average is not a Multiple of the 'Mode'.", "2-4-07-5-53")
                End If
            Else
                '���[�h���I�[�g�̏ꍇ�A�啶��AUTO�ɕϊ�
                If (UCase(.mode) = MODE_AUTO) Or (.mode = "") Then
                    .mode = MODE_AUTO
                End If
            End If
        End With
    Next i

Exit Sub

errPALSsub_CheckCategoryInfo:
    Call sub_errPALS("CheckCategoryInfo error at 'sub_CheckCategoryInfo'", "2-4-07-0-54")

End Sub