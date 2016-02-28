Attribute VB_Name = "XLibPatternMod"
'�T�v:
'   �p�^�[�����䃉�C�u�����Q
'
'�ړI:
'   �T�F�p�^�[���V�[�g����f�[�^���擾�����[�h����
'   �U�F�e���_�C���p�^�[������API���������̋@�\�u���b�N�Ƃ��Đ؂�o��
'
'   Revision History:
'   Data        Description
'   2010/07/27�@��TOPT�t���[�����[�N�����ɂ��d�l�ύX
'               �@StartPattern/RunPattern/StartStopPattern�֐��̈����փ^�C�~���O����ǉ�
'   2010/07/29�@�������W���[���̃T�u���[�`��Break(), DebugMsg()�̗��p��~
'                 �t�@���N�V����LoadPatternFile()����Ԓl��Ԃ���悤�ɕύX
'   2012/12/20  H.Arikawa
'               StopPattern��ҏW�B�e�X�^�[�^�C�v�ɉ�����Halt�̓����ύX����B
'               LoadPatternFile��ҏW�BPatGrps�V�[�g�ǂݍ��ݕ��ύX�B
'
'�쐬��:
'   0145206097

Option Explicit

Public Function LoadPatternFile() As Long
'���e:
'   PatGrps�V�[�g����p�^���O���[�v��ǂݍ��݃��[�h����
'
'�p�����[�^:
'
'�Ԓl:
'   ����:0(TL_SUCCESS)�A���s:1(TL_ERROR)
'
'���ӎ���:
'   ���[�N�V�[�g"PatGrps"��������Ȃ��Ƃ��A
'   �܂��̓��[�N�V�[�g����̃f�[�^�擾���ɃG���[�����������ꍇ��
'   ���b�Z�[�W�{�b�N�X�Ōx����\�����ATL_ERROR��Ԃ��܂��B

    Const DATA_SHEET_NAME = "PatGrps"
    Const PATTERN_GROUP = "GroupName"

    Dim targetWorkSheet As Worksheet
    Dim patGroupName As Range
    Dim tsbName As Range
    Dim tsbSheetName As String
    Dim rowIndex As Long

    Call StopPattern
    TheHdw.Digital.Patterns.UnloadAll
    TheHdw.Digital.Patgen.TimeoutEnable = False

    On Error GoTo ErrHandler
    Set targetWorkSheet = getWorkSheet(DATA_SHEET_NAME)

    With targetWorkSheet
        
        Set patGroupName = .Range(PATTERN_GROUP)
        Set tsbName = .Range("E3")
        
        rowIndex = patGroupName.Row + 1

        Do While .Cells(rowIndex, patGroupName.Column) <> ""
            tsbSheetName = .Cells(rowIndex, tsbName.Column)
            TheHdw.Digital.Timing.Load (tsbSheetName)
            TheHdw.Digital.Patterns.pat(.Cells(rowIndex, patGroupName.Column)).Load
            rowIndex = rowIndex + 1
        Loop

    End With
    LoadPatternFile = TL_SUCCESS
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation Or vbOKOnly, "Error"
    TheHdw.Digital.Patterns.UnloadAll
    TheHdw.Digital.Patgen.TimeoutEnable = False
    LoadPatternFile = TL_ERROR
End Function

Public Sub StartPattern( _
    ByVal patGroupName As String, _
    Optional ByVal startLabel As String = "START", _
    Optional timeSetName As String = "", _
    Optional categoryName As String = "", _
    Optional selectorName As String = "" _
)
'���e:
'   �p�^�[�����o�[�X�g����
'
'�p�����[�^:
'[patGroupName] In  �p�^�[���O���[�v��
'[startLabel]   In  �X�^�[�g���x��
'[timeSetName]  In  �^�C�~���O�Z�b�g��
'[categoryName] In�@�J�e�S����
'[selectorName] In  �Z���N�^��
'
'���ӎ���:
'   �o�[�X�g�I����҂����ɐ��䂪�v���O�����ɖ߂�
'

    With TheHdw.Digital
        .Timing.Load timeSetName, categoryName, selectorName
        .Patterns.pat(patGroupName).Start startLabel
    End With

End Sub

Public Sub RunPattern( _
    ByVal patGroupName As String, _
    Optional ByVal startLabel As String = "START", _
    Optional timeSetName As String = "", _
    Optional categoryName As String = "", _
    Optional selectorName As String = "" _
)
'���e:
'   �p�^�[�����o�[�X�g����
'
'�p�����[�^:
'[patGroupName] In  �p�^�[���O���[�v��
'[startLabel]   In  �X�^�[�g���x��
'[timeSetName]  In  �^�C�~���O�Z�b�g��
'[categoryName] In�@�J�e�S����
'[selectorName] In  �Z���N�^��
'
'���ӎ���:
'   �o�[�X�g�I����҂��Ă��琧�䂪�v���O�����ɖ߂�
'

    With TheHdw.Digital
        .Timing.Load timeSetName, categoryName, selectorName
        .Patterns.pat(patGroupName).Run startLabel
    End With

End Sub

Public Sub StartStopPattern( _
    ByVal patGroupName As String, _
    ByVal startLabel As String, _
    ByVal stopLabel As String, _
    Optional timeSetName As String = "", _
    Optional categoryName As String = "", _
    Optional selectorName As String = "" _
)
'���e:
'   �p�^�[�����o�[�X�g����
'
'�p�����[�^:
'[patGroupName] In  �p�^�[���O���[�v��
'[startLabel]   In  �X�^�[�g���x��
'[stopLabel]    In  �X�g�b�v���x��
'[timeSetName]  In  �^�C�~���O�Z�b�g��
'[categoryName] In�@�J�e�S����
'[selectorMame] In  �Z���N�^��
'
'���ӎ���:
'   �w��̃X�g�b�v���x����HALT��}����A�w��̃X�^�[�g���x������o�[�X�g����
'

    With TheHdw.Digital
        .Timing.Load timeSetName, categoryName, selectorName
        .Patterns.pat(patGroupName).StartStop startLabel, stopLabel
    End With

End Sub

Public Sub StopPattern()
'���e:
'   �p�^�[���o�[�X�g���I������
'
'�p�����[�^:
'
'���ӎ���:
'        �f�R�[�_�Ŏg�p����p�^�[���̃p�^�[���o�[�X�g�I�����s���B
'
    With TheHdw.Digital.Patgen
    
        If .IsRunningAnySite = True Then
            .Ccall = True
            .HaltWait
            .Ccall = False
        End If

    End With

End Sub

Public Sub StopPattern_Halt()
'���e:
'   �p�^�[���o�[�X�g���I������
'
'�p�����[�^:
'
'���ӎ���:
'        �e�X�^�[��IP750�̎��́AHalt�~�߂��Ȃ��悤��If���ŕ���B
'

    With TheHdw.Digital.Patgen
    
        If TesterType = "IP750" Then
            If .IsRunningAnySite = True Then
                .Ccall = True
                .HaltWait
                .Ccall = False
            End If
        Else
            If .IsRunningAnySite = True Then
                   .Halt
            End If
        End If

    End With

End Sub

Public Sub SetTimeOut( _
    Optional ByVal runOutStatus As Boolean = False, _
    Optional ByVal runOutTime As Long = 5 _
)
'���e:
'   �p�^�[���̃^�C���A�E�g������ݒ肷��
'
'�p�����[�^:
'[runOutStatus] In  �^�C���A�E�g�X�e�[�^�X
'[runOutTime]   In  �^�C���A�E�g����
'
'���ӎ���:
'

    With TheHdw.Digital.Patgen
        .TimeoutEnable = runOutStatus
        .TIMEOUT = runOutTime
    End With

End Sub

'{ XLibDcMod�ɓ��l�̊֐�����
Private Function getWorkSheet( _
    ByRef SheetName As String _
) As Worksheet

    On Error GoTo errMsg

    Set getWorkSheet = Worksheets(SheetName)

    Exit Function

errMsg:

    Call Err.Raise(Err.Number, Err.Source, "Not " & SheetName & " Sheet Exist !")
    '���{��G���[���b�Z�[�W�o��
'    Call Err.Raise(Err.Number, Err.Source, sheetName & " �V�[�g������܂���")

    Set getWorkSheet = Nothing

End Function
'}

Private Function TimingLoad_f() As Long
    Dim myInstance As String
    Dim myTSBName As String
    
    If First_Exec = 0 Then
        myInstance = TheExec.DataManager.InstanceName
        myTSBName = Mid(myInstance, InStr(myInstance, "_") + 1, Len(myInstance))
        TheHdw.Digital.Timing.Load myTSBName
    End If
    
End Function
Public Sub PatGrpsColorMake()

    With Worksheets("PatGrps").Range("E3")
        .Interior.color = RGB(0, 0, 255)
        .Font.color = vbWhite
    End With
    
End Sub
