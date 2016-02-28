VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEeeAuto_IlluminatorChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �R���f�B�V�����`�F�b�J�@����������擾���W���[��
'
'�ړI:
'   ���̂Ƃ��̏����ݒ����������擾����
'�쐬��:
'   �ێR
'   �L�� �ҏW Ver1.1 2013/02/25
'
Option Explicit

Private Const NIKON = "NIKON"
Private Const INTERACTION = "InterAction"
Private Const STATUS_READ_COMMAND = "Q"
Private Const ILLUM_MAKER_LIST_WRKSHT_NAME = "IllumLst"
Private Const DEFAULT_ILLUM_GPIB_ADDR_NIKON = 8 'NIKON Illuminator GPIB Addr No.8
Private Const DEFAULT_ILLUM_GPIB_ADDR_IA = 2 'InterAction GPIB Addr No.2

Private Const DARK_INDENTIFIER As String = "N5A1200" 'DARK������������

Private hOptPort As Integer 'For GPIB Direct Command Send Addr

'�����̏�Ԃ������\���́i�C���^�[�A�N�V�����A�j�R�����p�j
Private Type type_Illminator
    strCondition As String
    '�������牺�̓C���^�[�A�N�V�����͎g�p���Ȃ�
    dblLux As Long
    lFLTurret As Long
    lShutter As Long
    lColor As Long
End Type
Implements IEeeAuto_TestConditionChecker

Private mIllumConditions(2) As type_Illminator

Private Sub IEeeAuto_TestConditionChecker_CheckBeforeCondition()

'���e:
'   �����������@�������[�J�̊m�F
'
'���ӎ���:
'   �������[�J�̊m�F
'   ���̎��_�ł̒l�̎擾
'
'�E�C���^�[�A�N�V����
'�@�@�����̃f�o�C�X�f�B�X�N���v�^�̈����͍čl�������B
'�@�@�u�v���Z�X�Ԃŋ��L����v�ƃ}�j���A���ɂ���̂ŁA����ɕ�����܂�����
'�@�@�ł͏���ɊJ���Ă悢�̂��H�Ƃ�����Ƃǂ��ɂ��ʖڂƏ����ĂȂ���....�B
'
'�E�j�R��
'�@�@�W�����C�u�����̃e���_�C���I�u�W�F�N�g�𗘗p���邽�߁A�f�o�C�X�̏������͂��Ȃ��B
'

    Erase mIllumConditions
        
    Dim wsheet As Worksheet
    
    If OptCond.IllumMaker = NIKON Then '2012/11/15 175Debug Arikawa
        Set wsheet = ThisWorkbook.Worksheets("Opt(NSIS)")
        '�j�R�������̓e���_�C���I�u�W�F�N�g����擾����
        '�e���_�C���I�u�W�F�N�g�̎��̂��p�u���b�N�錾����Ă��邽��
        'OptIni�������Ă���Ή�������K�v�͂Ȃ��B
    Else
        Set wsheet = ThisWorkbook.Worksheets("Opt(IA)")
        Call ibdev(0, GetOptGpibAddress, 0, 13, 1, &H13, hOptPort)
    End If
        
    mIllumConditions(0) = GetIllumCondtion
        
    Set wsheet = Nothing
    
End Sub
Private Function IEeeAuto_TestConditionChecker_CheckAfterCondition() As Boolean

    'Default��Valid
    IEeeAuto_TestConditionChecker_CheckAfterCondition = True
    
    mIllumConditions(1) = GetIllumCondtion
        
    If (IsEqualTwoCondition(mIllumConditions(0), mIllumConditions(1))) Then
        '�������ꏏ�������ꍇ InValid�Ƃ���
        IEeeAuto_TestConditionChecker_CheckAfterCondition = False
    End If
    
End Function

Private Sub IEeeAuto_TestConditionChecker_SetOperationMode(ByVal strMode As String)
    
    '���[�h�͂��̃N���X�ɂ͂Ȃ�
    
End Sub

Private Sub IEeeAuto_TestConditionChecker_SetEndPosition(ByVal strPos As String)

    '�ŏI�ʒu�͂��̃N���X�ɂ͂Ȃ�

End Sub
'=================================================================

Private Function GetIllumCondtion() As type_Illminator
'���e:
'   �������牞�����擾�������̂��̂�������
'
'���ӎ���:
'�E�C���^�[�A�N�V����
'�@�@������������������擾���邽�߂ɂ͓��삪�������ĂȂ��Ă͂Ȃ�Ȃ�
'    ���̂��߁u�����҂����Ȃ��v�ݒ�œ����ꂽ�Ƃ��̂݁A�����҂�������B
'
'�E�j�R��
'�@�@�e���_�C���I�u�W�F�N�g����擾�B������ۂ������������Ă���
'    �������������j�R���͕K�������҂������Ă���悤�ȋC�z������̂ŁA�����̓P�A���Ȃ��ł���

    Dim IllumCondition As String

    If OptCond.IllumMaker = NIKON Then
        Dim nikonInfo As NSIS_II_CONDITION
        nikonInfo = InitOptCondition
        
        With nikonInfo
            NSIS_II.ReadDevices .Slit, .NDFilter, .WedgeFilter, .color, .Shutter, .LCShutter, .Diffusion, .Pattern, .FNumberIris, .Mirror, .LED, .Pupil, .Illuminant, .ColorTemperature, .FNumberTurret
            IllumCondition = "S" & CStr(.Shutter) & "/L" & CStr(NSIS_II.Level) & "/C" & CStr(.color) & "/F" & CStr(.FNumberTurret)
            GetIllumCondtion.dblLux = NSIS_II.Level
            GetIllumCondtion.lFLTurret = .FNumberTurret
            GetIllumCondtion.lColor = .color
            GetIllumCondtion.lShutter = .Shutter
        End With
        
    ElseIf OptCond.IllumMaker = INTERACTION Then
                         
         '�u�����҂������Ȃ��v�ݒ�̂Ƃ��͊������܂�
         If Flg_Illum_GPIB_StatusRead_IA <> True Then
            '���̃E�F�C�g���Ȃ��Ɠ��삪�ρA20ms�Ƃ������Ԃ͕W�����C�u�����ɂ��킹���B
            TheHdw.WAIT 20 * mS
            
            OptStatus_Rev
            
            TheHdw.WAIT 20 * mS '�ꉞ
        End If
        
        IllumCondition = Space$(50)
        Call ibwrt(hOptPort, "B")
        Call ibrd(hOptPort, IllumCondition)
              
    End If
    
    GetIllumCondtion.strCondition = TrimIllumCondition(IllumCondition)
    
End Function

Private Function TrimIllumCondition(ByVal strCondition As String) As String
'���e:
'   �������������₷���`�ɒ���
'
'[strCondition]    IN   String�^:    ��������̉������̂���
'
'���ӎ���:
'   �悭�킩�����s�����Ă���̂ŁA����
'   EOS���[�h�̐ݒ肪�����H

    Dim varTmp As Variant
    
    strCondition = Replace(strCondition, Chr(10), "") 'LF
    strCondition = Replace(strCondition, Chr(13), "") 'CR
    strCondition = Trim(strCondition)
    
    'DARK��ԂȂ當���������Ă����Ԃ�
    If IsDarkCondition(strCondition) Then
        TrimIllumCondition = DARK_INDENTIFIER
        Exit Function
    End If
    
    varTmp = Split(strCondition, "/")
    
    TrimIllumCondition = ""
    Dim strExpresssion As String
    
    Dim i As Long
    For i = 0 To UBound(varTmp)
        If Left(varTmp(i), 1) = "X" Or Left(varTmp(i), 1) = "Y" Or Left(varTmp(i), 2) = "DK" Then
        Else
            If TrimIllumCondition = "" Then
                strExpresssion = ""
            Else
                strExpresssion = "/"
            End If
            
            TrimIllumCondition = TrimIllumCondition + strExpresssion + varTmp(i)
        End If
    Next i

End Function

Private Function GetOptGpibAddress() As Long
'���e:
'   ������GPIB�A�h���X��Ԃ�
'
'���ӎ���:
'   �W�����C�u�������̂��̃p�N��

    Dim wkshtObj As Object
    Dim basePoint As Variant
    Dim gbibAddrPoint As Variant
    Dim nodeRow As Long
    Dim nodeColumn As Long
    
    On Error GoTo EndOpt

    '======= WorkSheet Select ========
    Set wkshtObj = ThisWorkbook.Sheets(ILLUM_MAKER_LIST_WRKSHT_NAME)
         
    '======= WorkSheet ErrorProcess ========
    If IsEmpty(wkshtObj) Then
        MsgBox "Non IllumLst WorkSheet For Sw_Node Select"
        Exit Function
    End If
    
    '======= Base Point Find ========
    Set basePoint = Worksheets(ILLUM_MAKER_LIST_WRKSHT_NAME).Range("a1:K100").Find("Sw_node")
    If basePoint Is Nothing Then
        MsgBox "Non KeyWord @IllumLst WorkSheet For Find Sw_node"
        Exit Function
    End If
    
    '======= GPIB Address Point Find ========
    Set gbibAddrPoint = Worksheets(ILLUM_MAKER_LIST_WRKSHT_NAME).Range("a1:K100").Find("GPIB Address")
    If gbibAddrPoint Is Nothing Then
        MsgBox "Non KeyWord @IllumLst WorkSheet For Find GPIB Address"
        Exit Function
    End If
    
    '======= Search Node Start Addres Define ========
    nodeRow = basePoint.Row + 1
    nodeColumn = basePoint.Column

    '======= Search Node ========
    Do While Not wkshtObj.Cells(nodeRow, nodeColumn) = ""
        If wkshtObj.Cells(nodeRow, nodeColumn) = Sw_Node Then
            GetOptGpibAddress = Int(wkshtObj.Cells(nodeRow, gbibAddrPoint.Column))
            If GetOptGpibAddress = 0 Then
                '======== Default Setting ========
                If OptCond.IllumMaker = NIKON Then
                    GetOptGpibAddress = DEFAULT_ILLUM_GPIB_ADDR_NIKON
                ElseIf OptCond.IllumMaker = INTERACTION Then
                    GetOptGpibAddress = DEFAULT_ILLUM_GPIB_ADDR_IA
                End If
                '=================================
            End If
            
            Exit Function
        End If
        nodeRow = nodeRow + 1
    Loop

EndOpt:
    MsgBox "GPIB Address Error@Getting Illuminator GPIB Address", vbSystemModal, "GPIB Address Error"

End Function

Private Function IsDarkCondition(ByRef strCondition As String) As Boolean
'���e:
'   �擾�����������ʂ����ă_�[�N�Ȃ̂�
'
'���ӎ���:
'�E�C���^�[�A�N�V����
'   DK0�R�}���h��DARK�ɂ��ꂽ�Ƃ��̑΍�
'   N5A1200�����o������Œ蕶����ɂ��Ă��܂�
'
'�E�j�R��
'�@ �V���b�^��"2"�Ȃ�Œ蕶����ɂ���
'   �Ɠx�v����0.1���Ԃ��Ă��ċC��������

    Dim varTmp As Variant
        
    varTmp = Split(strCondition, "/")

    Dim i As Long
    Dim IsDark1 As Boolean, IsDark2 As Boolean
    IsDark1 = False
    IsDark2 = False
    
    If OptCond.IllumMaker = NIKON Then
        For i = 0 To UBound(varTmp)
            If Left(varTmp(i), 1) = "S" Then
                If Mid(varTmp(i), 2, Len(varTmp(i)) - 1) = "2" Then
                    IsDark1 = True
                    IsDark2 = True
                    Exit For
                End If
            End If
        Next i
    Else
    
        For i = 0 To UBound(varTmp)
            If Left(varTmp(i), 1) = "N" Then
                If Mid(varTmp(i), 2, Len(varTmp(i)) - 1) = "5" Then
                    IsDark1 = True
                End If
            ElseIf Left(varTmp(i), 1) = "A" Then
                If Mid(varTmp(i), 2, Len(varTmp(i)) - 1) = "1200" Then
                    IsDark2 = True
                End If
            End If
        Next i
    
    End If
    
    If IsDark1 And IsDark2 Then
        IsDarkCondition = True
    Else
        IsDarkCondition = False
    End If
    
End Function

Private Function OptStatus_Rev() As Long
'���e:
'    �����҂��������Ȃ�
'
'���ӎ���:
'�@�@�W�����C�u��������uFlg_Illum_GPIB_StatusRead_IA�v�̕�����ʂ�
'    ���̕���͂ЂƂ�̊K�w�ɂ��炷
'�@  �j�R����OptSet�ŕK�������҂������Ă���悤�Ȃ̂ŁA������s���̃C���^�[�A�N�V�����̂݁B
'

    Dim buf As String * 100
    Dim sendcom As String
    Dim status As String
    Dim Count As Integer
    
    Const STATUS_READY As String = "0"
    Const STATUS_BUSY As String = "1"
    Const STATUS_LOCAL As String = "2"
    Const STATUS_ERROR As String = "E"
    
    Dim iStatus As Long

    '========= NIKON =========
    If OptCond.IllumMaker = NIKON Then
        iStatus = NSIS_II.status
    
        While (iStatus <> 0)
            iStatus = NSIS_II.status
        Wend
    '=========== InterAction ===========
    ElseIf OptCond.IllumMaker = INTERACTION Then
    
        buf = Space$(100)
        sendcom = STATUS_READ_COMMAND + Chr(13) + Chr(10)
        status = STATUS_BUSY
        
        While (status <> STATUS_READY)
            '====== Status Read =====
            Call ibwrt(hOptPort, sendcom)
            Call ibrd(hOptPort, buf)
            status = Left(buf, 1)
            '========================
        
            If status = STATUS_BUSY Then
                Count = Count + 1
                If Count = 10000 Then
                    MsgBox "Illuminator is Busy Status.Timeout Error?", vbQuestion
                    Stop
                    Exit Function
                End If
            End If
    
            If status = STATUS_LOCAL Then
                MsgBox "Illuminator is LOCAL Status", vbExclamation
                Stop
            End If
        
            If status = STATUS_ERROR Then
                MsgBox "Illuminstor is Status Error", vbExclamation
                Break
                Exit Function
'                Stop
            End If
        Wend
    End If
    
End Function

'2012/11/12 175JobMakeDebug
Private Function InitOptCondition() As NSIS_II_CONDITION
 '���e:
'    �j�R���������\���̏�����
'
'���ӎ���:
'�@�@�W�����C�u�����̃p�N��
'

    With InitOptCondition
     .Axis = -1
     .Level = -1
     .Slit = -1
     .NDFilter = -1
     .WedgeFilter = -1
     .color = -1
     .Shutter = -1
     .LCShutter = -1
     .Diffusion = -1
     .Pattern = -1
     .FNumberIris = -1
     .Mirror = -1
     .LED = -1
     .Pupil = -1
     .Illuminant = -1
     .ColorTemperature = -1
     .FNumberTurret = -1
    End With
End Function

Private Function IsEqualTwoCondition(ByRef strRef As type_Illminator, ByRef strConditionBack As type_Illminator) As Boolean
'���e:
'   ���������������ǂ���
'
'���ӎ���:
'�E�C���^�[�A�N�V����
'   �P���ɕ�����̔�r�̂�
'
'�E�j�R��
'�@ DARK��ԂȂ當����̔�r����
'   ����ȊO�Ȃ�\���̂̑S�Ă̗v�f�ɂ��Ĕ�r����
'   �Ɠx�͂΂���̂�1.5%�܂ŋ��e����i1.5%�̗R���̓}�j���A���ɏ����Ă���s�[�N�덷�j

    IsEqualTwoCondition = False
    
    If OptCond.IllumMaker = NIKON Then
    
        If strConditionBack.strCondition = DARK_INDENTIFIER Then
            If strRef.strCondition = strConditionBack.strCondition Then
                IsEqualTwoCondition = True
                Exit Function
            End If
        
        Else
        
            If strRef.lShutter = strConditionBack.lShutter Then
                If strRef.lColor = strConditionBack.lColor Then
                    If strRef.lFLTurret = strConditionBack.lFLTurret Then
                        Dim dblValue As Double
                        dblValue = Abs((strConditionBack.dblLux - strRef.dblLux) / strConditionBack.dblLux)
                        If dblValue < 0.015 Then
                            IsEqualTwoCondition = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        
    ElseIf OptCond.IllumMaker = INTERACTION Then
        If strRef.strCondition = strConditionBack.strCondition Then
            IsEqualTwoCondition = True
            Exit Function
        End If
    End If

End Function