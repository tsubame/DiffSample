VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSnapIP750"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@@@IP750EX Support
'�T�v:
'   IP750�X�i�b�v�V���b�g�@�\���
'
'�ړI:
'   �T:�X�i�b�v�V���b�g�@�\����ւ��\��
'   �U:ISnapshot�̎����ɂ��@�\�̒�
'
'�쐬��:
'   SLSI����
'
'XlibTesterInfoMod�����b�s���O
'
'Code Checked
'Comment Checked
'

Option Explicit

'Tool�Ή���ɃR�����g�O���Ď��������ɂ���B�@2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1                '1CUL1G�{�[�h�̎g�p�@  0�F���g�p�A0�ȊO�F�g�p  <Tester��IP750EX�Ȃ�Default:1�ɂ��Ă���>
#Const HDVIS_USE = 0      'HVDIS���\�[�X�̎g�p�@ 0�F���g�p�A0�ȊO�F�g�p  <PALS��EeeAuto�����Ŏg�p>
#Const CUB_UB_USE = 0    'CUB UB�̐ݒ�          0�F���g�p�A0�ȊO�F�g�p

Implements ISnapshot
Private mLogFileName As String
Private mSerialNo As Long

#If UB_SNAPSHOT_USE <> 0 Then
Private m_UBSetObj As CUtyBitController
#End If

#If ICUL1G_USE <> 0 Then
Private m_Icul1gSnapshot As CIcul1gSnapshotProvider
#End If

Private mLogOutputPlace As OUTPUT_PLACE
Private mOutputStatus As Boolean

'�X�i�b�v�V���b�g�{�[�h
Enum SNAPSHOT_BOARD
    snapALL = 0
    snapAPMU
    snapPPMU
    snapDPS
    snapBPMU
    snapD_CH
    #If HDVIS_USE <> 0 Then
    snapHDVIS
    #End If
    
    #If ICUL1G_USE <> 0 Then
    snapICUL1G
    #End If
    
    'UB�ݒ��肪��������܂ł͂������񕕈󂷂�
    #If UB_SNAPSHOT_USE <> 0 Then
    snapAPMU_UB
    #If CUB_UB_USE <> 0 Then
    snapCUB_UB
    #End If
    #End If
End Enum

'�X�i�b�v���O�o�͐�̒�`
Enum OUTPUT_PLACE
    snapTXT_FILE = 0
    snapIGXL_DATALOG = 1
    snapTXT_AND_DATALOG = 2
End Enum

'#Pass-Ok
Public Sub Initialize()
'���e:
'   �X�i�b�v�V���b�g�@�\�̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:

    Call InitSnapIP750

End Sub

#If UB_SNAPSHOT_USE <> 0 Then
'#Pass-Ok UB�̐ݒ�s������肪��������܂ł͕���
Public Property Set UBObj(ByVal UBController As CUtyBitController)
'���e:
'    ���[�e�B���e�BBit�ݒ�Object
'
'�p�����[�^:
'    [UBController]   In  ���[�e�B���e�BBit�ݒ�Object
'
'�߂�l:
'
'���ӎ���:

    Set m_UBSetObj = UBController

End Property
#End If

'#Pass-Ok
Public Property Let OutputSaveStatus(ByVal outPutMsg As Boolean)
'���e:
'    �f�[�^���O�ւ̃X�i�b�v�V���b�g�擾���b�Z�[�W�̏o�͐؂�ւ�
'
'�p�����[�^:
'    [outputMsg]   In   TRUE:���b�Z�[�W�o��ON�AFALSE:���b�Z�[�W�o��OFF
'
'�߂�l:
'
'���ӎ���:
'

    mOutputStatus = outPutMsg

End Property

'#Pass-Ok
Public Function AsISnapshot() As ISnapshot
'���e:
'   ISnapshot OBJ�擾
'
'�p�����[�^:
'
'�߂�l:
'   ISnapshot �I�u�W�F�N�g
'
'���ӎ���:
'
    Set AsISnapshot = Me

End Function

'#Pass-Ok
Public Property Let SerialNumber(ByVal Number As Long)
'���e:
'    �X�i�b�v�V���b�g���O�̃��x���ɕ\������V���A���ԍ�
'
'�p�����[�^:
'    [Number]   In  �V���A���ԍ�
'
'�߂�l:
'
'���ӎ���:
'   GetSnapshot�����s���閈�ɒl�����ݒl���+1����܂��B
'
    mSerialNo = Number

End Property

'#Pass-Ok
Public Property Get SerialNumber() As Long
'���e:
'    �X�i�b�v�V���b�g���O�̃��x���ɕ\������V���A���ԍ�
'
'�p�����[�^:
'
'�߂�l:
'   ���݂̃V���A���ԍ��l
'
'���ӎ���:
'   GetSnapshot�����s���閈�ɒl�����݂̒l���+1����܂��B
'
    SerialNumber = mSerialNo

End Property

'#Pass-Ok
Public Property Let LogFileName(ByVal FullPathFileName As String)
'���e:
'    �X�i�b�v�V���b�g���O��TXT�t�@�C���ɏo�͂���Ƃ��̃t�@�C���ݒ�
'
'�p�����[�^:
'    [FullPathFileName]   In  ���O�t�@�C����
'
'�߂�l:
'
'���ӎ���:
'   �t�@�C������FullPath�Ŏw�肵�Ă��������B
'   (��) "D:\Jobs\Imx000\Log\SnapshotLog.txt"
'
    mLogFileName = FullPathFileName

End Property

'#Pass-Ok
Public Property Get LogFileName() As String
'���e:
'    �X�i�b�v�V���b�g���O��TXT�t�@�C���ɏo�͂���Ƃ��̃t�@�C���ݒ�
'
'�p�����[�^:
'
'�߂�l:
'   ���O�o�̓t�@�C����
'
'���ӎ���:
'
    LogFileName = mLogFileName

End Property

'#Pass-Ok
Public Property Let OutputPlace(ByVal Place As OUTPUT_PLACE)
'���e:
'    �X�i�b�v�V���b�g���O�̏o�͐�̐ݒ�
'
'�p�����[�^:
'    [Place]   In  ���O�̏o�͏ꏊ
'
'�߂�l:
'
'���ӎ���:
'   �o�͐�Ƀe�L�X�g�t�@�C�����w�肵�����ɁALogFile
'   �̐ݒ肪�s���Ă��Ȃ��ꍇ�̓G���[�ƂȂ�܂��
'   ���LogFile�̐ݒ���s���Ă��������B
'
    If (Place = snapTXT_AND_DATALOG Or Place = snapTXT_FILE) And (mLogFileName = "") Then
        Call OutputErrMsg("SnapshotLog file setting is nothing")
    Else
        mLogOutputPlace = Place
    End If

End Property

'#Pass-Ok
Public Sub GetSnapshot(Optional ByVal Comment As String = "", Optional ByVal ip750board As SNAPSHOT_BOARD = snapALL)
'���e:
'    �X�i�b�v�V���b�g���擾�����ʂ��o�͂���
'
'�p�����[�^:
'    [Comment]   In   �X�i�b�v�V���b�g���O�ɕ\������R�����g
'    [ip750board] In  �X�i�b�v�V���b�g���擾����{�[�h�̎w��
'
'�߂�l:
'
'���ӎ���:
'   ���s��A�X�i�b�v�V���b�g�V���A���ԍ��� +1 ����܂�
'
    Dim outputflg As Long
    Dim LogFileName As String
    Dim snapIdMsg As String
    Dim dataLogMsg As String
    
    Call makeSnapshotSetting(outputflg, LogFileName)
    snapIdMsg = makeSnapIdMsg(Comment)
    
    Select Case ip750board
        Case snapALL
            Call GetTesterInfo(snapIdMsg, outputflg, LogFileName)
            #If ICUL1G_USE <> 0 Then
            Call m_Icul1gSnapshot.CreatePpmuInfo(snapIdMsg, outputflg, LogFileName)
            #End If

            'UB�ݒ��肪��������܂ł͂������񕕈󂷂�
            #If UB_SNAPSHOT_USE <> 0 Then
            If Not m_UBSetObj.AsAPMU Is Nothing Then
                Call GetApmuUbSnapshot(snapIdMsg, outputflg, LogFileName)
            End If
            #If CUB_UB_USE <> 0 Then
            If Not m_UBSetObj.AsCUB Is Nothing Then
                Call GetCubUbSnapshot(snapIdMsg, outputflg, LogFileName)
            End If
            #End If
            #End If
        
        Case snapAPMU
            Call CreateApmuInfo(snapIdMsg, outputflg, LogFileName)
        
        Case snapPPMU
            Call CreatePpmuInfo(snapIdMsg, outputflg, LogFileName)
        
        Case snapD_CH
            Call CreatePeInfo(snapIdMsg, outputflg, LogFileName)
        
        Case snapBPMU
            Call CreateBpmuInfo(snapIdMsg, outputflg, LogFileName)
        
        Case snapDPS
            Call CreateDpsInfo(snapIdMsg, outputflg, LogFileName)
        
        #If HDVIS_USE <> 0 Then
        Case snapHDVIS
            Call CreateHdvisInfo(snapIdMsg, outputflg, LogFileName)
        #End If
        
                
        #If ICUL1G_USE <> 0 Then
        Case snapICUL1G
            Call m_Icul1gSnapshot.CreatePpmuInfo(snapIdMsg, outputflg, LogFileName)
        #End If
                
        'UB�ݒ��肪��������܂ł͂������񕕈󂷂�
        #If UB_SNAPSHOT_USE <> 0 Then
        Case snapAPMU_UB
            If Not m_UBSetObj.AsAPMU Is Nothing Then
                Call GetApmuUbSnapshot(snapIdMsg, outputflg, LogFileName)
            Else
                Call OutputErrMsg("APMU UB object is Nothing! Please set UBObj Property")
            End If

        #If CUB_UB_USE <> 0 Then
        Case snapCUB_UB
            If Not m_UBSetObj.AsCUB Is Nothing Then
                Call GetCubUbSnapshot(snapIdMsg, outputflg, LogFileName)
            Else
                Call OutputErrMsg("CUB UB object is Nothing! Please set UBObj Property")
            End If
        #End If
        #End If
        
    
    End Select
    
    
    '�X�i�b�v�V���b�g�̎擾���s�������Ƃ��f�[�^���O�ɕ\������Ƃ��̏���
    If mOutputStatus = True Then
        If Comment <> "" Then
            dataLogMsg = "#" & TheExec.DataManager.InstanceName & "." & Comment & "." & Me.SerialNumber
        Else
            dataLogMsg = "#" & TheExec.DataManager.InstanceName & "." & Me.SerialNumber
        End If
        TheExec.Datalog.WriteComment "TESTER_SS_SAVED = " & dataLogMsg
    End If
        
    '�X�i�b�v�V���b�g��ID�p�̃V���A���ԍ� increment
    mSerialNo = mSerialNo + 1

End Sub

'UB�ݒ��肪��������܂ł͂������񕕈󂷂�
#If UB_SNAPSHOT_USE <> 0 Then
'#Pass-Ok
Private Sub GetApmuUbSnapshot(ByVal CommentMsg As String, Optional OutputDatalog As Long = 1, _
    Optional outputLogName As String = "")
'APMU UB�̃X�i�b�v�V���b�g�擾�p

    Dim snapMsg As String

    snapMsg = CommentMsg & ",[APMU_UB]," & "Board=" & m_UBSetObj.AsAPMU.BoardNumber & _
                            ",Hex=" & m_UBSetObj.AsAPMU.UBSnapshotInHex

    If OutputDatalog = 1 Then
        Call TheExec.Datalog.WriteComment(snapMsg)
        Call TheExec.Datalog.WriteComment("")
    End If
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, snapMsg)
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub

'#Pass-Ok
#If CUB_UB_USE <> 0 Then
Private Sub GetCubUbSnapshot(ByVal CommentMsg As String, Optional OutputDatalog As Long = 1, _
    Optional outputLogName As String = "")
'CUB UB�̃X�i�b�v�V���b�g�擾�p

    Dim snapMsg As String

    snapMsg = CommentMsg & ",[CUB_UB]," & "Hex=" & m_UBSetObj.AsCUB.UBSnapshotInHex

    If OutputDatalog = 1 Then
        Call TheExec.Datalog.WriteComment(snapMsg)
        Call TheExec.Datalog.WriteComment("")
    End If
    If outputLogName <> "" Then
        Call mf_OutPutLog(outputLogName, snapMsg)
        Call mf_OutPutLog(outputLogName, "")
    End If

End Sub
#End If

#End If


'#Pass-Ok
Private Function makeSnapIdMsg(ByVal Comment As String) As String
'�X�i�b�v�V���b�g�̃R�����g���b�Z�[�W�쐬
        
    '�X�i�b�v�V���b�gID�p�̃��x���쐬
    If Comment <> "" Then
        makeSnapIdMsg = "#" & TheExec.DataManager.InstanceName & "." & Comment & "." & mSerialNo
    Else
        makeSnapIdMsg = "#" & TheExec.DataManager.InstanceName & "." & mSerialNo
    End If

End Function

'#Pass-Ok
Private Sub makeSnapshotSetting(ByRef outputflg As Long, ByRef LogFileName As String)
'�X�i�b�v�V���b�g���s�O�̃v���p�e�B�̐ݒ�󋵂̊m�F

    Select Case mLogOutputPlace
        Case snapTXT_FILE
            outputflg = 0
            LogFileName = mLogFileName
                    
        Case snapIGXL_DATALOG
            outputflg = 1
            LogFileName = ""
        
        Case snapTXT_AND_DATALOG
            outputflg = 1
            LogFileName = mLogFileName
    End Select

End Sub

'#Pass-Ok
Private Sub mf_OutPutLog(ByVal LogFileName As String, outPutMessage As String)
'�X�i�b�v�V���b�g�p�̃��O���t�@�C���ɏo�͂���B
    Dim fp As Integer
    On Error GoTo OUT_PUT_LOG_ERR

    fp = FreeFile
    Open LogFileName For Append As fp
    Print #fp, outPutMessage
    Close fp

    Exit Sub

OUT_PUT_LOG_ERR:
    Call OutputErrMsg(LogFileName & " MsgOutPut Error" & " @mf_OutPutLog")
    Stop

End Sub

'#Pass-Ok
Private Sub InitSnapIP750()
'������
    mLogFileName = ""
    mLogOutputPlace = snapIGXL_DATALOG
    mSerialNo = 1
    
    #If UB_SNAPSHOT_USE <> 0 Then
    Set m_UBSetObj = Nothing
    #End If
    
    mOutputStatus = False

    #If ICUL1G_USE <> 0 Then
    Set m_Icul1gSnapshot = New CIcul1gSnapshotProvider
    #End If

End Sub

'#Pass-Ok
Private Sub Class_Initialize()
    
    Call InitSnapIP750

End Sub




'######################## �ȉ� ISnapshot�̎��� #############################

'#Pass-Ok
Private Property Get ISnapshot_FunctionName() As String
    
    ISnapshot_FunctionName = "IP750Snapshot"

End Property

'#Pass-Ok
Private Sub ISnapshot_SaveSnapshot(Optional ByVal CommentMsg As String = "")
    
    Call Me.GetSnapshot(CommentMsg)

End Sub
