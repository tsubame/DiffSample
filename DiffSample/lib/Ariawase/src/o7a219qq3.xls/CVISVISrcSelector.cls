VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVISVISrcSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@@@IP750EX Support
'�T�v:
'   ChannelType���A�{�[�h�w��ɂ��Ή�����
'   IP750�̓d���ݒ�OBJ��I��������
'
'�ړI:
'   ChannelType�������ɑΉ�����DC���\�[�X��I��
'   ���\�[�X�w��ɂ��I��
'   �g�p����DC���\�[�X�̃��X�g��ێ�
'
'�쐬��:
'   SLSI����
'
'���ӎ���:
'(�Ή����\�[�X)
'   APMU�ADPS�APPMU�ABPMU�AHDVIS�AHSD200-PPMU�AICUL1G�|PPMU
'
'   HDVIS���p���ɂ́A�����t�R���p�C��������
'   HDVIS_USE=1�̋L�q���K�v�ł��B
'
'   HSD200���p���ɂ́A�����t�R���p�C��������
'   HSD200_USE=1�̋L�q���K�v�ł��B
'
'   ICUL1G���p���ɂ́A�����t�R���p�C��������
'   ICUL1G_USE=1�̋L�q���K�v�ł��B
'

Option Explicit

'Tool�Ή���ɃR�����g�O���Ď��������ɂ���B�@2013/03/07 H.Arikawa
#Const HDVIS_USE = 0      'HVDIS���\�[�X�̎g�p�@ 0�F���g�p�A0�ȊO�F�g�p  <PALS��EeeAuto�����Ŏg�p>
#Const ICUL1G_USE = 1               '1CUL1G�{�[�h�̎g�p�@  0�F���g�p�A0�ȊO�F�g�p  <Tester��IP750EX�Ȃ�Default:1�ɂ��Ă���>
#Const HSD200_USE = 1               'HSD200�{�[�h�̎g�p�@  0�F���g�p�A0�ȊO�F�g�p  <Tester��IP750EX�Ȃ�Default:1�ɂ��Ă���>

'�Z���N�g�Ή�����{�[�h��`
'���V�����d���{�[�h�ւ̑Ή����ɂ̓R�R�ɏ��ǉ��̕K�v������B
Public Enum DC_BOARD_NAME
    CHAN_TYPE_SELECT = -1 'ChannelType���ɂ��Board�I���w��p
    dcAPMU = 0
    dcDPS = 1
    dcPPMU = 2
    dcBPMU = 3
    
    #If HDVIS_USE <> 0 Then
    dcHDVIS = 4
    #End If
    
    #If ICUL1G_USE <> 0 Then
    dcICUL1G = 5
    #End If
End Enum

'DC�{�[�hObject�i�[�p�R���N�V����
Private m_BoardList As Collection

'#Pass-Ok
Public Property Get BoardCount() As Long
'���e:
'   �{�[�h���X�g�ɓo�^����Ă���{�[�h�̐��擾
'
'�p�����[�^:
'
'�߂�l:
'   �{�[�h���X�g�ɓo�^����Ă���{�[�h��
'
'���ӎ���:
'
    BoardCount = m_BoardList.Count

End Property

'#Pass-Ok
Private Sub InitBoardList()
'���e:
'   BoardList�̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    Set m_BoardList = Nothing 'Pass
    Set m_BoardList = New Collection 'Pass
    
End Sub

'#Pass-Ok
Private Sub AddBoard(dcboard As IVISVISource)
'���e:
'   BoardList�R���N�V�����ɁA�d���ݒ�pObject��o�^
'
'�p�����[�^:
'   [dcboard]   In  �ǉ����s���d���ݒ�pObject
'
'�߂�l:
'
'���ӎ���:
'   �o�^�ł���̂́AIVISVISource���������Ă���OBJ�ł�
'
    
    Dim boardchk As IVISVISource
            
    If m_BoardList.Count <> 0 Then
        '�o�^���w�����ꂽ�{�[�h�����łɃ{�[�hList�ɑ��݂��邩���m�F
        Set boardchk = Nothing
        On Error Resume Next
        Set boardchk = m_BoardList.Item(dcboard.BoardName)
        On Error GoTo 0
    
        If boardchk Is Nothing Then
            '�V�K�̃{�[�h���m�F���ꂽ�̂œo�^
            Call m_BoardList.Add(dcboard, dcboard.BoardName)
        Else
            '���łɑ��݂���{�[�h�Ȃ̂ŃG���[�Ƃ���
            Call OutputErrMsg(dcboard.BoardName & " board is exist (at VI Source Selector)")
        End If
    Else
        '�o�^����Ă���{�[�h���܂����������ꍇ�́A���̂܂ܓo�^�i�d���m�F�Ȃ��j
        Call m_BoardList.Add(dcboard, dcboard.BoardName)
    End If
    
    Set boardchk = Nothing

End Sub

'#Pass-Ok
Public Function Pins(ByVal PinList As String, Optional ByVal DCBoardName As DC_BOARD_NAME = CHAN_TYPE_SELECT) As IVISVISource
'���e:
'   PinMap�ɒ�`�σs����ChannelType���A�s�����g�p���Ă���d���̎w����
'   �ɂ��Ή�����d���ݒ�pObject��Ԃ��܂�
'
'�p�����[�^:
'   [PinList]       In   PinMap�ɒ�`�ς݂̃s����
'�@ [DCBoardName]�@ In   �w�肵���s�����g�p����d���ݒ�{�[�h�� �I�v�V����(Default CHAN_TYPE_SELECT)
'
'�߂�l:
'   �d���ݒ�pIVISVISource�I�u�W�F�N�g
'
'���ӎ���:
'   �s�����𕡐��w�肷��ꍇ�́A���̃O���[�v�ɑΉ�����
'   �d���ݒ�pObject������ł���K�v������܂��B
'
'   (��)
'    �� object.Pins("P_APMU1, P_APMU2, P_APMU3")
'    �~ object.Pins("P_APMU, P_DPS, P_IO")
'
'   DCBoardName���w�肵�Ȃ��ꍇ�́A�s�����ɑΉ�����ChannelType�ɂ��
'   �d���ݒ�pObject���I������܂��B
'
'   ��ChannelType��chIO�Ƃ��Ē�`����Ă���s�����w�肵
'     DCBoardName���w�肵�Ȃ��ꍇ��PPMU���I������܂��
'
'    (��)
'     object.Pins("P_IO")        --> PPMU��I��
'     object.Pins("P_IO",dcBPMU) --> BPMU��I��
'
    
    Dim chanType As HdwDrivers.chtype
    Dim selectBoard As IVISVISource
    
    Set Pins = Nothing
    
    'Pin�����ChannelType�擾
    chanType = GetChanType(PinList)
      
    '### �{�[�h�w�肠��̏��� ###
    If DCBoardName <> CHAN_TYPE_SELECT Then
        
        'DCBoardName���L�[�ɓd���ݒ�pOBJ�����o��
        Set selectBoard = Nothing
        
        On Error Resume Next
        Set selectBoard = m_BoardList.Item(DcBoardNameToBoardKey(DCBoardName))
        On Error GoTo 0
    
        If selectBoard Is Nothing Then
            '�Ή�����d���ݒ�pOBJ��BoardList�ɂȂ������ꍇ�̃G���[
            Call OutputErrMsg("Select board " & DcBoardNameToBoardKey(DCBoardName) & " is nothing (at VI Source Selector)")
        Else
            '�Ή�����d���ݒ�pOBJ�Ƀs�����X�g��ݒ肷��
            selectBoard.PinList = PinList
            Set Pins = selectBoard
        End If
    
        Set selectBoard = Nothing
        
    '### �{�[�h�w��Ȃ��̏��� ###
    Else
        'ICUL1G�Ή�����
        #If ICUL1G_USE <> 0 Then
            '�Ώۃs�����AICUL1G��PPMU�s�����ǂ����𒲂ׂ�
            If TheHdw.ICUL1G.IsICUL1GPin(PinList, icul1gPinPPMU) = True Then
                'ICUL1G�{�[�h�I������
                Set selectBoard = Nothing
                On Error Resume Next
                Set selectBoard = m_BoardList.Item("dcICUL1G")
                On Error GoTo 0
                If selectBoard Is Nothing Then
                    '�Ή�����d���ݒ�pOBJ��BoardList�ɂȂ������ꍇ�̃G���[
                    Call OutputErrMsg(PinList & " use board is nothing (at VI Source Selector)")
                Else
                    '�Ή�����d���ݒ�pOBJ�Ƀs�����X�g��ݒ肷��
                    selectBoard.PinList = PinList
                    Set Pins = selectBoard
                End If
                Exit Function
            End If
        #End If
                        
        'ChannelType���L�[�ɓd���ݒ�pOBJ�����o��
        Set selectBoard = Nothing
        On Error Resume Next
        Set selectBoard = m_BoardList.Item(ChTypeToBoardKey(chanType))
        On Error GoTo 0
    
        If selectBoard Is Nothing Then
            '�Ή�����d���ݒ�pOBJ��BoardList�ɂȂ������ꍇ�̃G���[
            Call OutputErrMsg(PinList & " use board is nothing (at VI Source Selector)")
        Else
            '�Ή�����d���ݒ�pOBJ�Ƀs�����X�g��ݒ肷��
            selectBoard.PinList = PinList
            Set Pins = selectBoard
        End If
    
        Set selectBoard = Nothing
    
    End If

End Function


'#Pass-Ok
Public Sub Initialize()
'���e:
'   �{�[�h�Z���N�^�̏������i�g�p�J�n�����j
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    
    Call InitBoardList
    Call InitBoardSetting

End Sub

'#Pass-Ok
Private Sub InitBoardSetting()
'���e:
'   �{�[�h�����̏�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'
    '�{�[�h�N���X�̐����ƃ��X�g�쐬
    Call AddBoard(New CVISApmu)
    Call AddBoard(New CVISDps)
    
    'PPMU�{�[�h�̏����̓R���p�C���I�v�V�����̎w��l�ōs��
    #If HSD200_USE <> 0 Then
        Call AddBoard(New CHsd200Ppmu)    'HSD200-PPMU
    #Else
        Call AddBoard(New CVISPpmu)       'HSD100-PPMU
    #End If
    
    Call AddBoard(New CVISBpmu)
    
    'HDVIS�{�[�h�̏����̓R���p�C���I�v�V�����Ŏw�肳�ꂽ���̂ݍs���B
    '�R���p�C���G���[�΍�
    #If HDVIS_USE <> 0 Then
        Call AddBoard(New CVISHdvis)
    #End If

    'ICUL1G�{�[�h�̏����̓R���p�C���I�v�V�����Ŏw�肳�ꂽ���̂ݍs���B
    '�R���p�C���G���[�΍�
    'ICUL1G
    #If ICUL1G_USE <> 0 Then
        Call AddBoard(New CIcul1gPpmu)
    #End If

End Sub

'#Pass-Ok
Private Sub Class_Initialize()
    
    Call Me.Initialize

End Sub

'#Pass-Ok
Private Function ChTypeToBoardKey(ByVal ChannelType As HdwDrivers.chtype) As String
'ChannelType�̒l���{�[�h�I��p�̃L�[���ɕϊ�����
'���V�����d���{�[�h�ւ̑Ή����ɂ̓R�R�ɏ��ǉ��̕K�v������B

    Select Case ChannelType
        Case chAPMU
            ChTypeToBoardKey = "dcAPMU"
        Case chDPS
            ChTypeToBoardKey = "dcDPS"
        Case chIO
            ChTypeToBoardKey = "dcPPMU"    '�f�W�^��Pin(I/O Pin)�̓{�[�h�w�肪�����Ƃ��ɂ�PPMU�Ƃ��Ď�舵��
        #If HDVIS_USE <> 0 Then
        Case chHDVIS
            ChTypeToBoardKey = "dcHDVIS"
        #End If
    End Select

End Function

'#Pass-Ok
Private Function DcBoardNameToBoardKey(ByVal BoardName As DC_BOARD_NAME) As String
'DC_BOARD_NAME�̒l���{�[�h�I��p�̃L�[���ɕϊ�����
'���V�����d���{�[�h�ւ̑Ή����ɂ̓R�R�ɏ��ǉ��̕K�v������B

    Select Case BoardName
        Case CHAN_TYPE_SELECT
            DcBoardNameToBoardKey = "CHAN_TYPE_SELECT" '���g�p
        Case dcAPMU
            DcBoardNameToBoardKey = "dcAPMU"
        Case dcDPS
            DcBoardNameToBoardKey = "dcDPS"
        Case dcPPMU
            DcBoardNameToBoardKey = "dcPPMU"
        Case dcBPMU
            DcBoardNameToBoardKey = "dcBPMU"
        
        #If HDVIS_USE <> 0 Then
        Case dcHDVIS
            DcBoardNameToBoardKey = "dcHDVIS"
        #End If
    
        #If ICUL1G_USE <> 0 Then
        Case dcICUL1G
            DcBoardNameToBoardKey = "dcICUL1G"
        #End If
    
    End Select

End Function