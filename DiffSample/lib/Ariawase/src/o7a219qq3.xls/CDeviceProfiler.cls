VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDeviceProfiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   DeviceConfigurations�V�[�g�ɐݒ肵�������擾���ێ�����N���X
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfigurations�V�[�g�̏����擾���ێ�����@�\����������
'       2010/12/07  �s�v�R�[�h���폜����
'
'�쐬��:
'   0145184346
'

Option Explicit

Private Const ERR_NUMBER As Long = 9999                ' Error�ԍ���ێ�����
Private Const CLASS_NAME As String = "CDeviceProfiler" ' Class���̂�ێ�����

Private Const PARAMETER_START_RANGE As String = "B5"   ' �f�[�^�ݒ�J�n�Z����ێ�����
Private Const PARAMETER_RANGE  As String = "B5:K5"     ' �f�[�^�ݒ�\�Z����ێ�����
Private Const PARAMETER_MODE_RANGE  As String = "H5"   ' Mode�p�����[�^�ݒ�Z����ێ�����
Private Const PARAMETER_START_ROW As Long = 5          ' �f�[�^�ݒ�J�n�Z����ێ�����
Private Const PARAMETER_MODE_COLUMN As Long = 8        ' Mode�p�����[�^�ݒ�Z����ێ�����
Private Const PARAMETER_MODE_TITLE As String = "Mode"  ' ���͋K���̃p�����[�^���̂�ێ�����
Private Const MODE_STATIC_WORD As String = "Static"    ' ���͋K���̃p�����[�^�l��ێ�����
Private Const MODE_DYNAMIC_WORD As String = "Dynamic"  ' ���͋K���̃p�����[�^�l��ێ�����

Private mConfigurations As Collection ' �S�Ă�Config�f�[�^��ێ�����

Public Sub Initialize()
'���e:
'   DeviceConfigurations�f�[�^���擾���ĕێ�����
'
'�p�����[�^:
'
'�߂�l:
'
'���ӎ���:
'

    '#####  SheetReader�𗘗p���āADeviceConfigurations�V�[�g��ǂݍ���  #####
    Dim clsWrkShtRdr As CWorkSheetReader
    Set clsWrkShtRdr = GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeDeviceConfigurations)
    
    Dim IFileStream As IFileStream
    Dim IParamReader As IParameterReader

    With clsWrkShtRdr
        Set IFileStream = .AsIFileStream
        Set IParamReader = .AsIParameterReader
    End With

    Set clsWrkShtRdr = Nothing


    '#####  DeviceConfigurations����ǂݏo�����i�[����  #####
    Dim clsDevCfg As CDeviceConfiguration
    Set mConfigurations = New Collection

    Do While Not IFileStream.IsEOR
        On Error GoTo ErrHandler
        Set clsDevCfg = New CDeviceConfiguration
        Call clsDevCfg.Initialize(IParamReader)
        On Error GoTo ErrHandlerInstanceName
        mConfigurations.Add clsDevCfg, clsDevCfg.InstanceName
        On Error GoTo 0
        IFileStream.MoveNext
    Loop

    Set clsDevCfg = Nothing
    Set IFileStream = Nothing
    Set IParamReader = Nothing

    '#####  �I��  #####
    Exit Sub


ErrHandler:
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Sub
ErrHandlerInstanceName:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Already Set Instance Name.[" & clsDevCfg.InstanceName & "]")
    Exit Sub

End Sub

Public Property Get Configuration(ByRef InstanceName As String) As CDeviceConfiguration
'���e:
'   �w�肵��Instance���̂�DeviceConfiguration�f�[�^��߂�
'
'�p�����[�^:
'   [InstanceName]  In  Instance���̂�ێ�����
'
'�߂�l:
'   DeviceConfiguration�f�[�^
'
'���ӎ���:
'

    '#####  �w�肵��Instance���̂�Configuration�����擾����  #####
    On Error GoTo ErrHandler
    Set Configuration = mConfigurations.Item(InstanceName)
    Exit Property

ErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Configuration", "Illegal TestInstance Name.")
    Exit Property
    
End Property

Public Property Get ConfigInfo(ByRef InstanceName As String) As CDeviceConfigInfo
'���e:
'   �w�肵��Instance���̂�DeviceConfigInfo�f�[�^��߂�
'
'�p�����[�^:
'   [InstanceName]  In  Instance���̂�ێ�����
'
'�߂�l:
'   DeviceConfigInfo�f�[�^
'
'���ӎ���:
'

    '#####  �w�肵��Instance���̂�ConfigInfo���݂̂��擾����  #####
    Dim clsDevCfg As CDeviceConfiguration
    On Error GoTo ErrHandler
    Set clsDevCfg = mConfigurations.Item(InstanceName)
    On Error GoTo 0
    Set ConfigInfo = clsDevCfg.ConfigInfo
    Set clsDevCfg = Nothing
    Exit Property

ErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".ConfigInfo", "Illegal TestInstance Name.")
    Exit Property
    
End Property

Public Sub SelectionCellForm(ByRef rngTarget As Range)
'���e:
'   Cell�ɓ��͋K����t������
'
'�p�����[�^:
'   [rngTarget]  In  �I�����ꂽ�Z�����
'
'�߂�l:
'
'���ӎ���:
'

    '#####  ���͋K�������X�g�����Đݒ肷��  #####
    Const SET_VALUE_LIST As String = MODE_STATIC_WORD & "," & MODE_DYNAMIC_WORD
    With rngTarget
        If .Column = PARAMETER_MODE_COLUMN And .Row >= PARAMETER_START_ROW And .Count = 1 Then
            Call addValidationCell(rngTarget, PARAMETER_MODE_TITLE, SET_VALUE_LIST)
        End If
    End With

End Sub

Private Sub addValidationCell(ByRef pTargetRange As Range, ByVal pDataName As String, _
                            ByVal pDataItemList As String, Optional ByVal pShowInputMessage As Boolean = False)
    
    '#####  ��x���݂̓��͋K�����폜  #####
    Call pTargetRange.Validation.Delete
    
    '#####  ���͋K����V�K�ǉ�  #####
    Call pTargetRange.Validation.Add(Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                                    operator:=xlEqual, Formula1:=pDataItemList)
    
    '#####  �\�����b�Z�[�W���̐ݒ�  #####
    With pTargetRange.Validation
        .InputTitle = pDataName                     ' Balloon�̃^�C�g���l
        .InputMessage = "Please select"             ' Balloon�̃��b�Z�[�W�l
        .ShowInput = pShowInputMessage              ' Balloon�̕\���ݒ�
        .InCellDropdown = True                      ' Dropdown�̕\���ݒ�
        .ErrorTitle = "Unknown " & pDataName & " !" ' Error Msgbox �^�C�g��
        .ErrorMessage = "Input Value is unknown !"  ' Error Msgbox ���b�Z�[�W
        .IgnoreBlank = False                        ' �u�����N���͋��e�ݒ�
        .IMEMode = xlIMEModeNoControl               ' IME���[�h�ݒ�
        .ShowError = True
    End With

End Sub

Public Sub ChangeCellForm(ByRef strWorkSheetName As String)
'���e:
'   Cell�ɖԊ|�����������{����
'
'�p�����[�^:
'   [strWorkSheetName]  In  �g�p����V�[�g��
'
'�߂�l:
'
'���ӎ���:
'

    '#####  Sheet���擾  #####
    Dim shtDevCfg As Worksheet
    On Error GoTo ErrHandler
    Set shtDevCfg = Worksheets(strWorkSheetName)
    On Error GoTo 0

    '#####  �Ԋ|�����������{  #####
    Application.ScreenUpdating = False
    Call ClearDataSheetFormat(shtDevCfg)
    Call setDataSheetFormat(shtDevCfg)
    Application.ScreenUpdating = True
           
    '#####  �I��  #####
    Set shtDevCfg = Nothing
    Exit Sub
        
ErrHandler:
    Call errorRaise("DeviceConfigurations WorkSheet is Nothing")
    Exit Sub

End Sub

Private Sub setDataSheetFormat(ByRef shtWork As Worksheet)
    Dim rngCondition As Range ' �f�[�^�J�n�Z��
    Dim rngData As Range      ' Mode�p�����[�^�ݒ�Z��
    Dim rngTarget As Range    ' �ύX��Z��
    Set rngCondition = shtWork.Range(PARAMETER_START_RANGE)
    Set rngData = shtWork.Range(PARAMETER_MODE_RANGE)
    While IsEmpty(rngCondition) = False
        If rngData.Value = MODE_STATIC_WORD Then
            Set rngTarget = rngData.offset(0, 2)
        ElseIf rngData.Value = MODE_DYNAMIC_WORD Then
            Set rngTarget = rngData.offset(0, 1)
        Else
            Set rngTarget = Nothing
        End If
        If Not (rngTarget Is Nothing) Then
            With rngTarget.Interior
                .Pattern = xlGray8
                .ColorIndex = 15
            End With
        End If
        Set rngData = rngData.offset(1, 0)
        Set rngCondition = rngCondition.offset(1, 0)
    Wend
    Set rngCondition = Nothing
    Set rngData = Nothing
    Set rngTarget = Nothing
End Sub

Private Sub ClearDataSheetFormat(ByRef shtWork As Worksheet)
    Dim rngCondition As Range ' �f�[�^�J�n�Z��
    Dim rngData As Range      ' �f�[�^�ݒ�\�Z��
    Set rngCondition = shtWork.Range(PARAMETER_START_RANGE)
    Set rngData = shtWork.Range(PARAMETER_RANGE)
    While IsEmpty(rngCondition) = False
        With rngData.Interior
            .Pattern = xlSolid
            .ColorIndex = xlNone
        End With
        Set rngData = rngData.offset(1, 0)
        Set rngCondition = rngCondition.offset(1, 0)
    Wend
    Set rngCondition = Nothing
    Set rngData = Nothing
End Sub

Private Sub errorRaise(ByVal pErrMsg As String)
    If Not (TheError Is Nothing) Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME, pErrMsg)
    Else
        Call Err.Raise(ERR_NUMBER, CLASS_NAME, pErrMsg)
    End If
End Sub

Private Sub Class_Initialize()
    Set mConfigurations = Nothing
End Sub

Private Sub Class_Terminate()
    Set mConfigurations = Nothing
End Sub