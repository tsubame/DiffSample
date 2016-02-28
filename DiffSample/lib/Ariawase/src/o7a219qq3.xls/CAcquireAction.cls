VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcquireAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   AcquireFrameWork（FWImageAcquire）の情報を保持＆実行するクラス
'
'   Revision History:
'       Data        Description
'       2010/04/28  AcquireFrameWork（FWImageAcquire）を実行する機能を実装した
'       2010/06/11  データ構造見直しのため、プログラムコードを変更した
'       2010/06/22  実行前後の情報を追加した（m_Statusを変更）
'                   実行情報リセット機能を実装した
'                   Macro実行エラー処理を追加した
'
'作成者:
'   tomoyoshi.takase
'   0145184346
'

Option Explicit

Implements IAcquireAction

Private Const ERR_NUMBER = 9999             ' エラー番号を保持する
Private Const ERR_NUMBER_NOTFOUND_MACRO = 424
Private Const CLASS_NAME = "CAcquireAction" ' Class名称を保持する

Private m_FunctionInfo As CAcquireFunctionInfo ' ImageACQTBLシート1行のデータを保持する
'Private m_ActionLogger As CActionLogger        ' 実行LogデータをダンプするActionLoggerを保持する
Private m_InstanceName As String               ' TestInstance名称を保持する
Private m_ActionName As String                 ' AcquireFrameWorkが実行するUserMacro名称を保持する
Private m_Status As Long                       ' 実行結果を保持する（-1：実行前　0：正常終了　1：エラー発生）

Public Property Get AsIAcquireAction() As IAcquireAction
    Set AsIAcquireAction = Me
End Property

Private Function IAcquireAction_Execute() As Integer
'内容:
'   ImageAcquireのUserMacroを実行する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'
'   Arg(0):インスタンス名
'   Arg(1):FW名
'   Arg(2):転送先プレーングループ
'   Arg(3):転送先プレーンPMD名
'   Arg(4):BitDepth
'   以降ユーザ領域

    On Error GoTo VBAErrHandler

    Dim capPlane As CImgPlane
    Set capPlane = GetFreePlaneForAcq(m_FunctionInfo.Arg(2), ConvertStrToDepth(m_FunctionInfo.Arg(3)), , m_FunctionInfo.Arg(0) & " Capture Plane")

    Dim devInfo As CDeviceConfigInfo
    Set devInfo = TheDeviceProfiler.ConfigInfo(m_FunctionInfo.Arg(0))
    
    Dim capParam As CParamPlane
    Set capParam = New CParamPlane
    Set capParam.plane = capPlane
    Set capParam.DeviceConfigInfo = devInfo
    
    m_Status = Application.Run(m_ActionName, m_FunctionInfo, capParam)

    'GetFreePlaneが成功。
    If Not capPlane Is Nothing Then
        Call TheParameterBank.Add(capParam, m_FunctionInfo.Arg(0))
    End If

    On Error GoTo ErrHandler

    'ユーザーマクロからの返り値をTL_SUCCESSに変更するのはCAcquirenstance
    IAcquireAction_Execute = m_Status
    
    If m_Status = TL_ERROR Then
        Call TheError.Raise(ERR_NUMBER, m_ActionName, "User Macro returned Error!")
    ElseIf m_Status <> TL_ERROR Then
        m_Status = TL_SUCCESS
    End If
    
    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
VBAErrHandler:
    m_Status = TL_ERROR
    If Err.Number = ERR_NUMBER_NOTFOUND_MACRO Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Execute", _
                        "Can Not Find Acquire Macro." & "(" & m_ActionName & ")")
    Else
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Execute", _
                        "[" & Err.Number & ":" & Err.Source & "]" & Err.Description & "(" & m_ActionName & ")")
    End If
    Exit Function
ErrHandler:
    m_Status = TL_ERROR
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".Execute", "Image Acquire Error : " & m_InstanceName)
    Exit Function

End Function

Private Function ConvertStrToDepth(ByVal pStr As String) As IdpBitDepth
    If pStr Like "*S16" Then
        ConvertStrToDepth = idpDepthS16
    ElseIf pStr Like "*S32" Then
        ConvertStrToDepth = idpDepthS32
    ElseIf pStr Like "*F32" Then
        ConvertStrToDepth = idpDepthF32
    Else
        Err.Raise 9999, TypeName(Me), "Illegal IdpBitDepth: " & pStr
    End If
End Function

Private Function GetFreePlaneForAcq(ByVal pType As String, ByVal pBitDepth As IdpBitDepth, _
        Optional ByVal pClear As Boolean = False, Optional pComment As String) As CImgPlane
    
    Dim ret As CImgPlane
    On Error GoTo ErrExit
    Set ret = TheIDP.PlaneManager(pType).GetFreePlane(pBitDepth, pClear)
    ret.Comment = pComment
    
    Set GetFreePlaneForAcq = ret
    Exit Function
    
ErrExit:
    Dim Err As CErrInfo
    Set Err = TheError.LastError
    TheExec.Datalog.WriteComment pType & " : There is no free plane."
    TheExec.Datalog.WriteComment Err.Message
    Set GetFreePlaneForAcq = ret
    Exit Function

End Function

Private Function IAcquireAction_ToptStart() As Integer
'内容:
'   TestOptimizer機能を使って、AcquireFrameWorkを実行する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    '#####  AcquireFrameWorkを実行する  #####
    With m_FunctionInfo
        If .macroName <> "" Then
            On Error GoTo IGXLErrHandler
            Call TheHdw.TOPT.Start(.GroupName, .macroName, .InstanceName)
        Else
            On Error GoTo ErrHandler
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".ToptStart", "Can Not Found Acquire Frame Work Name ! ( " _
                                & .GroupName & " / " & .InstanceName & " )")
        End If
    End With


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
IGXLErrHandler:
    IAcquireAction_ToptStart = TL_ERROR
    m_Status = IAcquireAction_ToptStart
    Exit Function
ErrHandler:
    IAcquireAction_ToptStart = TL_ERROR
    m_Status = IAcquireAction_ToptStart
    Exit Function


End Function

Private Function IAcquireAction_ClearStatus() As Integer
'内容:
'   AcquireFrameWorkインスタンスの実行履歴を削除する
'
'パラメータ:
'
'戻り値:
'   TL_SUCCESS : 正常終了
'   TL_ERROR   : エラー終了
'
'注意事項:
'


    On Error GoTo ErrHandler


    '#####  AcquireFrameWorkインスタンスの実行履歴を削除する  #####
    Call TheHdw.TOPT.ClearStatus(m_FunctionInfo.GroupName)


    '#####  終了  #####
    Exit Function


'#####  エラーメッセージ処理＆終了  #####
ErrHandler:
    IAcquireAction_ClearStatus = TL_ERROR
    Exit Function


End Function

Private Property Get IAcquireAction_GetFrameWorkName() As String
    IAcquireAction_GetFrameWorkName = m_FunctionInfo.macroName
End Property

Private Property Get IAcquireAction_GetStatus() As Integer
    IAcquireAction_GetStatus = m_Status
End Property

Private Sub IAcquireAction_Reset()
    m_Status = -1
End Sub

Private Property Get IAcquireAction_AcquireFunctionInfo() As CAcquireFunctionInfo
    Set IAcquireAction_AcquireFunctionInfo = m_FunctionInfo
End Property

Private Sub IAcquireAction_Initialize(reader As IParameterReader)
    Set m_FunctionInfo = New CAcquireFunctionInfo
    Call m_FunctionInfo.Initialize(reader)
    m_InstanceName = reader.ReadAsString("Arg0@Parameters")
    m_ActionName = reader.ReadAsString("Arg1@Parameters")
    m_Status = -1
End Sub

Private Sub Class_Initialize()
    Set m_FunctionInfo = Nothing
    m_InstanceName = ""
    m_ActionName = ""
    m_Status = -1
End Sub

Private Sub Class_Terminate()
    Set m_FunctionInfo = Nothing
    m_InstanceName = ""
    m_ActionName = ""
    m_Status = -1
End Sub


