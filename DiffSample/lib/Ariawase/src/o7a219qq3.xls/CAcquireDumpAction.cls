VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcquireDumpAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   CAcquireAction,CPostAcquireActionに割り込んで、画像をDumpする。
'目的:
'   画像表示、書き込み用
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月8日
'
'備考
'CAcquireActionに被る


Option Explicit

Implements IAcquireAction
Implements IAcquireActionInstance

Private Const ERR_NUMBER = 9999             ' エラー番号を保持する
Private Const CLASS_NAME = "CAcquireDumpAction" ' Class名称を保持する

Private m_IAcquireAction As IAcquireAction
Private m_FlagShow As Boolean
Private m_FlagSaveImage As Boolean

Private m_Status As Long                       ' 実行結果を保持する（-1：実行前　0：正常終了　1：エラー発生）

Public Property Get AsIAcquireAction() As IAcquireAction
    Set AsIAcquireAction = Me
End Property

Public Property Let FlagShow(p_FlagShow As Boolean)
    m_FlagShow = p_FlagShow
End Property

Public Property Get FlagShow() As Boolean
    FlagShow = m_FlagShow
End Property

Public Property Let FlagSaveImage(p_FlagShow As Boolean)
    m_FlagSaveImage = p_FlagShow
End Property

Public Property Get FlagSaveImage() As Boolean
    FlagSaveImage = m_FlagSaveImage
End Property

Private Property Set IAcquireActionInstance_AcquireActionInstance(pInstance As IAcquireAction)
'内容:
'   このクラスにIAcquireActionインターフェイスのインスタンスをセットします。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   なし
'戻り値:
'   IAcquireAction
'
'注意事項:
'
    Set m_IAcquireAction = pInstance

End Property

Private Property Get IAcquireActionInstance_AcquireActionInstance() As IAcquireAction
'内容:
'   このクラスが持つIAcquireActionインターフェイスのインスタンスを返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   なし
'戻り値:
'   IAcquireAction
'
'注意事項:
'
    Set IAcquireActionInstance_AcquireActionInstance = m_IAcquireAction

End Property

Private Function IAcquireAction_Execute() As Integer
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   なし
'戻り値:
'   Integer
'
'注意事項:
'
'
    
    IAcquireAction_Execute = m_IAcquireAction.Execute
    
    Dim capParam As CParamPlane
    If TheParameterBank.IsExist(m_IAcquireAction.AcquireFunctionInfo.Arg(0)) Then
        Set capParam = TheParameterBank.Item(m_IAcquireAction.AcquireFunctionInfo.Arg(0))
    Else
        Exit Function
    End If
    
    '### 画像保存機能 ############
    If m_FlagSaveImage = True Then
        
        'CIGXLImageISTARFileWriter
        Dim ImgWriter As CIGXLImageISTARFileWriter
        Set ImgWriter = New CIGXLImageISTARFileWriter
        'Set Data
        Call ImgWriter.Initialize(m_IAcquireAction.AcquireFunctionInfo.Arg(8), m_IAcquireAction.AcquireFunctionInfo.InstanceName)
        Call capParam.AsIProcParameter.Dump(ImgWriter)
        
        Dim site As Long
        Dim extFlag As Boolean
        extFlag = False
        For site = 0 To TheExec.sites.ExistingCount - 1
            If Dir(ImgWriter.ImgFileName(site)) <> "" Then
                extFlag = True
                Exit For
            End If
        Next site
        If extFlag = True Then
            Dim logWriter As IProcParamWriter
            Set logWriter = New CDatalogWriter
            logWriter.WriteAsString "", "##### AcquireTable Wrote Image #####"
            For site = 0 To TheExec.sites.ExistingCount - 1
                If Dir(ImgWriter.ImgFileName(site)) <> "" And TheExec.sites.site(site).Active Then
                    logWriter.WriteAsBlockStart ""
                    logWriter.WriteAsString "WROTE IMAGE", " " & capParam.AsIProcParameter.Name & "[Site" & site & "] -> " & ImgWriter.ImgFileName(site)
                    logWriter.WriteAsBlockEnd ""
                End If
            Next site
            Set logWriter = Nothing
            Set ImgWriter = Nothing
        Else
            If TheExec.sites.ActiveCount > 0 Then
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Failed in Image writing : " & m_IAcquireAction.AcquireFunctionInfo.Arg(0))
            End If
        End If
    End If
    
    
    
    '### 画像表示機能 ############
    If m_FlagShow = True Then

        'CScenarioFormDebugger
        Dim FormWriter As CScenarioFormDebugger
        Dim QuitFlag As Boolean
        
        Set FormWriter = New CScenarioFormDebugger
        
        Call capParam.AsIProcParameter.Dump(FormWriter)
        
        Call FormWriter.AsIParameterViewer.View
        QuitFlag = FormWriter.AsIParameterViewer.EndStatus
        
        Set FormWriter = Nothing
        
    End If
    
    Set capParam = Nothing
    
    If QuitFlag = True Then
        m_Status = TL_ERROR
    End If
    
    Exit Function

End Function

Private Function IAcquireAction_ToptStart() As Integer
    IAcquireAction_ToptStart = m_IAcquireAction.ToptStart
End Function

Private Function IAcquireAction_ClearStatus() As Integer
    IAcquireAction_ClearStatus = m_IAcquireAction.ClearStatus
End Function

Private Property Get IAcquireAction_GetFrameWorkName() As String
    IAcquireAction_GetFrameWorkName = m_IAcquireAction.GetFrameWorkName
End Property

Private Property Get IAcquireAction_GetStatus() As Integer
    If m_Status = 1 Then
        IAcquireAction_GetStatus = m_Status
    Else
        IAcquireAction_GetStatus = m_IAcquireAction.GetStatus
    End If
End Property

Private Sub IAcquireAction_Reset()
    m_Status = -1
    Call m_IAcquireAction.Reset
End Sub

Private Sub IAcquireAction_Initialize(reader As IParameterReader)
    Call m_IAcquireAction.Initialize(reader)
End Sub

Private Property Get IAcquireAction_AcquireFunctionInfo() As CAcquireFunctionInfo
    Set IAcquireAction_AcquireFunctionInfo = m_IAcquireAction.AcquireFunctionInfo
End Property

Private Sub Class_Initialize()
    Set m_IAcquireAction = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_IAcquireAction = Nothing
End Sub




