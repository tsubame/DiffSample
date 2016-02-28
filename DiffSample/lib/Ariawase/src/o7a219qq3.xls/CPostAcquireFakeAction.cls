VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPostAcquireFakeAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   CPostAcquireAction用の画像をファイルから読み込みます。
'目的:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月10日

Option Explicit

Implements IAcquireAction
Implements IAcquireActionInstance

Private Const ERR_NUMBER = 9999             ' エラー番号を保持する
Private Const CLASS_NAME = "CPostAcquireFakeAction" ' Class名称を保持する

Private m_CPostAcquireAction As CPostAcquireAction

Public Property Get AsIAcquireAction() As IAcquireAction
    Set AsIAcquireAction = Me
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
    Set m_CPostAcquireAction = pInstance

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
    Set IAcquireActionInstance_AcquireActionInstance = m_CPostAcquireAction

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

    IAcquireAction_Execute = m_CPostAcquireAction.AsIAcquireAction.Execute

    Dim capParam As CParamPlane
    If TheParameterBank.IsExist(m_CPostAcquireAction.AsIAcquireAction.AcquireFunctionInfo.Arg(0)) Then
        Set capParam = TheParameterBank.Item(m_CPostAcquireAction.AsIAcquireAction.AcquireFunctionInfo.Arg(0))
    Else
        Exit Function
    End If
    
    '画像Read
    Dim pReadPlane As CImgPlane
    Set pReadPlane = capParam.plane
    If Not pReadPlane Is Nothing Then
        Dim ImgReader As CIGXLImageFileReader
        Dim pPath As String
        Dim pFileName As String
        Set ImgReader = New CIGXLImageFileReader
        Call DivPathAndFileName(m_CPostAcquireAction.AsIAcquireAction.AcquireFunctionInfo.Arg(6), pPath, pFileName)
        If pPath = "" And pFileName = "" Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Arg6 description is unjust! Instance Name : " & m_CPostAcquireAction.AsIAcquireAction.AcquireFunctionInfo.Arg(0))
        ElseIf pPath = "" Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "There are no pathname of a reading image file: " & m_CPostAcquireAction.AsIAcquireAction.AcquireFunctionInfo.Arg(0))
        ElseIf pFileName = "" Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "There are no filename of a reading image file: " & m_CPostAcquireAction.AsIAcquireAction.AcquireFunctionInfo.Arg(0))
        End If
        On Error GoTo READ_FAIL
        Call ImgReader.Initialize(pPath, pFileName)
        Call capParam.AsIProcParameter.Read(ImgReader)
        
        If TheExec.sites.ActiveCount > 0 Then
            Dim site As Long
            Dim logWriter As IProcParamWriter
            Set logWriter = New CDatalogWriter
            logWriter.WriteAsString "", "##### AcquireTable Read Image #####"
            For site = 0 To TheExec.sites.ExistingCount - 1
                If TheExec.sites.site(site).Active Then
                    logWriter.WriteAsBlockStart ""
                    logWriter.WriteAsString "READ IMAGE", " " & ImgReader.ImgFileName(site) & " -> " & capParam.AsIProcParameter.Name & "[Site" & site & "]"
                    logWriter.WriteAsBlockEnd ""
                End If
            Next site
            Set logWriter = Nothing
        End If
    End If
    
    Exit Function
    
    '画像読み込み失敗
READ_FAIL:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Failed in image file reading : " & m_CPostAcquireAction.AsIAcquireAction.AcquireFunctionInfo.Arg(0))
    
End Function

Private Function IAcquireAction_ToptStart() As Integer
    IAcquireAction_ToptStart = m_CPostAcquireAction.AsIAcquireAction.ToptStart
End Function

Private Function IAcquireAction_ClearStatus() As Integer
    IAcquireAction_ClearStatus = m_CPostAcquireAction.AsIAcquireAction.ClearStatus
End Function

Private Property Get IAcquireAction_GetFrameWorkName() As String
    IAcquireAction_GetFrameWorkName = m_CPostAcquireAction.AsIAcquireAction.GetFrameWorkName
End Property

Private Property Get IAcquireAction_GetStatus() As Integer
    IAcquireAction_GetStatus = m_CPostAcquireAction.AsIAcquireAction.GetStatus
End Property

Private Sub IAcquireAction_Reset()
    Call m_CPostAcquireAction.AsIAcquireAction.Reset
End Sub

Private Sub IAcquireAction_Initialize(reader As IParameterReader)
    Call m_CPostAcquireAction.AsIAcquireAction.Initialize(reader)
End Sub

Private Property Get IAcquireAction_AcquireFunctionInfo() As CAcquireFunctionInfo
    Set IAcquireAction_AcquireFunctionInfo = m_CPostAcquireAction.AsIAcquireAction.AcquireFunctionInfo
End Property

Private Sub DivPathAndFileName(ByVal pOrg As String, ByRef retPath As String, ByRef RetFileName As String)
    Dim pName() As String
    pOrg = Replace(pOrg, """", "")
    pName = Split(pOrg, ",")
    If UBound(pName) > 0 Then
        retPath = pName(0)
        RetFileName = pName(1)
    Else
        retPath = ""
        RetFileName = ""
    End If

End Sub

Private Sub Class_Initialize()
    Set m_CPostAcquireAction = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_CPostAcquireAction = Nothing
End Sub

