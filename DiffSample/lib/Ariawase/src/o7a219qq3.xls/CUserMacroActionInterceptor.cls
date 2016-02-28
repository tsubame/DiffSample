VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUserMacroActionInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IUserMacroAction

Private Const CLASS_NAME As String = "CUserMacroActionIntrerceptor" ' ClassñºèÃÇï€éùÇ∑ÇÈ

Private m_ActionLogger As CActionLogger
Private mUserMacroAcrtion As CUserMacroAction

Public Property Set actionLogger(pActionlog As CActionLogger)
    Set m_ActionLogger = pActionlog
End Property

Public Property Get AsIUserMacroAction() As IUserMacroAction
    Set AsIUserMacroAction = Me
End Property

Public Property Get TrueInstance() As IUserMacroAction
    Set TrueInstance = mUserMacroAcrtion
End Property

Public Property Set TrueInstance(Instance As IUserMacroAction)
    Set mUserMacroAcrtion = Instance
End Property

Private Function IUserMacroAction_Execute() As Integer
Debug.Print "CUserMacroActionInterceptor:Pre"
    Dim clsActionInfo As CActionInfo
    Set clsActionInfo = New CActionInfo

    With clsActionInfo
        .macroName = mUserMacroAcrtion.AsIUserMacroAction.GetUserMacroName
        .TestInstanceName = TheExec.DataManager.InstanceName
    End With
        
    Call m_ActionLogger.LogBeforeExecution(clsActionInfo)
    
    IUserMacroAction_Execute = mUserMacroAcrtion.AsIUserMacroAction.Execute
    
    Call m_ActionLogger.LogAfterExecution(clsActionInfo, CLng(IUserMacroAction_Execute))
Debug.Print "CUserMacroActionInterceptor:Post"

End Function

Private Function IUserMacroAction_GetUserMacroName() As String
    IUserMacroAction_GetUserMacroName = mUserMacroAcrtion.AsIUserMacroAction.GetUserMacroName
End Function

Private Property Let IUserMacroAction_SetUserMacroName(ByVal strName As String)
    mUserMacroAcrtion.AsIUserMacroAction.SetUserMacroName = strName
End Property

Private Sub IUserMacroAction_Initialize()
    Call mUserMacroAcrtion.AsIUserMacroAction.Initialize
End Sub

Private Sub Class_Terminate()
    Set mUserMacroAcrtion = Nothing
    Set m_ActionLogger = Nothing
End Sub

