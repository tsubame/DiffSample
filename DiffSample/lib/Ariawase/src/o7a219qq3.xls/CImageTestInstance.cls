VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImageTestInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   TestInstance����ێ������s����N���X
'
'   Revision History:
'       Data        Description
'       2010/04/28  TestInstance(UserMacro)�����s����@�\����������
'       2010/05/12  �v���O�����R�[�h�𐮗�����
'       2010/05/25  ActionLogger�̈�����ύX����
'       2010/05/31  Error�����𒲐�����
'       2010/06/11  �v���O�����R�[�h�𐮗�����
'       2010/06/22  �v���O�����R�[�h�𐮗�����
'       2011/02/10  I/F�쐬���āA���K�[��Interceptor��
'
'�쐬��:
'   0145184346
'

Option Explicit

Private m_FlagInterceptor As Boolean

Private Const CLASS_NAME  As String = "CImageTestInstance" ' Class���̂�ێ�����

Private mTestInstanceName As String    ' TestInstance���̂�ێ�����
Private mAcquireAction As Collection   ' AcquireInstance���̂�ێ�����
Private mUserMacroAction As Collection ' UserMacro���s�N���X��ێ�����

Public Function Execute() As Integer
'���e:
'   TestInstance(UserMacro)�����s����
'
'�p�����[�^:
'
'�߂�l:
'   TL_SUCCESS : ����I��
'   TL_ERROR   : �G���[�I��
'
'���ӎ���:
'


    On Error GoTo ErrHandler


    '#####  TestInstance(UserMacro)�����s����  #####
    Dim clsUsrMcrAct As IUserMacroAction
    For Each clsUsrMcrAct In mUserMacroAction
'        Call clsActionLogger.LogBeforeExecution(clsActionInfo)
        Execute = clsUsrMcrAct.Execute
        If Execute = TL_ERROR Then
            Call TheError.Raise(9999, CLASS_NAME & ".Execute", "User Macro Return TL_ERROR ! ( " _
                                & clsUsrMcrAct.GetUserMacroName & " )")
        End If
'        Call clsActionLogger.LogAfterExecution(clsActionInfo, CLng(Execute))
    Next clsUsrMcrAct

    Set clsUsrMcrAct = Nothing

    '#####  �I��  #####
    Exit Function


'#####  �G���[���b�Z�[�W�������I��  #####
ErrHandler:
    Execute = TL_ERROR
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Exit Function


End Function

Public Function GetTestInstanceName() As String
     GetTestInstanceName = mTestInstanceName
End Function

Public Property Let SetTestInstanceName(ByRef strInsName As String)
     mTestInstanceName = strInsName
End Property

Public Function GetAcquireInstanceCount() As Integer
     GetAcquireInstanceCount = mAcquireAction.Count
End Function

Public Function GetAcquireInstanceName(ByRef cnt As Integer) As String
     GetAcquireInstanceName = mAcquireAction.Item(cnt)
End Function

Public Property Let SetAcquireInstance(ByRef colAcqAct As Collection)
     Set mAcquireAction = colAcqAct
End Property

Public Function GetUserMacroCount() As Integer
     GetUserMacroCount = mUserMacroAction.Count
End Function

Public Property Let SetUserMacro(ByRef clsUsrMcrAct As IUserMacroAction)
     mUserMacroAction.Add clsUsrMcrAct
End Property

Public Sub Initialize()
    mTestInstanceName = ""
    Set mAcquireAction = New Collection
    Set mUserMacroAction = New Collection
End Sub

Private Sub Class_Initialize()
    mTestInstanceName = ""
    Set mAcquireAction = New Collection
    Set mUserMacroAction = New Collection
End Sub

Private Sub Class_Terminate()
    mTestInstanceName = ""
    Set mAcquireAction = New Collection
    Set mUserMacroAction = New Collection
End Sub

Public Sub EnableInterceptor(ByVal pFlag As Boolean, ByRef pLogger As CActionLogger)
'�܂�
    
    Dim pObj As IUserMacroAction
    Dim pCollection As Collection
    Dim pInterceptor As CUserMacroActionInterceptor

    If m_FlagInterceptor = pFlag Then
        Exit Sub
    Else
        Set pCollection = New Collection
        For Each pObj In mUserMacroAction
            If pFlag = True Then
                Set pInterceptor = New CUserMacroActionInterceptor
                Set pInterceptor.TrueInstance = pObj
                Set pInterceptor.actionLogger = pLogger
                Call pCollection.Add(pInterceptor)
            Else
                Set pInterceptor = pObj
                Set pObj = pInterceptor.TrueInstance
                Set pInterceptor = Nothing
                Call pCollection.Add(pObj)
            End If
        Next pObj
        Set mUserMacroAction = Nothing
        Set mUserMacroAction = pCollection

        m_FlagInterceptor = pFlag

    End If

End Sub
