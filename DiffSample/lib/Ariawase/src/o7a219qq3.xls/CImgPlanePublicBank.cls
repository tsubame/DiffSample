VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlanePublicBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   プレーンを保存するためのクラス
'目的:
'   ParameterBankの仲介者。実際に保持しているのはParameterBank
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'

Option Explicit

Implements IImgPlaneBank

Private m_ParameterBank As IParameterBank

Public Function AsIImgPlaneBank() As IImgPlaneBank
    Set AsIImgPlaneBank = Me
End Function

Public Sub Init(pParamBank As IParameterBank)
'ParameterBankの参照を持っておく必要がある。
'結合時に完成予定
    Set m_ParameterBank = pParamBank
End Sub

Private Sub IImgPlaneBank_Add(ByVal pName As String, ByRef pPlane As CImgPlane, Optional ByVal pFixedPlane As Boolean = False, Optional ByVal pReadOnly As Boolean = False)
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   [pName]         In  1):
'   [pPlane]        In  2):
'   [pFixedPlane]   In  3):
'   [pReadOnly]     In  4):
'戻り値:
'
'注意事項:
'PrameterBankメソッドとの違い
'   受ける型がCImgPlane型
'IProcParameter型の付属品
'   DeviceInfo  --> Nothing
'   Plane       --> CImgPlane

    Dim pParam As CParamPlane
    Set pParam = New CParamPlane
    Set pParam.plane = pPlane
    
    pParam.plane.ReadOnly = pReadOnly
    If pParam.plane.Comment = "" Then
        pParam.plane.Comment = pName
    End If
    
    Dim pProcParameter As IProcParameter
    Set pProcParameter = pParam         'キャスト
    pProcParameter.Name = pName
    
    Call m_ParameterBank.Add(pProcParameter, pName, pFixedPlane)

    Set pParam = Nothing
    Set pProcParameter = Nothing

End Sub

Private Property Get IImgPlaneBank_IsOverwriteMode() As Boolean
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   なし
'戻り値:
'   Boolean
'
'注意事項:
'PrameterBankメソッドとの違い
'
    IImgPlaneBank_IsOverwriteMode = m_ParameterBank.IsOverwriteMode
End Property

Private Property Let IImgPlaneBank_IsOverwriteMode(ByRef mode As Boolean)
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   [mode]  In/Out  1):
'戻り値:
'
'注意事項:
'PrameterBankメソッドとの違い
'   ParameterBank全てに影響します。
    m_ParameterBank.IsOverwriteMode = mode
End Property

Private Function IImgPlaneBank_Item(ByVal pName As String) As CImgPlane
'内容:
'   キーに該当するCImgPlane Objを返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   [pName] In/Out  1):
'戻り値:
'   CImgPlane
'
'注意事項:
'PrameterBankメソッドとの違い
'
'パラメータバンクから受け取って画像だったら画像オブジェクトを返す。
    Dim pParameter As IProcParameter
    Set pParameter = m_ParameterBank.Item(pName)
    If TypeName(pParameter) = "CParamPlane" Then
        Dim pParamPlane As CParamPlane
        Set pParamPlane = pParameter
        Set IImgPlaneBank_Item = pParamPlane.plane
        Exit Function
    End If
    
    Set IImgPlaneBank_Item = Nothing

End Function

Private Sub IImgPlaneBank_Delete(ByVal pName As String)
'内容:
'   キーの内容が画像かどうか判定して、画像なら削除。違ったらエラー。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   [pName] In  1):
'戻り値:
'
'注意事項:
'PrameterBankメソッドとの違い
'   画像のみ削除可能です。

    Dim pParameter As IProcParameter
    Set pParameter = m_ParameterBank.Item(pName)
    If TypeName(pParameter) = "CParamPlane" Then
        Call m_ParameterBank.Delete(pName)
        Exit Sub
    Else
        Call TheError.Raise(9999, "CImgPlanePublicBank", "This Key can not delete Plane: " & pName)
    End If

End Sub

'OK
Private Function IImgPlaneBank_isExisting(ByVal pName As String) As Boolean
'内容:
'   ParameterBankに登録されているかを返します。（画像のみ）
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   [pName] In/Out  1):
'戻り値:
'   Boolean
'
'注意事項:
'PrameterBankメソッドとの違い
'   登録されている画像のみ対象です。
    
    On Error GoTo NOTHING_OBJ
    
    If m_ParameterBank.IsExist(pName) = True Then
        If TypeName(m_ParameterBank.Item(pName)) = "CParamPlane" Then
            IImgPlaneBank_isExisting = True
            Exit Function
        End If
    Else
        IImgPlaneBank_isExisting = False
    End If
    
NOTHING_OBJ:
    IImgPlaneBank_isExisting = False

End Function

'OK
Private Sub IImgPlaneBank_Clear()
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'PrameterBankメソッドとの違い
'   ありません。
'   このメソッドを実行することは、ParameterBankのClearを実行することと同じです。
'
    Call m_ParameterBank.Clear
End Sub

'OK
Private Function IImgPlaneBank_List() As String
'内容:
'   画像のリストを返します。
'   ParameterBankに登録されている画像のリスト
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   なし
'戻り値:
'   String
'
'注意事項:
'   PrameterBankメソッドとの違い
'   画像のみのリストを返します。
'   返ってくるString型の書式も若干違います。
'   ・ParameterBankから貰えるList情報の書式
'   　キー１(Fixed or Non),オブジェクトの型名１,キー２(Fixed or Non),オブジェクトの型名２,・・・
'   ・PlaneBankのList情報の書式
'   　登録名１(Fixed or Non),画像名１ & vbcrlf & 登録名２(Fixed or Non),画像名２ & vbcrlf

    Dim listArr()
    Dim PlaneList As Variant
    Dim pParamPlaneObj As CParamPlane
    Dim pKey As String
        
    If IImgPlaneBank_Count = 0 Then
        IImgPlaneBank_List = ""
        Exit Function
    End If
    
    PlaneList = Split(m_ParameterBank.List, ",")
    
    ReDim listArr(UBound(PlaneList) / 2 - 1, 1)
    
    Dim i As Long
    For i = 0 To UBound(listArr, 1)
        listArr(i, 0) = PlaneList(2 * i + 0)
        listArr(i, 1) = PlaneList(2 * i + 1)
    Next i
    
    For i = 0 To UBound(listArr, 1)
        If listArr(i, 1) = "CParamPlane" Then
            pKey = Replace(listArr(i, 0), "(Fixed)", "")
            Set pParamPlaneObj = m_ParameterBank.Item(pKey)
            IImgPlaneBank_List = IImgPlaneBank_List & listArr(i, 0) & "," & pParamPlaneObj.plane.Name & vbCrLf
        End If
    Next i

End Function

'OK
Private Property Get IImgPlaneBank_Count() As Long
'内容:
'   登録画像の数を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   なし
'戻り値:
'   Long
'
'注意事項:
'   PrameterBankメソッドとの違い
'   画像のみのカウントを返します。


    Dim listArr()
    Dim PlaneList As Variant

    PlaneList = Split(m_ParameterBank.List, ",")
    
    '何もないとubound=-1,lbound=0
    If UBound(PlaneList) = -1 Then
        IImgPlaneBank_Count = 0
        Exit Property
    End If
    
    Dim i As Long
    Dim j As Long
    For i = 0 To UBound(PlaneList) - 1
        If PlaneList(i) = "CParamPlane" Then
            j = j + 1
        End If
    Next i
    
    IImgPlaneBank_Count = j

End Property

'OK
Private Sub IImgPlaneBank_Display()
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'PrameterBankメソッドとの違い
'   ParameterBankのメソッドに、Displayメソッドはありません。
    
    Dim disp As New PlaneBankDisplay
    Call disp.Display

End Sub

Private Sub Class_Initialize()
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'
    Set m_ParameterBank = Nothing
End Sub

Private Sub Class_Terminate()
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月9日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'
    Set m_ParameterBank = Nothing
End Sub

