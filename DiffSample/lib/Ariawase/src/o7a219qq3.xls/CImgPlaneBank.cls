VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   プレーンを保存するためのクラス
'
'目的:
'   グローバル変数の使用を避けるため、保存用のクラスを用意。
'   名前をつけてCImgPlaneを関連付ける。
'
'作成者:
'   0145184004
'
'2011/02/09
'ParameterBankの使用に伴いこのクラスは使用しません。
'TheIDPからPlaneBankを使用するときは、CImgPlanePublicBank

Option Explicit

Implements IImgPlaneBank

Private m_Names As Collection
Private m_Planes As Collection

Private m_FixedNames As Collection
Private m_FixedPlanes As Collection
Private m_OverWriteMode As Boolean


Private Sub IImgPlaneBank_Add(ByVal pName As String, ByRef pPlane As CImgPlane, Optional ByVal pFixedPlane As Boolean = False, Optional ByVal pReadOnly As Boolean = False)
'内容:
'   プレーンを登録する。
'   pFixedPlaneをTrueにするとClearで一括解放されなくなる。
'   pReadOnlyをTrueにすると、画像を変更する関数(例:WritePixel, MultiMean, etc)が使えなくなる。
'    ※入力には使うことは可能。Averageなどデータの取得も可能。
'
'[pName]       IN   String型:       識別名
'[pPlane]      IN   CImgPlane型:    登録するプレーン
'[pFixedPlane] IN   Boolean型:      固定プレーンオプション
'[pReadOnly]   IN   Boolean型:      リードオンリーオプション
'
'備考:
'   エラー条件
'   1,指定した名前が既に存在している時
'
'   サンプルコード
'    Dim src As CImgPlane
'    Set src = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call TheIDP.PlaneBank.Add("SOURCE", src)
'
    
    'エラー
''    If IImgPlaneBank_isExisting(pName) = True Then
''        Call TheError.Raise(999, "Add", ErrMsgNameAlreadyExists(pName))
''        Exit Sub
''    End If
    If IImgPlaneBank_IsOverwriteMode = False Then
        If IImgPlaneBank_isExisting(pName) = True Then
            Call TheError.Raise(999, "Add", ErrMsgNameAlreadyExists(pName))
            Exit Sub
        End If
    Else
        'Bank登録に上書き。TOPT Acquireやり直し用
        If isExisting_(m_FixedNames, pName) = True Then
            Call m_FixedNames.Remove(pName)
            Call m_FixedPlanes.Remove(pName)
        ElseIf isExisting_(m_Names, pName) = True Then
            Call m_Names.Remove(pName)
            Call m_Planes.Remove(pName)
        End If
        '登録無ければ何もしない
    End If

    
    pPlane.ReadOnly = pReadOnly
    If pPlane.Comment = "" Then
        pPlane.Comment = pName
    End If
    
    If pFixedPlane = True Then
        Call m_FixedNames.Add(pName, pName)
        Call m_FixedPlanes.Add(pPlane, pName)
    Else
        Call m_Names.Add(pName, pName)
        Call m_Planes.Add(pPlane, pName)
    End If
    
End Sub

Private Property Get IImgPlaneBank_IsOverwriteMode() As Boolean
    IImgPlaneBank_IsOverwriteMode = m_OverWriteMode
End Property

Private Property Let IImgPlaneBank_IsOverwriteMode(ByRef mode As Boolean)
    m_OverWriteMode = mode
End Property

Private Function IImgPlaneBank_Item(ByVal pName As String) As CImgPlane
'内容:
'   指定した名前のプレーンを取得する
'
'[pName]        IN      String型:       識別名
'
'備考:
'   エラー条件
'   1,指定した名前が存在しない時
'
'   サンプルコード
'    Dim src As CImgPlane
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'
    Set IImgPlaneBank_Item = GetPlane(pName, m_Planes)
    If IImgPlaneBank_Item Is Nothing Then
        Set IImgPlaneBank_Item = GetPlane(pName, m_FixedPlanes)
    End If

    'エラー
    If IImgPlaneBank_Item Is Nothing Then
        Call TheError.Raise(999, "Item", ErrMsgNameDoesntExist(pName))
    End If
End Function

Private Sub IImgPlaneBank_Delete(ByVal pName As String)
'内容:
'   指定した名前のプレーンの登録を破棄する。
'
'[pName]       IN/OUT String型:     識別名
'
'備考:
'   エラー条件
'   1,指定した名前が存在しない時
'
'   サンプルコード
'    Call TheIDP.PlaneBank.Delete("SOURCE")
'
    If Not GetPlane(pName, m_Planes) Is Nothing Then
        Call m_Names.Remove(pName)
        Call m_Planes.Remove(pName)
    ElseIf Not GetPlane(pName, m_FixedPlanes) Is Nothing Then
        Call m_FixedNames.Remove(pName)
        Call m_FixedPlanes.Remove(pName)
    Else
        'エラー
        Call TheError.Raise(999, "Delete", ErrMsgNameDoesntExist(pName))
    End If

End Sub

Private Function IImgPlaneBank_isExisting(ByVal pName As String) As Boolean
'内容:
'   指定した名前の登録があるかどうかの確認。
'
'[pName]       IN/OUT String型:     識別名
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.PlaneBank.isExisting("SOURCE")
'

    If (Not isExisting_(m_Names, pName)) And (Not isExisting_(m_FixedNames, pName)) Then
        IImgPlaneBank_isExisting = False
    Else
        IImgPlaneBank_isExisting = True
    End If
    
End Function

Private Function isExisting_(ByRef pPlanes As Collection, ByVal pName As String) As Boolean

    Dim tmp As String
    On Error GoTo NOT_FOUND
    tmp = pPlanes.Item(pName)
    
    isExisting_ = True
    Exit Function
    
NOT_FOUND:
    isExisting_ = False

End Function

Private Sub IImgPlaneBank_Clear()
'内容:
'   固定プレーン以外の登録されたプレーンをすべて解放する。
'
'備考:
'   サンプルコード
'    Call TheIDP.PlaneBank.Clear
'
    Set m_Planes = New Collection
    Set m_Names = New Collection
End Sub

Private Function IImgPlaneBank_List() As String
'内容:
'   登録されたプレーンのリストを取得
'
'備考:
'   カンマ区切りの文字列で出力。
'   登録名,プレーン名,…　の順
'
'   サンプルコード
'    Debug.Print TheIDP.PlaneBank.List
'
    Dim tmp As Variant
    For Each tmp In m_FixedNames
        IImgPlaneBank_List = IImgPlaneBank_List & tmp & "(Fixed)" & "," & m_FixedPlanes(tmp).Name & vbCrLf
    Next tmp
    
    For Each tmp In m_Names
        IImgPlaneBank_List = IImgPlaneBank_List & tmp & "," & m_Planes(tmp).Name & vbCrLf
    Next tmp
    
End Function

Private Property Get IImgPlaneBank_Count() As Long
'内容:
'   登録されたプレーンの数を取得
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.PlaneBank.Count
'
    IImgPlaneBank_Count = m_Planes.Count + m_FixedPlanes.Count
End Property

Private Sub IImgPlaneBank_Display()
'内容:
'   登録されたプレーンのリストを表示
'
'備考:
'   サンプルコード
'    Call TheIDP.PlaneBank.Display
'
    Dim disp As New PlaneBankDisplay
    Call disp.Display
End Sub

Private Function GetPlane(ByVal pName As String, ByRef pPlanes As Collection) As CImgPlane

On Error GoTo NOT_FOUND
    Set GetPlane = pPlanes(pName)
    
    Exit Function
    
NOT_FOUND:
    Set GetPlane = Nothing
    
End Function

Private Sub Class_Initialize()

    Set m_Names = New Collection
    Set m_Planes = New Collection
    Set m_FixedNames = New Collection
    Set m_FixedPlanes = New Collection
    
End Sub

Private Sub Class_Terminate()

    Set m_Names = Nothing
    Set m_Planes = Nothing
    Set m_FixedNames = Nothing
    Set m_FixedPlanes = Nothing
    
End Sub

Private Function ErrMsgNameDoesntExist(ByVal pName As String)
    ErrMsgNameDoesntExist = """" & pName & """ doesn't exist."
End Function

Private Function ErrMsgNameAlreadyExists(ByVal pName As String)
    ErrMsgNameAlreadyExists = """" & pName & """ already exists."
End Function

Public Function AsIImgPlaneBank() As IImgPlaneBank
    Set AsIImgPlaneBank = Me
End Function


