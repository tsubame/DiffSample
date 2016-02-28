VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgFlagPlanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   複数フラグプレーンの管理クラス
'
'目的:
'   フラグプレーンを複数利用可能にする

Option Explicit

Private m_FlagPlanes As Collection
Private Const MAX_FLAG_BIT As Long = 16

'#FlagExpansion
'初期化 #Pass
Public Sub Initialize(ByVal pGroupName As String)
    Set m_FlagPlanes = New Collection
    Dim FlagPlane As IImgFlag
    Set FlagPlane = New CImgFlag
    FlagPlane.group = pGroupName
    m_FlagPlanes.Add FlagPlane
End Sub

'#FlagExpansion
'インターセプター有効化 #Pass
Public Sub EnableInterceptor()
    
    Dim FlagPlane As IImgFlag
    Dim Interceptor As CImgFlagInterceptor
    Dim TempPlanes As Collection
    Set TempPlanes = New Collection
    
    '型判定による現状確認と割り込み設定
    For Each FlagPlane In m_FlagPlanes
        If TypeName(FlagPlane) = "CImgFlagInterceptor" Then
            '割り込み済み、そのまま登録
            TempPlanes.Add FlagPlane
        Else
            'インターセプター生成と割り込み登録
            Set Interceptor = New CImgFlagInterceptor
            Set Interceptor.TrueInstance = FlagPlane
            TempPlanes.Add Interceptor.AsIImgFlag
        End If
    Next
    
    '置き換え
    Set m_FlagPlanes = TempPlanes
    
End Sub

'#FlagExpansion
'インターセプター無効化 #Pass
Public Sub DisableInterceptor()
    
    Dim FlagPlane As IImgFlag
    Dim Interceptor As CImgFlagInterceptor
    Dim TempPlanes As Collection
    Set TempPlanes = New Collection
    
    '型判定による現状確認と割り込み設定
    For Each FlagPlane In m_FlagPlanes
        If TypeName(FlagPlane) = "CImgFlag" Then
            '割り込み無し、そのまま登録
            TempPlanes.Add FlagPlane
        Else
            'インターセプター取り外しと登録
            Set Interceptor = FlagPlane '型変換
            TempPlanes.Add Interceptor.TrueInstance
        End If
    Next
    
    '置き換え
    Set m_FlagPlanes = TempPlanes
    
End Sub

'#FlagExpansion
'すべてのフラグプレーンの強制開放 #Pass
Public Sub ClearAllFlag()
    'プレーン開放
    Dim FlagPlane As IImgFlag
    For Each FlagPlane In m_FlagPlanes
        FlagPlane.Clear
    Next
    'フラグインスタンスを1つだけの状態にする
    If m_FlagPlanes.Count >= 2 Then
        Dim MaxCount As Long
        Dim ItemIndex As Long
        MaxCount = m_FlagPlanes.Count
        For ItemIndex = 1 To MaxCount - 1
            m_FlagPlanes.Remove 2
        Next
    End If
End Sub

'#FlagExpansion
'存在するフラグプレーンすべてを返す #Pass
Public Function FlagPlanes() As Collection
    Set FlagPlanes = m_FlagPlanes
End Function

'#FlagExpansion
'指定フラグが存在するフラグプレーンを返す #Pass
Public Function FlagPlane(ByVal pFlagName As String) As IImgFlag
    
    Dim TargetPlane As IImgFlag
    
    '1:指定のフラグ名が登録されているプレーンを検索する
    For Each TargetPlane In m_FlagPlanes
        If TargetPlane.FlagBit(pFlagName) <> 0 Then
            Set FlagPlane = TargetPlane
            Exit Function 'すでに登録されているものを利用
        End If
    Next

    '2:フラグの空きがあるプレーンを検索する
    For Each TargetPlane In m_FlagPlanes
        If TargetPlane.Count < MAX_FLAG_BIT Then
            Set FlagPlane = TargetPlane '登録されているプレーンを返す
            Exit Function '空きの存在するものを返す
        End If
    Next
    
    '3:空きが存在しないので新規作成しそれを返す
    Dim ChkFlagPlane As IImgFlag
    Set ChkFlagPlane = m_FlagPlanes.Item(1)
    Dim NewFlagPlane As IImgFlag
    
    '現在のインターセプター状態を確認する
    If TypeName(ChkFlagPlane) = "CImgFlag" Then
        'インターセプター無し
        Set NewFlagPlane = New CImgFlag
        NewFlagPlane.group = ChkFlagPlane.group
        m_FlagPlanes.Add NewFlagPlane
        Set FlagPlane = NewFlagPlane
    Else
        'インターセプター有り
        Dim Interceptor As CImgFlagInterceptor
        Set Interceptor = New CImgFlagInterceptor
        Set NewFlagPlane = New CImgFlag
        NewFlagPlane.group = ChkFlagPlane.group
        Set Interceptor.TrueInstance = NewFlagPlane
        m_FlagPlanes.Add Interceptor.AsIImgFlag
        Set FlagPlane = Interceptor.AsIImgFlag
    End If

End Function

'#FlagExpansion
'フラグ利用個数0個のフラグプレーンを開放する関数
Public Function ReleaseUnusedFlagPlane() As Long
    Dim TargetPlane As IImgFlag
    Dim Count As Long
    For Each TargetPlane In m_FlagPlanes
        If TargetPlane.Count = 0 And (Not TargetPlane.FlgPlane Is Nothing) Then
            TargetPlane.Clear
            Count = Count + 1
        End If
    Next
    ReleaseUnusedFlagPlane = Count
End Function
