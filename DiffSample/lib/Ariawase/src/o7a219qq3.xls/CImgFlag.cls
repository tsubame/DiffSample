VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   フラグの管理
'
'目的:
'   フラグを名前で管理する。
'
'作成者:
'   0145184004
'
' 2009/06/09 D.Maruyama フラグのクリアを最小限にするために大幅変更
' 2010/11/12 t.takase   EeeJOB V3.0 idpColorTypeの指定を任意の文字で可能にするため変更

Option Explicit

Implements IImgFlag

'ビットの状態
Private Enum FlagUsingState
    Uninitialized   '初期状態、　なにが入っているか不明
    INITIALIZED     '初期化済み(クリア済み)
    Used            '使用されている
    NotUsed         '使用が終わって開放された
End Enum

Private Type sFlagState
    Count As Long
    flagName(1 To 16) As String
    FlagState(1 To 16) As FlagUsingState
End Type

Private m_FlgPlane As CImgPlane
Private m_FlgState As sFlagState
Private m_Group As String

'#FlagExpansion
Private Property Get IImgFlag_Count() As Long
    IImgFlag_Count = FlgCount_
End Property

'#FlagExpansion
Private Property Get IImgFlag_Group() As String
    IImgFlag_Group = m_Group
End Property

Private Property Let IImgFlag_Group(ByVal RHS As String)
    m_Group = RHS
End Property

Private Property Set IImgFlag_FlgPlane(ByRef pPlane As CImgPlane)
    If Not m_FlgPlane Is Nothing Then Exit Property
    
    Set m_FlgPlane = pPlane
    If m_FlgPlane.Comment = "" Then
        m_FlgPlane.Comment = "Shared Flag Plane"
    End If
End Property

Private Property Get IImgFlag_FlgPlane() As CImgPlane
    Set IImgFlag_FlgPlane = m_FlgPlane
End Property

Private Sub IImgFlag_Clear()
'内容:
'   登録したフラグの一括クリア
    
    Call IImgFlag_ResetFlagState
    Set m_FlgPlane = Nothing

End Sub

Private Property Get IImgFlag_FlagBit(ByVal pName As String) As Long
'内容:
'   指定した名前のフラグ位置を取得
'
'[pName]       IN String型:     名前指定
'
'備考:
'   指定した名前のフラグが存在しない場合0を返す。
'

    Dim i As Long
    
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    For i = minBit To maxBit
        If pName = m_FlgState.flagName(i) Then
            IImgFlag_FlagBit = i
            Exit Property
        End If
    Next i
    
    IImgFlag_FlagBit = 0

End Property

'#FlagExpansion
Private Function IImgFlag_SetFlagBit(ByVal pName As String) As Long
'内容:
'   指定した名前のフラグをセットする。
'   フラグがセットされた場合、そのビットを0クリアしフラグ番号を返す。
'   既存の名前の場合はその番号を、フラグが埋まっている場合は0を返す。
'
'[pName]       IN String型:     名前指定
'
'備考:
'   フラグの空きがない場合0を返す。
'

    IImgFlag_SetFlagBit = IImgFlag_FlagBit(pName)
    If IImgFlag_SetFlagBit <> 0 Then Exit Function
    
    Dim oldPmd As CImgPmdInfo
    

    '個数が0のときはフラグプレンがないので確保しにいく
    If FlgCount_ = 0 And m_FlgPlane Is Nothing Then
    
'2009/09/03 D.Maruayama ゼロクリアは一回だけにする
'        Set m_FlgPlane = TheIDP.PlaneManager(m_Group).GetFreePlane(idpDepthS16, True)
        Set m_FlgPlane = TheIDP.PlaneManager(m_Group).GetFreePlane(idpDepthS16)
        m_FlgPlane.Comment = "Shared Flag Plane"
        Call m_FlgPlane.SetPMD(m_FlgPlane.BasePMD).WritePixel(0)
        Call InitializeFlagState
                
    End If
    
    '初期化済み（一度も使っていない）フラグビットをとってくる
    IImgFlag_SetFlagBit = GetInitializedFlag
    
    
    'とってこれなかったら使用済みフラグをクリアして、初期化済みビットをとってくる
    If IImgFlag_SetFlagBit = 0 Then
    
        Set oldPmd = m_FlgPlane.CurrentPMD
        
        Call RecoverFlagPlane(m_FlgPlane)
        
        If oldPmd.Name = m_FlgPlane.Manager.VariablePMDName Then
            Call m_FlgPlane.SetPMD(oldPmd)
        Else
            Call m_FlgPlane.SetPMD(oldPmd.Name)
        End If
        
        IImgFlag_SetFlagBit = GetInitializedFlag
        
    End If
    
    
    'ここまでにとってこれたら状態を変更する,だめならエラーを通知
    If IImgFlag_SetFlagBit <> 0 Then
        Call SetUsedState(IImgFlag_SetFlagBit, pName)
        Exit Function
    End If
        
    Call TheError.Raise(999, "SetFlagBit", "There's no bit for the flag")
End Function

'#FlagExpansion
Private Sub IImgFlag_RemoveFlagBit(ByVal pName As String)
'内容:
'   指定した名前のフラグを削除する
'
'[pName]       IN String型:     名前指定
'

    Dim i As Long
        
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    For i = minBit To maxBit
        If m_FlgState.flagName(i) = pName Then
            Call SetNotUsedState(i)
'利用個数0個になっても自動開放はしない
'            If FlgCount_ = 0 Then
'                Set m_FlgPlane = Nothing
'            End If
            Exit Sub
        End If
    Next i

End Sub

'#FlagExpansion
Private Function GetFlgBit(ByVal pName As String) As Long
    '自身の中で確認
    Call OtherGroupGetFlag(pName)
    
    GetFlgBit = IImgFlag_FlagBit(pName)
    '他者のフラグプレーンに問い合わせ
    If GetFlgBit = 0 Then
        GetFlgBit = m_FlgPlane.Manager.GetSharedFlagPlane(pName).FlagBit(pName)
    End If
    'どこにも無い
    If GetFlgBit = 0 Then Call Err.Raise(999, "", pName & " is not exist")
End Function



'''' 0000130343
Private Function OtherGroupGetFlag(ByVal pFlgName As String)

    Dim i As Integer
    Dim sPlane As CImgPlane
    Dim resPlane As CImgPlane
    Dim resBit As Integer
    Dim dstPlane As CImgPlane
    Dim dstBit As Integer

    ' Debug matsu
    Dim resZone As String
    Dim dstZone As String

    '' pFlgName がNULLの場合は、この処理は不要
    If pFlgName = "" Then Exit Function

    '' 条件フラグが存在していること
    If IImgFlag_FlagBit(pFlgName) <> 0 Then
        Exit Function
    Else
    '' フラグが存在していない場合は、まず、ほかのプレーングループのフラグをチェックする
        For i = 1 To TheIDP.PlaneManagerCount
            If TheIDP.PlaneManager(i).GetSharedFlagPlane(pFlgName).FlagBit(pFlgName) <> 0 Then
                '' 結果のフラグビットを確保する。
                Call TheIDP.PlaneManager(m_Group).GetSharedFlagPlane(pFlgName).SetFlagBit(pFlgName)
'                Call IImgFlag_SetFlagBit(pFlgName)
                '' コピー前後のフラグプレーンを指定する。
                Set dstPlane = TheIDP.PlaneManager(m_Group).GetSharedFlagPlane(pFlgName).FlgPlane
                Set resPlane = TheIDP.PlaneManager(i).GetSharedFlagPlane(pFlgName).FlgPlane
                '' ZONE情報の保管
                resZone = resPlane.CurrentPMD.Name
                dstZone = dstPlane.CurrentPMD.Name
                
                '' ZONEのBitズレなどのチェックを入れるのであればここに記述する。
                dstBit = TheIDP.PlaneManager(m_Group).GetSharedFlagPlane(pFlgName).FlagBit(pFlgName)
                resBit = TheIDP.PlaneManager(i).GetSharedFlagPlane(pFlgName).FlagBit(pFlgName)
                '' 対象のBitのみコピーする。
                Call dstPlane.SetPMD(dstPlane.BasePMD.Name)
                Call resPlane.SetPMD(resPlane.BasePMD.Name)
                Call dstPlane.LOr(resPlane, resPlane, EEE_COLOR_FLAT, , , dstBit, resBit, resBit)
                
                '' ZONE情報をFULLからもとの状態に戻す。
                Call dstPlane.SetPMD(dstZone)
                Call resPlane.SetPMD(resZone)
                
                Exit Function
            End If
        Next i
    End If
    '' ここにくるのは、本当にどこでもフラグを生成していないケースのみ

End Function



'#FlagExpansion
Private Sub IImgFlag_LNot(ByVal pDstName As String, ByVal pSrcName As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'内容:
'   pSrcNameのビットのNot演算の結果をpDstNameのビットに入れる。
'   pDstNameが登録されていない場合、新たに登録する。
'   既に登録されている場合は、そのビットに入れる。
'
'[pSrcName]     IN String型:        データ元の名前
'[pDstName]     IN String型:        結果の名前
'[pColor]       IN IdpColorType型:  色指定
'
'備考:
'   pSrcNameとpDstNameが同一でも可能。
'   直前のSetPMDで指定されたゾーンが対象
'
    'エラー処理
    If m_FlgPlane Is Nothing Then
        Call TheError.Raise(999, "LNot", "There's no flag.")
    End If

On Error GoTo NO_FLAG
    Dim SrcFlagPlane As CImgPlane
    Set SrcFlagPlane = GetInputFlagPlane(pSrcName)
    With m_FlgPlane
        If IImgFlag_FlagBit(pDstName) = 0 Then Call IImgFlag_SetFlagBit(pDstName)
        Call .LNot(SrcFlagPlane, pColor, , GetFlgBit(pDstName), GetFlgBit(pSrcName))
    End With

    Exit Sub
    
NO_FLAG:
    Call TheError.Raise(999, "LNot", Err.Description)
End Sub

'#FlagExpansion
Private Sub IImgFlag_LOr(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'内容:
'   pSrcName1のビットとpSrcName2のビットのOr演算の結果をpDstNameのビットに入れる。
'   pDstNameが登録されていない場合、新たに登録する。
'   既に登録されている場合は、そのビットに入れる。
'
'[pSrcName1]    IN String型:        データ元の名前1
'[pSrcName2]    IN String型:        データ元の名前2
'[pDstName]     IN String型:        結果の名前
'[pColor]       IN IdpColorType型:  色指定
'
'備考:
'   pSrcName1とpSrcName2とpDstNameが同一でも可能。
'   直前のSetPMDで指定されたゾーンが対象
'
    'エラー処理
    If m_FlgPlane Is Nothing Then
        Call TheError.Raise(999, "LOr", "There's no flag.")
    End If
    
On Error GoTo NO_FLAG
    Dim SrcFlagPlane1 As CImgPlane
    Set SrcFlagPlane1 = GetInputFlagPlane(pSrcName1)
    Dim SrcFlagPlane2 As CImgPlane
    Set SrcFlagPlane2 = GetInputFlagPlane(pSrcName2)
    With m_FlgPlane
        If IImgFlag_FlagBit(pDstName) = 0 Then Call IImgFlag_SetFlagBit(pDstName)
        Call .LOr(SrcFlagPlane1, SrcFlagPlane2, pColor, , , GetFlgBit(pDstName), GetFlgBit(pSrcName1), GetFlgBit(pSrcName2))
    End With

    Exit Sub
    
NO_FLAG:
    Call TheError.Raise(999, "LOr", Err.Description)
End Sub

'#FlagExpansion
Private Sub IImgFlag_LAnd(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'内容:
'   pSrcName1のビットとpSrcName2のビットのAnd演算の結果をpDstNameのビットに入れる。
'   pDstNameが登録されていない場合、新たに登録する。
'   既に登録されている場合は、そのビットに入れる。
'
'[pSrcName1]    IN String型:        データ元の名前1
'[pSrcName2]    IN String型:        データ元の名前2
'[pDstName]     IN String型:        結果の名前
'[pColor]       IN IdpColorType型:  色指定
'
'備考:
'   pSrcName1とpSrcName2とpDstNameが同一でも可能。
'   直前のSetPMDで指定されたゾーンが対象
'
    'エラー処理
    If m_FlgPlane Is Nothing Then
        Call TheError.Raise(999, "LAnd", "There's no flag.")
    End If
    
On Error GoTo NO_FLAG
    Dim SrcFlagPlane1 As CImgPlane
    Set SrcFlagPlane1 = GetInputFlagPlane(pSrcName1)
    Dim SrcFlagPlane2 As CImgPlane
    Set SrcFlagPlane2 = GetInputFlagPlane(pSrcName2)
    With m_FlgPlane
        If IImgFlag_FlagBit(pDstName) = 0 Then Call IImgFlag_SetFlagBit(pDstName)
        Call .LAnd(SrcFlagPlane1, SrcFlagPlane2, pColor, , , GetFlgBit(pDstName), GetFlgBit(pSrcName1), GetFlgBit(pSrcName2))
    End With

    Exit Sub
    
NO_FLAG:
    Call TheError.Raise(999, "LAnd", Err.Description)
End Sub

'#FlagExpansion
Private Sub IImgFlag_LXor(ByVal pDstName As String, ByVal pSrcName1 As String, ByVal pSrcName2 As String, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'内容:
'   pSrcName1のビットとpSrcName2のビットのOr演算の結果をpDstNameのビットに入れる。
'   pDstNameが登録されていない場合、新たに登録する。
'   既に登録されている場合は、そのビットに入れる。
'
'[pSrcName1]    IN String型:        データ元の名前1
'[pSrcName2]    IN String型:        データ元の名前2
'[pDstName]     IN String型:        結果の名前
'[pColor]       IN IdpColorType型:  色指定
'
'備考:
'   pSrcName1とpSrcName2とpDstNameが同一でも可能。
'   直前のSetPMDで指定されたゾーンが対象
'
    'エラー処理
    If m_FlgPlane Is Nothing Then
        Call TheError.Raise(999, "LXor", "There's no flag.")
    End If
    
On Error GoTo NO_FLAG
    Dim SrcFlagPlane1 As CImgPlane
    Set SrcFlagPlane1 = GetInputFlagPlane(pSrcName1)
    Dim SrcFlagPlane2 As CImgPlane
    Set SrcFlagPlane2 = GetInputFlagPlane(pSrcName2)
    With m_FlgPlane
        If IImgFlag_FlagBit(pDstName) = 0 Then Call IImgFlag_SetFlagBit(pDstName)
        Call .LXor(SrcFlagPlane1, SrcFlagPlane2, pColor, , , GetFlgBit(pDstName), GetFlgBit(pSrcName1), GetFlgBit(pSrcName2))
    End With

    Exit Sub
    
NO_FLAG:
    Call TheError.Raise(999, "LXor", Err.Description)
End Sub

'#FlagExpansion
Private Function IImgFlag_SetPMD(ByVal pPmd As Variant) As IImgFlag
'内容:
'   指定したゾーンに設定
'
'[pPmd]         IN Variant型:       ゾーン指定
'
'備考:
'   pSrcName1とpSrcName2とpDstNameが同一でも可能。
'
    'プレーン確保処理
    If m_FlgPlane Is Nothing Then
        Set m_FlgPlane = TheIDP.PlaneManager(m_Group).GetFreePlane(idpDepthS16)
        m_FlgPlane.Comment = "Shared Flag Plane"
        Call m_FlgPlane.SetPMD(m_FlgPlane.BasePMD).WritePixel(0)
        Call InitializeFlagState
    End If
    
    Call m_FlgPlane.SetPMD(pPmd)
    Set IImgFlag_SetPMD = Me
End Function

'#FlagExpansion
Private Function IImgFlag_SetCustomPMD(ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long) As IImgFlag
'内容:
'   指定したゾーンに設定
'
'[pPmd]         IN Variant型:       ゾーン指定
'
'備考:
'   pSrcName1とpSrcName2とpDstNameが同一でも可能。
'
    'プレーン確保処理
    If m_FlgPlane Is Nothing Then
        Set m_FlgPlane = TheIDP.PlaneManager(m_Group).GetFreePlane(idpDepthS16)
        m_FlgPlane.Comment = "Shared Flag Plane"
        Call m_FlgPlane.SetPMD(m_FlgPlane.BasePMD).WritePixel(0)
        Call InitializeFlagState
    End If
    
    Call m_FlgPlane.SetCustomPMD(pX, pY, pWidth, pHeight)
    Set IImgFlag_SetCustomPMD = Me
    
End Function

Private Property Get IImgFlag_Name() As String
'内容:
'   フラグ用に設定されているプレーンの名前取得
    If m_FlgPlane Is Nothing Then
        IImgFlag_Name = ""
    Else
        IImgFlag_Name = m_FlgPlane.Name
    End If
End Property

Private Property Get IImgFlag_FlagBitList() As String
'内容:
'   現在設定されているフラグ番号と、付けられた名前のリストを取得

    Dim i As Long
    
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    For i = minBit To maxBit
        If m_FlgState.flagName(i) <> "" Then
            IImgFlag_FlagBitList = IImgFlag_FlagBitList & Format$(i, "##") & ":" & m_FlgState.flagName(i) & vbCrLf
        End If
    Next i
    
End Property


Private Sub IImgFlag_ResetFlagState()

    Dim i As Long
        
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    With m_FlgState
        For i = minBit To maxBit
            .FlagState(i) = Uninitialized
            .flagName(i) = ""
        Next i
    End With
    
    m_FlgState.Count = 0
    
End Sub

Private Sub InitializeFlagState()

    Dim i As Long
        
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    With m_FlgState
        For i = minBit To maxBit
            .FlagState(i) = INITIALIZED
        Next i
    End With
    
End Sub

Private Function GetInitializedFlag() As Long

    Dim i As Long
    
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
        
    With m_FlgState
        For i = minBit To maxBit
            If .FlagState(i) = INITIALIZED Then
                GetInitializedFlag = i
                Exit Function
            End If
        Next i
    End With
    
    GetInitializedFlag = 0

End Function

Private Sub RecoverFlagPlane(ByRef plane As CImgPlane)
    
    Dim i As Long
    Dim mask As Long
    
    mask = &HFFFF
    
    Dim minBit As Long, maxBit As Long
    minBit = MinIndex_
    maxBit = MaxIndex_
    
    With m_FlgState
        For i = minBit To maxBit
            If .FlagState(i) = NotUsed Then
                mask = mask - 2 ^ (i - 1)
                .FlagState(i) = INITIALIZED
            End If
        Next i
    End With
    
    Call m_FlgPlane.SetPMD(m_FlgPlane.BasePMD).WritePixel(0, , , , mask)
    
    
End Sub

Private Sub SetUsedState(ByVal pBit As Long, ByVal pName As String)
    
    On Error GoTo ErrorExit
    With m_FlgState
        .flagName(pBit) = pName
        .FlagState(pBit) = Used
        .Count = .Count + 1
    End With
    Exit Sub
    
ErrorExit:
    Call TheError.Raise(999, "SetUsedState", "Internal Error")
End Sub

Private Sub SetNotUsedState(ByVal pBit As Long)
    
    On Error GoTo ErrorExit
    With m_FlgState
        .flagName(pBit) = ""
        .FlagState(pBit) = NotUsed
        .Count = .Count - 1
    End With
    Exit Sub
    
ErrorExit:
    Call TheError.Raise(999, "SetNotUsedState", "Internal Error")
End Sub


Private Property Get FlgCount_() As Long
    FlgCount_ = m_FlgState.Count
End Property

Private Property Get MinIndex_() As Long
    MinIndex_ = LBound(m_FlgState.flagName)
End Property

Private Property Get MaxIndex_() As Long
    MaxIndex_ = UBound(m_FlgState.flagName)
End Property

Private Sub Class_Initialize()
'init

End Sub

Private Sub Class_Terminate()
'term
    Set m_FlgPlane = Nothing

End Sub

Public Function AsIImgFlag() As IImgFlag
    Set AsIImgFlag = Me
End Function

'#FlagExpansion
'フラグ名を元にフラグプレーンを取り出し、結果出力プレーン(自分)と同じPMDに設定して返す関数
Private Function GetInputFlagPlane(ByVal pFlagName As String) As CImgPlane
    'フラグプレーン取得
    Dim flag As IImgFlag
    Set flag = m_FlgPlane.Manager.GetSharedFlagPlane(pFlagName)
    'PMD設定
    If m_FlgPlane.CurrentPMD.Name = m_FlgPlane.Manager.VariablePMDName Then
        flag.SetPMD m_FlgPlane.CurrentPMD
    Else
        flag.SetPMD m_FlgPlane.CurrentPMD.Name
    End If
    Set GetInputFlagPlane = flag.FlgPlane
End Function


