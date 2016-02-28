VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   プレーンの管理
'
'目的:
'   プレーンの管理を行う。
'   サイズ、カラーマップが同じプレーンを束ねる(ビット違いは纏める)
'
'作成者:
'   0145184004
'
Option Explicit

Implements IImgPlaneManager

Private Const POSTFIX_VAR_PMD = "_VAR"

Dim m_Name As String
Dim WithEvents m_Int16 As CImgPlanes
Attribute m_Int16.VB_VarHelpID = -1
Dim WithEvents m_Int32 As CImgPlanes
Attribute m_Int32.VB_VarHelpID = -1
Dim WithEvents m_Flt32 As CImgPlanes
Attribute m_Flt32.VB_VarHelpID = -1

Dim m_PMDs As Collection
Dim m_VarPMDName As String
Dim m_PlaneMap As CImgPlaneMap

'#FlagExpansion
Dim m_FlagPlanes As CImgFlagPlanes

'#FlagExpansion
Public Sub EnableInterceptor(ByVal pEnable As Boolean)
'このプレーンマネージャが持つFlagインスタンス切り替え用
    If pEnable = True Then
        m_FlagPlanes.EnableInterceptor
    Else
        m_FlagPlanes.DisableInterceptor
    End If
End Sub

Public Function AsIImgPlaneManager() As IImgPlaneManager
    Set AsIImgPlaneManager = Me
End Function

Private Sub Class_Initialize()
    Set m_PMDs = New Collection
End Sub

'#FlagExpansion
Private Function IImgPlaneManager_ReleaseUnusedFlagPlane() As Long
    IImgPlaneManager_ReleaseUnusedFlagPlane = m_FlagPlanes.ReleaseUnusedFlagPlane
End Function

'#FlagExpansion
Private Sub IImgPlaneManager_ClearAllFlag()
    m_FlagPlanes.ClearAllFlag
End Sub

'#FlagExpansion
Private Function IImgPlaneManager_GetSharedFlagPlanes() As Collection
    Set IImgPlaneManager_GetSharedFlagPlanes = m_FlagPlanes.FlagPlanes
End Function

Private Sub IImgPlaneManager_GetPlanes(ByRef pPlanes As Collection)
    If Not m_Int16 Is Nothing Then Call m_Int16.GetPlanes(pPlanes)
    If Not m_Int32 Is Nothing Then Call m_Int32.GetPlanes(pPlanes)
    If Not m_Flt32 Is Nothing Then Call m_Flt32.GetPlanes(pPlanes)
End Sub

Private Sub IImgPlaneManager_GetNoUsePlanes(ByRef pPlanes As Collection)
    If Not m_Int16 Is Nothing Then Call m_Int16.GetNoUsePlanes(pPlanes)
    If Not m_Int32 Is Nothing Then Call m_Int32.GetNoUsePlanes(pPlanes)
    If Not m_Flt32 Is Nothing Then Call m_Flt32.GetNoUsePlanes(pPlanes)
End Sub

'#FlagExpansion
Private Sub IImgPlaneManager_CreateBasePMD(ByVal pName As String, ByVal pBasePmdName As String, ByVal pWidth As Long, ByVal pHeight As Long, ByRef pColorMap As CImgPlaneMap, ByVal pNumS16 As Long, ByVal pNumS32 As Long, ByVal pNumF32 As Long)
'内容:
'   ベースPMD作成
'
'[pName]        IN  String型:       プレーン名
'[pBasePmdName] IN  String型:       ベースPMD名
'[pWidth]       IN  Long型:         幅
'[pHeight]      IN  Long型:         高さ
'[pColorMap]    IN  String型:       カラーマップ
'[pNumS16]      IN  Long型:         16Bit整数プレーンの枚数
'[pNumS32]      IN  Long型:         32Bit整数プレーンの枚数
'[pNumF32]      IN  Long型:         32Bit浮動小数プレーンの枚数
'
'備考:
'   CImgIDPのみ呼び出せるようにしたい。
'   エラー条件
'   1,指定したベースPMDの名前が既に存在する時
'
    If TheIDP.isExistingPMD(pBasePmdName) = True Then
        'エラー
        Call TheError.Raise(999, "CreateBasePMD", ErrMsgNameAlreadyExists(pBasePmdName))
    End If
    
    Set m_Int16 = New CImgPlanes
    Set m_Int32 = New CImgPlanes
    Set m_Flt32 = New CImgPlanes
        
    Call m_Int16.CreateBasePMD(pBasePmdName, pWidth, pHeight, pName, idpDepthS16, pNumS16, pColorMap.IdpColorsPmd)
    Call m_Int32.CreateBasePMD(pBasePmdName, pWidth, pHeight, pName, idpDepthS32, pNumS32, pColorMap.IdpColorsPmd)
    Call m_Flt32.CreateBasePMD(pBasePmdName, pWidth, pHeight, pName, idpDepthF32, pNumF32, pColorMap.IdpColorsPmd)
    
    m_Name = pName
    
    Set m_FlagPlanes = New CImgFlagPlanes
    m_FlagPlanes.Initialize pName
            
    Call AddPMD(pBasePmdName, 1, 1, pWidth, pHeight)
    
    'CreateBasePMD時、TheIDPが持つCImgPlaneMapsのCImgPlaneMapのコピーを、PlaneManagerが持つ。
    'GetFreePlane時、さらにCImgPlaneがコピーしたインスタンスを持つ。
    Set m_PlaneMap = pColorMap
    Set pColorMap = Nothing
    
    m_VarPMDName = pName & POSTFIX_VAR_PMD
    Call IImgPlaneManager_CreateSubPMD(m_VarPMDName, 1, 1, 1, 1)
        
End Sub

'#FlagExpansion
Private Function IImgPlaneManager_GetSharedFlagPlane(ByVal pFlagName As String) As IImgFlag

    On Error GoTo INSUFFICIENT_PLANE

    Set IImgPlaneManager_GetSharedFlagPlane = m_FlagPlanes.FlagPlane(pFlagName)
    Exit Function

INSUFFICIENT_PLANE:
    TheError.Raise 999, "GetSharedFlagPlane", "Flag plane is nothing."
    Set IImgPlaneManager_GetSharedFlagPlane = Nothing

End Function

Private Sub IImgPlaneManager_CreateSubPMD(ByVal pSubPmdName As String, ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long)
'内容:
'   サブPMD作成
'
'[pSubPmdName] IN   String型:       サブPMD名
'[pX]          IN   Long型:         基点Xアドレス
'[pY]          IN   Long型:         基点Yアドレス
'[pWidth]      IN   Long型:         幅
'[pHeight]     IN   Long型:         高さ
'
'備考:
'   エラー条件
'   1,指定したサブPMDの名前が既に存在する時
'
'   サンプルコード
'    Call TheIDP.PlaneManager("vmcu").CreateSubPMD("ZONE2D", 80, 60, 640, 480)
'
    If TheIDP.isExistingPMD(pSubPmdName) = True Then
        'エラー
        Call TheError.Raise(999, "CreateSubPMD", ErrMsgNameAlreadyExists(pSubPmdName))
    End If
    
    Call m_Int16.CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    Call m_Int32.CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    Call m_Flt32.CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    
    Call AddPMD(pSubPmdName, pX, pY, pWidth, pHeight)

End Sub

Private Function IImgPlaneManager_GetFreePlane(ByVal pDepth As IdpBitDepth, Optional ByVal pClear As Boolean = False, Optional ByVal pComment As String = "-") As CImgPlane
'内容:
'   空きプレーンの取得
'
'[pDepth]       IN  IdpBitDepth型:  ビット指定
'[pClear]       IN  Boolean型:      クリアするかしないか
'[pComment]     IN  String型:      画像コメント
'
'備考:
'   エラー条件
'   1,用意したプレーンが全て使われている時
'
'   サンプルコード
'    Dim src As CImgPlane
'    Set src = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'
    Dim pPlane As CImgPlane
    Dim pColorMap As CImgPlaneMap
    Set pColorMap = m_PlaneMap
    
    Set pPlane = SelectPlanes(pDepth).GetFreePlane(pColorMap)
    Call pPlane.SetPMD(pPlane.BasePMD.Name)         'TOPT動作時の対策。Plane確保してからSetPMD
    pPlane.Comment = pComment
    If pClear = True Then
        Call pPlane.WritePixel(0)
    End If
    
    Set IImgPlaneManager_GetFreePlane = pPlane
    Set pPlane = Nothing
    Set pColorMap = Nothing
    
End Function

Private Property Get IImgPlaneManager_Name() As String
'内容:
'   名前の取得
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.PlaneManager("vmcu").Name
'
    IImgPlaneManager_Name = m_Name
End Property

Private Property Get IImgPlaneManager_BasePMD() As CImgPmdInfo
'内容:
'   BasePMDの情報取得
'
'備考:
'   サンプルコード
'    With TheIDP.PlaneManager("vmcu").BasePMD
'        ………
'    End With
'
    Set IImgPlaneManager_BasePMD = m_PMDs(1).Clone
End Property

Private Property Get IImgPlaneManager_PMD(ByVal pName As String) As CImgPmdInfo
'内容:
'   指定した名前のPMDの情報を取得
'
'[pName]       IN   String型:       プレーン名
'
'備考:
'

On Error GoTo NOT_FOUND
    Set IImgPlaneManager_PMD = m_PMDs(pName).Clone
    Exit Property

NOT_FOUND:
    Set IImgPlaneManager_PMD = Nothing
End Property

Private Property Get IImgPlaneManager_Count(ByVal pBitDepth As IdpBitDepth) As Long
'内容:
'   指定したビット深さのプレーンの枚数を取得
'
'[pBitDepth]    IN  IdpBitDepth型:      ビット深さ指定
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.PlaneManager("vmcu").Count(idpDepthS16)
'
    IImgPlaneManager_Count = SelectPlanes(pBitDepth).Count
End Property

Private Property Get IImgPlaneManager_UsedCount(ByVal pBitDepth As IdpBitDepth) As Long
'内容:
'   指定したビット深さのプレーンの枚数を取得
'
'[pBitDepth]    IN  IdpBitDepth型:      ビット深さ指定
'
'備考:
'   サンプルコード
'    Debug.Print TheIDP.PlaneManager("vmcu").Count(idpDepthS16)
'
    IImgPlaneManager_UsedCount = SelectPlanes(pBitDepth).UsedCount
End Property

Private Property Get IImgPlaneManager_VariablePMDName() As String
'内容:
'   可変PMDの基本名取得
'
'備考:
'
    IImgPlaneManager_VariablePMDName = m_VarPMDName
End Property

Private Property Get IImgPlaneManager_PlaneMap() As CImgPlaneMap
'内容:
'   マネージャーが持つカラーマップの情報取得
'
'備考:
'   サンプルコード
'    With TheIDP.PlaneManager("vmcu").PlaneMap
'        ………
'    End With
'
    Set IImgPlaneManager_PlaneMap = m_PlaneMap
End Property

'Private Sub IImgPlaneManager_zSetFlagInstance(instance As IImgFlag)
''このプレーンマネージャが持つFlagインスタンス切り替え用
'    Set m_FlgPlane = Nothing
'    Set m_FlgPlane = instance
'End Sub
'
Private Sub AddPMD(ByVal pName As String, ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long)
    
    'm_PMDsにPMDを追加する
    Dim newPMD As CImgPmdInfo
    
    Set newPMD = New CImgPmdInfo
    Call newPMD.Create(pName, pX, pY, pWidth, pHeight)
    Call m_PMDs.Add(newPMD, pName)
    
End Sub

Private Function SelectPlanes(ByVal pBitDepth As IdpBitDepth) As CImgPlanes

    Select Case pBitDepth
    Case idpDepthS16
        Set SelectPlanes = m_Int16
    Case idpDepthS32
        Set SelectPlanes = m_Int32
    Case idpDepthF32
        Set SelectPlanes = m_Flt32
    Case Else
        TheError.Raise 999, "SelectPlanes", """" & pBitDepth & """ is Illegal Number."
    End Select
    
End Function

Private Function ErrMsgNameAlreadyExists(ByVal pName As String)
    ErrMsgNameAlreadyExists = """" & pName & """ already exists."
End Function

Private Sub m_Flt32_EVGETMANAGER(ByRef pMng As IImgPlaneManager)
    Set pMng = Me
End Sub

Private Sub m_Int16_EVGETMANAGER(ByRef pMng As IImgPlaneManager)
    Set pMng = Me
End Sub

Private Sub m_Int32_EVGETMANAGER(ByRef pMng As IImgPlaneManager)
    Set pMng = Me
End Sub

'#FlagExpansion
Private Sub Class_Terminate()
'term

Set m_Int16 = Nothing
Set m_Int32 = Nothing
Set m_Flt32 = Nothing
Set m_PMDs = Nothing
Set m_PlaneMap = Nothing
Set m_FlagPlanes = Nothing

End Sub
