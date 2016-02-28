VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   画像プレーンを表したクラス
'
'目的:
'   画像処理関連の関数をラッピング
'
'作成者:
'   0145184004
'
Option Explicit

Implements IProcParamReplaceable

Private Const MAX_COLOR = 11

Private Const SAME_COLOR As String = "SAME_COLOR"

Private Const PARAMETER_REPLACER_PLANEGROUP As String = "PlaneGroup"
Private Const PARAMETER_REPLACER_PLANENAME As String = "PlaneName"
Private Const PARAMETER_REPLACER_BITDEPTH As String = "BitDepth"
Private Const PARAMETER_REPLACER_READONLY As String = "ReadOnly"
Private Const PARAMETER_REPLACER_CURRENTPMD As String = "CurrentPMD"
Private Const PARAMETER_REPLACER_COLORMAPNAME As String = "ColorMapName"
Private Const PARAMETER_REPLACER_PLANECOMMENT As String = "PlaneComment"

Private WithEvents m_Parent As CImgPlanes
Attribute m_Parent.VB_VarHelpID = -1

Private m_Name As String
Private m_BitDepth As IdpBitDepth
Private m_ReadOnly As Boolean

Private m_CurrentPMD As CImgPmdInfo
Private m_PlaneMap As CImgPlaneMap

Private m_Comment As String


Private Sub m_Parent_EVGETPLANES(ByVal pPlanes As Collection)
    Call pPlanes.Add(Me)
End Sub

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let Comment(ByVal RHS As String)
    m_Comment = RHS
End Property

Friend Sub Create(ByVal pName As String, ByVal pBitDepth As IdpBitDepth, ByRef pParent As CImgPlanes, ByRef pColorMap As CImgPlaneMap)
'内容:
'   プレーン名などのデータを設定
'
'[pName]       IN   String型:               プレーン名
'[pBitDepth]   IN   IdpBitDepth型:          ビット深さ
'[pParent]     IN   CImgPlaneManager型:     親マネージャ
'
'備考:
'   CImgPlaneManagerだけ呼び出せるようにしたい
'
    m_Name = pName
    m_BitDepth = pBitDepth
    m_ReadOnly = False
    Set m_Parent = pParent
    
    Set m_PlaneMap = pColorMap
    
'    Call SetPMD(BasePMD.name)         'TOPT動作時の対策でコメントアウト。Plane確保完了してからSetPMDする
    
End Sub

Private Sub Class_Terminate()
    '終了処理
    '親に通知する。
    
    '親が先に死んだ時エラーが起こるので、On Errorで回避。
    On Error Resume Next
    Set m_PlaneMap = Nothing
    Set m_CurrentPMD = Nothing
    Call m_Parent.ReleasePlane(Me)
    Set m_Parent = Nothing

End Sub

Public Property Get Name() As String
'内容:
'   プレーン名取得
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Debug.Print src.Name
'
    Name = m_Name
End Property

Public Property Get BasePMD() As CImgPmdInfo
'内容:
'   ベースPMDの情報取得
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    With src.BasePMD
'        ………
'    End With
'
    Set BasePMD = m_Parent.Manager.BasePMD
End Property

Public Property Get CurrentPMD() As CImgPmdInfo
'内容:
'   現在設定されているPMD情報を取得
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    With src.CurrentPMD
'        ………
'    End With
'
    Set CurrentPMD = m_CurrentPMD.Clone
End Property

Public Property Get CurrentPmdName() As String
'内容:
'   現在設定されているPMDの(実際の)名前を取得
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Debug.Print src.CurrentPmdName
'
    CurrentPmdName = m_CurrentPMD.Name & m_Parent.Postfix
End Property

Public Property Get ReadOnly() As Boolean
'内容:
'   Read Onlyの状態取得
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Debug.Print src.ReadOnly
'
    ReadOnly = m_ReadOnly
End Property

Public Property Let ReadOnly(ByVal RHS As Boolean)
'内容:
'   Read Onlyの状態設定
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    src.ReadOnly = True
'
    m_ReadOnly = RHS
End Property

Public Property Get planeGroup() As String
'内容:
'   プレーンのグループ名(基本のプレーン名)取得
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Debug.Print src.PlaneGroup
'
    planeGroup = m_Parent.Manager.Name
End Property

Public Property Get BitDepth() As IdpBitDepth
'内容:
'   ビット深さを取得
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Debug.Print src.BitDepth
'
    BitDepth = m_BitDepth
End Property

'EeeJOB V3.0化に伴い廃止
'Public Property Get Color(ByVal pX As Long, ByVal pY As Long) As idpColorType
''内容:
''   指定したアドレスのカラータイプ取得
''
''[pX]           IN  Long型:     Xアドレス(1スタート)
''[pY]           IN  Long型:     Yアドレス(1スタート)
''
''備考:
''   サンプルコード
''    Dim src As CImgPlane
''
''    Set src = TheIDP.PlaneBank.Item("SOURCE")
''    Debug.Print src.Color(10, 10)
''
'    If pX <= 0 Or pY <= 0 Or pX > BasePMD.Width Or pY > BasePMD.Height Then
'        Call TheError.Raise(999, "Color", "Illegal Parameter")
'        Exit Property
'    End If
'
'    With m_Parent.Manager.PlaneMap
'        Color = .Color(((pX - 1) Mod .Width) + 1, ((pY - 1) Mod .Height) + 1)
'    End With
'
'End Property
'
Public Property Get planeMap() As CImgPlaneMap
    Set planeMap = m_PlaneMap
End Property

Public Property Get Pixel(ByVal pX As Long, ByVal pY As Long, Optional ByVal pSite As Long = 0) As Double
'内容:
'   指定したアドレスの画素の値取得
'
'[pX]           IN  Long型:     Xアドレス(1スタート)
'[pY]           IN  Long型:     Yアドレス(1スタート)
'[pSite]        IN  Long型:     サイト指定
'
'備考:
'   こういうオプションの使い方はありだろうか?
'   取得はサイト0、設定は全サイト
'
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Debug.Print src.Pixel(10, 10, 0)
'
    Dim oldPmd As CImgPmdInfo
    Dim pixArray() As Double
    
    Set oldPmd = CurrentPMD
    Call SetCustomPMD(pX, pY, 1, 1)
    
    Call ReadPixel(pSite, pixArray)
    Pixel = pixArray(0)
    
    If oldPmd.Name = m_Parent.Manager.VariablePMDName Then
        Call SetPMD(oldPmd)
    Else
        Call SetPMD(oldPmd.Name)
    End If

End Property

Public Property Let Pixel(ByVal pX As Long, ByVal pY As Long, Optional ByVal pSite As Long = -1, ByVal RHS As Double)
'内容:
'   指定したアドレスの画素の値設定
'
'[pX]           IN  Long型:     Xアドレス(1スタート)
'[pY]           IN  Long型:     Yアドレス(1スタート)
'[pSite]        IN  Long型:     サイト指定
'
'備考:
'   こういうオプションの使い方はありだろうか?
'   取得はサイト0、設定は全サイト
'
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    src.Pixel(10, 10, 0) = 100
'
    Dim oldPmd As CImgPmdInfo
    
    Set oldPmd = CurrentPMD
    Call SetCustomPMD(pX, pY, 1, 1).WritePixel(RHS, , pSite)
    
    If oldPmd.Name = m_Parent.Manager.VariablePMDName Then
        Call SetPMD(oldPmd)
    Else
        Call SetPMD(oldPmd.Name)
    End If
    
End Property

Public Function SetPMD(ByVal pPmd As Variant) As CImgPlane
'内容:
'   プレーンを指定したPMDに設定する。
'
'[pPMD]        IN   Variant型:      PMD名、またはサイズを設定したCImgPmdInfo
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'
'    Dim newPMD As CImgPmdInfo
'    Set newPMD = New CImgPmdInfo
'    Call newPMD.Create("", 10, 10, 100, 100)
'    Call src.SetPMD(newPMD)
'
    If IsObject(pPmd) Then
        With pPmd
            Set SetPMD = SetCustomPMD(.XAdr, .YAdr, .width, .height)
        End With
    Else
        Set SetPMD = SetPMD_(pPmd)
    End If

End Function

Private Function SetPMD_(ByVal pPmd As String) As CImgPlane

    Dim tmpPmd As CImgPmdInfo
    Set tmpPmd = m_Parent.Manager.PMD(pPmd)
    If tmpPmd Is Nothing Then
        'エラー
        Call TheError.Raise(999, "SetPMD", ErrMsgNameDoesntExist(pPmd))
    End If
    
    Set m_CurrentPMD = tmpPmd
    Call TheHdw.IDP.SetPMD(m_Name, m_CurrentPMD.Name & m_Parent.Postfix)
    
    Set SetPMD_ = Me
End Function

Public Function SetCustomPMD(ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long) As CImgPlane
'内容:
'   プレーンを指定したサイズの一時的なPMDに設定する。
'
'[pX]          IN   Long型:         基点のX座標(1スタート)
'[pY]          IN   Long型:         基点のY座標(1スタート)
'[pWidth]      IN   Long型:         幅
'[pHeight]     IN   Long型:         高さ
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetCustomPMD(10, 10, 100, 100)
'
    Dim tmpPmd As CImgPmdInfo
    
    Set tmpPmd = New CImgPmdInfo
    Call tmpPmd.Create(m_Parent.Manager.VariablePMDName, pX, pY, pWidth, pHeight)

    On Error GoTo ERROR_IGXL
    With tmpPmd
        Call TheHdw.IDP.ModifySubPMD(.Name & m_Parent.Postfix, .XAdr, .YAdr, .width, .height)
        Call TheHdw.IDP.SetPMD(m_Name, .Name & m_Parent.Postfix)
    End With
    
    Set m_CurrentPMD = tmpPmd
    Set SetCustomPMD = Me
    
    Exit Function
    
ERROR_IGXL:
    Call ErrMsgIGXL("SetCustomPMD")
    
End Function



'##################################### 以下、画像処理用メソッド #####################################

Public Sub Average( _
    ByRef pResult() As Double, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   平均値取得
'
'[pResult()]    OUT  Double型:          結果を返す配列
'[pColor]       IN   IdpColorType型:    色指定
'[pFlgName]     IN   String型:          フラグ名
'[pMaskBit]     IN   Long型:            マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.Average(res, idpColorRed)
'
        
    On Error GoTo ERROR_IGXL
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    Call TheHdw.IDP.Accumulate(m_Name, pIdpColorType, idpAccumMean, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    If pIdpColorType = idpColorAll Then
        Call AverageAll(pResult)
    Else
        Call ReadResultFP(pIdpColorType, idpReadMean, pResult)
    End If
    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Average")
    
End Sub

Private Sub AverageAll(ByRef pResult() As Double)

    Dim tmpAvg(MAX_COLOR) As Double
    Dim tmpNum(MAX_COLOR) As Double
    Dim allNum() As Double
    
    Dim c As Long
    Dim site As Long
    
    Call Redim_(pResult)
    Call Redim_(allNum)
    For site = 0 To SiteLimit
        Call TheHdw.IDP.ReadResultFP(site, m_Name, idpColorAll, idpReadMean, tmpAvg)
        Call TheHdw.IDP.ReadResultFP(site, m_Name, idpColorAll, idpReadNum, tmpNum)
        For c = 0 To UBound(tmpAvg)
            If tmpNum(c) <> 0 Then
                pResult(site) = pResult(site) + tmpAvg(c) * tmpNum(c)
                allNum(site) = allNum(site) + tmpNum(c)
            End If
        Next c
        If allNum(site) = 0 Then
            pResult(site) = 0
        Else
            pResult(site) = pResult(site) / allNum(site)
        End If
    Next site
    
End Sub

Public Sub sum( _
    ByRef pResult() As Double, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   合計値取得
'
'[pResult()]   OUT  Double型:           結果を返す配列
'[pColor]      IN   IdpColorType型:     色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.Sum(res, idpColorRed)
    
    On Error GoTo ERROR_IGXL
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    Call TheHdw.IDP.Accumulate(m_Name, pIdpColorType, idpAccumSum, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    
    If pIdpColorType = idpColorAll Then
        Call SumAll(pResult)
    Else
        Call ReadResultFP(pIdpColorType, idpReadSum, pResult)
    End If
    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Sum")
    
End Sub

Private Sub SumAll(ByRef pResult() As Double)

    Dim tmpSum(MAX_COLOR) As Double
    Dim c As Long
    Dim site As Long
    
    Call Redim_(pResult)
    
    For site = 0 To SiteLimit
        Call TheHdw.IDP.ReadResultFP(site, m_Name, idpColorAll, idpReadSum, tmpSum)
        For c = 0 To UBound(tmpSum)
            pResult(site) = pResult(site) + tmpSum(c)
        Next c
    Next site
    
End Sub

Public Sub StdDev( _
    ByRef pResult() As Double, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   標準偏差取得
'
'[pResult()]   OUT  Double型:           結果を返す配列
'[pColor]      IN   IdpColorType型:     色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.StdDev(res, idpColorRed)
'
    
    On Error GoTo ERROR_IGXL
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    Call TheHdw.IDP.Accumulate(m_Name, pIdpColorType, idpStdDeviation, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    
    If pIdpColorType = idpColorAll Then
        Call StdDevAll(pResult)
    Else
        Call ReadResultFP(pIdpColorType, idpReadStdDeviation, pResult)
    End If
    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("StdDev")
    
End Sub

Private Sub StdDevAll(ByRef pResult() As Double)

    Dim tmpStdDev(MAX_COLOR) As Double
    Dim tmpAvg(MAX_COLOR) As Double
    Dim tmpNum(MAX_COLOR) As Double
    Dim lastAvg() As Double
    Dim lastNum() As Double
    
    Dim c As Long
    Dim site As Long
    
    Call Redim_(pResult)
    Call Redim_(lastAvg)
    Call Redim_(lastNum)
    
    With TheHdw.IDP
        For site = 0 To SiteLimit
            Call .ReadResultFP(site, m_Name, idpColorAll, idpReadStdDeviation, tmpStdDev)
            Call .ReadResultFP(site, m_Name, idpColorAll, idpReadMean, tmpAvg)
            Call .ReadResultFP(site, m_Name, idpColorAll, idpReadNum, tmpNum)
            For c = 0 To UBound(tmpStdDev)
                If tmpNum(c) <> 0 Then
                    pResult(site) = (lastNum(site) + tmpNum(c)) * ((lastNum(site) - 1) * pResult(site) + (tmpNum(c) - 1) * (tmpStdDev(c) ^ 2)) + lastNum(site) * tmpNum(c) * (lastAvg(site) - tmpAvg(c)) ^ 2
                    lastAvg(site) = DivOrZero(lastAvg(site) * lastNum(site) + tmpAvg(c) * tmpNum(c), lastNum(site) + tmpNum(c))
                    lastNum(site) = lastNum(site) + tmpNum(c)
                    pResult(site) = DivOrZero(pResult(site), lastNum(site) * (lastNum(site) - 1))
                End If
            Next c
            pResult(site) = Sqr(pResult(site))
        Next site
    End With
End Sub

Private Function DivOrZero(ByVal pVal1 As Double, ByVal pVal2 As Double) As Double
    If pVal2 = 0 Then
        DivOrZero = 0
    Else
        DivOrZero = pVal1 / pVal2
    End If
End Function

Public Sub Num( _
    ByRef pResult() As Double, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   対象の画素の個数を取得
'
'[pResult()]   OUT  Double型:           結果を返す配列
'[pColor]      IN   IdpColorType型:     色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.Num(res, idpColorRed)
'
    
    On Error GoTo ERROR_IGXL
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    Call TheHdw.IDP.Accumulate(m_Name, pIdpColorType, idpAccumSum, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    
    If pIdpColorType = idpColorAll Then
        Call NumAll(pResult)
    Else
        Call ReadResultFP(pIdpColorType, idpReadNum, pResult)
    End If
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Num")
    
End Sub

Private Sub NumAll(ByRef pResult() As Double)

    Dim tmpNum(MAX_COLOR) As Double
    Dim c As Long
    Dim site As Long
    
    Call Redim_(pResult)
    
    For site = 0 To SiteLimit
        Call TheHdw.IDP.ReadResultFP(site, m_Name, idpColorAll, idpReadNum, tmpNum)
        For c = 0 To UBound(tmpNum)
            pResult(site) = pResult(site) + tmpNum(c)
        Next c
    Next site
    
End Sub

Public Sub Min( _
    ByRef pResult() As Double, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   最小値取得
'
'[pResult()]   OUT  Double型:           結果を返す配列
'[pColor]      IN   IdpColorType型:     色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.Min(res, idpColorRed)
'
    
    On Error GoTo ERROR_IGXL
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    Call TheHdw.IDP.MinMax(m_Name, pIdpColorType, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    If pIdpColorType = idpColorAll Then
        '数を取得するためAccumulateを呼ぶ。
        Call TheHdw.IDP.Accumulate(m_Name, idpColorAll, idpAccumSum, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
        If m_BitDepth = idpDepthF32 Then
            Call MinAllFP(pResult)
        Else
            Call MinAll(pResult)
        End If
    Else
        If m_BitDepth = idpDepthF32 Then
            Call ReadResultFP(pIdpColorType, idpReadMin, pResult)
        Else
            Call ReadResult(pIdpColorType, idpReadMin, pResult)
        End If
    End If
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Min")
    
End Sub

Private Sub MinAll(ByRef pResult() As Double)

    Dim tmpMin(MAX_COLOR) As Long
    Dim tmpNum(MAX_COLOR) As Long
    Dim c As Long
    Dim site As Long
    Dim isFirst As Boolean
    
    Call Redim_(pResult)
    
    With TheHdw.IDP
        For site = 0 To SiteLimit
            Call .ReadResult(site, m_Name, idpColorAll, idpReadMin, tmpMin)
            Call .ReadResult(site, m_Name, idpColorAll, idpReadNum, tmpNum)
            isFirst = True
            For c = 0 To UBound(tmpMin)
                If tmpNum(c) <> 0 Then
                    If isFirst Then
                        pResult(site) = tmpMin(c)
                        isFirst = False
                    Else
                        If pResult(site) > tmpMin(c) Then
                            pResult(site) = tmpMin(c)
                        End If
                    End If
                End If
            Next c
        Next site
    End With
    
End Sub

Private Sub MinAllFP(ByRef pResult() As Double)

    Dim tmpMin(MAX_COLOR) As Double
    Dim tmpNum(MAX_COLOR) As Long
    Dim c As Long
    Dim site As Long
    Dim isFirst As Boolean
    
    Call Redim_(pResult)
    
    With TheHdw.IDP
        For site = 0 To SiteLimit
            Call .ReadResultFP(site, m_Name, idpColorAll, idpReadMin, tmpMin)
            Call .ReadResult(site, m_Name, idpColorAll, idpReadNum, tmpNum)
            isFirst = True
            For c = 0 To UBound(tmpMin)
                If tmpNum(c) <> 0 Then
                    If isFirst Then
                        pResult(site) = tmpMin(c)
                        isFirst = False
                    Else
                        If pResult(site) > tmpMin(c) Then
                            pResult(site) = tmpMin(c)
                        End If
                    End If
                End If
            Next c
        Next site
    End With
    
End Sub

Public Sub max( _
    ByRef pResult() As Double, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   最大値取得
'
'[pResult()]   OUT  Double型:           結果を返す配列
'[pColor]      IN   IdpColorType型:     色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.Max(res, idpColorRed)
'
        
    On Error GoTo ERROR_IGXL
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    Call TheHdw.IDP.MinMax(m_Name, pIdpColorType, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    If pIdpColorType = idpColorAll Then
        '数を取得するためAccumulateを呼ぶ。
        Call TheHdw.IDP.Accumulate(m_Name, idpColorAll, idpAccumSum, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
        If m_BitDepth = idpDepthF32 Then
            Call MaxAllFP(pResult)
        Else
            Call MaxAll(pResult)
        End If
    Else
        If m_BitDepth = idpDepthF32 Then
            Call ReadResultFP(pIdpColorType, idpReadMax, pResult)
        Else
            Call ReadResult(pIdpColorType, idpReadMax, pResult)
        End If
    End If
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Max")
    
End Sub

Private Sub MaxAll(ByRef pResult() As Double)

    Dim tmpMax(MAX_COLOR) As Long
    Dim tmpNum(MAX_COLOR) As Long
    Dim c As Long
    Dim site As Long
    Dim isFirst As Boolean
    
    Call Redim_(pResult)
    
    With TheHdw.IDP
        For site = 0 To SiteLimit
            Call .ReadResult(site, m_Name, idpColorAll, idpReadMax, tmpMax)
            Call .ReadResult(site, m_Name, idpColorAll, idpReadNum, tmpNum)
            isFirst = True
            For c = 0 To UBound(tmpMax)
                If tmpNum(c) <> 0 Then
                    If isFirst Then
                        pResult(site) = tmpMax(c)
                        isFirst = False
                    Else
                        If pResult(site) < tmpMax(c) Then
                            pResult(site) = tmpMax(c)
                        End If
                    End If
                End If
            Next c
        Next site
    End With
    
End Sub

Private Sub MaxAllFP(ByRef pResult() As Double)

    Dim tmpMax(MAX_COLOR) As Double
    Dim tmpNum(MAX_COLOR) As Long
    Dim c As Long
    Dim site As Long
    Dim isFirst As Boolean
    
    Call Redim_(pResult)
    
    With TheHdw.IDP
        For site = 0 To SiteLimit
            Call .ReadResultFP(site, m_Name, idpColorAll, idpReadMax, tmpMax)
            Call .ReadResult(site, m_Name, idpColorAll, idpReadNum, tmpNum)
            isFirst = True
            For c = 0 To UBound(tmpMax)
                If tmpNum(c) <> 0 Then
                    If isFirst Then
                        pResult(site) = tmpMax(c)
                        isFirst = False
                    Else
                        If pResult(site) < tmpMax(c) Then
                            pResult(site) = tmpMax(c)
                        End If
                    End If
                End If
            Next c
        Next site
    End With
    
End Sub

Public Sub AbsMax( _
    ByRef pResult() As Double, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   絶対値での最大値取得
'
'[pResult()]   OUT  Double型:           結果を返す配列
'[pColor]      IN   IdpColorType型:     色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'   符号はそのまま
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.AbsMax(res, idpColorRed)
'
    Dim resMin() As Double
    Dim site As Long
        
    On Error GoTo ERROR_IGXL
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    Call TheHdw.IDP.MinMax(m_Name, pIdpColorType, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    If pIdpColorType = idpColorAll Then
        '数を取得するためAccumulateを呼ぶ。
        Call TheHdw.IDP.Accumulate(m_Name, idpColorAll, idpAccumSum, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
        If m_BitDepth = idpDepthF32 Then
            Call MaxAllFP(pResult)
            Call MinAllFP(resMin)
        Else
            Call MaxAll(pResult)
            Call MinAll(resMin)
        End If
    Else
        If m_BitDepth = idpDepthF32 Then
            Call ReadResultFP(pIdpColorType, idpReadMax, pResult)
            Call ReadResultFP(pIdpColorType, idpReadMin, resMin)
        Else
            Call ReadResult(pIdpColorType, idpReadMax, pResult)
            Call ReadResult(pIdpColorType, idpReadMin, resMin)
        End If
    End If
    
    For site = 0 To SiteLimit
        If Abs(resMin(site)) > Abs(pResult(site)) Then
            pResult(site) = resMin(site)
        End If
    Next site
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("AbsMax")
    
End Sub

Public Sub DiffMinMax( _
    ByRef pResult() As Double, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   最大値と最小値の差を取得
'
'[pResult()]   OUT  Double型:           結果を返す配列
'[pColor]      IN   IdpColorType型:     色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.DiffMinMax(res, idpColorRed)
'
    Dim resMin() As Double
    Dim site As Long
        
    
    On Error GoTo ERROR_IGXL
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    Call TheHdw.IDP.MinMax(m_Name, pIdpColorType, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    If pIdpColorType = idpColorAll Then
        '数を取得するためAccumulateを呼ぶ。
        Call TheHdw.IDP.Accumulate(m_Name, idpColorAll, idpAccumSum, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
        If m_BitDepth = idpDepthF32 Then
            Call MaxAllFP(pResult)
            Call MinAllFP(resMin)
        Else
            Call MaxAll(pResult)
            Call MinAll(resMin)
        End If
    Else
        If m_BitDepth = idpDepthF32 Then
            Call ReadResultFP(pIdpColorType, idpReadMax, pResult)
            Call ReadResultFP(pIdpColorType, idpReadMin, resMin)
        Else
            Call ReadResult(pIdpColorType, idpReadMax, pResult)
            Call ReadResult(pIdpColorType, idpReadMin, resMin)
        End If
    End If
    
    For site = 0 To SiteLimit
        pResult(site) = pResult(site) - resMin(site)
    Next site
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("DiffMinMax")
    
End Sub

'=======2009/05/11 Add Maruyama ここから==============
Public Sub MinMax( _
    ByRef pResultMin() As Double, ByRef pResultMax() As Double, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   最大値と最小値を取得
'
'[pResultMin()]   OUT  Double型:        結果(最小値)を返す配列
'[pResultMax()]   OUT  Double型:        結果(最大値)を返す配列
'[pColor]      IN   IdpColorType型:     色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim resMin() As Double, resMax() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.DiffMinMax(resMin,resMax,idpColorRed)
'
    
    On Error GoTo ERROR_IGXL
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    Call TheHdw.IDP.MinMax(m_Name, pIdpColorType, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    If pIdpColorType = idpColorAll Then
        '数を取得するためAccumulateを呼ぶ。
        Call TheHdw.IDP.Accumulate(m_Name, idpColorAll, idpAccumSum, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
        If m_BitDepth = idpDepthF32 Then
            Call MaxAllFP(pResultMax)
            Call MinAllFP(pResultMin)
        Else
            Call MaxAll(pResultMax)
            Call MinAll(pResultMin)
        End If
    Else
        If m_BitDepth = idpDepthF32 Then
            Call ReadResultFP(pIdpColorType, idpReadMax, pResultMax)
            Call ReadResultFP(pIdpColorType, idpReadMin, pResultMin)
        Else
            Call ReadResult(pIdpColorType, idpReadMax, pResultMax)
            Call ReadResult(pIdpColorType, idpReadMin, pResultMin)
        End If
    End If
    
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("MinMax")

End Sub
'=======2009/05/11 Add Maruyama ここまで==============


Public Function AccumulateColumn( _
    ByRef pSrc As CImgPlane, _
    Optional ByVal pAccType As IdpAccumOption = idpAccumSum, Optional ByVal pDestRows As Long = 1, Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
) As CImgPlane
'内容:
'   列演算(縦に演算)
'
'[pSrc]        IN   CImgPlane型:        データ元のプレーン
'[pAccType]    IN   IdpAccumOption型:   演算の種類(合計、平均、etc)
'[pDestRows]   IN   Long型:             結果を入れる行(相対アドレス)
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.AccumulateColumn(src, idpAccumMean, 1)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set AccumulateColumn = Me
        Call TheError.Raise(999, "AccumulateColumn", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.AccumulateColumn(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pAccType, pMaskBit, pDestRows, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set AccumulateColumn = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("AccumulateColumn")
    
End Function

Public Function AccumulateRow( _
    ByRef pSrc As CImgPlane, _
    Optional ByVal pAccType As IdpAccumOption = idpAccumSum, Optional ByVal pDestCols As Long = 1, Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
) As CImgPlane
'内容:
'   行演算(横に演算)
'
'[pSrc]        IN   CImgPlane型:        データ元のプレーン
'[pAccType]    IN   IdpAccumOption型:   演算の種類(合計、平均、etc)
'[pDestCols]   IN   Long型:             結果を入れる列(相対アドレス)
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.AccumulateRow(src, idpAccumMean, 1)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set AccumulateRow = Me
        Call TheError.Raise(999, "AccumulateRow", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.AccumulateRow(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pAccType, pMaskBit, pDestCols, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set AccumulateRow = Me

    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("AccumulateRow")
    
End Function

Public Function Add( _
    ByVal pSrc1 As Variant, ByVal pSrc2 As Variant, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS1 As Variant = SAME_COLOR, Optional ByRef pColorS2 As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBit1 As Long, Optional ByVal pMaskBit2 As Long _
) As CImgPlane
'内容:
'   画像の加算
'
'[pSrc1]       IN   Variant型:          データ元1
'[pSrc2]       IN   Variant型:          データ元2
'[pColorD]     IN   IdpColorType型:     保存先の色指定
'[pColorS1]    IN   IdpColorType型:     データ元1の色指定
'[pColorS2]    IN   IdpColorType型:     データ元2の色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             保存先のマスクビット
'[pMaskBit1]   IN   Long型:             データ元1のマスクビット
'[pMaskBit2]   IN   Long型:             データ元2のマスクビット
'
'備考:
'   結果はこの関数の呼出元のプレーンに返る
'   pSrc1,pSrc2はプレーン、定数、サイト配列が指定可能
'
'   サンプルコード
'    Dim src1 As CImgPlane
'    Dim src2 As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src1 = TheIDP.PlaneBank.Item("SOURCE1")
'    Set src2 = TheIDP.PlaneBank.Item("SOURCE2")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src1.SetPMD("ZONE2D")
'    Call src2.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.Add(src1, src2, idpColorRed, idpColorGreen, idpColorGreen2)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set Add = Me
        Call TheError.Raise(999, "Add", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS1 As IdpColorType
    Dim pIdpColorTypeS2 As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS1 = SelectColor(pColorS1, pSrc1, pColorD)
    pIdpColorTypeS2 = SelectColor(pColorS2, pSrc2, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.Add(Var2PlaneName(pSrc1), pIdpColorTypeS1, Var2PlaneName(pSrc2), pIdpColorTypeS2, m_Name, pIdpColorTypeD, Var2Num(pSrc1), Var2Num(pSrc2), pMaskBit1, pMaskBit2, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set Add = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Add")
    
End Function

Public Function Subtract( _
    ByVal pSrc1 As Variant, ByVal pSrc2 As Variant, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS1 As Variant = SAME_COLOR, Optional ByRef pColorS2 As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBit1 As Long, Optional ByVal pMaskBit2 As Long _
) As CImgPlane
'内容:
'   画像の減算
'
'[pSrc1]       IN   Variant型:          データ元1
'[pSrc2]       IN   Variant型:          データ元2
'[pColorD]     IN   IdpColorType型:     保存先の色指定
'[pColorS1]    IN   IdpColorType型:     データ元1の色指定
'[pColorS2]    IN   IdpColorType型:     データ元2の色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             保存先のマスクビット
'[pMaskBit1]   IN   Long型:             データ元1のマスクビット
'[pMaskBit2]   IN   Long型:             データ元2のマスクビット
'
'備考:
'   結果はこの関数の呼出元のプレーンに返る
'   pSrc1,pSrc2はプレーン、定数、サイト配列が指定可能
'
'   サンプルコード
'    Dim src1 As CImgPlane
'    Dim src2 As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src1 = TheIDP.PlaneBank.Item("SOURCE1")
'    Set src2 = TheIDP.PlaneBank.Item("SOURCE2")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src1.SetPMD("ZONE2D")
'    Call src2.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.Subtract(src1, src2, idpColorRed, idpColorGreen, idpColorGreen2)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set Subtract = Me
        Call TheError.Raise(999, "Subtract", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS1 As IdpColorType
    Dim pIdpColorTypeS2 As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS1 = SelectColor(pColorS1, pSrc1, pColorD)
    pIdpColorTypeS2 = SelectColor(pColorS2, pSrc2, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.Subtract(Var2PlaneName(pSrc1), pIdpColorTypeS1, Var2PlaneName(pSrc2), pIdpColorTypeS2, m_Name, pIdpColorTypeD, Var2Num(pSrc1), Var2Num(pSrc2), pMaskBit1, pMaskBit2, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set Subtract = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Subtract")
    
End Function

Public Function Multiply( _
    ByVal pSrc1 As Variant, ByVal pSrc2 As Variant, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS1 As Variant = SAME_COLOR, Optional ByRef pColorS2 As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBit1 As Long, Optional ByVal pMaskBit2 As Long _
) As CImgPlane
'内容:
'   画像の乗算
'
'[pSrc1]       IN   Variant型:          データ元1
'[pSrc2]       IN   Variant型:          データ元2
'[pColorD]     IN   IdpColorType型:     保存先の色指定
'[pColorS1]    IN   IdpColorType型:     データ元1の色指定
'[pColorS2]    IN   IdpColorType型:     データ元2の色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             保存先のマスクビット
'[pMaskBit1]   IN   Long型:             データ元1のマスクビット
'[pMaskBit2]   IN   Long型:             データ元2のマスクビット
'
'備考:
'   結果はこの関数の呼出元のプレーンに返る
'   pSrc1,pSrc2はプレーン、定数、サイト配列が指定可能
'
'   サンプルコード
'    Dim src1 As CImgPlane
'    Dim src2 As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src1 = TheIDP.PlaneBank.Item("SOURCE1")
'    Set src2 = TheIDP.PlaneBank.Item("SOURCE2")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src1.SetPMD("ZONE2D")
'    Call src2.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.Multiply(src1, src2, idpColorRed, idpColorGreen, idpColorGreen2)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set Multiply = Me
        Call TheError.Raise(999, "Multiply", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS1 As IdpColorType
    Dim pIdpColorTypeS2 As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS1 = SelectColor(pColorS1, pSrc1, pColorD)
    pIdpColorTypeS2 = SelectColor(pColorS2, pSrc2, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.Multiply(Var2PlaneName(pSrc1), pIdpColorTypeS1, Var2PlaneName(pSrc2), pIdpColorTypeS2, m_Name, pIdpColorTypeD, Var2Num(pSrc1), Var2Num(pSrc2), pMaskBit1, pMaskBit2, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set Multiply = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Multiply")
    
End Function

Public Function Divide( _
    ByVal pSrc1 As Variant, ByVal pSrc2 As Variant, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS1 As Variant = SAME_COLOR, Optional ByRef pColorS2 As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBit1 As Long, Optional ByVal pMaskBit2 As Long _
) As CImgPlane
'内容:
'   画像の除算
'
'[pSrc1]       IN   Variant型:          データ元1
'[pSrc2]       IN   Variant型:          データ元2
'[pColorD]     IN   IdpColorType型:     保存先の色指定
'[pColorS1]    IN   IdpColorType型:     データ元1の色指定
'[pColorS2]    IN   IdpColorType型:     データ元2の色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             保存先のマスクビット
'[pMaskBit1]   IN   Long型:             データ元1のマスクビット
'[pMaskBit2]   IN   Long型:             データ元2のマスクビット
'
'備考:
'   結果はこの関数の呼出元のプレーンに返る
'   pSrc1,pSrc2はプレーン、定数、サイト配列が指定可能
'
'   サンプルコード
'    Dim src1 As CImgPlane
'    Dim src2 As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src1 = TheIDP.PlaneBank.Item("SOURCE1")
'    Set src2 = TheIDP.PlaneBank.Item("SOURCE2")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src1.SetPMD("ZONE2D")
'    Call src2.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.Divide(src1, src2, idpColorRed, idpColorGreen, idpColorGreen2)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set Divide = Me
        Call TheError.Raise(999, "Divide", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS1 As IdpColorType
    Dim pIdpColorTypeS2 As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS1 = SelectColor(pColorS1, pSrc1, pColorD)
    pIdpColorTypeS2 = SelectColor(pColorS2, pSrc2, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.Divide(Var2PlaneName(pSrc1), pIdpColorTypeS1, Var2PlaneName(pSrc2), pIdpColorTypeS2, m_Name, pIdpColorTypeD, Var2Num(pSrc1), Var2Num(pSrc2), pMaskBit1, pMaskBit2, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set Divide = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Divide")
    
End Function

Public Function AddColumns( _
    ByRef pSrc As CImgPlane, ByVal pSrcCols As Long, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBitS As Long _
) As CImgPlane
'内容:
'   列同士の加算
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pSrcCols]    IN   Long型:             列間の幅
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             マスクビット
'[pMaskBitS]   IN   Long型:             データ元のマスクビット
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.AddColumns(src, 2, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set AddColumns = Me
        Call TheError.Raise(999, "AddColumns", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.AddColumns(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pSrcCols, pMaskBitS, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set AddColumns = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("AddColumns")
    
End Function

Public Function AddRows( _
    ByRef pSrc As CImgPlane, ByVal pSrcRows As Long, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBitS As Long _
) As CImgPlane
'内容:
'   行同士の加算
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pSrcRows]    IN   Long型:             行間の幅
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             マスクビット
'[pMaskBitS]   IN   Long型:             データ元のマスクビット
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.AddRows(src, 2, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set AddRows = Me
        Call TheError.Raise(999, "AddRows", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.AddRows(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pSrcRows, pMaskBitS, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set AddRows = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("AddRows")
    
End Function

Public Function Convolution( _
    ByRef pSrc As CImgPlane, ByVal pKernel As String, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR _
) As CImgPlane
'内容:
'   コンボリューションフィルタ
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pKernel]     IN   String型:           フィルタ名
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.Convolution(src, "HPF", idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set Convolution = Me
        Call TheError.Raise(999, "Convolution", ErrMsgReadOnly)
        Exit Function
    End If
    
    If pSrc Is Me Then
        Call TheError.Raise(999, "Convolution", "pSrc is same plane.")
        Exit Function
    End If
    
    If (m_BitDepth = idpDepthF32) And (TheIDP.Kernel(pKernel).KernelType = idpKernelInteger) Then
        Call TheError.Raise(999, "Convolution", "Invalid Kernel Type")
        Exit Function
    ElseIf ((m_BitDepth = idpDepthS16) Or (m_BitDepth = idpDepthS32)) And (TheIDP.Kernel(pKernel).KernelType = idpKernelFloat) Then
        Call TheError.Raise(999, "Convolution", "Invalid Kernel Type")
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.Convolution(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pKernel)
    Set Convolution = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Convolution")
    
End Function

Public Function CopyAtIntervals( _
    ByRef pSrc As CImgPlane, ByVal pIncMode As IdpIncMode, ByVal pPixInc As Variant, ByVal pLineInc As Variant, _
    Optional ByVal pPixAccess As IdpPixAccess = idpPixReadFirst _
) As CImgPlane
'内容:
'   間隔を指定したコピー
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pIncMode]    IN   IdpIncMode型:       モード指定
'[pPixInc]     IN   Variant型:          横方向間隔指定
'[pLineInc]    IN   Variant型:          縦方向間隔指定
'[pPixAccess]  IN   IdpPixAccess型:     アクセス方法指定
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.CopyAtIntervals(src, idpIncFix, 2, 2, idpPixReadFirst)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set CopyAtIntervals = Me
        Call TheError.Raise(999, "CopyAtIntervals", ErrMsgReadOnly)
        Exit Function
    End If
    
    '色指定はidpColorFlatのみ(IPの仕様)
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.CopyAtIntervals(pSrc.Name, idpColorFlat, m_Name, idpColorFlat, pIncMode, pPixInc, pLineInc, pPixAccess)
    Set CopyAtIntervals = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("CopyAtIntervals")
    
End Function

Public Function CopyPlane( _
    ByRef pSrc As CImgPlane, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, Optional ByVal pCopyMode As IdpCopyMode = idpCopyNormal, _
    Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBitS As Long _
) As CImgPlane
'内容:
'   コピー
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'[pCopyMode]   IN   IdpCopyMode型:      コピーのモード
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             保存先のマスクビット
'[pMaskBitS]   IN   Long型:             データ元のマスクビット
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.CopyPlane(src, idpColorRed, idpColorGreen)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set CopyPlane = Me
        Call TheError.Raise(999, "CopyPlane", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.CopyPlane(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pMaskBitS, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName), pCopyMode)
    Set CopyPlane = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("CopyPlane")
    
End Function

'#FlagExpansion
Public Sub Count( _
    ByRef pResult() As Double, ByVal pCountType As IdpCountType, ByVal pLoLim As Variant, ByVal pHiLim As Variant, _
    Optional ByVal pLimitType As IdpLimitType = idpLimitExclude, Optional ByRef pColor As Variant = EEE_COLOR_FLAT, _
    Optional ByVal pFlgName As String = "", Optional ByVal pInputFlgName As String = "", Optional ByVal pMaskBit As Long _
)
'内容:
'   指定条件に合う画素の個数をカウントする。
'
'[pResult()]        OUT Double型:           結果を返す配列
'[pCountType]       IN  IdpCountType型:     カウント方法の指定
'[pLoLim]           IN  Variant型:          Low側のリミット(サイト配列も可)
'[pHiLim]           IN  Variant型:          High側のリミット(サイト配列も可)
'[pLimitType]       IN  IdpLimitType型:     リミット値を含む、含まないの指定
'[pColor]           IN  IdpColorType型:     色指定
'[pFlgName]         IN  String型:           フラグ名
'[pInputFlgName]    IN  String型:           入力フラグ名
'[pMaskBit]         IN  Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合、共用フラグプレーンのカウント対象の位置に1を立てる。
'   入力フラグ名を指定した場合、1の立っている箇所がカウントの対象になる。
'   パラメータはTheHdw.IDP.Countに準拠
'   実際使っているのはCountMultiLimit
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.Count(res, idpCountBetween, 0, 100, idpLimitInclude, idpColorRed)
'
'=======2009/04/28 追加 Maruyama ここから==============
    If Not IsValidLimitArray(pLoLim) Then
        Exit Sub
    End If
    If Not IsValidLimitArray(pHiLim) Then
        Exit Sub
    End If
'=======2009/04/28 Add Maruyama ここまで==============
    
    Dim loLimArr() As Double
    Dim hiLimArr() As Double
    
    Call MakeLimitArray(pLoLim, loLimArr)
    Call MakeLimitArray(pHiLim, hiLimArr)
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    On Error GoTo ERROR_IGXL
    
    If pFlgName <> "" Then
        Call GetSharedFlagPlane(pFlgName).SetFlagBit(pFlgName)
    End If
    Call TheHdw.IDP.CountMultiLimit(m_Name, pIdpColorType, pCountType, loLimArr, hiLimArr, idpLimitEachSite, pLimitType, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName), GetOptPlaneName(pInputFlgName), GetOptFlgBit(pInputFlgName))
    
    If pIdpColorType = idpColorAll Then
        Call CountAll(pResult)
    Else
        Call ReadResult(pIdpColorType, idpReadCount, pResult)
    End If
    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Count")
    
End Sub

'=======2009/05/19 Add Maruyama ここから==============
Public Sub CountForFlgBitImgPlane( _
    ByRef pResult() As Double, ByVal pCountType As IdpCountType, ByVal pLoLim As Variant, ByVal pHiLim As Variant, _
    ByRef pFlgPlane As CImgPlane, ByVal pFlgBit As Long, _
    Optional ByVal pLimitType As IdpLimitType = idpLimitExclude, Optional ByRef pColor As Variant = EEE_COLOR_FLAT, _
    Optional ByVal pInputFlgName As String = "", Optional ByVal pMaskBit As Long _
)
'内容:
'   指定条件に合う画素の個数をカウントする。(フラグビットをイメージプレンに立てる)
'
'[pResult()]        OUT Double型:           結果を返す配列
'[pCountType]       IN  IdpCountType型:     カウント方法の指定
'[pLoLim]           IN  Variant型:          Low側のリミット(サイト配列も可)
'[pHiLim]           IN  Variant型:          High側のリミット(サイト配列も可)
'[pFlgName]         IN  CImgPlane型:        フラグを出力するプレン
'[pFlgBit]          IN  pFlgBit型:          フラグビット
'[pLimitType]       IN  IdpLimitType型:     リミット値を含む、含まないの指定
'[pColor]           IN  IdpColorType型:     色指定
'[pInputFlgName]    IN  String型:           入力フラグ名
'[pMaskBit]         IN  Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合、共用フラグプレーンのカウント対象の位置に1を立てる。
'   入力フラグ名を指定した場合、1の立っている箇所がカウントの対象になる。
'   パラメータはTheHdw.IDP.Countに準拠
'   実際使っているのはCountMultiLimit
'
'   サンプルコード
'    Dim src As CImgPlane, flg As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Set flg = TheIDP.TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16, True)
'    Call src.SetPMD("ZONE2D")
'    Call src.CountForFlgBitImgPlane(res, idpCountBetween, 0, 100, flg, 1, idpLimitInclude, idpColorRed)
'   =>関数をまたがない連続点とか
'
    If Not IsValidLimitArray(pLoLim) Then
        Exit Sub
    End If
    If Not IsValidLimitArray(pHiLim) Then
        Exit Sub
    End If
    
    Dim loLimArr() As Double
    Dim hiLimArr() As Double
    
    Call MakeLimitArray(pLoLim, loLimArr)
    Call MakeLimitArray(pHiLim, hiLimArr)
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    On Error GoTo ERROR_IGXL
    
'    If pFlgName <> "" Then
'        Call GetSharedFlagPlane.SetFlagBit(pFlgName)
'    End If
    Call TheHdw.IDP.CountMultiLimit(m_Name, pIdpColorType, pCountType, loLimArr, hiLimArr, idpLimitEachSite, pLimitType, pMaskBit, pFlgPlane.Name, pFlgBit, GetOptPlaneName(pInputFlgName), GetOptFlgBit(pInputFlgName))
    
    If pIdpColorType = idpColorAll Then
        Call CountAll(pResult)
    Else
        Call ReadResult(pIdpColorType, idpReadCount, pResult)
    End If
    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("CountForFlgBitImgPlane")
    
End Sub
'=======2009/05/19 Add Maruyama ここまで==============

Private Sub CountAll(ByRef pResult() As Double)

    Dim tmpCount(MAX_COLOR) As Long
    Dim c As Long
    Dim site As Long
    
    Call Redim_(pResult)
    
    For site = 0 To SiteLimit
        Call TheHdw.IDP.ReadResult(site, m_Name, idpColorAll, idpReadCount, tmpCount)
        For c = 0 To UBound(tmpCount)
            pResult(site) = pResult(site) + tmpCount(c)
        Next c
    Next site
    
End Sub

Public Sub PutFlag( _
    ByVal pFlgName As String, ByVal pCountType As IdpCountType, ByVal pLoLim As Variant, ByVal pHiLim As Variant, _
    Optional ByVal pLimitType As IdpLimitType = idpLimitExclude, Optional ByRef pColor As Variant = EEE_COLOR_FLAT, _
    Optional ByVal pInputFlgName As String = "", Optional ByVal pMaskBit As Long _
)
'内容:
'   指定条件に合う画素の位置にフラグを立てる。
'
'[pFlgName]         IN  String型:           フラグ名
'[pCountType]       IN  IdpCountType型:     カウント方法の指定
'[pLoLim]           IN  Variant型:          Low側のリミット(サイト配列も可)
'[pHiLim]           IN  Variant型:          High側のリミット(サイト配列も可)
'[pLimitType]       IN  IdpLimitType型:     リミット値を含む、含まないの指定
'[pColor]           IN  IdpColorType型:     色指定
'[pInputFlgName]    IN  String型:           入力フラグ名
'[pMaskBit]         IN  Long型:             マスクビット
'
'備考:
'   入力フラグ名を指定した場合、1の立っている箇所がカウントの対象になる。
'   パラメータはTheHdw.IDP.Countに準拠
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.PutFlag(src, idpCountBetween, 0, 100, idpLimitInclude, idpColorRed, 1)
'
    
    Dim tmp() As Double
    
    Call Count(tmp, pCountType, pLoLim, pHiLim, pLimitType, pColor, pFlgName, pInputFlgName, pMaskBit)
    
End Sub

Public Function ExecuteLUT(ByRef pSrc As CImgPlane, ByVal pLUTName As String, Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR) As CImgPlane
'内容:
'   LUTの実行
'
'[pSrc]        IN   CImgPlane型:        データ元のプレーン
'[pLutName]    IN   String型:           LUTの名前
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.ExecuteLUT(src, "ABS", idpColorAll, idpColorAll)
'
    
    If m_ReadOnly = True Then
        'エラー処理
        Set ExecuteLUT = Me
        Call TheError.Raise(999, "ExecuteLUT", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.ExecuteLUT(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pLUTName)
    Set ExecuteLUT = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("ExecuteLUT")
    
End Function

Public Sub Labeling(ByRef pResult() As Double, _
    Optional pMinPixels As Long = 2, Optional ByRef pDstPlane As CImgPlane, Optional ByVal pLabelOutputMode As IdpLabelOutputMode = idpLabelNumberMode _
)
'内容:
'   ラベリング処理
'   値が0でない画素をグループ化し、最大グループの画素数を返す。
'
'[pResult()]        OUT Double型:               結果を返す配列
'[pMinPixels]       IN  Long型:                 連続と見做す画素の最小画素数
'[pDstPlane]        OUT CImgPlane型:            連続していると判断した画素の位置情報を書き込むプレーン
'[pLabelOutputMode] IN  IdpLabelOutputMode型:   位置情報の出力方法指定
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.Labeling(res, 2)
'
    If Not pDstPlane Is Nothing Then
        If pDstPlane.ReadOnly = True Then
            'エラー処理
            Call TheError.Raise(999, "Labeling", ErrMsgReadOnly)
            Exit Sub
        End If
    End If
    
    Dim tmp As Long
    Dim site As Long
    
    On Error GoTo ERROR_IGXL
    Call Redim_(pResult)
    For site = 0 To SiteLimit
        Call TheHdw.IDP.LabelingEx(site, m_Name, tmp, idpNeighborhood8, pMinPixels, OptPlaneName(pDstPlane), pLabelOutputMode)
        pResult(site) = tmp
    Next site
    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Labeling")
    
End Sub

Public Function LAnd( _
    ByRef pSrc1 As CImgPlane, ByVal pSrc2 As Variant, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS1 As Variant = SAME_COLOR, Optional ByRef pColorS2 As Variant = SAME_COLOR, _
    Optional ByVal pBitD As Long, Optional ByVal pBitS1 As Long, Optional ByVal pBitS2 As Long, _
    Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBitS1 As Long, Optional ByVal pMaskBitS2 As Long _
) As CImgPlane
'内容:
'   対象画像同士の論理積
'
'[pSrc1]       IN   CImgPlane型:        データ元1
'[pSrc2]       IN   Variant型:          データ元2
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS1]    IN   IdpColorType型:     データ元1の色指定
'[pColorS2]    IN   IdpColorType型:     データ元2の色指定
'[pBitD]       IN   Long型:             保存先のビット指定
'[pBitS1]      IN   Long型:             データ元1のビット指定
'[pBitS2]      IN   Long型:             データ元2のビット指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             保存先のマスクビット
'[pMaskBitS1]  IN   Long型:             データ元1のマスクビット
'[pMaskBitS2]  IN   Long型:             データ元2のマスクビット
'
'備考:
'   サンプルコード
'    Dim src1 As CImgPlane
'    Dim src2 As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src1 = TheIDP.PlaneBank.Item("SOURCE1")
'    Set src2 = TheIDP.PlaneBank.Item("SOURCE2")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src1.SetPMD("ZONE2D")
'    Call src2.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.LAnd(src1, src2, idpColorRed, idpColorGreen, idpColorGreen2, 1, 2, 3)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set LAnd = Me
        Call TheError.Raise(999, "LAnd", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS1 As IdpColorType
    Dim pIdpColorTypeS2 As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS1 = SelectColor(pColorS1, pSrc1, pColorD)
    pIdpColorTypeS2 = SelectColor(pColorS2, pSrc2, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.LAnd(pSrc1.Name, pIdpColorTypeS1, Var2PlaneName(pSrc2), pIdpColorTypeS2, m_Name, pIdpColorTypeD, pBitS1, pBitS2, pBitD, Var2Num(pSrc2), pMaskBitS1, pMaskBitS2, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set LAnd = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("LAnd")
    
End Function

Public Function LNot( _
    ByRef pSrc As CImgPlane, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByVal pBitD As Long, Optional ByVal pBitS As Long, _
    Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBitS As Long _
) As CImgPlane
'内容:
'   対象画像の論理否定
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'[pBitD]       IN   Long型:             保存先のビット指定
'[pBitS]       IN   Long型:             データ元のビット指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             保存先のマスクビット
'[pMaskBitS]   IN   Long型:             データ元のマスクビット
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.LNot(src, idpColorRed, idpColorGreen, 1, 2)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set LNot = Me
        Call TheError.Raise(999, "LNot", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.LNot(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pBitS, pBitD, pMaskBitS, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set LNot = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("LNot")
    
End Function

Public Function LOr( _
    ByRef pSrc1 As CImgPlane, ByVal pSrc2 As Variant, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS1 As Variant = SAME_COLOR, Optional ByRef pColorS2 As Variant = SAME_COLOR, _
    Optional ByVal pBitD As Long, Optional ByVal pBitS1 As Long, Optional ByVal pBitS2 As Long, _
    Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBitS1 As Long, Optional ByVal pMaskBitS2 As Long _
) As CImgPlane
'内容:
'   対象画像同士の論理和
'
'[pSrc1]       IN   CImgPlane型:        データ元1
'[pSrc2]       IN   Variant型:          データ元2
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS1]    IN   IdpColorType型:     データ元1の色指定
'[pColorS2]    IN   IdpColorType型:     データ元2の色指定
'[pBitD]       IN   Long型:             保存先のビット指定
'[pBitS1]      IN   Long型:             データ元1のビット指定
'[pBitS2]      IN   Long型:             データ元2のビット指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             保存先のマスクビット
'[pMaskBitS1]  IN   Long型:             データ元1のマスクビット
'[pMaskBitS2]  IN   Long型:             データ元2のマスクビット
'
'備考:
'   サンプルコード
'    Dim src1 As CImgPlane
'    Dim src2 As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src1 = TheIDP.PlaneBank.Item("SOURCE1")
'    Set src2 = TheIDP.PlaneBank.Item("SOURCE2")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src1.SetPMD("ZONE2D")
'    Call src2.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.LOr(src1, src2, idpColorRed, idpColorGreen, idpColorGreen2, 1, 2, 3)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set LOr = Me
        Call TheError.Raise(999, "LOr", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS1 As IdpColorType
    Dim pIdpColorTypeS2 As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS1 = SelectColor(pColorS1, pSrc1, pColorD)
    pIdpColorTypeS2 = SelectColor(pColorS2, pSrc2, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.LOr(pSrc1.Name, pIdpColorTypeS1, Var2PlaneName(pSrc2), pIdpColorTypeS2, m_Name, pIdpColorTypeD, pBitS1, pBitS2, pBitD, Var2Num(pSrc2), pMaskBitS1, pMaskBitS2, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set LOr = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("LOr")
    
End Function

Public Function LXor( _
    ByRef pSrc1 As CImgPlane, ByVal pSrc2 As Variant, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS1 As Variant = SAME_COLOR, Optional ByRef pColorS2 As Variant = SAME_COLOR, _
    Optional ByVal pBitD As Long, Optional ByVal pBitS1 As Long, Optional ByVal pBitS2 As Long, _
    Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBitS1 As Long, Optional ByVal pMaskBitS2 As Long _
) As CImgPlane
'内容:
'   対象画像同士の排他的論理和
'
'[pSrc1]       IN   CImgPlane型:        データ元1
'[pSrc2]       IN   Variant型:          データ元2
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS1]    IN   IdpColorType型:     データ元1の色指定
'[pColorS2]    IN   IdpColorType型:     データ元2の色指定
'[pBitD]       IN   Long型:             保存先のビット指定
'[pBitS1]      IN   Long型:             データ元1のビット指定
'[pBitS2]      IN   Long型:             データ元2のビット指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             保存先のマスクビット
'[pMaskBitS1]  IN   Long型:             データ元1のマスクビット
'[pMaskBitS2]  IN   Long型:             データ元2のマスクビット
'
'備考:
'   サンプルコード
'    Dim src1 As CImgPlane
'    Dim src2 As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src1 = TheIDP.PlaneBank.Item("SOURCE1")
'    Set src2 = TheIDP.PlaneBank.Item("SOURCE2")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src1.SetPMD("ZONE2D")
'    Call src2.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.LXor(src1, src2, idpColorRed, idpColorGreen, idpColorGreen2, 1, 2, 3)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set LXor = Me
        Call TheError.Raise(999, "LXor", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS1 As IdpColorType
    Dim pIdpColorTypeS2 As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS1 = SelectColor(pColorS1, pSrc1, pColorD)
    pIdpColorTypeS2 = SelectColor(pColorS2, pSrc2, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.LXor(pSrc1.Name, pIdpColorTypeS1, Var2PlaneName(pSrc2), pIdpColorTypeS2, m_Name, pIdpColorTypeD, pBitS1, pBitS2, pBitD, Var2Num(pSrc2), pMaskBitS1, pMaskBitS2, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set LXor = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("LXor")
    
End Function

Public Function MultiMean( _
    ByRef pSrc As CImgPlane, ByVal pDivX As Long, ByVal pDivY As Long, _
    Optional ByVal pMultiMeanfunc As IdpMultiMeanFunc = idpMultiMeanFuncSum, Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long _
) As CImgPlane
'内容:
'   マルチミーン
'
'[pSrc]             IN  CImgPlane型:        データ元
'[pDivX]            IN  Long型:             横方向の圧縮画素数
'[pDivY]            IN  Long型:             縦方向の圧縮画素数
'[pMultiMeanfunc]   IN  IdpMultiMeanFunc型: 演算方法
'[pColorD]          IN  IdpColorType型:     保存先の色指定
'[pColorS]          IN  IdpColorType型:     データ元の色指定
'[pFlgName]         IN  String型:           フラグ名
'[pMaskBit]         IN  Long型:             マスクビット
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.MultiMean(src, 2, 2, idpMultiMeanFuncMean, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set MultiMean = Me
        Call TheError.Raise(999, "MultiMean", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.MultiMean(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, idpMultiMeanModeColor, pMultiMeanfunc, pDivX, pDivY, idpMultiMeanDivPixel, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set MultiMean = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("MultiMean")
    
End Function

'=======2009/04/28 追加 Maruyama ここから==============
Public Function MultiMeanByBlock( _
    ByRef pSrc As CImgPlane, ByVal pDivX As Long, ByVal pDivY As Long, _
    Optional ByVal pMultiMeanfunc As IdpMultiMeanFunc = idpMultiMeanFuncSum, Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long _
) As CImgPlane
'内容:
'   マルチミーン
'
'[pSrc]             IN  CImgPlane型:        データ元
'[pDivX]            IN  Long型:             横方向の分割数
'[pDivY]            IN  Long型:             縦方向の分割数
'[pMultiMeanfunc]   IN  IdpMultiMeanFunc型: 演算方法
'[pColorD]          IN  IdpColorType型:     保存先の色指定
'[pColorS]          IN  IdpColorType型:     データ元の色指定
'[pFlgName]         IN  String型:           フラグ名
'[pMaskBit]         IN  Long型:             マスクビット
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.MultiMeanByBlock(src, 2, 2, idpMultiMeanFuncMean, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set MultiMeanByBlock = Me
        Call TheError.Raise(999, "MultiMeanByBlock", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.MultiMeanByBlock(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pDivX, pDivY, pMultiMeanfunc, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set MultiMeanByBlock = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("MultiMeanByBlock")
    
End Function
'=======2009/04/28 Add Maruyama ここまで==============

Public Function NonContinuousRankFilter( _
    ByRef pSrc As CImgPlane, ByVal pWindowWidth As Long, ByVal pWindowHeight As Long, ByVal pRank As Long, _
    ByVal pColInc As Long, ByVal pRowInc As Long, Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR _
) As CImgPlane
'内容:
'   連続しない画素に対するランクフィルタ処理
'
'[pSrc]             IN  CImgPlane型:        データ元
'[pWindowWidth]     IN  Long型:             フィルタ幅
'[pWindowHeight]    IN  Long型:             フィルタ高さ
'[pRank]            IN  Long型:             ランク
'[pColInc]          IN  Long型:             横方向の増分
'[pRowInc]          IN  Long型:             縦方向の増分
'[pColorD]          IN  IdpColorType型:     保存先の色指定
'[pColorS]          IN  IdpColorType型:     データ元の色指定
'
'備考:
'   2枚以上の空きプレーンが必要
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.NonContinuousRankFilter(src, 3, 3, 5, 2, 2, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set NonContinuousRankFilter = Me
        Call TheError.Raise(999, "NonContinuousRankFilter", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim wrk1 As CImgPlane
    Dim wrk2 As CImgPlane
    
    Set wrk1 = TheIDP.PlaneManager(pSrc.planeGroup).GetFreePlane(pSrc.BitDepth).SetPMD(pSrc.CurrentPMD.Name)
    Set wrk2 = TheIDP.PlaneManager(pSrc.planeGroup).GetFreePlane(pSrc.BitDepth).SetPMD(pSrc.CurrentPMD.Name)
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.NonContinuousRankFilter(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pWindowWidth, pRank, wrk1.Name, wrk2.Name, pColInc, pRowInc, pWindowHeight)
    Set NonContinuousRankFilter = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("NonContinuousRankFilter")
    
End Function

'#FlagExpansion
Public Sub PixelLog( _
    ByVal pSite As Long, ByRef pFlgName As String, ByRef pResultArray() As T_PIXINFO, _
    Optional ByVal pMaxNum As Long = 0, Optional ByVal pAddrMode As IdpAddrMode = idpAddrAbsolute _
)
'内容:
'   フラグプレーンのフラグに対応する位置の画素情報を取得
'
'[pSite]            IN  Long型:         サイト指定
'[pFlgName]         IN  String型:       フラグ名
'[pResultArray()]   OUT T_PIXINFO型:    データの格納先
'[pMaxNum]          IN  Long型:         データ数
'[pAddrMode]        IN  IdpAddrMode型:  アドレスの出力方法
'
'備考:
'   pMaxNumに0以下を指定すると、フラグ数を取得してその数だけ結果を返す。
'   もしフラグが立っていない場合は、(-1 To -1)の配列を返す。
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim flg As CImgPlane
'    Dim res() As T_PIXINFO
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set flg = TheIDP.PlaneBank.Item("FLAG")
'    Call src.SetPMD("ZONE2D")
'    Call flg.SetPMD("ZONE2D")
'    Call src.PixelLog(1, flg, res, 100, 1, idpAddrAbsolute)
'

    Dim dblTmp() As Double
    Dim lngTmp() As Long
    Dim i As Long
    Dim j As Long
    Dim tmpNum() As Double

    On Error GoTo ERROR_IGXL
    
    '指定が無い場合はカウントして最大取得個数にする。
    If pMaxNum <= 0 Then
        Call Num(tmpNum, idpColorFlat, pFlgName)
        pMaxNum = tmpNum(pSite)
        If pMaxNum = 0 Then
            ReDim pResultArray(-1 To -1)
            Exit Sub
        End If
    End If
    
    Dim FlagPlane As IImgFlag
    Set FlagPlane = GetSharedFlagPlane(pFlgName)
    If CurrentPMD.Name = Me.Manager.VariablePMDName Then
        FlagPlane.SetPMD CurrentPMD
    Else
        FlagPlane.SetPMD CurrentPMD.Name
    End If

    If m_BitDepth = idpDepthF32 Then
        ReDim dblTmp(pMaxNum, 2)
        Call TheHdw.IDP.PixelLog(pSite, m_Name, idpColorFlat, dblTmp, pMaxNum, GetSharedFlagPlane(pFlgName).Name, GetSharedFlagPlane(pFlgName).FlagBit(pFlgName), pAddrMode)

        If dblTmp(0, 0) = 0 Or dblTmp(0, 1) = 0 Then
            'X=0 or Y=0　=>　欠陥最大個数の指定はあったが、欠陥が一個も無い
            ReDim pResultArray(-1 To -1)
            Exit Sub
        Else
            ReDim pResultArray(pMaxNum - 1)
            For i = 0 To pMaxNum - 1
                pResultArray(i).x = dblTmp(i, 0)
                pResultArray(i).y = dblTmp(i, 1)
                pResultArray(i).Value = dblTmp(i, 2)
            Next i
        End If
    Else
        ReDim lngTmp(pMaxNum, 2)
        Call TheHdw.IDP.PixelLog(pSite, m_Name, idpColorFlat, lngTmp, pMaxNum, GetSharedFlagPlane(pFlgName).Name, GetSharedFlagPlane(pFlgName).FlagBit(pFlgName), pAddrMode)

        
        If lngTmp(0, 0) = 0 Or lngTmp(0, 1) = 0 Then
            'X=0 or Y=0　=>　欠陥最大個数の指定はあったが、欠陥が一個も無い
            ReDim pResultArray(-1 To -1)
            Exit Sub
        Else
            ReDim pResultArray(pMaxNum - 1)
            For i = 0 To pMaxNum - 1
                pResultArray(i).x = lngTmp(i, 0)
                pResultArray(i).y = lngTmp(i, 1)
                pResultArray(i).Value = lngTmp(i, 2)
            Next i
        End If
    End If

    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("PixelLog")
    
End Sub

Public Function RankFilter( _
    ByVal pSrc As CImgPlane, ByVal pWindowWidth As Long, ByVal pWindowHeight As Long, ByVal pRank As Long, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR _
) As CImgPlane
'内容:
'   ランクフィルタ処理
'
'[pSrc]             IN  CImgPlane型:        データ元
'[pWindowWidth]     IN  Long型:             フィルタ幅
'[pWindowHeight]    IN  Long型:             フィルタ高さ
'[pRank]            IN  Long型:             ランク値
'[pColorD]          IN  IdpColorType型:     保存先の色指定
'[pColorS]          IN  IdpColorType型:     データ元の色指定
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.RankFilter(src, 3, 3, 5, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set RankFilter = Me
        Call TheError.Raise(999, "RankFilter", ErrMsgReadOnly)
        Exit Function
    End If
    
    If pSrc Is Me Then
        Call TheError.Raise(999, "RankFilter", "pSrc is same plane.")
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.RankFilter(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pWindowWidth, pRank, pWindowHeight)
    Set RankFilter = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("RankFilter")
    
End Function

Public Function Median( _
    ByVal pSrc As CImgPlane, ByVal pWidth As Long, ByVal pHeight As Long, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR _
) As CImgPlane
'内容:
'   メディアンフィルタ処理
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pWidth]      IN   Long型:             フィルタ幅
'[pHeight]     IN   Long型:             フィルタ高さ
'[pColorD]     IN   IdpColorType型:     保存先の色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.Median(src, 3, 3, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set Median = Me
        Call TheError.Raise(999, "Median", ErrMsgReadOnly)
        Exit Function
    End If
    
    If pSrc Is Me Then
        Call TheError.Raise(999, "Median", "pSrc is same plane.")
        Exit Function
    End If
    
    Dim rank As Long
    If (pWidth Mod 2 = 0) Or (pHeight Mod 2 = 0) Then
        Call TheError.Raise(999, "Median", "Specify the odd number for pWidth and pHeight")
    End If
    rank = (pWidth * pHeight + 1) / 2
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.RankFilter(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pWidth, rank, pHeight)
    Set Median = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Median")
    
End Function

Public Function ReadFile(ByVal pSite As Long, ByVal pFileName As String, Optional ByVal pFileType As IdpFileFormat = idpFileBinary, Optional ByRef pColor As Variant = EEE_COLOR_FLAT) As CImgPlane
'内容:
'   ファイルから画像データの読み込み
'
'[pSite]       IN   Long型:             サイト指定
'[pFileName]   IN   String型:           ファイル名
'[pFileType]   IN   IdpFileFormat型:    データのタイプ
'[pColor]      IN   IdpColorType型:     色指定
'
'備考:
'   エラーを発行しない?
'
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.SetPMD("ZONE2D")
'    Call dst.ReadFile(0, "File", idpFileBinary)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set ReadFile = Me
        Call TheError.Raise(999, "ReadFile", ErrMsgReadOnly)
        Exit Function
    End If
    
    Call ErrMsgColorAll(pColor, "ReadFile")
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.ReadFile(pSite, m_Name, pIdpColorType, pFileName, pFileType)
    Set ReadFile = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("ReadFile")
    
End Function

Public Sub ReadPixel(ByVal pSite As Long, ByRef pResultArr() As Double, Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByVal pMaxPixel As Long = 0)
'内容:
'   現在指定されているPMDの画素情報を取得する。
'
'[pSite]       IN   Long型:             サイト指定
'[pResultArr]  OUT  Double型:           データの出力先
'[pColor]      IN   IdpColorType型:     色指定
'[pMaxPixel]   IN   Long型:             最大のデータ数
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res() As Double
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.ReadPixel(0, res, idpColorRed, 10)
'
    Dim lngTmp() As Long
    Dim i As Long
    
    Dim PixNum() As Double
    
    Call ErrMsgColorAll(pColor, "ReadPixel")
    
    Call Num(PixNum, pColor)
    
    If pMaxPixel <= 0 Or PixNum(pSite) < pMaxPixel Then
        pMaxPixel = PixNum(pSite)
    End If
    
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    On Error GoTo ERROR_IGXL
    If m_BitDepth = idpDepthF32 Then
        ReDim pResultArr(pMaxPixel - 1)
        Call TheHdw.IDP.ReadPixel(pSite, m_Name, pIdpColorType, pResultArr, pMaxPixel)
    Else
        ReDim lngTmp(pMaxPixel - 1)
        Call TheHdw.IDP.ReadPixel(pSite, m_Name, pIdpColorType, lngTmp, pMaxPixel)
        
        ReDim pResultArr(UBound(lngTmp))
        For i = 0 To UBound(lngTmp)
            pResultArr(i) = lngTmp(i)
        Next i
    End If
    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("ReadPixel")
    
End Sub

Public Function ShiftLeft(ByRef pSrc As CImgPlane, ByVal pNumShift As Long, Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR) As CImgPlane
'内容:
'   左シフト
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pNumShift]   IN   Long型:             シフト数
'[pColorD]     IN   IdpColorType型:     保存先の色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.ShiftLeft(src, 3, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set ShiftLeft = Me
        Call TheError.Raise(999, "ShiftLeft", ErrMsgReadOnly)
        Exit Function
    End If
    
    If pNumShift < 0 Then
        Call TheError.Raise(999, "ShiftLeft", "Don't support minus number")
        Exit Function
    End If
    
    If m_BitDepth = idpDepthF32 Then
        Call TheError.Raise(999, "ShiftLeft", "Don't support float plane.")
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.ShiftLeft(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pNumShift)
    Set ShiftLeft = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("ShiftLeft")
    
End Function

Public Function ShiftRight(ByRef pSrc As CImgPlane, ByVal pNumShift As Long, Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR) As CImgPlane
'内容:
'   右シフト
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pNumShift]   IN   Long型:             シフト数
'[pColorD]     IN   IdpColorType型:     保存先の色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.ShiftRight(src, 3, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set ShiftRight = Me
        Call TheError.Raise(999, "ShiftRight", ErrMsgReadOnly)
        Exit Function
    End If
    
    If pNumShift < 0 Then
        Call TheError.Raise(999, "ShiftRight", "Don't support minus number")
        Exit Function
    End If
    
    If m_BitDepth = idpDepthF32 Then
        Call TheError.Raise(999, "ShiftRight", "Don't support float plane.")
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.ShiftRight(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pNumShift)
    Set ShiftRight = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("ShiftRight")
    
End Function

Public Function SubColumns( _
    ByRef pSrc As CImgPlane, ByVal pSrcCols As Long, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBitS As Long _
) As CImgPlane
'内容:
'   列同士の減算
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pSrcCols]    IN   Long型:             列間の幅
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             マスクビット
'[pMaskBitS]   IN   Long型:             データ元のマスクビット
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.SubColumns(src, 2, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set SubColumns = Me
        Call TheError.Raise(999, "SubColumns", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.SubColumns(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pSrcCols, pMaskBitS, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set SubColumns = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("SubColumns")
    
End Function

Public Function SubRows( _
    ByRef pSrc As CImgPlane, ByVal pSrcRows As Long, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, Optional ByRef pFlgName As String = "", _
    Optional ByVal pMaskBitD As Long, Optional ByVal pMaskBitS As Long _
) As CImgPlane
'内容:
'   行同士の減算
'
'[pSrc]        IN   CImgPlane型:        データ元
'[pSrcRows]    IN   Long型:             行間の幅
'[pColorD]     IN   IdpColorType型:     色指定
'[pColorS]     IN   IdpColorType型:     データ元の色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBitD]   IN   Long型:             マスクビット
'[pMaskBitS]   IN   Long型:             データ元のマスクビット
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.SubRows(src, 2, idpColorAll, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set SubRows = Me
        Call TheError.Raise(999, "SubRows", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    
    pIdpColorTypeD = ColorToIdpColorType(pColorD)
    pIdpColorTypeS = SelectColor(pColorS, pSrc, pColorD)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.SubRows(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pSrcRows, pMaskBitS, pMaskBitD, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Set SubRows = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("SubRows")
    
End Function

Public Function WriteFile(ByVal pSite As Long, ByVal pFileName As String, Optional ByVal pFileType As IdpFileFormat = idpFileBinary, Optional ByRef pColor As Variant = EEE_COLOR_FLAT) As CImgPlane
'内容:
'   ファイルに画像データを書き込む
'
'[pSite]       IN   Long型:             サイト指定
'[pFileName]   IN   String型:           ファイル名
'[pFileType]   IN   IdpFileFormat型:    データのタイプ
'[pColor]      IN   IdpColorType型:     色指定
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.WriteFile(0, "File", idpFileBinary, idpColorGreen)
'
    Call ErrMsgColorAll(pColor, "WriteFile")
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.WriteFile(pSite, m_Name, pIdpColorType, pFileName, pFileType)
    Set WriteFile = Me
    
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("WriteFile")
    
End Function

Public Function WritePixel(ByVal pConst As Variant, Optional ByRef pColor As Variant = EEE_COLOR_ALL, Optional ByVal pSite As Long = -1, Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long) As CImgPlane
'内容:
'   現在のPMDにデータを書き込む
'
'[pConst]      IN   Variant型:          書き込む数値
'[pColor]      IN   IdpColorType型:     色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.SetPMD("ZONE2D")
'    Call dst.WritePixel(100, idpColorAll)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set WritePixel = Me
        Call TheError.Raise(999, "WritePixel", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    On Error GoTo ERROR_IGXL
    Dim site As Long
    If IsArray(pConst) Then
        If pSite = -1 Then
            For site = 0 To SiteLimit
                Call TheHdw.IDP.WritePixel(m_Name, pIdpColorType, pConst(site), site, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName), pMaskBit)
            Next site
        Else
            Call TheHdw.IDP.WritePixel(m_Name, pIdpColorType, pConst(pSite), pSite, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName), pMaskBit)
        End If
    Else
        Call TheHdw.IDP.WritePixel(m_Name, pIdpColorType, pConst, pSite, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName), pMaskBit)
    End If
    
    Set WritePixel = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("WritePixel")
    
End Function

Public Function WritePixelAddr(ByRef pDataArray() As T_PIXINFO, Optional ByVal pSite As Long = -1) As CImgPlane
'内容:
'   指定した位置にデータを書き込む
'
'[pDataArray()]     IN  T_PIXINFO型:    書き込むデータ配列
'[pSite]            IN  Long型:         サイト
'
'備考:
'   サンプルコード
'    Dim dst As CImgPlane
'    Dim data(0) As T_PIXINFO
'
'    data(0).X = 100
'    data(0).Y = 100
'    data(0).Value = 500
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.SetPMD("ZONE2D")
'    Call dst.WritePixelAddr(data, 0)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set WritePixelAddr = Me
        Call TheError.Raise(999, "WritePixelAddr", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim dblTmp() As Double
    Dim lngTmp() As Long
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ERROR_IGXL
    If m_BitDepth = idpDepthF32 Then
        ReDim dblTmp(UBound(pDataArray), 2)
        For i = 0 To UBound(pDataArray)
            dblTmp(i, 0) = pDataArray(i).x
            dblTmp(i, 1) = pDataArray(i).y
            dblTmp(i, 2) = pDataArray(i).Value
        Next i
        
        Call TheHdw.IDP.WritePixelAddr(m_Name, idpColorFlat, dblTmp, UBound(dblTmp, 1) + 1, pSite)
    Else
        ReDim lngTmp(UBound(pDataArray), 2)
        For i = 0 To UBound(pDataArray)
            lngTmp(i, 0) = pDataArray(i).x
            lngTmp(i, 1) = pDataArray(i).y
            lngTmp(i, 2) = pDataArray(i).Value
        Next i
        
        Call TheHdw.IDP.WritePixelAddr(m_Name, idpColorFlat, lngTmp, UBound(lngTmp, 1) + 1, pSite)
    End If
    
    Set WritePixelAddr = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("WritePixelAddr")
    
End Function

Public Function WritePixelRamp(ByVal pOrgData As Double, ByVal pIncX As Double, ByVal pIncY As Double, Optional ByRef pColor As Variant = EEE_COLOR_FLAT, Optional ByVal pSite As Long = -1) As CImgPlane
'内容:
'   線形のデータを書き込む
'
'[pOrgData]    IN   Double型:           原点のデータ
'[pIncX]       IN   Double型:           横方向の増分
'[pIncY]       IN   Double型:           縦方向の増分
'[pColor]      IN   IdpColorType型:     色指定
'
'備考:
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.SetPMD("ZONE2D")
'    Call dst.WritePixelRamp(100, 1, 2, idpColorAll, 0)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set WritePixelRamp = Me
        Call TheError.Raise(999, "WritePixelRamp", ErrMsgReadOnly)
        Exit Function
    End If
    
    Dim dblTmp() As Double
    Dim lngTmp() As Long
    
    Dim pIdpColorType As IdpColorType
    pIdpColorType = ColorToIdpColorType(pColor)
    
    On Error GoTo ERROR_IGXL
    If m_BitDepth = idpDepthF32 Then
        ReDim dblTmp(2)
        dblTmp(0) = pOrgData
        dblTmp(1) = pIncX
        dblTmp(2) = pIncY
        Call TheHdw.IDP.WritePixelRamp(m_Name, pIdpColorType, dblTmp, pSite)
    Else
        ReDim lngTmp(2)
        lngTmp(0) = pOrgData
        lngTmp(1) = pIncX
        lngTmp(2) = pIncY
        Call TheHdw.IDP.WritePixelRamp(m_Name, pIdpColorType, lngTmp, pSite)
    End If
    
    Set WritePixelRamp = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("WritePixelRamp")
    
End Function

'2009/09/15 D.Maruyama ColorAll処理追加 ここから
Public Sub AverageColorAll( _
    ByRef pResult As ColorAllResult, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   平均値取得
'
'[pResult]      OUT  ColorAllResult型:  結果格納用構造体(色別サイト分)
'[pFlgName]     IN   String型:          フラグ名
'[pMaskBit]     IN   Long型:            マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.AverageColorAll(res)
'
        
    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.Accumulate(m_Name, idpColorAll, idpAccumMean, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    Call ReadResultFPColorAll(idpReadMean, pResult)
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("AverageColorAll")
    
End Sub

Public Sub SumColorAll( _
    ByRef pResult As ColorAllResult, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   合計値取得
'
'[pResult]     OUT  ColorAllResult型:   結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.SumColorAll(res)
    
    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.Accumulate(m_Name, idpColorAll, idpAccumSum, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    
    Call ReadResultFPColorAll(idpReadSum, pResult)
    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("SumColorAll")
    
End Sub

Public Sub StdDevColorAll( _
    ByRef pResult As ColorAllResult, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   標準偏差取得
'
'[pResult]     OUT  ColorAllResult型:   結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.StdDevColorAll(res)
'
    
    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.Accumulate(m_Name, idpColorAll, idpStdDeviation, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    
    Call ReadResultFPColorAll(idpReadStdDeviation, pResult)

    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("StdDevColorAll")
    
End Sub

Public Sub NumColorAll( _
    ByRef pResult As ColorAllResult, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   対象の画素の個数を取得
'
'[pResult]     OUT  ColorAllResult型:   結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.NumColorAll(res)
'
    
    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.Accumulate(m_Name, idpColorAll, idpAccumSum, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    
    Call ReadResultFPColorAll(idpReadNum, pResult)

    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("NumColorAll")
    
End Sub

Public Sub MaxColorAll( _
    ByRef pResult As ColorAllResult, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   最大値取得
'
'[pResult]     OUT  ColorAllResult型:   結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.MaxColorAll(res)
'
        
    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.MinMax(m_Name, idpColorAll, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    If m_BitDepth = idpDepthF32 Then
        Call ReadResultFPColorAll(idpReadMax, pResult)
    Else
        Call ReadResultColorAll(idpReadMax, pResult)
    End If
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("MaxColorAll")
    
End Sub

Public Sub MinColorAll( _
    ByRef pResult As ColorAllResult, _
   Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   最小値取得
'
'[pResult]     OUT  ColorAllResult型:   結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.MinColorAll(res)
'
    
    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.MinMax(m_Name, idpColorAll, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    If m_BitDepth = idpDepthF32 Then
        Call ReadResultFPColorAll(idpReadMin, pResult)
    Else
        Call ReadResultColorAll(idpReadMin, pResult)
    End If
     
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("MinColorAll")
    
End Sub

Public Sub MinMaxColorAll( _
    ByRef pResultMin As ColorAllResult, ByRef pResultMax As ColorAllResult, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   最大値と最小値を取得
'
'[pResultMin]   OUT  ColorAllResult型:        結果(最小値)を返す配列
'[pResultMax]   OUT  ColorAllResult型:        結果(最大値)を返す配列
'[pColor]      IN   IdpColorType型:     色指定
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim resMin As ColorAllResult, resMax As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.MinMaxColorAll(resMin,resMax)
'
    On Error GoTo ERROR_IGXL

    Call TheHdw.IDP.MinMax(m_Name, idpColorAll, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    If m_BitDepth = idpDepthF32 Then
        Call ReadResultFPColorAll(idpReadMin, pResultMin)
        Call ReadResultFPColorAll(idpReadMax, pResultMax)
    Else
        Call ReadResultColorAll(idpReadMin, pResultMin)
        Call ReadResultColorAll(idpReadMax, pResultMax)
    End If
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("MinMaxColorAll")

End Sub

Public Sub DiffMinMaxColorAll( _
    ByRef pResult As ColorAllResult, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   最大値と最小値の差を取得
'
'[pResult]     OUT  ColorAllResult型:   結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.DiffMinMaxColorAll(res)
'
    Dim resMin As ColorAllResult
    Dim site As Long
        
    
    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.MinMax(m_Name, idpColorAll, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
        
    If m_BitDepth = idpDepthF32 Then
        Call ReadResultFPColorAll(idpReadMax, pResult)
        Call ReadResultFPColorAll(idpReadMin, resMin)
    Else
        Call ReadResultColorAll(idpReadMax, pResult)
        Call ReadResultColorAll(idpReadMin, resMin)
    End If

    
    Dim ColorIndex As Long
    For site = 0 To SiteLimit
        For ColorIndex = 0 To MAX_COLOR
            pResult.color(ColorIndex).SiteValue(site) = pResult.color(ColorIndex).SiteValue(site) - resMin.color(ColorIndex).SiteValue(site)
        Next ColorIndex
    Next site
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("DiffMinMaxColorAll")
    
End Sub

Public Sub AbsMaxColorAll( _
    ByRef pResult As ColorAllResult, _
    Optional ByRef pFlgName As String = "", Optional ByVal pMaskBit As Long = 0 _
)
'内容:
'   絶対値での最大値取得
'
'[pResult]     OUT  ColorAllResult型:   結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:           フラグ名
'[pMaskBit]    IN   Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合は1が立っている箇所が演算対象
'   符号はそのまま
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.AbsMaxColorAll(res)
'
    Dim resMin As ColorAllResult
    Dim site As Long
        
    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.MinMax(m_Name, idpColorAll, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName))
    
    If m_BitDepth = idpDepthF32 Then
        Call ReadResultFPColorAll(idpReadMax, pResult)
        Call ReadResultFPColorAll(idpReadMin, resMin)
    Else
        Call ReadResultColorAll(idpReadMax, pResult)
        Call ReadResultColorAll(idpReadMin, resMin)
    End If
    
    Dim ColorIndex As Long
    For site = 0 To SiteLimit
        For ColorIndex = 0 To MAX_COLOR
            If Abs(resMin.color(ColorIndex).SiteValue(site)) > Abs(pResult.color(ColorIndex).SiteValue(site)) Then
                pResult.color(ColorIndex).SiteValue(site) = resMin.color(ColorIndex).SiteValue(site)
            End If
        Next ColorIndex
    Next site
    
    Exit Sub
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("AbsMaxColorAll")
    
End Sub

'#FlagExpansion
Public Sub CountColorAll( _
    ByRef pResult As ColorAllResult, ByVal pCountType As IdpCountType, ByVal pLoLim As Variant, ByVal pHiLim As Variant, _
    ByVal pCountLimMode As IdpCountLimitMode, Optional ByVal pLimitType As IdpLimitType = idpLimitExclude, _
    Optional ByVal pFlgName As String = "", Optional ByVal pInputFlgName As String = "", Optional ByVal pMaskBit As Long _
)
'内容:
'   指定条件に合う画素の個数をカウントする。
'
'[pResult]          OUT  ColorAllResult型:   結果格納用構造体(色別サイト分)
'[pCountType]       IN  IdpCountType型:     カウント方法の指定
'[pLoLim]           IN  Variant型:          Low側のリミット(サイト配列も可)
'[pHiLim]           IN  Variant型:          High側のリミット(サイト配列も可)
'[pCountLimMode]    IN  IdpCountLimitMode:  リミット値のとり方
'[pLimitType]       IN  IdpLimitType型:     リミット値を含む、含まないの指定
'[pFlgName]         IN  String型:           フラグ名
'[pInputFlgName]    IN  String型:           入力フラグ名
'[pMaskBit]         IN  Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合、共用フラグプレーンのカウント対象の位置に1を立てる。
'   入力フラグ名を指定した場合、1の立っている箇所がカウントの対象になる。
'   パラメータはTheHdw.IDP.Countに準拠
'   実際使っているのはCountMultiLimit
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim res As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Call src.CountColorAll(res, idpCountBetween, 0, 100, idpLimitEachSite, idpLimitInclude)
'
    Dim loLimArr() As Double
    Dim hiLimArr() As Double
    
    If pCountLimMode = idpLimitEachSite Then
    
        If Not IsValidLimitArray(pLoLim) Then
            Exit Sub
        End If
        If Not IsValidLimitArray(pHiLim) Then
            Exit Sub
        End If
        Call MakeLimitArray(pLoLim, loLimArr)
        Call MakeLimitArray(pHiLim, hiLimArr)
        
    ElseIf pCountLimMode = idpLimitEachSiteColor Then
    
        loLimArr = pLoLim
        hiLimArr = pHiLim
        
    Else
        Call TheError.Raise(999, "CountColorAll", "idpLimitEachSite and idpLimitEachSiteColor are only available")
    End If
        
    On Error GoTo ERROR_IGXL
    
    If pFlgName <> "" Then
        Call GetSharedFlagPlane(pFlgName).SetFlagBit(pFlgName)
    End If
    Call TheHdw.IDP.CountMultiLimit(m_Name, idpColorAll, pCountType, loLimArr, hiLimArr, pCountLimMode, pLimitType, pMaskBit, GetOptPlaneName(pFlgName), GetOptFlgBit(pFlgName), GetOptPlaneName(pInputFlgName), GetOptFlgBit(pInputFlgName))
    
    Call ReadResultColorAll(idpReadCount, pResult)
    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("CountColorAll")
    
End Sub

Public Sub CountColorAllForFlgBitImgPlane( _
    ByRef pResult As ColorAllResult, ByVal pCountType As IdpCountType, ByVal pLoLim As Variant, ByVal pHiLim As Variant, _
    ByVal pCountLimMode As IdpCountLimitMode, _
    ByRef pFlgPlane As CImgPlane, ByVal pFlgBit As Long, _
    Optional ByVal pLimitType As IdpLimitType = idpLimitExclude, _
    Optional ByVal pInputFlgName As String = "", Optional ByVal pMaskBit As Long _
)
'内容:
'   指定条件に合う画素の個数をカウントする。(フラグビットをイメージプレンに立てる)
'
'[pResult]          OUT  ColorAllResult型:   結果格納用構造体(色別サイト分)
'[pCountType]       IN  IdpCountType型:     カウント方法の指定
'[pLoLim]           IN  Variant型:          Low側のリミット(サイト配列も可)
'[pHiLim]           IN  Variant型:          High側のリミット(サイト配列も可)
'[pCountLimMode]    IN  IdpCountLimitMode:  リミット値のとり方
'[pFlgName]         IN  CImgPlane型:        フラグを出力するプレン
'[pFlgBit]          IN  pFlgBit型:          フラグビット
'[pLimitType]       IN  IdpLimitType型:     リミット値を含む、含まないの指定
'[pInputFlgName]    IN  String型:           入力フラグ名
'[pMaskBit]         IN  Long型:             マスクビット
'
'備考:
'   フラグ名を指定した場合、共用フラグプレーンのカウント対象の位置に1を立てる。
'   入力フラグ名を指定した場合、1の立っている箇所がカウントの対象になる。
'   パラメータはTheHdw.IDP.Countに準拠
'   実際使っているのはCountMultiLimit
'
'   サンプルコード
'    Dim src As CImgPlane, flg As CImgPlane
'    Dim res As ColorAllResult
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.SetPMD("ZONE2D")
'    Set flg = TheIDP.TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16, True)
'    Call src.SetPMD("ZONE2D")
'    Call src.CountColorAllForFlgBitImgPlane(res, idpCountBetween, 0, 100, idpLimitEachSite, flg, 1, idpLimitInclude)
'   =>関数をまたがない連続点とか
'

    Dim loLimArr() As Double
    Dim hiLimArr() As Double
    
    If pCountLimMode = idpLimitEachSite Then
    
        If Not IsValidLimitArray(pLoLim) Then
            Exit Sub
        End If
        If Not IsValidLimitArray(pHiLim) Then
            Exit Sub
        End If
        Call MakeLimitArray(pLoLim, loLimArr)
        Call MakeLimitArray(pHiLim, hiLimArr)
        
    ElseIf pCountLimMode = idpLimitEachSiteColor Then
    
        loLimArr = pLoLim
        hiLimArr = pHiLim
        
    Else
        Call TheError.Raise(999, "CountColorAllForFlgBitImgPlane", "idpLimitEachSite and idpLimitEachSiteColor are only available")
    End If

    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.CountMultiLimit(m_Name, idpColorAll, pCountType, loLimArr, hiLimArr, pCountLimMode, pLimitType, pMaskBit, pFlgPlane.Name, pFlgBit, GetOptPlaneName(pInputFlgName), GetOptFlgBit(pInputFlgName))
    
    Call ReadResultColorAll(idpReadCount, pResult)

    
    Exit Sub

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("CountColorAllForFlgBitImgPlane")
    
End Sub

Public Sub PutFlagColorAll( _
    ByVal pFlgName As String, ByVal pCountType As IdpCountType, ByVal pLoLim As Variant, ByVal pHiLim As Variant, _
    ByVal pCountLimMode As IdpCountLimitMode, _
    Optional ByVal pLimitType As IdpLimitType = idpLimitExclude, _
    Optional ByVal pInputFlgName As String = "", Optional ByVal pMaskBit As Long _
)
'内容:
'   指定条件に合う画素の位置にフラグを立てる。
'
'[pFlgName]         IN  String型:           フラグ名
'[pCountType]       IN  IdpCountType型:     カウント方法の指定
'[pLoLim]           IN  Variant型:          Low側のリミット(サイト配列も可)
'[pHiLim]           IN  Variant型:          High側のリミット(サイト配列も可)
'[pCountLimMode]    IN  IdpCountLimitMode:  リミット値のとり方
'[pLimitType]       IN  IdpLimitType型:     リミット値を含む、含まないの指定
'[pInputFlgName]    IN  String型:           入力フラグ名
'[pMaskBit]         IN  Long型:             マスクビット
'
'備考:
'   入力フラグ名を指定した場合、1の立っている箇所がカウントの対象になる。
'   パラメータはTheHdw.IDP.Countに準拠
'
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.PutFlagColorAll(src, idpCountBetween, 0, 100, idpLimitEachSite, idpLimitInclude,  1)
'
    
    Dim tmp As ColorAllResult
    
    Call CountColorAll(tmp, pCountType, pLoLim, pHiLim, pCountLimMode, pLimitType, pFlgName, pInputFlgName, pMaskBit)
    
End Sub

Private Sub ReadResultColorAll(ByVal pReadType As IdpReadType, ByRef pResult As ColorAllResult)

    Dim site As Long
    Dim tmp(MAX_COLOR) As Long
    
    On Error GoTo ERROR_RAISE
    
    Dim ColorIndex As Long
    With pResult
        For ColorIndex = 0 To MAX_COLOR
            Call Redim_(.color(ColorIndex).SiteValue)
        Next ColorIndex
    End With
    
    For site = 0 To SiteLimit
        Call TheHdw.IDP.ReadResult(site, m_Name, idpColorAll, pReadType, tmp)
        With pResult
            For ColorIndex = 0 To MAX_COLOR
                .color(ColorIndex).SiteValue(site) = tmp(ColorIndex)
            Next ColorIndex
        End With
    Next site
    
    Exit Sub
    
ERROR_RAISE:
    Call Err.Raise(999, "ReadResultColorAll", Err.Description)
    
End Sub

Private Sub ReadResultFPColorAll(ByVal pReadType As IdpReadType, ByRef pResult As ColorAllResult)

    Dim site As Long
    Dim tmp(MAX_COLOR) As Double
    
    On Error GoTo ERROR_RAISE
    
    Dim ColorIndex As Long
    With pResult
        For ColorIndex = 0 To MAX_COLOR
            Call Redim_(.color(ColorIndex).SiteValue)
        Next ColorIndex
    End With
    
    For site = 0 To SiteLimit
        Call TheHdw.IDP.ReadResultFP(site, m_Name, idpColorAll, pReadType, tmp)
        With pResult
            For ColorIndex = 0 To MAX_COLOR
                .color(ColorIndex).SiteValue(site) = tmp(ColorIndex)
            Next ColorIndex
        End With
    Next site
    
    
    Exit Sub
    
ERROR_RAISE:
    Call Err.Raise(999, "ReadResultFPColorAll", Err.Description)
    
End Sub

'2009/09/15 D.Maruyama ColorAll処理追加 ここまで

Public Function Acquire( _
    ByVal pPins As String, ByVal pNumFrame As Long, ByVal pAverageMode As IdpAverageMode, ByVal pAcqMode As IdpAcquireMode, _
    Optional ByVal pNumFields As Long = 2 _
) As CImgPlane
'内容:
'   画像の撮り込み
'
'[pPins]        IN  String型:           対象のピンリスト
'[pNumFrame]    IN  Long型:             撮り込むフレーム数
'[pAverageMode] IN  IdpAverageMode型:   アベレージの方法
'[pAcqMode]     IN  IdpAcquireMode型:   撮り込み方法
'[pNumFields]   IN  Long型:             インターレース時のフレーム数
'
'備考:
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.SetPMD("FULL")
'    Call dst.Acquire("VOUT", 10, idpAverage, idpAcqNonInterlace)
'
    If m_ReadOnly = True Then
        'エラー処理
        Set Acquire = Me
        Call TheError.Raise(999, "Acquire", ErrMsgReadOnly)
        Exit Function
    End If
    
    On Error GoTo ERROR_IGXL
    Call TheHdw.IDP.Acquire(pPins, m_Name, pNumFrame, pAverageMode, pAcqMode, pNumFields)
    Set Acquire = Me
    
    Exit Function

ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("Acquire")
    
End Function

Public Function ColorByAddress(ByVal pX As Long, ByVal pY As Long, Optional pAddrMode As IdpAddrMode = idpAddrAbsolute) As CImgColor
'内容:
'   指定したXYアドレスの色オブジェクト(CImgColor)を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月1日
'パラメータ:
'   [pX]        In  1):
'   [pY]        In  2):
'   [pAddrMode] In  3):アドレスが、絶対アドレス(BasePMD)か相対アドレス(CurrentPMD)か
'戻り値:
'   CImgColor
'
'注意事項:
'
    Dim pColorX As Long
    Dim pColorY As Long
    
    Dim pBasePMD As CImgPmdInfo
    Dim pPlaneMap As CImgPlaneMap
    
    '入力X、YをCImgPlaneMap上のアドレスに換算
    If pAddrMode = idpAddrAbsolute Then
        Set pBasePMD = BasePMD
        
        '対象範囲外はエラー
        If pX < pBasePMD.Left Or pX > pBasePMD.Right Then
            Call TheError.Raise(9999, "CImgplane.ColorAddress", "Address is out of range. " & pX)
        End If
        If pY < pBasePMD.Top Or pY > pBasePMD.Bottom Then
            Call TheError.Raise(9999, "CImgplane.ColorAddress", "Address is out of range. " & pY)
        End If
        
        pColorX = (pX - 1) Mod m_PlaneMap.width + 1
        pColorY = (pY - 1) Mod m_PlaneMap.height + 1
    ElseIf pAddrMode = idpAddrRelative Then
        '対象範囲外はエラー
        If pX < 1 Or pX > m_CurrentPMD.width Then
            Call TheError.Raise(9999, "CImgplane.ColorAddress", "Address is out of range. " & pX)
        End If
        If pY < 1 Or pY > m_CurrentPMD.height Then
            Call TheError.Raise(9999, "CImgplane.ColorAddress", "Address is out of range. " & pY)
        End If
    
        pColorX = ((m_CurrentPMD.Left - 1 + pX) - 1) Mod m_PlaneMap.width + 1
        pColorY = ((m_CurrentPMD.Top - 1 + pY) - 1) Mod m_PlaneMap.height + 1
    Else
        'addrMode指定間違い
        Call TheError.Raise(9999, "CImgplane", "IdpAddrMode is different.")
    End If
    
    'PlaneMapアドレスでの色を返す。
    Set ColorByAddress = m_PlaneMap.color(m_PlaneMap.ColorNameAddress(pColorX, pColorY))

End Function

Private Sub ReadResult(ByVal pColor As IdpColorType, ByVal pReadType As IdpReadType, ByRef pResult() As Double)

    Dim site As Long
    Dim tmp(MAX_COLOR) As Long
    
    On Error GoTo ERROR_RAISE
    Call Redim_(pResult)
    For site = 0 To SiteLimit
        Call TheHdw.IDP.ReadResult(site, m_Name, pColor, pReadType, tmp)
        pResult(site) = tmp(ReadIndex(pColor))
    Next site
    
    Exit Sub
    
ERROR_RAISE:
    Call Err.Raise(999, "ReadResult", Err.Description)
    
End Sub

Private Sub ReadResultFP(ByVal pColor As IdpColorType, ByVal pReadType As IdpReadType, ByRef pResult() As Double)

    Dim site As Long
    Dim tmp(MAX_COLOR) As Double
    
    On Error GoTo ERROR_RAISE
    Call Redim_(pResult)
    For site = 0 To SiteLimit
        Call TheHdw.IDP.ReadResultFP(site, m_Name, pColor, pReadType, tmp)
        pResult(site) = tmp(ReadIndex(pColor))
    Next site
    
    Exit Sub
    
ERROR_RAISE:
    Call Err.Raise(999, "ReadResultFP", Err.Description)
    
End Sub

Private Function SiteLimit() As Long
    SiteLimit = TheExec.sites.ExistingCount - 1
End Function

Private Function ReadIndex(ByVal pColor As IdpColorType) As IdpColorType

    If pColor = idpColorFlat Then
        ReadIndex = 0
    Else
        ReadIndex = pColor
    End If
    
End Function

Private Function OptPlaneName(ByRef pPlane As CImgPlane) As String
    
    If pPlane Is Nothing Then
        OptPlaneName = ""
    Else
        OptPlaneName = pPlane.Name
    End If
End Function


'''' 0000130343
Private Function OtherGroupGetFlag(ByVal pFlgName As String)

    Dim i As Integer
    Dim sPlane As CImgPlane
    Dim resPlane As CImgPlane
    Dim resBit As Integer
    Dim dstPlane As CImgPlane
    Dim dstBit As Integer

    Dim resZone As String
    Dim dstZone As String

    '' pFlgName がNULLの場合は、この処理は不要
    If pFlgName = "" Then Exit Function

    '' 条件:フラグが存在していること
    If GetSharedFlagPlane(pFlgName).FlagBit(pFlgName) <> 0 Then
        Exit Function
    Else
    '' フラグが存在していない場合は、まず、ほかのプレーングループのフラグをチェックする
        For i = 1 To TheIDP.PlaneManagerCount
            If TheIDP.PlaneManager(i).GetSharedFlagPlane(pFlgName).FlagBit(pFlgName) <> 0 Then
                '' 結果のフラグビットを確保する。
                Call GetSharedFlagPlane(pFlgName).SetFlagBit(pFlgName)
                '' コピー前後のフラグプレーンを指定する。
                Set dstPlane = GetSharedFlagPlane(pFlgName).FlgPlane
                Set resPlane = TheIDP.PlaneManager(i).GetSharedFlagPlane(pFlgName).FlgPlane
                '' 対象のZONEを保管する
                resZone = resPlane.CurrentPMD.Name
                dstZone = dstPlane.CurrentPMD.Name
                
                '' ZONEのBitズレなどのチェックを入れるのであればここに記述する。
                dstBit = GetSharedFlagPlane(pFlgName).FlagBit(pFlgName)
                resBit = TheIDP.PlaneManager(i).GetSharedFlagPlane(pFlgName).FlagBit(pFlgName)
                '' 対象のBitのみコピーする。
                '' Bitのコピーだけであれば、もう少し早くできそうな気がする。
                Call dstPlane.SetPMD(dstPlane.BasePMD.Name)
                Call resPlane.SetPMD(resPlane.BasePMD.Name)
                
                Call dstPlane.LOr(resPlane, resPlane, EEE_COLOR_FLAT, , , dstBit, resBit, resBit)
                '' FULL → もとのプレーン情報に戻す。
                Call dstPlane.SetPMD(dstZone)
                Call resPlane.SetPMD(resZone)

                Exit Function
            End If
        Next i
    End If
    '' ここにくるのは、本当にどこでもフラグを生成していないケースのみ

End Function


'#FlagExpansion
Private Function GetOptPlaneName(ByVal pFlgName As String) As String
    
    Call OtherGroupGetFlag(pFlgName)
    
    If pFlgName = "" Then
        GetOptPlaneName = ""
    Else
        Dim FlagPlane As IImgFlag
        Set FlagPlane = GetSharedFlagPlane(pFlgName)
        
        If CurrentPMD.Name = Me.Manager.VariablePMDName Then
            FlagPlane.SetPMD CurrentPMD
        Else
            FlagPlane.SetPMD CurrentPMD.Name
        End If
        GetOptPlaneName = FlagPlane.Name
    End If
End Function

'#FlagExpansion
Private Function GetOptFlgBit(ByVal pFlgName As String) As Long
    
    Call OtherGroupGetFlag(pFlgName)
    
    If pFlgName = "" Then
        GetOptFlgBit = 0
    Else
        GetOptFlgBit = GetSharedFlagPlane(pFlgName).FlagBit(pFlgName)
        If GetOptFlgBit = 0 Then
            Call TheError.Raise(999, "", ErrMsgNameDoesntExist(pFlgName))
        End If
    End If
End Function

''#FlagExpansion
'Private Function SetOptFlgBit(ByVal pFlgName As String) As Long
'    If pFlgName = "" Then
'        SetOptFlgBit = 0
'    Else
'        SetOptFlgBit = GetSharedFlagPlane(pFlgName).SetFlagBit(pFlgName)
'        Call GetSharedFlagPlane(pFlgName).SetPMD(CurrentPMD.Name)
'        If SetOptFlgBit = 0 Then
'            Call TheError.Raise(999, "", ErrMsgNameDoesntExist(pFlgName))
'        End If
'    End If
'End Function

Private Function ColorToIdpColorType(ByRef pColor As Variant) As IdpColorType
'内容:
'   型不定な色指定をIdpColorTypeに変換します。
'   pColorはオブジェクト型、String型、IdpColorType型のどの型か不明。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月4日
'パラメータ:
'   [pColor]    In  1):
'戻り値:
'   IdpColorType
'
'注意事項:
'

    If TypeName(pColor) = "String" Then  'String型
        If pColor = EEE_COLOR_ALL Then
            ColorToIdpColorType = idpColorAll
        ElseIf pColor = EEE_COLOR_FLAT Then
            ColorToIdpColorType = idpColorFlat
        Else
            ColorToIdpColorType = m_PlaneMap.color(CStr(pColor)).IdpColor
        End If
        Exit Function
    ElseIf TypeName(pColor) = "Long" Then  'IdpColorType
        If pColor >= -2 And pColor <= 11 Then
            ColorToIdpColorType = pColor
            Exit Function
        End If
    ElseIf TypeName(pColor) = "CImgColor" Then  'CImgColor型
        ColorToIdpColorType = pColor.IdpColor
        Exit Function
    End If

Call TheError.Raise(9999, "", "Invalid Color!")

End Function

'#FlagExpansion
Public Function GetSharedFlagPlane(ByVal pFlagName As String) As IImgFlag
'内容:
'   属するグループの共用フラグプレーンを取得
'
'備考:
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    With dst.GetSharedFlagPlane
'        Debug.Print .FlagBitList
'    End With
'
    Set GetSharedFlagPlane = m_Parent.Manager.GetSharedFlagPlane(pFlagName)
End Function

Public Property Get Manager() As IImgPlaneManager
    Set Manager = m_Parent.Manager
End Property

Private Function Var2PlaneName(ByVal pVal As Variant) As String
    
    If IsObject(pVal) Then
        Var2PlaneName = pVal.Name
    Else
        Var2PlaneName = ""
    End If
End Function

Private Function Var2Num(ByVal pVal As Variant) As Variant
    
    If IsObject(pVal) Then
        Var2Num = 0
    Else
        Var2Num = pVal
    End If
End Function

Private Sub MakeLimitArray(ByVal pVarLim As Variant, ByRef pLimArr() As Double)
    
    Dim site As Long
    
    If IsArray(pVarLim) Then
        pLimArr = pVarLim
    Else
        ReDim pLimArr(SiteLimit)
        For site = 0 To SiteLimit
            pLimArr(site) = pVarLim
        Next site
    End If
    
End Sub

Private Function SelectColor(ByRef pColor As Variant, ByRef pPlane As Variant, ByRef pRefColor As Variant) As IdpColorType
'内容:
'   色指定がない場合(pColor=SAME_COLOR)は、同じ色に指定する。
'変更者:
'  tomoyoshi.takase
'変更日: 2010年11月16日
'パラメータ:
'   [pColor]    In  1):指定
'   [pPlane]    In  2):指定した色の画像
'   [pRefColor] In  3):
'戻り値:
'   idpColorType
'
'注意事項:
'   SAME_COLORの場合は同じ色指定の文字があること

    If TypeName(pPlane) <> "CImgPlane" Then      'pPlaneがCImgPlaneでない時。固定値演算用の配列変数。
        SelectColor = ColorToIdpColorType(pRefColor)
    End If
    
    If TypeName(pColor) = "CImgColor" Then
        SelectColor = pColor.IdpColor
    ElseIf TypeName(pColor) = "Long" Then
        SelectColor = pColor
    ElseIf TypeName(pColor) = "String" Then
        If pColor = SAME_COLOR Then
           SelectColor = ColorToIdpColorType(pRefColor)
        ElseIf pColor = EEE_COLOR_ALL Then
            SelectColor = idpColorAll
        ElseIf pColor = EEE_COLOR_FLAT Then
            SelectColor = idpColorFlat
        Else
            SelectColor = pPlane.planeMap.color(CStr(pColor)).IdpColor
        End If
    Else
    
    End If
    
End Function

Private Sub Redim_(ByRef pResult() As Double)

    On Error GoTo FIX_ARRAY
    ReDim pResult(SiteLimit)
    
    Exit Sub
    
FIX_ARRAY:
    If UBound(pResult) <> SiteLimit Then
        'エラー
        Call Err.Raise(999, "Redim_", "Invalid Array. (Num = " & UBound(pResult) & ")")
    End If
    Erase pResult
    
End Sub

'=======2009/04/28 追加 Maruyama ここから==============
Private Function IsValidLimitArray(ByRef pArray As Variant) As Boolean

    'オブジェクトはエラー！
    If IsObject(pArray) Then
        IsValidLimitArray = False
        Call TheError.Raise(999, "LimArray", "LimArray is Object.")
        Exit Function
    End If

    '配列でないときはよい
    If IsNumeric(pArray) Then
        IsValidLimitArray = True
        Exit Function
    End If
    
    '配列の場合は2次元目の要素数を見に行って取れたらだめ。
    On Error GoTo ONE_ORDER_ARRAY
    Dim tmp As Long
    tmp = UBound(pArray, 2)
    On Error GoTo 0
    IsValidLimitArray = False
    Call TheError.Raise(999, "LimArray", "LimArray is over two order array.")
    Exit Function
    
ONE_ORDER_ARRAY:
    IsValidLimitArray = True
    Exit Function
    
End Function
'=======2009/04/28 Add Maruyama ここまで==============

Public Sub View(Optional MaskName As String = "")
'内容:
'   IDVの起動
'
'備考:
'   暫定機能
'   IDVに対象プレーンの画像を表示することはできるが、操作できない
'
'   サンプルコード
'    Dim src As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Call src.View
'
    With theidv
        Call .OpenForm
        .PlaneNameGreen = m_Name
        ' matsuno
        If MaskName <> "" Then
            .HilightColor = "Red"
            .PlaneNameFlag = GetSharedFlagPlane(MaskName).Name
            .MsbFlag = GetSharedFlagPlane(MaskName).FlagBit(MaskName)
            .PMDName = CurrentPMD.Name & "_S"
        End If
        ' matsuno

        Call .Refresh
    End With
End Sub

Private Function ErrMsgReadOnly() As String
    ErrMsgReadOnly = "Property is Read Only."
End Function

Private Sub ErrMsgColorAll(ByRef pColor As Variant, ByVal pErrSource As String)
    Dim pIdpColor As IdpColorType
    pIdpColor = ColorToIdpColorType(pColor)
    If pIdpColor = idpColorAll Then
        Call TheError.Raise(999, pErrSource, """idpColorAll"" is invalid parameter")
    End If
End Sub

Private Function ErrMsgNameDoesntExist(ByVal pName As String)
    ErrMsgNameDoesntExist = """" & pName & """ doesn't exist."
End Function

Private Sub ErrMsgIGXL(ByVal pErrSource As String)
    Call TheError.Raise(999, pErrSource, Err.Description)
End Sub

'#FlagExpansion
Private Function GetFlgBit(ByVal pName As String) As Long
    GetFlgBit = GetSharedFlagPlane(pName).FlagBit(pName)
    If GetFlgBit = 0 Then Call Err.Raise(999, "", pName & " is not exist")
End Function

'#FlagExpansion
Public Function FlagNot(ByVal pFlgName As String, Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT) As CImgPlane
'内容:
'   共用フラグプレーンのNot演算の結果を取得
'
'[pFlgName]     IN   CImgPlane型:       データ元のフラグ名
'[pDstBit]      IN   Long型:            保存先ビット指定
'[pColor]       IN   IdpColorType型:    色指定
'
'備考:
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.FlagNot("FLAG", 1, idpColorRed)
'
    On Error GoTo ERROR_IGXL
    Dim flag As IImgFlag
    Set flag = GetSharedFlagPlane(pFlgName)
    If CurrentPMD.Name = Me.Manager.VariablePMDName Then
        flag.SetPMD CurrentPMD
    Else
        flag.SetPMD CurrentPMD.Name
    End If
    
    Call LNot(flag.FlgPlane, pColor, , pDstBit, GetFlgBit(pFlgName))
    
    Set FlagNot = Me
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("FlagNot")
    
End Function

'#FlagExpansion
Public Function FlagOr(ByVal pFlgName1 As String, ByVal pFlgName2 As String, Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT) As CImgPlane
'内容:
'   共用フラグプレーンのOr演算の結果を取得
'
'[pFlgName1]    IN   CImgPlane型:       データ元1のフラグ名
'[pFlgName2]    IN   CImgPlane型:       データ元2のフラグ名
'[pDstBit]      IN   Long型:            保存先ビット指定
'[pColor]       IN   IdpColorType型:    色指定
'
'備考:
'   pFlgName1、pFlgName2は同名の指定可能。
'
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.FlagOr("FLAG1", "FLAG2", 1, idpColorRed)
'
    On Error GoTo ERROR_IGXL
    
    Dim FlagPlane1 As CImgPlane
    Set FlagPlane1 = GetInputFlagPlane(pFlgName1)
    Dim FlagPlane2 As CImgPlane
    Set FlagPlane2 = GetInputFlagPlane(pFlgName2)
    Call LOr(FlagPlane1, FlagPlane2, pColor, , , pDstBit, GetFlgBit(pFlgName1), GetFlgBit(pFlgName2))
                
    Set FlagOr = Me
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("FlagOr")
    
End Function

'#FlagExpansion
Public Function FlagAnd(ByVal pFlgName1 As String, ByVal pFlgName2 As String, Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT) As CImgPlane
'内容:
'   共用フラグプレーンのAnd演算の結果を取得
'
'[pFlgName1]    IN   CImgPlane型:       データ元1のフラグ名
'[pFlgName2]    IN   CImgPlane型:       データ元2のフラグ名
'[pDstBit]      IN   Long型:            保存先ビット指定
'[pColor]       IN   IdpColorType型:    色指定
'
'備考:
'   pFlgName1、pFlgName2は同名の指定可能。
'
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.FlagAnd("FLAG1", "FLAG2", 1, idpColorRed)
'
    On Error GoTo ERROR_IGXL
    
    Dim FlagPlane1 As CImgPlane
    Set FlagPlane1 = GetInputFlagPlane(pFlgName1)
    Dim FlagPlane2 As CImgPlane
    Set FlagPlane2 = GetInputFlagPlane(pFlgName2)
    Call LAnd(FlagPlane1, FlagPlane2, pColor, , , pDstBit, GetFlgBit(pFlgName1), GetFlgBit(pFlgName2))
        
    Set FlagAnd = Me
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("FlagAnd")
    
End Function

'#FlagExpansion
Public Function FlagXor(ByVal pFlgName1 As String, ByVal pFlgName2 As String, Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT) As CImgPlane
'内容:
'   共用フラグプレーンのXor演算の結果を取得
'
'[pFlgName1]    IN   CImgPlane型:       データ元1のフラグ名
'[pFlgName2]    IN   CImgPlane型:       データ元2のフラグ名
'[pDstBit]      IN   Long型:            保存先ビット指定
'[pColor]       IN   IdpColorType型:    色指定
'
'備考:
'   pFlgName1、pFlgName2は同名の指定可能。
'
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.FlagXor("FLAG1", "FLAG2", 1, idpColorRed)
'
    On Error GoTo ERROR_IGXL
    
    Dim FlagPlane1 As CImgPlane
    Set FlagPlane1 = GetInputFlagPlane(pFlgName1)
    Dim FlagPlane2 As CImgPlane
    Set FlagPlane2 = GetInputFlagPlane(pFlgName2)
    Call LXor(FlagPlane1, FlagPlane2, pColor, , , pDstBit, GetFlgBit(pFlgName1), GetFlgBit(pFlgName2))
        
    Set FlagXor = Me
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("FlagXor")
    
End Function

Public Function FlagCopy(ByVal pFlgName As String, Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT) As CImgPlane
'内容:
'   共用フラグプレーンの指定したビットの値をコピー
'
'[pFlgName]     IN   CImgPlane型:       データ元のフラグ名
'[pDstBit]      IN   Long型:            保存先ビット指定
'[pColor]       IN   IdpColorType型:    色指定
'
'備考:
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.FlagCopy("FLAG", 1, idpColorRed)
'
    On Error GoTo ERROR_IGXL
    Set FlagCopy = FlagOr(pFlgName, pFlgName, pDstBit, pColor)
    Exit Function
    
ERROR_IGXL:
    'エラー処理
    Call ErrMsgIGXL("FlagCopy")
    
End Function

Private Sub IProcParamReplaceable_GetContents(ByVal pRoot As String, ByRef pRes As Collection)
'内容:
'   パラメータの設定内容を取得する
'
'パラメータ:
'   [pRoot]     IN  String型:  このパラメータクラスの階層情報
'   [pRes]      IN/OUT  Collection型:  パラメータの設定内容を格納するコレクション
'
'注意事項:
'
'

    With pRes
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PLANEGROUP, Me.planeGroup, "String", False, PARAMETER_REPLACER_PLANEGROUP)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PLANENAME, m_Name, "String", False, PARAMETER_REPLACER_PLANENAME)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_BITDEPTH, CStr(m_BitDepth), "IdpBitDepth", False, PARAMETER_REPLACER_BITDEPTH)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_READONLY, CStr(m_ReadOnly), "Boolean", False, PARAMETER_REPLACER_READONLY)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_CURRENTPMD, Me.CurrentPmdName, "String", False, PARAMETER_REPLACER_CURRENTPMD)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COLORMAPNAME, m_PlaneMap.ColorNamesPmd, "String", False, PARAMETER_REPLACER_COLORMAPNAME)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PLANECOMMENT, m_Comment, "String", False, PARAMETER_REPLACER_PLANECOMMENT)
    End With
End Sub

Private Sub IProcParamReplaceable_ReplaceContent(ByVal pRoot As String, ByRef pParamContents As CImpLocalParamsContent)
'内容:
'   パラメータの内容を書き換える
'
'パラメータ:
'   [pRoot]             IN  String型:  このパラメータクラスの階層情報
'   [pParamContents]    IN  CImpLocalParamsContent型:  書き換えを行うパラメータの内容
'
'注意事項:
'
'

End Sub

Private Sub AddParamContents(ByRef pRes As Collection, ByVal pPath As String, ByVal pValue As String, ByVal pTypeName As String, ByVal pEditable As Boolean, ByVal pLabel As String)
'内容:
'   パラメータの情報をコレクションに追加する
'
'パラメータ:
'   [pRes]          IN/OUT  Collection型:  パラメータの設定内容を追加するコレクション
'   [pPath]         IN  String型:  パラメータの階層情報
'   [pValue]        IN  String型:  パラメータの値
'   [pTypeName]     IN  String型:  パラメータのデータ型
'   [pEditable]     IN  Boolean型:  パラメータが編集可能かどうか
'   [pLabel]        IN  String型:  パラメータ名
'
'注意事項:
'
'

    Dim myContents As New CImpLocalParamsContent
    Call myContents.Initialize(pPath, pValue, pTypeName, pEditable, pLabel)
    Call pRes.Add(myContents)
End Sub

'#FlagExpansion
'フラグ名を元にフラグプレーンを取り出し、結果出力プレーン(自分)と同じPMDに設定して返す関数
Private Function GetInputFlagPlane(ByVal pFlagName As String) As CImgPlane
    'フラグプレーン取得
    Dim flag As IImgFlag
    Set flag = Me.Manager.GetSharedFlagPlane(pFlagName)
    'PMD設定
    If CurrentPMD.Name = Me.Manager.VariablePMDName Then
        flag.SetPMD CurrentPMD
    Else
        flag.SetPMD CurrentPMD.Name
    End If
    Set GetInputFlagPlane = flag.FlgPlane
End Function

'/* ======= 2012/01/18 Add 0145206097 ここから ============== */
#If IDP_40IP02_USE <> 0 Then
Public Function ReadFileEx( _
        ByVal pSite As Long, ByVal pFileName As Variant, _
        Optional ByRef pColor As Variant = EEE_COLOR_FLAT, _
        Optional ByVal pSrcPC As IdpPCType = IdpTesterPC, _
        Optional ByVal pProcessMode As IdpRWFileExProcessMode = idpRWFileExProcessBlocking _
) As CImgPlane
'内容:
'   プレーンに画像データをファイルに読み込む。従来のReadFileの機能拡張版。
'
'[pSite]        In  読み込む対象のサイト番号
'[pFileName]    In  画像データのファイル名またはファイル名配列
'[pColor]       In  プレーンのカラー指定
'[pSrcPC]       In  ファイル読み込み元のPC指定
'[pProcessMode] In  ReadFileExの動作モード指定
'
'戻り値:
'   自分自身の参照を返す
'
'注意事項:
'   IG-XL制約事項
'   【ファイル指定制約】:エラー検出はIG-XL APIに委ねる
'   ①サイト指定が-1の場合、指定ファイル名はサイト分の要素数の文字列配列が指定されていること
'   ②サイト指定が-1の場合、サイトで指定ファイル名が重複していないこと
'
'   【カラー指定制約】:Eee-JOBで検知
'   ③格納プレーンの色指定がColorAllでないこと
'
'備考:
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.SetPMD("ZONE2D")
'    Call dst.ReadFileEx(0, "Tmp.txt")
'
    Const FUNCTION_NAME = "ReadFileEx"

    '/* 書き込み禁止のプレーンだったらエラー */
    If m_ReadOnly = True Then
        Set ReadFileEx = Me
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx)
    End If

    '/* カラータイプ型変換 */
    On Error GoTo ErrHandler
    Dim pIdpColorType As IdpColorType
    pIdpColorType = GetIdpColorType(m_PlaneMap, pColor)
    '/* ColorAllでの読み込みは禁止 */
    If pIdpColorType = idpColorAll Then GoTo errColorAll

    '/* 読み込み開始 */
    Call TheHdw.IDP.ReadFileEx(pSite, m_Name, pIdpColorType, pFileName, pSrcPC, pProcessMode)
    Set ReadFileEx = Me
    Exit Function

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
errColorAll:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgNotSupported("ColorAll"))
End Function

Public Function WriteFileEx( _
        ByVal pSite As Long, ByVal pFileName As Variant, _
        Optional ByVal pFileType As IdpFileFormat = idpFileBinary, _
        Optional ByRef pColor As Variant = EEE_COLOR_FLAT, _
        Optional ByVal pSrcPC As IdpPCType = IdpTesterPC, _
        Optional ByVal pMode As IdpWriteFileMode = idpWriteFileCurrPMD, _
        Optional ByVal pPixelCount As Variant = "", _
        Optional ByVal pProcessMode As IdpRWFileExProcessMode = idpRWFileExProcessBlocking _
) As CImgPlane
'内容:
'   プレーンのデータを画像ファイルに書き出す。従来のWriteFileの機能拡張版。
'
'[pSite]        In  書き出す対象のサイト番号
'[pFileName]    In  画像データのファイル名またはファイル名配列
'[pFileType]    In  画像データファイルのフォーマット指定
'[pColor]       In  プレーンのカラー指定
'[pSrcPC]       In  ファイル書き出し先のPC指定
'[pMode]        In  書き出す対象のPMD指定
'[pPixelCount]  In  書き出す対象がピクセル指定の場合に必要なピクセル数指定
'[pProcessMode] In  WriteFileExの動作モード指定
'
'戻り値:
'   自分自身の参照を返す
'
'注意事項:
'   IG-XL制約事項
'   【ファイル指定制約】:エラー検出はIG-XL APIに委ねる
'   ①サイト指定が-1の場合、指定ファイル名はサイト分の要素数の文字列配列が指定されていること
'   ②サイト指定が-1の場合、サイトで指定ファイル名が重複していないこと
'
'   【カラー指定制約】:Eee-JOBで検知
'   ③出力プレーンの色指定がColorAllでないこと
'
'備考:
'   サンプルコード
'    Dim dst As CImgPlane
'
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.SetPMD("ZONE2D")
'    Call dst.WriteFileEx(0, "Tmp.txt")
'
    Const FUNCTION_NAME = "WriteFileEx"

    On Error GoTo ErrHandler
    '/* カラータイプ型変換 */
    Dim pIdpColorType As IdpColorType
    pIdpColorType = GetIdpColorType(m_PlaneMap, pColor)
    '/* ColorAllでの書き出しは禁止 */
    If pIdpColorType = idpColorAll Then GoTo errColorAll

    '/* 書き出し開始 */
    Call TheHdw.IDP.WriteFileEx(pSite, m_Name, pIdpColorType, pFileName, pFileType, pSrcPC, pMode, pPixelCount, pProcessMode)
    Set WriteFileEx = Me
    Exit Function

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
errColorAll:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgNotSupported("ColorAll"))
End Function

Public Function BlurFilter( _
    ByRef pSrc As CImgPlane, ByVal pWidth As Long, _
    Optional ByVal pHeight As Long = 1, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR _
) As CImgPlane
'内容:
'   指定されたウィンドウの平均値を求め、その結果を結果プレーン上の注目ピクセルに格納する。
'
'[pSrc]     In  演算対象プレーン
'[pWidth]   In  ウィンドウのXサイズ
'[pHeight]  In  ウィンドウのYサイズ
'[pColorD]  In  保存先の色指定
'[pColorS]  In  データ元の色指定
'
'戻り値:
'   自分自身の参照を返す
'
'注意事項:
'   IG-XL制約事項
'   【入力プレーン制約】:Eee-JOBで検知
'   ①入力プレーンが自分自身でないこと
'
'   【ウィンドウサイズ制約】:Eee-JOBで検知
'   ②ウィンドウの高さ、幅ともに1以上63以下であること
'
'   【カラー指定制約】:エラー検出はIG-XL APIに委ねる
'   ③入力プレーンの色指定がColorAll、ColorFlatの場合、自分自身の色指定もそれぞれColorAll、ColorFlatであること
'   ④色指定がColorAllの場合は自分自身が入力プレーンの持つ色を全て含んでいること
'　注）④以外の場合、Color Map DefinitionでのidpColorTypeの割り当てによっては演算可能なケースが発生するが動作保障外なので注意
'   ⑤入力プレーンの色指定が単独色の場合、自分自身の色指定も単独色であること（指定する色は異なっても構わない）
'
'備考:
'   サンプルコード
'    Dim src As CImgPlane
'    Dim dst As CImgPlane
'
'    Set src = TheIDP.PlaneBank.Item("SOURCE")
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call src.SetPMD("ZONE2D")
'    Call dst.SetPMD("ZONE2D")
'    Call dst.Blur(src, 5, 5)
'
    Const FUNCTION_NAME = "BlurFilter"
    '/* 書き込み禁止のプレーンだったらエラー */
    If m_ReadOnly = True Then
        Set BlurFilter = Me
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx)
    End If

    '/* pSrcがNothingだったらエラー */
    If pSrc Is Nothing Then Call TheError.Raise(999, FUNCTION_NAME, ErrMsgObjctIsNothing("pSrc"))
    '/* pSrcが自分自身だったらエラー */
    If pSrc Is Me Then Call TheError.Raise(999, FUNCTION_NAME, ErrMsgSamePlane)

    '/* ウィンドウサイズは 1以上63以下 */
    If (pWidth > 63 Or pWidth < 1) Or (pHeight > 63 Or pHeight < 1) Then
        Call TheError.Raise(999, FUNCTION_NAME, "The window size must be 63 or less and 1 or more.")
    End If

    On Error GoTo ErrHandler
    '/* カラータイプ型変換 */
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    pIdpColorTypeD = GetIdpColorType(m_PlaneMap, pColorD)
    pIdpColorTypeS = GetSrcIdpColorType(pSrc.planeMap, pColorS, pColorD)

    '/* ブラーフィルタリング */
    Call TheHdw.IDP.Blur(pSrc.Name, pIdpColorTypeS, m_Name, pIdpColorTypeD, pWidth, pHeight)
    Set BlurFilter = Me
    Exit Function

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Function

Public Function TemporalSum( _
    ByRef pSrc() As CImgPlane, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "" _
) As CImgPlane
'内容:
'   複数のプレーン2~Nの各ピクセルごとに串刺し合計演算を行い、演算結果を呼び出し元のプレーンに返す。
'
'[pSrc()]   In  合計演算対象のプレーン配列
'[pColorD]  In  演算結果保存先のプレーン色指定
'[pColorS]  In  合計値演算対象プレーンの色指定
'[pFlgName] In  フラグ名
'
'戻り値:
'   自分自身の参照を返す
'
'注意事項:
'   IG-XL制約事項:以下のエラー検出はIG-XL APIに委ねる
'   【入力プレーン制約】
'   ①入力プレーン配列の要素数は2以上100以下であること
'   ②入力プレーン配列内にNothingが含まれていないこと
'   ③入力プレーンが全て同じBit長であること
'   ④入力プレーンが全て同じBasePMDであること
'   ⑤入力プレーンに自分自身が含まれていないこと
'
'   【結果格納プレーン制約】
'   ⑥自分自身のBit長が入力プレーンのBit長と同じか、それより深いこと
'　（例：入力が32Bit整数型であれば、出力はそれと同じか、32Bit浮動小数点型であることが必要）
'
'   【カラー指定制約】
'   ⑦入力プレーンの色指定がColorAll、ColorFlatの場合、自分自身の色指定もそれぞれColorAll、ColorFlatであること
'   ⑧色指定がColorAllの場合は自分自身が入力プレーンの持つ色を全て含んでいること
'　注）⑧以外の場合、Color Map DefinitionでのidpColorTypeの割り当てによっては演算可能なケースが発生するが動作保障外なので注意
'   ⑨入力プレーンの色指定が単独色の場合、自分自身の色指定も単独色であること（指定する色は異なっても構わない）
'
'備考:
'   サンプルコード
'   Dim src(2) As CImgPlane
'   Set src(0) = srcPlane1
'   Set src(1) = srcPlane2
'   Set src(2) = srcPlane3
'
'   Dim dst As CImgPlane
'   Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'   Call dst.TemporalSum(src)
'
    Const FUNCTION_NAME = "TemporalSum"
    '/* 書き込み禁止のプレーンだったらエラー */
    If m_ReadOnly = True Then
        Set TemporalSum = Me
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx)
    End If

    On Error GoTo ErrHandler
    '/* pSrc()の最小インデックス番号の要素だけチェック */
    '/* CImgPlane型でなかったらエラー（Nothing含む）*/
    '/* pSrc()の入力に、CImgPlane型でないオブジェクトを一度バリアントに突っ込んでCImgPlane型にキャストされた変数が渡されるとなぜか型の不一致が起こらないケースに対応 */
    If TypeName(pSrc(LBound(pSrc))) <> "CImgPlane" Then
        On Error GoTo 0
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pSrc()"))
    End If

    '/* カラータイプ型変換 */
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    pIdpColorTypeD = GetIdpColorType(m_PlaneMap, pColorD)
    '/* pSrc()は先頭のプレーンマップを参照する */
    pIdpColorTypeS = GetSrcIdpColorType(pSrc(LBound(pSrc)).planeMap, pColorS, pColorD)

    '/* 串刺し演算 */
    Call TheHdw.IDP.TemporalStatistics(GetArrayPlaneNames(pSrc), pIdpColorTypeS, m_Name, pIdpColorTypeD, idpTemporalSum, GetOptFlagName(pFlgName), GetOptFlagBit(pFlgName))
    Set TemporalSum = Me
    Exit Function

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Function

Public Function TemporalAverage( _
    ByRef pSrc() As CImgPlane, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "" _
) As CImgPlane
'内容:
'   複数のプレーン2~Nの各ピクセルごとに串刺し平均演算を行い、演算結果を呼び出し元のプレーンに返す。
'
'[pSrc()]   In  平均値演算対象のプレーン配列
'[pColorD]  In  演算結果保存先のプレーン色指定
'[pColorS]  In  平均値演算対象プレーンの色指定
'[pFlgName] In  フラグ名
'
'戻り値:
'   自分自身の参照を返す
'
'注意事項:
'   IG-XL制約事項:以下のエラー検出はIG-XL APIに委ねる
'   【入力プレーン制約】
'   ①入力プレーン配列の要素数は2以上100以下であること
'   ②入力プレーン配列内にNothingが含まれていないこと
'   ③入力プレーンが全て同じBit長であること
'   ④入力プレーンが全て同じBasePMDであること
'   ⑤入力プレーンに自分自身が含まれていないこと
'
'   【結果格納プレーン制約】
'   ⑥自分自身のBit長が入力プレーンのBit長と同じか、それより深いこと
'　（例：入力が32Bit整数型であれば、出力はそれと同じか、32Bit浮動小数点型であることが必要）
'
'   【カラー指定制約】
'   ⑦入力プレーンの色指定がColorAll、ColorFlatの場合、自分自身の色指定もそれぞれColorAll、ColorFlatであること
'   ⑧色指定がColorAllの場合は自分自身が入力プレーンの持つ色を全て含んでいること
'　注）⑧以外の場合、Color Map DefinitionでのidpColorTypeの割り当てによっては演算可能なケースが発生するが動作保障外なので注意
'   ⑨入力プレーンの色指定が単独色の場合、自分自身の色指定も単独色であること（指定する色は異なっても構わない）
'
'備考:
'   サンプルコード
'   Dim src(2) As CImgPlane
'   Set src(0) = srcPlane1
'   Set src(1) = srcPlane2
'   Set src(2) = srcPlane3
'
'   Dim dst As CImgPlane
'   Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'   Call dst.TemporalAverage(src)
'
    Const FUNCTION_NAME = "TemporalAverage"
    '/* 書き込み禁止のプレーンだったらエラー */
    If m_ReadOnly = True Then
        Set TemporalAverage = Me
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx)
    End If

    On Error GoTo ErrHandler
    '/* pSrc()の最小インデックス番号の要素だけチェック */
    '/* CImgPlane型でなかったらエラー（Nothing含む）*/
    '/* pSrc()の入力に、CImgPlane型でないオブジェクトを一度バリアントに突っ込んでCImgPlane型にキャストされた変数が渡されるとなぜか型の不一致が起こらないケースに対応 */
    If TypeName(pSrc(LBound(pSrc))) <> "CImgPlane" Then
        On Error GoTo 0
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pSrc()"))
    End If

    '/* カラータイプ型変換 */
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    pIdpColorTypeD = GetIdpColorType(m_PlaneMap, pColorD)
    '/* pSrc()は先頭のプレーンマップを参照する */
    pIdpColorTypeS = GetSrcIdpColorType(pSrc(LBound(pSrc)).planeMap, pColorS, pColorD)

    '/* 串刺し演算 */
    Call TheHdw.IDP.TemporalStatistics(GetArrayPlaneNames(pSrc), pIdpColorTypeS, m_Name, pIdpColorTypeD, idpTemporalMean, GetOptFlagName(pFlgName), GetOptFlagBit(pFlgName))
    Set TemporalAverage = Me
    Exit Function

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Function

Public Function TemporalStdDev( _
    ByRef pSrc() As CImgPlane, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlgName As String = "" _
) As CImgPlane
'内容:
'   複数のプレーン2~Nの各ピクセルごとに串刺し標準偏差演算を行い、演算結果を呼び出し元のプレーンに返す。
'
'[pSrc()]   In  標準偏差演算対象のプレーン配列
'[pColorD]  In  演算結果保存先のプレーン色指定
'[pColorS]  In  標準偏差演算対象プレーンの色指定
'[pFlgName] In  フラグ名
'
'戻り値:
'   自分自身の参照を返す
'
'注意事項:
'   IG-XL制約事項:以下のエラー検出はIG-XL APIに委ねる
'   【入力プレーン制約】
'   ①入力プレーン配列の要素数は2以上100以下であること
'   ②入力プレーン配列内にNothingが含まれていないこと
'   ③入力プレーンが全て同じBit長であること
'   ④入力プレーンが全て同じBasePMDであること
'   ⑤入力プレーンに自分自身が含まれていないこと
'
'   【結果格納プレーン制約】
'   ⑥自分自身のBit長が入力プレーンのBit長と同じか、それより深いこと
'　（例：入力が32Bit整数型であれば、出力はそれと同じか、32Bit浮動小数点型であることが必要）
'
'   【カラー指定制約】
'   ⑦入力プレーンの色指定がColorAll、ColorFlatの場合、自分自身の色指定もそれぞれColorAll、ColorFlatであること
'   ⑧色指定がColorAllの場合は自分自身が入力プレーンの持つ色を全て含んでいること
'　注）⑧以外の場合、Color Map DefinitionでのidpColorTypeの割り当てによっては演算可能なケースが発生するが動作保障外なので注意
'   ⑨入力プレーンの色指定が単独色の場合、自分自身の色指定も単独色であること（指定する色は異なっても構わない）
'
'備考:
'   サンプルコード
'   Dim src(2) As CImgPlane
'   Set src(0) = srcPlane1
'   Set src(1) = srcPlane2
'   Set src(2) = srcPlane3
'
'   Dim dst As CImgPlane
'   Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'   Call dst.TemporalStdDev(src)
'
    Const FUNCTION_NAME = "TemporalStdDev"
    '/* 書き込み禁止のプレーンだったらエラー */
    If m_ReadOnly = True Then
        Set TemporalStdDev = Me
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx)
    End If

    On Error GoTo ErrHandler
    '/* pSrc()の最小インデックス番号の要素だけチェック */
    '/* CImgPlane型でなかったらエラー（Nothing含む）*/
    '/* pSrc()の入力に、CImgPlane型でないオブジェクトを一度バリアントに突っ込んでCImgPlane型にキャストされた変数が渡されるとなぜか型の不一致が起こらないケースに対応 */
    If TypeName(pSrc(LBound(pSrc))) <> "CImgPlane" Then
        On Error GoTo 0
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pSrc()"))
    End If

    '/* カラータイプ型変換 */
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    pIdpColorTypeD = GetIdpColorType(m_PlaneMap, pColorD)
    '/* pSrc()は先頭のプレーンマップを参照する */
    pIdpColorTypeS = GetSrcIdpColorType(pSrc(LBound(pSrc)).planeMap, pColorS, pColorD)

    '/* 串刺し演算 */
    Call TheHdw.IDP.TemporalStatistics(GetArrayPlaneNames(pSrc), pIdpColorTypeS, m_Name, pIdpColorTypeD, idpTemporalStdDev, GetOptFlagName(pFlgName), GetOptFlagBit(pFlgName))
    Set TemporalStdDev = Me
    Exit Function

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Function

Public Function TemporalMin( _
    ByRef pSrc() As CImgPlane, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlagPlane As Variant = "", Optional ByRef pFlgName As String = "" _
) As CImgPlane
'内容:
'   複数のプレーン2~Nの各ピクセルから最小値を持つピクセルを抽出し、ピクセルの値を呼び出し元のプレーンに返す。
'
'[pSrc()]       In  最小値抽出対象のプレーン配列
'[pColorD]      In  抽出結果保存先のプレーン色指定
'[pColorS]      In  最小値抽出対象プレーンの色指定
'[pFlgPlane]    In  出力フラグプレーン
'[pFlgName]     In  フラグ名
'
'戻り値:
'   自分自身の参照を返す
'
'注意事項:
'   IG-XL制約事項:以下のエラー検出はIG-XL APIに委ねる
'   【入力プレーン制約】
'   ①入力プレーン配列の要素数は2以上32以下であること
'   ②入力プレーン配列にNothingが含まれていないこと
'   ③入力プレーンが全て同じBit長であること
'   ④入力プレーンが全て同じBasePMDであること
'   ⑤入力プレーンに自分自身が含まれていないこと
'
'   【結果格納プレーン制約】
'   ⑥自分自身のBit長が入力プレーンのBit長と同じであること
'
'   【カラー指定制約】
'   ⑦入力プレーンの色指定がColorAll、ColorFlatの場合、自分自身の色指定もそれぞれColorAll、ColorFlatであること
'   ⑧色指定がColorAllの場合は自分自身が入力プレーンの持つ色を全て含んでいること
'　注）⑧以外の場合、Color Map DefinitionでのidpColorTypeの割り当てによっては演算可能なケースが発生するが動作保障外なので注意
'   ⑨入力プレーンの色指定が単独色の場合、自分自身の色指定も単独色であること（指定する色は異なっても構わない）
'
'   【出力フラグプレーン制約】
'   ⑩出力フラグプレーンを利用する場合は、入力プレーンが16枚以下の場合は1枚、17枚以上であれば2枚用意されていること
'   ⑪出力フラグプレーンを利用する場合は、Bit長が16Bit整数型であること
'   ⑫出力フラグプレーンを利用する場合は、入力プレーンに含まれていないこと
'   ⑬出力フラグプレーンを利用する場合は、自分自身が含まれていないこと
'   ⑭出力フラグプレーンを配列指定で利用する場合は、その配列が②～④を満たすこと
'
'備考:
'   サンプルコード
'   Dim src(2) As CImgPlane
'   Set src(0) = srcPlane1
'   Set src(1) = srcPlane2
'   Set src(2) = srcPlane3
'
'   Dim dst As CImgPlane
'   Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'   Call dst.TemporalMin(src)
'
    Const FUNCTION_NAME = "TemporalMin"
    '/* 書き込み禁止のプレーンだったらエラー */
    If m_ReadOnly = True Then
        Set TemporalMin = Me
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx)
    End If

    On Error GoTo ErrHandler
    '/* pSrc()の最小インデックス番号の要素だけチェック */
    '/* CImgPlane型でなかったらエラー（Nothing含む）*/
    '/* pSrc()の入力に、CImgPlane型でないオブジェクトを一度バリアントに突っ込んでCImgPlane型にキャストされた変数が渡されるとなぜか型の不一致が起こらないケースに対応 */
    If TypeName(pSrc(LBound(pSrc))) <> "CImgPlane" Then
        On Error GoTo 0
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pSrc()"))
    End If

    '/* カラータイプ型変換 */
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    pIdpColorTypeD = GetIdpColorType(m_PlaneMap, pColorD)
    '/* pSrc()は先頭のプレーンマップを参照する */
    pIdpColorTypeS = GetSrcIdpColorType(pSrc(LBound(pSrc)).planeMap, pColorS, pColorD)

    '/* 串刺し演算 */
    Call TheHdw.IDP.TemporalMinMax(GetArrayPlaneNames(pSrc), pIdpColorTypeS, m_Name, pIdpColorTypeD, idpTemporalMin, GetArrayOptPlaneNames(pFlagPlane), GetOptFlagName(pFlgName), GetOptFlagBit(pFlgName))
    Set TemporalMin = Me
    Exit Function

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Function

Public Function TemporalMax( _
    ByRef pSrc() As CImgPlane, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByRef pFlagPlane As Variant = "", Optional ByRef pFlgName As String = "" _
) As CImgPlane
'内容:
'   複数のプレーン2~Nの各ピクセルから最大値を持つピクセルを抽出し、ピクセルの値を呼び出し元のプレーンに返す。
'
'[pSrc()]       In  最大値抽出対象のプレーン配列
'[pColorD]      In  抽出結果保存先のプレーン色指定
'[pColorS]      In  最大値抽出対象プレーンの色指定
'[pFlgPlane]    In  出力フラグプレーン
'[pFlgName]     In  フラグ名
'
'戻り値:
'   自分自身の参照を返す
'
'注意事項:
'   IG-XL制約事項:以下のエラー検出はIG-XL APIに委ねる
'   【入力プレーン制約】
'   ①入力プレーン配列の要素数は2以上32以下であること
'   ②入力プレーン配列にNothingが含まれていないこと
'   ③入力プレーンが全て同じBit長であること
'   ④入力プレーンが全て同じBasePMDであること
'   ⑤入力プレーンに自分自身が含まれていないこと
'
'   【結果格納プレーン制約】
'   ⑥自分自身のBit長が入力プレーンのBit長と同じであること
'
'   【カラー指定制約】
'   ⑦入力プレーンの色指定がColorAll、ColorFlatの場合、自分自身の色指定もそれぞれColorAll、ColorFlatであること
'   ⑧色指定がColorAllの場合は自分自身が入力プレーンの持つ色を全て含んでいること
'　注）⑧以外の場合、Color Map DefinitionでのidpColorTypeの割り当てによっては演算可能なケースが発生するが動作保障外なので注意
'   ⑨入力プレーンの色指定が単独色の場合、自分自身の色指定も単独色であること（指定する色は異なっても構わない）
'
'   【出力フラグプレーン制約】
'   ⑩出力フラグプレーンを利用する場合は、入力プレーンが16枚以下の場合は1枚、17枚以上であれば2枚用意されていること
'   ⑪出力フラグプレーンを利用する場合は、Bit長が16Bit整数型であること
'   ⑫出力フラグプレーンを利用する場合は、入力プレーンに含まれていないこと
'   ⑬出力フラグプレーンを利用する場合は、自分自身が含まれていないこと
'   ⑭出力フラグプレーンを配列指定で利用する場合は、その配列が②～④を満たすこと
'
'備考:
'   サンプルコード
'   Dim src(2) As CImgPlane
'   Set src(0) = srcPlane1
'   Set src(1) = srcPlane2
'   Set src(2) = srcPlane3
'
'   Dim dst As CImgPlane
'   Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'   Call dst.TemporalMax(src)
'
    Const FUNCTION_NAME = "TemporalMax"
    '/* 書き込み禁止のプレーンだったらエラー */
    If m_ReadOnly = True Then
        Set TemporalMax = Me
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx)
    End If

    On Error GoTo ErrHandler
    '/* pSrc()の最小インデックス番号の要素だけチェック */
    '/* CImgPlane型でなかったらエラー（Nothing含む）*/
    '/* pSrc()の入力に、CImgPlane型でないオブジェクトを一度バリアントに突っ込んでCImgPlane型にキャストされた変数が渡されるとなぜか型の不一致が起こらないケースに対応 */
    If TypeName(pSrc(LBound(pSrc))) <> "CImgPlane" Then
        On Error GoTo 0
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pSrc()"))
    End If

    '/* カラータイプ型変換 */
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    pIdpColorTypeD = GetIdpColorType(m_PlaneMap, pColorD)
    '/* pSrc()は先頭のプレーンマップを参照する */
    pIdpColorTypeS = GetSrcIdpColorType(pSrc(LBound(pSrc)).planeMap, pColorS, pColorD)

    '/* 串刺し演算 */
    Call TheHdw.IDP.TemporalMinMax(GetArrayPlaneNames(pSrc), pIdpColorTypeS, m_Name, pIdpColorTypeD, idpTemporalMax, GetArrayOptPlaneNames(pFlagPlane), GetOptFlagName(pFlgName), GetOptFlagBit(pFlgName))
    Set TemporalMax = Me
    Exit Function

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Function

Public Function TemporalCount( _
    ByRef pSrc() As CImgPlane, ByVal pCountType As IdpCountType, _
    ByVal pLoLim As Variant, ByVal pHiLim As Variant, ByVal pCountLimMode As IdpCountLimitMode, _
    Optional ByRef pColorD As Variant = EEE_COLOR_FLAT, Optional ByRef pColorS As Variant = SAME_COLOR, _
    Optional ByVal pLimitType As IdpLimitType = idpLimitExclude, _
    Optional ByRef pFlagPlane As Variant = "", Optional ByRef pFlgName As String = "" _
) As CImgPlane
'内容:
'   複数のプレーン2~Nの各ピクセルを通して指定した範囲内の値を持つピクセルの数をカウントし、カウントした値を呼び出し元のプレーンに返す。
'
'[pSrc()]           In  カウント対象のプレーン配列
'[pCountType]       In  カウント条件の指定
'[pLoLim]           In  Low側のスライス値
'[pHiLim]           In  High側のスライス値
'[pCountLimMode]    In  スライス値のモード指定
'[pColorD]          In  カウント結果保存先のプレーン色指定
'[pColorS]          In  カウント対象プレーンの色指定
'[pLimitType]       In  スライス値の境界指定
'[pFlgPlane]        In  出力フラグプレーン
'[pFlgName]         In  フラグ名
'
'戻り値:
'   自分自身の参照を返す
'
'注意事項:
'   IG-XL制約事項:以下のエラー検出はIG-XL APIに委ねる
'   【入力プレーン制約】
'   ①入力プレーン配列の要素数は2以上32以下であること
'   ②入力プレーン配列にNothingが含まれていないこと
'   ③入力プレーンが全て同じBit長であること
'   ④入力プレーンが全て同じBasePMDであること
'   ⑤入力プレーンに自分自身が含まれていないこと
'
'   【結果格納プレーン制約】
'   ⑥自分自身のBit長が入力プレーンのBit長と同じであること
'
'   【カラー指定制約】
'   ⑦入力プレーンの色指定がColorAll、ColorFlatの場合、自分自身の色指定もそれぞれColorAll、ColorFlatであること
'   ⑧色指定がColorAllの場合は自分自身が入力プレーンの持つ色を全て含んでいること
'　注）⑧以外の場合、Color Map DefinitionでのidpColorTypeの割り当てによっては演算可能なケースが発生するが動作保障外なので注意
'   ⑨入力プレーンの色指定が単独色の場合、自分自身の色指定も単独色であること（指定する色は異なっても構わない）
'
'   【出力フラグプレーン制約】
'   ⑩出力フラグプレーンを利用する場合は、入力プレーンが16枚以下の場合は1枚、17枚以上であれば2枚用意されていること
'   ⑪出力フラグプレーンを利用する場合は、Bit長が16Bit整数型であること
'   ⑫出力フラグプレーンを利用する場合は、入力プレーンに含まれていないこと
'   ⑬出力フラグプレーンを利用する場合は、自分自身が含まれていないこと
'   ⑭出力フラグプレーンを配列指定で利用する場合は、その配列が②～④を満たすこと
'
'  【スライス値指定制約】
'   ⑮スライス値のモード指定がidpLimitEachSiteの場合、Low側、High側共に要素数がサイト数と等しい1次元配列が入力されていること
'   ⑯スライス値のモード指定がidpLimitEachColorの場合、入力、出力プレーンの色指定が共にColorAllであること
'   ⑰スライス値のモード指定がidpLimitEachColorの場合、Low側、High側共に要素数11の1次元配列が入力されていること
'   ⑱スライス値のモード指定がidpLimitEachSiteColorの場合、入力、出力プレーンの色指定が共にColorAllであること
'   ⑲スライス値のモード指定がidpLimitEachSiteColorの場合、Low側、High側共に1次元目の要素数がサイト数と等しく、
'   　2次元目の要素数が11の2次元配列が入力されていること
'   ⑳スライス値のモード指定がidpLimitConstantの場合、Low側、High側共に固定値であること
'
'備考:
'   サンプルコード
'    Dim src(2) As CImgPlane
'    Set src(0) = srcPlane1
'    Set src(1) = srcPlane2
'    Set src(2) = srcPlane3
'
'    Dim dst As CImgPlane
'    Set dst = TheIDP.PlaneManager("vmcu").GetFreePlane(idpDepthS16)
'    Call dst.TemporalCount(src, idpCountAbove, 1, 1, idpLimitConstant)
'
    Const FUNCTION_NAME = "TemporalCount"
    '/* 書き込み禁止のプレーンだったらエラー */
    If m_ReadOnly = True Then
        Set TemporalCount = Me
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx)
    End If

    On Error GoTo ErrHandler
    '/* pSrc()の最小インデックス番号の要素だけチェック */
    '/* CImgPlane型でなかったらエラー（Nothing含む）*/
    '/* pSrc()の入力に、CImgPlane型でないオブジェクトを一度バリアントに突っ込んでCImgPlane型にキャストされた変数が渡されるとなぜか型の不一致が起こらないケースに対応 */
    If TypeName(pSrc(LBound(pSrc))) <> "CImgPlane" Then
        On Error GoTo 0
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pSrc()"))
    End If

    '/* カラータイプ型変換 */
    Dim pIdpColorTypeD As IdpColorType
    Dim pIdpColorTypeS As IdpColorType
    pIdpColorTypeD = GetIdpColorType(m_PlaneMap, pColorD)
    '/* pSrc()は先頭のプレーンマップを参照する */
    pIdpColorTypeS = GetSrcIdpColorType(pSrc(LBound(pSrc)).planeMap, pColorS, pColorD)

    '/* 串刺し演算 */
    Call TheHdw.IDP.TemporalCount(GetArrayPlaneNames(pSrc), pIdpColorTypeS, m_Name, pIdpColorTypeD, pCountType, pLoLim, pHiLim, pCountLimMode, pLimitType, GetArrayOptPlaneNames(pFlagPlane), GetOptFlagName(pFlgName), GetOptFlagBit(pFlgName))
    Set TemporalCount = Me
    Exit Function

ErrHandler:
    Call TheError.Raise(999, FUNCTION_NAME, Err.Description)
End Function

'/* 以下ユーティリティ関数 一部は既存の関数を見直し */
Private Function GetArrayPlaneNames(ByRef pSrc() As CImgPlane, Optional ByVal pCheckReadOnly As Boolean = False) As String()
'内容:
'   プレーンオブジェクト配列をプレーン名の文字列配列に変換する。
'
'[pSrc()]           In  プレーンオブジェクト配列
'[pCheckReadOnly]   In  プレーンの書き込み禁止プロパティチェックフラグ
'
'戻り値:
'   プレーン名の文字列配列
'
'概要:
'   ①配列の最小インデックス番号は問わない
'   ②配列要素内にNothingがあった場合は空白にする
'   ③書込み禁止チェックがONの時にプレーンが書込み禁止だったらエラーにする
'
'/* 2012/2/27 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetArrayPlaneNames"
    On Error GoTo errArrayMismatch
    Dim pNames() As String
    ReDim pNames(UBound(pSrc) - LBound(pSrc))
    Dim pIndex As Long
    Dim tmpName As String
    For pIndex = 0 To UBound(pNames)
        '/* 配列要素のNothingは無視する */
        If Not pSrc(pIndex + LBound(pSrc)) Is Nothing Then
            '/* 書き込み禁止プロパティチェック */
            If pCheckReadOnly Then
                If pSrc(pIndex + LBound(pSrc)).ReadOnly Then
                    tmpName = pSrc(pIndex + LBound(pSrc)).Name
                    GoTo errReadOnly
                End If
            End If
            pNames(pIndex) = pSrc(pIndex + LBound(pSrc)).Name
        End If
    Next pIndex
    GetArrayPlaneNames = pNames
    Exit Function
errArrayMismatch:
    Call TheError.Raise(Err.Number, FUNCTION_NAME, ErrMsgInvalidArrayType)
errReadOnly:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx(tmpName))
End Function

Private Function GetArrayOptPlaneNames(ByRef pOptSrc As Variant) As Variant
'内容:
'   オプショナルのプレーン情報（文字列、プレーンオブジェクト、プレーンオブジェクト配列）を文字列情報に変換する。
'
'[pOptSrc]           In  オプションプレーン指定
'
'戻り値:
'   空白、プレーン名の文字列、またはプレーン名の文字列配列
'
'概要:
'   ①入力は文字列、プレーンオブジェクト、プレーンオブジェクト配列以外はエラーにする
'   ②文字列入力は空白以外はエラーにする
'   ③プレーンが書込み禁止だったらエラーにする
'
'/* 2012/2/27 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetArrayOptPlaneNames"
    On Error GoTo ErrHandler
    '/* 入力フラグプレーンが文字列だった場合 */
    If TypeName(pOptSrc) = "String" Then
        If pOptSrc = "" Then
            GetArrayOptPlaneNames = ""
        '/* 文字列入力は空白以外は許可しない */
        Else
            GoTo errTypeMismatch
        End If
    '/* 入力フラグプレーンがCImgPlane型だった場合 */
    ElseIf TypeName(pOptSrc) = "CImgPlane" Then
        '/* 書込み禁止だったらエラー */
        If pOptSrc.ReadOnly Then GoTo errReadOnly
        GetArrayOptPlaneNames = pOptSrc.Name
    '/* 入力フラグプレーンがオブジェクト型配列だった場合 */
    ElseIf TypeName(pOptSrc) = "Object()" Then
        '/* CImgPlane型配列でない場合はキャストでエラー */
        Dim tmpSrc() As CImgPlane
        ReDim tmpSrc(UBound(pOptSrc) - LBound(pOptSrc))
        Dim pIndex As Long
        For pIndex = 0 To UBound(tmpSrc)
            Set tmpSrc(pIndex) = pOptSrc(LBound(pOptSrc) + pIndex)
        Next pIndex
        '/* 書込み禁止チェックをONにする */
        GetArrayOptPlaneNames = GetArrayPlaneNames(tmpSrc, True)
    '/* それ以外の変数型の入力は許可しない */
    Else
        GoTo errTypeMismatch
    End If
    Exit Function
ErrHandler:
    If Err.Number = 9 Then
        Call TheError.Raise(Err.Number, FUNCTION_NAME, ErrMsgInvalidArrayType)
    ElseIf Err.Number = 13 Then
        Call TheError.Raise(Err.Number, FUNCTION_NAME, ErrMsgTypeMismatch("pOptSrc"))
    Else
        Call TheError.Raise(Err.Number, FUNCTION_NAME, Err.Description)
    End If
errTypeMismatch:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pOptSrc"))
errReadOnly:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgReadOnlyForEx(pOptSrc.Name))
End Function

Private Function GetIdpColorType(ByRef pRefMap As CImgPlaneMap, ByRef pColor As Variant) As IdpColorType
'内容:
'   カラー情報（文字列、カラーオブジェクト）をIdpColorTypeに変換する。
'
'[pRefMap]  In  演算対象のプレーンが持つマップオブジェクト
'[pColor」  In  演算対象プレーンに指定するカラー
'
'戻り値:
'   idpColorType列挙型
'
'概要:
'   ①idpColorTypeの入力はエラーにする
'   ②単独色指定はpRefMapに存在しない色だとエラーになる
'
'/* 2012/2/28 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetIdpColorType"
    On Error GoTo ErrHandler
    Dim tmpColor As Variant
    '/* 色指定がCImgColor型だった場合 */
    If TypeName(pColor) = "CImgColor" Then
        tmpColor = pColor.Name
        GetIdpColorType = pRefMap.color(pColor.Name).IdpColor
    '/* 色指定が文字列だった場合 */
    ElseIf TypeName(pColor) = "String" Then
        '/* カラーオールの場合 */
        If pColor = EEE_COLOR_ALL Then
            GetIdpColorType = idpColorAll
        '/* カラーフラットの場合 */
        ElseIf pColor = EEE_COLOR_FLAT Then
            GetIdpColorType = idpColorFlat
        '/* 単独色指定の場合 */
        Else
            tmpColor = pColor
            GetIdpColorType = pRefMap.color(CStr(pColor)).IdpColor
        End If
    '/* それ以外の変数型の入力は許可しない */
    Else
        GoTo errTypeMismatch
    End If
    Exit Function
ErrHandler:
    If Err.Number = 9999 Then
        Call TheError.Raise(999, FUNCTION_NAME, ErrMsgNotExistColor(pRefMap.MapName, CStr(tmpColor)))
    Else
        Call TheError.Raise(Err.Number, FUNCTION_NAME, Err.Description)
    End If
errTypeMismatch:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgTypeMismatch("pColor"))
End Function

Private Function GetSrcIdpColorType(ByRef pRefMap As CImgPlaneMap, ByRef pColorS As Variant, ByRef pColorD As Variant) As IdpColorType
'内容:
'   GetIdpColorType関数の拡張。
'   pColorSにSAME_COLORが入力されるとpColorDに置き換えて変換する。
'
'[pRefMap]  In  演算対象のプレーンが持つマップオブジェクト
'[pColorS]  In  演算対象プレーンに指定するカラー
'[pColorD]  In  置き換え用のカラー
'
'戻り値:
'   idpColorType列挙型
'
'概要:
'   pColorS及びpColorDの型判定、単独色指定のカラーマップへの存在の有無などはGetIdpColorType関数に任せる
'
'/* 2012/2/28 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetSrcIdpColorType"
    On Error GoTo ErrHandler
    If TypeName(pColorS) = "String" Then
        '/* pColorSの色指定がSAME_COLORならpColorDに置き換える */
        If pColorS = SAME_COLOR Then
            GetSrcIdpColorType = GetIdpColorType(pRefMap, pColorD)
        Else
            GetSrcIdpColorType = GetIdpColorType(pRefMap, pColorS)
        End If
    Else
        GetSrcIdpColorType = GetIdpColorType(pRefMap, pColorS)
    End If
    Exit Function
ErrHandler:
    Call TheError.Raise(Err.Number, FUNCTION_NAME, Err.Description)
End Function

'#FlagExpansion
Private Function GetOptFlagName(ByVal pFlgName As String) As String
'内容:
'   このプレーンの共用フラグプレーン名を返す
'
'[pFlgName] In  フラグ名
'
'戻り値:
'   空白文字列または共用フラグプレーン名
'
'概要:
'   空白文字列が入力された場合はそのまま空白文字列を返す
'
'/* 2012/3/1 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetOptFlagName"
    On Error GoTo ErrHandler
    If pFlgName = "" Then
        GetOptFlagName = ""
    Else
        Dim FlagPlane As IImgFlag
        Set FlagPlane = GetSharedFlagPlane(pFlgName)
        
        If CurrentPMD.Name = Me.Manager.VariablePMDName Then
            FlagPlane.SetPMD CurrentPMD
        Else
            FlagPlane.SetPMD CurrentPMD.Name
        End If
        GetOptFlagName = FlagPlane.Name
    End If
    Exit Function
ErrHandler:
    Call TheError.Raise(Err.Number, FUNCTION_NAME, Err.Description)
End Function

'#FlagExpansion
Private Function GetOptFlagBit(ByVal pFlgName As String) As Long
'内容:
'   フラグ名が割り当てられているフラグビット番号を返す
'
'[pFlgName] In  フラグ名
'
'戻り値:
'   フラグビット番号
'
'概要:
'   空白文字列が入力された場合は0を返す
'
'/* 2012/3/1 全コード実行及び動作確認済み 0145206097 */
    Const FUNCTION_NAME = "GetOptFlagBit"
    On Error GoTo ErrHandler
    If pFlgName = "" Then
        GetOptFlagBit = 0
    Else
        GetOptFlagBit = Me.GetSharedFlagPlane(pFlgName).FlagBit(pFlgName)
        If GetOptFlagBit = 0 Then GoTo errFlagName
    End If
    Exit Function
ErrHandler:
    Call TheError.Raise(Err.Number, FUNCTION_NAME, Err.Description)
errFlagName:
    On Error GoTo 0
    Call TheError.Raise(999, FUNCTION_NAME, ErrMsgNotExistSharedFlag(pFlgName))
End Function

'/* 以下エラーメッセージ集 */
Private Function ErrMsgReadOnlyForEx(Optional ByVal pName As String = "Destination") As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgReadOnlyForEx = "[" & pName & "] is Read Only."
End Function

Private Function ErrMsgNotSupported(ByVal pType As String) As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgNotSupported = pType & " is not supported."
End Function

Private Function ErrMsgSamePlane() As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgSamePlane = "The source plane must be different from the destination plane."
End Function

Private Function ErrMsgObjctIsNothing(ByRef pType As String) As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgObjctIsNothing = pType & " is Nothing."
End Function

Private Function ErrMsgTypeMismatch(ByVal pType As String) As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgTypeMismatch = "Type of " & pType & " is mismatch."
End Function

Private Function ErrMsgInvalidArrayType() As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgInvalidArrayType = "Invalid array type."
End Function

Private Function ErrMsgNotExistColor(ByVal pMapName As String, ByVal pColorName As String) As String
'/* 2012/2/29 確認済み 0145206097 */
    ErrMsgNotExistColor = "The color [" & pColorName & "] is not included in plane map [" & pMapName & "]."
End Function

Private Function ErrMsgNotExistSharedFlag(ByVal pFlgName As String) As String
'/* 2012/3/1確認済み 0145206097 */
    ErrMsgNotExistSharedFlag = "Flag [" & pFlgName & "] does not exist."
End Function
#End If
'/* ======= 2012/01/18 Add 0145206097 ここまで ============== */
