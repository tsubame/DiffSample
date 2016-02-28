VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgColorAllResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ColorAllResult型のラッパー
'
'目的:
'   画像処理関連の関数をラッピング
'
'作成者:
'
Option Explicit

Private m_PlaneMap As CImgPlaneMap
Private m_ColorAllResult As ColorAllResult
Private m_ColorList() As String

Private Enum COLOAR_ALL_RESULT_STATE
    UNKOWN
    Init
    SET_PARAMED
End Enum
Private m_State As COLOAR_ALL_RESULT_STATE

Private Sub Class_Initialize()
    Set m_PlaneMap = Nothing '初期状態Nothing
    m_State = Init
End Sub

Private Sub Class_Terminate()
    Set m_PlaneMap = Nothing
    m_State = UNKOWN
End Sub

Public Sub SetParam(ByRef srcPlane As CImgPlane, ByRef sRsult As ColorAllResult)
    
    If (m_State <> Init) Then
        Err.Raise 9999, "CImgColorAllResult", "SetParam has benn called!"
        Exit Sub
    End If
    
    m_ColorAllResult = sRsult
    Set m_PlaneMap = srcPlane.planeMap
    
    Dim Exist_C As Boolean
    Dim ColorList_C() As String
    Dim Exist_R As Boolean
    Dim ColorList_R() As String

    If (srcPlane.CurrentPMD.width >= srcPlane.planeMap.width) And (srcPlane.CurrentPMD.height >= srcPlane.planeMap.height) Then
        '' カラーマップの範囲を超える場合
        m_ColorList = m_PlaneMap.ColorList
    Else
        '' カラーマップの範囲を超えない場合
        Exist_C = GetColorListRow(srcPlane, ColorList_C)
        Exist_R = GetColorListCol(srcPlane, ColorList_R)
        m_ColorList = ColorListCombine(ColorList_C, ColorList_R, Exist_C, Exist_R)
    End If
    
    m_State = SET_PARAMED
End Sub

Public Property Get color(ByVal ColorName As String) As SiteValues

    If (m_State <> SET_PARAMED) Then
        Err.Raise 9999, "CImgColorAllResult", "SetParam must be called!"
        Exit Property
    End If

    If (ExistColorCheck(ColorName)) Then
        color = m_ColorAllResult.color(m_PlaneMap.color(ColorName).IdpColor)
    Else
        Err.Raise 9999, "CImgColorAllResult", "NotFondColor:" & "ColorName"
        Exit Property
    End If

End Property

Public Property Get ColorList() As String()

    If (m_State <> SET_PARAMED) Then
        Err.Raise 9999, "CImgColorAllResult", "SetParam must be called!"
        Exit Property
    End If

    ColorList = m_ColorList

End Property

Public Sub SetData(ByVal ColorName As String, ByVal lSite As Long, ByVal dblData As Double)

    m_ColorAllResult.color(m_PlaneMap.color(ColorName).IdpColor).SiteValue(lSite) = dblData
    
End Sub

Public Function Clone() As CImgColorAllResult

    Dim cloneobj As CImgColorAllResult
    Set cloneobj = New CImgColorAllResult
    Call cloneobj.SetCImgColor(m_ColorAllResult, m_PlaneMap, m_ColorList)
    Set Clone = cloneobj
    Set cloneobj = Nothing

End Function

Public Sub SetCImgColor(ByRef sRsult As ColorAllResult, ByRef planeMap As CImgPlaneMap, ByRef o_colorList() As String)

    m_ColorAllResult = sRsult
    Set m_PlaneMap = planeMap
    m_ColorList = o_colorList
    m_State = SET_PARAMED

End Sub


Public Function CreateFlat(ret() As Double)
    ReDim m_ColorList(0) As String
    Dim site As Integer
    m_ColorList(0) = "FLAT"
    Set m_PlaneMap = TheIDP.PlaneMapManager.planeMap("COLOR_FLAT")

    m_ColorAllResult.color(m_PlaneMap.color("FLAT").IdpColor).SiteValue = ret
    m_State = SET_PARAMED

End Function

Private Function ColorListCombine(in1ColorList() As String, in2ColorList() As String, in1_Exist As Boolean, in2_Exist As Boolean) As String()

    If (in1_Exist) Then ColorListCombine = in2ColorList
    If (in2_Exist) Then ColorListCombine = in1ColorList
    If (in1_Exist Or in2_Exist) Then Exit Function

    Dim result() As String
    ReDim result(0) As String

    Dim MatchColor As String
    Dim i As Integer, j As Integer

    For i = 0 To UBound(in2ColorList)
        MatchColor = ""
        For j = 0 To UBound(in1ColorList)
            If (in2ColorList(i) = in1ColorList(j)) Then MatchColor = in2ColorList(i)
        Next j

        If MatchColor <> "" Then
            If result(0) <> "" Then
                ReDim Preserve result(UBound(result) + 1) As String
            End If
            result(UBound(result)) = in2ColorList(i)
        End If
    Next i
    ColorListCombine = result
End Function


Private Function GetColorListRow(ByRef srcPlane As CImgPlane, result() As String) As Boolean
    '' すべて存在する場合は、フラグをTrueに設定
    GetColorListRow = True
    If srcPlane.CurrentPMD.width >= srcPlane.planeMap.width Then Exit Function

    Dim Add As Integer
    Dim RowPoint As Integer, ColPoint As Integer
    
    ReDim result(0) As String

    For RowPoint = 1 To m_PlaneMap.height
        For ColPoint = 1 To srcPlane.CurrentPMD.width
            Add = ((srcPlane.CurrentPMD.Left + ColPoint - 2) Mod m_PlaneMap.width) + 1
            If result(0) <> "" Then
                ReDim Preserve result(UBound(result) + 1) As String
            End If
            result(UBound(result)) = m_PlaneMap.ColorNameAddress(Add, RowPoint)
        Next ColPoint
    Next RowPoint

    GetColorListRow = False
End Function
Private Function GetColorListCol(ByRef srcPlane As CImgPlane, result() As String) As Boolean
    '' すべて存在する場合は、フラグをTrueに設定
    GetColorListCol = True
    If srcPlane.CurrentPMD.height >= srcPlane.planeMap.height Then Exit Function

    Dim Add As Integer
    Dim RowPoint As Integer, ColPoint As Integer
    
    ReDim result(0) As String

    For RowPoint = 1 To srcPlane.CurrentPMD.height
        For ColPoint = 1 To m_PlaneMap.width
            Add = ((srcPlane.CurrentPMD.Top + RowPoint - 2) Mod m_PlaneMap.height) + 1
            If result(0) <> "" Then
                ReDim Preserve result(UBound(result) + 1) As String
            End If
            result(UBound(result)) = m_PlaneMap.ColorNameAddress(ColPoint, Add)
        Next ColPoint
    Next RowPoint

    GetColorListCol = False
End Function


Private Function ExistColorCheck(ByVal ColorName As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(m_ColorList)
        If (m_ColorList(i) = ColorName) Then
            ExistColorCheck = True
            Exit Function
        End If
    Next i
    ExistColorCheck = False
End Function

