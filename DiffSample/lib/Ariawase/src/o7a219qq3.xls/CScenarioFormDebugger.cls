VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CScenarioFormDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   フォームに対するパラメータダンプクラス
'
'目的:
'   フォームに情報を出力するライター
'   画像の場合はイメージビューワーで出力する
'
'作成者:
'   0145184306
'
Option Explicit

Implements IProcParamWriter
Implements IParameterViewer

Private Const ERR_NUMBER = 9999         ' エラー番号を保持する
Private Const CLASS_NAME = "CScenarioFormDebugger" ' Class名称を保持する
Private Const INDENT_SPACE = vbTab

Private m_ParamForm As ScenarioParameterViewer
Private m_ViewPlane As Collection
Private m_IGXLDebugDisplay As Object
'For Debug
Private Const DUMPKEYWORD_FOR_PARAM_NAME As String = "Parameter Name"
Private Const DUMPKEYWORD_FOR_CLASS_NAME As String = "Parameter Type"
Private Const DUMPKEYWORD_FOR_COMMENT As String = "Comment"
Private Const DUMPKEYWORD_FOR_NULL As String = ""
Private Const TAB_WIDTH = 4
Private m_IndentLevel As Long

Public Function AsIProcParamWriter() As IProcParamWriter
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IProcParamWriter型オブジェクト
'
'備考:
'
    Set AsIProcParamWriter = Me
End Function

Public Function AsIParameterViewer() As IParameterViewer
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IParameterViewer型オブジェクト
'
'備考:
'
    Set AsIParameterViewer = Me
End Function

Private Sub IParameterViewer_View()
'内容:
'   ダンプした情報をフォームに表示する
'   画像データがある場合はイメージデータビューワーで表示する
'
'備考:
'
'
    If m_ParamForm.EndStatus = True Then Exit Sub
    
    Dim vPlane As CImgPlane
    'Planeの表示
    If m_ViewPlane.Count > 0 Then
        For Each vPlane In m_ViewPlane
'            vPlane.View
            With theidv
                Call .OpenForm
                .PlaneNameGreen = vPlane.Name
                .PMDName = vPlane.CurrentPmdName
                Call .Refresh
            End With
        Next vPlane
    End If
    m_IGXLDebugDisplay.Show
    'Formの表示
    Call m_ParamForm.Display
    
    Call m_ParamForm.Hide
    Set m_ViewPlane = New Collection
    
End Sub

Private Sub IProcParamWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
'内容:
'   Boolean型情報をダンプする。
'
'引数:
'[dataName]  　IN String型:     データ名
'[blnData]　　 IN Boolean型:     データ
'
'備考:
'
'
    With Me.AsIProcParamWriter
        .WriteAsString dataName, blnData
    End With
End Sub

Private Sub IProcParamWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
'内容:
'   Double型情報をダンプする。
'
'引数:
'[dataName]  　IN String型:     データ名
'[dblData]　　 IN Double型:     データ
'
'備考:
'
'
    With Me.AsIProcParamWriter
        .WriteAsString dataName, dblData
    End With
End Sub

Private Sub IProcParamWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
'内容:
'   Integer型情報をダンプする。
'
'引数:
'[dataName]  　IN String型:     データ名
'[intData]　　 IN Integer型:     データ
'
'備考:
'
'
    With Me.AsIProcParamWriter
        .WriteAsString dataName, intData
    End With
End Sub

Private Sub IProcParamWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
'内容:
'   Long型情報をダンプする。
'
'引数:
'[dataName]  　IN String型:     データ名
'[lngData]　　 IN Long型:     データ
'
'備考:
'
'
    With Me.AsIProcParamWriter
        .WriteAsString dataName, lngData
    End With
End Sub

Private Sub IProcParamWriter_WriteAsPlane(ByVal pParamName As String, ByRef pImgPlane As CImgPlane)
'内容:
'   イメージプレーンの画像と情報をダンプする。
'
'[pParamName]  IN String型:     ヘッダーラベル
'[pParamPlane] IN CImgPlane型:  イメージプレーン
'
'備考:
'
'
    If Not pImgPlane Is Nothing Then
        With Me.AsIProcParamWriter
            .WriteAsString "Plane Name", pImgPlane.Name
            .WriteAsString "Plane Group", pImgPlane.planeGroup
            .WriteAsString "Bit Depth", idpDepthToStr(pImgPlane.BitDepth)
            .WriteAsString "Color Map Name", pImgPlane.planeMap.Name
            .WriteAsString "Color Map Definition", pImgPlane.planeMap.ColorNamesPmd
            .WriteAsString "Current PMD", pImgPlane.CurrentPmdName
            .WriteAsString "Read Only", pImgPlane.ReadOnly
            .WriteAsString "Plane Comment", pImgPlane.Comment
        End With
    
        m_ViewPlane.Add pImgPlane
    Else
        Me.AsIProcParamWriter.WriteAsString INDENT_SPACE & "Plane", "Nothing"
    End If

End Sub

Private Sub IProcParamWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
'内容:
'   テキスト情報をフォームにダンプする。
'
'[dataName]    IN String型:     データ名
'[strData]     IN String型:     データ
'
'備考:
'
'
    With m_ParamForm.ScenarioParamView
        Select Case dataName
            Case DUMPKEYWORD_FOR_PARAM_NAME
                    If .Text = "" Then
                        .Text = "[" & Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & "]" & " = " & strData
                    Else
                        .Text = .Text & vbCrLf & vbCrLf & "[" & Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & "]" & " = " & strData
                    End If
            Case DUMPKEYWORD_FOR_CLASS_NAME, DUMPKEYWORD_FOR_COMMENT
                    If .Text = "" Then
                        .Text = "[" & Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & "]" & " = " & strData
                    Else
                        .Text = .Text & vbCrLf & "[" & Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & "]" & " = " & strData
                    End If
            Case DUMPKEYWORD_FOR_NULL
                    If .Text = "" Then
                        .Text = Strings.Space$(m_IndentLevel * TAB_WIDTH) & strData
                    Else
                        .Text = .Text & vbCrLf & Strings.Space$(m_IndentLevel * TAB_WIDTH) & strData
                    End If
            Case Else
                    If .Text = "" Then
                        .Text = Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & " = " & strData
                    Else
                        .Text = .Text & vbCrLf & Strings.Space$(m_IndentLevel * TAB_WIDTH) & dataName & " = " & strData
                    End If
        End Select
    End With
End Sub

Private Property Get IParameterViewer_EndStatus() As Boolean
'内容:
'   フォームの終了ステータスを取得する
'
'戻り値:
'   強制終了フラグ(Boolean型)
'
'備考:
'
'
    IParameterViewer_EndStatus = m_ParamForm.EndStatus
End Property
Private Property Let IParameterViewer_EndStatus(ByVal pStatus As Boolean)
'内容:
'   フォームの終了ステータスを取得する
'
'戻り値:
'   強制終了フラグ(Boolean型)
'
'備考:
'
'
    m_ParamForm.EndStatus = pStatus
End Property
Private Property Get IProcParamWriter_MyName() As String
'内容:
'   自分の名前を取得する
'
'戻り値:
'   クラス名(String型)
'
'備考:
'
'
    IProcParamWriter_MyName = CLASS_NAME
End Property

Private Property Get IParameterViewer_MyName() As String
'内容:
'   自分の名前を取得する
'
'戻り値:
'   クラス名(String型)
'
'備考:
'
'
    IParameterViewer_MyName = CLASS_NAME
End Property

Private Sub Class_Initialize()
    Set m_ParamForm = ScenarioParameterViewer
    Set m_ViewPlane = New Collection
    Set m_IGXLDebugDisplay = CreateObject("CMU.IGXLDebugDisplay")
    m_IndentLevel = 0
End Sub

Private Sub Class_Terminate()
'    Unload m_ParamForm
    Set m_ParamForm = Nothing
    Set m_ViewPlane = Nothing
    Set m_IGXLDebugDisplay = Nothing
End Sub

Private Sub setHeader(ByVal pParamName As String)
    With m_ParamForm.ScenarioParamView
        If .Text = "" Then
            .Text = pParamName
        Else
            .Text = .Text & vbCrLf & pParamName
        End If
    End With
End Sub

Private Function composeString(ByRef strList() As String) As String

    Dim dataIndex As Long
    For dataIndex = 0 To UBound(strList)
        If dataIndex = 0 Then
            composeString = strList(dataIndex)
        Else
            composeString = composeString & "," & strList(dataIndex)
        End If
    Next dataIndex
End Function

Private Function idpDepthToStr(ByVal idpDepth As IdpBitDepth) As String
    Select Case idpDepth
        Case idpDepthF32
            idpDepthToStr = "idpDepthF32"
        Case idpDepthS16
            idpDepthToStr = "idpDepthS16"
        Case idpDepthS32
            idpDepthToStr = "idpDepthS32"
        Case Else
            idpDepthToStr = "-"
    End Select
End Function

Private Sub IProcParamWriter_WriteAsBlockEnd(ByVal pParamName As String)
    m_IndentLevel = m_IndentLevel - 1
End Sub

Private Sub IProcParamWriter_WriteAsBlockStart(ByVal pParamName As String)
    m_IndentLevel = m_IndentLevel + 1
End Sub

Public Sub IProcParamWriter_ReturnToRootBlock()
    m_IndentLevel = 0
End Sub
