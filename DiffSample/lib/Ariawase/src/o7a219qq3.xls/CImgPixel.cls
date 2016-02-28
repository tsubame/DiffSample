VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPixel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   画素ひとつに相当する情報のクラス
'目的:
'   画素のアドレス情報を返します。Originは１
'作成者:
'  tomoyoshi.takase
'作成日: 2011年1月31日

Option Explicit

Implements IProcParamReplaceable

Private Const PARAMETER_REPLACER_X As String = "x"
Private Const PARAMETER_REPLACER_Y As String = "y"

Private mX As Long
Private mY As Long

Public Sub Init(pX As Long, pY As Long)
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月3日
'パラメータ:
'   [pX]    In/Out  1):
'   [pY]    In/Out  2):
'戻り値:
'
'注意事項:
'
'

    
    If mX > 0 Or mY > 0 Then
        Call TheError.Raise(9999, "", "CImgPixel is set already!")
    End If
    
    mX = pX
    mY = pY

End Sub

Public Function Clone() As CImgPixel
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月3日
'パラメータ:
'   なし
'戻り値:
'   CImgPixel
'
'注意事項:
'
'

    Dim p_Pixel As CImgPixel
    
    Set p_Pixel = New CImgPixel
    Call p_Pixel.Init(mX, mY)
    Set Clone = p_Pixel

End Function

Public Property Get x() As Long
    x = mX
End Property

Public Property Get y() As Long
    y = mY
End Property

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
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_X, CStr(mX), "Long", False, PARAMETER_REPLACER_X)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_Y, CStr(mY), "Long", False, PARAMETER_REPLACER_Y)
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
