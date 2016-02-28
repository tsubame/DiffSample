VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   カラーマップの情報
'目的:
'   カラーマップ情報１つ分を管理します。
'   カラーマップに関するString情報を、idpColorTypeの情報に解釈します
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月22日
'   2011/02/02  CImgPlaneMapはIdpColorAllやIdpColorFlatは面倒見ない(CImgPlaneが見る)
'               それらの指定がされた場合はエラー。
'   2012/02/20  idpColorTypeを0番から順に割り振る仕組みから、CIdpColorMapperクラスで
'               予め割り振っておいたidpColorTypeを使用するように変更。

Option Explicit

Implements IProcParamReplaceable

Private Const ERR_NUMBER = 9999                         ' Error番号
Private Const CLASS_NAME = "CImgPlaneMap"           ' Class名称

Private Const ORIGIN_X As Integer = 1
Private Const ORIGIN_Y As Integer = 1

Private Const PARAMETER_REPLACER_MAPNAME As String = "MapName"
Private Const PARAMETER_REPLACER_COMMENT As String = "Comment"
Private Const PARAMETER_REPLACER_WIDTH As String = "Width"
Private Const PARAMETER_REPLACER_HEIGHT As String = "Height"
Private Const PARAMETER_REPLACER_COLORLIST As String = "ColorList"
Private Const PARAMETER_REPLACER_IDPCOLORSPMD As String = "IdpColorsPmd"

Private m_MapName As String                     'このクラスのColor Map Name
Private m_Comment As String

Private m_Width As Integer                  'このカラーマップの幅
Private m_Height As Integer                 'このカラーマップの高さ

'For PMDシートフォーマット
Private m_ColorNamesPmd As String                        'String PMDフォーマット。カンマ、カッコ区切り。
Private m_IdpColorsPmd As String                   'idpColorType番号 PMDフォーマット。カンマ、カッコ区切り。

'For idp <=> Str　Change
Private m_IdpStrs As Collection              'ユーザー定義のidpColorType情報をキーにStringを持つ
Private m_StrIdps As Collection              'ユーザー定義のString情報をキーにidpColorTypeを持つ

'For Colors
Private m_Colors As Collection              'ユーザー定義のString情報をキーにCImgColorを持つ

'For Width,Height,retArray
Private m_ArrayColorName() As String
Private m_ArrayIdpColorNumber() As IdpColorType

'For idpColorMapper
Private m_IdpColorMapper As CIdpColorMapper     'String -> idpColorType変換情報を持つ

'OK
Public Sub Create(ByVal pMapName As String, ByRef pMap() As String, ByRef pIdpColorMapper As CIdpColorMapper, Optional pComment As String = "")
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月25日
'パラメータ:
'   [pMapName]          In  1):Color Map Name
'   [pMap]              In  2):Color Map 定義文字列(ユーザー定義の文字状態で配列型)
'   [pIdpColorMapper]   In  3):色名->idpColorType変換情報
'   [pComment]          In  4):コメント
'戻り値:
'
'注意事項:
'

    'マップ情報作成
    m_MapName = pMapName
    m_Comment = pComment

    Set m_IdpColorMapper = pIdpColorMapper
    
    'idpColorTypeへ変換
    Call StrToIdpColor(pMap)
    
    'CImgColor,CImgPixel生成
    Dim p_Color As Variant
    Dim p_X As Long, p_Y As Long
    Dim p_ColorObj As CImgColor
    Dim p_PixelObj As CImgPixel
    Dim p_PixelObjs As Collection
    
    Set m_Colors = New Collection
    For Each p_Color In m_IdpStrs
        Set p_PixelObjs = New Collection
        For p_Y = ORIGIN_Y To m_Height
            For p_X = ORIGIN_X To m_Width
                If CStr(p_Color) = m_ArrayColorName(p_X, p_Y) Then
                    Set p_PixelObj = New CImgPixel
                    Call p_PixelObj.Init(p_X, p_Y)
                    Call p_PixelObjs.Add(p_PixelObj)
                    Set p_PixelObj = Nothing
                End If
            Next p_X
        Next p_Y
        Set p_ColorObj = New CImgColor
        Call p_ColorObj.Init(Me.Name, CStr(p_Color), CIdpColor(m_StrIdps.Item(CStr(p_Color))), p_PixelObjs)
        Call m_Colors.Add(p_ColorObj, CStr(p_Color))
        Set p_PixelObjs = Nothing
    Next p_Color
    
End Sub

'OK
Public Function Clone() As CImgPlaneMap
'内容:
'   このインスタンスのクローンを返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月25日
'パラメータ:
'   なし
'戻り値:
'   CImgPlaneMap
'
'注意事項:
'

    Dim pClone As CImgPlaneMap
    Dim pStr() As String
    Set pClone = New CImgPlaneMap
    
    With Me
        pStr = m_ArrayColorName
        Call pClone.Create(.Name, pStr, m_IdpColorMapper, .Comment)
    End With
    
    Set Clone = pClone
    
End Function

Public Property Get ColorList() As String()
'内容:
'   登録済みのColor Map名をString型の動的配列で返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年1月25日
'パラメータ:
'   [)] In  1):
'戻り値:
'   String()
'
'注意事項:
'
'

    Dim pTmp As Variant
    Dim pArray() As String
    Dim PNum As Long

    ReDim pArray(m_IdpStrs.Count - 1)
    
    For Each pTmp In m_IdpStrs
        pArray(PNum) = pTmp
        PNum = PNum + 1
    Next pTmp

    ColorList = pArray

End Property

'OK
Public Property Get Name() As String
'内容:
'   このインスタンスの名前(Color Map Definition に定義された名前)を取得
'
'備考:
'
    Name = m_MapName
End Property

'OK
Public Property Get width() As Long
'内容:
'   マップの幅を取得
'
'備考:
'
    width = m_Width
End Property

'OK
Public Property Get height() As Long
'内容:
'   マップの高さを取得
'
'備考:
'
    height = m_Height
End Property

'OK
Public Property Get Comment() As String
'内容:
'   マップのコメントを取得
'
'備考:
'
    Comment = m_Comment
End Property

'OK
Public Property Get IdpColorAddress(ByVal pX As Long, ByVal pY As Long) As IdpColorType
'内容:
'   指定アドレスの色情報を取得
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月22日
'パラメータ:
'   [pX]    In  1):     Xアドレス(1スタート)
'   [pY]    In  2):     Yアドレス(1スタート)
'戻り値:
'   idpColorType
'
'注意事項:
'

    If (pX <= 0) Or (pX > m_Width) Or (pY <= 0) Or (pY > m_Height) Then
        Call TheError.Raise(9999, "Color Map Infomation", "Invalid Parameter")
        Exit Property
    End If
    
    IdpColorAddress = m_ArrayIdpColorNumber(pX, pY)

End Property

'OK
Public Property Get ColorNameAddress(ByVal pX As Long, ByVal pY As Long) As String
'内容:
'   指定アドレスの色情報を取得
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月22日
'パラメータ:
'   指定アドレスの色情報を取得
'   [pX]    In  1):     Xアドレス(1スタート)
'   [pY]    In  2):     Yアドレス(1スタート)
'戻り値:
'   String
'
'注意事項:
'

    If (pX <= 0) Or (pX > m_Width) Or (pY <= 0) Or (pY > m_Height) Then
        Call TheError.Raise(9999, "Color Map Infomation", "Invalid Parameter")
        Exit Property
    End If
    
    ColorNameAddress = m_ArrayColorName(pX, pY)

End Property

'OK
Public Property Get IdpColorsPmd() As String
'内容:
'   マップ定義の文字列を返す（ユーザー定義idpColorType情報）
'
'備考:
'
    IdpColorsPmd = m_IdpColorsPmd
End Property

'OK
Public Property Get ColorNamesPmd() As String
'内容:
'   マップ定義の文字列を返す（ユーザー定義String情報）
'
'備考:
'
    ColorNamesPmd = m_ColorNamesPmd
End Property

Public Property Get ColorIdpType(p_IdpColorType As IdpColorType) As CImgColor
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月1日
'パラメータ:
'   [pIdpColor] In  1):
'戻り値:
'   CImgColor
'
'注意事項:
'

    On Error GoTo NOTHING_KEY
    Set ColorIdpType = m_Colors.Item(CStr(ColorName(p_IdpColorType))).Clone
    Exit Property
NOTHING_KEY:
    Call TheError.Raise(9999, "idp", "This idpColorType is no use! Color Map Definition: " & p_IdpColorType)

End Property

Public Property Get color(p_ColorName As String) As CImgColor
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月1日
'パラメータ:
'   [pColorName]    In  1):
'戻り値:
'   CImgColor
'
'注意事項:
'

    On Error GoTo NOTHING_KEY
    Set color = m_Colors.Item(p_ColorName)
    Exit Property

NOTHING_KEY:
    Call TheError.Raise(9999, "idp", "This Color name is no use! Color Map Definition: " & p_ColorName)

End Property

'OK
'Private Property Get IdpColorNumber(pColor As Variant) As IdpColorType
''内容:
''   該当名のidpColorTypeを返す（このクラスに割り振られたidpColorType情報）
''
''備考:
''
'    On Error GoTo NOTHING_KEY
'
'    If pColor = EEE_COLOR_ALL Then
'        IdpColorNumber = idpColorAll
'        Exit Property
'    ElseIf pColor = EEE_COLOR_FLAT Then
'        IdpColorNumber = idpColorFlat
'        Exit Property
'    End If
'
'    IdpColorNumber = m_StrIdps.Item(pColor)
'    Exit Property
'
'NOTHING_KEY:
'    Call TheError.Raise(9999, "idp", "This Color name is no use! Color Map Definition: " & pColor)
'
'End Property

'OK
Private Property Get ColorName(pIdpColorType As IdpColorType) As String
'内容:
'   該当idpColorTypeの文字列を返す（ユーザー定義String情報）
'
'備考:
'
    On Error GoTo NOTHING_KEY
    ColorName = m_IdpStrs.Item(CStr(pIdpColorType))
    Exit Property
NOTHING_KEY:
    Call TheError.Raise(9999, "idp", "This idpColorType is no use! Color Map Definition: " & pIdpColorType)

End Property

'OK
Public Property Get Count() As Integer
'内容:
'   登録されている色数を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月19日
'パラメータ:
'   なし
'戻り値:
'   Integer
'
'注意事項:
'
    Count = m_IdpStrs.Count
End Property

Public Function OddRow() As Collection
'内容:
'   奇数行にある色を全て返します。
'   CImgPixelのCollectionで返します。
'   対象、非対象の範囲で、どちらにも該当する場合、対象とします。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月4日
'パラメータ:
'   なし
'戻り値:
'   Collection
'
'注意事項:
'   カラーマップ定義における奇数行です。
'   PMD定義における奇数行とは意味が異なる場合があります。

    Dim p_Color As Variant
    Dim p_Pixel As Variant
    Dim p_ColorCollection As Collection
    
    Set p_ColorCollection = New Collection
    
    For Each p_Color In m_Colors
        For Each p_Pixel In p_Color.Pixels
            If (((p_Pixel.y - 1) Mod 2) + 1) = 1 Then
                Call p_ColorCollection.Add(p_Color)
                Exit For
            End If
        Next p_Pixel
    Next p_Color

    Set OddRow = p_ColorCollection

End Function

Public Function OddColumn() As Collection
'内容:
'   奇数列にある色を全て返します。
'   CImgPixelのCollectionで返します。
'   対象、非対象の範囲で、どちらにも該当する場合、対象とします。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月4日
'パラメータ:
'   なし
'戻り値:
'   Collection
'
'注意事項:
'   カラーマップ定義における奇数列です。
'   PMD定義における奇数列とは意味が異なる場合があります。

    Dim p_Color As Variant
    Dim p_Pixel As Variant
    Dim p_ColorCollection As Collection
    
    Set p_ColorCollection = New Collection
    
    For Each p_Color In m_Colors
        For Each p_Pixel In p_Color.Pixels
            If (((p_Pixel.x - 1) Mod 2) + 1) = 1 Then
                Call p_ColorCollection.Add(p_Color)
                Exit For
            End If
        Next p_Pixel
    Next p_Color

    Set OddColumn = p_ColorCollection

End Function

Public Function EvenRow() As Collection
'内容:
'   偶数行にある色を全て返します。
'   CImgPixelのCollectionで返します。
'   対象、非対象の範囲で、どちらにも該当する場合、対象とします。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月4日
'パラメータ:
'   なし
'戻り値:
'   Collection
'
'注意事項:
'   カラーマップ定義における偶数行です。
'   PMD定義における偶数行とは意味が異なる場合があります。

    Dim p_Color As Variant
    Dim p_Pixel As Variant
    Dim p_ColorCollection As Collection
    
    Set p_ColorCollection = New Collection
    
    For Each p_Color In m_Colors
        For Each p_Pixel In p_Color.Pixels
            If (((p_Pixel.y - 1) Mod 2) + 1) = 2 Then
                Call p_ColorCollection.Add(p_Color)
                Exit For
            End If
        Next p_Pixel
    Next p_Color

    Set EvenRow = p_ColorCollection

End Function

Public Function EvenColumn() As Collection
'内容:
'   偶数列にある色を全て返します。
'   CImgPixelのCollectionで返します。
'   対象、非対象の範囲で、どちらにも該当する場合、対象とします。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月4日
'パラメータ:
'   なし
'戻り値:
'   Collection
'
'注意事項:
'   カラーマップ定義における偶数列です。
'   PMD定義における偶数列とは意味が異なる場合があります。

    Dim p_Color As Variant
    Dim p_Pixel As Variant
    Dim p_ColorCollection As Collection
    
    Set p_ColorCollection = New Collection
    
    For Each p_Color In m_Colors
        For Each p_Pixel In p_Color.Pixels
            If (((p_Pixel.x - 1) Mod 2) + 1) = 2 Then
                Call p_ColorCollection.Add(p_Color)
                Exit For
            End If
        Next p_Pixel
    Next p_Color

    Set EvenColumn = p_ColorCollection

End Function

Private Sub StrToIdpColor(ByRef pMap() As String)
'内容:
'   ユーザーが定義したカラーマップのシート情報をidpColorTypeの定義に変換します。
'   カラー番号はCIdpColorMapperクラスの情報を使って割り振ります。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月22日
'パラメータ:
'   [pMap]   In  1):ユーザーが定義したマップのデータ領域(Rangeオブジェクト)
'戻り値:
'   なし
'
'注意事項:
'   最大８色です(IG-XL仕様)

    Dim i As Integer, j As Integer

    '同色の横方向で間隔違いエラーチェック用(IG-XL仕様、EeeJOBでエラー出力します)
    Dim pHorizontalColorMap(1 To 8) As Integer
    Dim colorCounter As Integer
    colorCounter = 0

    m_Width = UBound(pMap, 1)
    m_Height = UBound(pMap, 2)

    ReDim m_ArrayIdpColorNumber(ORIGIN_X To m_Width, ORIGIN_Y To m_Height)

    m_ArrayColorName = pMap          '動的配列に値渡し

    Set m_StrIdps = New Collection
    Set m_IdpStrs = New Collection

    For i = ORIGIN_Y To m_Height
        
        Erase pHorizontalColorMap
        For j = ORIGIN_X To m_Width
            Dim tmpStr As String
            tmpStr = pMap(j, i)
            Dim tmpIdpColor As Integer
            tmpIdpColor = m_IdpColorMapper.ColorNumber(tmpStr)  'idpColorType取得

            If Not IsKey(tmpStr, m_StrIdps) Then           'Key(色定義の名前)が作成されていなければ登録
                colorCounter = colorCounter + 1
                If colorCounter > 8 Then
                    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Name is Too Much!" & vbCrLf & _
                                                                "At most 8 colors" & vbCrLf & _
                                                                "Error Color Map Name: " & m_MapName)
                End If

                pHorizontalColorMap(j) = tmpIdpColor    '色の横方向間隔チェック用

                Call m_StrIdps.Add(tmpIdpColor, tmpStr)
                Call m_IdpStrs.Add(tmpStr, CStr(tmpIdpColor))
            Else
                pHorizontalColorMap(j) = tmpIdpColor    '色の横方向間隔チェック用
            End If
            m_ArrayIdpColorNumber(j, i) = tmpIdpColor
        Next j
        If j < 9 Then
            pHorizontalColorMap(j) = -1            '色の横方向間隔チェック用(番兵)
        End If
        Call CheckIntervalSameColor(pHorizontalColorMap)
    Next i

    'IG-XL PMD Map情報生成
    m_IdpColorsPmd = ""
    m_ColorNamesPmd = ""
    
    For i = ORIGIN_Y To m_Height
        m_IdpColorsPmd = m_IdpColorsPmd & "("
        m_ColorNamesPmd = m_ColorNamesPmd & "("
        
        For j = ORIGIN_X To m_Width
            m_IdpColorsPmd = m_IdpColorsPmd & CStr(m_ArrayIdpColorNumber(j, i))
            m_ColorNamesPmd = m_ColorNamesPmd & pMap(j, i)
            
            If j <> m_Width Then
                m_IdpColorsPmd = m_IdpColorsPmd & ","
                m_ColorNamesPmd = m_ColorNamesPmd & ","
            End If
        Next j
        
        m_IdpColorsPmd = m_IdpColorsPmd & ")"
        m_ColorNamesPmd = m_ColorNamesPmd & ")"
    Next i
    
End Sub

'OK
Private Function IsKey(ByVal pKey As String, ByVal pObj As Collection) As Boolean
'内容:
'   該当Collectionオブジェクトにキーが存在するか調べる
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月22日
'パラメータ:
'   [pKey]  In  1):
'   [pObj]  In  2):
'戻り値:
'   Boolean :Trueすでに存在します       Falseまだ無し
'
'注意事項:
'
    On Error GoTo ALREADY_REG
    Call pObj.Item(pKey)
    IsKey = True
    Exit Function

ALREADY_REG:
    IsKey = False

End Function

'OK
Private Function CIdpColor(pInt As Integer) As IdpColorType
    CIdpColor = pInt
End Function

Private Sub CheckIntervalSameColor(pHorizontalColorMap() As Integer)
'内容:
'   同色、横方向の間隔チェック。
'
'   配列は横方向アドレス、値はIdpColorType
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月22日
'パラメータ:
'   [pHorizontalColorMap]  In  1):横方向のidpColorType情報
'戻り値:
'
'注意事項:

Dim i As Integer
Dim pSameColorIndex As Collection
Dim pSameColorAddress(11) As Integer
Dim pSameColorInterval(11) As Integer

    Set pSameColorIndex = New Collection
    Erase pSameColorAddress
    Erase pSameColorInterval
    
    For i = 1 To 8
        If pHorizontalColorMap(i) = -1 Then Exit For       '番兵チェック

        '色ごとに0番から番号を割り振る(間隔チェック用)
        On Error Resume Next
        Call pSameColorIndex.Add(pSameColorIndex.Count, CStr(pHorizontalColorMap(i)))
        On Error GoTo 0
        Dim tmpColorIndex As Integer
        tmpColorIndex = pSameColorIndex(CStr(pHorizontalColorMap(i)))
        
        If pSameColorAddress(tmpColorIndex) = 0 Then
            '対応色にアドレス代入
            pSameColorAddress(tmpColorIndex) = i
        Else
            If pSameColorInterval(tmpColorIndex) = 0 Then
                '対応色に同色の間隔を代入
                pSameColorInterval(tmpColorIndex) = i - pSameColorAddress(tmpColorIndex)
                pSameColorAddress(tmpColorIndex) = i
            Else
                '対応色の間隔チェック
                If pSameColorInterval(tmpColorIndex) <> (i - pSameColorAddress(tmpColorIndex)) Then
                    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Map Definition Sheet Error!" & vbCrLf & _
                                                                "Interval of same color is different!" & vbCrLf & vbCrLf & _
                                                                "Color Map Name: " & m_MapName & vbCrLf & _
                                                                "Color Name: " & m_IdpStrs.Item(CStr(pHorizontalColorMap(i))) & vbCrLf & _
                                                                "Error Intervals of number: " & pSameColorInterval(tmpColorIndex) & "," & (i - pSameColorAddress(tmpColorIndex)))
                End If
                pSameColorAddress(tmpColorIndex) = i
            End If
        End If
    Next i

End Sub

''OK
'Private Sub CheckIntervalSameColor(pHorizontalColorMap() As Integer)
''内容:
''   同色、横方向の間隔チェック。
''
''   配列は横方向アドレス、値はIdpColorType
''作成者:
''  tomoyoshi.takase
''作成日: 2010年11月22日
''パラメータ:
''   [pHorizontalColorMap]  In  1):横方向のidpColorType情報
''戻り値:
''
''注意事項:
'
'Dim i As Integer
'Dim pSameColorAddress(11) As Integer
'Dim pSameColorInterval(11) As Integer
'
'    Erase pSameColorAddress
'    Erase pSameColorInterval
'
'    For i = 1 To 8
'        If pHorizontalColorMap(i) = -1 Then Exit For       '番兵チェック
'
'        If pSameColorAddress(pHorizontalColorMap(i)) = 0 Then
'            '対応色にアドレス代入
'            pSameColorAddress(pHorizontalColorMap(i)) = i
'        Else
'            If pSameColorInterval(pHorizontalColorMap(i)) = 0 Then
'                '対応色に同色の間隔を代入
'                pSameColorInterval(pHorizontalColorMap(i)) = i - pSameColorAddress(pHorizontalColorMap(i))
'                pSameColorAddress(pHorizontalColorMap(i)) = i
'            Else
'                '対応色の間隔チェック
'                If pSameColorInterval(pHorizontalColorMap(i)) <> (i - pSameColorAddress(pHorizontalColorMap(i))) Then
'                    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Map Definition Sheet Error!" & vbCrLf & _
'                                                                "Interval of same color is different!" & vbCrLf & vbCrLf & _
'                                                                "Color Map Name: " & m_MapName & vbCrLf & _
'                                                                "Color Name: " & m_IdpStrs.Item(CStr(pHorizontalColorMap(i))) & vbCrLf & _
'                                                                "Error Intervals of number: " & pSameColorInterval(pHorizontalColorMap(i)) & "," & (i - pSameColorAddress(pHorizontalColorMap(i))))
'                End If
'                pSameColorAddress(pHorizontalColorMap(i)) = i
'            End If
'        End If
'    Next i
'
'End Sub


'OK
Private Sub Class_Initialize()
'内容:
'   コンストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月26日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'

    Set m_IdpStrs = New Collection
    Set m_StrIdps = New Collection

End Sub

'OK
Private Sub Class_Terminate()
'内容:
'   デストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月26日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'

    Set m_IdpStrs = Nothing
    Set m_StrIdps = Nothing
    Set m_Colors = Nothing
    Set m_IdpColorMapper = Nothing
End Sub

Public Property Get MapName() As String
'内容:
'   このクラスのColor Map Nameを戻す
'
'パラメータ:
'
'戻り値:
'   このクラスのColor Map Nameを戻す
'
'注意事項:
'
    MapName = m_MapName
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
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_MAPNAME, m_MapName, "String", False, PARAMETER_REPLACER_MAPNAME)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COMMENT, m_Comment, "String", False, PARAMETER_REPLACER_COMMENT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_WIDTH, CStr(m_Width), "Integer", False, PARAMETER_REPLACER_WIDTH)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_HEIGHT, CStr(m_Height), "Integer", False, PARAMETER_REPLACER_HEIGHT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COLORLIST, m_ColorNamesPmd, "String", False, PARAMETER_REPLACER_COLORLIST)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_IDPCOLORSPMD, m_IdpColorsPmd, "String", False, PARAMETER_REPLACER_IDPCOLORSPMD)
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
