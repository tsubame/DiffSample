VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIdpColorMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   それぞれの色名にidpColorTypeを割り振る
'
'目的:
'   登録されたMap情報を元に、使用するidpColorType数を少なく抑えつつも
'   同じMap内で番号が衝突しないようなidpColorTypeの割り振りを行う
'
'作成者:
'   0145184931
'
Option Explicit

Private Const ERR_NUMBER As Long = 9999                 ' エラー番号を保持する
Private Const CLASS_NAME As String = "CIdpColorMapper"  ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty     ' Default値"Empty"を保持する

Private Const ORIGIN_X As Integer = 1
Private Const ORIGIN_Y As Integer = 1

Private mColorNumbers As Collection         '色番号のリスト
Private mColorNames As Collection           '色名のリスト
Private mReferenceMapList As Collection     '各色名を使用するMapの番号のリスト (key1: 色名, key2: Map番号)
Private mMapNames As Collection             'Map名のリスト

Public Sub AddMap(ByRef pMapData() As String, ByVal pName As String)
'内容:
'   Map情報を追加する
'
'[pMapData()]  IN String型: Mapデータ(色名の2次元配列)を指定
'[pName]       IN String型: Map名を指定
'
'備考:
'   重複するMap名を登録することはできない
'
    '### Map名を登録 ###
    'Map番号を割り当てる(既に同じMap名が登録されている場合はエラー)
    Call mMapNames.Add(pName, pName)

    '### 色名登録/色名を使用するMap番号登録 ###
    'Mapから色名を全て取得して情報を登録する
    Dim myWidth As Integer
    myWidth = UBound(pMapData, 1)
    Dim myHeight As Integer
    myHeight = UBound(pMapData, 2)
    Dim mapIndexY As Integer
    For mapIndexY = ORIGIN_Y To myHeight
        Dim mapIndexX As Integer
        For mapIndexX = ORIGIN_X To myWidth
            Dim tmpColorName As String
            '現在の座標の色名を取得
            tmpColorName = pMapData(mapIndexX, mapIndexY)

            '色名未登録の場合のみ追加する
            If KeyExists(mColorNames, tmpColorName) = False Then
                Call mColorNames.Add(tmpColorName, tmpColorName)
                Call mReferenceMapList.Add(New Collection, tmpColorName)
            End If

            '色名を使用するMap名を登録
            On Error Resume Next
            Call mReferenceMapList(tmpColorName).Add(pName, pName)
            On Error GoTo 0
        Next mapIndexX
    Next mapIndexY

    Exit Sub
End Sub

Public Sub Execute()
'内容:
'   登録されたMap情報を元にしてそれぞれの色名にidpColorTypeを割り振る
'
'備考:
'   必要なMap情報を全て登録した後で実行すること
'
    Set mColorNumbers = New Collection

    '色番号未設定の色名リストを作成
    Dim bufUnregistered As Collection
    Set bufUnregistered = New Collection
    Dim tmpStr As Variant
    For Each tmpStr In mColorNames
        Call bufUnregistered.Add(tmpStr, tmpStr)
    Next tmpStr

    '色番号 (0から順に割り振る)
    Dim colorCounter As Integer
    colorCounter = 0

    '全ての色に色番号が割り振られるまでループ
    Do While bufUnregistered.Count > 0
        'Map名記録用Collection (色番号の重複チェック用)
        Dim tmpMapRegister As Collection
        Set tmpMapRegister = New Collection

        Dim tmpColorName As Variant
        For Each tmpColorName In bufUnregistered
            Dim tmpRefMaps As Collection
            '色名:tmpColorNameを使用するMap名のリストを取得
            Set tmpRefMaps = mReferenceMapList(tmpColorName)

            '同じMap内で別々の色に対して同じ色番号が割り振られないか確認
            If IsOverlapping(tmpMapRegister, tmpRefMaps) = False Then
                '色名に色番号を登録する
                Call mColorNumbers.Add(colorCounter, tmpColorName)
                '同じ色名を使用するMapを記録する
                Call RegisterMaps(tmpMapRegister, tmpRefMaps)

                '番号未設定リストから除く
                Call bufUnregistered.Remove(tmpColorName)
            End If
        Next tmpColorName
        colorCounter = colorCounter + 1
    Loop
End Sub

Public Property Get ColorNumber(ByVal pName As String) As Integer
'内容:
'   色名に割り振られたidpColorTypeを返す
'
'[pName]        IN String型:    idpColorTypeを取得する色名
'
'備考:
'   Execute実行後に使用すること
'
    ColorNumber = mColorNumbers(pName)
End Property

Private Sub Class_Initialize()
    Set mColorNames = New Collection
    Set mReferenceMapList = New Collection
    Set mMapNames = New Collection
End Sub

Private Sub Class_Terminate()
    Set mColorNumbers = Nothing
    Set mColorNames = Nothing
    Set mReferenceMapList = Nothing
    Set mMapNames = Nothing
End Sub

Private Sub RegisterMaps(ByRef pRegister As Collection, ByRef pMaps As Collection)
'内容:
'   リストのMap名をコレクションキーとして登録しておく
'
'[pRegister]    IN/OUT Collection型:    記録用コレクション
'[pMaps]        IN Collection型:        記録するMap名のリスト
'
'備考:
'
    Dim bufKey As Variant
    For Each bufKey In pMaps
        Call pRegister.Add(True, bufKey)
    Next bufKey
End Sub

Private Function IsOverlapping(ByRef pRegister As Collection, ByRef pMaps As Collection) As Boolean
'内容:
'   リストのMap名の中にコレクションキーとして登録済みのものがあるかどうかを返す
'
'[pRegister]    IN Collection型:    記録確認先コレクション
'[pMaps]        IN Collection型:    確認するMap名のコレクション
'
'備考:
'
    Dim bufKey As Variant
    For Each bufKey In pMaps
        If KeyExists(pRegister, bufKey) = True Then
            IsOverlapping = True
            Exit Function
        End If
    Next bufKey
    IsOverlapping = False
End Function

Private Function KeyExists(ByRef pCollection As Collection, ByVal pKey As String) As Boolean
'内容:
'   コレクションに登録済みのキーかどうかを返す
'
'[pCollection]  IN Collection型:    確認するコレクション
'[pKey]         IN String型:        確認するキー
'
'備考:
'
    On Error GoTo ErrHandler
    Call pCollection(pKey)
    KeyExists = True
    Exit Function
ErrHandler:
    KeyExists = False
End Function
