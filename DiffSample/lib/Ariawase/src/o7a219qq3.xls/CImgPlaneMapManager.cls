VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPlaneMapManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ColorMapInfomationシートの情報を管理するクラス
'目的:
'   ColorMapInfomationシートの全情報を管理、CImgPlaneMapに振り分ける。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月21日
'

Option Explicit

Private Const ERR_NUMBER = 9999                         ' Error番号
Private Const CLASS_NAME = "CImgPlaneMapManager"   ' Class名称

Private Const SHT_LABEL_COLOR_MAP_NAME As String = "Color Map Name"
Private Const SHT_LABEL_COLOR_MAP As String = "Color Map"
Private Const SHT_LABEL_COLOR As String = "Color"
Private Const SHT_LABEL_COLOR_ARG_START As Integer = 1
Private Const SHT_LABEL_COLOR_ARG_END As Integer = 8
Private Const SHT_LABEL_COMMENT As String = "Comment"

Private mColorMaps As Collection       '
Private mIdpColorMapper As CIdpColorMapper  'idpColorType割り当てクラスを保持

'OK
Public Sub Init()
'内容:
'   ColorMapシートデータを取得して、CImgPlaneMapインスタンス群を生成します
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月21日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'

    '#### Color Map シート情報 ####
    Dim pColorMapName As String        'Color Map Name
    Dim pColorMap As CImgPlaneMap       'Color Map Instance
    Dim pComment As String
    Dim pChkNumColPreLine As Integer
    Dim pChkNumColNowLine As Integer
    Dim pMapData() As String
    Dim pTmpMapData(1 To 8, 1 To 9) As String

    Set mColorMaps = New Collection

    '### Color Map情報一時保存
    Dim pMapNames As Collection
    Set pMapNames = New Collection
    Dim pComments As Collection
    Set pComments = New Collection
    Dim pMapDataList As Collection
    Set pMapDataList = New Collection

    '### idpColorType割り振りクラス ###
    Set mIdpColorMapper = New CIdpColorMapper
    
    '#### 行のデータ確認カウンタ ####
    Dim intDataCnt As Integer
    Dim intGroupRowCnt As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ERROR_PLANEMAP
    '#####  SheetReaderを利用して、ColorMapInfoシートを読み込む  #####
    Dim clsWrkShtRdr As CWorkSheetReader
    Set clsWrkShtRdr = GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeColorMap)
    
    Dim strSheetName As String
    strSheetName = GetWkShtReaderManagerInstance.GetActiveSheetName(eSheetType.shtTypeColorMap)

    Dim IFileStream As IFileStream
    Dim IParamReader As IParameterReader

    With clsWrkShtRdr
        Set IFileStream = .AsIFileStream
        Set IParamReader = .AsIParameterReader
    End With

    Set clsWrkShtRdr = Nothing

    '#####  ColorMapInfoシートの情報を読み出し＆格納する  #####
    Do While Not IFileStream.IsEOR
    
        '#####  ColorMapTable 名称を取得し、配列情報をCImgPlaneMapに渡す  #####
        If IParamReader.ReadAsString(SHT_LABEL_COLOR_MAP_NAME) <> "" Then
            '#### 前行までのテーブル情報をコレクションに格納 ####
            If Not (pColorMap Is Nothing) Then
                '#### 前の行のColor Map Name終端処理 ####
                If intGroupRowCnt > 8 Then Call ErrorRowOver8(pColorMapName, intGroupRowCnt)
                
                ReDim pMapData(1 To pChkNumColPreLine, 1 To intGroupRowCnt)
                For i = 1 To UBound(pMapData, 2)
                    For j = 1 To UBound(pMapData, 1)
                        pMapData(j, i) = pTmpMapData(j, i)
                    Next j
                Next i

                'Map情報を一時的に保持する
                Call pMapDataList.Add(pMapData, pColorMapName)
                Call pMapNames.Add(pColorMapName)
                Call pComments.Add(pComment, pColorMapName)

                'idpColorType割り振りクラスにMapデータ追加
                Call mIdpColorMapper.AddMap(pMapData, pColorMapName)

'                Call pColorMap.Create(pColorMapName, pMapData, pComment)
                Call mColorMaps.Add(pColorMap, pColorMapName)
                
                Erase pTmpMapData
                Set pColorMap = Nothing
            End If
            
            '##### 現在の行のテーブル名取得 #####
            intGroupRowCnt = 0
            pChkNumColPreLine = 0
            pChkNumColNowLine = 0
            pColorMapName = ""
            
            Set pColorMap = New CImgPlaneMap
            pColorMapName = IParamReader.ReadAsString(SHT_LABEL_COLOR_MAP_NAME)
            pComment = IParamReader.ReadAsString(SHT_LABEL_COMMENT)
            If IsKey(pColorMapName, mColorMaps) Then       'Color Map Name重複チェック
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "This Color Map Name Already is Exist! " & pColorMapName)
            End If
        End If
        
        If intGroupRowCnt > 8 Then Call ErrorRowOver8(pColorMapName, intGroupRowCnt)
        
        Dim pStr As String
        For i = SHT_LABEL_COLOR_ARG_START To SHT_LABEL_COLOR_ARG_END
            pStr = IParamReader.ReadAsString(SHT_LABEL_COLOR & i & "@" & SHT_LABEL_COLOR_MAP)
            If pStr <> "" Then
                pTmpMapData(pChkNumColNowLine + 1, intGroupRowCnt + 1) = pStr
                pChkNumColNowLine = pChkNumColNowLine + 1
            End If
            
            '#### カラムの数のチェック ####
            If pStr = "" Or i = SHT_LABEL_COLOR_ARG_END Then
            
                '  カラムの数が前の行と違う               And １行目でない
                If pChkNumColPreLine <> pChkNumColNowLine And pChkNumColPreLine <> 0 Then
                    Call ErrorDifferentColumnNum(pColorMapName, pChkNumColPreLine, pChkNumColNowLine)    '横軸の情報の数に違いがある。
                End If
                pChkNumColPreLine = pChkNumColNowLine
                pChkNumColNowLine = 0
                Exit For
            End If
        Next i
        
        intDataCnt = intDataCnt + 1
        intGroupRowCnt = intGroupRowCnt + 1
        IFileStream.MoveNext
    Loop

    '##### 最後のカラーマップ定義の終端処理 #####
    If Not (pColorMap Is Nothing) Then
        If intGroupRowCnt > 8 Then Call ErrorRowOver8(pColorMapName, intGroupRowCnt)
        
        ReDim pMapData(1 To pChkNumColPreLine, 1 To intGroupRowCnt)
        For i = 1 To UBound(pMapData, 2)
            For j = 1 To UBound(pMapData, 1)
                pMapData(j, i) = pTmpMapData(j, i)
            Next j
        Next i

        'Map情報を一時的に保持する
        Call pMapDataList.Add(pMapData, pColorMapName)
        Call pMapNames.Add(pColorMapName)
        Call pComments.Add(pComment, pColorMapName)

        'idpColorType割り振りクラスにMapデータ追加
        Call mIdpColorMapper.AddMap(pMapData, pColorMapName)

'        Call pColorMap.Create(pColorMapName, pMapData, pComment)
        Call mColorMaps.Add(pColorMap, pColorMapName)
        
        Erase pTmpMapData
        Set pColorMap = Nothing
    End If
    
    Set IFileStream = Nothing
    Set IParamReader = Nothing

    '##### idpColorType割り振り計算実行 #####
    Call mIdpColorMapper.Execute

    '##### テーブル情報をインスタンスに格納 #####
    Dim mapIndex As Long
    For mapIndex = 1 To pMapDataList.Count
        Set pColorMap = mColorMaps(mapIndex)
        pMapData = pMapDataList(mapIndex)
        Call pColorMap.Create(pMapNames(mapIndex), pMapData, mIdpColorMapper, pComments(mapIndex))
    Next mapIndex
    Exit Sub
    
ERROR_PLANEMAP:
    Set mColorMaps = New Collection
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

Private Sub Class_Initialize()
'内容:
'   コンストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月25日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'

    Set mColorMaps = New Collection
'    Call Initialize
End Sub

Private Sub Class_Terminate()
'内容:
'   デストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月25日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'

    Set mColorMaps = Nothing
    Set mIdpColorMapper = Nothing
End Sub

'OK
Public Function PlaneMapList() As String()
'内容:
'   登録済みのColor Map 名を動的配列のStringで返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月25日
'パラメータ:
'   なし
'戻り値:
'   String
'
'注意事項:
'

    Dim pTmp As Variant
    Dim pArray() As String
    Dim PNum As Long

    ReDim pArray(mColorMaps.Count - 1)
    
    For Each pTmp In mColorMaps
        pArray(PNum) = CStr(pTmp.Name)
        PNum = PNum + 1
    Next pTmp

    PlaneMapList = pArray

End Function

'OK
Public Function planeMap(pColorMapName As String) As CImgPlaneMap
'内容:
'   登録済みのColor Mapを返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月25日
'パラメータ:
'   なし
'戻り値:
'   String
'
'注意事項:
'
    On Error GoTo NOTHING_COLOR_MAP
    Set planeMap = mColorMaps.Item(pColorMapName)
    On Error GoTo 0
    Exit Function
    
NOTHING_COLOR_MAP:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Map Name is Nothing!: " & pColorMapName & vbCrLf)
End Function

'OK
Public Property Get Count() As Long
'内容:
'   登録済みのColor Mapの数を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月25日
'パラメータ:
'   なし
'戻り値:
'   Long
'
'注意事項:
'
    Count = mColorMaps.Count
End Property

Public Function IsExist(pColorMapName As String) As Boolean
'内容:
'   対象のColorMapの有無を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   [pColorMapName]  In  1):
'戻り値:
'   Boolean
'
'注意事項:
'
    On Error GoTo NOT_FOUND
    Call mColorMaps.Item(pColorMapName)
    IsExist = True
    Exit Function
NOT_FOUND:
    IsExist = False
End Function

Private Sub ErrorRowOver8(pTableName As String, pOverNum As Integer)
'内容:
'   ColorMapの定義が
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月11日
'パラメータ:
'   [pTableName]    In  1):
'   [pOverNum]      In  2):
'戻り値:
'
'注意事項:

    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Map Row Data is Too Much!" & vbCrLf _
                                                & "Row Data Max 8,But Written: " & pOverNum & vbCrLf _
                                                & "Error Color Map Name: " & pTableName)
End Sub

Private Sub ErrorDifferentColumnNum(pTableName As String, pPreNum As Integer, pNowNum As Integer)
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月11日
'パラメータ:
'   [pTableName]    In  1):
'   [pPreNum]       In  2):
'   [pNowNum]       In  3):
'戻り値:
'
'注意事項:

    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "Color Map Name is differnet Collummn Number! " & pTableName & vbCrLf & _
                                                "Number Of Column: " & pPreNum & " and " & pNowNum)
End Sub

Private Function IsKey(ByVal pKey As String, ByVal pObj As Collection) As Boolean
'内容:
'   該当Collectionオブジェクトにキーが存在するか調べる
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月22日
'パラメータ:
'   なし
'戻り値:
'   Boolean :Trueすでにadd済み。存在します。       Falseまだ無し
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



