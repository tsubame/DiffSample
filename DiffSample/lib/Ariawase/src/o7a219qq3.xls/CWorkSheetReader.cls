VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   汎用ワークシートReader
'
'目的:
'   対象ワークシートを読込む
'   最初にキャッシュを取り、以降キャッシュを利用
'
'作成者:
'   a_oshima
'
'前提条件:
'   グローバルオブジェクト TheError:CErrManager が存在すること
'
'利用手順:
'   ①インスタンス生成後、Initializeメソッドをコール。基点セルは初期位置に設定される
'       ※SetLocationによってPrimaryKeyをサーチして基点セルを設定することも可能
'   ②Do While でEndOfRowを超えない間、
'       ②-1 IparameterReaderのReadAsメソッドで必要な列の値を読込み
'       ②-2 MoveNextで次の行に進む
'   ③Do While終わり
'

Option Explicit

Implements IParameterReader
Implements IFileStream

Private WithEvents mEventWatcher As CEventWatcher      '読み出し対象ワークシート名
Attribute mEventWatcher.VB_VarHelpID = -1
Private mTargetWorkSheet As Worksheet      '読み出し対象ワークシート名
Attribute mTargetWorkSheet.VB_VarHelpID = -1
Private mSheetInfo As CSheetInfo    'データラベル名とセルアドレスの対応を保持する。定数のように利用。クラスのInitializeイベントで初期化。
Private mEnableReportCurrentCell As Boolean

Private Const Comment As String = "Comment"

Private mCache As Variant                           'キャッシュ
Private mCacheRowOffset As Long                     'キャッシュのオフセット
Private mCacheColumnOffset As Long                  'キャッシュのオフセット
Private mCacheRowCount As Long                      'キャッシュの大きさ
Private mCacheColumnCount As Long                   'キャッシュの大きさ
Private mCurrentRow As Long                         '基点セル(mDateCell)の代わり
Private mCurrentColumn As Long                      '基点セル(mDateCell)の代わり
Private mEndOfRow As Long                           'EndOfRow(mEorCell)の代わり(データの終わりの行のセル)

Public Event CacheUpdate(Msg As String)
Public Event ButtonClick(Msg As String)


Property Get sheetInfo() As CSheetInfo
'内容:
'   CSheetInfoのインスタンスを返します
'
'パラメータ:
'   なし
'
'戻り値:
'    CSheetInfoのインスタンス
'
'例外:
'   SheetInfoのインスタンスが未作成のときは例外9999を投げます
'
'注意事項:

    If mSheetInfo Is Nothing Then
        Call TheError.Raise(9999, TypeName(Me) & ".SheetInfo", "WorkSheet Reader has not initialized !")
    Else
        Set sheetInfo = mSheetInfo
    End If
End Property



Property Let EnableReportCurrentCell(ByVal pFlagEnable As Boolean)
'内容:
'   デバッグ用。基点セル報告機能をOn/Offします。
'   　”基点セル報告機能”：読込み基点セルが移動する度に基点セルの位置をイミディエイトウインドウに出力します。
'
'パラメータ:
'   [pFlgEnable]  In  Boolean型:  True：基点セル報告機能をOn、False：基点セル報告機能をOff
'
'戻り値:
'   なし
'
'注意事項:
'
    mEnableReportCurrentCell = pFlagEnable
End Property
Property Get EnableReportCurrentCell() As Boolean
'内容:
'   デバッグ用。基点セル報告機能の現在設定を返します。
'
'パラメータ:
'   なし
'
'戻り値:
'    True：基点セル報告機能On、False：基点セル報告機能Off
'
'注意事項:
'
    EnableReportCurrentCell = mEnableReportCurrentCell
End Property

Private Sub Init()
    
    With mTargetWorkSheet.UsedRange
        mCache = .Value
        mCacheRowOffset = .Row - 1
        mCacheColumnOffset = .Column - 1
        mCacheRowCount = .Rows.Count
        mCacheColumnCount = .Columns.Count
    End With
    If mEnableReportCurrentCell = True Then
        Debug.Print "Cache Offset(" & mCacheRowOffset & "," & mCacheColumnOffset & ")"
        Debug.Print "Cache Size(" & mCacheRowCount & "," & mCacheColumnCount & ")"
    End If
    
    Me.AsIFileStream.Rewind
    updataEor
    
End Sub
Public Sub Initialize(ByRef pBook As Workbook, ByVal pSheetName As String, ByRef pSheetInfoFactory As ISheetInfoFactory)
'内容:
'   初期化メソッド。下記を行う。
'     ①読込み対象ワークシート名をメンバ変数に設定
'     ②キャッシュ取込
'     ③Rewind
'     ④EndOfRow更新
'
'パラメータ:
'   [pBook]       IN  Workbook型:  読込み対象のブックのインスタンス
'   [pSheetName]  IN  String型:    読込み対象のワークシート名
'   [pSheetInfoFactory]  IN  CSheetInfoFactory型:   CSheetInfoFactoryのインスタンス
'
'戻り値:
'   なし
'
'例外:
'   初期化処理に失敗すると例外9999を投げます
'
'注意事項:
'
'
    
    On Error GoTo ErrorHandler
    Set mTargetWorkSheet = pBook.Worksheets(pSheetName)
    Set mSheetInfo = pSheetInfoFactory.CreateSheetInfo(mTargetWorkSheet)
    Set mEventWatcher = New CEventWatcher
    Set mEventWatcher.SetSheet = mTargetWorkSheet
    Call Init
    Exit Sub
ErrorHandler:
    Call TheError.Raise(9999, TypeName(Me) & ".Initialize", "Can Not Initialize WorkSheet Reader !")
End Sub

Private Sub Class_Initialize()
'    Set mSheetInfo = New CSheetInfo
End Sub

Private Sub Class_Terminate()
    mCache = Empty
    Set mEventWatcher = Nothing
    Set mSheetInfo = Nothing
    Set mTargetWorkSheet = Nothing
End Sub



Public Function AsIFileStream() As IFileStream
'内容:
'   IFileStream型へのキャスト
'
'パラメータ:
'   なし
'
'戻り値:
'    IFileStream型の自オブジェクト
'
    Set AsIFileStream = Me
End Function

Private Sub IFileStream_SetLocation(ByVal pDataName As String)
'内容:
'   PrimaryKeyをサーチして基点セルを移動させます
'   （Findメソッドで引数に指定した文字列に一致するセルを探しだします）
'
'パラメータ:
'    [pDataName] IN  String型:  Test Instance Name列をサーチする文字列
'
'例外:
'    サーチの結果、見つからなかった場合は例外9999を投げます
'
'戻り値:
'    なし
'
    Call CheckInitializeCalled
    Dim col As String
    col = GetCacheColumn(GetAddress(mSheetInfo.GetPrimaryKey))
    Dim i As Long
    Dim flgFound As Boolean
    For i = 1 To mCacheRowCount
        If mCache(mCacheRowOffset + i, col) = pDataName Then
            flgFound = True
            mCurrentRow = i
            mCurrentColumn = col
        End If
    Next i
    If flgFound = False Then
        TheError.Raise 9999, TypeName(Me) & ".SetLocation", "Can Not Find The Primary Key [" & pDataName & "] !"
    End If
    
    If mEnableReportCurrentCell = True Then
        Debug.Print "SetLocation: Cache(" & mCurrentRow & "," & mCurrentColumn & ")"
    End If
End Sub

Private Sub IFileStream_MoveNext()
'内容:
'   基点セルを1つ進めます（行方向）
'
'戻り値:
'    なし
'
    Call CheckInitializeCalled
    mCurrentRow = mCurrentRow + 1
    If mEnableReportCurrentCell = True Then
        Debug.Print "MoveNext: Cache(" & mCurrentRow & "," & mCurrentColumn & ")"
    End If
End Sub

Private Sub IFileStream_Rewind()
'内容:
'   基点セルを初期位置に戻します（初期位置：Primary Keyの見出しの直下のデータセル）
'
'戻り値:
'    なし
'
    Call CheckInitializeCalled
    mCurrentRow = GetCacheRow(GetAddress(mSheetInfo.GetPrimaryKey)) + 1
    mCurrentColumn = GetCacheColumn(GetAddress(mSheetInfo.GetPrimaryKey))
'シートオブジェクト取得時の動作に合わせてチェックしない。2011/03/22
'    If mCurrentRow > mCacheRowOffset + mCacheRowCount Then
'        TheError.Raise 9999, TypeName(Me) & ".Rewind", "Can Not Find The Beginning Of Data !"
'    End If
    
    If mEnableReportCurrentCell = True Then
        Debug.Print "Rewind: Cache(" & mCurrentRow & "," & mCurrentColumn & ")"
    End If
End Sub

Private Function IFileStream_IsEOR() As Boolean
'内容:
'   基点セルがEndOfRowを超えたどうかを判定します
'
'戻り値:
'    True：基点セルがEndOfRowを超えた、False：基点セルはEndOfRowを超えていない
    Call CheckInitializeCalled
    IFileStream_IsEOR = (mCurrentRow > mEndOfRow)
End Function



Public Function AsIParameterReader() As IParameterReader
'内容:
'   IParameterReader型へのキャスト
'
'パラメータ:
'   なし
'
'戻り値:
'    IParameterReader型の自オブジェクト
'
    Set AsIParameterReader = Me
End Function

Private Function IParameterReader_ReadAsBoolean(ByVal pDataName As String) As Boolean
    Dim Value As Variant
    Value = mCache(mCurrentRow, dataColumn(pDataName))
    If IsNumeric(Value) Then
        IParameterReader_ReadAsBoolean = CBool(Value)
    Else
        TheError.Raise 9999, TypeName(Me) & ".ReadAsBoolean", "Does Not Numeric:" & Value & " Cache(" & mCurrentRow & "," & dataColumn(pDataName) & ")"
    End If
End Function

Private Function IParameterReader_ReadAsDouble(ByVal pDataName As String) As Double
    Dim Value As Variant
    Value = mCache(mCurrentRow, dataColumn(pDataName))
    If IsNumeric(Value) Then
        IParameterReader_ReadAsDouble = CDbl(Value)
    Else
        TheError.Raise 9999, TypeName(Me) & ".ReadAsDouble", "Does Not Numeric:" & Value & " Cache(" & mCurrentRow & "," & dataColumn(pDataName) & ")"
    End If
End Function

Private Function IParameterReader_ReadAsInteger(ByVal pDataName As String) As Integer
'pDataNameに!Rowが指定されるとシートの行番号が返る
    Dim Value As Variant
    If pDataName = "!Row" Then
        Value = mCurrentRow
    Else
        Value = mCache(mCurrentRow, dataColumn(pDataName))
    End If
    If IsNumeric(Value) Then
        IParameterReader_ReadAsInteger = CInt(Value)
    Else
        TheError.Raise 9999, TypeName(Me) & ".ReadAsInteger", "Does Not Numeric:" & Value & " Cache(" & mCurrentRow & "," & dataColumn(pDataName) & ")"
    End If
End Function

Private Function IParameterReader_ReadAsLong(ByVal pDataName As String) As Long
'pDataNameに!Rowが指定されるとシートの行番号が返る
    Dim Value As Variant
    If pDataName = "!Row" Then
        Value = mCurrentRow
    Else
        Value = mCache(mCurrentRow, dataColumn(pDataName))
    End If
    If IsNumeric(Value) Then
        IParameterReader_ReadAsLong = CLng(Value)
    Else
        TheError.Raise 9999, TypeName(Me) & ".ReadAsLong", "Does Not Numeric:" & Value & " Cache(" & mCurrentRow & "," & dataColumn(pDataName) & ")"
    End If
End Function

Private Function IParameterReader_ReadAsString(ByVal pDataName As String) As String
'内容:
'   データラベルから列を割り出し、値を読み取って文字列型で返します
'
'パラメータ:
'    [pDataName] IN  String型:  データラベル（候補は大文字Propertyで取得可能。またLabelで全てを取得可能）
'
'例外:
'   データラベルが想定外の場合、例外9999を投げます
'
'戻り値:
'    特定されたセルの値（Valueメソッドの戻り値）
    IParameterReader_ReadAsString = mCache(mCurrentRow, dataColumn(pDataName))
End Function


Private Function dataColumn(ByVal pDataName As String) As Long
    Call CheckInitializeCalled
    If IFileStream_IsEOR Then
        Call TheError.Raise(9999, TypeName(Me) & ".dataColumn", "Out Of Range !")
    End If
    If IsExist(pDataName) Then
        dataColumn = GetCacheColumn(GetAddress(pDataName))
    Else
        Call TheError.Raise(9999, TypeName(Me) & ".dataCell", "Can Not Find The Parameter [" & pDataName & "] !")
    End If
    If mEnableReportCurrentCell = True Then
        Debug.Print "dataCell: Cache(" & mCurrentRow & "," & GetCacheColumn(GetAddress(pDataName)) & ")"
    End If
End Function

Private Function IsExist(ByVal pKey As String) As Boolean
    IsExist = mSheetInfo.IsExist(pKey)
End Function

Private Sub updataEor()

    If mSheetInfo.IsExist(Comment) Then
        mCacheColumnCount = mTargetWorkSheet.Range(mSheetInfo.GetAddress(Comment)).Column - mCacheColumnOffset
    End If

    mEndOfRow = mCacheRowOffset + mCacheRowCount

    Dim i As Long
    For i = mCurrentRow To mCacheRowCount
        Dim flg_empty As Boolean
        flg_empty = True
        Dim j As Long
        For j = 1 To mCacheColumnCount
            If mCache(i, j) <> vbNullString Then
                flg_empty = False
                Exit For
            End If
        Next j
        If flg_empty = True Then
            mEndOfRow = i - 1 + mCacheRowOffset
            Exit For
        Else
            Dim col As String
            col = GetCacheColumn(GetAddress(mSheetInfo.GetPrimaryKey))
            If mCache(i, col) = "END" Then
                mEndOfRow = i - 1 + mCacheRowOffset
                Exit For
            End If
        End If
    Next i

    mCacheRowCount = mEndOfRow

    If mEnableReportCurrentCell = True Then
       Debug.Print "updateEor: Cache EndOfRow: " & mEndOfRow
    End If
End Sub

Private Sub CheckInitializeCalled()
    If mTargetWorkSheet Is Nothing Then
        TheError.Raise 9999, TypeName(Me) & ".CheckInitializeCalled", "Call Initialize method first!"
    End If
End Sub

Private Function GetCacheRow(ByVal pAddress As String) As Long
    Dim rowStr As Long
    If pAddress Like "[A-Z][A-Z]*" Then
        rowStr = Strings.Mid$(pAddress, 3)
    ElseIf pAddress Like "[A-Z]*" Then
        rowStr = Strings.Mid$(pAddress, 2)
    Else
        TheError.Raise 9999, TypeName(Me) & ".GetCacheRow", "Illegal Address:" & pAddress
    End If
    
    Dim Row As Long
    If Not IsNumeric(rowStr) Then
        TheError.Raise 9999, TypeName(Me) & ".GetCacheRow", "Illegal Address:" & pAddress
    End If
    Row = CLng(rowStr)
    
    GetCacheRow = Row - mCacheRowOffset
End Function

Private Function GetCacheColumn(ByVal pAddress As String) As Long
    Dim secondFigure As Long
    Dim firstFigure As Long
    If pAddress Like "[A-Z][A-Z]*" Then
        secondFigure = Asc(Strings.Left$(pAddress, 1)) - Asc("A") + 1
        firstFigure = Asc(Strings.Mid$(pAddress, 2, 1)) - Asc("A") + 1
    ElseIf pAddress Like "[A-Z]*" Then
        secondFigure = 0
        firstFigure = Asc(Strings.Left$(pAddress, 1)) - Asc("A") + 1
    Else
        TheError.Raise 9999, TypeName(Me) & ".GetCacheColumn", "Illegal Address:" & pAddress
    End If
    
    Dim col As Long
    col = secondFigure * 26 + firstFigure
      
    GetCacheColumn = col - mCacheColumnOffset
End Function
Private Function GetAddress(ByVal pName As String) As String
'見出しのアドレスを返す
    GetAddress = mSheetInfo.GetAddress(pName)
End Function



Private Sub mEventWatcher_ButtonClick(Msg As String)
    RaiseEvent ButtonClick(Msg)
End Sub

Private Sub mEventWatcher_WorksheetChange(ByRef TargetRng As Range)
    Call UpdateCache
    RaiseEvent CacheUpdate("EndOfRow:" & mEndOfRow)
End Sub


'Public Property Get WorkSheetCanged() As Boolean
'    WorkSheetCanged = False
'End Property

Private Sub UpdateCache()
    Erase mCache
    Call Init
    If mEnableReportCurrentCell = True Then
        MsgBox "Update Cache!" & vbCrLf _
            & "  EndOfRow:" & mEndOfRow & vbCrLf _
            & "  Cache Size:(" & mCacheRowCount & "," & mCacheColumnCount & ")"
    End If
End Sub

Public Function DataAddress(ByVal pDataName As String) As String
'内容:
'   基点セルのある行の指定データのアドレスを返します
'   （エラーメッセージのセル位置情報提供用）
'
'パラメータ:
'    [pDataName] IN  String型:  データラベル
'
'例外:
'   データラベルが想定外の場合、例外9999を投げます
'
'戻り値:
'    特定されたセルのアドレス
    Dim addr As String
    addr = GetAddress(pDataName)
    Dim w As Long
    If addr Like "[A-Z][A-Z][0-9]*" Then
        w = 2
    ElseIf addr Like "[A-Z][0-9]*" Then
        w = 1
    End If
    DataAddress = Left$(addr, w) & mCurrentRow
End Function

Public Property Get endRow() As Long
'内容:
'   最終行をかえす

    endRow = mEndOfRow
End Property
