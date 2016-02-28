VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFlagBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   プログラム制御用のフラグ情報を管理するクラス
'
'   Revision History:
'       Data        Description
'       2010/10/07  プログラム制御用のフラグ情報を管理する機能を実装した
'                   CVarBank一式を流用している
'       2010/10/22  コメント文を追加＆変更した
'       2010/10/28  Property名称を変更した(Condition → State)
'       2011/03/04　不具合修正(by 0145206097)
'                   初期化(Initialize)及びコンストラクタロジックの変更
'                   インターセプター切り替えロジックの変更及びメソッド名称変更
'                   ・SaveLogFile→Save
'                   ・SaveMode→Dump
'                   論理否定(LNot)の複数フラグ入力時のエラーを追加
'                   排他的論理和(LXor)の追加
'                   フラグクリア(Clear)の追加
'                   エラーレイズ強化
'       2011/03/22  LAnd, LOr, LXorに対して空文字列を指定した場合エラーになるよう修正
'
'作成者:
'   0145184346
'

Option Explicit

Private Const ERR_NUMBER = 9999        ' Error番号
Private Const CLASS_NAME = "CFlagBank" ' Class名称を保持する

'/** フラグバンクが利用するインターフェイス **/
Private mIVarBank As IVarBank
'/** フラグバンク専用のVarBankの実態 **/
Private mMyBank As CVarBank
'/** デフォルトログファイル名**/
Private Const DEFAULT_FILE_NAME As String = "EeeJOBLogTheFlagBank.csv"

Public Function Add(ByRef pFlagName As String, ByRef pState As Boolean, Optional ByRef pComment As String = "")
'内容:
'   Flag名と初期値を登録する
'
'パラメータ:
'   [pFlagName] In  Flag名
'   [pState]    In  Flag初期論理値
'   [pComment]  In  コメント
'
'戻り値:
'
'注意事項:
'   VarBankへはFixedで登録する
'
    If pFlagName = "" Then GoTo IsAnonymous
    On Error GoTo AlreadyExist
    Call mIVarBank.Add(pFlagName, pState, True, pComment)
    Exit Function
IsAnonymous:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".Add", "Unnamed Flag Added!"
AlreadyExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".Add", "This Flag Is Already Exist! " & pFlagName
End Function

Public Property Let state(ByRef pFlagName As String, ByRef pState As Boolean)
'内容:
'   Flagの論理値を変更する
'
'パラメータ:
'   [pFlagName] In  Flag名
'   [pState]    In  Flag論理値
'
'戻り値:
'
'注意事項:
'   複数Flagの入力が可能（カンマ区切り）
'
    On Error GoTo NotExist
    Dim flgArray() As String
    Dim flgIndex As Integer
    flgArray = Split(Replace(pFlagName, " ", ""), ",")
    For flgIndex = 0 To UBound(flgArray)
        If (flgArray(flgIndex) <> "") Then
            With mIVarBank
                .ReadOnly(flgArray(flgIndex)) = False
                .Value(flgArray(flgIndex)) = pState
            End With
        End If
    Next flgIndex
    Exit Property
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LetState", "This Flag Is Not Exist! " & flgArray(flgIndex)
End Property

Public Property Get state(ByRef pFlagName As String) As Boolean
'内容:
'   Flagの論理値を返す
'
'パラメータ:
'   [pFlagName] In  Flag名
'
'戻り値:
'   Flag論理値
'
'注意事項:
'
    On Error GoTo NotExist
    state = mIVarBank.Value(pFlagName)
    Exit Property
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".GetState", "This Flag Is Not Exist! " & pFlagName
End Property

Public Property Let Comment(ByRef pFlagName As String, pComment As String)
'内容:
'   Flagのコメントを変更する
'
'パラメータ:
'   [pFlagName] In  Flag名
'   [pComment]  In  コメント
'
'戻り値:
'
'注意事項:
'
    On Error GoTo NotExist
    mIVarBank.Comment(pFlagName) = pComment
    Exit Property
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LetComment", "This Flag Is Not Exist! " & pFlagName
End Property

Public Property Get Comment(ByRef pFlagName As String) As String
'内容:
'   Flagのコメントを返す
'
'パラメータ:
'   [pFlagName] In  Flag名
'
'戻り値:
'   コメント
'
'注意事項:
'
    On Error GoTo NotExist
    Comment = mIVarBank.Comment(pFlagName)
    Exit Property
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LetComment", "This Flag Is Not Exist! " & pFlagName
End Property

Public Property Get Count() As Long
'内容:
'   登録してあるFlag数を返す
'
'パラメータ:
'
'戻り値:
'   Flag数
'
'注意事項:
'
    Count = mIVarBank.Count
End Property

Public Property Get IsExist(ByRef pFlagName As String) As Boolean
'内容:
'   指定したFlagが登録されているかどうか確認する
'
'パラメータ:
'   [pFlagName] In  Flag名
'
'戻り値:
'   Flagの登録あり(True)
'   Flagの登録なし(False)
'
'注意事項:
'
    IsExist = mIVarBank.IsExist(pFlagName)
End Property

Public Function Search(ByRef SearchPattern As String) As String
'内容:
'   キーワードによるFlag検索
'
'パラメータ:
'   [SearchPattern] In  Flagの検索キーワード
'
'戻り値:
'   Patternと一致したFlag名(","区切り)
'
'注意事項:
'
    Search = mIVarBank.Search(SearchPattern)
End Function

Public Function SearchState(ByRef pState As Boolean) As String
'内容:
'   指定した論理値に一致するFlag名を返す
'
'パラメータ:
'   [pState]  In  Flag論理値
'
'戻り値:
'   論理値と一致したFlag名(","区切り)
'
'注意事項:
'
    Dim flgArray() As String
    Dim flgIndex As Integer
    flgArray = Split(Me.Search("*"), ",")
    For flgIndex = 0 To UBound(flgArray)
        If (flgArray(flgIndex) <> "") Then
            With mIVarBank
                If (pState = .Value(flgArray(flgIndex))) Then
                    SearchState = SearchState & flgArray(flgIndex) & ","
                End If
            End With
        End If
    Next flgIndex
End Function

Public Sub Clear()
'内容:
'   登録したFlagを全て初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'   2011/03/04追加 0145206097
'
    mIVarBank.AllClear
End Sub

Public Function LNot(ByRef pFlagName As String) As Boolean
'内容:
'   登録したFlagの論理否定を返す（Not）
'
'パラメータ:
'   [pFlagName]  In  Flag名
'
'戻り値:
'   論理否定
'
'注意事項:
'   登録されている論理値は変更されない
'   複数Flagの入力は不可
'
    On Error GoTo NotExist
    LNot = Not (mIVarBank.Value(pFlagName))
    Exit Function
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LNot", "This Flag Is Not Exist! " & pFlagName
End Function

Public Function LAnd(ByRef pFlagName As String) As Boolean
'内容:
'   登録したFlagの論理積を返す（And）
'
'パラメータ:
'   [pFlagName] In  Flag名称
'
'戻り値:
'   論理積
'
'注意事項:
'   登録されている論理値は変更されない
'   複数Flagの入力が可能（カンマ区切り）
'   単独Flagの場合はそのFlagの論理値が返る
'
    On Error GoTo NotExist
    Dim flgArray() As String
    Dim flgIndex As Integer
    flgArray = Split(Replace(pFlagName, " ", ""), ",")
    LAnd = True
    
    If UBound(flgArray) < 0 Then
        On Error GoTo 0
        TheError.Raise ERR_NUMBER, TypeName(Me) & ".LAnd", "Please Input Flag Name!"
    End If

    For flgIndex = 0 To UBound(flgArray)
        If (flgArray(flgIndex) <> "") Then
            LAnd = LAnd And (mIVarBank.Value(flgArray(flgIndex)))
        End If
    Next flgIndex
    Exit Function
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LAnd", "This Flag Is Not Exist! " & flgArray(flgIndex)
End Function

Public Function LOr(ByRef pFlagName As String) As Boolean
'内容:
'   登録したFlagの論理和を返す（Or）
'
'パラメータ:
'   [pFlagName] In  Flag名称
'
'戻り値:
'   論理和
'
'注意事項:
'   登録されている論理値は変更されない
'   複数Flagの入力が可能（カンマ区切り）
'   単独Flagの場合はそのFlagの論理値が返る
'
    On Error GoTo NotExist
    Dim flgArray() As String
    Dim flgIndex As Integer
    flgArray = Split(Replace(pFlagName, " ", ""), ",")
    LOr = False

    If UBound(flgArray) < 0 Then
        On Error GoTo 0
        TheError.Raise ERR_NUMBER, TypeName(Me) & ".LOr", "Please Input Flag Name!"
    End If

    For flgIndex = 0 To UBound(flgArray)
        If (flgArray(flgIndex) <> "") Then
            LOr = LOr Or (mIVarBank.Value(flgArray(flgIndex)))
        End If
    Next flgIndex
    Exit Function
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LOr", "This Flag Is Not Exist! " & flgArray(flgIndex)
End Function

Public Function LXor(ByRef pFlagName As String) As Boolean
'内容:
'   登録したFlagの排他的論理和を返す（LXor）
'
'パラメータ:
'   [pFlagName] In  Flag名称
'
'戻り値:
'   排他的論理和
'
'注意事項:
'   2011/03/04追加 0145206097
'   登録されている論理値は変更されない
'   複数Flagの入力が可能（カンマ区切り）
'   単独Flagの場合はそのFlagの論理値が返る
'
    On Error GoTo NotExist
    Dim flgArray() As String
    Dim flgIndex As Integer
    flgArray = Split(Replace(pFlagName, " ", ""), ",")
    LXor = False
    
    If UBound(flgArray) < 0 Then
        On Error GoTo 0
        TheError.Raise ERR_NUMBER, TypeName(Me) & ".LXor", "Please Input Flag Name!"
    End If

    For flgIndex = 0 To UBound(flgArray)
        If (flgArray(flgIndex) <> "") Then
            LXor = LXor Xor (mIVarBank.Value(flgArray(flgIndex)))
        End If
    Next flgIndex
    Exit Function
NotExist:
    TheError.Raise ERR_NUMBER, TypeName(Me) & ".LXor", "This Flag Is Not Exist! " & flgArray(flgIndex)
End Function

Public Sub Initialize()
'内容:
'   TheFlagBankの初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Set mMyBank = New CVarBank
    Set mIVarBank = mMyBank
End Sub

Public Sub Dump(ByVal pDump As Boolean)
'内容:
'   TheFlagBankのダンプ
'
'パラメータ:
'   [pDump]    In  ダンプモード指定
'
'戻り値:
'
'注意事項:
'
    Dim inDumping As Boolean
    inDumping = IsInterceptor
    If pDump And Not inDumping Then
        Dim myInterceptor As CVarBankInterceptor
        Set myInterceptor = New CVarBankInterceptor
        Call myInterceptor.Initialize(mMyBank)
        Set mIVarBank = myInterceptor
        TheExec.Datalog.WriteComment "Eee JOB Output Log! :TheFlagBank Condition"
    ElseIf Not pDump And inDumping Then
        Set mIVarBank = mMyBank
    End If
End Sub

Public Function Save(Optional saveFileName As String = DEFAULT_FILE_NAME) As Long
'内容:
'   ダンプデータを外部ファイルへ出力する
'
'パラメータ:
'   [SaveFileName]  In  ファイル名指定（オプション）
'
'戻り値:
'
'注意事項:
'   ファイル名の指定がない場合はデフォルト名でカレントフォルダに出力する
'
    If Not IsInterceptor Then Exit Function
    If saveFileName = "" Then saveFileName = DEFAULT_FILE_NAME
    Dim myInterceptor As CVarBankInterceptor
    Set myInterceptor = mIVarBank
    Call myInterceptor.SaveLogFile(saveFileName)
    Set mIVarBank = mMyBank
End Function

Private Function IsInterceptor() As Boolean
'/* 既にインターセプターを利用しているかどうかを型チェックで判断する */
    Dim myInterceptor As CVarBankInterceptor
    On Error GoTo IsInterceptor
    Set myInterceptor = mIVarBank
    IsInterceptor = True
    Exit Function
IsInterceptor:
    IsInterceptor = False
End Function

Private Sub Class_Initialize()
    Me.Initialize
End Sub

Private Sub Class_Terminate()
    Set mMyBank = Nothing
    Set mIVarBank = Nothing
End Sub
