VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csPALS_LoopMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
' クラス名  ：  csPALS_LoopMain.cls
' 概要      ：  LOOP測定時に使用する情報をもつクラス
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/16　新規作成        K.Sumiyashiki
'==========================================================================

Private xLoopCategory() As csPALS_LoopCategoryParams        'csPALS_LoopCategoryParamsクラスの定義

Public CategoryInfoList As New Collection                   '条件名(OF,ML,etc..)とインデックスを紐付ける為のコレクション
Public CategoryKeyList  As New Collection                   'キーの重複チェックの為のコレクション


'*********************************************************************************
' 関数名    ：  Class_Initialize
' 概要      ：  インスタンス生成時の初期化
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Initialize()

    '紐付け用コレクション初期化
    Call sub_InitCollection(CategoryInfoList)
    Call sub_InitCollection(CategoryKeyList)
    
    'TestConditionのデータ読込
    Call sub_ReadCategoryInfo
    
    'TestConditionのデータに不備がないかチェック
    Call sub_CheckCategoryInfo

End Sub


'*********************************************************************************
' 関数名    ：  Class_Terminate
' 概要      ：  インスタンス開放時の処理
' 引数      ：  なし
' 備考      ：  なし
' 更新履歴  ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Private Sub Class_Terminate()

    '紐付け用コレクション初期化
    Call sub_InitCollection(CategoryInfoList)
    Call sub_InitCollection(CategoryKeyList)

End Sub


'*********************************************************************************
' プロパティ名  ：  LoopCategory
' 概要          ：  カテゴリ番号に対応するLoopCategoryオブジェクトを返すプロパティ
' 引数          ：  index :カテゴリ名を示すインデックス
' 備考          ：  読み取り専用
' 更新履歴      ：  Rev1.0      2010/08/18　新規作成        K.Sumiyashiki
'*********************************************************************************
Public Property Get LoopCategory(ByVal index As Integer) As csPALS_LoopCategoryParams

    Set LoopCategory = xLoopCategory(index)

End Property


'********************************************************************************************
' 名前: sub_ReadCategoryInfo
' 内容: TestConditionのパラメータをcsLoopParamsクラスで定義されたLoopParams配列に格納。
'       LoopParamsは条件数分の配列が動的に作成される。
' 引数: なし
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/17　新規作成   K.Sumiyashiki
' 更新履歴： Rev1.1      2011/06/02　処理追加   K.Sumiyashiki
'                                    ⇒Eee-JOB対応
'********************************************************************************************
Private Sub sub_ReadCategoryInfo()

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_ReadCategoryInfo

    Dim TC_y As Long            'Y方向のセル位置を格納する変数
    Dim strbuf As String        '条件名を格納する変数　例)"OF"、"HL"等
    
    'TestConditionシートの存在チェック
    If Not sub_SheetNameCheck(TESTCONDITION) Then

'>>>2011/06/02 K.SUMIYASHIKI ADD
        If sub_SheetNameCheck(ACQTBLSHTNAME) Then
Jump_Set_ACQTBL:
            '条件数でLoopParamsを再定義
            ReDim xLoopCategory(1)
            
            'csLoopParamsクラスのインスタンス生成
            Set xLoopCategory(1) = New csPALS_LoopCategoryParams
            
            strbuf = "DummyData"
            
            'カテゴリ名を入力
            xLoopCategory(1).category = strbuf
            
            '条件名をキーに、LoopParamsの配列番号をコレクションに追加
            CategoryInfoList.Add Item:=1, key:=strbuf
            
            'キー重複チェックの為に、条件名を値とキーにしたコレクション作成
            CategoryKeyList.Add Item:=strbuf, key:=strbuf
        
            'WaitとAverageデータの入力
            With xLoopCategory(CategoryInfoList(strbuf))
                .Average = 1
                .WAIT = 1
            End With
    
            Exit Sub
'<<<2011/06/02 K.SUMIYASHIKI ADD
        Else
            Call sub_errPALS("TestCondition not found at 'sub_ReadCategoryInfo'", "2-4-01-8-37")
            Exit Sub
        End If
    End If
    
    With Worksheets(TESTCONDITION)
    
        'TestConditionシートのY方向のスタートセル位置を設定
        TC_y = TC_START_Y
        
        'Condition Nameの最後まで繰り返し
        Do While Len(.Cells(TC_y, TC_CONDINAME_X).Value)
            
            'Sw_Nodeが一致する行のみ処理
            If Sw_Node = .Cells(TC_y, TC_SWNODE_X).Value Then
            
                'strBufに条件名を代入
                strbuf = .Cells(TC_y, TC_CONDINAME_X).Value
                
                'コレクションにデータがない場合(1項目目)の処理
                If CategoryInfoList.Count = 0 Then
                    'csLoopParamsクラスのインスタンスを生成し、データを追加
                    Call sub_Add_CategoryInfo(strbuf, TC_y)
                Else
                    '既に条件がコレクションに追加されているかのチェック
                    If key_check(CategoryKeyList, strbuf) Then
                        '追加されていた場合、データを追加
                        Call sub_InputTestCondition(strbuf, TC_y)
                    Else
                        'csLoopParamsクラスのインスタンスを生成し、データを追加
                        Call sub_Add_CategoryInfo(strbuf, TC_y)
                    End If
                End If
            End If
        '次の行へ
        TC_y = TC_y + 1
        Loop
    End With

    If CategoryInfoList.Count = 0 Then
        If sub_SheetNameCheck(ACQTBLSHTNAME) Then
            GoTo Jump_Set_ACQTBL
        End If
        Call sub_errPALS("NotFound Info@Sheet[" & TESTCONDITION & "], Node[" & Sw_Node & "] at 'sub_ReadCategoryInfo'", "2-4-01-5-38")
    End If

Exit Sub

errPALSsub_ReadCategoryInfo:
    Call sub_errPALS("Read IG-XL sheet error at 'sub_ReadCategoryInfo'", "2-4-01-0-39")

End Sub


'********************************************************************************************
' 名前: key_check
' 内容: 同じ値がコレクションに存在しているかチェックする
' 引数: col:検索するコレクション
'       query:検索対象の文字列
' 戻値: 同じ値あり=>True
'       同じ値なし=>False
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/17　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Function key_check(ByRef col As Collection, ByRef query As String) As Boolean

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSkey_check

    Dim Item As Variant
    'コレクションの最後まで繰り返し
    For Each Item In col
        '同じ値があった場合、Trueを返し関数を抜ける
        If Item = query Then
            key_check = True
            Exit Function
        End If
    Next
    '同じ値が無い場合、Falseを返す
    key_check = False

Exit Function

errPALSkey_check:
    Call sub_errPALS("key check error at 'csPALS_LoopMain -> key_check'", "2-4-02-0-40")
    
End Function


'********************************************************************************************
' 名前: sub_add_CategoryInfo
' 内容: csLoopParamsクラスのインスタンスを生成し、データを追加する。
'       インスタンスは条件数だけ、動的に作られる。
'       CategoryInfoListコレクションは、キーを条件名に、LoopParamsの配列番号がItemに追加される。
'       CategoryKeyListコレクションは、キー検索の為に作成している。
' 引数: strBuf:条件名
'       TC_y:セルのY方向の位置
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/17　新規作成   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_Add_CategoryInfo(ByRef strbuf As String, ByVal TC_y As Long)
    
    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALSsub_Add_CategoryInfo

    'LoopParamsの配列番号の設定
    Dim intItemCnt As Integer
    
    If CategoryInfoList.Count = 0 Then
        '配列が無い場合、配列番号を1に設定
        '紐付け用のコレクションのキーに0を設定出来ない為、1から開始
        intItemCnt = 1
    Else
        '配列が存在する場合、配列数をインクリメント
        intItemCnt = UBound(xLoopCategory) + 1
    End If

    '条件数でLoopParamsを再定義
    ReDim Preserve xLoopCategory(intItemCnt)
    
    'csLoopParamsクラスのインスタンス生成
    Set xLoopCategory(intItemCnt) = New csPALS_LoopCategoryParams
    
    'カテゴリ名を入力
    xLoopCategory(intItemCnt).category = strbuf
    
    '条件名をキーに、LoopParamsの配列番号をコレクションに追加
    CategoryInfoList.Add Item:=intItemCnt, key:=strbuf
    
    'キー重複チェックの為に、条件名を値とキーにしたコレクション作成
    CategoryKeyList.Add Item:=strbuf, key:=strbuf
    
'    '条件数を加算
    
    'データ入力
    Call sub_InputTestCondition(strbuf, TC_y)

Exit Sub

errPALSsub_Add_CategoryInfo:
    Call sub_errPALS("Add CategoryInfo error at 'sub_Add_CategoryInfo'", "2-4-03-0-41")

End Sub


'********************************************************************************************
' 名前: sub_InputTestCondition
' 内容: TestConditionシートのデータ取得
'       データの種類によって、代入変数を分けている
' 引数: strBuf:条件名
'       TC_y:セルのY方向の位置
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/17　新規作成   K.Sumiyashiki
'            Rev1.1      2011/05/16　エラーコメント修正   K.Sumiyashiki
'********************************************************************************************
Private Sub sub_InputTestCondition(ByVal strbuf As String, ByVal TC_y As Long)

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If

On Error GoTo errPALSsub_InputTestCondition

    Dim strTmp As String        'セルデータを一時保存
    
    With xLoopCategory(CategoryInfoList(strbuf))
                
        'セルデータを一時格納
        strTmp = Worksheets(TESTCONDITION).Cells(TC_y, TC_PROCEDURENAME_X).Value
        
        Select Case strTmp
            'Waitデータの場合
            Case SET_WAIT
                If .WAIT = -1 Then
                    'データが初期値の場合、データ入力
                    .WAIT = Worksheets(TESTCONDITION).Cells(TC_y, TC_ARG1_X).Value
                Else
                    '同じデータが複数ある場合、エラーコメント表示
'>>>2011/5/12 K.SUMIYASHIKI CHANGE
'old                    Call sub_errPALS("Error : Wait data overlaps! of 'TestCondition sheet'")
                    Call sub_errPALS("Wait data overlaps! of 'TestCondition sheet'" & vbCrLf _
                                    & "Condition Name -> " & .category & vbCrLf _
                                    & "Error Row -> " & TC_y, "2-4-04-2-42")
'<<<2011/5/12 K.SUMIYASHIKI CHANGE
                End If
            'Averageデータの場合
            Case SET_AVERAGE
                If .Average = -1 Then
                    'データが初期値の場合、データ入力
                    .Average = Worksheets(TESTCONDITION).Cells(TC_y, TC_ARG1_X).Value
                Else
'>>>2011/5/12 K.SUMIYASHIKI CHANGE
                    '同じデータが複数ある場合、エラーコメント表示
'old                    Call sub_errPALS("Error : Average data overlaps! of 'TestCondition sheet'")
                    Call sub_errPALS("Average data overlaps! of 'TestCondition sheet'" & vbCrLf _
                                    & "Condition Name -> " & .category & vbCrLf _
                                    & "Error Row -> " & TC_y, "2-4-04-2-43")
                End If
'<<<2011/5/12 K.SUMIYASHIKI CHANGE
            'Modeデータの場合
            Case ACQUIRE_MODE
                If Len(.mode) = 0 Then
                    'データが初期値の場合、データ入力
                    .mode = Worksheets(TESTCONDITION).Cells(TC_y, TC_ARG1_X).Value
                Else
'>>>2011/5/12 K.SUMIYASHIKI CHANGE
                    '同じデータが複数ある場合、エラーコメント表示
'old                    Call sub_errPALS("Error : Mode data overlaps! of 'TestCondition sheet'")
                    Call sub_errPALS("Mode data overlaps! of 'TestCondition sheet'" & vbCrLf _
                                    & "Condition Name -> " & .category & vbCrLf _
                                    & "Error Row -> " & TC_y, "2-4-04-2-44")
'<<<2011/5/12 K.SUMIYASHIKI CHANGE
                End If
            Case Else
                '上の文字列以外の場合はスル
                'Eee-JOBで他の文字列を使用している為
        End Select
    End With

Exit Sub

errPALSsub_InputTestCondition:
    Call sub_errPALS("Input TestCondition Data error at 'sub_InputTestCondition'", "2-4-04-0-45")

End Sub


'********************************************************************************************
' 名前: ChangeLoopParams
' 内容: 特性値のバラツキがあった際に、TestConditionシートのパラメータ変更を実施する関数
' 引数: strCategory    :カテゴリ名
'       strWaitOrAve   :Wait or Averageを示す文字列
'       strValue       :変更したい値
' 戻値: True           :変更時エラーなし
'       False          :変更時エラーあり
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/17　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Function ChangeLoopParams(ByVal strCategory As String, ByVal strWaitOrAve As String, ByVal strValue As String) As Boolean

    If g_ErrorFlg_PALS Then
        Exit Function
    End If

On Error GoTo errPALSChangeLoopParams

    Dim TC_y As Long                    'テストコンディションのY方向のセル位置
    Dim Tmp_strWaitOrAve As String
    
    Select Case strWaitOrAve
    Case "Wait"
        Tmp_strWaitOrAve = SET_WAIT
    Case "Average"
        Tmp_strWaitOrAve = SET_AVERAGE
    Case Else
        MsgBox "Error : Program argument error at 'ChangeLoopParams'" & vbCrLf & "ErrCode.2-4-05-4-47", vbExclamation, PALS_ERRORTITLE
        ChangeLoopParams = False
        Exit Function
    End Select
    
    'テストコンディションシート
    With Worksheets(TESTCONDITION)
    
        'TestConditionシートのY方向のスタートセル位置を設定
        TC_y = TC_START_Y
        
        'Condition Nameの最後まで繰り返し
        Do While Len(.Cells(TC_y, TC_CONDINAME_X).Value) > 0
            
            'Condition NameとProcedure Nameが一致する場合、セルの値を上書き
            If (.Cells(TC_y, TC_CONDINAME_X).Value = strCategory) And (.Cells(TC_y, TC_PROCEDURENAME_X).Value = Tmp_strWaitOrAve) _
                And (.Cells(TC_y, TC_SWNODE_X).Value = Sw_Node) Then
                
                'セルの値を上書き
                .Cells(TC_y, TC_ARG1_X).Value = val(strValue)
                
                'フラグをTrueに変更
                ChangeLoopParams = True
                
                '関数を抜ける
                Exit Function
            End If
        '次の行へ
        TC_y = TC_y + 1
        Loop
    End With
    
    '一致する項目が無かった場合、エラーメッセージ
    MsgBox "Error : '" & strCategory & "' not found at 'TestCondition sheet'." & vbCrLf & "         I can't change parameter." & vbCrLf & "ErrCode.2-4-05-4-48", vbExclamation, PALS_ERRORTITLE

Exit Function

errPALSChangeLoopParams:
    Call sub_errPALS("Change LoopParameter error at 'ChangeLoopParams'", "2-4-05-0-46")

End Function


'********************************************************************************************
' 名前: CategoryCount
' 内容: カテゴリ数を返す関数
' 引数: なし
' 戻値: カテゴリ数
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/20　新規作成   K.Sumiyashiki
'********************************************************************************************
Public Function CategoryCount() As Integer

On Error GoTo errPALSCategoryCount

    CategoryCount = UBound(xLoopCategory)

Exit Function

errPALSCategoryCount:
    Call sub_errPALS("Get category count error at 'csPALS_LoopMain -> CategoryCount'", "2-4-06-0-49")

End Function


'********************************************************************************************
' 名前: sub_CheckCategoryInfo
' 内容: TestConditionシートに記載されているパラメータに不備がないかチェックする関数
' 引数: なし
' 戻値: なし
' 備考    ： なし
' 更新履歴： Rev1.0      2010/08/17　新規作成   K.Sumiyashiki
'            Rev1.1      2011/03/21　処理変更   K.Sumiyashiki
'                                    ⇒アベレージ回数チェックを解除
'********************************************************************************************
Private Sub sub_CheckCategoryInfo()

    If g_ErrorFlg_PALS Then
        Exit Sub
    End If
    
On Error GoTo errPALSsub_CheckCategoryInfo

    Dim i As Long       'カテゴリー数を示すループカウンタ

    '全カテゴリを繰り返す
    For i = 1 To UBound(xLoopCategory)
        With xLoopCategory(i)
            'アベレージ回数が512回以上でないかチェック
'>>>2011/03/21 K.SUMIYASHIKI COMMENT OUT
'            If .Average > 511 Then
'                Call sub_errPALS("Error : " & .Category & " average count is over 512.")
'<<<2011/03/21 K.SUMIYASHIKI COMMENT OUT
            If .Average < 1 Then
                Call sub_errPALS(.category & " average count is under 0.", "2-4-07-5-50")
            End If
            
            'アベレージ or Wait設定がシートに記述されているかチェック
            If .Average = -1 Or .WAIT = -1 Then
                Call sub_errPALS(.category & " Averge or Wait is not set.", "2-4-07-5-51")
            End If
            
            'モードに格納されている変数が文字列かどうかチェック
            'IsNumeric関数⇒文字列:False、数値:True
            If IsNumeric(.mode) Then
                '設定されているAverage回数がモードで指定された数値以下かチェック
                If (.Average <> -1) And (.Average < .mode) Then
                    Call sub_errPALS(.category & "  average < Set Mode count.", "2-4-07-5-52")
                End If
                
                '設定されているAverage回数がモードで指定された数値の倍数になっているかチェック
                If (.Average <> -1) And (.Average Mod .mode) Then
                    Call sub_errPALS(.category & "  average is not a Multiple of the 'Mode'.", "2-4-07-5-53")
                End If
            Else
                'モードがオートの場合、大文字AUTOに変換
                If (UCase(.mode) = MODE_AUTO) Or (.mode = "") Then
                    .mode = MODE_AUTO
                End If
            End If
        End With
    Next i

Exit Sub

errPALSsub_CheckCategoryInfo:
    Call sub_errPALS("CheckCategoryInfo error at 'sub_CheckCategoryInfo'", "2-4-07-0-54")

End Sub
