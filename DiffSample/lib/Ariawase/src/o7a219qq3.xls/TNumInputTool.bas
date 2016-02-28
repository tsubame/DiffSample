Attribute VB_Name = "TNumInputTool"
'概要:
'   Flow Tableワークシートへ、テスト番号ルールに従ったテスト番号を入力する
'
'目的:
'   テスト番号ルールの遵守
'   テスト番号入力禁止部への番号入力トラブル防止
'
'作成者:
'   SLSI 今手
'
'開発履歴:
'   2009-05-25 試作初版完成。
'
'   2009-05-27 テスト番号の削除処理を、手動時の操作に合わせて 入力値を""からEmptyに変更。
'              テスト番号入力禁止部に値が入っていたときのメッセージWindowを変更。
'
'   2009-06-15 実行時にテスト番号ルールのVersionを指定することで複数のテスト番号ルールに対応。
'              2008-12-17に合意されたテスト番号ルールのVersion2を実装。
'              テスト番号ルールVersion1のdcラベルを廃止(JobチェックToolの検出仕様に合わせた)。
'              TOOLのVersion番号公開機能追加
'
'   2009-06-18 正式版公開にむけてモジュール名変更(ZZZ_InputTestNumTool-->TNumInputTool)
'   2012-12-26 H.Arikawa 自動化のTNumberに合わせてVER3を追加

Option Explicit

'Flow Tableシートに関する定義など
Private Const LABEL_INDEX = "B5"
Private Const OPCODE_INDEX = "G5"
Private Const PARAMETER_INDEX = "H5"
Private Const RESULT_INDEX = "O5"
Private Const LAST_OPCODE_VALUE = "set-device"
Private Const FLOW_TABLE_SHT_NAME = "Flow Table"
Private Const TEST_CATEGORY_HEADER_NAME = "SEQ"
Private Const PROHIBITION_PLACE_MSG = "It is a test number input prohibition place"

'テスト番号ルールを実装した関数名の定義(※新ルール追加時にはココをメンテ)
Private Const TNUM_RULE_SUPPLIER_V1 = "RetStartTestNumber_V1"
Private Const TNUM_RULE_SUPPLIER_V2 = "RetStartTestNumber_V2"
Private Const TNUM_RULE_SUPPLIER_V3 = "RetStartTestNumber_V3"

Public Enum TNumRuleVer
    VER_1 = 1#
    VER_2 = 2#
    VER_3 = 3#
End Enum

'Parameter セルのIndex用
Private m_ParameterRng As Range

'TOOLのVersion公開用
Private Const TOOL_VERSION = "1.00"


Public Sub InputTestNumber(ByVal pRuleVerNo As TNumRuleVer, _
Optional ByVal pTgtFlowTable As Worksheet = Nothing)
'内容:
'   Flow Tableシートにテスト番号を入力する
'
'パラメータ:
'   [pRuleVerNo]       In   テスト番号ルールのVersion番号
'   [pTgtFlowTable]    In   番号を入力するFlowTabelシートObject（オプション）
'
'戻り値:
'
'注意事項:
'
    Dim Label As String
    Dim OpcodeRng As Range
    Dim LabelRng As Range
    Dim ResultRng As Range
    Dim TestNumber As Long
    Dim answer As Long
    Dim GetTestNumber As String
        
    '使用するテスト番号ルールの選択処理(※新ルール追加時にはココをメンテ)
    Select Case pRuleVerNo
        Case VER_1
            GetTestNumber = TNUM_RULE_SUPPLIER_V1
        Case VER_2
            GetTestNumber = TNUM_RULE_SUPPLIER_V2
        Case VER_3
            GetTestNumber = TNUM_RULE_SUPPLIER_V3
        Case Else
            Call Err.Raise(9999#, "InputTestNumber", "Rule Version = " & pRuleVerNo & " is unknown test number rule !")
            Exit Sub
    End Select
            
    'ワークシートの指定省略時は、Flow Tableを勝手にターゲットとする
    '(本当はシートが無いときのエラー処理があるほうがよい）
    If pTgtFlowTable Is Nothing Then
        Set pTgtFlowTable = Worksheets(FLOW_TABLE_SHT_NAME)
    End If
                
    '初回のインデックス設定
    Set LabelRng = pTgtFlowTable.Range(LABEL_INDEX)
    Set OpcodeRng = pTgtFlowTable.Range(OPCODE_INDEX)
    Set m_ParameterRng = pTgtFlowTable.Range(PARAMETER_INDEX)
    Set ResultRng = pTgtFlowTable.Range(RESULT_INDEX)
        
    '一番最初のラベル
    Label = LabelRng.Value
    TestNumber = Application.Run(GetTestNumber, Label)
    
    While OpcodeRng <> LAST_OPCODE_VALUE
        'ラベルの値確認(新しいラベルが登場したら)
        If (LabelRng.Value <> "") And (Label <> LabelRng.Value) Then
            '新しいラベル用のテスト番号取得
            Label = LabelRng.Value
            TestNumber = Application.Run(GetTestNumber, Label)
        End If
        
        'Enable Wordが空欄、Resultが空欄、Parameterの値がSEQのところは
        ' テスト番号は入力されないでOK？
        If (LabelRng.offset(0#, 1#).Value <> "") And (ResultRng.Value <> "") _
        And (m_ParameterRng.Value <> TEST_CATEGORY_HEADER_NAME) Then
                
            'パラメータが空欄でない、かつTNameが空欄でない
            If (m_ParameterRng.Value <> "") And (TNameRng.Value <> "") Then
                'Test番号をセルに入力
                TNumberRng.Value = TestNumber
                'テスト番号を1つインクリメントする
                TestNumber = TestNumber + 1#
            Else
                'Test番号を入力しない（テスト番号入力禁止の部分なので空欄にする）※ここの処理は不要かも
                If TNumberRng.Value <> "" Then
                    answer = MsgBox(MakeEraseConfirmMsg, vbYesNo + vbExclamation, PROHIBITION_PLACE_MSG)
                    If answer = vbYes Then
                       TNumberRng.Value = Empty
                    End If
                End If
            End If
        Else
            'Test番号を入力しない（テスト番号入力禁止の部分なので空欄にする）
                If TNumberRng.Value <> "" Then
                    answer = MsgBox(MakeEraseConfirmMsg, vbYesNo + vbExclamation, PROHIBITION_PLACE_MSG)
                    If answer = vbYes Then
                        TNumberRng.Value = Empty
                    End If
                End If
        End If
    
        'Indexをひとつ下に進める
        Set LabelRng = LabelRng.offset(1#, 0#)
        Set m_ParameterRng = m_ParameterRng.offset(1#, 0#)
        Set OpcodeRng = OpcodeRng.offset(1#, 0#)
        Set ResultRng = ResultRng.offset(1#, 0#)
    
    Wend
    
    Call MsgBox("Done", vbInformation, "InputTestNumber")

End Sub

Public Function TNumInputToolVer() As String
'内容:
'   TNumInputToolのVersion番号を返す
'
'パラメータ:
'
'戻り値:
'   TNumInputToolのVersion番号
'
'注意事項:
'
    TNumInputToolVer = TOOL_VERSION

End Function

'--------------------------------------------------------------------------------
'以下 Private Function

'#Pass
'TName Rangeの取得
Private Function TNameRng() As Range
    Set TNameRng = m_ParameterRng.offset(0#, 1#)
End Function

'#Pass
'TNum Rangeの取得
Private Function TNumberRng() As Range
    Set TNumberRng = m_ParameterRng.offset(0#, 2#)
End Function

'#Pass
'テスト番号の消去確認用のメッセージ作成用
Private Function MakeEraseConfirmMsg() As String
    MakeEraseConfirmMsg = "Address= " & m_ParameterRng.Address & vbCrLf & _
    "Parameter= " & m_ParameterRng.Value & vbCrLf & _
    "Test number= " & m_ParameterRng.offset(0#, 2#).Value & vbCrLf & _
    "Do you erase a test number?"
End Function


'--------------------------------------------------------------------------------
'以下 テスト番号ルールの実装(※新ルール追加時にはココをメンテ)

'#Pass
'テスト番号ルール Version1
'2009/05/22時点でのLegacyルールの実装(ソース=木村さんの情報より)
Private Function RetStartTestNumber_V1(ByVal pLabel As String) As Long
    
    Select Case pLabel
        Case "dcpar"
            RetStartTestNumber_V1 = 2#
            Exit Function
        Case "image"
            RetStartTestNumber_V1 = 1002#
            Exit Function
        Case "color"
            RetStartTestNumber_V1 = 2002#
            Exit Function
        Case "flmura"
            RetStartTestNumber_V1 = 3002#
            Exit Function
        Case "grade"
            RetStartTestNumber_V1 = 4002#
            Exit Function
        Case "shiroten"
            RetStartTestNumber_V1 = 5002#
            Exit Function
        Case "nashiji"
            RetStartTestNumber_V1 = 6002#
            Exit Function
        Case "margin"
            RetStartTestNumber_V1 = 7002#
            Exit Function
        Case Else
            Call Err.Raise(9999#, "RetStartTestNumber_V1", pLabel & " is UnKnown Label !")
    End Select

End Function

'#Pass
'テスト番号ルール Version2
'2008/12/17に厚木、熊本、長崎で合意された新ルールの実装(ソース=米田さんの情報より)
Private Function RetStartTestNumber_V2(ByVal pLabel As String) As Long
    
    Select Case pLabel
        Case "dcpar"
            RetStartTestNumber_V2 = 2#
            Exit Function
        Case "image"
            RetStartTestNumber_V2 = 1002#
            Exit Function
        Case "color"
            RetStartTestNumber_V2 = 3002#
            Exit Function
        Case "grade"
            RetStartTestNumber_V2 = 5002#
            Exit Function
        Case "shiroten"
            RetStartTestNumber_V2 = 6002#
            Exit Function
        Case "margin"
            RetStartTestNumber_V2 = 8002#
            Exit Function
        Case Else
            Call Err.Raise(9999#, "RetStartTestNumber_V2", pLabel & " is UnKnown Label !")
    End Select

End Function

'#Pass
'テスト番号ルール Version3
'2012/11/12 自動化に対応したルール実装
Private Function RetStartTestNumber_V3(ByVal pLabel As String) As Long
    
    Select Case pLabel
        Case "dcpar"
            RetStartTestNumber_V3 = 2#
            Exit Function
        Case "image"
            RetStartTestNumber_V3 = 1002#
            Exit Function
        Case "grade"
            RetStartTestNumber_V3 = 5002#
            Exit Function
        Case "shiroten"
            RetStartTestNumber_V3 = 6002#
            Exit Function
        Case "margin"
            RetStartTestNumber_V3 = 8002#
            Exit Function
        Case Else
            Call Err.Raise(9999#, "RetStartTestNumber_V3", pLabel & " is UnKnown Label !")
    End Select

End Function

