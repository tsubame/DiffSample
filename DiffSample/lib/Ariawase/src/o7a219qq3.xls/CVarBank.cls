VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVarBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   Public変数の代用クラス
'目的:
'   Public変数の使用によるトラブル回避用です。
'   このクラスに登録した値は固定となり、不用意な変更のリスクを避けます。
'   値を修正しようとするとエラーとなります。値を修正する必要がある場合は、設定を変更してから修正します。
'   特にTOPT使用時は、プロシージャ処理途中で別のプロシージャへ処理が移るため、意図しない処理が発生する可能性があり注意が必要です。
'特記事項:
'   ##### 返り値は、入力時と同じ型で受けてください。違う型で受けた場合、全く違う値となることがあります #####
'   例）Double型をInteger型で受ける
'
'   キー文字の大小違いは同じとみなします。
'
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月12日

Option Explicit

Implements IVarBank

Private m_VarBankElements As Collection         '要素のCollection
Private m_FixedVarBankElements As Collection    '要素のCollection Fixedは別オブジェクトで保持。Clearメソッドで消去されない。

'Like演算子　大文字小文字無視オプション
Option Compare Text

Private Property Get IVarBank_Value(VariableName As String) As Variant
'内容:
'   登録済みの変数名の値を取得します。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]  In  1):変数名
'戻り値:
'   Variant
'
'注意事項:
'
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".GetValue")
    If Not (p_Element Is Nothing) Then
        IVarBank_Value = p_Element.val
    End If
    Set p_Element = Nothing
End Property

Private Property Let IVarBank_Value(VariableName As String, pValue As Variant)
'内容:
'   登録済みの変数に値を設定します。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]  In  1):変数名
'   [pValue]        In  2):設定値
'戻り値:
'
'注意事項:
'   設定するには ReadOnly プロパティをFalse に設定しないといけません。
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".LetValue")
    If Not (p_Element Is Nothing) Then
        If p_Element.ReadOnly Then
            Call TheError.Raise(9999, TypeName(Me) & ".LetValue", "This Variable is Now Status ReadOnly! " & VariableName)
            Exit Property
        Else
            p_Element.val = pValue
        End If
    End If
    Set p_Element = Nothing
End Property

Private Property Get IVarBank_ReadOnly(VariableName As String) As Boolean
'内容:
'   変数のReadOnlyプロパティ状態を取得します。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]  In  1):変数名
'戻り値:
'   Boolean
'
'注意事項:
'
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".GetReadOnly")
    If Not (p_Element Is Nothing) Then
        IVarBank_ReadOnly = p_Element.ReadOnly
    End If
    Set p_Element = Nothing
End Property

Private Property Let IVarBank_ReadOnly(VariableName As String, pReadOnly As Boolean)
'内容:
'
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]  In  1):変数名
'   [pReadOnly]     In  2):設定する状態
'戻り値:
'
'注意事項:
'
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".LetReadOnly")
    If Not (p_Element Is Nothing) Then
        p_Element.ReadOnly = pReadOnly
    End If
    Set p_Element = Nothing
End Property

Private Property Get IVarBank_Fixed(VariableName As String) As Boolean
'内容:
'   保持状態かどうかを取得します。
'   保持状態にすると、Clearメソッドにて破棄されません。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]  In  1):変数名
'戻り値:
'   Boolean
'
'注意事項:
'   AllClearメソッドでは破棄されます。
'

    If ExistFixed(VariableName) Then
        IVarBank_Fixed = True
        Exit Property
    ElseIf ExistNonFixed(VariableName) Then
        IVarBank_Fixed = False
        Exit Property
    Else
        Call TheError.Raise(9999, TypeName(Me) & ".GetFixed", "This Variable is not Exist! " & VariableName)
    End If
End Property

Private Property Let IVarBank_Fixed(VariableName As String, pFixed As Boolean)
'内容:
'   保持状態を設定します。
'   保持状態にすると、Clearメソッドにて破棄されません。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]  In  1):変数名
'   [pFixed]        In  2):設定する保持状態
'戻り値:
'
'注意事項:
'   AllClearメソッドでは破棄されます。
'

    If ExistFixed(VariableName) Then
        If pFixed = False Then
            Call m_VarBankElements.Add(m_FixedVarBankElements.Item(VariableName), VariableName)
            Call m_FixedVarBankElements.Remove(VariableName)
            Exit Property
        End If
    ElseIf ExistNonFixed(VariableName) Then
        If pFixed = True Then
            Call m_FixedVarBankElements.Add(m_VarBankElements.Item(VariableName), VariableName)
            Call m_VarBankElements.Remove(VariableName)
            Exit Property
        End If
    Else
        Call TheError.Raise(9999, TypeName(Me) & ".LetFixed", "This Variable is not Exist! " & VariableName)
    End If
End Property

Private Property Get IVarBank_Comment(VariableName As String) As String
'内容:
'   変数のコメントを参照します。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]  In  1):変数名
'戻り値:
'   String
'
'注意事項:
'
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".GetComment")
    If Not (p_Element Is Nothing) Then
        IVarBank_Comment = p_Element.Comment
    End If
    Set p_Element = Nothing
End Property

Private Property Let IVarBank_Comment(VariableName As String, pComment As String)
'内容:
'   変数のコメントを変更します。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]  In  1):変数名
'   [pComment]      In  2):コメント
'戻り値:
'
'注意事項:
'
'

    Dim p_Element As CVarBankElement
    Set p_Element = GetVariable(VariableName, ".LetComment")
    If Not (p_Element Is Nothing) Then
        p_Element.Comment = pComment
    End If
    Set p_Element = Nothing
End Property

Private Property Get IVarBank_Count() As Long
'内容:
'   登録されている変数の数を返します。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   なし
'戻り値:
'   Long
'
'注意事項:
'
'

    If Not (m_VarBankElements Is Nothing) Then                           '登録変数群があるかどうか
        IVarBank_Count = m_VarBankElements.Count
    End If
    If Not (m_FixedVarBankElements Is Nothing) Then                   '登録変数群(固定版)があるかどうか
        IVarBank_Count = IVarBank_Count + m_FixedVarBankElements.Count
    End If
End Property

Private Function IVarBank_Add(VariableName As String, FirstValue As Variant, Optional FixedStatus As Boolean, Optional Comment As String) As Boolean
'内容:
'   変数名と値を登録します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月23日
'パラメータ:
'   [VariableName]       In  1):登録名
'   [Value]              In  2):登録値
'   [Fixed]              In  3):保持状態
'   [Comment]            In  4):登録に関してのコメント
'戻り値:
'
'注意事項:
'

    If IVarBank_IsExist(VariableName) Then                                                          '登録変数群に変数があるかどうか
        Call TheError.Raise(9999, TypeName(Me) & ".Add", "This Variable Already is Exist! " & VariableName)
        IVarBank_Add = False
        Exit Function
    Else                                                                                    '変数に登録されていない
        Dim pVariable As CVarBankElement
        Set pVariable = New CVarBankElement
        Call pVariable.Initialize(VariableName, FirstValue, Comment)
        If FixedStatus Then
            Call m_FixedVarBankElements.Add(pVariable, VariableName)
        Else
            Call m_VarBankElements.Add(pVariable, VariableName)
        End If
    End If
    
    IVarBank_Add = True
    
End Function

Private Function IVarBank_IsExist(VariableName As String) As Boolean
'内容:
'   該当するVariableNameが登録済みかどうかをBooleanで返す。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]       In  1):変数名
'戻り値:
'   Boolean
'注意事項:
'
    
    If ExistFixed(VariableName) Then
        IVarBank_IsExist = True
        Exit Function
    ElseIf ExistNonFixed(VariableName) Then
        IVarBank_IsExist = True
        Exit Function
    End If
    
    IVarBank_IsExist = False

End Function

Private Function IVarBank_Search(ByVal SearchPattern As String, Optional NotFlag As Boolean = False) As String
'内容:
'   登録されている変数から、SearchPatternに合う変数を探します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月15日
'パラメータ:
'   [SearchPattern] In  1):探す文字
'   [NotFlag]       In  2):True:含まない文字を返す。False:含む文字を返す。
'戻り値:
'   String
'
'注意事項:
'   文字の大小違いは無視されます
'   SearchPatternはLike演算子で比較されます
'
'  ○探したいSearchPattern例
'   全部探したい場合                  ："*"
'   頭にTestが付く文字を探したい場合  ："Test*"
'   末尾に7が付く文字を探したい場合   ："*7"
'   HLを含む文字を探したい場合        ："*HL*"


    Dim i As Integer
    Dim var As Variant
    Dim keyName As String
    Dim retStr As String
    
    If Not (m_VarBankElements Is Nothing) Then
        For Each var In m_VarBankElements
            If i > 5000 Then
                MsgBox "SearchVariable Function. Too Much Variable!"
                Exit Function
            End If
            If (var.Name Like SearchPattern) Xor NotFlag Then
                retStr = retStr & var.Name & ","
                i = i + 1
            End If
        Next var
    End If
    
    If Not (m_FixedVarBankElements Is Nothing) Then
        For Each var In m_FixedVarBankElements
            If i > 5000 Then
                MsgBox "SearchVariable Function. Too Much Variable!"
                Exit Function
            End If
            If (var.Name Like SearchPattern) Xor NotFlag Then
                retStr = retStr & var.Name & ","
                i = i + 1
            End If
        Next var
    End If
    
    IVarBank_Search = retStr

End Function

Private Sub IVarBank_Clear()
'内容:
'   登録された変数群を初期化します。Fixed版は初期化しません。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月14日
'パラメータ:
'戻り値:
'
'注意事項:
'
    Set m_VarBankElements = Nothing
    Set m_VarBankElements = New Collection
End Sub

Private Sub IVarBank_AllClear()
'内容:
'   登録された変数群を全て初期化します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月14日
'パラメータ:
'戻り値:
'
'注意事項:
'
    Set m_VarBankElements = Nothing
    Set m_VarBankElements = New Collection
    Set m_FixedVarBankElements = Nothing
    Set m_FixedVarBankElements = New Collection
End Sub

Private Function ExistNonFixed(VariableName As String) As Boolean
'内容:
'   該当するVariableNameがNonFixed collectionに登録済みかどうかをBooleanで返す。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]       In  1):変数名
'戻り値:
'   Boolean
'注意事項:
'
    Dim p_Element As CVarBankElement
    
    On Error GoTo NOT_FOUND_NONFIX
    Set p_Element = m_VarBankElements.Item(VariableName)
    ExistNonFixed = True
    Set p_Element = Nothing
    Exit Function
    
NOT_FOUND_NONFIX:
    
    ExistNonFixed = False

End Function

Private Function ExistFixed(VariableName As String) As Boolean
'内容:
'   該当するVariableNameがFixed collectionに登録済みかどうかをBooleanで返す。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]       In  1):変数名
'戻り値:
'   Boolean
'注意事項:
'
    Dim p_Element As CVarBankElement
    
    On Error GoTo NOT_FOUND_FIX
    Set p_Element = m_FixedVarBankElements.Item(VariableName)
    ExistFixed = True
    Set p_Element = Nothing
    Exit Function
    
NOT_FOUND_FIX:
    
    ExistFixed = False

End Function

Private Function GetVariable(VariableName As String, MethodName As String) As CVarBankElement
'内容:
'   該当する変数が存在するかどうか。存在すればオブジェクトを返す。無ければNothing
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [VariableName]       In  1):変数名
'戻り値:
'   CVarBankElement
'
'注意事項:
'

    If ExistFixed(VariableName) Then
        Set GetVariable = m_FixedVarBankElements.Item(VariableName)
        Exit Function
    ElseIf ExistNonFixed(VariableName) Then
        Set GetVariable = m_VarBankElements.Item(VariableName)
        Exit Function
    End If
    
    Call TheError.Raise(9999, TypeName(Me) & MethodName, "This Variable is not Exist! " & VariableName)
    Set GetVariable = Nothing

End Function

Private Sub Class_Initialize()
    Set m_VarBankElements = New Collection
    Set m_FixedVarBankElements = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_VarBankElements = Nothing
    Set m_FixedVarBankElements = Nothing
End Sub

