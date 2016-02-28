VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CActionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   Actionの情報
'
'目的:
'   Actionの情報を格納するデータオブジェクト。下記情報を持つ。
'     ・テストインスタンス名
'     ・テストプロシージャ名（UserMacro）
'     ・Action名（フレームワーク名）
'     ・AcquireGroup名
'     ・Acquireインスタンス名
'     ・CAcquireFunctionInfo型オブジェクト（マクロのパラメータを保持）
'
'作成者:
'   a_oshima
'
'前提条件:
'   CAcquireFunctionInfoクラスを利用
'   CAcquireFunctionInfoクラスの下記メソッドを利用
'      ・Public Sub CopyObj(OrgObj As CAcquireFunctionInfo)
'　　　・Public Property Get ArgParameterCount() As Long
'　　　・Public Property Get Arg(ByVal pArgNumber As Long) As Variant
'
'利用手順:



Option Explicit

Private mTestInstanceName As String
Private mMacroName As String
Private mFrameworkName As String
Private mAcquireGroupName As String
Private mAcquireInstanceName As String
Private mParameters As CAcquireFunctionInfo




Property Get TestInstanceName() As String
    TestInstanceName = mTestInstanceName
End Property
Property Let TestInstanceName(ByRef pName As String)
    mTestInstanceName = pName
End Property

Property Get FrameworkName() As String
    FrameworkName = mFrameworkName
End Property
Property Let FrameworkName(ByRef pName As String)
    mFrameworkName = pName
End Property

Property Get AcquireGroupName() As String
    AcquireGroupName = mAcquireGroupName
End Property
Property Let AcquireGroupName(ByRef pName As String)
    mAcquireGroupName = pName
End Property

Property Get AcquireInstanceName() As String
    AcquireInstanceName = mAcquireInstanceName
End Property
Property Let AcquireInstanceName(ByRef pName As String)
    mAcquireInstanceName = pName
End Property

Property Get macroName() As String
    macroName = mMacroName
End Property
Property Let macroName(ByRef pName As String)
    mMacroName = pName
End Property

Property Get Parameters() As CAcquireFunctionInfo
    Set Parameters = mParameters
End Property
Property Set Parameters(ByRef pParameters As CAcquireFunctionInfo)
'    'コピーを格納
'    Set mParameters = New CAcquireFunctionInfo
'    Call mParameters.CopyObj(pParameters)
    '参照元を格納
    Set mParameters = pParameters
End Property

Private Sub Class_Initialize()
    Set mParameters = New CAcquireFunctionInfo
End Sub

Private Sub Class_Terminate()
    Set mParameters = Nothing
End Sub

Public Function Clone() As CActionInfo
'内容:
'   自オブジェクトの複製を返す（Deep Copy）
'
'戻り値:
'   自オブジェクトの複製
'
'備考:
'
'
    Dim objCopy As CActionInfo
    Set objCopy = New CActionInfo
    objCopy.TestInstanceName = mTestInstanceName
    objCopy.FrameworkName = mFrameworkName
    objCopy.AcquireGroupName = mAcquireGroupName
    objCopy.AcquireInstanceName = mAcquireInstanceName
    objCopy.macroName = mMacroName
    Set objCopy.Parameters = mParameters.Clone
    
    Set Clone = objCopy
End Function

Public Function IsSameName(ByRef pTargetActionInfo As CActionInfo) As Boolean
'内容:
'   同型オブジェクトと各Propertyをそれぞれ比較する
'   （パラメータの情報は比較しない）
'
'パラメータ:
'   [pTargetActionInfo]  IN  CActionInfo型:  比較対象のActionInfo
'
'戻り値:
'   True：パラメータを除く全Propertyが同じ、False：Propertyが１つでも異なる
'
'備考:
'
'
    If mTestInstanceName = pTargetActionInfo.TestInstanceName _
            And mFrameworkName = pTargetActionInfo.FrameworkName _
            And mAcquireGroupName = pTargetActionInfo.AcquireGroupName _
            And mAcquireInstanceName = pTargetActionInfo.AcquireInstanceName _
            And mMacroName = pTargetActionInfo.macroName Then
        IsSameName = True
    Else
        IsSameName = False
    End If
End Function

Public Function CSVIndex(Optional ByVal ArgCount As Long = 0&) As String
'内容:
'   CSVのヘッダを返す。
'
'戻り値:
'   CSVのヘッダ（カンマ区切りの文字列。最後に改行は付加しない）
'
'備考:
'   各値はMakeStringMsgで特殊文字をエスケープ済み
'
    
    Dim str As String
    str = MakeStringMsg("[TestInstanceName]") _
            & "," & MakeStringMsg("[TestProcedureName]") _
            & "," & MakeStringMsg("[ActionName]") _
            & "," & MakeStringMsg("[GroupName]") _
            & "," & MakeStringMsg("[InstanceName]")

    Dim loopEnd As Long
    loopEnd = ArgCount - 1
    If ArgCount < mParameters.ArgParameterCount Then
        loopEnd = mParameters.ArgParameterCount - 1
    End If
    
    Dim i As Long
    For i = 0& To loopEnd
        str = str & "," & MakeStringMsg("[Arg" & i & "]")
    Next i
    
    CSVIndex = str
End Function
Public Function ToCSVString(Optional ByVal ArgCount As Long = 0&) As String
'内容:
'   CSVのデータを返す。
'
'戻り値:
'   CSVのデータ（カンマ区切りの文字列。最後に改行は付加しない）
'
'備考:
'   各値はMakeStringMsgで特殊文字をエスケープ済み
'
    Dim str As String
    str = MakeStringMsg(mTestInstanceName) _
            & "," & MakeStringMsg(mMacroName) _
            & "," & MakeStringMsg(mFrameworkName) _
            & "," & MakeStringMsg(mAcquireGroupName) _
            & "," & MakeStringMsg(mAcquireInstanceName)

    Dim i As Long
    For i = 0& To mParameters.ArgParameterCount - 1
        str = str & "," & MakeStringMsg(mParameters.Arg(i))
    Next i
    
    If ArgCount > mParameters.ArgParameterCount Then
        For i = mParameters.ArgParameterCount To ArgCount - 1
            str = str & ","
        Next i
    End If
    
    ToCSVString = str
End Function
Public Function ToString(Optional ByVal pIndentWidth As Long = 0&) As String
'内容:
'   デバッグ用。データを文字列に書き出す。
'
'パラメータ:
'   [pIndentWidth]  IN  Long型:  インデント幅（デフォルト：0）
'
'戻り値:
'   文字列に整形済みのデータ
'
'備考:
'
'
    Dim str As String
    str = "TestInstanceName:" & mTestInstanceName & vbCrLf _
            & "TestProcedureName:" & mMacroName & vbCrLf _
            & "ActionName:" & mFrameworkName & vbCrLf _
            & "GroupName:" & mAcquireGroupName & vbCrLf _
            & "InstanceName:" & mAcquireInstanceName & vbCrLf _
            & "Param:[" & vbCrLf
    
    Dim i As Long
    For i = 0& To mParameters.ArgParameterCount - 1
        str = str & Space$(4) & "Arg" & i & ":" & mParameters.Arg(i) & vbCrLf
    Next i
    
    str = str & "]" & vbCrLf
    
    If pIndentWidth > 0 Then
        ToString = Indent(str, pIndentWidth)
    Else
        ToString = str
    End If
End Function

Public Function CsvInformation(Optional ByVal ArgCount As Long = 0&) As String
'内容:
'   CSVのデータを返す。
'
'戻り値:
'   CSVのデータ（カンマ区切りの文字列。最後に改行を付加）
'
'備考:
'   各値はMakeStringMsgで特殊文字をエスケープ済み
'
    CsvInformation = ToCSVString(ArgCount) & vbCrLf
End Function
Public Function Information() As String
'内容:
'   デバッグ用。データを文字列に書き出す。
'
'戻り値:
'   文字列に整形済みのデータ（最後に文字列を付加）
'
'備考:
'
'
    Information = ToString() & vbCrLf
End Function




Private Function Indent(ByVal pMessage As String, ByVal width As Long) As String
    Dim lines As Variant
    lines = Split(pMessage, vbCrLf)
    Dim i As Long
    Dim Msg As String
    For i = LBound(lines) To UBound(lines) - 1
        Msg = Msg & Space$(width) & lines(i) & vbCrLf
    Next i
    If lines(UBound(lines)) <> vbNullString Then
        Msg = Msg & Space$(width) & lines(UBound(lines))
    End If
    Indent = Msg
End Function

'#Pass-OK
Private Function MakeStringMsg(Message As String) As String
'CSVの文字列要素を作成する関数
    Dim makeMsg As String
        
    If Len(Message) > 0 Then
        '1:CR、LF、CR+LF等の改行文字を置き換える
        '2:文字列のカンマを \, に置き換える
        '3:文字列のダブルコーテーションを ""に置き換える
        makeMsg = ReplaceDoubleQuotation(ReplaceComma(ReplaceCrLf(Message)))
        
        '文字列パラメータの長さが1以上の場合は、文字列をダブルコーテーションでくくる
        makeMsg = Chr$(34) & makeMsg & Chr$(34)
        MakeStringMsg = makeMsg
    
    Else
        MakeStringMsg = Message
    End If
    
End Function

'#Pass-OK
Private Function ReplaceCrLf(Message As String) As String
'文字列に含まれている改行文字を置き換えて返す関数
'vbCrLf-->\r\n ,vbCr --> \r, vbLf --> \n
    Dim makeMsg As String
    makeMsg = Replace(Replace(Replace(Message, vbCrLf, "\r\n"), vbCr, "\r"), vbLf, "\n")
    ReplaceCrLf = makeMsg
End Function

'#Pass-OK
Private Function ReplaceDoubleQuotation(Message As String) As String
'文字列に含まれているDoubleQuotationを置き換えて返す関数
'" --> ""
    Dim makeMsg As String
    makeMsg = Replace(Message, """", """" & """")
    ReplaceDoubleQuotation = makeMsg
End Function

'#Pass-OK
Private Function ReplaceComma(Message As String) As String
'文字列に含まれているCommaを置き換えて返す関数
', --> \,
    Dim makeMsg As String
    makeMsg = Replace(Message, ",", "\,")
    ReplaceComma = makeMsg
End Function

'Private Function ReplaceBackSlash(Message As String) As String
''文字列に含まれているBackslashを置き換えて返す関数(未使用)
''\ --> \\
'    Dim makeMsg As String
'    makeMsg = Replace(Message, "\", "\\")
'    ReplaceBackSlash = makeMsg
'End Function



