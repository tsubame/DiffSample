VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CActionExecutionHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   Actionの実行履歴の1件相当の情報
'
'目的:
'   Actionの実行履歴の1件相当の情報を格納するデータオブジェクト。下記を保持。
'     ・記録番号
'     ・記録日
'     ・記録時間
'     ・時刻
'     ・トリガー（Start/End）
'     ・ActionInfo型オブジェクト
'     ・Error番号（0:Success、Not 0:Error）
'     ・実行ステータス（CSVには出力されない）状態：未実行、開始済み、実行済み、エラー
'
'作成者:
'   a_oshima
'
'前提条件:
'   CActionInfoクラスを利用
'   グローバルオブジェクト TheError:CErrManager が存在すること
'
'利用手順:
'


Option Explicit

Private mRecordNumber As String
Private mTrigger As String
Private mRecordDate As String
Private mRecordTime As String
Private mLapTime As Single
Private mActionInfo As CActionInfo
Private mErrorNo As Long

Property Let Trigger(ByVal pTrigger As String)
    mTrigger = pTrigger
End Property
Property Get Trigger() As String
    Trigger = mTrigger
End Property

Property Get RecordNumber() As String
    RecordNumber = mRecordNumber
End Property
Property Let RecordNumber(ByRef pId As String)
    mRecordNumber = pId
End Property

Property Get RecordDate() As String
    RecordDate = mRecordDate
End Property
Property Let RecordDate(ByRef pDay As String)
'パラメータ:
'   [pDay]  IN  String型:  日付を表す文字列。日付でない場合は例外発生

    If Not IsDate(pDay) Then
        Call TheError.Raise(9999, TypeName(Me) & ".RecordDate", "Not Date:" & pDay)
    End If
    mRecordDate = pDay
End Property

Property Get RecordTime() As String
    RecordTime = mRecordTime
End Property
Property Let RecordTime(ByRef pTim As String)
    mRecordTime = pTim
End Property

Property Get LapTime() As Single
    LapTime = mLapTime
End Property
Property Let LapTime(ByRef pTim As Single)
'パラメータ:
'   [pTim]  IN  Single型:  範囲：0または正の値。負の場合は例外発生
    
    If pTim < 0! Then
        Call TheError.Raise(9999, TypeName(Me) & ".LapTime", "Tim must 0 or positive value:" & pTim)
    End If
    mLapTime = pTim
End Property

Property Get ActionInfo() As CActionInfo
    Set ActionInfo = mActionInfo
End Property
Property Set ActionInfo(ByRef pActionInfo As CActionInfo)
    Set mActionInfo = pActionInfo
End Property

Property Get ErrorNo() As Long
    ErrorNo = mErrorNo
End Property

Property Let ErrorNo(ByRef pErrorNo As Long)
    mErrorNo = pErrorNo
End Property

Property Get ReturnCode() As String
    If mTrigger = "Start" Then
        ReturnCode = "-"
    Else
        ReturnCode = CStr(mErrorNo)
    End If
End Property

Private Sub Class_Initialize()
    Set mActionInfo = New CActionInfo
End Sub

Private Sub Class_Terminate()
    Set mActionInfo = Nothing
End Sub

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
    str = MakeStringMsg("[RecordNumber]") _
            & "," & MakeStringMsg("[RecordDate]") _
            & "," & MakeStringMsg("[RecordTime]") _
            & "," & MakeStringMsg("[LapTime(s)]") _
            & "," & MakeStringMsg("[Trigger]") _
            & "," & mActionInfo.CSVIndex(ArgCount) _
            & "," & MakeStringMsg("[ReturnCode]")
            
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
    str = MakeStringMsg(mRecordNumber) _
            & "," & MakeStringMsg(mRecordDate) _
            & "," & MakeStringMsg(mRecordTime) _
            & "," & mLapTime _
            & "," & MakeStringMsg(mTrigger) _
            & "," & mActionInfo.ToCSVString(ArgCount) _
            & "," & MakeStringMsg(ReturnCode())
    
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
    Dim str As String
    str = "RecordNumber:" & mRecordNumber & vbCrLf _
            & "RecordDate:" & mRecordDate & vbCrLf _
            & "RecordTime:" & mRecordTime & vbCrLf _
            & "LapTime(s):" & CStr(mLapTime) & vbCrLf _
            & "Trigger:" & mTrigger & vbCrLf _
            & "ActionInfo:[" & vbCrLf _
                & mActionInfo.ToString(4) _
            & "]" & vbCrLf _
            & "ReturnCode:" & ReturnCode() & vbCrLf

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




