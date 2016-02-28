VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CErrInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   エラー情報構造体用クラス
'
'目的:
'   エラー情報の要素をまとめる
'   エラー情報と情報提供機能を合わせて管理する
'
'作成者:
'   SLSI今手
'
'注意点:
'   エラーマネージャObjectが本クラスを利用します。
'
Option Explicit

Private m_SerialNumber As Long   'エラー記録シリアル番号
Private m_Number As Long         'エラー番号
Private m_Source As String       'エラーソース
Private m_Message As String      'エラーメッセージ
Private m_HelpFile As String     'エラー時に表示するヘルプファイル
Private m_HelpContext As Long 'ヘルプファイルのトピックを表すコンテキストIDを含む整数
Private m_Date As String         'エラー発生日
Private m_Time As String         'エラー発生時間
Private Const ThisClassName = "CErrInfo" 'クラス名称
'#Pass-OK
Public Property Let SerialNumber(ByVal ErrRecordNumber As Long)
'内容:
'   エラー情報シリアル番号設定
'
'パラメータ:
'   [ErrRecordNumber]  In  エラー記録番号
'
'戻り値:
'
'注意事項:
'
    m_SerialNumber = ErrRecordNumber

End Property

'#Pass-OK
Public Property Get SerialNumber() As Long
'内容:
'   エラー情報シリアル番号取得
'
'パラメータ:
'
'戻り値:
'   エラー記録時のシリアル番号
'
'注意事項:
'
    SerialNumber = m_SerialNumber

End Property

'#Pass-OK
Public Property Let ErrDate(ByVal ErrDate As String)
'内容:
'   エラー発生日設定
'
'パラメータ:
'   [ErrDate]  In  エラー発生日文字列
'
'戻り値:
'
'注意事項:
'
    m_Date = ErrDate

End Property

'#Pass-OK
Public Property Get ErrDate() As String
'内容:
'   エラー発生日取得
'
'パラメータ:
'
'戻り値:
'   エラー発生日文字列
'
'注意事項:
'
    ErrDate = m_Date

End Property

'#Pass-OK
Public Property Let ErrTime(ByVal ErrTime As String)
'内容:
'   エラー発生時間設定
'
'パラメータ:
'   [ErrTime]  In  エラー発生時間文字列
'
'戻り値:
'
'注意事項:
'
    m_Time = ErrTime

End Property

'#Pass-OK
Public Property Get ErrTime() As String
'内容:
'   エラー発生時間取得
'
'パラメータ:
'
'戻り値:
'   エラー発生時間文字列
'
'注意事項:
'
    ErrTime = m_Time

End Property

'#Pass-OK
Public Property Let HelpContext(ByVal HelpContextNumber As Long)
'内容:
'   エラー時ヘルプファイルのヘルプID設定
'
'パラメータ:
'   [HelpContextNumber]  In  ヘルプID番号
'
'戻り値:
'
'注意事項:
'
    m_HelpContext = HelpContextNumber

End Property

'#Pass-OK
Public Property Get HelpContext() As Long
'内容:
'   エラー時ヘルプファイルのヘルプID取得
'
'パラメータ:
'
'戻り値:
'   ヘルプID番号
'
'注意事項:
'
    HelpContext = m_HelpContext

End Property

'#Pass-OK
Public Property Let HelpFile(ByVal HelpFile As String)
'内容:
'   エラー時に表示するヘルプファイル名設定
'
'パラメータ:
'   [HelpFile]  In  ヘルプファイル名文字列
'
'戻り値:
'
'注意事項:
'   ヘルプファイル名は、FullPathで設定します
'   (例) d:\HelpFiles\JobErrHelp.mht
'
    m_HelpFile = HelpFile

End Property

'#Pass-OK
Public Property Get HelpFile() As String
'内容:
'   エラー時に表示するヘルプファイル名取得
'
'パラメータ:
'
'戻り値:
'   ヘルプファイル文字列
'
'注意事項:
'
    HelpFile = m_HelpFile

End Property

'#Pass-OK
Public Property Let Number(ByVal ErrNumber As Long)
'内容:
'   エラー番号設定
'
'パラメータ:
'   [ErrNumber]  In  エラー番号値
'
'戻り値:
'
'注意事項:
'   エラー番号には、65535以下の値を設定して下さい。
'
    
    If ErrNumber > 65535 Then
        Call Err.Raise(9999, ThisClassName, _
        "error number=" & ErrNumber & " is a value that cannot be set.  [error number <= 65535]")
    End If
    
    m_Number = ErrNumber

End Property

'#Pass-OK
Public Property Get Number() As Long
'内容:
'   エラー番号取得
'
'パラメータ:
'
'戻り値:
'   エラー番号値
'
'注意事項:
'
    Number = m_Number

End Property

'#Pass-OK
Public Property Let Message(ByVal ErrorMessage As String)
'内容:
'   エラーメッセージ
'
'パラメータ:
'   [ErrorMessage]  In  エラー時のメッセージ
'
'戻り値:
'
'注意事項:
'
    m_Message = ErrorMessage

End Property

'#Pass-OK
Public Property Get Message() As String
'内容:
'   エラーメッセージ
'
'パラメータ:
'
'戻り値:
'   エラー時のメッセージ
'
'注意事項:
'
    Message = m_Message

End Property

'#Pass-OK
Public Property Let Source(ByVal ErrSource As String)
'内容:
'   エラーソース名
'
'パラメータ:
'   [ErrSource]  In  エラーソース名称
'
'戻り値:
'
'注意事項:
'
    m_Source = ErrSource

End Property

'#Pass-OK
Public Property Get Source() As String
'内容:
'   エラーソース名
'
'パラメータ:
'
'戻り値:
'   エラーソース名称
'
'注意事項:
'
    Source = m_Source

End Property

'#Pass-OK
Public Property Get CsvInformation() As String
'内容:
'   エラー情報をCSV（カンマ区切り形式）文字列で提供する
'
'パラメータ:
'
'戻り値:
'   CSV形式エラー情報
'
'注意事項:
'  出力Format仕様
'   [SerialNumber],"[Date]","[Time]",[Number],"[Source]","[Message]","[HelpFile]",[HelpContext]
'
'    ＊改行文字、カンマ、ダブルコーテーションはエスケープ文字に置き換えられます。
'      vbCrLf-->\r\n ,vbCr --> \r, vbLf --> \n
'      , --> \,
'      " --> ""
'
    Dim CsvMsg As String
            
    With Me
        CsvMsg = .SerialNumber & "," & _
         MakeStringMsg(.ErrDate) & "," & _
         MakeStringMsg(.ErrTime) & "," & _
        .Number & "," & _
         MakeStringMsg(.Source) & "," & _
         MakeStringMsg(.Message) & "," & _
         MakeStringMsg(.HelpFile) & "," & _
        .HelpContext & vbCrLf
    End With
    
    CsvInformation = CsvMsg

End Property

'#Pass-OK
Public Property Get Information() As String
'内容:
'   エラー情報提供用関数
'
'パラメータ:
'
'戻り値:
'   エラー情報表示用メッセージ
'
'注意事項:
'   出力される要素は
'   日時、シリアル番号、エラー番号、エラーソース、エラーメッセージとなります
'
    Dim PrintMsg As String
    
    'エラーメッセージ作成
    With Me
        PrintMsg = "[Date] " & .ErrDate & " " & .ErrTime & vbCrLf & _
        "[Serial] " & .SerialNumber & vbCrLf & _
        "[ErrNum] " & .Number & vbCrLf & _
        "[Src] " & MakeStringMsg(.Source) & vbCrLf & _
        "[Msg] " & MakeStringMsg(.Message) & vbCrLf
    End With
    
    Information = PrintMsg

End Property

'#Pass-OK
Public Sub DebugPrint()
'内容:
'   エラー情報をイミディエイトWindowに出力する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Debug.Print Me.Information

End Sub

'#Pass-OK
Public Sub DispMsgBox()
'内容:
'   エラー情報をメッセージBOXに出力する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Call MsgBox(Me.Information, vbInformation, "Error information")

End Sub

'-- 以下、Private機能群 -------------------------------------------------------

'#Pass-OK
Private Sub Class_Initialize()
'constructor：内部データ初期化
    m_SerialNumber = 0
    m_Number = 0
    m_Source = ""
    m_Message = ""
    m_HelpFile = ""
    m_HelpContext = 0
    m_Date = ""
    m_Time = ""
End Sub

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
        makeMsg = Chr(34) & makeMsg & Chr(34)
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



