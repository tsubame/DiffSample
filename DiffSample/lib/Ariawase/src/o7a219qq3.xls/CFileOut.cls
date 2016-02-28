VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ファイルの出力管理クラス
'目的:
'   ファイル出力のコーディング効率向上。
'   インスタンスをファイルの窓口とする。
'使い方:
'   ・New でインスタンス生成
'   ・SetFileName で保存ファイル名設定。
'   ・以下の命令でテキストをファイルに出力
'     ファイル入出力時に使う命令は４つ
'   　①OpenFile   ：ファイルを開く
'   　②FileNum    ：出力用。Print #インスタンス名.FileNum, 出力文字列・・・ と使う
'   　③CloseFile  ：ファイルを閉じる
'
'     ④PrintOut   ：上記３つを ①→②(Print)→③まとめて実行する
'記述サンプル:
'   Dim pFileOut As CFileOut
'   Set pFileOut = New CFileOut
'   Call pFileOut.SetFileName("tmp.txt")
'   Call pFileOut.PrintOut("hogehogehoge")
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月26日

Option Explicit

Public Enum eWriteMode
    OverWrite
    AddWrite
End Enum

Private m_WriteMode As eWriteMode
Private m_FileNum As Integer
Private m_FileName As String                'フルパス付きファイル名
Private m_OrgFileName As String             'ユーザーが設定したファイル名

Public Sub SetFileName(ByVal fileName As String, Optional WriteMode As eWriteMode = AddWrite)
'内容:
'   ファイル名の設定
'   ファイル名のみ、フルパス名どちらも可。
'   ファイル名のみの場合、アクティブワークブックのディレクトリに保存。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月26日
'パラメータ:
'   [FileName]  In      1):FileNameの指定。
'戻り値:
'
'注意事項:

    m_OrgFileName = fileName
    If (fileName Like "?:*") Or (fileName Like "\\*") Then
        m_FileName = fileName                                               'ファイル名にローカルドライブ名orネットワークドライブ名を含んでいたら、名前そのまま
    Else
        m_FileName = ActiveWorkbook.Path & "\" & fileName                   'ファイル名だけなら、アクティブワークブックのフルパスを追加
    End If
    m_WriteMode = WriteMode
'    'ファイルを開けるか試してみる
'    Call OpenFile
'    Call CloseFile

End Sub

Public Property Get fileNum() As Integer
'内容:
'   このインスタンスが開いているファイル番号を返します。
'   開いてなければ、ゼロでエラー。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月26日
'パラメータ:
'   なし
'戻り値:
'   Integer
'
'注意事項:

    If m_FileNum > 0 Then
        fileNum = m_FileNum
    Else
        Call Err.Raise(9999, , "File is not opened. =>" & m_FileName)
    End If

End Property

Public Sub OpenFile()
'内容:
'   SetFileName で設定した情報でファイルを開く
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月26日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
    
    On Error GoTo ErrHandler
    m_FileNum = FreeFile                                                      '１～２５５の空いてるファイル番号を取得
    If m_WriteMode = OverWrite Then
        Open m_FileName For Output As m_FileNum                                   '上書きでファイルを開く
    Else
        Open m_FileName For Append As m_FileNum                                   '追記でファイルを開く
    End If
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    Call Err.Raise(9999, , "File Open is Failed! => " & m_FileName)

End Sub

Public Sub CloseFile()
'内容:
'   開いたファイルを閉じる
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月26日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:

    'ファイルを開いたことが無い or 閉じられている => 何もしない。
    If m_FileNum <> 0 Then
        Close #m_FileNum
        m_FileNum = 0                   '閉じたことの明示用
    End If
End Sub

Public Sub PrintOut(PrintMessage As String)
'内容:
'   このインスタンスで管理するファイルへ Open → 書き込み → Close をセットで実行
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月26日
'パラメータ:
'   [PrintMessage]  In  1):ファイルに書き込む文字
'戻り値:
'
'注意事項:

    Call OpenFile
    Print #m_FileNum, PrintMessage
    Call CloseFile

End Sub

Public Property Get GetFileName() As String
'内容:
'   このインスタンスが持っているファイル名を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月26日
'パラメータ:
'   なし
'戻り値:
'   String
'
'注意事項:
    GetFileName = m_OrgFileName
End Property

Private Sub Class_Initialize()
'内容:
'   コンストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月26日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:

    m_FileNum = 0
    m_FileName = ""
End Sub

Private Sub Class_Terminate()
'内容:
'   ファイルが閉じられてなければColse実行。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月26日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
    
    'ファイルが開かれていない or 閉じられていれば何もしない。
    If m_FileNum <> 0 Then
        Close #m_FileNum
        m_FileNum = 0
    End If
    m_FileName = ""
End Sub


