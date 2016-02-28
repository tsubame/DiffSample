VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTextTreeViewWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ワークシートナビゲーションツリービューのテキスト用ライタークラス
'
'目的:
'   ツリービューデータをテキストへ出力する
'
'   Revision History:
'   Data        Description
'   2009/06/15  作成
'
'作成者:
'   0145206097
'
Option Explicit

Implements IParameterWriter
Implements IFileStream

Dim mFileNo As Integer
Dim mFileName As String

Private Const FILE_NAME = "WorksheetTreeViewer.csv"

Public Sub OpenFile(ByVal filePath As String)
    On Error GoTo ErrorHandler
    mFileNo = FreeFile
    mFileName = filePath & "\" & getToday & "_" & FILE_NAME
    Open mFileName For Output As mFileNo
    Exit Sub
ErrorHandler:
    Close mFileNo
    Err.Raise 9999, "CTextTreeViewWriter.OpenFile", "Opening File Is Failed"
End Sub

Public Sub CloseFile()
    Close mFileNo
End Sub

Private Sub IParameterWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
    '### このインターフェイス機能の動作を無視する ###################
End Sub

Private Sub IParameterWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
    '### このライターはこのインターフェイス機能を保障しない #########
    Err.Raise 999, "CTextTreeViewWriter.AsIParameterWriter.WriteAsDouble", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
    '### このライターはこのインターフェイス機能を保障しない #########
    Err.Raise 999, "CTextTreeViewWriter.AsIParameterWriter.WriteAsInteger", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
    '### このライターはこのインターフェイス機能を保障しない #########
    Err.Raise 999, "CTextTreeViewWriter.AsIParameterWriter.WriteAsLong", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
    '### ファイルへデータ出力 #######################################
    Select Case dataName:
        Case SUPPLIER_NAME:
            Print #mFileNo, strData;
        Case TOOL_NAME:
            Print #mFileNo, "," & strData;
        Case sheet_name:
            Print #mFileNo, ",," & strData;
    End Select
End Sub

Private Sub IFileStream_MoveNext()
    '### データの改行 ###############################################
    Print #mFileNo,
End Sub

Private Function IFileStream_IsEOR() As Boolean
    '### このインターフェイス機能の動作を無視する ###################
End Function

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    '### このライターはこのインターフェイス機能を保障しない #########
    Err.Raise 999, "CTextTreeViewWriter.AsIFileStream.SetLocation", "Does Not Support This Method !"
End Sub

Private Sub IFileStream_Rewind()
    '### このライターはこのインターフェイス機能を保障しない #########
    Err.Raise 999, "CTextTreeViewWriter.AsIFileStream.Rewind", "Does Not Support This Method !"
End Sub

Public Function AsIParameterWriter() As IParameterWriter
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IParameterWriter型オブジェクト
'
'注意事項:
'
    Set AsIParameterWriter = Me
End Function

Public Function AsIFileStream() As IFileStream
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IFileStreamr型オブジェクト
'
'注意事項:
'
    Set AsIFileStream = Me
End Function

Private Function getToday() As String
    getToday = Format(DateTime.Now, "yymmdd") & "_" & Format(DateTime.Now, "hhnnss")
End Function
