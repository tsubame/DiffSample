VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctionConstantsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' メンバ変数
Private Name_ As String
Private Dic_ As Object

' ヘッダ情報
Private Const NameHeader = "_ConstantDefTable_"
Private Const ValueHeader = "_Value_"

Private Sub Class_Initialize()
    Name_ = ""
    Set Dic_ = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set Dic_ = Nothing
End Sub

' 初期化
' @param values テーブルの内容が記述されたセル領域
Public Sub Initialize(values As Range)
    
    Dim ri As Integer
    Dim ci As Integer
    Dim ni As Integer ' Name(Key) Index
    Dim vi As Integer ' Value Index
    Dim str_tmp As String
    
    On Error GoTo ReleaseObjects
    
    ' 名前
    Name_ = Trim(values.Cells(1, 1))
    
    ' 見出しの位置を取得
    ' 2列目が"_"で始まる行の取得
    For ri = 1 To values.Rows.Count
        If InStr(1, Trim(values.Cells(ri, 2)), "_", vbBinaryCompare) = 1 Then
            Exit For
        End If
    Next
    
    For ci = 2 To values.Columns.Count
        str_tmp = Trim(values.Cells(ri, ci))
        If StrComp(str_tmp, NameHeader, vbBinaryCompare) = 0 Then
            ni = ci
        ElseIf StrComp(str_tmp, ValueHeader, vbBinaryCompare) = 0 Then
            vi = ci
        End If
    Next
    
    ' 見出し行の次行からデータとする
    For ri = ri + 1 To values.Rows.Count
        ' Dictionaryに追加
        Call Dic_.Add(values.Cells(ri, ni).Value, values.Cells(ri, vi).Value)
    Next
    
ReleaseObjects:
    ' Nothing
        
End Sub

' 名前の取得
' @return 名前
Public Function GetName() As String
    GetName = Name_
End Function

' キーの一覧を取得
' @return キー(定数名)一覧
Public Function GetNames() As String()

    Dim ii As Integer
    Dim values() As String
    
    If IsNull(Dic_) Or Dic_.Count <= 0 Then
        Exit Function
    End If
    
    ReDim values(Dic_.Count - 1)
    
    ' コピー
    For ii = 0 To Dic_.Count - 1
        values(ii) = Dic_.Keys()(ii)
    Next
    
    GetNames = values
        
End Function

' 値の取得
' @param key キー(定数名)
' @return 定数
Public Function GetValue(key As String) As Variant
    
    If IsNull(Dic_) Or Dic_.Count <= 0 Or Not Dic_.Exists(key) Then
        Exit Function
    End If
    
    GetValue = Dic_.Item(key)
        
End Function
