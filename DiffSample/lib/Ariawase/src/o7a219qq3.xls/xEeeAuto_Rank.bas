Attribute VB_Name = "xEeeAuto_Rank"
'概要:
'   Rank Sheetの存在確認
'
'目的:
'   Rankシートを見に行ってあったらTrueを返す
'
'作成者:
'   2012/03/12 D.Maruyama
'   2012/11/12 H.Arikawa  RankSheetに項目記載があるかチェックするルーチンを追加。

Private Enum EeeAutoRankState
    UNKOWN
    INITALIZED
End Enum

Private m_IsRankSheet_Exist As Boolean
Private m_State As EeeAutoRankState

Private Const RANK_SHEET_NAME As String = "rank_sheet"

Option Explicit

'内容:
'   このモジュールの初期化
'
'備考:
'   RANKSHEETのあるないを判断する
'
Public Sub InitializeEeeAutoRank()

    m_IsRankSheet_Exist = False

    Dim mySheet As Worksheet
    
    For Each mySheet In ThisWorkbook.Worksheets
        If mySheet.Name = RANK_SHEET_NAME Then
            m_IsRankSheet_Exist = False
            Exit For
        End If
    Next mySheet
    
    m_State = INITALIZED
        
End Sub

'内容:
'   RANK処理をするかどうかを返す
'
'備考:
'   RANKSHEETのある場合 True、ない場合 False
'
'有川コメント
'   RANKSHEETは全タイプ挿入されるので、中身の記載があるかをチェックする。
'   セル指定で値が入っているかをチェックさせる。

Public Function IsRankEnable() As Boolean

    Dim tname As String

    If m_State <> INITALIZED Then
        Err.Raise 9999, "IsRankEnable", "xEeeAutoRank is not Initialized!"
        IsRankEnable = False
        Exit Function
    End If
        
    IsRankEnable = m_IsRankSheet_Exist
    
    tname = Worksheets("Tenken").Range("B9").Value
    
    If tname = "" Then
        IsRankEnable = False
    End If
    
End Function
