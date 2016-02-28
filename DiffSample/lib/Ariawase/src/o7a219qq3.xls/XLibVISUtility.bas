Attribute VB_Name = "XLibVISUtility"
'概要:
'   VISクラスで使用するユーティリティ関数群
'
'目的:
'   各VISクラス共通に使用するユーティリティ関数をまとめる
'
'作成者:
'   SLSI今手
'
'   XlibSTD_CommonDCMod_V01内の、共通処理を切り出し
'   エラーメッセージ表示用サブルーチン追加
'
'
'Code Checked
'Comment Checked
'

Option Explicit

Private Const ALL_SITE = -1

'#Pass
Public Function ConvertVariableToArray(ByRef DstVar As Variant) As Boolean
'内容:
'   変数を要素数が存在Site数の配列変数に変換する
'
'パラメータ:
'    [DstVar]    In   変換対象変数
'    [DstVar]    Out  変換後配列変数
'
'戻り値:
'   ステータス（正常終了=True）
'
'注意事項:
'
    Dim VarArray() As Double
    Dim site As Long

    If IsArray(DstVar) Then
        If UBound(DstVar) <> CountExistSite Then
            ConvertVariableToArray = False
        Else
            ConvertVariableToArray = True
        End If
    Else
        ReDim VarArray(CountExistSite)

        For site = 0 To UBound(VarArray)
            VarArray(site) = DstVar
        Next site
        DstVar = VarArray
        ConvertVariableToArray = True
    End If

End Function

'#Pass
Public Sub GetChanList(ByVal PinList As String, ByVal site As Long, ByVal chanType As chtype, ByRef retChannels() As Long)
'内容:
'   対象ピンのChannel番号を取得する
'
'パラメータ:
'    [PinList]        In    対象ピンリスト
'    [site]           In    Site番号
'    [chanType]       In    ChannelType
'    [retChannels()]  Out   取得したChannel番号
'
'戻り値:
'
'注意事項:
'
    Dim ChanNum As Long
    Dim siteNum As Long
    Dim errMsg As String

    Call TheExec.DataManager.GetChanList(PinList, site, chanType, retChannels, ChanNum, siteNum, errMsg)
    If errMsg <> "" Then
        Call OutputErrMsg(errMsg & " (at GetChanList)")
    End If
End Sub

'#Pass
Public Sub GetActiveChanList(ByVal PinList As String, ByVal chanType As chtype, ByRef retChannels() As Long)
'内容:
'   選択されているSiteの対象ピンのChannel番号を取得する
'
'パラメータ:
'    [PinList]         In    対象ピンリスト
'    [chanType]        In    ChannelType
'    [retChannels()]   Out   取得したChannel番号
'
'戻り値:
'
'注意事項:
'
    Dim ChanNum As Long
    Dim siteNum As Long
    Dim errMsg As String

    Call TheExec.DataManager.GetChanListForSelectedSites(PinList, chanType, retChannels, ChanNum, siteNum, errMsg)
    If errMsg <> "" Then
        Call OutputErrMsg(errMsg & " (at GetActiveChanList)")
    End If
End Sub

'#Pass
Public Function CountExistSite() As Long
'内容:
'   存在するSite数を取得する
'
'パラメータ:
'
'戻り値:
'   存在Site数
'
'注意事項:
'
    CountExistSite = TheExec.sites.ExistingCount - 1

End Function

'#Pass
Public Function CountActiveSite() As Long
'内容:
'   Active Site数を取得する
'
'パラメータ:
'
'戻り値:
'   ActiveSite数
'
'注意事項:
'   シリアルLOOP中は戻り値=1となる
'
    With TheExec.sites
        If .InSerialLoop Then
            CountActiveSite = 1
        Else
            CountActiveSite = .ActiveCount
        End If
    End With
    
End Function

'#Pass
Public Function IsActiveSite(ByVal site As Long) As Boolean
'内容:
'   SiteがActive状態であるか確認する
'
'パラメータ:
'    [site]     In    確認Site番号
'
'戻り値:
'   確認結果(Active状態=True)
'
'注意事項:
'
    IsActiveSite = TheExec.sites.site(site).Selected

End Function

'#Pass
Public Function GetChanType(ByVal PinList As String) As chtype
'内容:
'   指定ピンのChannelTypeを取得する
'
'パラメータ:
'    [PinList]    In    確認対象ピン
'
'戻り値:
'   確認対象ピンのChannelType
'
'注意事項:
'   確認対象ピンリストに異なるChannelTypeの
'   Pinが指定された場合は、戻り値がchUnkとなる
'
    Dim chanType As chtype
    Dim pinArr() As String
    Dim pinNum As Long
    Dim i As Long

    With TheExec.DataManager
        chanType = .chanType(PinList)

        If chanType = chUnk Then
            On Error GoTo INVALID_PINLIST
            Call .DecomposePinList(PinList, pinArr, pinNum)
            On Error GoTo -1

            chanType = .chanType(pinArr(0))
            For i = 0 To pinNum - 1
                If chanType <> .chanType(pinArr(i)) Then
                    chanType = chUnk
                    Exit For
                End If
            Next i
        End If
    End With

    GetChanType = chanType
    Exit Function

INVALID_PINLIST:
    GetChanType = chUnk

End Function

'#Pass
Public Sub SeparatePinList(ByVal PinList As String, ByRef retPinNames() As String)
'内容:
'   ピンリストのピン情報を配列形式に分離
'
'パラメータ:
'    [PinList]          In    分離対象ピン
'    [retPinNames()]    Out   分離後ピン
'
'戻り値:
'
'注意事項:
'
    Dim pinNum As Long
    Call TheExec.DataManager.DecomposePinList(PinList, retPinNames, pinNum)
    
End Sub

'#Pass
Public Function CreateEmpty2DArray(ByVal Dim1 As Long, ByVal Dim2 As Long) As Variant
'内容:
'   指定サイズの二次元配列変数を作成
'
'パラメータ:
'    [Dim1]    In   配列次元1の数
'    [Dim2]    In   配列次元2の数
'
'戻り値:
'   Dim1×Dim2のサイズの2次元配列
'
'注意事項:
'   配列の値は0
'
    Dim ret2DArr() As Variant
    Dim tmp() As Double
    Dim i As Long
    
    ReDim ret2DArr(Dim1)
    ReDim tmp(Dim2)
    
    For i = 0 To UBound(ret2DArr)
        ret2DArr(i) = tmp
    Next i
    
    CreateEmpty2DArray = ret2DArr
    
End Function

'#Pass
Public Function IsValidSite(ByVal site As Long) As Boolean
'内容:
'   サイトが有効なサイトか確認
'
'パラメータ:
'    [site]    In   確認するサイト番号
'
'戻り値:
'   確認結果(有効Site = True)
'
'注意事項:
'
    If site = ALL_SITE Then
        IsValidSite = True
    ElseIf 0 <= site And site <= CountExistSite Then
        IsValidSite = True
    Else
        IsValidSite = False
    End If

End Function

'#Pass
Public Function CreateLimit(ByVal dstVal As Variant, ByVal loLim As Double, ByVal hiLim As Double) As Variant
'内容:
'   Limit値を上限値と下限値より生成
'
'パラメータ:
'    [dstVal]    In   設定値
'    [loLim]     In   下限値
'    [hiLim]     In   上限値
'
'戻り値:
'   Limit値
'
'注意事項:
'
    Dim i As Long

    If IsArray(dstVal) Then
        For i = 0 To UBound(dstVal)
            If dstVal(i) < loLim Then dstVal(i) = loLim
            If dstVal(i) > hiLim Then dstVal(i) = hiLim
        Next i
    Else
        If dstVal < loLim Then dstVal = loLim
        If dstVal > hiLim Then dstVal = hiLim
    End If

    CreateLimit = dstVal

End Function

'#Pass
Public Function ReadMultiResult(ByVal PinName As String, ByRef retResult() As Double, ByRef Results As Collection) As Boolean
'内容:
'    ピン名をキーにコレクションの要素を取り出す
'
'パラメータ:
'    [PinName]        In   コレクションのキーとなるピン名
'    [retResult()]    Out  コレクションから取り出した値
'    [results]        In   要素を取り出すコレクション
'
'戻り値:
'   ステータス（正常終了=True）
'
'注意事項:
'
    Dim site As Long
    Dim result As Variant

    On Error GoTo NOT_FOUND
    result = Results(PinName)
    On Error GoTo 0
    
    For site = 0 To CountExistSite
        retResult(site) = result(site)
    Next site

    ReadMultiResult = True
    Exit Function
    
NOT_FOUND:
    ReadMultiResult = False
    
End Function

'#Pass
Public Function IsGangPinlist(ByVal PinList As String, ByVal chtype As chtype) As Boolean
'内容:
'    '指定されたピンリストにギャングピンが含まれているか確認
'
'パラメータ:
'    [PinName]       In   確認を行うPinList
'    [chtype]        In   対象となるボードのChannelType
'
'戻り値:
'   確認結果（ギャングピンが含まれている=True）
'
'注意事項:
'
    Dim pinNames() As String
    Dim Channels() As Long
    
    Call GetChanList(PinList, ALL_SITE, chtype, Channels)
    Call SeparatePinList(PinList, pinNames)
    
    If (UBound(pinNames) + 1) * (CountExistSite + 1) <> UBound(Channels) + 1 Then
        IsGangPinlist = True  'ギャングピンがある
    Else
        IsGangPinlist = False 'ギャングピンはない
    End If

End Function

Public Function IsGangMultiPinlist(ByVal PinList As String) As Boolean
'内容:
'   指定されたピンリストのPinGpがすべてGANG接続用のPinGpか確認する
'
'パラメータ:
'   [PinList]   In  確認を行うPinList
'
'戻り値:
'   確認結果（すべてGANG接続用のPinGpである=True）
'
'注意事項:
'
    Dim pinListArr() As String
    'ピングループを展開せずにカンマ区切り形式の配列に変換
    Call ConvertStrPinListToArrayPinList(PinList, pinListArr)

    Dim tmpPinGp As Variant
    For Each tmpPinGp In pinListArr
        If IsGangPinlist(tmpPinGp, GetChanType(tmpPinGp)) <> True Then
            IsGangMultiPinlist = False
'            Call MsgBox(tmpPinGp & " はGANG接続用のPinGpではありません")
            Exit Function
        End If
    Next tmpPinGp

    IsGangMultiPinlist = True

End Function

Public Sub ConvertStrPinListToArrayPinList(ByVal StrPinList As String, ByRef ArrayPinList() As String)
'内容:
'   カンマ区切り文字列形式のピンリストを、配列形式のピンリストに変換
'
'パラメータ:
'   [StrPinList]    In   変換対象の文字列ピンリスト
'   [ArrayPinList]  Out  変換後の配列形式ピンリスト
'
'戻り値:
'
'注意事項:
'   PinListにPinGpを指定したときには、PinGpのメンバーは展開しません
'   形式変換のみとなります。
'
    Dim ret As Long
    Dim i As Long
    
    Erase ArrayPinList()

    Do
        ret = InStr(1, StrPinList, ",")
        If ret = 0 Then
            ReDim Preserve ArrayPinList(i)
            ArrayPinList(i) = StrPinList
            Exit Do
        End If
        ReDim Preserve ArrayPinList(i)
        ArrayPinList(i) = Left(StrPinList, ret - 1)
        StrPinList = Right(StrPinList, Len(StrPinList) - ret)
        i = i + 1
    Loop

End Sub
