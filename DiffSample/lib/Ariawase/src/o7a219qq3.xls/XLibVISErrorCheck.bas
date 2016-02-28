Attribute VB_Name = "XLibVISErrorCheck"
'概要:
'   VISクラスで使用するチェック用の関数群
'
'目的:
'   各VISクラス共通に使用するチェック用関数をまとめる
'
'作成者:
'   SLSI今手
'
'   XlibSTD_CommonDCMod_V01内の、DC関連エラーチェック
'   処理を切り出しまとめたもの。
'
'
'Code Checked
'Comment Checked
'

Option Explicit

Private Const ALL_SITE = -1

'#Pass
Public Function CheckPinList(ByVal PinList As String, ByVal chanType As chtype, ByVal FunctionName As String) As Boolean
'内容:
'    指定ピンが対象ChannelTypeとして定義されているかを確認
'
'パラメータ:
'    [PinList]        In  確認対象ピンリスト。
'    [chanType]       In  確認チャンネルタイプ
'    [FunctionName]   In  呼び出し元関数名
'
'戻り値:
'   確認結果
'
'注意事項:
'
    If GetChanType(PinList) <> chanType Then
        Call OutputErrMsg(PinList & " is Invalid Channel Type at " & FunctionName & "().")
        CheckPinList = False
    Else
        CheckPinList = True
    End If
    
End Function

'#Pass
Public Function CheckSinglePins(ByVal PinList As String, ByVal chanType As chtype, ByVal FunctionName As String) As Boolean
'内容:
'    指定ピンが同一ChannelTypeの物か確認
'
'パラメータ:
'    [PinList]        In  確認対象ピンリスト。
'    [chanType]       In  確認チャンネルタイプ
'    [FunctionName]   In  呼び出し元関数名
'
'戻り値:
'   確認結果
'
'注意事項:
'
    Dim Channels() As Long
    Dim ChanNum As Long
    Dim siteNum As Long
    Dim errMsg As String
    Call TheExec.DataManager.GetChanList(PinList, ALL_SITE, chanType, Channels, ChanNum, siteNum, errMsg)
    
    If ChanNum <> siteNum Then
        Call OutputErrMsg("Don't Support Multi Pins at " & FunctionName & "().")
        CheckSinglePins = False
    Else
        CheckSinglePins = True
    End If
    
End Function

'#Pass
Public Function CheckForceVariantValue(ByVal ForceVal As Variant, ByVal loLim As Double, ByVal hiLim As Double, ByVal FunctionName As String) As Boolean
'内容:
'    Force値が下限値と上限値の間の値であるかを確認
'
'パラメータ:
'    [ForceVal]       In  Force値
'    [loLim]          In  下限値
'    [hiLim]          In  上限値
'    [FunctionName]   In  呼び出し元関数名
'
'戻り値:
'   確認結果
'
'注意事項:
'
    Dim site As Long
    
    If IsArray(ForceVal) Then
        If UBound(ForceVal) <> CountExistSite Then
            Call OutputErrMsg("ForceVal is Invalid Site Array at " & FunctionName & "().")
            CheckForceVariantValue = False
            Exit Function
        End If
        
        For site = 0 To CountExistSite
            If (ForceVal(site) < loLim Or hiLim < ForceVal(site)) Then
                Call OutputErrMsg("ForceVal(= " & ForceVal(site) & ") must be between " & loLim & " and " & hiLim & " at " & FunctionName & "().")
                CheckForceVariantValue = False
                Exit Function
            End If
        Next site
        
    Else
        If (ForceVal < loLim Or hiLim < ForceVal) Then
            Call OutputErrMsg("ForceVal(= " & ForceVal & ") must be between " & loLim & " and " & hiLim & " at " & FunctionName & "().")
            CheckForceVariantValue = False
            Exit Function
        End If
    End If
    
    CheckForceVariantValue = True
    
End Function

'#Pass
Public Function CheckClampValue(ByVal clampVal As Double, ByVal loLim As Double, ByVal hiLim As Double, ByVal FunctionName As String) As Boolean
'内容:
'    Clamp値が、下限値と上限値の間の値であるかを確認
'
'パラメータ:
'    [clampVal]       In  Clamp値
'    [loLim]          In  下限値
'    [hiLim]          In  上限値
'    [FunctionName]   In  エラーメッセージに表示する呼び出し元関数名
'
'戻り値:
'   確認結果
'
'注意事項:
'
    If (clampVal < loLim Or hiLim < clampVal) Then
        Call OutputErrMsg("ClampVal(= " & clampVal & ") must be between " & loLim & " and " & hiLim & " at " & FunctionName & "().")
        CheckClampValue = False
    Else
        CheckClampValue = True
    End If

End Function

'#Pass
Public Function IsExistSite(ByVal site As Long, ByVal FunctionName As String) As Boolean
'内容:
'    指定番号のSiteが存在するか確認
'
'パラメータ:
'    [site]           In  確認Site番号
'    [FunctionName]   In  エラーメッセージに表示する呼び出し元関数名
'
'戻り値:
'   確認結果
'
'注意事項:
'
    If site <> ALL_SITE And (site < 0 Or CountExistSite < site) Then
        Call OutputErrMsg("Site(= " & site & ") must be -1 or between 0 and " & CountExistSite & " at " & FunctionName & "().")
        IsExistSite = False
    Else
        IsExistSite = True
    End If

End Function

'#Pass
Public Function CheckResultArray(ByRef retResult() As Double, ByVal FunctionName As String) As Boolean
'内容:
'    結果格納用の配列変数の要素数が存在するSite数と合っているか確認
'
'パラメータ:
'    [retResult()]      In  結果格納用の配列変数
'    [FunctionName]   In  エラーメッセージに表示する呼び出し元関数名
'
'戻り値:
'   確認結果
'
'注意事項:
'
    If UBound(retResult) <> CountExistSite Then
        Call OutputErrMsg("Elements of retResult() is Different from Number of Site at " & FunctionName & "().")
        CheckResultArray = False
    Else
        CheckResultArray = True
    End If

End Function

'#Pass
Public Function CheckAvgNum(ByVal avgNum As Long, ByVal FunctionName As String) As Boolean
'内容:
'    メータリード時のアベレージ回数値が1未満でないことを確認
'
'パラメータ:
'    [avgNum]         In  アベレージ回数値
'    [FunctionName]   In  エラーメッセージに表示する呼び出し元関数名
'
'戻り値:
'   確認結果
'
'注意事項:
'

    If avgNum < 1 Then
        Call OutputErrMsg("AvgNum must be 1 or More at " & FunctionName & "().")
        CheckAvgNum = False
    Else
        CheckAvgNum = True
    End If
    
End Function

Public Function CheckFailSiteExists(ByVal FunctionName As String) As Boolean
'内容:
'    存在するサイトにFAILサイトがあるかどうかを確認する
'
'パラメータ:
'    [FunctionName]   In  エラーメッセージに表示する呼び出し元関数名
'
'戻り値:
'   確認結果（True：FAILサイトが存在する）
'
'注意事項:
'

    With TheExec.sites
        If .ExistingCount <> .ActiveCount Then
            CheckFailSiteExists = True
            Call MsgBox("ExistingSites=" & .ExistingCount & _
            ", ActiveSites=" & .ActiveCount & _
            " FAIL site exists.  at " & FunctionName & "()", vbCritical, "CheckFailSiteExists")
            Exit Function
        End If
    End With

    CheckFailSiteExists = False
    
End Function

