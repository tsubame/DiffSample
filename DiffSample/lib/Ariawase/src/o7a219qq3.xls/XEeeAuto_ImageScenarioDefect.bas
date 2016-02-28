Attribute VB_Name = "XEeeAuto_ImageScenarioDefect"
'概要:
'   撮像シナリオから呼ばれるマクロ群
'
'目的:
'
'
'作成者:
'   2012/01/23 Ver0.1 D.Maruyama
'   2012/01/31 Ver0.2 H.Yamanaka : OffsetZone情報を追加


Option Explicit

'Defectの情報格納箇所
'VarBankがオブジェクト（Collection）をうけとってくれないのと
'この情報の持ち方を変更すると、いっぱいほかの箇所を直さないといけないので
'145の情報の持ち方にのっかる。
Public Defect_Infomation(nSite) As type_point
Public Type type_point
    Label() As String                       'Label
    X_address() As Long                     'X座標
    Y_address() As Long                     'Y座標
    Value() As Double                       '値（MAX，MIN, AVE）
    Enable() As Integer                     'BlowEnable
End Type

Public m_fileLunDefectFile As Integer

'内容:
'   Defect構造体配列の初期化を行う
'
'
'注意事項:
'   dc_setup内で必ず読んでほしい
'

Public Sub InitializeDefectInformation()

    Dim siteIdx As Long
    
    For siteIdx = 0 To nSite
        With Defect_Infomation(siteIdx)
            Erase .Label
            Erase .X_address
            Erase .Y_address
            Erase .Value
            Erase .Enable
        End With
    Next siteIdx

End Sub

Public Sub UninitializeDefectInformation()

    Dim siteIdx As Long
    
    For siteIdx = 0 To nSite
        With Defect_Infomation(siteIdx)
            Erase .Label
            Erase .X_address
            Erase .Y_address
            Erase .Value
            Erase .Enable
        End With
    Next siteIdx

End Sub
'こちらに移動
Public Function mf_OpenDefectFile() As Boolean


    If Sw_Ana = 1 Then
        m_fileLunDefectFile = FreeFile
        Open Defect_full_fname For Append As m_fileLunDefectFile
    End If
    
    mf_OpenDefectFile = True
    
End Function

'こちらに移動
Public Function mf_CloseDefectFile() As Boolean


    If Sw_Ana = 1 Then Close m_fileLunDefectFile
    
    mf_CloseDefectFile = True

End Function

Public Sub WriteDefect_New(ByRef srcPlane As CImgPlane, _
    ByVal srcZone As String, _
    ByRef srcValue() As Double, _
    ByRef Lsb() As Double, _
    ByVal FlgName As String, _
    ByVal LabelName As String, _
    ByVal Unit As String, _
    ByVal Correction As Double, _
    ByVal DefectMaxNum As Long, _
    ByVal ArrayLabel As String, _
    ByVal offsetZone As String, _
    Optional ByVal IsOtp As Boolean = False)
    
    

    'ZoneのOffset値を算出する。
    Dim pOffsetX As Long
    Dim pOffsetY As Long
    With TheIDP.PMD(offsetZone)
        pOffsetX = .Left - 1
        pOffsetY = .Top - 1
    End With
    
    Dim site As Long
    For site = 0 To nSite - 1
        If TheExec.sites.site(site).Active Then
            If (Flg_Debug = 1 Or Sw_Ana = 1) And (srcValue(site) > 0) Then
                TheExec.Datalog.WriteComment "******** " & LabelName & " DEFECT ADDRESS & DATA (SITE:" & site & ") *********"
            End If
            'OTPのための欠陥情報保存のため、d_read_vmcuはフラグによらず実行する
            '関数内でフラグに応じた出力をするので、必ず呼んでも問題ない
            '保存する時間ももったいないと言うのであれば考えよう
            Call d_read_vmcu(site, srcPlane, srcZone, FlgName, srcValue(site), DefectMaxNum, Lsb, LabelName, Unit, ArrayLabel, pOffsetX, pOffsetY, Correction, IsOtp)
        End If
    Next site
    
End Sub

''内容:
''   DefectFileへの書き出しを行う
''
''
''注意事項:
''
''
'Private Function WriteDefect(ByRef pParams As CUsrMacroParams) As Long
''/* Defect吐き出しマクロ */
'
'    On Error GoTo ErrHandler
'    With pParams
'        '/* インプットパラメータの取得 */
'        Dim srcPlane As CParamPlane
'        Set srcPlane = .GetInParam("SrcPlane")
'        Dim srcValue As CParamImgResult
'        Set srcValue = .GetInParam("SrcValue")
'
'         '/* コンディションパラメータの取得 */
'        Dim pSrcZone As String
'        pSrcZone = .GetProperty("srcZone")
'        Dim pFlagName As String
'        pFlagName = .GetProperty("FlagName")
'        Dim pLabelName As String
'        pLabelName = .GetProperty("LabelName")
'        Dim pUnit As String
'        pUnit = .GetProperty("Unit")
'        Dim pCorrection As Double
'        pCorrection = .GetProperty("Correction")
'        Dim pDefectMaxNum As Long
'        pDefectMaxNum = .GetProperty("DefectMaxNum")
'        Dim pArrayLabel As String
'        pArrayLabel = .GetProperty("ArrayLabel")
'        Dim poffsetZone As String
'        poffsetZone = .GetProperty("offsetZone")
'    End With
'
'    'LSBは使わないかもしれないけども取得しておく
'    Dim pLSB() As Double
'    pLSB = srcPlane.DeviceConfigInfo.LSB.AsDouble
'
'    'ZoneのOffset値を算出する。
'    Dim pOffsetX As Long
'    Dim pOffsetY As Long
'    With TheIDP.PMD(poffsetZone)
'        pOffsetX = .Left - 1
'        pOffsetY = .Top - 1
'    End With
'
'    Dim siteIdx As Long
'    For siteIdx = 0 To srcValue.CountSite - 1
'        If srcValue.Flat.Site(siteIdx).STATUS Then
'            If (Flg_Debug = 1 Or Sw_Ana = 1) And (srcValue.Flat.Site(siteIdx).Value > 0) Then
'                TheExec.Datalog.WriteComment "******** " & pLabelName & " DEFECT ADDRESS & DATA (SITE:" & siteIdx & ") *********"
'            End If
'            'OTPのための欠陥情報保存のため、d_read_vmcuはフラグによらず実行する
'            '関数内でフラグに応じた出力をするので、必ず呼んでも問題ない
'            '保存する時間ももったいないと言うのであれば考えよう
'            Call d_read_vmcu(siteIdx, srcPlane.Plane, pSrcZone, pFlagName, srcValue.Flat.Site(siteIdx).Value, pDefectMaxNum, pLSB, pLabelName, pUnit, pArrayLabel, pOffsetX, pOffsetY, pCorrection, True)
'        End If
'    Next siteIdx
'
'    '/* マクロ実行成功 */
'    Exit Function
'
'    '/* エラーハンドリング */
'ErrHandler:
'    Set pParams.err = err
'End Function

'/* シナリオマクロ用欠陥出力関数 */
'/* オリジナルからの変更点はグローバルフラグのみ */
Private Sub d_read_vmcu(ByVal site As Long, ByVal DefPmd As CImgPlane, ByVal DefZone As String, _
                                ByVal FlgName As String, ByVal Num As Long, ByVal MaxNum, _
                                ByRef Lsb() As Double, ByVal signature As String, ByVal Unit As String, _
                                ByVal pArrayLabel As String, _
                                ByVal pOffsetX As String, ByVal pOffsetY As String, _
                                Optional ByVal BaseVal As Double = 1, _
                                Optional ByVal IsDefectSave As Boolean = False _
                                )
    
    'DefectDataも保存せず、ログ出力もせず、Defectもはかないなら、何もしないでぬける
    If (Not IsDefectSave) And Flg_Debug <> 1 And Sw_Ana <> 1 Then
        Exit Sub
    End If

    Dim PixelLogResult() As T_PIXINFO
    
    Dim fileNum As Integer
    Dim i As Long
    Dim x As Long, y As Long
    Dim Data As Double
    
    If Num <= 0 Then Exit Sub
    If BaseVal = 0 Then BaseVal = 1
        
    If Num > MaxNum Then Num = MaxNum
    
    ReDim PixelLogResult(Num)
    
    Dim lBeforeSize As Long
    
    
    With DefPmd
        Call .SetPMD(DefZone)
        Call .PixelLog(site, FlgName, PixelLogResult, Num, idpAddrAbsolute)
    End With
    
'    If Sw_Ana = 1 Then
'        FileNum = FreeFile
'        Open Defect_full_fname For Append As FileNum
'    End If
    
    If IsDefectSave Then
        Dim lCurStart As Long
        lCurStart = AllocateDefectInformation(site, Num)
    End If
    
    For i = 0 To Num - 1
        x = PixelLogResult(i).x - pOffsetX
        y = PixelLogResult(i).y - pOffsetY
        Select Case Unit
            Case "mV"
                Data = PixelLogResult(i).Value * Lsb(site) / mV
            Case "%"
                Data = PixelLogResult(i).Value / BaseVal * 100
            Case Else
                Data = PixelLogResult(i).Value * Lsb(site)
        End Select
        
        If Sw_Ana = 1 Then
            Print #fileNum, _
                CStr(WaferNo) & Format(CStr(DeviceNumber_site(site)), "0000") & " " & signature & " " & Unit & " " _
                & Format(x, "#### "); Format(y, "#### "); Format(Data * 1000, "######") & ""
        End If
        
        If Flg_Debug = 1 Then
            Call TheExec.Datalog.WriteComment( _
                signature & ":(" & CStr(x) & ", " & CStr(y) & ") = " & Format(Data, "0.##0") & " " & Unit)
        End If
                                
    Next i
    
    
    If IsDefectSave Then
        For i = 0 To Num - 1
            With Defect_Infomation(site)
                .Label(lCurStart + i) = pArrayLabel
                .X_address(lCurStart + i) = x
                .Y_address(lCurStart + i) = y
                .Value(lCurStart + i) = PixelLogResult(i).Value
    '            .Enable (lNum + lCur)
            End With
        Next i
    End If
    
'    If Sw_Ana = 1 Then
'        Close FileNum
'    End If
    
End Sub

'内容:
'   Defect構造体配列のアロケートを行う
'
'
'注意事項:
'
'
Private Function AllocateDefectInformation(ByVal lSite As Long, ByVal lNum As Long) As Long

    Dim lCur As Long
    On Error Resume Next
    lCur = UBound(Defect_Infomation(lSite).Label)
    
    'UBoundの成功でErrは0(最低一度はRedimされている)
    If Err = 0 Then
        AllocateDefectInformation = lCur + 1
        With Defect_Infomation(lSite)
            ReDim Preserve .Label(lNum + lCur)
            ReDim Preserve .X_address(lNum + lCur)
            ReDim Preserve .Y_address(lNum + lCur)
            ReDim Preserve .Value(lNum + lCur)
            ReDim Preserve .Enable(lNum + lCur)
        End With
    'UBoundの成功でErrは0以外(要素は空)
    Else
        AllocateDefectInformation = 0
        Err.Clear
        With Defect_Infomation(lSite)
            ReDim .Label(lNum - 1)
            ReDim .X_address(lNum - 1)
            ReDim .Y_address(lNum - 1)
            ReDim .Value(lNum - 1)
            ReDim .Enable(lNum - 1)
        End With
    End If
    
    '下のコマンドでErrがクリアされるので、もとに戻すのはここ
    On Error GoTo 0
    
End Function




