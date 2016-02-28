Attribute VB_Name = "Otp_PageWithSram_SubMod"
Option Explicit

Public blnFlg_BlowCheck As Boolean

Public Sub OtpVariableClear()
'OTP Initialize Sub.
'OTP測定で使用している変数のクリア

Erase AddrParPage
Erase BitParPage
Erase BlowDataAllBin
Erase BlowDataAllBin2
Erase ReadDataAllBin
Erase FFBlowInfo



Dim NowPage As Integer
For NowPage = 0 To OtpPageEnd
    Flg_ModifyPage(NowPage) = False
Next NowPage

End Sub

Public Sub Set_OtpCondition(ByRef ArgArr() As String)

    'パラメータの取得
    '想定数より小さければエラーコード

    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_VARIABLE_PARAM) Then
        Err.Raise 9999, "OTP", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
    End If

    '終了文字列が見つからないのもだめ
    Dim IsFound As Boolean
    Dim lCount As Long
    Dim i As Long

    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "OTP", """#EOP"" is not found! [" & GetInstanceName & "] !"
    End If

    Dim testConditionList() As String
    testConditionList = Split(ArgArr(0), ",")
    For i = 0 To UBound(testConditionList)
        If Trim(testConditionList(i) <> "") Then
            Call TheCondition.SetCondition(testConditionList(i))
        End If
    Next i

End Sub

Public Sub OtpInitialize_Get_AddressParPage()
'OTP Initialize Sub.
'OTPMAPシートから各PageのAddress数とBit数をGETするよ。

'注意事項（この関数の制約事項。下記の制約に違反している場合には動作保証しない。）
'制約1：Page番号は、最終Pageの最終Addressまで1セルずつ番号を記入しておく。前セル同じPageだからといって、記入をしなかったり、セルを統合してはいけない。


    Dim NowPage As Integer
    Dim RowCount As Long

    Worksheets("OTPMAP").Select                                                                         'OTP情報が記載されているSheetを選択(OTPMAP)

    For NowPage = 0 To OtpPageSize - 1                                                                  'Pageまわし

        Do While Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value <> ""          'OTPMAPのPage列先頭から最後尾までのLOOP

            If NowPage = Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value Then   'Page毎のAddress数をGET
                AddrParPage(NowPage) = AddrParPage(NowPage) + 1
            End If

            RowCount = RowCount + BitParHex
        Loop

        BitParPage(NowPage) = BitParHex * AddrParPage(NowPage)                                          'Page毎のBit数をGET
        RowCount = 0

    Next NowPage

End Sub

Public Sub OtpInitialize_Get_FixedValue()
'OTP Initialize Sub.
'OTPMAPシートから固定値情報をGET。BlowとVerifyで使用するよ。

'注意事項（この関数の制約事項。下記の制約に違反している場合には動作保証しない。）
'制約1：Value(Bin)欄は空欄NG。記載できる文字はこれだけ⇒"0" or "1" or "X"


    Dim site As Long
    Dim NowPage As Integer
    Dim NowBit As Integer
    Dim RowCount As Long
    Dim TotalRowCount As Long
    
    Worksheets("OTPMAP").Select                                                                         'OTP情報が記載されているSheetを選択(OTPMAP)
    
    '===== 固定値情報をBin変数に入れ込み ======================================
    For site = 0 To nSite
'        If TheExec.Sites.site(site).Active = True Then
            
            For NowPage = 0 To OtpPageSize - 1                                                          'Pageまわし
                For NowBit = 0 To (BitParPage(NowPage)) - 1                                             'Bitまわし
                    
                    ReadDataAllBin(site, NowPage, NowBit) = Cells(NowBit + TotalRowCount + OtpInfoSheet_Row_Value, OtpInfoSheet_Column_Value).Value     'Verify用変数に固定値情報を格納
                    BlowDataAllBin(site, NowPage, NowBit) = ReadDataAllBin(site, NowPage, NowBit)                                                       'Blow用変数に固定値情報を格納
                    BlowDataAllBin2(site, NowPage, NowBit) = ReadDataAllBin(site, NowPage, NowBit)                                                       'Blow用変数に固定値情報を格納
                    If Cells(OtpInfoSheet_Row_BlowInfo + RowCount, OtpInfoSheet_Column_BlowInfo).Value = "SRAM" Then        '現在のBitがSRAM用であれば、そのBitのBlow用変数は0の固定値とする(冗長後に上書きをしないように)
                        BlowDataAllBin(site, NowPage, NowBit) = "0"
                        BlowDataAllBin2(site, NowPage, NowBit) = "X"
                    End If
                    
                    RowCount = RowCount + 1
                    
                Next NowBit
                TotalRowCount = RowCount
            Next NowPage
            
            RowCount = 0
            TotalRowCount = 0
            
'        End If
    Next site
    
End Sub

Public Sub OtpInitialize_Get_FFBlowInfo()
'OTP Initialize Sub.
'OTPMAPシートからFFBlowするPageとBit情報をGETするよ。

'注意事項（この関数の制約事項。下記の制約に違反している場合には動作保証しない。）
'制約1：FF書き込みの列で空欄はNG。FF書き込みする⇒"1"。FF書き込みしない⇒"0"。1か0のどちらかを必ず記載してあること。
'制約2：FF書き込みが複数ページにまたがるのはNG。ある1ページだけにしてね。ある1ページ内であれば、何BitでもBlowできます(FF(11111111)じゃなくてもイケる(EEとかもOK))。


    Dim site As Long
    Dim NowPage As Integer
    Dim NowBit As Integer
    Dim RowCount As Long
    
    Worksheets("OTPMAP").Select                                                             'OTP情報が記載されているSheetを選択(OTPMAP)
    
    '===== 固定値情報をBin変数に入れ込み ======================================
    Do While Cells(OtpInfoSheet_Row_FF + RowCount, OtpInfoSheet_Column_FF).Value <> ""      'OTPMAPのFF情報が記載されている列の先頭から最後尾までをLOOP

        NowPage = Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value   'LOOP中の現在のPageとそのBit番号をGET
        NowBit = Cells(OtpInfoSheet_Row_Bit + RowCount, OtpInfoSheet_Column_Bit).Value
        
        If Cells(OtpInfoSheet_Row_FF + RowCount, OtpInfoSheet_Column_FF).Value = 1 Then     'FFBlow用の変数作成。あと、Page情報もGET。
            FFBlowInfo(NowPage, NowBit) = 1
            FFBlowPage = NowPage
        Else
            FFBlowInfo(NowPage, NowBit) = 0
        End If
        
        RowCount = RowCount + 1
        
    Loop
                
End Sub

Public Sub OtpInitialize_Select_OtpBlow_Page()
'OTP Initialize Sub.
'OTPMAPシートからOTPBLOWが必要となるページ情報を取得。そのページでのBlowValueが全て"0"であれば、そのページはBlowのPatRunを行わない。

'注意事項（この関数の制約事項。下記の制約に違反している場合には動作保証しない。）
'制約1：Value(Bin)欄は空欄NG。記載できる文字はこれだけ⇒"0" or "1" or "X"

    Dim NowPage As Integer
    Dim RowCount As Long
    
    Worksheets("OTPMAP").Select                                                                     'OTP情報が記載されているSheetを選択(OTPMAP)
    
    For NowPage = 0 To OtpPageSize - 1
        Flg_OtpBlowPage(NowPage) = False                                                            'まずは全ページFalseにしておく。全ページBlow実行無し。
    Next NowPage
        
    Do While Cells(OtpInfoSheet_Row_Value + RowCount, OtpInfoSheet_Column_Value).Value <> ""        '各Pageの各Bit分LOOP
        If Cells(OtpInfoSheet_Row_Value + RowCount, OtpInfoSheet_Column_Value).Value <> "0" Then
            NowPage = Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value       '現在のPageで"0"以外のBlowValueが存在すれば、TrueにしてBlow実行フラグを立てる
            Flg_OtpBlowPage(NowPage) = True
            
            If Cells(OtpInfoSheet_Row_Value + RowCount, OtpInfoSheet_Column_Value).Value = 1 Then   '現在のPageで"1"のBlowValue(=固定値)が存在すればそのPage情報を保持。対象Pageが複数あれば、その中の最後のPage情報を保持。
                Flg_OtpBlowFixValPage = Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value
            End If
            
        End If
        RowCount = RowCount + 1
    Loop
        
End Sub

Private Function Dec2Bin(myDecvalue As String, OutBit As Integer) As String
'OTP Standard Function.
'10進数を2進数に変換する

    Dim lngdecnumber As Long
    Dim strbinnumber As String
    strbinnumber = ""
    lngdecnumber = 0

    lngdecnumber = CLng(myDecvalue)

    Do
        strbinnumber = CStr(lngdecnumber Mod 2) & strbinnumber
        lngdecnumber = Fix(lngdecnumber / 2)
    Loop While lngdecnumber > 0

    Do While Len(strbinnumber) < OutBit
        strbinnumber = "0" & strbinnumber
    Loop

    Dec2Bin = strbinnumber

End Function

Public Sub OtpInitialize_Make_FixedValuePattern()
'OTP Initialize Sub.
'固定値情報は最初にModifyしておく。


    Dim site As Long
    Dim NowPage As Integer
    Dim NowBit As Integer

    Dim PageBit As Long
    Dim BinData As String
    Const PageInfoBit = 8
    Const ByteBit = 8
    Dim PG_ARY(PageInfoBit - 1) As String
    Dim ByteCount As Long
    Dim LenLen As Long
    Dim iiii As Long
    Dim ModifyDataBlowDeb As String
    Dim ModifyDataVerifyDeb As String
    Dim ModifyDataFFDeb As String
    Dim i As Integer

    If Flg_Simulator = 1 Then Exit Sub
    
    'Active Site PickUp (1site de OK)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Exit For
        End If
    Next site
            
    Call StopPattern
     
    For NowPage = 0 To OtpPageEnd
    
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        '■Page情報をModify-> Blowパターン、Verifyパターン、固定値パターン、Blankパターン、FFBlowパターン
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    
        BinData = Dec2Bin(CStr(NowPage), PageInfoBit)
        
        For PageBit = 0 To PageInfoBit - 1
            PG_ARY(PageBit) = Mid(BinData, PageBit + 1, 1)
        Next PageBit
    
        TheHdw.Digital.Patterns.pat("OtpBlowPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_Page_OtpBlow & CStr(NowPage), 0, RejiIn, PG_ARY
        TheHdw.Digital.Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_Page_OtpVerify & CStr(NowPage), 0, RejiOut, PG_ARY
        TheHdw.Digital.Patterns.pat("OtpFixedValueCheckPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_Page_OtpFixedValueCheck & CStr(NowPage), 0, RejiOut, PG_ARY
        TheHdw.Digital.Patterns.pat("BlankCheckPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_Page_BlankCheck & CStr(NowPage), 0, RejiOut, PG_ARY
        If FFBlowPage = NowPage Then    'FFBlowするページだけModify(FFBlowするページは1Pageであること前提)
            TheHdw.Digital.Patterns.pat("OtpBlow_Break_Pat").ModifyPinVectorBlockData Label_Page_OtpBlow_Break, 0, RejiIn, PG_ARY
        End If
    
    
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        '■固定値Blow情報をBlowパターンへModify
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        ByteCount = 0
        ModifyDataBlowDeb = ""
        '===== Make Blow Modify Data ==================================
        For NowBit = 0 To BitParPage(NowPage) - 1
            ByteCount = ByteCount + 1
            If ByteCount < ByteBit Then
                ModifyDataBlowDeb = ModifyDataBlowDeb & BlowDataAllBin(site, NowPage, NowBit)
            ElseIf ByteCount = ByteBit Then
                ModifyDataBlowDeb = ModifyDataBlowDeb & BlowDataAllBin(site, NowPage, NowBit) & "X"
            End If
            If ByteCount = 8 Then ByteCount = 0
        Next NowBit
        
        LenLen = Len(ModifyDataBlowDeb)
        ReDim ModifyDataBlow(LenLen - 1) As String
        For iiii = 1 To LenLen
            ModifyDataBlow(iiii - 1) = Mid(ModifyDataBlowDeb, iiii, 1)
        Next iiii
        
        '===== Modify Blow Pattern (Fixed Value) ======================
        TheHdw.Digital.Patterns.pat("OtpBlowPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_OtpBlow & CStr(NowPage), 0, RejiIn, ModifyDataBlow
    
    
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        '■AutoBlow情報をBlowパターンへModify
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

        Dim ModifyDataBlowAuto(ByteBit - 1) As String
        For iiii = 1 To ByteBit
            ModifyDataBlowAuto(iiii - 1) = "0"
        Next iiii
        
        '===== Modify Blow Pattern (Fixed Value) ======================
        TheHdw.Digital.Patterns.pat("OtpBlowPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_OtpBlowAuto & CStr(NowPage), 0, RejiIn, ModifyDataBlowAuto

    
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        '■固定値Blow情報をVerifyパターンへModify
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        ByteCount = 0
        ModifyDataVerifyDeb = ""
        '===== Make Verify Modify Data ================================
        For NowBit = 0 To BitParPage(NowPage) - 1

            ByteCount = ByteCount + 1

            If ByteCount < ByteBit Then
                If BlowDataAllBin(site, NowPage, NowBit) = "0" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "L"
                ElseIf BlowDataAllBin(site, NowPage, NowBit) = "1" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "H"
                ElseIf BlowDataAllBin(site, NowPage, NowBit) = "X" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "X"
                End If
            ElseIf ByteCount = ByteBit Then
                If BlowDataAllBin(site, NowPage, NowBit) = "0" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "LX"
                ElseIf BlowDataAllBin(site, NowPage, NowBit) = "1" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "HX"
                ElseIf BlowDataAllBin(site, NowPage, NowBit) = "X" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "XX"
                End If

                ByteCount = 0
            End If

        Next NowBit

        LenLen = Len(ModifyDataVerifyDeb)
        ReDim ModifyDataVerifyFix(LenLen - 1) As String
        For iiii = 1 To LenLen
            ModifyDataVerifyFix(iiii - 1) = Mid(ModifyDataVerifyDeb, iiii, 1)
        Next iiii

        '===== Modify Verify Pattern (Fixed Value) ====================
        TheHdw.Digital.Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_OtpVerify & CStr(NowPage), 0, RejiOut, ModifyDataVerifyFix


        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        '■固定値Blow情報を固定値パターンへModify
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        ByteCount = 0
        ModifyDataVerifyDeb = ""
        '===== Make Verify Modify Data ================================
        For NowBit = 0 To BitParPage(NowPage) - 1

            ByteCount = ByteCount + 1

            If ByteCount < ByteBit Then
                If BlowDataAllBin2(site, NowPage, NowBit) = "0" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "L"
                ElseIf BlowDataAllBin2(site, NowPage, NowBit) = "1" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "H"
                ElseIf BlowDataAllBin2(site, NowPage, NowBit) = "X" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "X"
                End If
            ElseIf ByteCount = ByteBit Then
                If BlowDataAllBin2(site, NowPage, NowBit) = "0" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "LX"
                ElseIf BlowDataAllBin2(site, NowPage, NowBit) = "1" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "HX"
                ElseIf BlowDataAllBin2(site, NowPage, NowBit) = "X" Then
                    ModifyDataVerifyDeb = ModifyDataVerifyDeb & "XX"
                End If

                ByteCount = 0
            End If

        Next NowBit

        LenLen = Len(ModifyDataVerifyDeb)
        ReDim ModifyDataVerifyFix2(LenLen - 1) As String
        For iiii = 1 To LenLen
            ModifyDataVerifyFix2(iiii - 1) = Mid(ModifyDataVerifyDeb, iiii, 1)
        Next iiii

        '===== Modify Verify Pattern (Fixed Value) ====================
        TheHdw.Digital.Patterns.pat("OtpFixedValueCheckPage" & CStr(NowPage) & "_Pat").ModifyPinVectorBlockData Label_OtpFixedValueCheck & CStr(NowPage), 0, RejiOut, ModifyDataVerifyFix2


        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        '■"FF"Blow情報をFF専用BlowパターンへModify
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        ByteCount = 0
        ModifyDataFFDeb = ""
        If FFBlowPage = NowPage Then    'FFBlowするページだけModify(FFBlowするページは1Pageであること前提)

            '===== Make Blow Modify Data ==================================
            For NowBit = 0 To BitParPage(NowPage) - 1
                
                ByteCount = ByteCount + 1
                
                If ByteCount < ByteBit Then
                    ModifyDataFFDeb = ModifyDataFFDeb & FFBlowInfo(NowPage, NowBit)
                ElseIf ByteCount = ByteBit Then
                    ModifyDataFFDeb = ModifyDataFFDeb & FFBlowInfo(NowPage, NowBit) & "X"
                End If
                If ByteCount = 8 Then ByteCount = 0
            Next NowBit

            LenLen = Len(ModifyDataFFDeb)
            ReDim ModifyDataFF(LenLen - 1) As String
            For iiii = 1 To LenLen
                ModifyDataFF(iiii - 1) = Mid(ModifyDataFFDeb, iiii, 1)
            Next iiii

            '===== Modify Blow Pattern (Fixed Value) ======================
            TheHdw.Digital.Patterns.pat("OtpBlow_Break_Pat").ModifyPinVectorBlockData Label_OtpBlow_Break, 0, RejiIn, ModifyDataFF


            '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
            '■AutoBlow情報をFF専用BlowパターンへModify
            '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    
            Dim ModifyDataFFBlowAuto(ByteBit - 1) As String
            For iiii = 1 To ByteBit
                ModifyDataFFBlowAuto(iiii - 1) = "0"
            Next iiii
            
            '===== Modify Blow Pattern (Fixed Value) ======================
            TheHdw.Digital.Patterns.pat("OtpBlow_Break_Pat").ModifyPinVectorBlockData Label_OtpBlowAuto_Break, 0, RejiIn, ModifyDataFFBlowAuto

        End If


    Next NowPage

End Sub

Private Function MakeModifyData(ByRef InBinData() As String, ByRef OutBinData() As String, ByRef PageSelect() As Long, ByRef BitSelect() As Long, ByVal site As Long, Optional ByVal Case_SramRep As Boolean = False) As String
'OTPの個々のBlowデータを、Modify用のまとめた変数へ入れなおす

    Dim strSteps As Long
    Dim width As Long
    
    width = Len(InBinData(site))

    For strSteps = 0 To width - 1
    
        If Mid(InBinData(site), strSteps + 1, 1) = "0" Then OutBinData(site, PageSelect(strSteps), BitSelect(strSteps)) = "0"
        If Mid(InBinData(site), strSteps + 1, 1) = "1" Then OutBinData(site, PageSelect(strSteps), BitSelect(strSteps)) = "1"
               
        If Case_SramRep = False Then                       'SRAM冗長の場合はこのフラグは立てない。SRAM冗長専用フラグは別に保持。
            Flg_ModifyPage(PageSelect(strSteps)) = True    'このPageはBlowが必要ということで、BlowPageフラグをTrueにするよ。必要最低限のBlow実行でいいようにね。
        End If
        
    Next

End Function

Public Sub OtpInitialize_Get_PageBit(ByVal Label As String, ByVal BitWidthAll As Long, ByRef Page() As Long, ByRef Bit() As Long)
'OTP Initialize Sub.
'OTPの各変動値情報のPageとBit情報をGETするよ。

'Arg1 Input:  ラベル(変動値情報毎に持っているもの)。このラベルをOTPMAPシート内で検索して、引っかかったPageとBit番号を変数へ格納して保存しておくよ。
'Arg2 Input:  Bit幅。ラベルに割り当てられているOTPメモリBit数のこと。
'Arg3 Output: Page情報。ラベルに割り当てられているOTPメモリのあるBitが、どのPageに当たるのかを示す変数。
'Arg4 Output: Bit情報。ラベルに割り当てられているOTPメモリのあるBitが、どのBitに当たるのかを示す変数。

'注意事項（この関数の制約事項。下記の制約に違反している場合には動作保証しない。）


    Dim RowCount As Long
    Dim NowPage As Integer
    Dim ii As Long
    
    Worksheets("OTPMAP").Select                                                                             'OTP情報が記載されているSheetを選択(OTPMAP)
        
    '========== Get Page&Address Start Infomation =============================
    Do While Cells(OtpInfoSheet_Row_BlowInfo + RowCount, OtpInfoSheet_Column_BlowInfo).Value <> Label       'OTP情報の先頭から、ラベルまでの行数をカウント
        RowCount = RowCount + 1
        If Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value = "" Then Exit Do        '無限LOOP防止
    Loop
    
    Do While ii <> BitWidthAll
        If Cells(OtpInfoSheet_Row_BlowInfo + RowCount, OtpInfoSheet_Column_BlowInfo).Value = Label Then     'カウントした行数から、このラベルが持つBit幅分のPage情報とBit番号情報をGET
            Page(ii) = Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value
            Bit(ii) = Cells(OtpInfoSheet_Row_Bit + RowCount, OtpInfoSheet_Column_Bit).Value
            ii = ii + 1
        End If
        RowCount = RowCount + 1
        If Cells(OtpInfoSheet_Row_Page + RowCount, OtpInfoSheet_Column_Page).Value = "" Then Exit Do        '無限LOOP防止
    Loop
    
End Sub

Public Function ActiveSite_Check_OTP() As Long
'OTP Standard Function.
'ActiveなSiteにフラグを立てる。この後、ActiveSiteを一時的にDisableにする時などに用いる。
    
    Erase Flg_ActiveSite_OTP
    Dim site As Long

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Flg_ActiveSite_OTP(site) = 1
        End If
    Next site

End Function

Public Function ActiveSite_Return_OTP() As Long
'OTP Standard Function.
'ActiveSite_Checkでフラグを立てておいたSiteをActiveにする。これ以前に一時的にDisableにしていたSiteをActiveに戻す時などに用いる。

    Dim site As Long

    For site = 0 To nSite
        If Flg_ActiveSite_OTP(site) = 1 Then
            TheExec.sites.site(site).Active = True
        End If
    Next site

End Function

Public Sub Output_OtpBlowData()

'OTPBLOWのデバッグログ吐き出し関数
'データログを保存後、Excelに貼り付けて、『:』で区切ればきれいに並びます。
'固定値は、OTPMAPとそのまま比較すればいいよ。

Dim site As Long
Dim NowBit As Long
Dim NowPage As Long
Dim PageA As Long
Dim Data As String
Dim PageData As String
Dim first As Boolean
first = True
    
    TheExec.Datalog.WriteComment "OTP BLOW DATA"
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            PageA = 0
            For NowBit = 0 To BitParPage(PageA) - 1

                For NowPage = 0 To OtpPageSize - 1
                    If first = True Then
                        PageData = PageData + " : " & "P" & NowPage
                    End If
                    Data = Data + " :  " & BlowDataAllBin(site, NowPage, NowBit)
                Next NowPage
                
                If first = True Then
                    TheExec.Datalog.WriteComment "---- - : --- - : ----" & PageData
                End If
                
                first = False
                TheExec.Datalog.WriteComment "Site " & site & " : " & "Bit " & NowBit & " : " & "Data" & Data
                Data = ""
                
                If PageA < OtpPageSize Then
                    PageA = PageA + 1
                End If

            Next NowBit
            
            PageData = ""
            first = True
        End If
    Next site
                    

End Sub

Public Sub Output_OtpReadData()

Dim site As Long
Dim NowBit As Long
Dim NowPage As Long
Dim PageA As Long
Dim Data As String
Dim PageData As String
Dim first As Boolean
first = True


    Dim site_status As Long
    Dim NowCapData As Long
    Dim VectorOffset As Long
    Dim HramLoop As Long
    Dim HramSizeMod As Long
    Dim NowHramLoop As Long
    Dim HramSetSize As Long
    Const HramSize As Integer = 256
    Const BitParByte As Long = 8
    Dim DataOffset As Long
    Dim ReadErr(nSite) As Double
    Dim Deb_ReadDataAllBin(nSite, OtpPageSize - 1, OtpMaxBitParPage - 1) As String
    


    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■RollCall実行
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    If TheExec.sites.ActiveCount > 0 Then
        
        For NowPage = 0 To OtpPageEnd
            
            '========== PATTERN LOAD ==========================================
            With TheHdw.Digital
                .Timing.Load
                .Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").Load
            End With
        
            DataOffset = 0
            VectorOffset = 0
            HramLoop = Int(BitParPage(NowPage) / HramSize) + 1
            HramSizeMod = BitParPage(NowPage) Mod HramSize
            If HramSizeMod = 0 Then
                HramLoop = HramLoop - 1
                HramSizeMod = HramSize
            End If
            
            For NowHramLoop = 0 To (HramLoop - 1)
                
                '========== HRAM SETUP OFFSET CALCULATE ===========================
                If NowHramLoop <> (HramLoop - 1) Then
                    VectorOffset = NowHramLoop * (HramSize / BitParByte * ByteParVector_VerifyPat)
                    HramSetSize = HramSize
                Else
                    VectorOffset = NowHramLoop * (HramSize / BitParByte * ByteParVector_VerifyPat)
                    HramSetSize = HramSizeMod
                End If
                
                '========== HRAM SETUP ============================================
                Call StopPattern
                With TheHdw.Digital
                    .Patgen.EventCycleEnabled = False
                    .Patgen.NoHaltMode = noHaltAlways                                                           ' noHaltAlways:パターンの中に記述してあるHalt、又はVBA上でHaltが実行されるまでパターンは止まらない
                    .Patgen.EventSetVector True, "OtpVerifyPage" & CStr(NowPage) & "_Pat", Label_OtpVerify & CStr(NowPage), VectorOffset
                    .HRAM.SetTrigger trigFirst, True, 0, True                                                           ' trigFail:最初のFailサイクルから取り込み開始    True:EventCycleCountを有効    0:取り込み開始サイクルの何サイクル前から取り込むか   True:取り込んでいるサイクル数がHRAMサイズに達した場合にはそこで取り込みをやめる
                    .HRAM.SetCapture captSTV, True                                                               ' captFail:Failサイクルのみを取り込む   Ture: リピート文のVectorである場合は最後のリピートサイクルだけ取り込む
                    .HRAM.Size = HramSetSize                                                                       ' とりあえず最大値を指定
                End With
                
                '========== PATTERN RUN ===========================================
                TheHdw.Digital.Patterns.pat("OtpVerifyPage" & CStr(NowPage) & "_Pat").Run ("Start")

                '========== HRAM DATA SITE ============================
                site_status = TheExec.sites.SelectFirst
                While site_status <> loopDone
                    site = TheExec.sites.SelectedSite
                    
                    For NowCapData = 0 To (HramSize - 1)
                        If TheHdw.Digital.HRAM.Pins(RejiOut).PinData(NowCapData) = "L" Then
                            Deb_ReadDataAllBin(site, NowPage, NowCapData + DataOffset) = "0"
                        ElseIf TheHdw.Digital.HRAM.Pins(RejiOut).PinData(NowCapData) = "H" Then
                            Deb_ReadDataAllBin(site, NowPage, NowCapData + DataOffset) = "1"
                        End If
                    Next NowCapData
                    
                    site_status = TheExec.sites.SelectNext(site_status)
                Wend
                
                DataOffset = DataOffset + HramSize

            Next NowHramLoop
            
        Next NowPage
        
        '========== PATTERN SETUP =================================
        With TheHdw.Digital.Patgen
            .EventCycleEnabled = False
            .EventCycleCount = 1
            .MaskTilCycle = False
        End With

    End If



    TheExec.Datalog.WriteComment "OTP READ DATA"
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            PageA = 0
            For NowBit = 0 To BitParPage(PageA) - 1

                For NowPage = 0 To OtpPageSize - 1
                    If first = True Then
                        PageData = PageData + " : " & "P" & NowPage
                    End If
                    Data = Data + " :  " & Deb_ReadDataAllBin(site, NowPage, NowBit)
                Next NowPage
                
                If first = True Then
                    TheExec.Datalog.WriteComment "---- - : --- - : ----" & PageData
                End If
                
                first = False
                TheExec.Datalog.WriteComment "Site " & site & " : " & "Bit " & NowBit & " : " & "Data" & Data
                Data = ""
                
                If PageA < OtpPageSize Then
                    PageA = PageA + 1
                End If

            Next NowBit
            
            PageData = ""
            first = True
        End If
    Next site
                               

End Sub

Private Sub Get_ChipId_Debug()
'OTP Debug Sub. GET LOTNo/WaferNo/ChipNo Dummy Data.

    Dim TypeName_OTP_Debug As String
    TypeName_OTP_Debug = Mid(NormalJobName, 4, 3)
            
    LotName = "98M" & TypeName_OTP_Debug & "76543"
    WaferNo = "4"
    
    Dim site As Long
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            DeviceNumber_site(site) = site + 1
        End If
    Next site

End Sub

Public Sub MakeBlowData_Lot(ByVal BitWidth As Integer, ByRef Page() As Long, ByRef Bit() As Long, ByRef Beam As Long)

    Dim site As Long
    Dim BlowData(nSite) As String

    If Flg_OTP_BLOW = 0 Then Exit Sub
    
    '========== For Otp Blow Debug ============================================
    If Flg_AutoMode = False And Flg_OTP_BLOW = 1 Then Call Get_ChipId_Debug
    
    '========== Dec -> Bin ============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then

            BlowData(site) = "XXXX"

            If Mid(LotName, Beam, 1) = "0" Then BlowData(site) = "0000"
            If Mid(LotName, Beam, 1) = "1" Then BlowData(site) = "0001"
            If Mid(LotName, Beam, 1) = "2" Then BlowData(site) = "0010"
            If Mid(LotName, Beam, 1) = "3" Then BlowData(site) = "0011"
            If Mid(LotName, Beam, 1) = "4" Then BlowData(site) = "0100"
            If Mid(LotName, Beam, 1) = "5" Then BlowData(site) = "0101"
            If Mid(LotName, Beam, 1) = "6" Then BlowData(site) = "0110"
            If Mid(LotName, Beam, 1) = "7" Then BlowData(site) = "0111"
            If Mid(LotName, Beam, 1) = "8" Then BlowData(site) = "1000"
            If Mid(LotName, Beam, 1) = "9" Then BlowData(site) = "1001"
            If Mid(LotName, Beam, 1) = "A" Then BlowData(site) = "1010"
            If Mid(LotName, Beam, 1) = "B" Then BlowData(site) = "1011"
            If Mid(LotName, Beam, 1) = "C" Then BlowData(site) = "1100"
            If Mid(LotName, Beam, 1) = "D" Then BlowData(site) = "1101"
            If Mid(LotName, Beam, 1) = "E" Then BlowData(site) = "1110"
            If Mid(LotName, Beam, 1) = "F" Then BlowData(site) = "1111"

            If Mid(LotName, Beam, 1) = "S" Then BlowData(site) = "1000"
            If Mid(LotName, Beam, 1) = "M" Then BlowData(site) = "0100"

            If BlowData(site) = "XXXX" Then
                blnFlg_BlowCheck = True
            End If

            '========== Blow用変数作成 ========================================
            Call MakeModifyData(BlowData, BlowDataAllBin, Page, Bit, site)

        End If
    Next site
    
    '### DEBUG LOG OUTPUT ###
    If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site**]" & " " & "LotNo" & Beam & " = " & Mid(LotName, Beam, 1)
    
End Sub

Public Sub MakeBlowData_Wafer(ByVal BitWidth As Integer, ByRef Page() As Long, ByRef Bit() As Long)

    Dim site As Long
    Dim BlowData(nSite) As String

    If Flg_OTP_BLOW = 0 Then Exit Sub

    '========== For Otp Blow Debug ============================================
    If Flg_AutoMode = False And Flg_OTP_BLOW = 1 Then Call Get_ChipId_Debug
    
    '========== Dec -> Bin ============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
                        
            BlowData(site) = Dec2Bin(CStr(WaferNo), BitWidth)
            
            '========== Blow用変数作成 ========================================
            Call MakeModifyData(BlowData, BlowDataAllBin, Page, Bit, site)

        End If
    Next site
    
    '### DEBUG LOG OUTPUT ###
    If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site**]" & " " & "WaferNo" & " = " & CStr(WaferNo)

End Sub

Public Sub MakeBlowData_Chip(ByVal BitWidth As Integer, ByRef Page() As Long, ByRef Bit() As Long)

    Dim site As Long
    Dim BlowData(nSite) As String

    If Flg_OTP_BLOW = 0 Then Exit Sub

    '========== For Otp Blow Debug ============================================
    If Flg_AutoMode = False And Flg_OTP_BLOW = 1 Then Call Get_ChipId_Debug
    
    '========== Dec -> Bin ============================================
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
                
            BlowData(site) = Dec2Bin(CStr(DeviceNumber_site(site)), BitWidth)

            '========== Blow用変数作成 ========================================
            Call MakeModifyData(BlowData, BlowDataAllBin, Page, Bit, site)

            '### DEBUG LOG OUTPUT ###
            If Flg_Debug = 1 Then TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "ChipNo" & " = " & CStr(DeviceNumber_site(site))

        End If
    Next site
    
    
End Sub

Public Sub MakeBlowData_Temp(ByVal BitWidth_Ofs As Integer, ByVal BitWidth_Slp As Integer, ByRef Page() As Long, ByRef Bit() As Long, ByVal TEMP_OFS As String, ByVal TEMP_SLP As String)
                                
'OTP欠陥補正Blowデータ生成用Sub（TEMP系補正）

'Arg1 Input:温度計オフセット情報のBit幅
'Arg2 Input:温度計傾き情報のBit幅
'Arg3 Input:各Blow情報Bitの対象Page番号
'Arg4 Input:各Blow情報Bitの対象Bit番号
'Arg5 Input:温度計オフセット情報のラベル名（項目名）
'Arg6 Input:温度計傾き情報のラベル名（項目名）

'注意事項（この関数の制約事項。下記の制約に違反している場合には動作保証しない。）
'制約1:温度計のオフセット情報と傾き情報は連続Bitであること。さらに、オフセット情報の方が若いアドレスであること。


    Dim site As Long
    Dim TempInfo_Ofs() As Double
    Dim TempInfo_Slp() As Double
    Dim AllTempInfo(nSite) As String
    Dim data_ofs(nSite) As String
    Dim data_slp(nSite) As String

    '========== GET TEMP INFOMATION =======================
    TheResult.GetResult TEMP_OFS, TempInfo_Ofs
    TheResult.GetResult TEMP_SLP, TempInfo_Slp
    
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
        
            '========== Dec->Bin & Offset + Slope =========================
            data_ofs(site) = Dec2Bin(CStr(TempInfo_Ofs(site)), BitWidth_Ofs)
            data_slp(site) = Dec2Bin(CStr(TempInfo_Slp(site)), BitWidth_Slp)
            
            AllTempInfo(site) = data_ofs(site) & data_slp(site)
        
            '========== Blow用変数作成 ====================================
            Call MakeModifyData(AllTempInfo, BlowDataAllBin, Page, Bit, site)


            '##### DEBUG LOG OUTPUT #####
            If Flg_Debug = 1 Then
                TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "TEMP:" & "OFS" & " = " & TempInfo_Ofs(site)
                TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "TEMP:" & "SLP" & " = " & TempInfo_Slp(site)
            End If

        End If
    Next site
                        
End Sub

Public Sub MakeBlowData_Sram(tsite As Long, ByVal BitWidth_En As Integer, ByVal BitWidth_Addr As Integer, ByVal BitWidth_Data As Integer, ByRef Page() As Long, ByRef Bit() As Long)
                                
'OTP欠陥補正Blowデータ生成用Sub（SRAM冗長）

'Arg1 Input:SRAM冗長のBit幅。Fuse冗長や未使用Bitは含めたら駄目です
'Arg2 Input:各Blow情報Bitの対象Page番号
'Arg3 Input:各Blow情報Bitの対象Bit番号

'注意事項（この関数の制約事項。下記の制約に違反している場合には動作保証しない。）
'制約1:EN -> Address -> Data の並びで1冗長データが構成されている仕様であること。


    Dim RepairNo As Integer
    Dim BlowData_SramRep(nSite) As String
    Dim i As Long
    
    '===== SRAM冗長データをつなげる ===========================================
    For RepairNo = 1 To MAX_EF_BIST_RD_BIT
        BlowData_SramRep(tsite) = BlowData_SramRep(tsite) + Dec2Bin(CStr(Ef_Enbl_Addr(tsite, RepairNo - 1)), BitWidth_En)
        BlowData_SramRep(tsite) = BlowData_SramRep(tsite) + Dec2Bin(CStr(Ef_Rcon_Addr(tsite, RepairNo - 1)), BitWidth_Addr)
        BlowData_SramRep(tsite) = BlowData_SramRep(tsite) + Dec2Bin(CStr(Ef_Repr_Data(tsite, RepairNo - 1)), BitWidth_Data)
    Next RepairNo

    '========== Blow用変数作成 ====================================
    Call MakeModifyData(BlowData_SramRep, SramBlowDataAllBin, Page, Bit, tsite, True)
    
    '========== SRAM冗長のModifyを行うPageを選択 ==================
    For i = 0 To Len(BlowData_SramRep(tsite)) - 1
        Flg_ModifyPageSRAM(Page(i)) = True
    Next i
                        
End Sub

Public Sub MakeBlowData_Defect_SinCpFd(ByVal MaxRepairNo As Integer, ByVal BitWidthN As Integer, ByVal BitWidthX As Integer, ByVal BitWidthY As Integer, ByVal BitWidthS As Integer, ByVal BitWidthD As Integer, _
                                       ByVal SourceType As String, ByVal NgAdd_LeftS As Long, ByVal NgAdd_LeftE As Long, ByVal NgAdd_RightS As Long, ByVal NgAdd_RightE As Long, _
                                       ByRef Page() As Long, ByRef Bit() As Long, ByRef OverFlowCheck() As Double, ByRef NgAddressCheck() As Double, ParamArray DefectInfo() As Variant)
                                
'OTP欠陥補正Blowデータ生成用Sub（Single/Couplet/FD補正領域が共通のタイプ）

'Arg1  Input:最大可能補正個数
'Arg2  Input:個数情報のbit幅
'Arg3  Input:Xアドレス情報のbit幅(1欠陥当たりの)
'Arg4  Input:Yアドレス情報のbit幅(1欠陥当たりの)
'Arg5  Input:Sorce情報のbit幅(1欠陥当たりの)
'Arg6  Input:Direction情報のbit幅(1欠陥当たりの)
'Arg7  Input:Sorce情報のタイプ
'Arg9  Input:欠陥が存在してはいけないZONEの左側スタートアドレス
'Arg10 Input:欠陥が存在してはいけないZONEの左側エンドアドレス
'Arg11 Input:欠陥が存在してはいけないZONEの右側スタートアドレス
'Arg12 Input:欠陥が存在してはいけないZONEの右側エンドアドレス
'Arg13 Input:各Blow情報Bitの対象Page番号
'Arg14 Input:各Blow情報Bitの対象Bit番号
'Arg15 Output:補正上限オーバーをお知らせするフラグ
'Arg16 Output:欠陥が存在してはいけないZONEの欠陥有無をお知らせするフラグ
'Arg17以降 Input:撮像から引き渡される欠陥情報。DKやHLといった欠陥種類数に応じて、Arg18やArg19と引数が増える。


'注意事項（この関数の制約事項。下記の制約に違反している場合には動作保証しない。）
'制約1：Blowデータ全体のつながりは（欠陥が3個だった場合の例）、個数⇒欠陥①SourceSorce⇒欠陥①Direction⇒欠陥①Xアドレス⇒欠陥①Yアドレス⇒欠陥②SourceSorce⇒欠陥②Direction⇒欠陥②Xアドレス⇒欠陥②Yアドレス⇒欠陥③SourceSorce⇒欠陥③Direction⇒欠陥③Xアドレス⇒欠陥③Yアドレス⇒　という仕様であること。
'制約2：OTP内の欠陥補正のBlowbitアドレスが個数情報Blowから最終Direction情報まで全て連続していること。個数⇒欠陥①SourceSorce⇒CP欠陥①Direction⇒欠陥①Xアドレス⇒CP欠陥①Yアドレス⇒　LOT名　⇒欠陥②SourceSorce⇒欠陥②Direction⇒欠陥②Xアドレス⇒C陥②Yアドレス　みたいな仕様はダメ。
    
    Dim site As Long
    
    Dim BlowData_All(nSite) As String
    Dim amari(nSite) As Integer
    Dim ParamLoop As Long
    Dim ParamNo As Long
    ParamNo = UBound(DefectInfo)
    Dim ParamLabel As String
    Dim NowRepairNo As Integer
    Dim i(nSite) As Long
    Dim SameAddress As Boolean
    Dim Go_NgAddCheck As Boolean
    Dim NgLeft As Long
    Dim NgRight As Long
    
    Dim DefInfo_Num() As Double
    Dim DefInfo_Hadd() As Double
    Dim DefInfo_Vadd() As Double
    Dim DefInfo_Dire() As Double
    Dim DefInfo_Src1() As Double
    Dim DefInfo_Src2 As Double

    Dim AllDefInfo_Num(nSite) As Double
    ReDim AllDefInfo_Hadd(nSite, MaxRepairNo - 1) As Double
    ReDim AllDefInfo_Vadd(nSite, MaxRepairNo - 1) As Double
    ReDim AllDefInfo_Dire(nSite, MaxRepairNo - 1) As Double
    ReDim AllDefInfo_Src(nSite, MaxRepairNo - 1) As Double
    
    For ParamLoop = 0 To ParamNo                                                                            '欠陥の種類数分LOOP
        ParamLabel = CStr(DefectInfo(ParamLoop))                                                            '欠陥項目名をGET
        
        '========== GET DEFECT INFOMATION ===============================                                   '現在の欠陥項目のDefect情報をGET
        TheResult.GetResult ParamLabel & "_Info_Num", DefInfo_Num
        TheResult.GetResult ParamLabel & "_Info_Hadd", DefInfo_Hadd
        TheResult.GetResult ParamLabel & "_Info_Vadd", DefInfo_Vadd
        TheResult.GetResult ParamLabel & "_Info_Dire", DefInfo_Dire
        TheResult.GetResult ParamLabel & "_Info_Sorc", DefInfo_Src1
        
        For site = 0 To nSite                                                                               'siteまわし
            If TheExec.sites.site(site).Active = True Then
                If OverFlowCheck(site) = 0 Then                                                             '補正上限をオーバーしていなければ、補正用変数に補正情報を格納

                    For NowRepairNo = 0 To DefInfo_Num(site) - 1                                            '現在の欠陥項目の個数分LOOP
                        SameAddress = False                                                                 '欠陥アドレスかぶりCheckフラグをクリア
                        
                        '===== ADDRESS CHECK =====
                        If AllDefInfo_Num(site) > 0 Then
                            For i(site) = 0 To AllDefInfo_Num(site) - 1
                                If AllDefInfo_Hadd(site, i(site)) = DefInfo_Hadd(site, NowRepairNo) And _
                                   AllDefInfo_Vadd(site, i(site)) = DefInfo_Vadd(site, NowRepairNo) Then
                                    SameAddress = True                                                      '欠陥アドレスかぶりが発生したらフラグをTrue
                                    Exit For
                                End If
                            Next i
                        End If
                        
                        '===== DEFECT ADDITION =====
                        If SameAddress = False Then                                                         '欠陥アドレスかぶりでなければ、補正用変数に補正情報を格納
                            AllDefInfo_Num(site) = AllDefInfo_Num(site) + 1                                 'Tatalの補正個数をインクリメント
                            
                            If AllDefInfo_Num(site) <= MaxRepairNo Then                                     '補正上限を超えていないかをチェック
                                '----- X address -----
                                AllDefInfo_Hadd(site, AllDefInfo_Num(site) - 1) = DefInfo_Hadd(site, NowRepairNo)
                                '----- Y address -----
                                AllDefInfo_Vadd(site, AllDefInfo_Num(site) - 1) = DefInfo_Vadd(site, NowRepairNo)
                                '----- Direction -----
                                AllDefInfo_Dire(site, AllDefInfo_Num(site) - 1) = DefInfo_Dire(site, NowRepairNo)
                                '----- Sorce -----
                                Call BlowSorceInfoCheck(SourceType, DefInfo_Src1(site, NowRepairNo), DefInfo_Src2, Go_NgAddCheck)    'Single(PD)/CP/FDのSourceコードの確認
                                AllDefInfo_Src(site, AllDefInfo_Num(site) - 1) = DefInfo_Src2
                            
                                '----- NG ADDRESS CHECK -----
                                If Go_NgAddCheck = True Then
                                    '----- ZONE LEFT (Center Pix) -----
                                    If NgAdd_LeftS <= DefInfo_Hadd(site, NowRepairNo) And DefInfo_Hadd(site, NowRepairNo) <= NgAdd_LeftE Then
                                        NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                    End If
                                    '----- ZONE LEFT (Direction Pix) -----
                                    If DefInfo_Dire(site, NowRepairNo) = 0 Then     'Direction=Right
                                        If NgAdd_LeftS <= DefInfo_Hadd(site, NowRepairNo) + 2 And DefInfo_Hadd(site, NowRepairNo) + 2 <= NgAdd_LeftE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 1 Then     'Direction=RightBottom
                                        If NgAdd_LeftS <= DefInfo_Hadd(site, NowRepairNo) + 2 And DefInfo_Hadd(site, NowRepairNo) + 2 <= NgAdd_LeftE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 2 Then     'Direction=Bottom
                                        'Bottom = X-Address is Same Center Address
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 3 Then     'Direction=LeftBottom
                                        If NgAdd_LeftS <= DefInfo_Hadd(site, NowRepairNo) - 2 And DefInfo_Hadd(site, NowRepairNo) - 2 <= NgAdd_LeftE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                    
                                    '----- ZONE RIGHT (Center Pix) -----
                                    If NgAdd_RightS <= DefInfo_Hadd(site, NowRepairNo) And DefInfo_Hadd(site, NowRepairNo) <= NgAdd_RightE Then
                                        NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                    End If
                                    '----- ZONE RIGHT (Direction Pix) -----
                                    If DefInfo_Dire(site, NowRepairNo) = 0 Then     'Direction=Right
                                        If NgAdd_RightS <= DefInfo_Hadd(site, NowRepairNo) + 2 And DefInfo_Hadd(site, NowRepairNo) + 2 <= NgAdd_RightE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 1 Then     'Direction=RightBottom
                                        If NgAdd_RightS <= DefInfo_Hadd(site, NowRepairNo) + 2 And DefInfo_Hadd(site, NowRepairNo) + 2 <= NgAdd_RightE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 2 Then     'Direction=Bottom
                                        'memo : Bottom = X-Address is Same Center Address
                                    End If
                                    If DefInfo_Dire(site, NowRepairNo) = 3 Then     'Direction=LeftBottom
                                        If NgAdd_RightS <= DefInfo_Hadd(site, NowRepairNo) - 2 And DefInfo_Hadd(site, NowRepairNo) - 2 <= NgAdd_RightE Then
                                            NgAddressCheck(site) = 1 + NgAddressCheck(site)
                                        End If
                                    End If
                                
                                End If

                            Else
                                OverFlowCheck(site) = 1                                                     '補正上限を超えていたらフラグを立てる
                                Exit For
                            End If
                        End If
                    Next NowRepairNo
    
                End If
                
                '##### DEBUG LOG OUTPUT #####
                If Flg_Debug = 1 Then
                    TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Num" & " = " & DefInfo_Num(site)
                    For NowRepairNo = 0 To DefInfo_Num(site) - 1
                        TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Hadd" & " = " & DefInfo_Hadd(site, NowRepairNo)
                        TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Vadd" & " = " & DefInfo_Vadd(site, NowRepairNo)
                        If DefInfo_Src1(site, NowRepairNo) = 0 Then
                            TheExec.Datalog.WriteComment "SourceCode Nothing"
                        End If
                        If DefInfo_Src1(site, NowRepairNo) = 1 Then
                            TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Sorc" & " = " & "Single"
                        End If
                        If DefInfo_Src1(site, NowRepairNo) = 2 Then
                            TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Sorc" & " = " & "Couplet"
                        End If
                        If DefInfo_Src1(site, NowRepairNo) = 3 Then
                            TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Sorc" & " = " & "FD"
                        End If
                        TheExec.Datalog.WriteComment "Blow Infomation" & " " & "[Site" & site & "]" & " " & "DEFECT-REP:" & ParamLabel & "_Info_Dire" & " = " & DefInfo_Dire(site, NowRepairNo)
                    Next NowRepairNo
                End If
                
            End If
        Next site
                        
    Next ParamLoop


    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
                        
            If OverFlowCheck(site) = 0 And NgAddressCheck(site) = 0 Then
            
                '========== 補正個数情報をModify変数へ格納 ================
                BlowData_All(site) = Dec2Bin(CStr(AllDefInfo_Num(site)), BitWidthN)
                
                '========== CUPLET補正アドレス情報をModify変数へ格納 ======
                For NowRepairNo = 0 To AllDefInfo_Num(site) - 1
                        BlowData_All(site) = BlowData_All(site) _
                                            + Dec2Bin(CStr(AllDefInfo_Src(site, NowRepairNo)), BitWidthS) _
                                            + Dec2Bin(CStr(AllDefInfo_Dire(site, NowRepairNo)), BitWidthD) _
                                            + Dec2Bin(CStr((AllDefInfo_Hadd(site, NowRepairNo) + OtpPixOffset_X)), BitWidthX) _
                                            + Dec2Bin(CStr((AllDefInfo_Vadd(site, NowRepairNo) + OtpPixOffset_Y)), BitWidthY)
                Next NowRepairNo
                        
                
                '========== 補正上限に足りない時は"0"でうめる =============　　'前チップのModify情報がパターンに残ってるから余りには0書きが必要
                amari(site) = MaxRepairNo - AllDefInfo_Num(site)
                
                If amari(site) > 0 Then
                    For NowRepairNo = 0 To amari(site) - 1
                            BlowData_All(site) = BlowData_All(site) _
                                                + Dec2Bin("0", BitWidthS) _
                                                + Dec2Bin("0", BitWidthD) _
                                                + Dec2Bin("0", BitWidthX) _
                                                + Dec2Bin("0", BitWidthY)
                    Next NowRepairNo
                End If
                            
                '========== Blow用変数作成 ====================================
                Call MakeModifyData(BlowData_All, BlowDataAllBin, Page, Bit, site)

            End If
            
        End If
    Next site

End Sub

Public Sub BlowSorceInfoCheck(ByVal SourceType As String, ByRef SrcIn As Double, ByRef SrcOut As Double, ByRef Go_NgAddCheck As Boolean)
'2013/8/2　SrcType1のみ。Sourceコードのナンバリングが変わったタイプが出てきたら、その都度追加が必要。

    Go_NgAddCheck = False
    
    Select Case SourceType
        Case "SrcType1"
            If SrcIn = 1 Then SrcOut = 1
            If SrcIn = 2 Then
                SrcOut = 2
                Go_NgAddCheck = True
            End If
            If SrcIn = 3 Then SrcOut = 3
        Case Else
            MsgBox "Error!! Nothing SourceType!!"
            Stop
    End Select
        
End Sub

