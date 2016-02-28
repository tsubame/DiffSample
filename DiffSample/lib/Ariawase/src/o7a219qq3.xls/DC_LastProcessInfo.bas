Attribute VB_Name = "DC_LastProcessInfo"
'概要:
'　前工程で作成された不良情報ファイルを読み込む
'
'目的:
'　前工程検査にてNGになったチップを強制不良とする。
'
'作成者:
'   2012/10/16 Ver0.1 A.Hamaya
'   2013/01/16 Ver0.2 A.Hamaya
'   2013/01/24 Ver0.3 H.Arikawa 長崎を想定した処理を追加。
'
'
'使用するには、下記コードをdc_setupへ記入
'
'    '### 初期化 ###
'    If Flg_AutoMode = True Then
'        If CInt(DeviceNumber_site(0)) = 1 Then  'デバイスNo.が1の時
'            Call Init_LastProcessInfoFILE
'        End If
'    End If
'

Option Explicit

#Const EEE_AUTO_JOB_LOCATE = 2      '1:長崎200mm,2:長崎300mm,3:熊本

'### 関数定義 ###
Public USonic(nSite) As Double
Public Wasavi(nSite) As Double
Public PadClo(nSite) As Double
Public Fmura(nSite) As Double
Public Proces(nSite) As Double

'### 各工程のナンバー定義 ###
Private Const USonicNum As Integer = 1
Private Const WasaviNum As Integer = 2
Private Const PadCloNum As Integer = 3
Private Const FmuraNum As Integer = 4

Private NowWaferID As String
Private NGchipCNT As Integer            'NGチップの個数
Private NGdataCNT As Integer            'データの個数
Private NGChipNo() As String            'NGチップNo.格納用
Private NGChipData() As String          'NGチップデータ格納用
Private FileStatus() As String          'ファイル状態格納用 exist/not-exist

Private flg_NoFILE As Boolean           'エンドファイルが無かった場合に立つフラグ
Private flg_NoEndFILE As Boolean        'エンドファイルが無かった場合に立つフラグ

Private LastProcessInfoFILE As String       'ファイル本体
Private LastProcessInfoFILE_END As String   'ファイルのエンドファイル

Private Const LastProcessInfo_FilePATH_K As String = "f:\job\failchipdetection\"        '前工程不良情報ファイルのパス(熊本)
Private Const LastProcessInfo_FilePATH_N As String = "f:\job\failchipdetection\"        '前工程不良情報ファイルのパス(長崎)　予約

Public Function ultrasonic_f() As Double

    Dim site As Long
    Call SiteCheck

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            USonic(site) = Get_NgChipInfo(USonicNum, site)
        End If
    Next site
    
    Call test(USonic)

End Function

Public Function wasavi_f() As Double

    Dim site As Long
    Call SiteCheck

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Wasavi(site) = Get_NgChipInfo(WasaviNum, site)
        End If
    Next site

    Call test(Wasavi)

End Function

Public Function padclosing_f() As Double

    Dim site As Long
    Call SiteCheck

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            PadClo(site) = Get_NgChipInfo(PadCloNum, site)
        End If
    Next site

    Call test(PadClo)

End Function

Public Function fmura_f() As Double

    Dim site As Long
    Call SiteCheck

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Fmura(site) = Get_NgChipInfo(FmuraNum, site)
        End If
    Next site

    Call test(Fmura)

End Function

'###　強制不良チップがあった場合、この関数でFailにする。 ###

Public Function processng_f() As Double

    Dim site As Long
    Call SiteCheck

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then

            If USonic(site) = 1 Or Wasavi(site) = 1 Or PadClo(site) = 1 Or Fmura(site) = 1 Then
                Proces(site) = 1
            Else
                Proces(site) = 0
            End If
        
        End If
    Next site

    Call test(Proces)

End Function

'現在のウェーハIDを取得する

Private Sub Get_waferID()
    
    Dim typeFolder As String
    
    '### Production IFシート読み込み ###
    Dim wkshtObj As Object
    Set wkshtObj = ThisWorkbook.Sheets("Production IF")
    '======= WorkSheet ErrorProcess ========
    If wkshtObj Is Nothing Then
        MsgBox "Not Find Sheet : " & " Production IF"
        Exit Sub
    End If

    '### Production IFシートからWaferIDを取得 ###
'    NowWaferID = wkshtObj.Cells(WaferNo + 2, 10)
    NowWaferID = "ESD105706-08"         'DEBUG!!!!!!!
    
    typeFolder = Mid(NowWaferID, 3, 4)  'ex)M105
    
    NowWaferID = typeFolder + "\" + NowWaferID      'ex)M105\29M105001-01
    
End Sub

Private Function Open_File() As Boolean

    Dim FileNo As Integer                   'ファイルナンバー
    Dim strText As String                   '読み込んだ内容を格納します。
    Dim i, j As Integer
    
    Dim fileData, fileData2 As Variant      'ファイルから読み込んだNGチップデータ格納用
    
    Call Get_waferID    'ウェーハID取得
    
    #If EEE_AUTO_JOB_LOCATE = 1 Or EEE_AUTO_JOB_LOCATE = 2 Then
        LastProcessInfoFILE = LastProcessInfo_FilePATH_N & NowWaferID & ".txt"
        LastProcessInfoFILE_END = LastProcessInfo_FilePATH_N & NowWaferID & ".txt.END"
    #ElseIf EEE_AUTO_JOB_LOCATE = 3 Then
        LastProcessInfoFILE = LastProcessInfo_FilePATH_K & NowWaferID & ".txt"
        LastProcessInfoFILE_END = LastProcessInfo_FilePATH_K & NowWaferID & ".txt.END"
    #End If

    '### 対象のファイルが無ければ抜ける ###
    flg_NoFILE = False
    If Dir(LastProcessInfoFILE) = "" Then
        Open_File = False
        flg_NoFILE = True
        Exit Function
    End If

    '### 対象のエンドファイルが無ければ抜ける ###
    flg_NoEndFILE = False
    If Dir(LastProcessInfoFILE_END) = "" Then
        Open_File = False
        flg_NoEndFILE = True
        Exit Function
    End If

    '### ファイルを開く ###
    FileNo = FreeFile
    Open LastProcessInfoFILE For Input As #FileNo
    On Error GoTo CloseFile

    '### ファイルからNGチップ数／データ数を取得する ###
    NGchipCNT = 0
    NGdataCNT = 0
    Do Until EOF(FileNo)
        Line Input #FileNo, strText
        fileData = Split(strText, ":")
        fileData2 = Split(fileData(1), ",")
        '=== 2行目からデータ数を取得する ===
        If NGchipCNT = 1 Then
            For j = 0 To UBound(fileData2)
                NGdataCNT = NGdataCNT + 1               'データ個数
            Next j
        End If
        NGchipCNT = NGchipCNT + 1
    Loop
    NGchipCNT = NGchipCNT - 2                           'NGチップの個数

    '### ファイルを閉じる ###
    Close #FileNo

    Open_File = True


FILE_end:

Exit Function

CloseFile:

    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    Close #FileNo
    flg_NoEndFILE = True
'    MsgBox ("File Open Error! Please Check a File")
    GoTo FILE_end

End Function

'テキストファイルからデータ読み込み、配列変数へ書き込む

Public Function Init_LastProcessInfoFILE() As Boolean

    Dim FileNo As Integer                   'ファイルナンバー
    Dim strText As String                   '読み込んだ内容を格納します。
    Dim i, j As Integer
    
    Dim fileData, fileData2 As Variant      'ファイルから読み込んだNGチップデータ格納用
    
    '### File Search&Open ###
    If Open_File = False Then
        Exit Function
    End If
    '#################
    
    '--- 変数宣言 ---
    ReDim NGChipNo(NGchipCNT)                   'NGチップNo.格納用
    ReDim NGChipData(NGchipCNT, NGdataCNT)      'NGチップデータ格納用
    ReDim FileStatus(NGdataCNT)                 'ファイル状態格納用 exist/not-exist
    '----------------

    '### ファイルを開く ###
    FileNo = FreeFile
    Open LastProcessInfoFILE For Input As #FileNo
    On Error GoTo CloseFile

    '### ファイルからNGチップNo.とデータを取得する ###
    NGchipCNT = 0
    Do Until EOF(FileNo)
        Line Input #FileNo, strText
        fileData = Split(strText, ":")
        fileData2 = Split(fileData(1), ",")
        '=== 2行目のファイル状態読み込み ===
        If NGchipCNT = 1 Then
            If fileData(0) <> "File" Then
                GoTo CloseFile
            End If
            For j = 0 To UBound(fileData2)
                If fileData2(j) = "exist" Or fileData2(j) = "not-exist" Or fileData2(j) = "" Then
                    FileStatus(j) = fileData2(j)                    'ファイル状態取得
                Else
                    GoTo CloseFile
                End If
            Next j
        End If
        '=== 3行目からのデータ読み込み ===
        If NGchipCNT > 1 Then
            If CInt(fileData(0)) > 0 Then
                NGChipNo(NGchipCNT - 2) = fileData(0)                   'NGチップNo.
            Else
                GoTo CloseFile
            End If
            For j = 0 To UBound(fileData2)
                If fileData2(j) = "0" Or fileData2(j) = "1" Or fileData2(j) = "-1" Or fileData2(j) = "" Then
                    NGChipData(NGchipCNT - 2, j) = fileData2(j)     'NGチップデータ
                Else
                    GoTo CloseFile
                End If
            Next j
        End If
        NGchipCNT = NGchipCNT + 1
    Loop
    NGchipCNT = NGchipCNT - 2
    
    '### ファイルを閉じる ###
    Close #FileNo

FILE_end:

Exit Function

CloseFile:

    'xxxxxxxxx  FileOpenError Routine  xxxxxxx
    Close #FileNo
    flg_NoEndFILE = True
'    MsgBox ("File Open Error! Please Check a File")
    GoTo FILE_end
    
End Function

Public Function Get_NgChipInfo(failNUM As Integer, site As Long) As Integer

    Dim i, j As Integer

    Dim flg_Not_Exist As Boolean

    If Flg_AutoMode = True Then
    
        '### 対象のエンドファイルが無かった場合 ###
        If flg_NoFILE = True Then
            Exit Function
        End If
        
        '### 対象のエンドファイルが無かった場合 ###
        If flg_NoEndFILE = True Then
            Get_NgChipInfo = -1
            Exit Function
        End If

        '### ファイル状態の確認 ###
        If FileStatus(failNUM - 1) = "not-exist" Then
            Get_NgChipInfo = -1
            Exit Function
        End If

        '### 各工程の情報をテスト結果として返す ###
        For i = 0 To NGchipCNT - 1
            If CInt(DeviceNumber_site(site)) = CInt(NGChipNo(i)) Then
                If NGChipData(i, failNUM - 1) <> "" Then
                    Get_NgChipInfo = CInt(NGChipData(i, failNUM - 1))     'NGチップデータ取得
                End If
                Exit For
            End If
        Next i
        
    End If
    
End Function

