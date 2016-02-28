Attribute VB_Name = "RegisterVersionManagementMod"
Option Explicit
'   ///Version 1.1///
'
'   Update history
'Ver1.1 2013/10/9 H.Arikawa HashCodeエラー時にStopPMCで止めるように処理追加。

'================================================================================
' For Hash Code Definition
'================================================================================
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
                            (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
                             ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
                            (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" _
                            (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, _
                             ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
                            (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" _
                            (ByVal hHash As Long, pbData As Any, ByVal cbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
                            (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, ByRef pcbData As Long, _
                             ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL   As Long = 1
Private Const PROV_RSA_AES    As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Private Const HP_HASHVAL      As Long = 2
Private Const HP_HASHSIZE     As Long = 4

Private Const ALG_TYPE_ANY    As Long = 0
Private Const ALG_CLASS_HASH  As Long = 32768

Private Const ALG_SID_MD2     As Long = 1
Private Const ALG_SID_MD4     As Long = 2
Private Const ALG_SID_MD5     As Long = 3
Private Const ALG_SID_SHA     As Long = 4
Private Const ALG_SID_SHA_256 As Long = 12
Private Const ALG_SID_SHA_512 As Long = 14

Private Const CALG_MD2        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
Private Const CALG_MD4        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
Private Const CALG_MD5        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
Private Const CALG_SHA        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
Private Const CALG_SHA_256    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)
Private Const CALG_SHA_512    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512)

'================================================================================
' Module Definition
'================================================================================
Private Const RVVM_LARGE_FILE_SIZE_HASHCODE As String = "This File Size is too large."
Private Const RVVM_FILE_NOT_FOUND_HASHCODE As String = "This File is not found."
Private Const RVVM_DONT_MANAGEMENT_VERSION As String = "This File is not out of version management!"

Private Type HashCode_information
    PatName As String
    filePath As String
    HashCode As String
End Type

Private Enum RegisterVersionManagementModState
    
    RVMM_STATE_UNINITIALIZED = 0
    RVMM_STATE_INITIALIZED = 1
    RVMM_STATE_HASE_CREATED = 2

End Enum

'================================================================================
' Module Variables
'================================================================================
Private HashCode_Data() As HashCode_information '今のハッシュコードの情報
Private m_RvmmState As RegisterVersionManagementModState
Public Flg_HashCheckResult As Boolean           'ハッシュコードチェック結果フラグ

Public Function myState() As Long
    myState = m_RvmmState
End Function

'================================================================================
' Public Functions
'================================================================================
Public Sub RVMM_Initialize()

    Erase HashCode_Data
    
    '状態遷移
    m_RvmmState = RVMM_STATE_INITIALIZED
    
End Sub


Public Sub RVMM_CreateRegisterHashCode()
        
        
'    HASHCODEシートがなかったら動かない
    If Not IsHashCodeFunctionEnable() Then
        Call MsgBox("HashCode Sheet is not found! This function is disable!", vbCritical, "RVMM_CreateRegisterHashCode")
        Exit Sub
    End If
    
    '変数宣言
    Dim VerX As Long
    Dim HashCode_Data_BeforeVersion() As HashCode_information '一つ前のVersionのハッシュコードの情報
        
        
'    初期化チェック
    If m_RvmmState = RVMM_STATE_UNINITIALIZED Then
        Call MsgBox("Register Version Management Mod is not be initialized!", vbCritical, "RVMM_CreateRegisterHashCode")
        Exit Sub
    End If
    
    'LoadPatをしたものに関しては最後に消したくないので保存しておく
    Dim ArrayMax As Long
    ArrayMax = GetUBoundHashCode_Data(HashCode_Data)

'    LoadPatが呼ばれていることが前提
'    PatGrpシートとTestInstanceから情報収集をする｡
    Call GetHashCode_information
    

'    パスを全部調べてVersionが全部一緒でないとエラーとする
    If Not IsSamePattenVersion(HashCode_Data, VerX) Then
        Call MsgBox("The Register Versions is not same!", vbCritical, "RVMM_CreateRegisterHashCode")
        GoTo ErrorExit
    End If
    
    
'    HashCode_Dataを基準として
'    REGVERフォルダを全部みて､最新のREGVERフォルダでなければエラーとする
    If Not IsLatestRegversion(VerX) Then
        Call MsgBox("Pat files are not the latest!", vbCritical, "RVMM_CreateRegisterHashCode")
        GoTo ErrorExit
    End If
    
'    今のVersionのHASHCODEを作成する
    Call CreateHashCode_impl(HashCode_Data)
    
'  ファイルが存在しない場合はここで引っ掛ける
    Dim strNotFoundFile As String
    If Not IsAllFileHashCodeCreated(HashCode_Data, strNotFoundFile) Then
        Call MsgBox("Pat file " & strNotFoundFile & " is not found!", vbCritical, "RVMM_CreateRegisterHashCode")
        GoTo ErrorExit
    End If
    
'    今のVersionが1でないなら、前のバージョンのパスを作成し、ハッシュコードを作成する
    If VerX <> 1 Then
        Call ConvertBeforeVersionPath(HashCode_Data, HashCode_Data_BeforeVersion)
        Call CreateHashCode_impl(HashCode_Data_BeforeVersion)
    End If

    
'    今のVersionがCreateHashで更新されない場合はエラーを出す｡
    If Not IsUpdateRegisterVersion() Then
        Call MsgBox("All HashCode is same. Please check if you move pat files to ""RegVerX"" folder!", vbCritical, "RVMM_CreateRegisterHashCode")
        GoTo ErrorExit
    End If
        
'    シートに記載をする前にソートを行う。Verが1の場合はソートしない
    If VerX <> 1 Then
        '今のVersionを基準に前のVersionのパスを並び替え
         Call SortHashCodeInformation(HashCode_Data, HashCode_Data_BeforeVersion)
    End If
      
'    今のバージョンのHASHCODEと､前のバージョンのHASHCODEをHASHCODEシートに書く
'    シートのクリアも行う
    Call WrtieHashCode(VerX <> 1, HashCode_Data, HashCode_Data_BeforeVersion)
    
'    今のVersionと前のVersionのHASHコードに違いがあったらしるしをつける
    Call CheckHashCoceWithBeforeVersion
      
'    CreatHashしたRegVerを外部ファイルに履歴をのこす
    Call WriteCreateHashCodeRecord(VerX)
        
'    後始末
    Erase HashCode_Data_BeforeVersion
    Call RecoverHashCodeData(HashCode_Data, ArrayMax)

    
    '状態遷移
    m_RvmmState = RVMM_STATE_HASE_CREATED
    
    Call MsgBox("RVMM_CreateRegisterHashCode is succeeded !", , "Congratulation")
    
    Exit Sub
    
ErrorExit:
'    後始末
    Erase HashCode_Data_BeforeVersion
    Call RecoverHashCodeData(HashCode_Data, ArrayMax)
         
End Sub

Public Function RVMM_GetRegisterVersion() As Long
    
    Dim VerX As Long
    
'    HASHCODEシートがなかったら動かない
    If Not IsHashCodeFunctionEnable() Then
        Call OutpuMessageToIgxlDataLog("HashCode Sheet is not found! This function is disable!")
        RVMM_GetRegisterVersion = 0
        Flg_HashCheckResult = False
        Exit Function
    End If
    
'    初期化チェック
    If m_RvmmState = RVMM_STATE_UNINITIALIZED Then
        Call OutpuMessageToIgxlDataLog("Register Version Management Mod is not be initialized!")
        RVMM_GetRegisterVersion = -1
        Flg_HashCheckResult = False
        Exit Function
    End If
    
'    LoadPatで追加された分は消したくないので値を保持する
    Dim ArrayMax As Long
    ArrayMax = GetUBoundHashCode_Data(HashCode_Data)
    
'    LoadPatが呼ばれていることが前提
'    PatGrpシートとTestInstanceから情報収集をする｡
    Call GetHashCode_information

'    パスを全部調べてVersionが全部一緒でないとエラーとする
    If Not IsSamePattenVersion(HashCode_Data, VerX) Then
        Call OutpuMessageToIgxlDataLog("The Register Versions is not same!")
        RVMM_GetRegisterVersion = -2
        Call DisableAllTest 'テストの停止(EeeJob関数)
        GoTo ErrorExit
    End If
    
'    外部ファイルの履歴からすべてのRegVerフォルダが全部HASHCODE変換されたかチェックする
    If Not IsAllRegVerHashCreated Then
        Call OutpuMessageToIgxlDataLog("All ""RegVerX"" folder is not created hashcode!")
        RVMM_GetRegisterVersion = -3
        Call DisableAllTest 'テストの停止(EeeJob関数)
        GoTo ErrorExit
    End If
    
'    パスとパタン名からハッシュコードを生成して､HASHCODEシートのHashコードと比較し､合致しないとエラー
'    今のVersionのHASHCODEを作成する
    Call CreateHashCode_impl(HashCode_Data)

'   パタンファイルがなかった場合の処理
    Dim strNotFoundFile As String
    If Not IsAllFileHashCodeCreated(HashCode_Data, strNotFoundFile) Then
        Call OutpuMessageToIgxlDataLog("Pat file " & strNotFoundFile & " is not found!")
        RVMM_GetRegisterVersion = -4
        Call DisableAllTest 'テストの停止(EeeJob関数)
        GoTo ErrorExit
    End If
    
'    シートと比較する
    If Not IsEqaulToHashCode(HashCode_Data, strNotFoundFile) Then
        If (Len(strNotFoundFile) = 0) Then
            Call OutpuMessageToIgxlDataLog("HashCode Sheet is empty!")
        Else
            Call OutpuMessageToIgxlDataLog("Pat file " & strNotFoundFile & "'s hashcode is mismatch!")
        End If
        RVMM_GetRegisterVersion = -5
        Call DisableAllTest 'テストの停止(EeeJob関数)
        GoTo ErrorExit
    End If
  
    '返り値セット
    RVMM_GetRegisterVersion = VerX
    
'    後始末
    Call RecoverHashCodeData(HashCode_Data, ArrayMax)
    
    '状態遷移
    m_RvmmState = RVMM_STATE_HASE_CREATED
    
    Exit Function
    
ErrorExit:
'    後始末
    Call RecoverHashCodeData(HashCode_Data, ArrayMax)
    First_Exec = 0
    Flg_HashCheckResult = False
End Function

Public Sub RVMM_LoadPat(ByVal PatPath As String)

'    松野さん
'
'    PassからPatNameだけを読み取って､それを構造体にパターンの名前として扱う
'    パターンのロードもする｡ (HashCode_information)

'    ' パターンを読みこむ(FULLパスだったか、ファイル名だったか・・・)
#If 1 Then
    Call TheHdw.Digital.Patterns.pat(PatPath).Load
#End If

'    HASHCODEシートがなかったらこれ以上は動かない
    If Not IsHashCodeFunctionEnable() Then
        Exit Sub
    End If

'    初期化チェック
    If m_RvmmState <> RVMM_STATE_INITIALIZED Then
        Call MsgBox("Register Version Management Mod is not be initialized!", vbCritical, "RVMM_LoadPat")
        Exit Sub
    End If

    'PathからPatNameを読み取る、Loadできたことから"PatPass"はファイルパスだと考えてよい
    Dim i As Integer
    Dim j As Integer
    i = InStrRev(PatPath, "\") + 1
    j = InStrRev(UCase(PatPath), UCase(".pat"))
        
    Dim tempstr As String
    tempstr = Mid(PatPath, i, j - i)


    '構造体にパターンのパスと名前を格納(2回目以降)
On Error GoTo FIRST_CYCLE
    Dim elem_max As Integer
    elem_max = UBound(HashCode_Data) + 1
On Error GoTo 0
    ReDim Preserve HashCode_Data(elem_max) As HashCode_information
    
    HashCode_Data(elem_max).filePath = PatPath
    HashCode_Data(elem_max).PatName = tempstr
        
    Exit Sub
    
    '構造体にパターンのパスと名前を格納(初回)
FIRST_CYCLE:
    ReDim HashCode_Data(0) As HashCode_information
    HashCode_Data(0).filePath = PatPath
    HashCode_Data(0).PatName = tempstr

End Sub

'================================================================================
' Private Functions
'================================================================================

Private Sub OutpuMessageToIgxlDataLog(ByRef strMsg As String)

#If 1 Then
        Call TheExec.Datalog.WriteComment(strMsg)
#Else
        Debug.Print strMsg
#End If

End Sub


Private Function GetUBoundHashCode_Data(ByRef HashCodeArray() As HashCode_information)

    On Error GoTo FirstArrayAlloc
    GetUBoundHashCode_Data = UBound(HashCodeArray)
    GoTo AllocEnd
FirstArrayAlloc:
    GetUBoundHashCode_Data = -1
AllocEnd:
     On Error GoTo 0

End Function

Private Sub RecoverHashCodeData(ByRef HashCodeArray() As HashCode_information, ByVal lRecoverSize As Long)

    If lRecoverSize = -1 Then
        Erase HashCodeArray
    Else
        ReDim Preserve HashCodeArray(lRecoverSize)
    End If

End Sub

Private Function IsAllFileHashCodeCreated(ByRef HashCodeArray() As HashCode_information, ByRef strNotFoundFile As String) As Boolean

    Dim lBegin As Long, lEnd As Long
    
    lBegin = LBound(HashCodeArray)
    lEnd = UBound(HashCodeArray)
    
    Dim i As Long
    
    For i = lBegin To lEnd
        If StrComp(HashCodeArray(i).HashCode, RVVM_FILE_NOT_FOUND_HASHCODE) = 0 Then
            strNotFoundFile = HashCodeArray(i).PatName
            IsAllFileHashCodeCreated = False
            Exit Function
        End If
    Next i

    IsAllFileHashCodeCreated = True

End Function

Private Function GetHashCode_information() As Double

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■テストインスタンスの名前とパスを取ってくる
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Dim ShtTestInstances As Worksheet
    Set ShtTestInstances = ThisWorkbook.Worksheets("Test Instances")

    Dim ArrMax As Long
    Dim j As Long
    Dim i As Long
    
    On Error GoTo FirstArrayAlloc
    ArrMax = UBound(HashCode_Data)
    GoTo AllocEnd
FirstArrayAlloc:
    ArrMax = -1
AllocEnd:
     On Error GoTo 0
   
    j = ArrMax + 1 'GetRegVerへ入れる配列の初期値
    i = 5   'テストインスタンスのCheck開始行を指定（固定）
    
    Do Until Len(Trim(ShtTestInstances.Cells(i, 2))) = 0  'Group Nameが空白セル（ Len(Trim(空白セル)) ）まで行を変えてCheckする
        If ShtTestInstances.Cells(i, 3) = "IG-XL Template" Then
            ReDim Preserve HashCode_Data(j)
            HashCode_Data(j).PatName = ShtTestInstances.Cells(i, 2)
            HashCode_Data(j).filePath = ShtTestInstances.Cells(i, 14)
            j = j + 1
        End If
        i = i + 1
    Loop

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■PatGrpsの名前とパスを取ってくる
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    Dim objsheet_PatGrps As Object
    Dim shtPatGrp As Worksheet
    
    For Each objsheet_PatGrps In Worksheets
        If objsheet_PatGrps.Name = "PatGrps" Then
            Set shtPatGrp = ThisWorkbook.Worksheets("PatGrps")
            i = 4   'PatGrpsのCheck開始行を指定（固定）
            
            Do Until Len(Trim(shtPatGrp.Cells(i, 2))) = 0  'Group Nameが空白セル（ Len(Trim(空白セル)) ）まで行を変えてCheckする
                ReDim Preserve HashCode_Data(j)
                HashCode_Data(j).PatName = shtPatGrp.Cells(i, 2)
                HashCode_Data(j).filePath = shtPatGrp.Cells(i, 3)
                j = j + 1
                i = i + 1
            Loop
        End If
    Next

End Function

'HASHCODEシートがなかったら動かない
'田中さん
Private Function IsHashCodeFunctionEnable() As Boolean

    IsHashCodeFunctionEnable = False
    
    Dim shtHashCode As Worksheet
    
    On Error GoTo errLable

    '======= WorkSheet Select ========
    Set shtHashCode = ThisWorkbook.Sheets("HashCode")
    
    IsHashCodeFunctionEnable = True
    
    Set shtHashCode = Nothing
    
    Exit Function
    
errLable:
    IsHashCodeFunctionEnable = False
End Function

'REGVERフォルダを全部みて､最新のREGVERフォルダでなければエラーとする
Private Function IsLatestRegversion(ByVal VersionX As Long) As Boolean
'高木(真)
Dim strRet As String
Dim strChar As String
Dim strOrg As String
Dim PatFolder() As String
Dim PatPath, FolderName
Dim GetFolderName As Integer
Dim FolderNoA As Integer
Dim FolderNoB As Integer
Dim LatestVer As Integer
Dim NameLength As Integer
Dim LatestFolderNo As Integer

    '使用するディレクトリ指定
    PatPath = ThisWorkbook.Path & "\PAT\"
    'Regverという名前のフォルダの検索(最初のフォルダの値が入る。)
    FolderName = Dir(PatPath & "Regver*", vbDirectory)
    '初期設定
    FolderNoA = 0
    LatestVer = 0
    ReDim PatFolder(0)
    
    'ディレクトリ内のRegverという名前のフォルダーを全てPatFolderに格納。
    'Regverがフォルダであることを確認する。
    Do While FolderName <> ""
        '現在のフォルダと親フォルダは無視。
        If FolderName <> "." And FolderName <> ".." Then
            If (GetAttr(PatPath & FolderName) And vbDirectory) = vbDirectory Then
                ReDim Preserve PatFolder(FolderNoA)
                PatFolder(FolderNoA) = FolderName
                FolderNoA = FolderNoA + 1
            End If
        End If
        FolderName = Dir '次のフォルダ名を返す
    Loop
    
    For FolderNoB = 0 To FolderNoA - 1
        strRet = ""
        strOrg = PatFolder(FolderNoB)
        
        'フォルダ名からVersionとなる数字を抜き出す。
        For NameLength = 1 To Len(strOrg)
            strChar = Mid(strOrg, NameLength, 1)
            If IsNumeric(strChar) Then
                strRet = strRet & strChar
            End If
        Next NameLength
        
        If LatestVer < strRet Then
            LatestVer = strRet
        ElseIf LatestVer = strRet Then
            '同じverは存在しないためエラーとする。
            MsgBox "同じversionが存在しています。"
            Exit Function
        End If
    Next FolderNoB
    
    If LatestVer = VersionX Then
        IsLatestRegversion = True
    End If
    
End Function
    
' パスを全部調べてVersionが全部一緒でないとエラーとする
Private Function IsSamePattenVersion(ByRef HashCodeArray() As HashCode_information, ByRef VerX As Long) As Boolean
    '赤坂
    Const Offset_RegVer As Long = 6   'RegVerXのXを読む為に6(RegVerの文字数)をOffsetとして入力
    
    Dim i As Long, lBegin As Long, lEnd As Long
    
    '配列の大きさ確認
    On Error GoTo NOT_ARRAY_ALLOC
    lBegin = LBound(HashCodeArray)
    lEnd = UBound(HashCodeArray)
    On Error GoTo 0

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■PathからRegVerXのXの値を比較して同じならVerを返す。違っていたら-1を返す。
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Dim RegVer_Posision As Long
    Dim counter As Long
    Dim lArraySize As Long
    lArraySize = 0
    For counter = lBegin To lEnd       '配列個数分比較する。
        Dim VerNumber As Long
        Dim ArrRegVerX() As String
        
        VerNumber = 0   'Xの文字数(初期値0)

        'PathからRegVerの位置を取得します。
        RegVer_Posision = InStr(UCase(HashCodeArray(counter).filePath), UCase("RegVer"))

        If RegVer_Posision <> 0 Then
            'RegVerXのXが何桁かを取得します(aを求める)
            Do While IsNumeric(Mid(HashCodeArray(counter).filePath, RegVer_Posision + Offset_RegVer + VerNumber, 1))
                VerNumber = VerNumber + 1   'もし切り取った文字が数字だったら次の文字を見に行く。
            Loop
        
            'Xの値を求めます。
            ReDim Preserve ArrRegVerX(lArraySize)
            ArrRegVerX(lArraySize) = Mid(HashCodeArray(counter).filePath, RegVer_Posision + Offset_RegVer, VerNumber)
            
            'Xの値を比較します。
            If lArraySize <> 0 Then
                If ArrRegVerX(lArraySize) <> ArrRegVerX(lArraySize - 1) Then
                    IsSamePattenVersion = False
                    VerX = -1
                    Exit Function
                End If
                VerX = ArrRegVerX(lArraySize)
            End If
            lArraySize = lArraySize + 1
        End If
    Next

    'RegVerXのX全てが等しいのでIsSamePattenVersionにTrueを返します。
    IsSamePattenVersion = True
    
     Exit Function
NOT_ARRAY_ALLOC:
    Err.Raise 9999, "Test"
    
End Function
    
Private Sub ConvertBeforeVersionPath(ByRef NowArray() As HashCode_information, ByRef BeforeArray() As HashCode_information)
    '赤坂
    Const Offset_RegVer As Long = 6   'RegVerXのXを読む為に6(RegVerの文字数)をOffsetとして入力
    
    Dim i As Long, lBegin As Long, lEnd As Long
    
    '配列の大きさ確認
    On Error GoTo NOT_ARRAY_ALLOC
    lBegin = LBound(NowArray)
    lEnd = UBound(NowArray)
    On Error GoTo 0

    
    '配列確保
    ReDim BeforeArray(lEnd)
    
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■今のVersionパスから前のVersionのパスを作り出す。
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Dim RegVer_Posision As Long
    Dim counter As Long

    counter = 0     '初期化

    For counter = lBegin To lEnd       '配列個数分比較する。
        Dim VerNumber As Long
        Dim ArrRegVerX() As String
        Dim lX() As Long
        Dim BeforeX() As String
        Dim BeforelX() As Long
        Dim FilePath_No As Long
        Dim FirstPath As String
        Dim LatterPath As String
        Dim FirstFilePath_No As Long
        
       
        VerNumber = 0   'Xの文字数(初期値0)

        'PassからRegVerの位置を取得します。
        RegVer_Posision = InStr(UCase(NowArray(counter).filePath), UCase("RegVer"))

        If RegVer_Posision <> 0 Then
            'RegVerXのXが何桁かを取得します(aを求める)
            Do While IsNumeric(Mid(NowArray(counter).filePath, RegVer_Posision + Offset_RegVer + VerNumber, 1))
                VerNumber = VerNumber + 1   'もし切り取った文字が数字だったら次の文字を見に行く。
            Loop
        
            'Xの値達を求めます。
            ReDim Preserve ArrRegVerX(counter)
            ReDim Preserve lX(counter)
            ReDim Preserve BeforeX(counter)
            ReDim Preserve BeforelX(counter)
            ArrRegVerX(counter) = Mid(NowArray(counter).filePath, RegVer_Posision + Offset_RegVer, VerNumber)
            lX(counter) = val(ArrRegVerX(counter))   '文字列を数値に変換
            
            'Xの値を-1します。（Verを一つ下げます。）
            BeforelX(counter) = lX(counter) - 1
            BeforeX(counter) = str(BeforelX(counter))
            
            'Pathを合成する為に全体の文字数を数えます。
            FilePath_No = Len(NowArray(counter).filePath)
            
            'RegVerまでのPathと文字数を取得します。
            FirstPath = Left(NowArray(counter).filePath, RegVer_Posision + Offset_RegVer - 1)
            FirstFilePath_No = Len(FirstPath)
            
            'RegVerX以降のPathを取得します。
            LatterPath = Right(NowArray(counter).filePath, FilePath_No - FirstFilePath_No - VerNumber)
            
            '前のVersionのパスを作り出します。
            BeforeArray(counter).filePath = FirstPath & LTrim(BeforeX(counter)) & LatterPath
            BeforeArray(counter).PatName = NowArray(counter).PatName
        Else
            BeforeArray(counter).filePath = RVVM_DONT_MANAGEMENT_VERSION
            BeforeArray(counter).PatName = NowArray(counter).PatName
        End If
    Next
    
     Exit Sub
NOT_ARRAY_ALLOC:
    Err.Raise 9999, "Test"
End Sub

'今のVersionのHASHCODEを作成する
'前のVersionのHASHCODEを作成する
Private Sub CreateHashCode_impl(ByRef HashCodeArray() As HashCode_information)
    '丸山
    Const MAX_SIZE As Long = 10# * 1024 * 1024
   
    Dim i As Long, lBegin As Long, lEnd As Long
    
    '配列の大きさ確認
    On Error GoTo NOT_ARRAY_ALLOC
    lBegin = LBound(HashCodeArray)
    lEnd = UBound(HashCodeArray)
    On Error GoTo 0
    
    For i = lBegin To lEnd
        
        'ファイルのサイズを取得
        Dim File_Size As Long
    
        If StrComp(HashCodeArray(i).filePath, RVVM_DONT_MANAGEMENT_VERSION) = 0 Then 'バージョン管理しないファイルの場合
            File_Size = -2 'バージョン管理しないときは-1とする
        ElseIf Dir(HashCodeArray(i).filePath) <> "" Then 'ファイルの存在確認
            File_Size = FileLen(HashCodeArray(i).filePath)
        Else
            File_Size = -1  'ファイルの存在しないときは-1とする
        End If

        If (File_Size = 0) Then
            HashCodeArray(i).HashCode = ""
        ElseIf (File_Size = -1) Then
            HashCodeArray(i).HashCode = RVVM_FILE_NOT_FOUND_HASHCODE
        ElseIf (File_Size = -2) Then
            HashCodeArray(i).HashCode = RVVM_DONT_MANAGEMENT_VERSION
        ElseIf (File_Size > MAX_SIZE) Then
            HashCodeArray(i).HashCode = RVVM_LARGE_FILE_SIZE_HASHCODE
        Else
            HashCodeArray(i).HashCode = CreateHashFile(HashCodeArray(i).filePath, CALG_MD5)
        End If
'        HashCodeArray(i).HashCode = CreateMD5HashString(HashCodeArray(i).FilePath)
    Next i
    
    Exit Sub
    
NOT_ARRAY_ALLOC:
    Err.Raise 9999, "CreateHashCode_impl", "Memory Allocation Error!"
End Sub
    
'今のVersionがCreateHashで更新されない場合はエラーを出す｡
Private Function IsUpdateRegisterVersion() As Boolean
    '丸山
    
    IsUpdateRegisterVersion = False
    
    'シートの取得
    Dim shtHashCode As Worksheet
    Set shtHashCode = ThisWorkbook.Worksheets("HashCode")
    
    'HashCode_Data配列の大きさ確認
    Dim lBegin As Long, lEnd As Long
    On Error GoTo NOT_ARRAY_ALLOC
    lBegin = LBound(HashCode_Data)
    lEnd = UBound(HashCode_Data)
    On Error GoTo 0
    
    
    '===大方針としてはデバッグ時追加のことを考えて、HashCode_Dataを基準にシートを確認していく====
    'まずシートから情報の取得、空っぽの場合はVersionUpがなされるとしてすぐぬける
    Dim i As Long, j As Long
    Dim shtDataArray() As HashCode_information
    
    i = 4
    j = 0
    If Len(Trim(shtHashCode.Cells(i, 2))) = 0 Then
        Set shtHashCode = Nothing 'シートの開放
        IsUpdateRegisterVersion = True
        Exit Function
    End If
    
    Do Until Len(Trim(shtHashCode.Cells(i, 2))) = 0  'Group Nameが空白セル（ Len(Trim(空白セル)) ）まで行を変えてCheckする
        ReDim Preserve shtDataArray(j)
        With shtDataArray(j)
            .PatName = shtHashCode.Cells(i, 2)
            .HashCode = shtHashCode.Cells(i, 3)
        End With
        j = j + 1
        i = i + 1
    Loop
    
    Set shtHashCode = Nothing 'シートの開放
    
    'shtDataArray配列の大きさ確認
    Dim lBeginSht As Long, lEndSht As Long
    On Error GoTo NOT_ARRAY_ALLOC
    lBeginSht = LBound(shtDataArray)
    lEndSht = UBound(shtDataArray)
    On Error GoTo 0
    
    
    '次にHashCode_Dataを基準にシートの情報を確認していく
    Dim IsMatch As Boolean
    j = 0
    For i = lBegin To lEnd
        IsMatch = False 'いったんFalseにして
        For j = lBeginSht To lEndSht
            If (StrComp(HashCode_Data(i).PatName, shtDataArray(j).PatName) = 0) Then
                If (StrComp(HashCode_Data(i).HashCode, RVVM_LARGE_FILE_SIZE_HASHCODE) = 0) And _
                    (StrComp(shtDataArray(i).HashCode, RVVM_LARGE_FILE_SIZE_HASHCODE) = 0) Then
                    IsMatch = False 'PatNameが一致してファイルサイズが大きい場合Falseとする
                    Exit For
                ElseIf (StrComp(HashCode_Data(i).HashCode, shtDataArray(j).HashCode) = 0) Then
                    IsMatch = True 'PatNameとハッシュコードが一致する場合Trueとする
                    Exit For
                Else
                    IsMatch = False 'PatNameが一致してハッシュコードが異なる場合Falseとする
                    Exit For
                End If
            End If
            
        Next j
        
        'ここでIsMathch=Falseの場合は
        '・HashCode_Dataの情報がシートにみつからなかった
        '・HashCode_Dataの情報がシートと異なっていた
        'ことを意味するのでバージョンアップがなされたとして抜けてよい
        If Not IsMatch Then
            Erase shtDataArray
            IsUpdateRegisterVersion = True
            Exit Function
        End If
    Next i
    
    IsUpdateRegisterVersion = False
    
    Exit Function
    
NOT_ARRAY_ALLOC:
    Err.Raise 9999, "IsUpdateRegisterVersion", "Memory Allocation Error!"
    
End Function

        
'今のバージョンのHASHCODEと､前のバージョンのHASHCODEをHASHCODEシートに書く
'中でソートすること
Private Sub WrtieHashCode(ByVal IsWriteBeforeVersion As Boolean, _
        ByRef NowArray() As HashCode_information, ByRef BeforeArray() As HashCode_information)
    
    '丸山
    'エラー処理はどこまでしようか？ひとまず何もしないでおく
    
    Const COLUMN_PAT_NAME As Long = 2
    Const COLUMN_PAT_NOW_HASH As Long = 3
    Const COLUMN_PAT_BEFORE_HASH As Long = 4
    
    Const ROW_START As Long = 4
    
    Dim shtHashCode As Worksheet
    Set shtHashCode = ThisWorkbook.Worksheets("HashCode")
    
    '今あるもののクリア
    Call ClearWorkShet(shtHashCode)
        
    '要素数の取得、並び替え後なので、数もそろっているはず
    Dim i As Long, lBegin As Long, lEnd As Long
    lBegin = LBound(NowArray)
    lEnd = UBound(NowArray)
    
    'ライトする
    If IsWriteBeforeVersion Then
        For i = lBegin To lEnd
            With shtHashCode
                .Cells(i + ROW_START, COLUMN_PAT_NAME) = NowArray(i).PatName
                .Cells(i + ROW_START, COLUMN_PAT_NOW_HASH) = NowArray(i).HashCode
                .Cells(i + ROW_START, COLUMN_PAT_BEFORE_HASH) = BeforeArray(i).HashCode
            End With
        Next i
    Else
        For i = lBegin To lEnd
            With shtHashCode
                .Cells(i + ROW_START, COLUMN_PAT_NAME) = NowArray(i).PatName
                .Cells(i + ROW_START, COLUMN_PAT_NOW_HASH) = NowArray(i).HashCode
            End With
        Next i
    End If
    
    Set shtHashCode = Nothing 'シートの開放
    
End Sub
    
'今のVersionと前のVersionのHASHコードに違いがあったらしるしをつける
Private Sub CheckHashCoceWithBeforeVersion()

    '田中さん
    Const COLUMN_PAT_NAME As Long = 2
    Const COLUMN_PAT_NOW_HASH As Long = 3
    Const COLUMN_PAT_BEFORE_HASH As Long = 4
    Const COLUMN_PAT_DIFF As Long = 5
    Const ROW_START As Long = 4
    
    Dim shtHashCode As Worksheet
    Set shtHashCode = ThisWorkbook.Worksheets("HashCode")
    
    Dim i As Long
    Dim strNowHash As String
    Dim strBeforeHash As String
    
    Do Until Len(Trim(shtHashCode.Cells(i + ROW_START, COLUMN_PAT_NAME))) = 0 'Group Nameが空白セル（ Len(Trim(空白セル)) ）まで行を変えてCheckする
        strNowHash = shtHashCode.Cells(i + ROW_START, COLUMN_PAT_NOW_HASH)
        strBeforeHash = shtHashCode.Cells(i + ROW_START, COLUMN_PAT_BEFORE_HASH)
        
        If StrComp(RVVM_LARGE_FILE_SIZE_HASHCODE, strNowHash) = 0 Or StrComp(RVVM_LARGE_FILE_SIZE_HASHCODE, strBeforeHash) = 0 Then
            shtHashCode.Cells(i + ROW_START, COLUMN_PAT_DIFF) = "*"
        ElseIf StrComp(RVVM_DONT_MANAGEMENT_VERSION, strBeforeHash) = 0 Then
            shtHashCode.Cells(i + ROW_START, COLUMN_PAT_DIFF) = "+"
        ElseIf StrComp(strNowHash, strBeforeHash) <> 0 Then
            shtHashCode.Cells(i + ROW_START, COLUMN_PAT_DIFF) = "O"
        End If
        i = i + 1
    Loop

End Sub
  
Private Sub WriteCreateHashCodeRecord(ByRef VerX As Long)
    '赤坂
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■CreatHashしたRegVerを外部ファイルに履歴をのこす。
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    Dim BookName As String
    Dim TypeName As String
    Dim DerivationName As String
    Dim RecordFullPathName As String
    Dim intFileNum As Integer
    Dim RecordDate As Date

    Const StartTypeNamePosition As String = 4   'Job名のTypeが書いてある開始位置(Ex.p7e104lq4なら4文字目から)
    Const TypeNameNumber As String = 3          'Typeの文字数（083,104等）
    Const StartDerivationNamePosition As String = 7   'Job名の派生が書いてある開始位置(Ex.p7e104lq4なら7文字目から)
    Const DerivationNameNumber As String = 2          '派生の文字数（lq,cq,aq等）
    

    'Typeを取得(ExcelのJobNameから取得)　p7e104lq4
    BookName = ThisWorkbook.Name
    TypeName = Mid(BookName, StartTypeNamePosition, TypeNameNumber)
    
    If Not IsNumeric(TypeName) Then
        Call MsgBox("Pleas Check Job TypeName (ExcelFileName)")     'もし切り取った文字が数字で無かったらエラーメッセージを表示。
        Err.Raise 9999, "Test"
    End If
    
    '派生を取得
    DerivationName = Mid(BookName, StartDerivationNamePosition, DerivationNameNumber)

    '日付を取得
    RecordDate = Date

    '履歴があるかどうか確認(まず、JobのFilePathを取得)
    RecordFullPathName = ThisWorkbook.Path & "\PAT\" & "HashCodeRecorde" & "_" & TypeName & "_" & DerivationName & ".txt"

    '履歴の作成
    If Dir(RecordFullPathName) = "" Then   '無ければ新規作成
        intFileNum = FreeFile
        Open RecordFullPathName For Output As intFileNum
        Print #intFileNum, RecordDate & " " & "RegVer" & LTrim(str(VerX))
        Close #intFileNum
    Else                                    '既に履歴があれば上書き
        intFileNum = FreeFile
        Open RecordFullPathName For Append As intFileNum
        Print #intFileNum, RecordDate & " " & "RegVer" & LTrim(str(VerX))
        Close #intFileNum
    End If
    
End Sub
    
Private Function IsAllRegVerHashCreated() As Boolean
    '赤坂
Dim PatPath As String
Dim FolderName As String
Dim FolderNo As Integer

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■外部ファイルの履歴からすべてのRegVerフォルダが全部HASHCODE変換されたかチェックする
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■


    '■■■■■フォルダからRegVerを取得する■■■■■
    '使用するディレクトリ指定
    PatPath = ThisWorkbook.Path & "\PAT\"
    'Regverという名前のフォルダの検索(最初のフォルダ)
    FolderName = Dir(PatPath & "Regver*", vbDirectory)
    '初期設定
    FolderNo = 0
    ReDim ExistenceFolder(0)
    
    '  ディレクトリ内のRegverという名前のフォルダーを全てExistenceFolderに格納。
    'Regverがフォルダであることを確認する。
    Do While FolderName <> ""
        '現在のフォルダと親フォルダは無視。
        If FolderName <> "." And FolderName <> ".." Then
            If (GetAttr(PatPath & FolderName) And vbDirectory) = vbDirectory Then
                'ExistenceFolderのサイズに合わせて配列の大きさを変える。
                ReDim Preserve ExistenceFolder(FolderNo)
                ExistenceFolder(FolderNo) = FolderName
                FolderNo = FolderNo + 1
            End If
        End If
        FolderName = Dir '次のフォルダ名を返す
    Loop
    
    
    '■■■■■履歴が有るかを確認する■■■■■
    Dim BookName As String
    Dim TypeName As String
    Dim DerivationName As String
    Dim RecordFullPathName As String
    Dim intFileNum As Integer

    Const StartTypeNamePosition As String = 4   'Job名のTypeが書いてある開始位置(Ex.p7e104lq4なら4文字目から)
    Const TypeNameNumber As String = 3          'Typeの文字数（083,104等）
    Const StartDerivationNamePosition As String = 7   'Job名の派生が書いてある開始位置(Ex.p7e104lq4なら7文字目から)
    Const DerivationNameNumber As String = 2          '派生の文字数（lq,cq,aq等）

    'Typeを取得(ExcelのJobNameから取得)　p7e104lq4
    BookName = ThisWorkbook.Name
    TypeName = Mid(BookName, StartTypeNamePosition, TypeNameNumber)
    
    If Not IsNumeric(TypeName) Then
        Call MsgBox("Pleas Check Job TypeName (ExcelFileName)")     'もし切り取った文字が数字で無かったらエラーメッセージを表示。
        Err.Raise 9999, "Test"
    End If
    
    '派生を取得
    DerivationName = Mid(BookName, StartDerivationNamePosition, DerivationNameNumber)

    '履歴があるかどうか確認(まず、JobのFilePathから取得)
    RecordFullPathName = ThisWorkbook.Path & "\PAT\" & "HashCodeRecorde" & "_" & TypeName & "_" & DerivationName & ".txt"
    
    
    If Dir(RecordFullPathName) <> "" Then   '履歴がある場合のみ比較する
        '■■■■■txtからRegVerを取得する■■■■■
        
        Const Offset_RegVer As Long = 6   'RegVerXのXを読む為に6(RegVerの文字数)をOffsetとして入力
    
        Dim LineDate As String
        Dim VerNumber As Long
        Dim RecodeDate() As String
        Dim lCounter As Long
        Dim RegVer_Posision As Long

        lCounter = 0 '初期化
        
        intFileNum = FreeFile
        
        Open RecordFullPathName For Input As intFileNum
        While Not EOF(intFileNum)
            Line Input #intFileNum, LineDate
                '読み込んだデータからRegVerの位置を取得します。
                RegVer_Posision = InStr(1, LineDate, "RegVer", vbBinaryCompare)
        
                'RegVerXのXが何桁かを取得します(aを求める)
                VerNumber = 0   'Xの文字数(初期値0)
                Do While IsNumeric(Mid(LineDate, RegVer_Posision + Offset_RegVer + VerNumber, 1))
                    VerNumber = VerNumber + 1   'もし切り取った文字が数字だったら次の文字を見に行く。
                Loop
            
                '読み込んだデータのRegVerXを求めます。
                ReDim Preserve RecodeDate(lCounter)
                RecodeDate(lCounter) = Mid(LineDate, RegVer_Posision, VerNumber + Offset_RegVer)
                lCounter = lCounter + 1
        Wend
        
        Close intFileNum
        
        '■■■■■フォルダとテキストのRegVerを比較する■■■■■
        Dim Loop1 As Long
        Dim Loop2 As Long
        Dim IsFolderCheck As Boolean
        
        FolderNo = FolderNo - 1
        lCounter = lCounter - 1
        
        For Loop1 = 0 To FolderNo
            IsFolderCheck = False
            For Loop2 = 0 To lCounter
                If StrComp(UCase(ExistenceFolder(Loop1)), UCase(RecodeDate(Loop2))) = 0 Then
                    IsFolderCheck = True
                    Exit For
                End If
            Next
            If IsFolderCheck = False Then
                IsAllRegVerHashCreated = False
                Exit Function
            End If
        Next
    Else
        IsAllRegVerHashCreated = False
        Exit Function
    End If
    IsAllRegVerHashCreated = True
End Function

'シートと比較する
Private Function IsEqaulToHashCode(ByRef HashCodeArray() As HashCode_information, ByRef strPatName As String) As Boolean
    '田中さん
    
    'HashCode_Data配列の大きさ確認
    Dim lBegin As Long, lEnd As Long
    On Error GoTo NOT_ARRAY_ALLOC
    lBegin = LBound(HashCodeArray)
    lEnd = UBound(HashCodeArray)
    On Error GoTo 0
    
    'ワークシートの取得
    Dim shtHashCode As Worksheet
    Set shtHashCode = ThisWorkbook.Worksheets("HashCode")
     
    
    '===読み込まれたパタンファイルがHashCodeシートのハッシュコードと等しいことを確認したい====
    '===HashCodeシートに余分なものが書いてあることに関してはOKとする====
    'まずシートから情報の取得、空っぽの場合はVersionUpがなされるとしてすぐぬける
    Dim i As Long, j As Long
    Dim shtDataArray() As HashCode_information
    i = 4
    j = 0
    If Len(Trim(shtHashCode.Cells(i, 2))) = 0 Then
        Set shtHashCode = Nothing 'シートの開放
        IsEqaulToHashCode = False
        strPatName = ""
        Exit Function
    End If
    
    Do Until Len(Trim(shtHashCode.Cells(i, 2))) = 0  'Group Nameが空白セル（ Len(Trim(空白セル)) ）まで行を変えてCheckする
        ReDim Preserve shtDataArray(j)
        With shtDataArray(j)
            .PatName = shtHashCode.Cells(i, 2)
            .HashCode = shtHashCode.Cells(i, 3)
        End With
        j = j + 1
        i = i + 1
    Loop
    
     
    'shtDataArray配列の大きさ確認
    Dim lBeginSht As Long, lEndSht As Long
    On Error GoTo NOT_ARRAY_ALLOC
    lBeginSht = LBound(shtDataArray)
    lEndSht = UBound(shtDataArray)
    On Error GoTo 0
     
     
     '次にHashCode_Dataを基準にシートの情報を確認していく
    Dim IsMatch As Boolean
    j = 0
    For i = lBegin To lEnd
        IsMatch = False 'いったんFalseにして
        For j = lBeginSht To lEndSht
            If (StrComp(HashCodeArray(i).PatName, shtDataArray(j).PatName) = 0) Then
                If (StrComp(HashCodeArray(i).HashCode, shtDataArray(j).HashCode) = 0) Then
                    IsMatch = True 'PatNameとハッシュコードが一致する場合Trueとする
                    Exit For
                Else
                    IsMatch = False 'PatNameが一致してとハッシュコードが異なる場合Falseとする
                    Exit For
                End If
            End If
            
        Next j
        
        'ここでIsMathch=Falseの場合は
        '・HashCode_Dataの情報がシートにみつからなかった
        '・HashCode_Dataの情報がシートと異なっていた
        'ことを意味するので抜けてよい
        If Not IsMatch Then
            Erase shtDataArray
            IsEqaulToHashCode = False
            strPatName = HashCodeArray(i).PatName
            Exit Function
        End If
    Next i
   
   'ここまできたら完全一致をプレゼント
   IsEqaulToHashCode = True
   
   Exit Function
   
NOT_ARRAY_ALLOC:
    Err.Raise 9999, "IsEqaulToHashCode", "Memory Allocation Error!"
    
End Function

Private Sub SortHashCodeInformation(ByRef NowArray() As HashCode_information, ByRef BeforeArray() As HashCode_information)
'高木(真)
    Dim BuffHensu() As HashCode_information    '値をスワップするための作業域
  
    Dim lngBaseNumber As Long        '中央の要素番号を格納する変数
    Dim iLoopNow  As Long            'ループカウンタ(現verのHushCode_data)
    Dim iLoopBefore  As Long         'ループカウンタ(前verのHushCode_data)
    Dim lngEnd As Long
    
    
    lngEnd = UBound(HashCode_Data)   'ループカウンタ(現verのHushCode_data)の終了数
    ReDim BuffHensu(UBound(HashCode_Data))

    '現PatNameと前PatNameが同じであれば、一時的にBuffHensuに移動。
    For iLoopNow = 0 To lngEnd
        For iLoopBefore = 0 To lngEnd
            If NowArray(iLoopNow).PatName = BeforeArray(iLoopBefore).PatName Then
                BuffHensu(iLoopNow).filePath = BeforeArray(iLoopBefore).filePath
                BuffHensu(iLoopNow).HashCode = BeforeArray(iLoopBefore).HashCode
                BuffHensu(iLoopNow).PatName = BeforeArray(iLoopBefore).PatName
                Exit For
            End If
        Next iLoopBefore
    Next iLoopNow

    'BuffHensuにあるデータを前versionのデータとして移動。
    For iLoopNow = 0 To lngEnd
    BeforeArray(iLoopNow) = BuffHensu(iLoopNow)
    Next iLoopNow

End Sub

'================================================================================
' Function Level2
'================================================================================
Private Sub ClearWorkShet(ByRef sht As Worksheet)
'内容:
'   前出力したやつをクリアする
'
'[CndChk_wkst]    IN   Worksheet:    けすワークシート
'
'注意事項:
'
    Const COLUMN_PAT_NAME As Long = 2
    Const ROW_START As Long = 4
    
    '描画をきる
    Application.ScreenUpdating = False
        
    '最後のセルを取得
    Dim rgLast As Range
    Set rgLast = sht.Cells.SpecialCells(xlCellTypeLastCell)

    '対象領域を選択して、ありとあらゆるものをクリア
    With sht.Range(sht.Cells(ROW_START, COLUMN_PAT_NAME), rgLast)
        .ClearContents
        .Interior.ColorIndex = 0
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
            
    '描画をもどす
    Application.ScreenUpdating = True
    
    '後始末
    Set rgLast = Nothing
    
End Sub

'================================================================================
' Hash code Produce Functions
'================================================================================
Private Function CreateHashFile(ByVal strFileName As String, ByVal lngAlgID As Long) As String
    Dim abytData() As Byte
    Dim intFile As Integer
    Dim lngError As Long
    On Error Resume Next
        If Len(Dir(strFileName)) > 0 Then
            intFile = FreeFile
            Open strFileName For Binary Access Read Shared As #intFile
            abytData() = InputB(LOF(intFile), #intFile)
            Close #intFile
        End If
        lngError = Err.Number
    On Error GoTo 0
    If lngError = 0 Then CreateHashFile = CreateHashFromBinary(abytData(), lngAlgID) _
                    Else CreateHashFile = ""
End Function

   
' Create Hash
Private Static Function CreateHashFromBinary(abytData() As Byte, ByVal lngAlgID As Long) As String
    Dim hProv As Long, hHash As Long
    Dim abytHash(0 To 63) As Byte
    Dim lngLength As Long
    Dim lngResult As Long
    Dim strHash As String
    Dim i As Long
    strHash = ""
    If CryptAcquireContext(hProv, vbNullString, vbNullString, _
                           IIf(lngAlgID >= CALG_SHA_256, PROV_RSA_AES, PROV_RSA_FULL), _
                           CRYPT_VERIFYCONTEXT) <> 0& Then
        If CryptCreateHash(hProv, lngAlgID, 0&, 0&, hHash) <> 0& Then
            lngLength = UBound(abytData()) - LBound(abytData()) + 1
            If lngLength > 0 Then lngResult = CryptHashData(hHash, abytData(LBound(abytData())), lngLength, 0&) _
                             Else lngResult = CryptHashData(hHash, ByVal 0&, 0&, 0&)
            If lngResult <> 0& Then
                lngLength = UBound(abytHash()) - LBound(abytHash()) + 1
                If CryptGetHashParam(hHash, HP_HASHVAL, abytHash(LBound(abytHash())), lngLength, 0&) <> 0& Then
                    For i = 0 To lngLength - 1
                        strHash = strHash & Right$("0" & Hex$(abytHash(LBound(abytHash()) + i)), 2)
                    Next
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0&
    End If
    CreateHashFromBinary = LCase$(strHash)
End Function

' MD5
Public Static Function CreateMD5Hash(abytData() As Byte) As String
    CreateMD5Hash = CreateHashFromBinary(abytData(), CALG_MD5)
End Function

Public Static Function CreateMD5HashString(ByVal strData As String) As String
    CreateMD5HashString = CreateHashString(strData, CALG_MD5)
End Function
' Create Hash From String(Shift_JIS)
Private Static Function CreateHashString(ByVal strData As String, ByVal lngAlgID As Long) As String
    CreateHashString = CreateHashFromBinary(StrConv(strData, vbFromUnicode), lngAlgID)
End Function

