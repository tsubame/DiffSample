VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgKernelManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   Kernel情報の管理用
'目的:
'   Kernelの登録情報問い合わせ用。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日

Option Explicit

Private Const ERR_NUMBER = 9999                         ' Error番号
Private Const CLASS_NAME = "CImgKernelManager"   ' Class名称

Private Const SHT_LABEL_KERNEL_NAME As String = "Kernel Name"
Private Const SHT_LABEL_KERNEL_WIDTH As String = "Width"
Private Const SHT_LABEL_KERNEL_HEIGHT As String = "Height"
Private Const SHT_LABEL_KERNEL_ANCHOR_X As String = "X Anchor"
Private Const SHT_LABEL_KERNEL_ANCHOR_Y As String = "Y Anchor"
Private Const SHT_LABEL_KERNEL_SHIFTR As String = "ShiftR"
Private Const SHT_LABEL_KERNEL_TYPE As String = "Type"
Private Const SHT_LABEL_KERNEL_COMMENT As String = "Comment"
Private Const SHT_LABEL_VAL As String = "Val "
Private Const SHT_LABEL_VAL_ARG_START As Integer = 1
Private Const SHT_LABEL_VAL_ARG_END As Integer = 64

Private mKernelInfos As Collection
Private mTmpKernelInfos As Collection   'カーネル情報
Private mKernelErr As String                'EeeJOBチェックでのエラー記録用
Private mFlagKernelError As Boolean         'IG-XLエラーが発生したかどうかの確認用。EeeJOBとIG-XLで情報の差分発生防止。

Public Sub Init()
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月10日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'
'

    Call RegKernelShtInfo

End Sub

Public Property Get IsErrIGXL() As Boolean
'内容:
'   TheIDPを初期化すべきか否かの判断材料となる情報
'   このフラグがTrueのときはTheIDP.Removeresoucesする必要があります
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月10日
'パラメータ:
'   なし
'戻り値:
'   Boolean
'
'注意事項:
'
'
    IsErrIGXL = mFlagKernelError

End Property

Public Sub CreateKernel( _
    pKernelName As String, ByVal pKernelType As IdpKernelType, _
    ByVal pNumCol As Long, ByVal pNumRow As Long, ByVal pShiftR As Long, ByVal pKernelVal As String, Optional ByVal pComment As String = "" _
)
'内容:
'   カーネル定義を保持します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   [pKernelName]   In  1):カーネル名
'   [pKernelType]   In  2):カーネルタイプ
'   [pNumCol]       In  3):横方向の数
'   [pNumRow]       In  4):縦方向の数
'   [pShiftR]       In  5):右シフトbit数
'   [pKernelVal]    In  6):カーネル定義の文字列
'戻り値:
'
'注意事項:
'

    On Error GoTo ERROR_IGXL
    
    Call TheHdw.IDP.CreateKernel(pKernelName, pKernelType, pNumCol, pNumRow, pShiftR, pKernelVal)

    'テーブルにデータ追加
    Dim newKernel As CImgKernelInfo
    Set newKernel = New CImgKernelInfo
    Call newKernel.Create(pKernelName, pKernelType, pNumCol, pNumRow, pShiftR, pKernelVal, pComment)
    Call mKernelInfos.Add(newKernel, pKernelName)
    
    Exit Sub
    
ERROR_IGXL:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "CreateKernel Error")


End Sub

Public Function IsExist(pKernelName As String) As Boolean
'内容:
'   対象のKernelの有無を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   [pKernelName]  In  1):
'戻り値:
'   Boolean
'
'注意事項:
'
    On Error GoTo NOT_FOUND
    Call mKernelInfos.Item(pKernelName)
    IsExist = True
    Exit Function
NOT_FOUND:
    IsExist = False
End Function

Public Function Kernel(pKernelName As String) As CImgKernelInfo
'内容:
'   対象のKernelのインスタンスを返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年12月2日
'パラメータ:
'   [pKernelName]  In  1):Kernel名
'戻り値:
'   CImgKernelInfo
'
'注意事項:
'   CImgIDPにも同様のメソッドKernelがあります。
'   旧バージョンとの互換確保のために残してます。

    On Error GoTo NOT_FOUND
    Set Kernel = mKernelInfos.Item(pKernelName)
    Exit Function
NOT_FOUND:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, "This Kernel Name is not found!" & vbCrLf & "Kernel Name:" & pKernelName)
End Function

Public Property Get Count() As Long
'内容:
'   このクラスが持つKernelの数を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年12月2日
'パラメータ:
'   なし
'戻り値:
'   Long
'
'注意事項:

    Count = mKernelInfos.Count

End Property

Private Sub RegKernelShtInfo()
'内容:
'   kernelシートがあれば、情報を貰って登録する
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月3日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'

    '#### グループ ####
    Dim pGroupErrCnt As Integer
    Dim pGroupStart As Integer
    Dim pGroupEnd As Integer
    Dim pGroupInfo As Collection
    Dim pTmp As Variant
    
    Set pGroupInfo = New Collection
    
    '#### 行のデータ確認カウンタ ####
    Dim intStartRow As Integer
    Dim intDataCnt As Integer
    Dim intGroupRowCnt As Integer
    
    Dim pAllKernelErrCnt As Long
    Dim pKernelErrCnt As Long
    
    '#### Kernelのパラメータ ####
    Dim pName As String
    Dim pWidth As Integer
    Dim pHeight As Integer
    Dim pShiftR As Integer
    Dim pKernelType As IdpKernelType
    Dim pData As String
    Dim pComment As String
    
    Dim pNameForChk As Collection
    
    Dim pNowHeight As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim pStr As String
    
    '#####  SheetReaderを利用して、Kernelシートを読み込む  #####
    On Error GoTo NOTHING_SHT       'シート読み込みエラーの場合(シートが無い場合何もしない)
    Dim clsWrkShtRdr As CWorkSheetReader
    Set clsWrkShtRdr = GetWkShtReaderManagerInstance.GetReaderInstance(eSheetType.shtTypeKernel)
    On Error GoTo 0
    
    Dim IFileStream As IFileStream
    Dim IParamReader As IParameterReader

    With clsWrkShtRdr
        Set IFileStream = .AsIFileStream
        Set IParamReader = .AsIParameterReader
    End With
    
    mKernelErr = "Image Kernel Definition Sheet Error!" & vbCrLf
    
    '#####  Kernelシートの読み込み  #####
    Set pNameForChk = New Collection
    Do While Not IFileStream.IsEOR

        '#####  カーネル名称を取得  #####
        If IParamReader.ReadAsString(SHT_LABEL_KERNEL_NAME) <> "" Then
            '##### 前行までの定義終端処理(ちゃんと定義された情報のみ登録。エラーは最後にまとめて出す) #####
            If pName <> "" Then         '最初の定義以外
                If pNowHeight <> pHeight Then   'データ高さチェック
                    pKernelErrCnt = pKernelErrCnt + 1
                End If
                If pKernelErrCnt = 0 Then
                    pData = LTrim(pData)        '最初のスペース文字削除
                    Call StockKernelDefinition(pName, pKernelType, pWidth, pHeight, pShiftR, pData, pComment)
                Else
                    Call StockErr(pName)
                End If
            End If
            
            '#####  カーネルパラメータを取得  #####
            pName = IParamReader.ReadAsString(SHT_LABEL_KERNEL_NAME)
            pWidth = IParamReader.ReadAsInteger(SHT_LABEL_KERNEL_WIDTH)
            pHeight = IParamReader.ReadAsInteger(SHT_LABEL_KERNEL_HEIGHT)
            pShiftR = IParamReader.ReadAsInteger(SHT_LABEL_KERNEL_SHIFTR)
            pKernelType = CIdpKernel(IParamReader.ReadAsString(SHT_LABEL_KERNEL_TYPE))
            pComment = IParamReader.ReadAsString(SHT_LABEL_KERNEL_COMMENT)
            
            pData = ""
            pAllKernelErrCnt = pAllKernelErrCnt + pKernelErrCnt
            pKernelErrCnt = 0
            pNowHeight = 0
            
            If IsKey(pName, pNameForChk) = True Then
                pKernelErrCnt = 1
            Else
                Call pNameForChk.Add(pName, pName)
            End If
            If ChkSize(pWidth) = False Or _
               ChkSize(pHeight) = False Or _
               ChkShiftR(pShiftR) = False Or _
               pKernelType = -1 Then
                pKernelErrCnt = 1
            End If
        
        End If
        
        '#####  カーネル定義データを取得  #####
        For i = SHT_LABEL_VAL_ARG_START To SHT_LABEL_VAL_ARG_END
            pStr = IParamReader.ReadAsString(SHT_LABEL_VAL & i)
            If pStr = "" Then
                If i <= pWidth Then
                    '定義データの幅がWidth設定以下
                    pKernelErrCnt = pKernelErrCnt + 1
                End If
            Else
                If i > pWidth Then
                    '定義データの幅がWidth設定よりも大きい
                    pKernelErrCnt = pKernelErrCnt + 1
                Else
                    pData = pData & " " & pStr
                End If
            End If
        Next i
        
        pNowHeight = pNowHeight + 1

        IFileStream.MoveNext
    Loop

    '##### 前行までの定義終端処理(ちゃんと定義された情報のみ登録。エラーは最後にまとめて出す) #####
    If pName <> "" Then         '最初の定義以外
        If pNowHeight <> pHeight Then   'データ高さチェック
            pKernelErrCnt = pKernelErrCnt + 1
        End If
        If pKernelErrCnt = 0 Then
            pData = LTrim(pData)        '最初のスペース文字削除
            Call StockKernelDefinition(pName, pKernelType, pWidth, pHeight, pShiftR, pData, pComment)
        Else
            Call StockErr(pName)
        End If
    End If
    pAllKernelErrCnt = pAllKernelErrCnt + pKernelErrCnt
    
    'シートの整形
    Call ControlShtFormatKernel
    
    Set clsWrkShtRdr = Nothing
    Set IFileStream = Nothing
    Set IParamReader = Nothing
    Set pGroupInfo = Nothing
    
    If pAllKernelErrCnt > 0 Then
'        MsgBox mKernelErr
        Set mTmpKernelInfos = New Collection
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME, mKernelErr)
    End If
        Call TriggerRegDifinition
    Exit Sub

NOTHING_SHT:
    
End Sub

Private Sub StockErr(ByVal pName As String)
'内容:
'   エラーをためておく
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月4日
'パラメータ:
'   [pName] In/Out  1):
'戻り値:
'
'注意事項:
'
'
    mKernelErr = mKernelErr & vbCrLf & "Definition Name: " & pName

End Sub

Private Sub StockKernelDefinition( _
    pKernelName As String, ByVal pKernelType As IdpKernelType, _
    ByVal pNumCol As Long, ByVal pNumRow As Long, ByVal pShiftR As Long, ByVal pKernelVal As String, Optional ByVal pComment As String = "" _
)
'シート上の定義情報を蓄えておきます。

    'テーブルにデータ追加
    Dim newKernel As CImgKernelInfo
    Set newKernel = New CImgKernelInfo
    Call newKernel.Create(pKernelName, pKernelType, pNumCol, pNumRow, pShiftR, pKernelVal, pComment)
    Call mTmpKernelInfos.Add(newKernel, pKernelName)

End Sub

Private Sub TriggerRegDifinition()
'内容:
'   蓄えたシート上の定義をIG-XLに登録します。
'   全部のデータが正しいことを確認して、IG-XLに登録します。
'   無用なTheIDPの破棄を避けます。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月10日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'
'


    On Error GoTo ERROR_IGXL
    
    Dim pKernel As CImgKernelInfo
    For Each pKernel In mTmpKernelInfos
    
        Call TheHdw.IDP.CreateKernel(pKernel.Name, pKernel.KernelType, pKernel.width, pKernel.height, pKernel.ShiftR, pKernel.Value)
    
    Next pKernel

    'EeeJOBチェック問題なし、IG-XL登録問題なし。カーネルシート情報登録完了
    Set mKernelInfos = mTmpKernelInfos
    Set mTmpKernelInfos = New Collection
    
    Exit Sub

ERROR_IGXL:
    'このエラーの場合、EeeJOBとIG-XLで情報に差分が発生してるので
    'mFlagKernelErrorをTrueにして、外部からTheIDP.RemoveResoucesできるようにしておく。
    Set mTmpKernelInfos = New Collection
    mFlagKernelError = True
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME, mKernelErr)

End Sub

Private Function ChkShiftR(ByVal pShiftRbit As Integer) As Boolean
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月3日
'パラメータ:
'   [pShiftRbit]    In/Out  1):
'戻り値:
'   Integer
'
'注意事項:
'
'
    If pShiftRbit >= 0 And pShiftRbit <= 16 Then
        ChkShiftR = True
    Else
        ChkShiftR = False
    End If
    
End Function

Private Function CIdpKernel(ByVal pKernelType As String) As IdpKernelType
'内容:
'   文字情報をidpKernelTypeに変換。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月3日
'パラメータ:
'   [pKernelType]   In/Out  1):
'戻り値:
'   IdpKernelType
'
'注意事項:
'   該当しない場合は-1

    pKernelType = UCase(pKernelType)    '大小文字無視
    
    If pKernelType = "INTEGER" Then
        CIdpKernel = idpKernelInteger
    ElseIf pKernelType = "FLOAT" Then
        CIdpKernel = idpKernelFloat
    Else
        CIdpKernel = -1
    End If

End Function

Private Function ChkSize(pSize As Integer) As Boolean
'内容:
'   大きさが1～64かチェックして問題なければそのまま返す。エラーなら-1を返す。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年3月4日
'パラメータ:
'   [pSize] In/Out  1):
'戻り値:
'   Integer
'
'注意事項:
'
'
    If pSize >= 1 And pSize <= 64 Then
        ChkSize = True
    Else
        ChkSize = False
    End If

End Function

Private Function IsKey(ByVal pKey As String, ByRef pObj As Collection) As Boolean
'内容:
'   該当Collectionオブジェクトにキーが存在するか調べる
'作成者:
'  tomoyoshi.takase
'作成日: 2010年10月22日
'パラメータ:
'   なし
'戻り値:
'   Boolean :Trueすでにadd済み。存在します。       Falseまだ無し
'
'注意事項:
'
    On Error GoTo ALREADY_REG
    Call pObj.Item(pKey)
    IsKey = True
    Exit Function

ALREADY_REG:
    IsKey = False

End Function

Private Sub Class_Initialize()
'内容:
'   コンストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'
'
    Set mKernelInfos = New Collection
    Set mTmpKernelInfos = New Collection

End Sub

Private Sub Class_Terminate()
'内容:
'   デストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'
'
    Set mKernelInfos = Nothing
    Set mTmpKernelInfos = Nothing

End Sub
