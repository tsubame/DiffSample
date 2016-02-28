VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   プレーン情報のラッパー。(旧CTraceablePlane)
'   デバイスプロファイラなどプレーンの生成環境の情報をセットで保持する。
'
'   Revision History:
'       Data        Description
'       2011/02/01  TraceablePlaneよりインターフェース化してCParamPlaneとして実装
'       2011/02/15  シナリオエンジン組み込みに伴いもろもろ修正
'       2011/02/18  Plane,DevInfoのWriteOnceエラー処理を削除(上書き可能), Dumpの引数pWriterをByVal=>ByRef
'       2011/02/23  結合検証結果より最新ソース取得。Initialize Byval => ByRef
'       2011/11/01  Readの実装
'
'目的:
'   デバイス情報、CImgPlaneを持ち、必要に応じてデバイス情報、画像インスタンスを返す。
'
'作成者:
'   0145184304
'

Option Explicit

Implements IProcParameter
Implements IProcParamReplaceable

Private Const ERR_NUMBER As Long = 9999                  ' Error番号を保持する
Private Const CLASS_NAME As String = "CParamPlane"       ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default値"Empty"を保持する

Private Const DUMPKEYWORD_FOR_PARAM_NAME As String = "Parameter Name"
Private Const DUMPKEYWORD_FOR_CLASS_NAME As String = "Parameter Type"
Private Const DUMPKEYWORD_FOR_COMMENT As String = "Comment"

Private Const PARAMETER_REPLACER_NAME As String = "Name"
Private Const PARAMETER_REPLACER_PLANE As String = "Plane"
Private Const PARAMETER_REPLACER_DEVICEINFO As String = "DeviceInfo"
Private Const PARAMETER_REPLACER_COMMENT As String = "Comment"
Private Const PARAMETER_REPLACER_TYPE As String = "Type"

Private mName As String
Private mCImgPlane As CImgPlane
Private mCDeviceConfigInfo As CDeviceConfigInfo
Private mComment As String

Public Sub Initialize(ByVal paramName As String, ByRef clsImgPlane As CImgPlane, ByRef clsDevInfo As CDeviceConfigInfo, Optional ByVal pComment As String = "")
'内容:
'   初期化
'
'[paramName]   IN String型:     パラメータ名を指定
'[clsImgPlane] IN CImgPlane型:  格納するCImagPlaneを指定
'[clsDevInfo]  IN CDeviceConfigInfo型:  格納するCDeviceConfigInfoを指定
'
'備考:
'

    mName = paramName
    
    Set mCImgPlane = clsImgPlane
    Set mCDeviceConfigInfo = clsDevInfo
    
    mComment = pComment

End Sub
Private Function IProcParameter_Clone() As IProcParameter
'内容:
'   Cloneを作成する
'
'戻り値:
'   CParamPlaneを返す（クローン）
'
'備考:
'
    
    '##### 自身が持っているDeviceInfoのクローンを作る #####
    Dim objFrameRate As CParamFrameRate
    Dim objAccTime As CParamAccTime
    Dim objGain As CParamGain
    Dim objLSB As CParamLsb
    
    Dim objDevInfo As CDeviceConfigInfo
    
    If Not Me.DeviceConfigInfo Is Nothing Then
        Set objDevInfo = Me.DeviceConfigInfo.AsIProcParameter.Clone
    Else
        Set objDevInfo = Nothing
    End If
    
    '#####  自身のCloneを作成して戻す  #####
    Dim cloneobj As CParamPlane
    Set cloneobj = New CParamPlane
    cloneobj.Initialize Me.AsIProcParameter.Name, Me.plane, objDevInfo, Me.AsIProcParameter.Comment
    Set IProcParameter_Clone = cloneobj
    Set cloneobj = Nothing

    Set objDevInfo = Nothing
    Set objFrameRate = Nothing
    Set objAccTime = Nothing
    Set objGain = Nothing
    Set objLSB = Nothing

End Function
Private Sub IProcParameter_Dump(ByRef pWriter As IProcParamWriter)
'内容:
'   PlaneとDeviceInfoとCommentをDumpする(pWriterにデータを引き渡す)
'
'[pWriter]   IN IProcParamWriter型:     DumpするWriterを指定する
'
'備考:
'

    '#####  自身のクラス名 / プレーン / デバイス情報 / コメント をDump #####
    ' Parameter Block
    pWriter.WriteAsString "", "Parameter " & Me.AsIProcParameter.Name & " {"
    pWriter.WriteAsBlockStart "Dump"
    pWriter.WriteAsString "", "Type: " & CLASS_NAME                             ' クラス名
    ' Plane Block
    pWriter.WriteAsString "", "Plane {"
    pWriter.WriteAsBlockStart Me.AsIProcParameter.Name
    pWriter.WriteAsPlane Me.AsIProcParameter.Name, Me.plane                     ' プレーン
    pWriter.WriteAsBlockEnd "ParamPlane"
    pWriter.WriteAsString "", "}"
    ' DeviceConfigInfo Block
    If Not Me.DeviceConfigInfo Is Nothing Then
        Me.DeviceConfigInfo.AsIProcParameter.Dump pWriter                           ' デバイス情報
    End If
    pWriter.WriteAsString "", "Comment: " & Me.AsIProcParameter.Comment         ' コメント
    pWriter.WriteAsBlockEnd "Dump"
    pWriter.WriteAsString "", "}"
    
    
End Sub

Private Property Let IProcParameter_Name(ByVal pName As String)
'内容:
'   パラメータ名を保持する
'
'[pName]   IN String型:     保持するパラメータ名を指定する
'
'備考:
'

    mName = pName

End Property

Private Property Get IProcParameter_Name() As String
'内容:
'   保持したパラメータ名を返す
'
'戻り値:
'   保持したパラメータ名
'
'備考:
'

    IProcParameter_Name = mName

End Property

Private Property Let IProcParameter_Comment(ByVal pComment As String)
'内容:
'   コメントを保持する
'
'[pComment]   IN String型:     保持するコメントを指定する
'
'備考:
'

    mComment = pComment

End Property

Private Property Get IProcParameter_Comment() As String
'内容:
'   保持したコメントを返す
'
'戻り値:
'   保持したコメント
'
'備考:
'

    IProcParameter_Comment = mComment

End Property

Public Property Set plane(pPlane As CImgPlane)
'内容:
'   プレーンを保持する
'
'[pPlane]   IN CImgPlane型:     保持するプレーンを指定する
'
'備考:
'

        Set mCImgPlane = pPlane

End Property

Public Property Get plane() As CImgPlane
'内容:
'   保持したプレーンを返す
'
'戻り値:
'   保持したプレーン
'
'備考:
'

    Set plane = mCImgPlane

End Property


Public Property Set DeviceConfigInfo(pDeviceInfo As CDeviceConfigInfo)
'内容:
'   デバイス情報を保持する
'
'[pDeviceInfo]   IN CDeviceConfigInfo型:     保持するデバイス情報を指定する
'
'備考:
'
        Set mCDeviceConfigInfo = pDeviceInfo

End Property

Public Property Get DeviceConfigInfo() As CDeviceConfigInfo
'内容:
'   保持したデバイス情報を返す
'
'戻り値:
'   保持したデバイス情報
'
'備考:
'

    Set DeviceConfigInfo = mCDeviceConfigInfo

End Property

Public Function PixelLog(ByVal pFlgName As String, Optional ByVal pMaxNum As Long = 0, _
                            Optional ByVal pAddrMode As IdpAddrMode = idpAddrAbsolute) As CParamDefect
'内容:
'   該当するフラグのPixelログをCParamDefect型で返します。
'パラメータ:
'   [pFlgName]  In  1):
'   [pMaxNum]   In  2):
'   [pAddrMode] In  3):
'戻り値:
'   CParamDefect
'
'注意事項:
'   CImgPlaneのPixelLogはSubですが、こちらはFunctionです。


    '-- 変数定義など ---------------------------------------
    Dim pSiteCount As Long                       'SiteLoopカウント用
    Dim pTypePointDefects() As T_PIXINFO         '欠陥最大個数分の画素情報構造体
    Dim pDefectPixel As CParamPixelLog           '1画素分の欠陥情報
    Dim pThisSiteDefects As CParamSitePixelLog   '1Site分の点欠陥情報
    Dim pAllSiteDefects() As CParamSitePixelLog  '全サイトの点欠陥情報全部入り
    Dim pDefects As CParamDefect                 '全サイトの点欠陥情報全部入りオブジェクト
    Dim pColor As CImgColor                     '
    Dim pPixelCount As Long                      '欠陥個数
    '--------------------------------------------------------
    
    '-- 初期化処理 ------------------------------------------
    '全サイト欠陥情報(箱はサイトの有効無効に限らず全サイト数分用意)
    ReDim pAllSiteDefects(TheExec.sites.ExistingCount - 1)
    '--------------------------------------------------------
    
    'SiteLOOPで処理実行
    For pSiteCount = 0 To UBound(pAllSiteDefects) Step 1
        '有効サイトのみ情報取得
        If TheExec.sites.site(pSiteCount).Active = True Then
            Call Me.plane.PixelLog(pSiteCount, pFlgName, pTypePointDefects, pMaxNum, pAddrMode)
            'サイト毎欠陥情報作成
            Set pThisSiteDefects = New CParamSitePixelLog
            Call pThisSiteDefects.Initialize(pSiteCount)
            '欠陥画素LOOP
            If UBound(pTypePointDefects) <> -1 Then
                For pPixelCount = 0 To UBound(pTypePointDefects) Step 1
'Withステートメント使うと動的配列がロックされる
'                    With pTypePointDefects(pPixelCount)
'                        If .X = 0 Or .Y = 0 Then
'                            Exit For
'                        End If
'                        Set pDefectPixel = New CParamPixelLog
'                        Set pColor = mCImgPlane.ColorByAddress(.X, .Y, pAddrMode)
'                        Call pDefectPixel.Initialize(.X, .Y, .Value, pColor)
'                    End With
                    If pTypePointDefects(pPixelCount).x = 0 Or pTypePointDefects(pPixelCount).y = 0 Then
                        Exit For
                    End If
                    Set pDefectPixel = New CParamPixelLog
                    Set pColor = mCImgPlane.ColorByAddress(pTypePointDefects(pPixelCount).x, pTypePointDefects(pPixelCount).y, pAddrMode)
                    Call pDefectPixel.Initialize(pTypePointDefects(pPixelCount).x, pTypePointDefects(pPixelCount).y, pTypePointDefects(pPixelCount).Value, pColor)
                    Call pThisSiteDefects.Add(pDefectPixel)
                Next
            End If
            'サイトの欠陥情報を設定
            Set pAllSiteDefects(pSiteCount) = pThisSiteDefects
        Else
            'サイトの欠陥情報を設定
            Set pAllSiteDefects(pSiteCount) = Nothing
        End If
    Next
    
    '最終結果登録処理
    '（無効サイトの値を取得した際は実行時エラーでなくNothingが戻ることになる）
    Set pDefects = New CParamDefect
    For pSiteCount = 0 To UBound(pAllSiteDefects) Step 1
        Call pDefects.Add(pAllSiteDefects(pSiteCount), pSiteCount)
    Next
    Set pDefects.DeviceConfig = mCDeviceConfigInfo
    Set PixelLog = pDefects

End Function

Private Sub Class_Initialize()

    Set mCImgPlane = Nothing
    Set mCDeviceConfigInfo = Nothing

End Sub

Private Sub Class_Terminate()

    mName = INITIAL_EMPTY_VALUE
    Set mCImgPlane = Nothing
    Set mCDeviceConfigInfo = Nothing
    mComment = INITIAL_EMPTY_VALUE

End Sub
Public Function AsIProcParameter() As IProcParameter
    Set AsIProcParameter = Me
End Function

Public Sub IProcParameter_Read(ByRef pReader As IProcParamReader)
'内容:
'   指定Planeにstbファイル画像を読み込む
'
'引数:
'[pReader]     IN/OUT IProcParamReader型:   Reader
'
'備考:
'
'

    Call pReader.ReadAsBlockStart("ParamPlane")
    Call pReader.ReadAsPlane(Me.AsIProcParameter.Name, mCImgPlane)
    Call pReader.ReadAsBlockEnd("ParamPlane")
End Sub

Private Sub IProcParamReplaceable_GetContents(ByVal pRoot As String, ByRef pRes As Collection)
'内容:
'   パラメータの設定内容を取得する
'
'パラメータ:
'   [pRoot]     IN  String型:  このパラメータクラスの階層情報
'   [pRes]      IN/OUT  Collection型:  パラメータの設定内容を格納するコレクション
'
'注意事項:
'
'

    Dim bufSubParam As IProcParamReplaceable
    With pRes
        '   2012-01-27 (0145184931): Nameを編集不可に変更
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_NAME, mName, "String", False, PARAMETER_REPLACER_NAME)
        Set bufSubParam = mCImgPlane
        If Not bufSubParam Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PLANE, vbNullString, "CImgPlane", False, PARAMETER_REPLACER_PLANE)
            Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_PLANE, pRes)
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_PLANE, "Nothing", "CImgPlane", False, PARAMETER_REPLACER_PLANE)
        End If
        Set bufSubParam = mCDeviceConfigInfo
        If Not bufSubParam Is Nothing Then
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, vbNullString, "CDeviceConfigInfo", False, PARAMETER_REPLACER_DEVICEINFO)
            Call bufSubParam.GetContents(pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, pRes)
        Else
            Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_DEVICEINFO, "Nothing", "CDeviceConfigInfo", False, PARAMETER_REPLACER_DEVICEINFO)
        End If
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_COMMENT, mComment, "String", True, PARAMETER_REPLACER_COMMENT)
        Call AddParamContents(pRes, pRoot & "\" & PARAMETER_REPLACER_TYPE, CLASS_NAME, "String", False, PARAMETER_REPLACER_TYPE)
    End With
End Sub

Private Sub IProcParamReplaceable_ReplaceContent(ByVal pRoot As String, ByRef pParamContents As CImpLocalParamsContent)
'内容:
'   パラメータの内容を書き換える
'
'パラメータ:
'   [pRoot]             IN  String型:  このパラメータクラスの階層情報
'   [pParamContents]    IN  CImpLocalParamsContent型:  書き換えを行うパラメータの内容
'
'注意事項:
'
'

    Dim targetSubParam As IProcParamReplaceable
    Dim targetIndex As Long
    With pParamContents
        Dim subParamLabel As String
        subParamLabel = XLibImpExperimentUtility.GetSubParamLabel(.Path, pRoot)
        Select Case subParamLabel
'   2012-01-27 (0145184931): Nameを編集不可に変更
'            Case PARAMETER_REPLACER_NAME
'                mName = .value
            Case PARAMETER_REPLACER_COMMENT
                mComment = .Value
        End Select
    End With
End Sub

Private Sub AddParamContents(ByRef pRes As Collection, ByVal pPath As String, ByVal pValue As String, ByVal pTypeName As String, ByVal pEditable As Boolean, ByVal pLabel As String)
'内容:
'   パラメータの情報をコレクションに追加する
'
'パラメータ:
'   [pRes]          IN/OUT  Collection型:  パラメータの設定内容を追加するコレクション
'   [pPath]         IN  String型:  パラメータの階層情報
'   [pValue]        IN  String型:  パラメータの値
'   [pTypeName]     IN  String型:  パラメータのデータ型
'   [pEditable]     IN  Boolean型:  パラメータが編集可能かどうか
'   [pLabel]        IN  String型:  パラメータ名
'
'注意事項:
'
'

    Dim myContents As New CImpLocalParamsContent
    Call myContents.Initialize(pPath, pValue, pTypeName, pEditable, pLabel)
    Call pRes.Add(myContents)
End Sub
