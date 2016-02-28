VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIGXLImageFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   stbファイル画像読み込みクラス
'
'目的:
'   stbファイルの画像を読み込むリーダー
'
'作成者:
'   0145184306
'
Option Explicit

Implements IProcParamReader

Private Const ERR_NUMBER = 9999         ' エラー番号を保持する
Private Const CLASS_NAME = "CIGXLImageFileReader" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default値"Empty"を保持する

Private m_FilePath As String
Private mParamName As String
Private mImgFileName() As String

Public Function AsIProcParamReader() As IProcParamReader
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IParameterReaderX20型オブジェクト
'
'備考:
'
    Set AsIProcParamReader = Me
End Function

Public Sub Initialize(pFolder As String, pParamName As String)
'内容:
'   イニシャライズ
'   リーダーが検索を行うフォルダをセット
'   フォルダの存在確認も行う
'   パラメータ名をセット
'   入力ファイル情報出力のため、CDatalogWriterオブジェクトを生成
'
'引数:
'[strFileName]  IN String型:     書き込み先フォルダ
'[pParamName]  IN String型:      パラメータ名
'
'備考:
'
    'フォルダの存在確認
    If Dir(pFolder, vbDirectory) = "" Then
        TheError.Raise ERR_NUMBER, CLASS_NAME & ".Initialize", "This Folder doesn't exist!"
    Else
        m_FilePath = pFolder
    End If
    mParamName = pParamName
End Sub

Private Function IProcParamReader_ReadAsBoolean(ByVal dataName As String) As Boolean
    TheError.Raise ERR_NUMBER, CLASS_NAME, "Unsupported Value Type!!"
End Function


Private Function IProcParamReader_ReadAsDouble(ByVal dataName As String) As Double
    TheError.Raise ERR_NUMBER, CLASS_NAME, "Unsupported Value Type!!"
End Function

Private Function IProcParamReader_ReadAsInteger(ByVal dataName As String) As Integer
    TheError.Raise ERR_NUMBER, CLASS_NAME, "Unsupported Value Type!!"
End Function

Private Function IProcParamReader_ReadAsLong(ByVal dataName As String) As Long
    TheError.Raise ERR_NUMBER, CLASS_NAME, "Unsupported Value Type!!"
End Function

Private Function IProcParamReader_ReadAsPlane(ByVal pParamName As String, ByRef pImgPlane As CImgPlane)
'内容:
'   指定Planeにstbファイル画像を読み込む
'
'引数:
'[pParamName]  IN String型:     インスタンス名
'[pImgPlane]   IN/OUT CImgPlane型:  画像取り込み用Plane
'
'備考:
'   ファイル名はpParamName_site.stbに固定
'

    Dim site As Long
    Dim ImgFileName As String
    ReDim mImgFileName(TheExec.sites.ExistingCount - 1)
    
    For site = 0 To TheExec.sites.ExistingCount - 1
        ImgFileName = makeReadFileName(site, mParamName)
        ImgFileName = m_FilePath & "\" & ImgFileName
        'ファイルの存在確認
        If Dir(ImgFileName, vbNormal) = "" Then GoTo ErrorHandle
        With pImgPlane
            Call .ReadFile(site, ImgFileName)
        End With
        mImgFileName(site) = ImgFileName
    Next site
    Exit Function
ErrorHandle:
    TheError.Raise ERR_NUMBER, CLASS_NAME & ".ReadAsPlane", ImgFileName & " File does'nt exist!!"
End Function

Private Function IProcParamReader_ReadAsString(ByVal dataName As String) As String
    TheError.Raise ERR_NUMBER, CLASS_NAME, "Unsupported Value Type!!"
End Function

Private Property Get IProcParamReader_MyName() As String
    IProcParamReader_MyName = CLASS_NAME
End Property

Private Sub Class_Initialize()
    m_FilePath = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    m_FilePath = INITIAL_EMPTY_VALUE
End Sub

Private Function makeReadFileName(ByVal pSite As Long, ByVal pParamName As String) As String
'内容:
'   ISTAR形式のファイル名を作成する。
'
'引数:
'[pSite]       IN Integer型:    サイト
'[pParamName] IN String型:   テストインスタンス名
'
'戻り値:
'   ファイル名
'
'備考:
'   stbファイルのみ
'
    Const FILE_EXTENSION As String = ".stb"
    
    makeReadFileName = pParamName & "_" & pSite & FILE_EXTENSION
    
End Function

Public Function IProcParamReader_ReadAsBlockEnd(ByVal pParamName As String) As IProcParamReader
End Function

Public Function IProcParamReader_ReadAsBlockStart(ByVal pParamName As String) As IProcParamReader
End Function

Public Property Get ImgFileName(ByVal pIndex As Integer) As String
    ImgFileName = mImgFileName(pIndex)
End Property

