VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIGXLImageISTARFileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   IDV形式で画像出力するライター
'
'目的:
'   指定されたフォルダーにISTAR形式のファイル名をつけて画像を出力する
'
'作成者:
'   0145184306
'
Option Explicit

Implements IProcParamWriter

Private Const ERR_NUMBER = 9999         ' エラー番号を保持する
Private Const CLASS_NAME = "CIGXLImageISTARFileWriter" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default値"Empty"を保持する

Private m_FilePath As String
Private mParamName As String
Private mImgFileName() As String

Public Function AsIProcParamWriter() As IProcParamWriter
'内容:
'   自分自身が実装しているインターフェイスオブジェクト型へのキャスト
'
'戻り値：
'   IProcParamWriter型オブジェクト
'
'備考:
'
    Set AsIProcParamWriter = Me
End Function

Public Sub Initialize(ByVal pFolder As String, pParamName As String)
'内容:
'   イニシャライズ
'   ライターがファイル出力するフォルダをセット
'   フォルダの存在確認も行う
'   パラメータ名をセット
'   出力ファイル情報出力のため、CDatalogWriterオブジェクトを生成
'
'引数:
'[strFileName]  IN String型:     書き込み先フォルダ
'[pParamName[ IN String型：      パラメータ名
'
'備考:
'
    If Dir(pFolder, vbDirectory) = "" Then
        TheError.Raise ERR_NUMBER, CLASS_NAME & ".Initialize", "This Folder doesn't exist!"
    Else
        m_FilePath = pFolder
    End If
    mParamName = pParamName
End Sub

Private Sub IProcParamWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
End Sub

Private Sub IProcParamWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
End Sub

Private Sub IProcParamWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
End Sub

Private Sub IProcParamWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
End Sub

Private Sub IProcParamWriter_WriteAsPlane(ByVal pParamName As String, ByRef pImgPlane As CImgPlane)
'内容:
'   指定Plane画像を出力する
'
'引数:
'[pParamName]  IN String型:     テストインスタンス名
'[pParamPlane] IN CImgPlane型:  出力Plane
'
'備考:
'
'
    Dim site As Long
    Dim ImgFileName As String
    ReDim mImgFileName(TheExec.sites.ExistingCount - 1)

    If m_FilePath = INITIAL_EMPTY_VALUE Then
        TheError.Raise ERR_NUMBER, CLASS_NAME & ".WriteAsPlane", "Initialize Function is not done!!"
        Exit Sub
    Else
        For site = 0 To TheExec.sites.ExistingCount - 1
            ImgFileName = makeISTARImgFileName(site, mParamName)
            ImgFileName = m_FilePath & "\" & ImgFileName
            Call pImgPlane.WriteFile(site, ImgFileName)
            mImgFileName(site) = ImgFileName
            'File名重複回避
            TheHdw.WAIT 1 * S
        Next site
    End If
End Sub

Private Sub IProcParamWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
End Sub

Private Property Get IProcParamWriter_MyName() As String
    IProcParamWriter_MyName = CLASS_NAME
End Property

Private Function makeISTARImgFileName(ByVal pSite As Long, ByVal pInstanceName As String) As String
'内容:
'   ISTAR形式のファイル名を作成する
'
'引数:
'[pSite]       IN/OUT Integer型:    サイト
'[pInstanceName] IN/OUT String型:   テストインスタンス名
'
'戻り値:
'   ISTAR形式ファイル名
'
'備考:
'
'
    makeISTARImgFileName = "Test_" & pSite & ".stb"    'for debug

    Dim site As Long
    Dim strDeviceType As String
    Dim lngLotNo As Long
    Dim lngWaferNo As Long
    Dim lngChipNo As Long
    Dim lngXAddress As Long
    Dim lngYAddress As Long
    Dim lngDate As Variant

    Const REDUCE_VAL As Integer = 1   '縮小倍率係数
    Const FILE_EXTENSION As String = ".stb"
    'システム情報取得
    With TheSystemInfo
        strDeviceType = .ProductName
        lngLotNo = .LotNo
        lngWaferNo = .WaferNo
        lngChipNo = .GetChipNoOfSite(pSite)
        lngXAddress = .GetXAddressOfSite(pSite)
        lngYAddress = .GetYAddressOfSite(pSite)
    End With
    '日付取得
    lngDate = Now

    'ISTAR用ファイル名を作成
    makeISTARImgFileName = strDeviceType & "_LOTNAME" & Format(CStr(lngLotNo), "00") & "-" & _
                                Format(CStr(lngWaferNo), "00") & Format(CStr(lngChipNo), "0000") & "-" & _
                                    Format(CStr(lngXAddress), "00") & "-" & Format(CStr(lngYAddress), "00") & "-" & _
                                        pInstanceName & "-" & REDUCE_VAL & "-" & Format(CStr(lngDate), "yyyymmddHHMMSS") & FILE_EXTENSION

End Function

Private Sub Class_Initialize()
    m_FilePath = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    m_FilePath = INITIAL_EMPTY_VALUE
End Sub

Private Sub IProcParamWriter_WriteAsBlockEnd(ByVal pParamName As String)
End Sub

Private Sub IProcParamWriter_WriteAsBlockStart(ByVal pParamName As String)
End Sub

Public Sub IProcParamWriter_ReturnToRootBlock()
End Sub

Public Property Get ImgFileName(ByVal pIndex As Integer) As String
    ImgFileName = mImgFileName(pIndex)
End Property

