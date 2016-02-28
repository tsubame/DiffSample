Attribute VB_Name = "XLibImageMod"
'概要:
'   画像処理ライブラリ
'
'目的:
'   一般的によく使う処理の纏め
'
'作成者:
'   0145184004
'
Option Explicit

Private Const TMP_SIZE = 11

'2009/09/15 D.Maruyama ColorAll処理追加 ここから
Public Type SiteValues
    SiteValue() As Double
End Type

Public Type ColorAllResult
    color(TMP_SIZE) As SiteValues
End Type
'2009/09/15 D.Maruyama ColorAll処理追加 ここまで

Public Sub Average( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'内容:
'   平均値を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[retResult()] OUT  Double型:       結果格納用配列(サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.Average(retResult, srcColor, pFlgName)

End Sub

Public Sub sum( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'内容:
'   合計値を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[retResult()] OUT  Double型:       結果格納用配列(サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.sum(retResult, srcColor, pFlgName)

End Sub

Public Sub StdDev( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'内容:
'   標準偏差を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[retResult()] OUT  Double型:       結果格納用配列(サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.StdDev(retResult, srcColor, pFlgName)

End Sub

Public Sub GetPixelCount( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'内容:
'   対象の画素数を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[retResult()] OUT  Double型:       結果格納用配列(サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.Num(retResult, srcColor, pFlgName)

End Sub

Public Sub Min( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'内容:
'   最小値を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[retResult()] OUT  Double型:       結果格納用配列(サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.Min(retResult, srcColor, pFlgName)

End Sub

Public Sub max( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'内容:
'   最大値を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[retResult()] OUT  Double型:       結果格納用配列(サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.max(retResult, srcColor, pFlgName)

End Sub

Public Sub MinMax( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retMin() As Double, ByRef retMax() As Double, _
    Optional pFlgName As String = "" _
)
'内容:
'   最小値、最大値を一度に取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[retMin()]    OUT  Double型:       最小値格納用配列(サイト分)
'[retMax()]    OUT  Double型:       最大値格納用配列(サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)

'=======2009/05/11 変更 Maruyama ここから==============
'    Call srcPlane.Min(retMin, srcColor, pFlgName)
'    Call srcPlane.Max(retMax, srcColor, pFlgName)
    Call srcPlane.MinMax(retMin, retMax, srcColor, pFlgName)
'=======2009/05/11 変更 Maruyama ここまで==============

End Sub

Public Sub DiffMinMax( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'内容:
'   最小値と最大値の差を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[retResult()] OUT  Double型:       結果格納用配列(サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.DiffMinMax(retResult, srcColor, pFlgName)

End Sub

Public Sub AbsMax( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'内容:
'   最小値と最大値の内絶対値の大きい方を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[retResult()] OUT  Double型:       結果格納用配列(サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.AbsMax(retResult, srcColor, pFlgName)

End Sub

Public Sub Count( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal limitType As IdpLimitType, ByRef retResult() As Double, _
    Optional ByVal pFlgName As String = "", Optional ByVal pInputFlgName As String = "" _
)
'内容:
'   条件に該当する点の個数を取得する。
'
'[srcPlane]         IN   CImgPlane型:    対象プレーン
'[srcZone]          IN   String型:       対象プレーンのゾーン指定
'[srcColor]         IN   IdpColorType型: 対象プレーンの色指定
'[countType]        IN   IdpCountType型: カウント条件指定
'[loLim]            IN   Variant型:      下限値
'[hiLim]            IN   Variant型:      上限値
'[limitType]        IN   IdpLimitType型: 境界値を含む、含まない指定
'[retResult()]      OUT  Double型:       結果格納用配列(動的配列)
'[pFlgName]         IN   String型:       出力フラグ名
'[pInputFlgName]    IN   String型:       入力フラグ名
'
'備考:
'   hiLim,loLimをVariantで定義しているのは、サイト別で境界値が違う場合に対応するため。
'   サイト配列を入れると各サイトごと別々に、定数を入れると全サイトに同じ値が適用される。
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.Count(retResult, countType, loLim, hiLim, limitType, srcColor, pFlgName, pInputFlgName)

End Sub

'=======2009/05/19 Add Maruyama ここから==============
Public Sub CountForFlgBitImgPlane( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal limitType As IdpLimitType, ByRef retResult() As Double, _
    ByRef pFlgPlane As CImgPlane, ByVal pFlgBit As Long, _
    Optional ByVal pInputFlgName As String = "" _
)
'内容:
'   条件に該当する点の個数を取得する。(フラグビットをイメージプレンに立てる)
'
'[srcPlane]         IN   CImgPlane型:    対象プレーン
'[srcZone]          IN   String型:       対象プレーンのゾーン指定
'[srcColor]         IN   IdpColorType型: 対象プレーンの色指定
'[countType]        IN   IdpCountType型: カウント条件指定
'[loLim]            IN   Variant型:      下限値
'[hiLim]            IN   Variant型:      上限値
'[limitType]        IN   IdpLimitType型: 境界値を含む、含まない指定
'[retResult()]      OUT  Double型:       結果格納用配列(動的配列)
'[pFlgName]         IN   CImgPlane型:       出力フラグ名
'[pFlgBit]        　IN   Long型:       出力フラグ名
'[pInputFlgName]    IN   String型:       入力フラグ名
'
'備考:
'   hiLim,loLimをVariantで定義しているのは、サイト別で境界値が違う場合に対応するため。
'   サイト配列を入れると各サイトごと別々に、定数を入れると全サイトに同じ値が適用される。
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call pFlgPlane.SetPMD(srcZone)
    Call srcPlane.CountForFlgBitImgPlane(retResult, countType, loLim, hiLim, pFlgPlane, pFlgBit, limitType, srcColor, pInputFlgName)

End Sub
'=======2009/05/19 Add Maruyama ここまで==============

Public Sub PutFlag( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal limitType As IdpLimitType, _
    ByVal pFlgName As String, Optional ByVal pInputFlgName As String _
)
'内容:
'   条件に該当する点にフラグを立てる。(Countからフラグを立てることだけに特化)
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[countType]   IN   IdpCountType型: カウント条件指定
'[loLim]       IN   Variant型:      下限値
'[hiLim]       IN   Variant型:      上限値
'[limitType]   IN   IdpLimitType型: 境界値を含む、含まない指定
'[pFlgName]         IN   String型:       出力フラグ名
'[pInputFlgName]    IN   String型:       入力フラグ名
'
'備考:
'   hiLim,loLimをVariantで定義しているのは、サイト別で境界値が違う場合に対応するため。
'   サイト配列を入れると各サイトごと別々に、定数を入れると全サイトに同じ値が適用される。
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.PutFlag(pFlgName, countType, loLim, hiLim, limitType, srcColor, pInputFlgName)

End Sub

Public Sub Add( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'内容:
'   対象画像を加算する
'
'[srcPlane]    IN   CImgPlane型:    元プレーン1
'[srcZone]     IN   String型:       元プレーン1のゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーン1の色指定
'[srcPlane2]   IN   CImgPlane型:    元プレーン2
'[srcZone2]    IN   String型:       元プレーン2のゾーン指定
'[srcColor2]   IN   IdpColorType型: 元プレーン2の色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'
'備考:
'   各プレーンには同じものを指定することが可能。
'   その場合ゾーン指定は最後のもので統一されてしまうので要注意。
'   例)
'       Add(dst, "ZONE3", EEE_COLOR_FLAT, src, "ZONE3" EEE_COLOR_FLAT, dst, "ZONE3_2", EEE_COLOR_FLAT)
'       この場合dstは ZONE3_2 だけが対象になる。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Add(srcPlane, srcPlane2, dstColor, srcColor, srcColor2)

End Sub

Public Sub AddConst( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal addVal As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'内容:
'   対象画像に指定値を足す。
'
'[srcPlane]    IN   CImgPlane型:    元プレーン
'[srcZone]     IN   String型:       元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーンの色指定
'[addVal]      IN   Variant型:      加算値
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'
'備考:
'   各プレーンには同じものを指定することが可能。
'   その場合ゾーン指定は最後のもので統一されてしまうので要注意。
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Add(srcPlane, addVal, dstColor, srcColor)

End Sub

Public Sub Subtract( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'内容:
'   対象画像を減算する
'
'[srcPlane]    IN   CImgPlane型:    元プレーン1
'[srcZone]     IN   String型:       元プレーン1のゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーン1の色指定
'[srcPlane2]   IN   CImgPlane型:    元プレーン2
'[srcZone2]    IN   String型:       元プレーン2のゾーン指定
'[srcColor2]   IN   IdpColorType型: 元プレーン2の色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'
'備考:
'   各プレーンには同じものを指定することが可能。
'   その場合ゾーン指定は最後のもので統一されてしまうので要注意。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Subtract(srcPlane, srcPlane2, dstColor, srcColor, srcColor2)

End Sub

Public Sub SubtractConst( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal subVal As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'内容:
'   対象画像から指定値を引く。
'
'[srcPlane]    IN   CImgPlane型:    元プレーン
'[srcZone]     IN   String型:       元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーンの色指定
'[subVal]      IN   Variant型:      減算値
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'
'備考:
'   各プレーンには同じものを指定することが可能。
'   その場合ゾーン指定は最後のもので統一されてしまうので要注意。
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Subtract(srcPlane, subVal, dstColor, srcColor)

End Sub

Public Sub Multiply( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'内容:
'   対象画像を乗算する
'
'[srcPlane]    IN   CImgPlane型:    元プレーン1
'[srcZone]     IN   String型:       元プレーン1のゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーン1の色指定
'[srcPlane2]   IN   CImgPlane型:    元プレーン2
'[srcZone2]    IN   String型:       元プレーン2のゾーン指定
'[srcColor2]   IN   IdpColorType型: 元プレーン2の色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'
'備考:
'   各プレーンには同じものを指定することが可能。
'   その場合ゾーン指定は最後のもので統一されてしまうので要注意。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Multiply(srcPlane, srcPlane2, dstColor, srcColor, srcColor2)

End Sub

Public Sub MultiplyConst( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal mulVal As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'内容:
'   対象画像に指定値を掛ける。
'
'[srcPlane]    IN   CImgPlane型:    元プレーン
'[srcZone]     IN   String型:       元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーンの色指定
'[mulVal]      IN   Variant型:      乗算値
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'
'備考:
'   各プレーンには同じものを指定することが可能。
'   その場合ゾーン指定は最後のもので統一されてしまうので要注意。
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Multiply(srcPlane, mulVal, dstColor, srcColor)

End Sub

Public Sub MultiplyConstFlag( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal mulVal As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, Optional ByVal pInputFlgName As String = "" _
)
'内容:
'   対象画像に指定値を掛ける。
'
'[srcPlane]    IN   CImgPlane型:    元プレーン
'[srcZone]     IN   String型:       元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーンの色指定
'[mulVal]      IN   Variant型:      乗算値
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'
'備考:
'   各プレーンには同じものを指定することが可能。
'   その場合ゾーン指定は最後のもので統一されてしまうので要注意。
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Multiply(srcPlane, mulVal, dstColor, srcColor, , pInputFlgName)

End Sub

Public Sub Divide( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'内容:
'   対象画像を除算する
'
'[srcPlane]    IN   CImgPlane型:    元プレーン1
'[srcZone]     IN   String型:       元プレーン1のゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーン1の色指定
'[srcPlane2]   IN   CImgPlane型:    元プレーン2
'[srcZone2]    IN   String型:       元プレーン2のゾーン指定
'[srcColor2]   IN   IdpColorType型: 元プレーン2の色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'
'備考:
'   各プレーンには同じものを指定することが可能。
'   その場合ゾーン指定は最後のもので統一されてしまうので要注意。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Divide(srcPlane, srcPlane2, dstColor, srcColor, srcColor2)

End Sub

Public Sub DivideConst( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal divVal As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'内容:
'   対象画像を指定値で割る。
'
'[srcPlane]    IN   CImgPlane型:    元プレーン
'[srcZone]     IN   String型:       元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーンの色指定
'[divVal]      IN   Variant型:      除算値
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'
'備考:
'   各プレーンには同じものを指定することが可能。
'   その場合ゾーン指定は最後のもので統一されてしまうので要注意。
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Divide(srcPlane, divVal, dstColor, srcColor)

End Sub

Public Sub Median( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal width As Long, ByVal height As Long _
)
'内容:
'   対象画像にメディアンフィルタを掛ける。
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'[Width]       IN   Long型:         フィルタ幅
'[Height]      IN   Long型:         フィルタ高さ
'
'備考:
'
    Dim Center As Long

    Center = (width * height + 1) / 2

    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.RankFilter(srcPlane, width, height, Center, dstColor, srcColor)

End Sub

Public Sub MedianHV( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal width As Long, ByVal height As Long _
)
'内容:
'   対象画像にメディアンフィルタを掛ける。
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'[Width]       IN   Long型:         フィルタ幅
'[Height]      IN   Long型:         フィルタ高さ
'
'備考:
'
    Dim tmpPlane As CImgPlane
    Set tmpPlane = TheIDP.PlaneManager(srcPlane.planeGroup).GetFreePlane(srcPlane.BitDepth)
    
    Call Median(srcPlane, srcZone, srcColor, tmpPlane, srcZone, srcColor, width, 1)
    Call Median(tmpPlane, srcZone, srcColor, dstPlane, dstZone, dstColor, 1, height)
    
End Sub

Public Sub MedianVH( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal width As Long, ByVal height As Long _
)
'内容:
'   対象画像にメディアンフィルタを掛ける。
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'[Width]       IN   Long型:         フィルタ幅
'[Height]      IN   Long型:         フィルタ高さ
'
'備考:
'
    Dim tmpPlane As CImgPlane
    Set tmpPlane = TheIDP.PlaneManager(srcPlane.planeGroup).GetFreePlane(srcPlane.BitDepth)
    
    Call Median(srcPlane, srcZone, srcColor, tmpPlane, srcZone, srcColor, 1, height)
    Call Median(tmpPlane, srcZone, srcColor, dstPlane, dstZone, dstColor, width, 1)
    
End Sub

Public Sub Convolution( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal Kernel As String, Optional ByVal divVal As Long = 0 _
)
'内容:
'   対象画像にコンボリューションフィルタを掛ける。
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'[kernel]      IN   String型:       フィルタ名
'[divVal]      IN   Long型:         割戻し用の値
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Convolution(srcPlane, Kernel, dstColor, srcColor)

    If divVal <> 0 Then
        Call DivideConst(dstPlane, dstZone, dstColor, divVal, dstPlane, dstZone, dstColor)
    End If

End Sub

Public Sub ExecuteLUT( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal lutName As String _
)
'内容:
'   対象画像にルックアップテーブルを掛ける。
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 対象プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'[lutName]     IN   String型:       LUT名
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.ExecuteLUT(srcPlane, lutName, dstColor, srcColor)

End Sub

Public Sub Copy( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional pmask As String = "")
'内容:
'   対象画像にメディアンフィルタを掛ける。
'
'[srcPlane]    IN   CImgPlane型:    コピー元プレーン
'[srcZone]     IN   String型:       コピー元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: コピー元プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:    コピー先プレーン
'[dstZone]     IN   String型:       コピー先プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: コピー先プレーンの色指定
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.CopyPlane(srcPlane, dstColor, srcColor, , pmask)

End Sub

Public Sub WritePixel(ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, ByVal writeVal As Double, Optional ByVal mask As Long = 0)
'内容:
'   対象画像に値を書き込む。
'
'[dstPlane]     OUT CImgPlane型:    対象プレーン
'[dstZone]      IN  String型:       対象プレーンのゾーン指定
'[dstColor]     IN  IdpColorType型: 対象プレーンの色指定
'[writeVal]     IN  Double型:       書き込む値
'[mask]         IN  Long型:         マスク指定
'
'備考:
'   maskを指定すると1の立ったビットは無視される。
'   例)
'       WritePixel("vmcu00", "ZONE3", 0, &HFFF0)　とした場合下位4bitのみに0が書き込まれる。
'
'   本当はこの機能は使わないようにすべきと思う。
'   元プログラムの中でこれを使っていたので入れたが、
'   ビット演算をしたいのであればLOr,LAndなどが用意されているからそれを使うべき。

    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.WritePixel(writeVal, dstColor, , , mask)

End Sub

Public Sub MultiMean( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal multiMeanFunc As IdpMultiMeanFunc, ByVal width As Long, ByVal height As Long _
)
'内容:
'   マルチミーンを行う
'
'[srcPlane]    IN   CImgPlane型:        元プレーン
'[srcZone]     IN   String型:           元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型:     元プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:        結果格納プレーン
'[dstZone]     IN   String型:           結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型:     結果格納プレーンの色指定
'[multiMeanFunc] IN IdpMultiMeanFunc型: 演算方法指定(Max,Min,Mean,Sum)
'[Width]       IN   Long型:             幅指定
'[Height]      IN   Long型:             高さ指定
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.MultiMean(srcPlane, width, height, multiMeanFunc, dstColor, srcColor)

End Sub

Public Sub MultiMeanByBlock( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal multiMeanFunc As IdpMultiMeanFunc, ByVal DivX As Long, ByVal DivY As Long _
)
'内容:
'   マルチミーンを行う
'
'[srcPlane]    IN   CImgPlane型:        元プレーン
'[srcZone]     IN   String型:           元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型:     元プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:        結果格納プレーン
'[dstZone]     IN   String型:           結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型:     結果格納プレーンの色指定
'[multiMeanFunc] IN IdpMultiMeanFunc型: 演算方法指定(Max,Min,Mean,Sum)
'[DivX]        IN   Long型:             横分割指定
'[DivY]        IN   Long型:             縦分割指定
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.MultiMeanByBlock(srcPlane, DivX, DivY, multiMeanFunc, dstColor, srcColor)

End Sub


Public Sub AccumulateRow( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional ByVal accOption As IdpAccumOption = idpAccumSum, Optional ByVal dstCol As Long = 1 _
)
'内容:
'   横方向に演算
'
'[srcPlane]    IN   CImgPlane型:        元プレーン
'[srcZone]     IN   String型:           元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型:     元プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:        結果格納プレーン
'[dstZone]     IN   String型:           結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型:     結果格納プレーンの色指定
'[accOption]   IN   IdpAccumOption型:   演算方法指定(Mean,Sum,StdDeviation)
'[dstCol]      IN   Long型:             結果格納列指定
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.AccumulateRow(srcPlane, accOption, dstCol, dstColor, srcColor)

End Sub

Public Sub AccumulateColumn( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional ByVal accOption As IdpAccumOption = idpAccumSum, Optional ByVal dstRow As Long = 1 _
)
'内容:
'   縦方向に演算
'
'[srcPlane]    IN   CImgPlane型:        元プレーン
'[srcZone]     IN   String型:           元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型:     元プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:        結果格納プレーン
'[dstZone]     IN   String型:           結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型:     結果格納プレーンの色指定
'[accOption]   IN   IdpAccumOption型:   演算方法指定(Mean,Sum,StdDeviation)
'[dstRow]      IN   Long型:             結果格納行指定
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.AccumulateColumn(srcPlane, accOption, dstRow, dstColor, srcColor)

End Sub

Public Sub SubRows( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal diffRows As Long _
)
'内容:
'   指定幅分隣接する行同士を減算
'
'[srcPlane]    IN   CImgPlane型:    元プレーン
'[srcZone]     IN   String型:       元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'[diffRows]    IN   Long型:         幅指定
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.SubRows(srcPlane, diffRows, dstColor, srcColor)

End Sub

Public Sub SubColumns( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal diffCols As Long _
)
'内容:
'   ###このプロシージャの役割などをできるだけ詳しく記述してください###
'
'[srcPlane]    IN   CImgPlane型:    元プレーン
'[srcZone]     IN   String型:       元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'[diffCols]    IN   Long型:         幅指定
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.SubColumns(srcPlane, diffCols, dstColor, srcColor)

End Sub

Public Sub LOr( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional ByVal srcBit As Long = 0, Optional ByVal srcBit2 As Long = 0, Optional ByVal dstBit As Long = 0 _
)
'内容:
'   対象画像同士をOR演算する
'
'[srcPlane]    IN   CImgPlane型:    元プレーン1
'[srcZone]     IN   String型:       元プレーン1のゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーン1の色指定
'[srcPlane2]   IN   CImgPlane型:    元プレーン2
'[srcZone2]    IN   String型:       元プレーン2のゾーン指定
'[srcColor2]   IN   IdpColorType型: 元プレーン2の色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'[srcBit]      IN   Long型:         元プレーン1のビット指定
'[srcBit2]     IN   Long型:         元プレーン2のビット指定
'[dstBit]      IN   Long型:         結果格納プレーンのビット指定
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.LOr(srcPlane, srcPlane2, dstColor, srcColor, srcColor2, dstBit, srcBit, srcBit2)

End Sub

Public Sub LAnd( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional ByVal srcBit As Long = 0, Optional ByVal srcBit2 As Long = 0, Optional ByVal dstBit As Long = 0 _
)
'内容:
'   対象画像同士をAND演算する
'
'[srcPlane]    IN   CImgPlane型:    元プレーン1
'[srcZone]     IN   String型:       元プレーン1のゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーン1の色指定
'[srcPlane2]   IN   CImgPlane型:    元プレーン2
'[srcZone2]    IN   String型:       元プレーン2のゾーン指定
'[srcColor2]   IN   IdpColorType型: 元プレーン2の色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'[srcBit]      IN   Long型:         元プレーン1のビット指定
'[srcBit2]     IN   Long型:         元プレーン2のビット指定
'[dstBit]      IN   Long型:         結果格納プレーンのビット指定
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.LAnd(srcPlane, srcPlane2, dstColor, srcColor, srcColor2, dstBit, srcBit, srcBit2)

End Sub

Public Sub ShiftLeft( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal shiftNum As Long _
)
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.ShiftLeft(srcPlane, shiftNum, dstColor, srcColor)
End Sub

Public Sub ShiftRight( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal shiftNum As Long _
)
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.ShiftRight(srcPlane, shiftNum, dstColor, srcColor)
End Sub

Public Sub LNot( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional ByVal srcBit As Long = 0, Optional ByVal dstBit As Long = 0 _
)
'内容:
'   対象画像をNOT演算する
'
'[srcPlane]    IN   CImgPlane型:    元プレーン
'[srcZone]     IN   String型:       元プレーンのゾーン指定
'[srcColor]    IN   IdpColorType型: 元プレーンの色指定
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[dstZone]     IN   String型:       結果格納プレーンのゾーン指定
'[dstColor]    IN   IdpColorType型: 結果格納プレーンの色指定
'[srcBit]      IN   Long型:         元プレーンのビット指定
'[dstBit]      IN   Long型:         結果格納プレーンのビット指定
'
'備考:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.LNot(srcPlane, dstColor, srcColor, dstBit, srcBit)

End Sub

'=======2009/04/28 Add Maruyama 引数追加　この関数==============
'=======2013/02/15 Add JOB自動化対応 引数順変更==============
Public Sub ReadPixel( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByVal site As Long, _
    ByVal dataNum As Long, ByVal pFlgName As String, _
    ByRef retPixArr() As T_PIXINFO, ByRef AddrMode As IdpAddrMode _
)
'内容:
'   フラグプレーンで指定された画素のデータを読み込む
'
'[site]        IN   Long型:         サイト指定(必須)
'[dataNum]     IN   Long型:         読み込むデータの個数
'[srcPlane]    IN   CImgPlane型:    データ元プレーン
'[srcZone]     IN   String型:       データ元プレーンのゾーン指定
'[pFlgName]    IN   String型:       フラグ名
'[retPixArr()] OUT  T_PIXINFO型:    結果格納用配列
'[AddrMode]    IN   IdpAddrMode型:  アドレスの返し方
'
'備考:
'   retPixArrは不定長配列を指定する。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.PixelLog(site, pFlgName, retPixArr, dataNum, AddrMode)

End Sub
'=======2009/04/28 Add Maruyama ここまで==============

Public Sub EdgeCorrect(ByRef dstPlane As CImgPlane, ByRef dstColor As Variant, ByVal srcZoneArray As Variant, ByVal dstZoneArray As Variant)
'内容:
'   srcZoneArrayで指定されたデータを、dstZoneArrayの対応するゾーンにコピーする
'
'[dstPlane]     IN  対象プレーン
'[dstColor]     IN  対象プレーンの色指定
'[srcZoneArray] IN  コピー元ゾーン配列
'[dstZoneArray] IN  コピー先ゾーン配列
'
'備考:
'
    Dim i As Long
    Dim workPlane As CImgPlane

    Set workPlane = TheIDP.PlaneManager(dstPlane.planeGroup).GetFreePlane(dstPlane.BitDepth)
    For i = 0 To UBound(srcZoneArray)
        Call Copy(dstPlane, srcZoneArray(i), dstColor, workPlane, dstZoneArray(i), dstColor)
        Call Copy(workPlane, dstZoneArray(i), dstColor, dstPlane, dstZoneArray(i), dstColor)
    Next i

End Sub

Public Sub Extention( _
    ByRef pSrcPlane As CImgPlane, ByVal pZone As String, ByRef pDstPlane As CImgPlane, ByVal pExLeft As Long, ByVal pExRight As Long, ByVal pExTop As Long, ByVal pExBottom As Long, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT _
)
'内容:
'   pSrcPlaneのpZoneのデータをpDstPlaneにコピーし、指定した幅の分拡張する。
'
'[pSrcPlane]    IN  データ元のプレーン
'[pZone]        IN  対象のゾーン
'[pDstPlane]    IN  対象のプレーン(pSrcPlaneと同じものも可能)
'[pExLeft]      IN  左側の拡張数
'[pExRight]     IN  右側の拡張数
'[pExTop]       IN  上側の拡張数
'[pExBottom]    IN  下側の拡張数
'[pColor]       IN  色指定
'
'備考:
'   負の値を指定すると、ゾーンの内部で拡張する
'       正:ゾーンの内側のデータを外側にコピー
'       負:ゾーンの内側にさらに内側からコピー
'   pColorにEEE_COLOR_FLAT以外を指定すると、拡張幅×カラーマップの幅に拡張し、指定した色のデータをコピーする。
'
    Dim tmpSrcPMD As CImgPmdInfo
    Dim tmpDstPMD As CImgPmdInfo
    Dim tmpPlane As CImgPlane

    Set tmpPlane = TheIDP.PlaneManager(pSrcPlane.planeGroup).GetFreePlane(pSrcPlane.BitDepth)

    If pColor <> EEE_COLOR_FLAT Then
        With pSrcPlane.planeMap
            pExLeft = pExLeft * .width
            pExRight = pExRight * .width
            pExTop = pExTop * .height
            pExBottom = pExBottom * .height
        End With
    End If

    '元をコピー
    If Not pSrcPlane Is pDstPlane Then
        Call pSrcPlane.SetPMD(pZone)
        Call pDstPlane.SetPMD(pZone)
        Call pDstPlane.CopyPlane(pSrcPlane, pColor)
    End If

    With TheIDP.PMD(pZone)
        '左辺をコピー
        If pExLeft <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr + (Abs(pExLeft) - pExLeft) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, Abs(pExLeft), .height - (Abs(pExTop) - pExTop) / 2 - (Abs(pExBottom) - pExBottom) / 2)
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, pColor)
                Call pDstPlane.SetCustomPMD(.XAdr - Abs(pExLeft), .YAdr, .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane, pColor)
            End With
        End If

        '右辺をコピー
        If pExRight <> 0 Then
            Call tmpPlane.SetCustomPMD(.Right + 1 - Abs(pExRight) - (Abs(pExRight) - pExRight) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, Abs(pExRight), .height - (Abs(pExTop) - pExTop) / 2 - (Abs(pExBottom) - pExBottom) / 2)
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, pColor)
                Call pDstPlane.SetCustomPMD(.XAdr + Abs(pExRight), .YAdr, .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane, pColor)
            End With
        End If

        '上辺をコピー
        If pExTop <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr - (Abs(pExLeft) + pExLeft) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, .width + (Abs(pExLeft) + pExLeft) / 2 + (Abs(pExRight) + pExRight) / 2, Abs(pExTop))
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, pColor)
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr - Abs(pExTop), .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane, pColor)
            End With
        End If

        '下辺をコピー
        If pExBottom <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr - (Abs(pExLeft) + pExLeft) / 2, .Bottom + 1 - Abs(pExBottom) - (Abs(pExBottom) - pExBottom) / 2, .width + (Abs(pExLeft) + pExLeft) / 2 + (Abs(pExRight) + pExRight) / 2, Abs(pExBottom))
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, pColor)
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr + Abs(pExBottom), .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane, pColor)
            End With
        End If
    End With

End Sub

Public Sub ExtentionMirror( _
    ByRef pSrcPlane As CImgPlane, ByVal pZone As String, ByRef pDstPlane As CImgPlane, ByVal pExLeft As Long, ByVal pExRight As Long, ByVal pExTop As Long, ByVal pExBottom As Long _
)
'内容:
'   pSrcPlaneのpZoneのデータをpDstPlaneにコピーし、指定した幅の分拡張する。
'   拡張は鏡像コピーでされる。
'
'[pSrcPlane]    IN  データ元のプレーン
'[pZone]        IN  対象のゾーン
'[pDstPlane]    IN  対象のプレーン(pSrcPlaneと同じものも可能)
'[pExLeft]      IN  左側の拡張数
'[pExRight]     IN  右側の拡張数
'[pExTop]       IN  上側の拡張数
'[pExBottom]    IN  下側の拡張数
'
'備考:
'   負の値を指定すると、ゾーンの内部で拡張する
'       正:ゾーンの内側のデータを外側にコピー
'       負:ゾーンの内側にさらに内側からコピー
'   色の指定はできない。フラットのみ。
'
    Dim tmpSrcPMD As CImgPmdInfo
    Dim tmpDstPMD As CImgPmdInfo
    Dim tmpPlane As CImgPlane

    Set tmpPlane = TheIDP.PlaneManager(pSrcPlane.planeGroup).GetFreePlane(pSrcPlane.BitDepth)

    '元をコピー
    If Not pSrcPlane Is pDstPlane Then
        Call pSrcPlane.SetPMD(pZone)
        Call pDstPlane.SetPMD(pZone)
        Call pDstPlane.CopyPlane(pSrcPlane)
    End If

    With TheIDP.PMD(pZone)
        '左辺をコピー
        If pExLeft <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr + (Abs(pExLeft) - pExLeft) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, Abs(pExLeft), .height - (Abs(pExTop) - pExTop) / 2 - (Abs(pExBottom) - pExBottom) / 2)
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, , , idpCopyMirrorHorizontal)
                Call pDstPlane.SetCustomPMD(.XAdr - Abs(pExLeft), .YAdr, .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane)
            End With
        End If

        '右辺をコピー
        If pExRight <> 0 Then
            Call tmpPlane.SetCustomPMD(.Right + 1 - Abs(pExRight) - (Abs(pExRight) - pExRight) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, Abs(pExRight), .height - (Abs(pExTop) - pExTop) / 2 - (Abs(pExBottom) - pExBottom) / 2)
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, , , idpCopyMirrorHorizontal)
                Call pDstPlane.SetCustomPMD(.XAdr + Abs(pExRight), .YAdr, .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane)
            End With
        End If

        '上辺をコピー
        If pExTop <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr - (Abs(pExLeft) + pExLeft) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, .width + (Abs(pExLeft) + pExLeft) / 2 + (Abs(pExRight) + pExRight) / 2, Abs(pExTop))
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, , , idpCopyMirrorVertical)
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr - Abs(pExTop), .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane)
            End With
        End If

        '下辺をコピー
        If pExBottom <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr - (Abs(pExLeft) + pExLeft) / 2, .Bottom + 1 - Abs(pExBottom) - (Abs(pExBottom) - pExBottom) / 2, .width + (Abs(pExLeft) + pExLeft) / 2 + (Abs(pExRight) + pExRight) / 2, Abs(pExBottom))
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, , , idpCopyMirrorVertical)
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr + Abs(pExBottom), .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane)
            End With
        End If
    End With

End Sub

Private Sub Extention_(ByRef pSrcPlane As CImgPlane, ByRef pDstPlane As CImgPlane, ByRef pWrkPlane As CImgPlane, ByRef pSrcZone As CImgPmdInfo, ByRef pDstZone As CImgPmdInfo)

    Call pSrcPlane.SetPMD(pSrcZone)
    Call pWrkPlane.SetPMD(pSrcZone)
    Call pDstPlane.SetPMD(pDstZone)

    Call pWrkPlane.CopyPlane(pSrcPlane)
    Call pDstPlane.CopyPlane(pWrkPlane)

End Sub

Public Sub GetFreePlane(ByRef pDst As CImgPlane, ByVal pType As String, ByVal pBitDepth As IdpBitDepth, Optional ByVal pClear As Boolean = False, Optional pComment As String = "-")
'内容:
'   プレーンを取得する。
'
'[pDst]         IN  対象のプレーン
'[pType]        IN  マネージャ名指定
'[pBitDepth]    IN  ビット深さ指定
'[pClear]       IN  データを0クリアするかどうか
'[pComment]     IN  コメント
'
'備考:
'

    On Error GoTo flagcheck
    
    Set pDst = Nothing
    Set pDst = TheIDP.PlaneManager(pType).GetFreePlane(pBitDepth, pClear, pComment)
'IDPログに情報出力するため変更    pDst.Comment = pComment

    Exit Sub

'2014/04/02 T.Akasaka
'Flag拡張機能の為、Plane枚数が足りない場合は空きのFlagPlaneを確認する処理を追加
flagcheck:
    Dim ReleaseCount As Long
    
    ReleaseCount = TheIDP.PlaneManager(pType).ReleaseUnusedFlagPlane
    Set pDst = Nothing
    Set pDst = TheIDP.PlaneManager(pType).GetFreePlane(pBitDepth, pClear, pComment)
    
End Sub

'=======2009/04/30 追加 Maruyama ここから==============
Public Function GetFreePlaneForTOPT(ByRef pDst As CImgPlane, ByVal pType As String, ByVal pBitDepth As IdpBitDepth, _
        Optional ByVal pClear As Boolean = False, Optional pComment As String = "-") As Boolean
'内容:
'   プレーンを取得する。
'
'[pDst]         IN  対象のプレーン
'[pType]        IN  マネージャ名指定
'[pBitDepth]    IN  ビット深さ指定
'[pClear]       IN  データを0クリアするかどうか
'[pComment]     IN  コメント
'
'備考:
'
    
    Set pDst = Nothing
    On Error GoTo ErrExit
    Set pDst = TheIDP.PlaneManager(pType).GetFreePlane(pBitDepth, pClear, pComment)
'IDPログに情報出力するため変更    pDst.Comment = pComment
    
    GetFreePlaneForTOPT = True
    Exit Function
    
ErrExit:
    Dim Err As CErrInfo
    Set Err = TheError.LastError
    TheExec.Datalog.WriteComment pType & " : There is no free plane."
    TheExec.Datalog.WriteComment Err.Message
    GetFreePlaneForTOPT = False
    Exit Function

End Function
'=======2009/04/30 追加 Maruyama ここまで==============

Public Sub ReleasePlane(ByRef pDst As CImgPlane)
'内容:
'   プレーンを解放する。
'
'備考:
'
    Set pDst = Nothing
End Sub

Public Sub GetRegisteredPlane(ByVal pName As String, ByRef pDst As CImgPlane, Optional pComment As String, Optional IsDelete As Boolean = False)
'内容:
'   PlaneBankに登録されたプレーンを取得する。
'
'[pName]        IN  登録した名前を指定
'[pDst]         IN  対象のプレーン
'
'備考:
'
    Set pDst = Nothing
    Set pDst = TheIDP.PlaneBank.Item(pName)
    
'=======2009/04/28 Add Maruyama ここから==============
    pDst.Comment = pComment
        
    If IsDelete Then
        TheIDP.PlaneBank.Delete (pName)
        pDst.ReadOnly = False
    End If
'=======2009/04/28 Add Maruyama ここまで==============
End Sub

'=======2009/04/30 Add Maruyama ここから==============
Public Sub SharedFlagNot(ByRef pPlaneGroup As Variant, ByVal pZone As Variant, _
        ByVal pDstName As String, ByVal pSrcName As String, _
        Optional ByRef pColor As Variant = EEE_COLOR_ALL)
'内容:
'   SharedFlagのFlagを反転する
'
'[pPlaneGroup]         IN Variant型　　　　対象のプレーン,タイプ
'[pZone]        IN Variant型:       対象プレーンのゾーン指定
'[pSrcName]     IN String型:        データ元の名前
'[pDstName]     IN String型:        結果の名前
'[pColor]       IN IdpColorType型:  色指定
'
'備考:
'   pSrcNameとpDstNameが同一でも可能。
'

    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pDstName)
        Call .SetPMD(pZone)
        Call .LNot(pDstName, pSrcName, pColor)
    End With
    
End Sub

Public Sub SharedFlagOr(ByRef pPlaneGroup As Variant, ByVal pZone As Variant, ByVal pDstName As String, _
        ByVal pSrcName1 As String, ByVal pSrcName2 As String, _
        Optional ByRef pColor As Variant = EEE_COLOR_ALL)
'内容:
'   pSrcName1のビットとpSrcName2のビットのOr演算の結果をpDstNameのビットに入れる。
'   pDstNameが登録されていない場合、新たに登録する。
'   既に登録されている場合は、そのビットに入れる。
'
'[pPlaneGroup]         IN Variant型　　　　対象のプレーン,タイプ
'[pZone]        IN Variant型:       対象プレーンのゾーン指定
'[pSrcName1]    IN String型:        データ元の名前1
'[pSrcName2]    IN String型:        データ元の名前2
'[pDstName]     IN String型:        結果の名前
'[pColor]       IN IdpColorType型:  色指定
'
'備考:
'   pSrcName1とpSrcName2とpDstNameが同一でも可能。
'
    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pDstName)
        Call .SetPMD(pZone)
        Call .LOr(pDstName, pSrcName1, pSrcName2, pColor)
    End With

End Sub

Public Sub SharedFlagAnd(ByRef pPlaneGroup As Variant, ByVal pZone As Variant, ByVal pDstName As String, _
        ByVal pSrcName1 As String, ByVal pSrcName2 As String, _
        Optional ByRef pColor As Variant = EEE_COLOR_ALL)
'内容:
'   pSrcName1のビットとpSrcName2のビットのAnd演算の結果をpDstNameのビットに入れる。
'   pDstNameが登録されていない場合、新たに登録する。
'   既に登録されている場合は、そのビットに入れる。
'
'[pPlaneGroup]         IN Variant型　　　　対象のプレーン,タイプ
'[pZone]        IN Variant型:       対象プレーンのゾーン指定
'[pSrcName1]    IN String型:        データ元の名前1
'[pSrcName2]    IN String型:        データ元の名前2
'[pDstName]     IN String型:        結果の名前
'[pColor]       IN IdpColorType型:  色指定
'
'備考:
'   pSrcName1とpSrcName2とpDstNameが同一でも可能。

    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pDstName)
        Call .SetPMD(pZone)
        Call .LAnd(pDstName, pSrcName1, pSrcName2, pColor)
    End With

End Sub

Public Sub SharedFlagXor(ByRef pPlaneGroup As Variant, ByVal pZone As Variant, ByVal pDstName As String, _
        ByVal pSrcName1 As String, ByVal pSrcName2 As String, _
        Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'内容:
'   pSrcName1のビットとpSrcName2のビットのXOr演算の結果をpDstNameのビットに入れる。
'   pDstNameが登録されていない場合、新たに登録する。
'   既に登録されている場合は、そのビットに入れる。
'
'[pPlaneGroup]         IN Variant型　　　　対象のプレーン,タイプ
'[pZone]        IN Variant型:       対象プレーンのゾーン指定
'[pSrcName1]    IN String型:        データ元の名前1
'[pSrcName2]    IN String型:        データ元の名前2
'[pDstName]     IN String型:        結果の名前
'[pColor]       IN IdpColorType型:  色指定
'
'備考:
'   pSrcName1とpSrcName2とpDstNameが同一でも可能。

    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pDstName)
        Call .SetPMD(pZone)
        Call .LXor(pDstName, pSrcName1, pSrcName2, pColor)
    End With

End Sub

Public Sub FlagNot(ByRef pDst As CImgPlane, ByVal pZone As Variant, ByVal pFlgName As String, _
    Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'内容:
'   共用フラグプレーンのNot演算の結果を取得
'
'[pDst]         IN   CImgPlane型　　　　対象のプレーン
'[pZone]        IN   Variant型:       対象プレーンのゾーン指定
'[pFlgName]     IN   String型:          データ元のフラグ名
'[pDstBit]      IN   Long型:            保存先ビット指定
'[pColor]       IN   IdpColorType型:    色指定
'
'備考:
    With pDst
        Call .SetPMD(pZone)
        Call .FlagNot(pFlgName, pDstBit, pColor)
    End With
End Sub

Public Sub FlagOr(ByRef pDst As CImgPlane, ByVal pZone As Variant, ByVal pFlgName1 As String, ByVal pFlgName2 As String, _
    Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'内容:
'   共用フラグプレーンのOr演算の結果を取得
'
'[pDst]         IN   CImgPlane型　　　　対象のプレーン
'[pZone]        IN   Variant型:         対象プレーンのゾーン指定
'[pFlgName1]    IN   String型:          データ元1のフラグ名
'[pFlgName2]    IN   String型:          データ元2のフラグ名
'[pDstBit]      IN   Long型:            保存先ビット指定
'[pColor]       IN   IdpColorType型:    色指定
'
'備考:
'   pFlgName1、pFlgName2は同名の指定可能。
    With pDst
        Call .SetPMD(pZone)
        Call .FlagOr(pFlgName1, pFlgName2, pDstBit, pColor)
    End With

End Sub

Public Sub FlagAnd(ByRef pDst As CImgPlane, ByVal pZone As Variant, ByVal pFlgName1 As String, ByVal pFlgName2 As String, _
    Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'内容:
'   共用フラグプレーンのAnd演算の結果を取得
'
'[pDst]         IN   CImgPlane型　　　　対象のプレーン
'[pZone]        IN   Variant型:       対象プレーンのゾーン指定
'[pFlgName1]    IN   String型:          データ元1のフラグ名
'[pFlgName2]    IN   String型:          データ元2のフラグ名
'[pDstBit]      IN   Long型:            保存先ビット指定
'[pColor]       IN   IdpColorType型:    色指定
'
'備考:
'   pFlgName1、pFlgName2は同名の指定可能。
    With pDst
        Call .SetPMD(pZone)
        Call .FlagAnd(pFlgName1, pFlgName2, pDstBit, pColor)
    End With
End Sub

Public Sub FlagXor(ByRef pDst As CImgPlane, ByVal pZone As Variant, ByVal pFlgName1 As String, ByVal pFlgName2 As String, _
    Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'内容:
'   共用フラグプレーンのXor演算の結果を取得
'
'[pDst]         IN   CImgPlane型　　　　対象のプレーン
'[pZone]        IN   Variant型:         対象プレーンのゾーン指定
'[pFlgName1]    IN   String型:          データ元1のフラグ名
'[pFlgName2]    IN   String型:          データ元2のフラグ名
'[pDstBit]      IN   Long型:            保存先ビット指定
'[pColor]       IN   IdpColorType型:    色指定
'
'備考:
'   pFlgName1、pFlgName2は同名の指定可能。
    With pDst
        Call .SetPMD(pZone)
        Call .FlagXor(pFlgName1, pFlgName2, pDstBit, pColor)
    End With

End Sub

Public Sub FlagCopy(ByRef pDst As CImgPlane, ByVal pZone As Variant, ByVal pFlgName As String, _
    Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_ALL)
'内容:
'   共用フラグプレーンの指定したビットの値をコピー
'
'[pDst]         IN   CImgPlane型　　　　対象のプレーン
'[pZone]        IN   Variant型:         対象プレーンのゾーン指定
'[pFlgName]     IN   String型:          データ元のフラグ名
'[pDstBit]      IN   Long型:            保存先ビット指定
'[pColor]       IN   IdpColorType型:    色指定
    With pDst
        Call .SetPMD(pZone)
        Call .FlagCopy(pFlgName, pDstBit, pColor)
    End With

End Sub
'=======2009/04/30 Add Maruyama ここまで==============

'=======2009/06/03 Add Maruyama ここから==============
'藤後さん作成
Public Function IsFlgExist( _
        ByRef pPlaneGroup As Variant, ByVal pFlgName As String) As Boolean

    Dim Bit As Long
    
    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pFlgName)
        Bit = .FlagBit(pFlgName)
        If Bit = 0 Then
            IsFlgExist = False
            Exit Function
        End If
    End With
    IsFlgExist = True
    
End Function
'=======2009/06/03 Add Maruyama ここまで==============

'2009/09/15 D.Maruyama ColorAll処理追加 ここから
Public Sub AverageColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'内容:
'   平均値を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[retResult]   OUT  CImgColorAllResult型:       結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.AverageColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub SumColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'内容:
'   合計値を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[retResult]   OUT  CImgColorAllResult型:       結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.SumColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub StdDevColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'内容:
'   標準偏差を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[retResult]   OUT  CImgColorAllResult型:       結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.StdDevColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub GetPixelCountColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'内容:
'   対象の画素数を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[retResult]   OUT  CImgColorAllResult型:       結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.NumColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub MinColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'内容:
'   最小値を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[retResult]   OUT  CImgColorAllResult型:       結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.MinColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub MaxColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'内容:
'   最大値を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[retResult]   OUT  CImgColorAllResult型:       結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.MaxColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub MinMaxColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retMin As CImgColorAllResult, ByRef retMax As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'内容:
'   最小値、最大値を一度に取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[retMin]      OUT  CImgColorAllResult型:       最小値格納用配列(サイト分)
'[retMax]      OUT  CImgColorAllResult型:       最大値格納用配列(サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'
    Dim tempResutMin As ColorAllResult
    Dim tempResutMax As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.MinMaxColorAll(tempResutMin, tempResutMax, pFlgName)

    If retMax Is Nothing Then
        Set retMax = New CImgColorAllResult
    End If
    
    Call retMax.SetParam(srcPlane, tempResutMax)

    If retMin Is Nothing Then
        Set retMin = New CImgColorAllResult
    End If
    
    Call retMin.SetParam(srcPlane, tempResutMin)

End Sub

Public Sub DiffMinMaxColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'内容:
'   最小値と最大値の差を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[retResult]   OUT  CCImgColorAllResult型:       結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.DiffMinMaxColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub AbsMaxColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'内容:
'   最小値と最大値の内絶対値の大きい方を取得
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[retResult]   OUT  CImgColorAllResult型:       結果格納用構造体(色別サイト分)
'[pFlgName]    IN   String型:       フラグ名
'
'備考:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.AbsMaxColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub


Public Sub CountColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal pCountLimMode As IdpCountLimitMode, _
    ByVal limitType As IdpLimitType, _
    ByRef retResult As CImgColorAllResult, _
    Optional ByVal pFlgName As String = "", Optional ByVal pInputFlgName As String = "" _
)
'内容:
'   条件に該当する点の個数を取得する。
'
'[srcPlane]         IN   CImgPlane型:    対象プレーン
'[srcZone]          IN   String型:       対象プレーンのゾーン指定
'[countType]        IN   IdpCountType型: カウント条件指定
'[loLim]            IN   Variant型:      下限値
'[hiLim]            IN   Variant型:      上限値
'[pCountLimMode]    IN   IdpCountLimitMode型: 境界値のとりかた（サイト別？サイト色別？）
'[limitType]        IN   IdpLimitType型: 境界値を含む、含まない指定
'[retResult]        OUT  CImgColorAllResult型:       結果格納用構造体(色別サイト分)
'[pFlgName]         IN   String型:       出力フラグ名
'[pInputFlgName]    IN   String型:       入力フラグ名
'
'備考:
'   hiLim,loLimをVariantで定義しているのは、サイト別で境界値が違う場合に対応するため。
'   サイト配列を入れると各サイトごと別々に、定数を入れると全サイトに同じ値が適用される。
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.CountColorAll(tempResut, countType, loLim, hiLim, pCountLimMode, limitType, pFlgName, pInputFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub CountColorAllForFlgBitImgPlane( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal pCountLimMode As IdpCountLimitMode, _
    ByVal limitType As IdpLimitType, ByRef retResult As CImgColorAllResult, _
    ByRef pFlgPlane As CImgPlane, ByVal pFlgBit As Long, _
    Optional ByVal pInputFlgName As String = "" _
)
'内容:
'   条件に該当する点の個数を取得する。(フラグビットをイメージプレンに立てる)
'
'[srcPlane]         IN   CImgPlane型:    対象プレーン
'[srcZone]          IN   String型:       対象プレーンのゾーン指定
'[countType]        IN   IdpCountType型: カウント条件指定
'[loLim]            IN   Variant型:      下限値
'[hiLim]            IN   Variant型:      上限値
'[pCountLimMode]    IN   IdpCountLimitMode型: 境界値のとりかた（サイト別？サイト色別？）
'[limitType]        IN   IdpLimitType型: 境界値を含む、含まない指定
'[retResult]        OUT  CImgColorAllResult型:       結果格納用構造体(色別サイト分)
'[pFlgName]         IN   CImgPlane型:       出力フラグ名
'[pFlgBit]        　IN   Long型:       出力フラグ名
'[pInputFlgName]    IN   String型:       入力フラグ名
'
'備考:
'   hiLim,loLimをVariantで定義しているのは、サイト別で境界値が違う場合に対応するため。
'   サイト配列を入れると各サイトごと別々に、定数を入れると全サイトに同じ値が適用される。
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call pFlgPlane.SetPMD(srcZone)
    Call srcPlane.CountColorAllForFlgBitImgPlane(tempResut, countType, loLim, hiLim, pCountLimMode, pFlgPlane, pFlgBit, limitType, pInputFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub PutFlagColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal pCountLimMode As IdpCountLimitMode, _
    ByVal limitType As IdpLimitType, _
    ByVal pFlgName As String, Optional ByVal pInputFlgName As String _
)
'内容:
'   条件に該当する点にフラグを立てる。(Countからフラグを立てることだけに特化)
'
'[srcPlane]    IN   CImgPlane型:    対象プレーン
'[srcZone]     IN   String型:       対象プレーンのゾーン指定
'[countType]   IN   IdpCountType型: カウント条件指定
'[loLim]       IN   Variant型:      下限値
'[hiLim]       IN   Variant型:      上限値
'[pCountLimMode]    IN   IdpCountLimitMode型: 境界値のとりかた（サイト別？サイト色別？）
'[limitType]   IN   IdpLimitType型: 境界値を含む、含まない指定
'[pFlgName]         IN   String型:       出力フラグ名
'[pInputFlgName]    IN   String型:       入力フラグ名
'
'備考:
'   hiLim,loLimをVariantで定義しているのは、サイト別で境界値が違う場合に対応するため。
'   サイト配列を入れると各サイトごと別々に、定数を入れると全サイトに同じ値が適用される。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.PutFlagColorAll(pFlgName, countType, loLim, hiLim, pCountLimMode, limitType, pInputFlgName)

End Sub
'2009/09/15 D.Maruyama ColorAll処理追加 ここまで

'以下はモジュール内でのみの使用 ##################################################################################################################################
Private Sub SetOptionalPMD(ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByVal refZone As Variant)

    If dstPlane Is Nothing Then Exit Sub

    If IsEmpty(dstZone) Then
        Call dstPlane.SetPMD(refZone)
    Else
        Call dstPlane.SetPMD(dstZone)
    End If

End Sub

Private Function Var2PlaneNameFlag(ByVal pVal As Variant) As String
    
    If IsObject(pVal) Then
        Var2PlaneNameFlag = pVal.Manager.Name
    Else
        Var2PlaneNameFlag = pVal
    End If
End Function



