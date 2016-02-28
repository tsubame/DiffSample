VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVISVISource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   IP750 電源Board I/F
'
'目的:
'   各種電源ボードの制御方法の統一
'
'作成者:
'   SLSI今手
'
'
'Code Checked
'Comment Checked
'

Option Explicit

Private Const ALL_SITE = -1

Public Property Get BoardName() As String
'内容:
'   IP750ボードの名前
'
'パラメータ:
'
'戻り値:
'   IP750ボード名
'
'注意事項:
'   Read Only
'
End Property

Public Property Let PinList(ByVal PinList As String)
'内容:
'   制御対象ピンリスト
'
'パラメータ:
'    [PinList]   In  制御対象ピンリスト。
'
'戻り値:
'
'注意事項:
'
End Property

Public Property Get PinList() As String
'内容:
'   制御対象ピンリスト
'
'パラメータ:
'
'戻り値:
'   制御対象ピンリスト
'
'注意事項:
'
End Property

Public Sub SetFVMI(ByVal ForceV As Variant, _
    ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True)
'内容:
'   PinListに設定されたピンを電圧印加状態に設定する｡
'
'パラメータ:
'    [ForceV]      In   印加電圧。配列指定可能。
'    [ClampI]      In   クランプ電流値。
'    [Site]        In   サイト指定。オプション(Default -1)
'    [ConnectOn]   In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンを電圧印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceVで印加電圧を指定。ForceVは数値 or サイト数分の配列。
'    ■数値の場合全サイト同じ値を設定｡配列の場合各サイトに対応する配列要素の値を設定｡
'    ■ClampIでクランプ電流を設定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■電圧印加レンジはForceV､電流測定レンジはClampIから判定して設定する｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
End Sub
'V21-Release
Public Sub SetFVMIMulti(ByVal ForceV As Variant, _
    ByVal ClampI As Double, _
    Optional ByVal ConnectOn As Boolean = True)
'内容:
'   PinListに設定されたピンを電圧印加状態に設定する｡(Site同時)
'
'パラメータ:
'    [ForceV]      In   印加電圧。配列指定可能。
'    [ClampI]      In   クランプ電流値。
'    [ConnectOn]   In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンを電圧印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceVで印加電圧を指定。ForceVは数値 or サイト数分の配列。
'    ■数値の場合全サイト同じ値を設定｡配列の場合各サイトに対応する配列要素の値を設定｡
'    ■ClampIでクランプ電流を設定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■電圧印加レンジはForceV､電流測定レンジはClampIから判定して設定する｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
End Sub
Public Sub SetFIMV(ByVal ForceI As Variant, _
    ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True)
'内容:
'   PinListに設定されたピンを電流印加状態に設定する｡
'
'パラメータ:
'    [ForceI]       In    印加電流。配列指定可能。
'    [ClampV]       In    クランプ電圧値。
'    [Site]         In    サイト指定。オプション(Default -1)
'    [ConnectOn]    In    デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンを電流印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceIで印加電流を指定。ForceVは数値 or サイト数分の配列。
'    ■数値の場合全サイト同じ値を設定｡配列の場合各サイトに対応する配列要素の値を設定｡
'    ■ClampVでクランプ電圧を設定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■電流印加レンジはForceI､電圧測定レンジはClampVから判定して設定する｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
End Sub
'V21-Release
Public Sub SetFIMVMulti(ByVal ForceI As Variant, _
    ByVal ClampV As Double, _
    Optional ByVal ConnectOn As Boolean = True)
'内容:
'   PinListに設定されたピンを電流印加状態に設定する｡(Site同時)
'
'パラメータ:
'    [ForceI]       In    印加電流。配列指定可能。
'    [ClampV]       In    クランプ電圧値。
'    [ConnectOn]    In    デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンを電流印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceIで印加電流を指定。ForceVは数値 or サイト数分の配列。
'    ■数値の場合全サイト同じ値を設定｡配列の場合各サイトに対応する配列要素の値を設定｡
'    ■ClampVでクランプ電圧を設定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■電流印加レンジはForceI､電圧測定レンジはClampVから判定して設定する｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
End Sub
Public Sub SetMV(ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True)
'内容:
'   PinListに設定されたピンを電圧測定状態に設定する。
'
'パラメータ:
'    [ClampV]       In    クランプ電圧値。
'    [Site]         In    サイト指定。オプション(Default -1)
'    [ConnectOn]    In    デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampVでクランプ電圧を設定｡(APMUリソースのみクランプ機能有効)
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■電圧測定レンジはClampVから判定して設定する｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
End Sub
'V21-Release
Public Sub SetMVMulti(ByVal ClampV As Double, _
    Optional ByVal ConnectOn As Boolean = True)
'内容:
'   PinListに設定されたピンを電圧測定状態に設定する。
'
'パラメータ:
'    [ClampV]       In    クランプ電圧値。
'    [ConnectOn]    In    デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampVでクランプ電圧を設定｡(APMUリソースのみクランプ機能有効)
'    ■非アクティブサイトに対しては何もしない｡
'    ■電圧測定レンジはClampVから判定して設定する｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
End Sub
Public Sub ConnectPins(Optional ByVal site As Long = ALL_SITE)
'内容:
'   PinListに設定されたピンをデバイスに接続する｡
'
'パラメータ:
'    [Site]    In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'
End Sub

Public Sub DisconnectPins(Optional ByVal site As Long = ALL_SITE)
'内容:
'   PinListに設定されたピンをデバイスから切り離す｡
'
'パラメータ:
'    [Site]    In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■切り離した後、0V印加状態に設定する(接続はしない)。
'
End Sub
'V21-Release
Public Sub DisconnectPinsMulti()
'内容:
'   PinListに設定されたピンをデバイスから切り離す｡(Site同測)
'
'パラメータ:
'
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■切り離した後、0V印加状態に設定する(接続はしない)。
'
End Sub
Public Sub SetGND(Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True)
'内容:
'   PinListに設定されたピンを接地(0V印加)状態にする
'
'パラメータ:
'    [Site]         In    サイト指定。オプション(Default -1)
'    [ConnectOn]    In    デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンを接地する｡(0V印加状態)
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
End Sub

Public Sub ChangeMIRange(ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE)
'内容:
'   PinListに設定されたピンの電流測定レンジを変更する｡
'
'パラメータ:
'    [ClampI]    In    クランプ電流値。
'    [Site]      In    サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンの印加電圧を変えず､電流測定レンジを変更する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampIでクランプ電流を設定｡
'    ■電流測定レンジはClampIから判定して設定する｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'
End Sub
'V21-Release
Public Sub ChangeMIRangeMulti(ByVal ClampI As Double)
'内容:
'   PinListに設定されたピンの電流測定レンジを変更する｡
'
'パラメータ:
'    [ClampI]    In    クランプ電流値。

'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンの印加電圧を変えず､電流測定レンジを変更する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampIでクランプ電流を設定｡
'    ■電流測定レンジはClampIから判定して設定する｡
'    ■非アクティブサイトに対しては何もしない｡
'
End Sub

Public Sub MeasureV(ByRef retResult() As Double, _
    ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE)
'内容:
'   PinListに設定されたピンの電圧値を測定する｡
'
'パラメータ:
'    [retResult()]   Out   結果を格納する配列。(サイト数分の要素が必要)
'    [AvgNum]        In    平均回数。
'    [Site]          In    サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンの電圧値を測定する｡
'    ■PinListは単独ピンのみ。複数ピン、ピングループ(ギャングピンを除く)の指定は不可。
'    ■retResultに結果を返す｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■Siteが-1以外の場合、対象サイト以外のretResult()の値は変化させない。
'    ■非アクティブサイトの結果は0を返す。Siteが-1の時も同様。
'
End Sub

Public Sub MeasureI(ByRef retResult() As Double, _
    ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE)
'内容:
'   PinListに設定されたピンの電流値を測定する｡
'
'パラメータ:
'    [retResult()]    Out   結果を格納する配列。(サイト数分の要素が必要)
'    [AvgNum]         In    平均回数。
'    [Site]           In    サイト指定。オプション(Default -1)
'
'戻り値:
'
'詳細動作
'    ■PinListに対応するピンの電流値を測定する｡
'    ■PinListは単独ピンのみ。複数ピン、ピングループ(ギャングピンを除く)の指定は不可。
'    ■retResultに結果を返す｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■Siteが-1以外の場合、対象サイト以外のretResult()の値は変化させない。
'    ■非アクティブサイトの結果は0を返す。Siteが-1の時も同様。
'
End Sub

Public Sub MeasureIMulti(ByVal avgNum As Long)
'内容:
'   PinListに設定されたピンの電流値を測定する｡
'
'パラメータ:
'    [AvgNum]    In   平均回数。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンの電流値を測定する｡結果はReadMIMultiResultを使って読み出す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ギャングピンは対応しない｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■全サイトが対象になる｡非アクティブサイトの結果は0になる｡
'
End Sub

Public Sub MeasureVMulti(ByVal avgNum As Long)
'内容:
'   PinListに設定されたピンの電圧値を測定する｡
'
'パラメータ:
'    [AvgNum]    In    平均回数。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンの電圧値を測定する｡結果はReadMVMultiResultを使って読み出す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ギャングピンは対応しない｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■全サイトが対象になる｡非アクティブサイトの結果は0になる｡
'
End Sub

Public Sub ReadMVMultiResult(ByRef retResult() As Double)
'内容:
'   MeasureVMultiで測定した電圧値を読み出す｡
'
'パラメータ:
'    [retResult()]    Out   結果を格納する配列。(サイト数分の要素が必要)
'
'戻り値:
'
'注意事項:
'    ■PinListは最後に使ったMeasureVMultiのピンリストに含まれるピン名を指定｡
'    ■retResultに結果を返す｡
'
End Sub

Public Sub ReadMIMultiResult(ByRef retResult() As Double)
'内容:
'   MeasureIMultiで測定した電流値を読み出す｡
'
'パラメータ:
'    [retResult()]    Out   結果を格納する配列。(サイト数分の要素が必要)
'
'戻り値:
'
'注意事項:
'    ■PinListは最後に使ったMeasureIMultiのピンリストに含まれるピン名を指定｡
'    ■retResultに結果を返す｡
'
End Sub

Public Function GetMeasureRange() As Collection
'内容:
'   PinListに設定されたピンのメジャーレンジを取得
'
'パラメータ:
'
'戻り値:
'   メジャーレンジ取得結果
'
'注意事項:
'結果の取り出し方:
'   取得結果コレクションより、
'   PinListのピン名をキーに、ピン毎の
'   測定レンジ値を取り出すことができます
'
'MIモード時のGANG(MEARGE)接続の対応:
'   APMU/HDVISリソースに関して、GANG(MEARGE)接続
'   を行うためのPinGPをPinListに指定した時、
'   MIモードの場合は、並列接続を考慮した
'   電流レンジ値を返します。
'   (DPSリソースのGANG接続には対応していません)
'
'GangGpピンの取得例:
'   FVMIモードに設定された3CHのGangGP(PowerTypeのPinGp)である
'   GANG_APMU（G_APMU1,G_APMU2,G_APMU3）を指定すると
'
'   取得命令 TheDC.Pins("GANG_APMU").GetMeasureRange
'   取得結果 150[mA] (IRange50[mA]×並列接続数3)
'
'GANG接続GPのマルチ指定:
'   GANG接続用のマルチPinGP指定は対応しません｡
'   GANG接続用のPinGPを複数同時にPinListに指定すると
'   エラーとなります
'
'   (エラー例)
'    × TheDC.Pins("GANG_ApmuGp1,GANG_ApmuGp2).GetMeasureRange
'
'FAILサイトがある場合:
'   存在するサイトのうち1つでもActiveでないサイトが存在する
'   状態で命令を実行したときはエラーとなります。
'   存在サイトがすべてActiveな状態で実行してください
'
'IG-XLシミュレータ環境での実行:
'   IG-XLシミュレータ環境においてはBPMUピンのレンジは取得できません。
'   実行するとエラーメッセージ出力後、取得結果はNothingとなります

End Function

Public Function GetForceVLimit() As Double()
'内容:
'   電圧印加の時に許される最大値と最小値を取得
'
'パラメータ:
'
'戻り値:
'   電圧印加時の設定可能な最大値、最小値
'   要素数0-->最小値、要素数1-->最大値
'
'注意事項:
'
'
End Function

Public Function GetForceILimit() As Double()
'内容:
'   電流印加の時に許される最大値と最小値を取得
'
'パラメータ:
'
'戻り値:
'   電流印加時の設定可能な最大値、最小値
'   要素数0-->最小値、要素数1-->最大値
'
'注意事項:
'GANG（MEARGE）接続のサポート:
'   APMU/HDVISリソースに関して、GANG(MEARGE)接続
'   を行うためのPinGPをPinListに指定した時は、
'   並列接続数を考慮したForce電流値を返します。
'
'GangGpピンの取得例:
'   3CHのGangGP(PowerTypeのPinGp)である
'   GANG_APMU（G_APMU1,G_APMU2,G_APMU3）を指定すると
'
'   取得命令 TheDC.Pins("GANG_APMU").GetForceILimit
'   取得結果 150[mA] (50[mA]×並列接続数3)
'
'GANG接続GPのマルチ指定:
'   GANG接続用のマルチPinGP指定は対応しません｡
'   GANG接続用のPinGPを複数同時にPinListに指定すると
'   エラーとなります
'
'   (エラー例)
'   × TheDC.Pins("GANG_ApmuGp1,GANG_ApmuGp2).GetForceILimit
'
'DPSリソース:
'   DPSリソースはFIモードが存在しないため
'   命令を実行するとエラーとなります
'
End Function

Public Function GetForceValue() As Collection
'内容:
'   PinListに設定されたピンのForce値を取得
'
'パラメータ:
'
'戻り値:
'   Force値、取得結果
'
'注意事項:
'結果の取り出し方:
'   取得結果コレクションより、
'   PinListのピン名をキーに、ピン毎の
'   Force値を取り出すことができます
'
'GANG接続GPのマルチ指定:
'   GANG接続用のマルチPinGP指定は対応しません｡
'   GANG接続用のPinGPを複数同時にPinListに指定すると
'   エラーとなります
'
'   (エラー例)
'    × TheDC.Pins("GANG_ApmuGp1,GANG_ApmuGp2).GetMeasureRange
'
'FAILサイトがある場合:
'   存在するサイトのうち1つでもActiveでないサイトが存在する
'   状態で実行したときはエラーとなります。
'   存在サイトがすべてActiveな状態で実行してください
'
'DPSリソースの取得:
'   DPSのForce値、取得結果は常に「dpsPrimaryVoltage」の値となります
'
'IG-XLシミュレータ環境での実行:
'   IG-XLシミュレータ環境においてはBPMUピンのForce値は取得できません。
'   実行するとエラーメッセージ出力後、取得結果はNothingとなります
'
End Function

Public Property Get ForceIRangeList() As String
'内容:
'   PinListに設定されたピンのリソースで
'   設定可能なForce電流レンジリストを返す
'
'パラメータ:
'
'戻り値:
'   カンマ区切り形式の電流レンジリスト
'   (例) 40uA,200uA,1mA,5mA,50mA
'   ＊レンジパラメータが存在しないときは"None"となります。
'   ＊APMUのGANG対象ピンは50mAを返します
'
'注意事項:
'   DPSにはFIモードは存在しません
'   Auto、Smartは出力しません
'
End Property

Public Property Get ForceVRangeList() As String
'内容:
'   PinListに設定されたピンのリソースで
'   設定可能なForce電圧レンジリストを返す
'
'パラメータ:
'
'戻り値:
'   カンマ区切り形式の電圧レンジリスト
'   (例) 2V,5V,10V,35V
'   ＊レンジパラメータが存在しないときは"None"となります。
'
'注意事項:
'   DPS、PPMUには電圧レンジのパラメータは存在しません。
'   APMUの40Vは出力しません
'   Auto、Smartは出力しません
'
End Property

Public Property Get MeasIRangeList() As String
'内容:
'   PinListに設定されたピンのリソースで
'   設定可能なMeasure電流レンジリストを返す
'
'パラメータ:
'
'戻り値:
'   カンマ区切り形式の電流レンジリスト
'   (例) 200nA,2uA,10uA,40uA,200uA,1mA,5mA,50mA
'   ＊レンジパラメータが存在しないときは"None"となります。
'   ＊APMUのGANG対象ピンは50mAを返します
'
'注意事項:
'   BPMUの2uAは出力しません
'   Auto、Smartは出力しません
'
End Property

Public Property Get MeasVRangeList() As String
'内容:
'   PinListに設定されたピンリソースで
'   設定可能なMeasure電圧レンジリストを返す
'
'パラメータ:
'
'戻り値:
'   カンマ区切り形式の電圧レンジリスト
'   (例) 2V,5V,10V,35V
'   ＊レンジパラメータが存在しないときは"None"となります。
'
'注意事項:
'   APMUの40Vは出力しません
'   Auto、Smartは出力しません
'
End Property

