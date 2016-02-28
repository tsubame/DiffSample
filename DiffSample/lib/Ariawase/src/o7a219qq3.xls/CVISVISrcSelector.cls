VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVISVISrcSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@@@IP750EX Support
'概要:
'   ChannelType情報、ボード指定により対応する
'   IP750の電源設定OBJを選択させる
'
'目的:
'   ChannelType情報を元に対応するDCリソースを選択
'   リソース指定による選択
'   使用するDCリソースのリストを保持
'
'作成者:
'   SLSI今手
'
'注意事項:
'(対応リソース)
'   APMU、DPS、PPMU、BPMU、HDVIS、HSD200-PPMU、ICUL1G－PPMU
'
'   HDVIS利用時には、条件付コンパイル引数に
'   HDVIS_USE=1の記述が必要です。
'
'   HSD200利用時には、条件付コンパイル引数に
'   HSD200_USE=1の記述が必要です。
'
'   ICUL1G利用時には、条件付コンパイル引数に
'   ICUL1G_USE=1の記述が必要です。
'

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const HDVIS_USE = 0      'HVDISリソースの使用　 0：未使用、0以外：使用  <PALSとEeeAuto両方で使用>
#Const ICUL1G_USE = 1               '1CUL1Gボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>
#Const HSD200_USE = 1               'HSD200ボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

'セレクト対応するボード定義
'＊新しい電源ボードへの対応時にはココに情報追加の必要がある。
Public Enum DC_BOARD_NAME
    CHAN_TYPE_SELECT = -1 'ChannelType情報によるBoard選択指定用
    dcAPMU = 0
    dcDPS = 1
    dcPPMU = 2
    dcBPMU = 3
    
    #If HDVIS_USE <> 0 Then
    dcHDVIS = 4
    #End If
    
    #If ICUL1G_USE <> 0 Then
    dcICUL1G = 5
    #End If
End Enum

'DCボードObject格納用コレクション
Private m_BoardList As Collection

'#Pass-Ok
Public Property Get BoardCount() As Long
'内容:
'   ボードリストに登録されているボードの数取得
'
'パラメータ:
'
'戻り値:
'   ボードリストに登録されているボード数
'
'注意事項:
'
    BoardCount = m_BoardList.Count

End Property

'#Pass-Ok
Private Sub InitBoardList()
'内容:
'   BoardListの初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Set m_BoardList = Nothing 'Pass
    Set m_BoardList = New Collection 'Pass
    
End Sub

'#Pass-Ok
Private Sub AddBoard(dcboard As IVISVISource)
'内容:
'   BoardListコレクションに、電源設定用Objectを登録
'
'パラメータ:
'   [dcboard]   In  追加を行う電源設定用Object
'
'戻り値:
'
'注意事項:
'   登録できるのは、IVISVISourceを実装しているOBJです
'
    
    Dim boardchk As IVISVISource
            
    If m_BoardList.Count <> 0 Then
        '登録を指示されたボードがすでにボードListに存在するかを確認
        Set boardchk = Nothing
        On Error Resume Next
        Set boardchk = m_BoardList.Item(dcboard.BoardName)
        On Error GoTo 0
    
        If boardchk Is Nothing Then
            '新規のボードが確認されたので登録
            Call m_BoardList.Add(dcboard, dcboard.BoardName)
        Else
            'すでに存在するボードなのでエラーとする
            Call OutputErrMsg(dcboard.BoardName & " board is exist (at VI Source Selector)")
        End If
    Else
        '登録されているボードがまったく無い場合は、そのまま登録（重複確認なし）
        Call m_BoardList.Add(dcboard, dcboard.BoardName)
    End If
    
    Set boardchk = Nothing

End Sub

'#Pass-Ok
Public Function Pins(ByVal PinList As String, Optional ByVal DCBoardName As DC_BOARD_NAME = CHAN_TYPE_SELECT) As IVISVISource
'内容:
'   PinMapに定義済ピンのChannelType情報、ピンが使用している電源の指定情報
'   により対応する電源設定用Objectを返します
'
'パラメータ:
'   [PinList]       In   PinMapに定義済みのピン名
'　 [DCBoardName]　 In   指定したピンが使用する電源設定ボード名 オプション(Default CHAN_TYPE_SELECT)
'
'戻り値:
'   電源設定用IVISVISourceオブジェクト
'
'注意事項:
'   ピン名を複数指定する場合は、そのグループに対応する
'   電源設定用Objectが同一である必要があります。
'
'   (例)
'    ○ object.Pins("P_APMU1, P_APMU2, P_APMU3")
'    × object.Pins("P_APMU, P_DPS, P_IO")
'
'   DCBoardNameを指定しない場合は、ピン名に対応するChannelTypeにより
'   電源設定用Objectが選択されます。
'
'   ＊ChannelTypeがchIOとして定義されているピンを指定し
'     DCBoardNameを指定しない場合はPPMUが選択されます｡
'
'    (例)
'     object.Pins("P_IO")        --> PPMUを選択
'     object.Pins("P_IO",dcBPMU) --> BPMUを選択
'
    
    Dim chanType As HdwDrivers.chtype
    Dim selectBoard As IVISVISource
    
    Set Pins = Nothing
    
    'Pin名よりChannelType取得
    chanType = GetChanType(PinList)
      
    '### ボード指定ありの処理 ###
    If DCBoardName <> CHAN_TYPE_SELECT Then
        
        'DCBoardNameをキーに電源設定用OBJを取り出し
        Set selectBoard = Nothing
        
        On Error Resume Next
        Set selectBoard = m_BoardList.Item(DcBoardNameToBoardKey(DCBoardName))
        On Error GoTo 0
    
        If selectBoard Is Nothing Then
            '対応する電源設定用OBJがBoardListになかった場合のエラー
            Call OutputErrMsg("Select board " & DcBoardNameToBoardKey(DCBoardName) & " is nothing (at VI Source Selector)")
        Else
            '対応する電源設定用OBJにピンリストを設定する
            selectBoard.PinList = PinList
            Set Pins = selectBoard
        End If
    
        Set selectBoard = Nothing
        
    '### ボード指定なしの処理 ###
    Else
        'ICUL1G対応処理
        #If ICUL1G_USE <> 0 Then
            '対象ピンが、ICUL1GのPPMUピンかどうかを調べる
            If TheHdw.ICUL1G.IsICUL1GPin(PinList, icul1gPinPPMU) = True Then
                'ICUL1Gボード選択処理
                Set selectBoard = Nothing
                On Error Resume Next
                Set selectBoard = m_BoardList.Item("dcICUL1G")
                On Error GoTo 0
                If selectBoard Is Nothing Then
                    '対応する電源設定用OBJがBoardListになかった場合のエラー
                    Call OutputErrMsg(PinList & " use board is nothing (at VI Source Selector)")
                Else
                    '対応する電源設定用OBJにピンリストを設定する
                    selectBoard.PinList = PinList
                    Set Pins = selectBoard
                End If
                Exit Function
            End If
        #End If
                        
        'ChannelTypeをキーに電源設定用OBJを取り出し
        Set selectBoard = Nothing
        On Error Resume Next
        Set selectBoard = m_BoardList.Item(ChTypeToBoardKey(chanType))
        On Error GoTo 0
    
        If selectBoard Is Nothing Then
            '対応する電源設定用OBJがBoardListになかった場合のエラー
            Call OutputErrMsg(PinList & " use board is nothing (at VI Source Selector)")
        Else
            '対応する電源設定用OBJにピンリストを設定する
            selectBoard.PinList = PinList
            Set Pins = selectBoard
        End If
    
        Set selectBoard = Nothing
    
    End If

End Function


'#Pass-Ok
Public Sub Initialize()
'内容:
'   ボードセレクタの初期化（使用開始準備）
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    
    Call InitBoardList
    Call InitBoardSetting

End Sub

'#Pass-Ok
Private Sub InitBoardSetting()
'内容:
'   ボード準備の初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    'ボードクラスの生成とリスト作成
    Call AddBoard(New CVISApmu)
    Call AddBoard(New CVISDps)
    
    'PPMUボードの準備はコンパイルオプションの指定値で行う
    #If HSD200_USE <> 0 Then
        Call AddBoard(New CHsd200Ppmu)    'HSD200-PPMU
    #Else
        Call AddBoard(New CVISPpmu)       'HSD100-PPMU
    #End If
    
    Call AddBoard(New CVISBpmu)
    
    'HDVISボードの準備はコンパイルオプションで指定された時のみ行う。
    'コンパイルエラー対策
    #If HDVIS_USE <> 0 Then
        Call AddBoard(New CVISHdvis)
    #End If

    'ICUL1Gボードの準備はコンパイルオプションで指定された時のみ行う。
    'コンパイルエラー対策
    'ICUL1G
    #If ICUL1G_USE <> 0 Then
        Call AddBoard(New CIcul1gPpmu)
    #End If

End Sub

'#Pass-Ok
Private Sub Class_Initialize()
    
    Call Me.Initialize

End Sub

'#Pass-Ok
Private Function ChTypeToBoardKey(ByVal ChannelType As HdwDrivers.chtype) As String
'ChannelTypeの値をボード選択用のキー情報に変換する
'＊新しい電源ボードへの対応時にはココに情報追加の必要がある。

    Select Case ChannelType
        Case chAPMU
            ChTypeToBoardKey = "dcAPMU"
        Case chDPS
            ChTypeToBoardKey = "dcDPS"
        Case chIO
            ChTypeToBoardKey = "dcPPMU"    'デジタルPin(I/O Pin)はボード指定が無いときにはPPMUとして取り扱い
        #If HDVIS_USE <> 0 Then
        Case chHDVIS
            ChTypeToBoardKey = "dcHDVIS"
        #End If
    End Select

End Function

'#Pass-Ok
Private Function DcBoardNameToBoardKey(ByVal BoardName As DC_BOARD_NAME) As String
'DC_BOARD_NAMEの値をボード選択用のキー情報に変換する
'＊新しい電源ボードへの対応時にはココに情報追加の必要がある。

    Select Case BoardName
        Case CHAN_TYPE_SELECT
            DcBoardNameToBoardKey = "CHAN_TYPE_SELECT" '未使用
        Case dcAPMU
            DcBoardNameToBoardKey = "dcAPMU"
        Case dcDPS
            DcBoardNameToBoardKey = "dcDPS"
        Case dcPPMU
            DcBoardNameToBoardKey = "dcPPMU"
        Case dcBPMU
            DcBoardNameToBoardKey = "dcBPMU"
        
        #If HDVIS_USE <> 0 Then
        Case dcHDVIS
            DcBoardNameToBoardKey = "dcHDVIS"
        #End If
    
        #If ICUL1G_USE <> 0 Then
        Case dcICUL1G
            DcBoardNameToBoardKey = "dcICUL1G"
        #End If
    
    End Select

End Function
