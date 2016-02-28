VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataSheetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   コントロールシート（Ex．Job List）のマネージャクラス
'
'目的:
'   コントロールシート上に定義されたアクティブシートを取得する
'
'作成者:
'   0145206097
'
Option Explicit

Const DATATOOL_NAME = "B1"

Dim mCtrlSheet As Worksheet
Dim mDataSheet As Collection

Private Sub Class_Initialize()
    Set mCtrlSheet = Nothing
    Set mDataSheet = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCtrlSheet = Nothing
    Set mDataSheet = Nothing
End Sub

Public Sub Initialize(ByVal ctrlSheet As String)
'内容:
'   コントロールシートの指定
'
'[ctrlSheet]   IN String型:     コントロールシート名
'
'備考:
'   コントロールシートが二つ以上ある場合はエラー処理生成を行う
'
    Dim wsheet As Collection
    Set wsheet = collectSheet(ctrlSheet)
    If wsheet Is Nothing Then
        Set mCtrlSheet = Nothing
    ElseIf wsheet.Count = 1 Then
        Set mCtrlSheet = wsheet.Item(1)
    Else
        Err.Raise 9999, "CDataSheetManager.Initialize", "Find More Than One Control Sheet [" & ctrlSheet & "] !"
    End If
End Sub

Public Function GetActiveDataSht(ByVal DataSheet As String) As Worksheet
'内容:
'   取得したいData Toolの種類からアクティブなシートを特定しオブジェクトを返す
'
'[dataSheet]   IN String型:     取得したいData Tool名
'
'備考:
'
    Set GetActiveDataSht = Nothing
    '### Book内の対象となるDataToolシートのオブジェクトを取得 #################
    Set mDataSheet = collectSheet(DataSheet)
    '### Book内に対象となるDataToolシートが存在しないときはエラー処理生成 #####
    If mDataSheet Is Nothing Then
        Exit Function
'        Err.Raise 9999, "CDataSheetManager.GetActiveDataSht", "Can Not Find The Data Tool Sheet [" & datasheet & "] !"
    End If

    Dim currSheet As Worksheet
    Select Case DataSheet
        '### 対象となるDataToolシートがChannel Mapの場合#######################
        Case CHAN_MAP_TOOL:
            For Each currSheet In mDataSheet
                If currSheet.Name = GetCurrentChanMap Then
                    Set GetActiveDataSht = currSheet
                End If
            Next currSheet
        Case Else
        '### コントロールシートが存在しない場合 ###############################
        If mCtrlSheet Is Nothing Then
            If mDataSheet.Count = 1 Then
                Set GetActiveDataSht = mDataSheet.Item(1)
            Else
                '<<< 対象のDataToolシートが複数ある場合はエラー処理生成
'                Err.Raise 9999, "CDataSheetManager.GetActiveDataSht", "Find More Then One Data Tool Sheet [" & datasheet & "] !"
            End If

        '### コントロールシートが存在する場合 #################################
        Else
            '<<< 各コントロールシート用のリーダーオブジェクトを生成
            Dim shReader As IParameterReader
            Select Case mCtrlSheet.Range(DATATOOL_NAME).Value
                Case JOB_LIST_TOOL:
                    Dim crtlSheet As New CJobListSheetReader
                    crtlSheet.Initialize mCtrlSheet.Name
                    crtlSheet.AsIFileStream.SetLocation GetCurrentJobName
                    Set shReader = crtlSheet
                Case SHEET_MANAGER_TOOL:
                    Dim manageSheet As New CManagerSheetReader
                    manageSheet.Initialize mCtrlSheet.Name
                    manageSheet.AsIFileStream.SetLocation GetCurrentJobName
                    Set shReader = manageSheet
                Case Else
                    Err.Raise 9999, "CDataSheetManager.GetActiveDataSht", "Does Not Support The Contorl Sheet [" & mCtrlSheet.Range(DATATOOL_NAME).Value & "] !"
            End Select
            '<<< コントロールシートから対象のDataToolシート名を取得
            Dim paramName As CParamString
            Set paramName = CreateCParamString
            Dim targetSheet As String
            With paramName.AsIParameter
                .Name = DataSheet
                .Read shReader
                targetSheet = .AsString
            End With
            If targetSheet <> "" Then
                '<<< Book内のDataToolシート名とコントロールシート上のDataToolシート名を比較
                For Each currSheet In mDataSheet
                    If currSheet.Name = targetSheet Then
                        Set GetActiveDataSht = currSheet
                    End If
                Next currSheet
                '<<< 不一致の場合はエラー処理生成
'                If GetActiveDataSht Is Nothing Then
''                    Err.Raise 9999, "CDataSheetManager.GetActiveDataSht", "Can Not Find The Sheet Name [" & targetSheet & "] In Data Tool Sheet [" & datasheet & "] !"
'                End If
            End If
        End If
    End Select
End Function

Public Property Get ManagerSheet() As Worksheet
    Set ManagerSheet = mCtrlSheet
End Property

Private Function collectSheet(ByVal DataSheet As String) As Collection
    Dim currSheet As Worksheet
    Set collectSheet = New Collection
    For Each currSheet In Worksheets
        If currSheet.Range(DATATOOL_NAME).Value = DataSheet Then
            collectSheet.Add currSheet
        End If
    Next currSheet
    If collectSheet.Count = 0 Then
        Set collectSheet = Nothing
    End If
End Function
