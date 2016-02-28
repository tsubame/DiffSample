Imports System
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Public Class MainMenuForm

    ' 定数

    ' ライブラリのディレクトリ　実行ファイルからの相対位置
    Const LIB_DIR_PATH = "\..\..\..\lib\"

    ' AriAwaseライブラリ
    Const ARI_DIR_PATH = LIB_DIR_PATH + "Ariawase\" ' ディレクトリ名
    Const BAT_FILE_NAME = "command.bat"             ' バッチファイル名
    Const EXCEL_FILE_DIR = "excels"                      ' ジョブを入れるディレクトリ
    Const COPY_EXCEL_DIR = "bin"                      ' 実行時にジョブをコピーするディレクトリ
    'Const EXCEL_DIR_PATH = Application.ExecutablePath & ARI_DIR_PATH & ""

    ' ロード時に実行
    Private Sub MainMenuForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        copyJobs()
    End Sub


    ' バッチファイルの実行
    Private Sub execBat()
        Try
            ' カレントディレクトリの移動
            IO.Directory.SetCurrentDirectory(Application.ExecutablePath & ARI_DIR_PATH)
            Debug.Print("バッチファイルを実行します。")

            'Call Shell(BAT_FILE_NAME, vbHide)

            ''ProcessStartInfoオブジェクトを作成する
            Dim psi As New System.Diagnostics.ProcessStartInfo()
            ''メモ帳の実行ファイルのパスを指定する
            psi.FileName = BAT_FILE_NAME
            ''WindowStyleにMinimizedを指定して、最小化された状態で起動されるようにする
            psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
            ''アプリケーションを起動する
            Dim p As Process = Process.Start(psi)

            ''Dim p As Process
            ''p = Process.Start(BAT_FILE_NAME, vbHide)
            p.WaitForExit()
            '''終了したか確認する
            ''If p.HasExited Then
            Debug.Print("終了しました。")
            ''Else
            ''    MessageBox.Show("終了していません。")
            ''End If

        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

    End Sub

    ' ジョブファイルをコピー
    Private Sub copyJobs()
        ' ジョブのディレクトリに移動
        Dim jobsOrgDir = Application.ExecutablePath & ARI_DIR_PATH & EXCEL_FILE_DIR
        Debug.Print(jobsOrgDir)
        Directory.SetCurrentDirectory(jobsOrgDir)
        ' ファイルの一覧を取得
        Dim files As String() = Directory.GetFiles(Directory.GetCurrentDirectory(), "*", SearchOption.AllDirectories)

        Try
            Dim fName As String
            For Each fPath In files
                Debug.Print(fPath)
                fName = Path.GetFileName(fPath)
                ' コピー
                File.Copy(fPath, "..\" & COPY_EXCEL_DIR & "\" & fName)
            Next

        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub

    Private Sub converExcelToCSV(filePath As String)

        Const CSV_DIR_NAME = "src"

        Dim excel As New Excel.Application() 'EXCELの宣言
        Dim Book As Excel.Workbook

        Try
            ' EXCEL　の画面表示
            excel.Visible = False
            ' EXCEL　のﾒｯｾｰｼﾞ
            excel.Application.DisplayAlerts = False
            ' EXCEL　のｵｰﾌﾟﾝ
            Book = excel.Application.Workbooks.Open(Filename:=filePath)
            ' 1000 ミリ秒 (1秒) 待機する
            System.Threading.Thread.Sleep(1000)
            ' EXCELの保存 ﾌｧｲﾙﾌｫｰﾏｯﾄ 42:ﾀﾌﾞ区切CSV 43:EXL 44:XML
            For Each sheet In Book.Sheets
                Debug.Print(sheet.ToString)
            Next

            'excel.ActiveWorkbook.SaveAs(Filename:=filePath + ".csv", FileFormat:=42, CreateBackup:=False)
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

        excel.Quit()
        excel = Nothing

        'Dim FPN As String 'ﾌｧｲﾙパス
        'Dim STR As String 'ファイル名

        ' 出力フォルダ作成
        'Directory.CreateDirectory(Path.GetTempPath() & CSV_DIR_NAME)

    End Sub

    Private Sub toCSV(fPath As String)
        Dim xls = New Excel.Application

        xls.Visible = False
        xls.DisplayAlerts = False

        xls.Workbooks.Open(fPath)

        With xls.ActiveWorkbook
            Dim i
            For i = 1 To .Sheets.Count
                .Sheets(i).Select
                Dim csvPath = Path.GetFileName(fPath)
                Dim csv = COPY_EXCEL_DIR & "\" & .Sheets(i).Name & ".csv"
                .SaveAs(csv, 6)
            Next
        End With

        xls.Quit()
        xls = Nothing
    End Sub

    ' ボタンクリック
    Private Sub execButton_Click(sender As Object, e As EventArgs) Handles execButton.Click
        ' ジョブのディレクトリに移動
        Dim dir = Application.ExecutablePath & ARI_DIR_PATH & EXCEL_FILE_DIR
        Directory.SetCurrentDirectory(dir)
        ' ファイルの一覧を取得
        Dim files As String() = Directory.GetFiles(Directory.GetCurrentDirectory(), "*", SearchOption.AllDirectories)

        Try
            'Dim fName As String

            For Each fPath In files
                toCSV(fPath)
            Next

        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

    End Sub
End Class