Attribute VB_Name = "XLibImgPmdSheet"
'概要:
'   ###このモジュールの役割などを１～２行でわかるように記述してください###
'
'目的:
'   ###目的等の詳細を記述してください###
'
'作成者:
'   0145184004
'
Option Explicit

Private m_PmdSht As CImgPmdSheet

Public Sub Initialize(ByVal pShtName As String)
'内容:
'   指定したシートの機能を有効にする。
'
'[pShtName]    IN   String:         対象のシート名
'
'備考:
'
    Set m_PmdSht = New CImgPmdSheet
    Set m_PmdSht.targetSheet = Worksheets(pShtName)
End Sub

Public Sub CreatePMD(ByVal pShtName As String)
'内容:
'   指定したシートのデータに従って、PMDを作成する。
'
'[pShtName]    IN   String:         対象のシート名
'
'備考:
'
    Call Initialize(pShtName)
    Application.StatusBar = "Creating Base PMD..."
    Call m_PmdSht.CreatePMD
    Application.StatusBar = False
End Sub

Public Sub AddPmdSheet()
'内容:
'   PMDシートをJobに追加する。
'
'備考:
'
        
    Dim shtEnd As Worksheet
    
    With Worksheets
        Set shtEnd = .Item(.Count)
    End With
    
    Call ShtPMD.Copy(, shtEnd)
    
End Sub

Public Sub CreatePMDIfNothing()
    On Error GoTo ErrHandler
    If TheIDP.PlaneManagerCount = 0 Then
        Call CreatePMD(GetWkShtReaderManagerInstance.GetActiveSheetName(shtTypePMDDefinition))
    End If
    Exit Sub
ErrHandler:
    DestroyPMDSheet
    DestroyTheIDP
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub DestroyPMDSheet()
    Set m_PmdSht = Nothing
End Sub
