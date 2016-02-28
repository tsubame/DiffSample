Attribute VB_Name = "ModuleExporter"
Option Explicit


Public Sub showModules(ByVal path As String)

    Dim book As Excel.Workbook
    Dim books As Excel.Workbooks

    Set book = Workbooks.Open(path)
    book.Activate
    Application.Visible = False

    Dim comp As Variant
    Dim comps As Collection

    Debug.Print ""
    Set comps = ThisWorkbook.VBProject.VBComponents
    
    MsgBox ("show.")
'    For Each comp In ThisWorkbook.VBProject.VBComponents
'
'        'Debug.Print comp
'
'    Next
End Sub


Public Sub exec()

    Dim comp As Variant
    Dim comps As Collection

    Debug.Print ""
    For Each comp In ThisWorkbook.VBProject.VBComponents
    
        Debug.Print comp
    
    Next

End Sub
