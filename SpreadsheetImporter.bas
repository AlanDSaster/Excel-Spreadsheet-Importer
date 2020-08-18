Attribute VB_Name = "SpreadsheetImporter"
Option Explicit

Sub ImportSheets()
    
    Dim directory As String
    Dim sFile As String
    Dim wb As Workbook
    Dim ws_input As Worksheet
    Dim fullpath As String
    
    directory = SelectFolder
    
    If directory <> "" Then
        
        Set wb = ThisWorkbook
        
        sFile = Dir(directory & "\*.xl*")
        Do While Len(sFile) > 0
            fullpath = directory & "\" & sFile
            Set ws_input = Workbooks.Open(fullpath).Sheets(1)
            ws_input.Copy wb.Sheets(1)
            ws_input.Parent.Close
            sFile = Dir
        Loop
        
        MsgBox "Processing Complete"
        
    End If
    
End Sub

Function SelectFolder() As String
    
    SelectFolder = ""
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then SelectFolder = .SelectedItems(1)
    End With
End Function
