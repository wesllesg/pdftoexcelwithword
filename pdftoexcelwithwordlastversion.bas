'pathAndFileName = “C:\…\Power Spreadsheets Examples\Statistical Tables.pdf”

Private sub convertpd (pathAndFileName as string)
Dim myWorksheet As Worksheet
Dim wordApp As Word.Application
Dim myWshShell As wshShell
Dim pathAndFileName As String
Dim registryKey As String
Dim wordVersion As String

Set myWorksheet = ActiveWorkbook.Worksheets(“Word Early Binding”)
Set wordApp = New Word.Application
Set myWshShell = New wshShell

wordVersion = wordApp.Version

registryKey = “HKCU\SOFTWARE\Microsoft\Office\” & wordVersion & “\Word\Options\”

myWshShell.RegWrite registryKey & “DisableConvertPdfWarning”, 1, “REG_DWORD”

wordApp.Documents.Open Filename:=pathAndFileName, ConfirmConversions:=False

myWshShell.RegWrite registryKey & “DisableConvertPdfWarning”, 0, “REG_DWORD”

wordApp.ActiveDocument.Content.Copy

With myWorksheet
End With

Range(“B4”).Select
PasteSpecial Format:=”Text”

wordApp.Quit SaveChanges:=wdDoNotSaveChanges

Set wordApp = Nothing
Set myWshShell = Nothing

End Sub
