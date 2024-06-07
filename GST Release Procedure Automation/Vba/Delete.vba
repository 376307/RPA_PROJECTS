Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Sheet1").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("2A").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("6A").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Filtered 6A").Select
    ActiveWindow.SelectedSheets.Delete
    Range("A1").Select
End Sub
