Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Sheet1").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    ActiveWorkbook.Save
    Range("C2").Select
End Sub
