Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Month Wise").Select
    Columns("A:A").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    ActiveWorkbook.Save
End Sub
