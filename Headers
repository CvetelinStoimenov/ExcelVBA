Sub Macro1()

'Clearing the clipboard.

Application.CutCopyMode = False

'Adding column on the right.

    Selection.Insert Shift:=xlToRight
    
'Formats the column.

        Selection.NumberFormat = _
        "_-* #,##0.00 [$ˆ-1]_-;-* #,##0.00 [$ˆ-1]_-;_-* ""-""?? [$ˆ-1]_-;_-@_-"
    
    
    With Selection.Font
        
        .Color = -16776961
        .TintAndShade = 0
    
    End With

ActiveCell.FormulaR1C1 = "SUM"

Selection.ColumnWidth = 16

'Adding column on the right.

    Selection.Insert Shift:=xlToRight

'Formats the column.

        Selection.NumberFormat = _
        "_-* #,##0.00 [$ˆ-1]_-;-* #,##0.00 [$ˆ-1]_-;_-* ""-""?? [$ˆ-1]_-;_-@_-"
    
    With Selection.Font
        
        .Color = -16776961
        .TintAndShade = 0
    
    End With
    
ActiveCell.FormulaR1C1 = "Diff."

Selection.ColumnWidth = 16

'Adding column on the right.

    Selection.Insert Shift:=xlToRight

'Formats the column.

    Selection.NumberFormat = _
        "_-* #,##0.00 [$ˆ-1]_-;-* #,##0.00 [$ˆ-1]_-;_-* ""-""?? [$ˆ-1]_-;_-@_-"
    
    With Selection.Font
        
        .Color = -16776961
        .TintAndShade = 0
    
    End With

ActiveCell.FormulaR1C1 = "Net price, EXW Sofia"

Selection.ColumnWidth = 25

    Application.DisplayAlerts = True

'Close the Workbook

    Workbooks("header.xlsm").Close SaveChanges:=False

End Sub
