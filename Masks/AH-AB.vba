Sub Macro1()
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+q
'
Dim wbk As Workbook
strMIN = "C:\Users\Natalia\Desktop\2 AH-AB\AB-MIN.xls"
strNALI4NOST = "C:\Users\Natalia\Desktop\2 AH-AB\AH-NALI4.xls"

    Range("A1").Select
    
Set wbk = Workbooks.Open(strMIN)
    
    Range("A1:H3").Select
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("AH-AB.xlsm").Worksheets("AH-AB").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    
    Application.CutCopyMode = False
    
    Workbooks("AB-MIN.XLS").Close SaveChanges:=False
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$D$9706").AutoFilter Field:=1, Criteria1:="=0008", _
        Operator:=xlAnd
     
     lr = Cells(Rows.Count, 1).End(xlUp).Row
    
        If lr > 1 Then
        
            Range("A2:A" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        
        End If
    
    Selection.AutoFilter
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    
Set wbk = Workbooks.Open(strNALI4NOST)
    
    Windows("AH-AB.xlsm").Activate
    
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(C[-2],'[AH-NALI4.xls]SheetName'!C1:C2,2,0)"
    
    Windows("AH-NALI4.xls").Activate
    
    Windows("AH-AB.xlsm").Activate
    
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-2]"
    
    Range("D2:E2").Select
    ActiveCell.Offset(0, -1).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    ActiveCell.Offset(0, -1).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 2).Select
    
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    Rows("1:1").Select
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Range("E2").Select

    Dim A As String
    
    A = InputBox("Enter Criteria: LESS THAN ") ' take your input from the user
    
        If IsNumeric(strTemp) Then
            
            'strTemp is a numeric value
        
        Else
            
            'strTemp is not a numeric value
        
        End If
        
    Columns("A:E").Select ' original macro code
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$A$53").AutoFilter Field:=5, Criteria1:="<" & A, Operator:=xlAnd
    
    lr = Cells(Rows.Count, 1).End(xlUp).Row
    
    If lr > 1 Then
        
        Range("A2:A" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    
    End If


    ActiveSheet.Range("$A$1:$E$33139").AutoFilter Field:=5, Criteria1:="#N/A"
    
    lr = Cells(Rows.Count, 1).End(xlUp).Row
    
        If lr > 1 Then
            
            Range("A2:A" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        
        End If
    
    Range("A1:E1").Select
    Selection.AutoFilter
    Columns("A:E").Select
    Application.CutCopyMode = False
    
    Workbooks("AH-NALI4.xls").Close SaveChanges:=False
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Êîëè÷åñòâî çà íàëèâàíå"
    Columns("D:D").Select
    Selection.Copy
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "1"
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="0000"
    
    Range("B1:D2").Select
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks.Add
    
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\Natalia\Desktop\2 AH-AB\0000.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="0001"
    
    Range("B1:D2").Select
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks.Add
    
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\Natalia\Desktop\2 AH-AB\0001.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="0006"
    
    Range("B1:D2").Select
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks.Add
    
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\Natalia\Desktop\2 AH-AB\0006.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
Dim wb As Workbook
Dim ws As Worksheet
Dim ws1 As Worksheet
Dim rng1 As Range
Set wb = ActiveWorkbook
Set ws = wb.Sheets("AH-AB")
Set rng1 = ws.Cells.Find("*", ws.[a1], xlFormulas, , , xlPrevious)
    
    Sheets("OBSHT_TRANSFER").Select
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    "AH-AB!R1C1:R" & CStr(rng1.Row) & "C" & CStr(rng1.Column), Version:=xlPivotTableVersion10).CreatePivotTable _
    TableDestination:="OBSHT_TRANSFER!R3C1", TableName:="PivotTable1", DefaultVersion _
    :=xlPivotTableVersion10
    
    Sheets("OBSHT_TRANSFER").Select
    Cells(3, 1).Select
    

        With ActiveSheet.PivotTables("PivotTable1").PivotFields("Êîä àðòèêóë")
            
            .Orientation = xlRowField
            .Position = 1
        
        End With
    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Êîëè÷åñòâî çà íàëèâàíå"), _
        "Sum of Êîëè÷åñòâî çà íàëèâàíå", xlSum
    
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("1:3").Select
    
    Range("A2").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    
    Columns("B:B").Select
    Selection.Copy
    
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(-1, 0).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Copy
    
    Workbooks.Add
    
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\Natalia\Desktop\2 AH-AB\OBSHT TRANSFER.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Sheets("AH-AB").Select
    
    Sheets("OBSHT_TRANSFER").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    
    Sheets("AH-AB").Select
    Range("A1").Select
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Save
    
    Application.DisplayAlerts = True
    
    Application.Quit
    
End Sub
