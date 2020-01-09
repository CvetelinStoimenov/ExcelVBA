Sub Macro1()

' Keyboard Shortcut: Ctrl+q

    Dim wbk As Workbook
    strMIN = "C:\Users\irina\Desktop\1 AB-AB\AB-MIN.xls"
    strNALI4NOST = "C:\Users\irina\Desktop\1 AB-AB\AB-NALI4NOST.xls"
    
        Range("A1").Select
    
    Set wbk = Workbooks.Open(strMIN)
    
        Range("A1:H3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        
        Workbooks("AB-AB.xlsm").Worksheets("AB-AB").Activate
        
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
    
    Windows("AB-AB.xlsm").Activate
    
        Range("D2").Select
        
        ActiveCell.FormulaR1C1 = _
            "=VLOOKUP(C[-2],'[AB-NALI4NOST.xls]SheetName'!C1:C2,2,0)"
        
        Windows("AB-NALI4NOST.xls").Activate
        
        Windows("AB-AB.xlsm").Activate
        
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
        
        Selection.AutoFilter
        
        Range("E2").Select
        
    Dim A As String
    
    A = InputBox("Enter Criteria: LESS THAN ") ' take your input from the user
    
        If IsNumeric(strTemp) Then
        
            'strTemp is a numeric value
        
        Else
            
            'strTemp is not a numeric value
        
        End If
    
    Columns("A:E").Select ' original macro code

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
    
    Workbooks("AB-NALI4NOST.xls").Close SaveChanges:=False
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Êîëè÷åñòâî çà íàëèâàíå"
    Columns("C:C").Select
    Selection.Copy
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    
    ActiveSheet.Range("$A$1:$D$1123").AutoFilter Field:=1, Criteria1:="0000"
    
    Range("B1:D2").Select
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Desktop\1 AB-AB\0000.txt" _
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
    "C:\Users\irina\Desktop\1 AB-AB\0001.txt" _
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
    "C:\Users\irina\Desktop\1 AB-AB\0006.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
    
    Application.Quit

End Sub
