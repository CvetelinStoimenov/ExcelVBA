Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "q\n14"

' Merging the Files.

Dim wbk As Workbook

strNALI4NOST = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-NALI4NOST.xls"
str1001 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1001.xls"
str1002 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1002.xls"
str1003 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1003.xls"
str1004 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1004.xls"
str1005 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1005.xls"
str1006 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1006.xls"
str1007 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1007.xls"
str1008 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1008.xls"
str1009 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1009.xls"
str1010 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1010.xls"
str1011 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1011.xls"
str1012 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1012.xls"
str1013 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1013.xls"
str1014 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1014.xls"
str1015 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1015.xls"
str1016 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1016.xls"
str1017 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1017.xls"
str1018 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1018.xls"
str1020 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1020.xls"
str1021 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1021.xls"
str1022 = "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\1022.xls"

    Range("A1").Select

Set wbk = Workbooks.Open(str1001)

    Range("A1:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False

    Workbooks("1001.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1002)

    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1002.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1003)

    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1003.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1004)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1004.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1005)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1005.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1006)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1006.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1007)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    
    Application.CutCopyMode = False
    
    Workbooks("1007.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1008)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1008.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1009)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1009.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1010)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1010.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1011)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1011.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1012)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1012.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1013)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1013.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1014)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1014.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1015)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1015.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1016)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1016.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1017)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1017.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1018)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1018.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1020)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1020.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1021)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1021.XLS").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1022)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks("ALL.xlsm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    Workbooks("1022.XLS").Close SaveChanges:=False
    
    'Processing the merged file.
    
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.Range("$A$1:$E$1").AutoFilter Field:=5, Criteria1:=">=0", _
    Operator:=xlAnd
    
    lr = Cells(Rows.Count, 1).End(xlUp).Row
    
        If lr > 1 Then
        
            Range("A2:A" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        
        End If
    
    ActiveSheet.Range("$A$1:$E$1").AutoFilter Field:=5
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft

Set wbk = Workbooks.Open(strNALI4NOST)
    
    Windows("ALL.xlsm").Activate
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
    "=VLOOKUP(C[-2],'[AB-NALI4NOST.xls]SheetName'!C1:C2,2,0)"
    
    Windows("AB-NALI4NOST.xls").Activate
    
    Windows("ALL.xlsm").Activate
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


Dim A As String

    A = InputBox("Enter Criteria: LESS THAN ") ' Taking input from the user.
        
        If IsNumeric(strTemp) Then
        
            'strTemp is a numeric value
        
        Else
            
            'strTemp is not a numeric value
        
        End If
    
    Columns("A:E").Select ' Original macro code.
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
    
    Workbooks("AB-NALI4NOST.xls").Close SaveChanges:=False
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Количество за наливане"
    Columns("C:C").Select
    Selection.Copy
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight

Dim wb As Workbook
Dim ws As Worksheet
Dim ws1 As Worksheet
Dim rng1 As Range
Set wb = ActiveWorkbook
Set ws = wb.Sheets("ALL")
Set rng1 = ws.Cells.Find("*", ws.[a1], xlFormulas, , , xlPrevious)

    Sheets("OBSHT_TRANSFER").Select
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    "ALL!R1C1:R" & CStr(rng1.Row) & "C" & CStr(rng1.Column), Version:=xlPivotTableVersion10).CreatePivotTable _
    TableDestination:="OBSHT_TRANSFER!R3C1", TableName:="PivotTable1", DefaultVersion _
    :=xlPivotTableVersion10
    
    Sheets("OBSHT_TRANSFER").Select
    Cells(3, 1).Select
    
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("Код артикул")
        
            .Orientation = xlRowField
            .Position = 1
        
        End With
    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
    "PivotTable1").PivotFields("Количество за наливане"), _
    "Sum of Количество за наливане", xlSum
    
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
    
    'Recording the transfers to text files.
    
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AH-OBSHT-TRANSFER.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    
    Range("B8").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.End(xlUp).Select
    ActiveCell.FormulaR1C1 = "1"
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\OBSHTA PRODAJBA.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Sheets("ALL").Select
    Sheets("OBSHT_TRANSFER").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Sheets("ALL").Select
    Range("A1:D1").Select

    ActiveSheet.Range("$A$1:$D$1123").AutoFilter Field:=1, Criteria1:="1001"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1001.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1002"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1002.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1003"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1003.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1004"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1004.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1005"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1005.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1006"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1006.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1007"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1007.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1008"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1008.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1009"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1009.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1010"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1010.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1011"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1011.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1012"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1012.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1013"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1013.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1014"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1014.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1015"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1015.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1016"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1016.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1017"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1017.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1018"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1018.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1020"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1020.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1021"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1021.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1022"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Box Sync\Excel\Macros\3 AB-AH\AB-AH-TXT\1022.txt" _
    , FileFormat:=xlText, CreateBackup:=False
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    Selection.AutoFilter
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Save
    
    Application.DisplayAlerts = True
    
    Application.Quit

End Sub

