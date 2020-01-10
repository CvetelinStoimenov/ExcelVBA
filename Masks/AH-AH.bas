Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "q\n14"

'Merging the Files.

Dim wbk As Workbook

strALL = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\ALL.xlsxm"
str1001 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1001.xlsx"
str1002 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1002.xlsx"
str1003 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1003.xlsx"
str1004 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1004.xlsx"
str1005 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1005.xlsx"
str1006 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1006.xlsx"
str1007 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1007.xlsx"
str1008 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1008.xlsx"
str1009 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1009.xlsx"
str1010 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1010.xlsx"
str1011 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1011.xlsx"
str1012 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1012.xlsx"
str1013 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1013.xlsx"
str1014 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1014.xlsx"
str1015 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1015.xlsx"
str1016 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1016.xlsx"
str1017 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1017.xlsx"
str1018 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1018.xlsx"
str1020 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1020.xlsx"
str1021 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1021.xlsx"
str1022 = "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\1022.xlsx"

    Range("A1").Select

Set wbk = Workbooks.Open(str1001)

    Range("A1:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1001.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1002)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1002.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1003)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1003.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1004)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1004.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1005)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1005.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1006)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1006.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1007)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1007.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1008)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1008.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1009)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1009.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1010)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1010.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1011)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1011.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1012)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1012.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1013)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1013.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1014)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1014.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1015)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1015.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1016)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1016.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1017)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1017.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1018)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1018.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1020)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1020.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1021)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1021.xlsx").Close SaveChanges:=False

Set wbk = Workbooks.Open(str1022)
    
    Range("A2:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("ALL.xlsxm").Worksheets("ALL").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Application.CutCopyMode = False
    Workbooks("1022.xlsx").Close SaveChanges:=False
    
''Processing the merged file.
    
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$E$33907").AutoFilter Field:=5, Criteria1:=">=0", _
    Operator:=xlAnd
    
    lr = Cells(Rows.Count, 1).End(xlUp).Row
        
        If lr > 1 Then
           
           Range("A2:A" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        
        End If
    
    ActiveSheet.Range("$A$1:$E$64148").AutoFilter Field:=5
    ActiveSheet.Range("$A$1:$E$64148").AutoFilter Field:=4, Criteria1:= _
    "=*1000 ÖÅÍÒÐÀËÅÍ ÑÊËÀÄ - 1.000;*", Operator:=xlAnd
    
    lr = Cells(Rows.Count, 1).End(xlUp).Row
        
        If lr > 1 Then
            
            Range("A2:A" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        
        End If
    
    ActiveSheet.Range("$A$1:$E$56456").AutoFilter Field:=4, Criteria1:= _
    "<>*1000 ÖÅÍÒÐÀËÅÍ ÑÊËÀÄ*", Operator:=xlAnd
    
    lr = Cells(Rows.Count, 1).End(xlUp).Row
        
        If lr > 1 Then
            
            Range("A2:A" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        
        End If
    
    ActiveSheet.Range("$A$1:$E$2158").AutoFilter Field:=4
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Copy
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Application.CutCopyMode = False
    
'Recording the transfers to text files.

    ActiveSheet.Range("$A$1:$C$1123").AutoFilter Field:=1, Criteria1:="1001"
    Range("B1:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs Filename:= _
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1001.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1002.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1003.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1004.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1005.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1006.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1007.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1008.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1009.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1010.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1011.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1012.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1013.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1014.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1015.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1016.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1017.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1018.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1020.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1021.txt" _
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
    "C:\Users\irina\Desktop\4 AH-AH\AH-AH_21\AH-AH-TXT-20\1022.txt" _
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

