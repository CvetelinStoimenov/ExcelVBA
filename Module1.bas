Attribute VB_Name = "Module1"
Sub ABC()

    Dim Wb As Workbook
    Dim rng As Range
    Dim LastRow As Long
    Set Wb = ActiveWorkbook
    Dim Wb2 As Workbook
    strFAG = "C:\Users\irina\Desktop\ABC.xlsx"
    Set rng = Application.InputBox("Избор на клетка за сравнение:", "", Type:=8)
    Set wbk = Workbooks.Open(strFAG)
    Wb.Activate
    ActiveCell.FormulaR1C1 = "=VLOOKUP(" & rng.Address(ReferenceStyle:=xlR1C1) & " ,[ABC.xlsx]SheetName!C1:C2,2,0)"
        '/ Set it to relative removing dollar sign
    ActiveCell.Formula = Application.ConvertFormula(ActiveCell.Formula, xlA1, xlA1, 4)
    LastRow = Range(rng.Address).End(xlDown).Row
    Range(Selection, Selection.Offset(LastRow - Range(rng.Address).Row, 0)).Select
    Selection.FillDown
    
    Workbooks("ABC.xlsx").Close SaveChanges:=False
    Workbooks("ABC.xlsm").Close SaveChanges:=False
End Sub
