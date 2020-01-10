Attribute VB_Name = "Module1"
Sub SplitOrdersFromFile()

Dim nSheet As Integer
Dim nTasks As Integer
Dim nSourceRow As Long
Dim nDestRow As Long
Dim wkb As Workbook
Dim wksSource As Worksheet
Dim wksDest As Worksheet
Dim rng As Range

Set rng = Application.InputBox("œŒ—À≈ƒÕ¿  ŒÀŒÕ¿", "»«¡Œ– Õ¿", Type:=8)

Set wkb = ActiveWorkbook
Set wksSource = ActiveSheet

For nTasks = wksSource.Range("C1").Column To rng.Column
    nSheet = nTasks - wksSource.Range("C1").Column + 1
    
    With wkb.Sheets
        
        If .Count < nSheet Then    ' Checks if sheet count on wkb exceeded
            
            Set wksDest = .Add(After:=.Item(.Count), Type:=xlWorksheet)
        
        Else
            
            Set wksDest = .Item(nSheet)    ' Keeps from having empty sheets
        
        End If
        
        Set wksDest = Sheets.Add(After:=Sheets(Sheets.Count))
        wksDest.Name = wksSource.Cells(1, nTasks)
    
    End With

    With wksSource
        wksDest.Cells(1, 1) = .Cells(.UsedRange.Row, 1)   ' Col A
        wksDest.Cells(1, 2) = "QTY"  ' Add header row to sheet
        wksDest.Cells(1, 3) = .Cells(.UsedRange.Row, 2)   ' Col B
        nDestRow = 2
        
        For nSourceRow = .UsedRange.Row + 1 To .UsedRange.Rows.Count
            
            If .Cells(nSourceRow, nTasks).Value <> "" Then

                wksDest.Cells(nDestRow, 1).FormulaR1C1 = _
                    .Range("A" & nSourceRow).Value
                wksDest.Cells(nDestRow, 2).FormulaR1C1 = _
                    .Cells(nSourceRow, nTasks).Value
                wksDest.Cells(nDestRow, 3).FormulaR1C1 = _
                    .Range("B" & nSourceRow).Value
                nDestRow = nDestRow + 1
            
            End If
        
        Next nSourceRow
    
    End With

Next nTasks

    With Application
        
        .ScreenUpdating = False
        .DisplayAlerts = False
    
    End With


Dim fldr As FileDialog
Dim sItem As String
    
    ' Select path for saving files
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fldr
        
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    
    End With

NextCode:
    
    GetFolder = sItem
    Set fldr = Nothing


    'Recording all sheets to text files except listed after ws.Name <>
    
Dim xWs As Worksheet
Dim xTextFile As String
   
    For Each xWs In Application.ActiveWorkbook.Worksheets
    
        If xWs.Name <> "DATA" And xWs.Name <> "OC" And xWs.Name <> "Sheet1" And xWs.Name <> "Sheet2" Then
    
        'Delete columns to all sheets except ws.Name <>
        
        xWs.Columns("D:E").EntireColumn.Delete
        
        xWs.Copy
         
        'Save to folder selected from user
        
        xTextFile = sItem & "/" & xWs.Name & ".txt"
        Application.ActiveWorkbook.SaveAs Filename:=xTextFile, FileFormat:=xlText
        Application.ActiveWorkbook.Saved = True
        Application.ActiveWorkbook.Close
    
        
        End If
    
    Next
    
        'Fill column in all sheets
    
    Dim sht As Worksheet
    Dim lRow As Long

    For Each sht In ActiveWorkbook.Worksheets
    
    'Skip sheets with name in ""
    
    If sht.Name <> "DATA" And sht.Name <> "OC" And sht.Name <> "Sheet1" And sht.Name <> "Sheet2" Then
        
        With sht
        
        'Find last row
            
            lRow = .Range("A" & .Rows.Count).End(xlUp).Row
            
            'Add SUM in cell D2
            
            .Cells(1, 4) = "SUM"
            
            'Fill entire column D with formula
            
            .Range("D2:D" & lRow).Value = "=RC[-1]*RC[-2]"
            
            'Add FOR in cell E2
            
            .Cells(1, 5) = "FOR"
            
            'Fill entire column E with name of the sheet
            
            .Range("E2:E" & lRow).Value = .Name
            
            Total = Total + Application.Sum(.Columns("B"))
            Sum = Sum + Application.Sum(.Columns("D"))
            
        End With
        
        End If
    
    Next sht
    
    Sheets("OC").Select
    Range("E65536").End(xlUp).Offset(4, 0).Value = Total
    Range("J65536").End(xlUp).Offset(4, 0).Value = Sum

Dim sh, intSheets As Integer
    
    intSheets = Application.Worksheets.Count
    
    For sh = 5 To intSheets
        
        Sheets(sh).Select
        Range("B65536").End(xlUp).Offset(2, 0).Value = _
        WorksheetFunction.Sum(Range("B2:B" & Cells.SpecialCells(xlLastCell).Row))
            
        Range("D65536").End(xlUp).Offset(2, 0).Value = _
        WorksheetFunction.Sum(Range("D2:D" & Cells.SpecialCells(xlLastCell).Row))
    
    Next sh
        
        With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    
    End With
    
    Workbooks("SplitOrdersFromFile.xlsm").Close SaveChanges:=False
    
End Sub


