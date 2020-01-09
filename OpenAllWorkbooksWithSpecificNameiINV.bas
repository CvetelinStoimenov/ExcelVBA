Attribute VB_Name = "Module1"
Public Sub OpenAllWorkbooks()

    Dim vFiles As Variant
    Dim vFile As Variant
    Dim wb As Workbook
    
    Dim fldr As FileDialog
    Dim sItem As String

    'Headers names
    
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "No."
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "EN Description"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Number"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "M.U."
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Qty"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "PPU EURO"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Total EURO"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Sizes"""
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "KG"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "SUM KG"
    
    'Prompt user to select directory
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With

    'Find file which start with specific name

NextCode:
    GetFolder = sItem
    Set fldr = Nothing
    vFiles = EnumerateFiles(sItem + "\", "xls*")
 
    For Each vFile In vFiles
        Workbooks.Open vFile
        
    Set wb = ActiveWorkbook
    Sheets("AUTO HELP EN").Select
       
    'Find specific cell
    
    Cells.Find(What:="Description", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
    'Move one cell left
    
    ActiveCell.Offset(0, -1).Select
    
    'Add filter to entire row
    
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$13:$L$78").AutoFilter Field:=1, Criteria1:="<>"
    
    'Find specific cell
    
    Cells.Find(What:="Description", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
    'Move one cell left and one down
        
    ActiveCell.Offset(1, -1).Select
    
    'Select ranges
    
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    
    Selection.Copy
    
    'Open destination file
    
    Windows("OpenAllWorkbooksWithSpecificNameiINV.xlsm").Activate
    
    'Finding last row
    
    Range("A1048576").End(xlUp).Offset(1, 0).Select
    
    'Paste Values only
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
      
      
    wb.Activate
    
          
    ActiveWorkbook.Close False
        
    'Disable some dialog boxes
    
    Application.CutCopyMode = False
    Application.AskToUpdateLinks = False
    Application.AskToUpdateLinks = True
    Application.CutCopyMode = True
      
    Next vFile

End Sub

Public Function EnumerateFiles(sDirectory As String, _
            Optional sFileSpec As String = "MON", _
            Optional InclSubFolders As Boolean = True) As Variant

    EnumerateFiles = Filter(Split(CreateObject("WScript.Shell").Exec _
        ("CMD /C DIR """ & sDirectory & "*PL of Inv_110000*." & sFileSpec & """ " & _
        IIf(InclSubFolders, "/S ", "") & "/B /A:-D").StdOut.ReadAll, vbCrLf), ".")

End Function

