VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "all Conditions"
   ClientHeight    =   12090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   23940
   OleObjectBlob   =   "all Conditions.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dynamic List in Combobox

Dim con As Object
Dim rs As Object
Dim sql As String

Private Sub ComboBox1_Change()

    If Not ComboBox1.Text = "" Then
    
        Call Listbox
        Call Combo(sql)
        
    End If

End Sub

Private Sub ComboBox2_Change()

    If Not ComboBox2.Text = "" Then
    
        Call Listbox
        Call Combo(sql)
        
    End If

End Sub


Private Sub ComboBox3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Workbooks("all Conditions.xlsm").Worksheets("Sheet1").Activate


With Me.ListBox2

    .ColumnHeads = True
    .ColumnCount = 10
    .ColumnWidths = "20,100,80,110,55,55,55,55,110,70"
    .RowSource = "Sheet1!A2:J2"

End With
    
    Me.ComboBox3.DropDown

End Sub

Private Sub ComboBox3_Change()

    If Not ComboBox3.Text = "" Then
    
        Call Listbox
        Call Combo(sql)
        
    End If
    
End Sub

Private Sub ComboBox4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Workbooks("all Conditions.xlsm").Worksheets("Sheet1").Activate

With Me.ListBox2

    .ColumnHeads = True
    .ColumnCount = 10
    .ColumnWidths = "20,100,80,110,55,55,55,55,110,70"
    .RowSource = "Sheet1!A2:J2"

End With

    Me.ComboBox4.DropDown

End Sub

Private Sub ComboBox4_Change()

    If Not ComboBox4.Text = "" Then
    
        Call Listbox
        Call Combo(sql)
        
    End If

End Sub

Private Sub CommandButton6_Click()

    Set con = Nothing
    
           ComboBox3 = Empty
           ComboBox4 = Empty
           
    ListBox1.Clear
    
    Call Userform_initialize

End Sub

Private Sub ToggleButton1_Click()

    If ToggleButton1.Value = False Then
    
        Application.Visible = False
    
           End If
    
       If ToggleButton1.Value = True Then
    
        Application.Visible = True
        UserForm1.Hide
    
    End If
    
End Sub

Private Sub ToggleButton2_Click()
    
    Call ToggleButton1_Click
    
        UserForm1.Hide
            
End Sub

Private Sub Userform_initialize()


    Set con = CreateObject("adodb.connection")
    
        #If VBA7 And Win64 Then
        
            con.Open "provider=microsoft.ace.oledb.12.0;data source=" & ThisWorkbook.FullName & ";extended properties=""excel 12.0;hdr=no"""
            
        #Else
            
            con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & ThisWorkbook.FullName & ";extended properties=""excel 8.0;hdr=no"""
            
        #End If
    
    Call Combo("")

End Sub

Sub Listbox()

    sql = "select * from [Sheet1$A:J] Where F1 is not null"
    
        If ComboBox3.Text <> "" Then sql = sql & " and f3 = '" & ComboBox3.Value & "'"
        If ComboBox4.Text <> "" Then sql = sql & " and f4 = '" & ComboBox4.Value & "'"
    
    Set rs = con.Execute(sql)
    
        If rs.BOF And rs.EOF Then
        
            MsgBox "Please enter valid record!"
            
        Else
    
    ListBox1.ColumnCount = rs.Fields.Count
    ListBox1.Column = rs.GetRows(rs.RecordCount)
    ListBox1.ColumnWidths = "20,100,80,110,55,55,55,55,110,70"
        
        End If
                
'    'Sorts ListBox
'
'    Dim i As Long
'    Dim j As Long
'    Dim sTemp As String
'    Dim sTemp2 As String
'    Dim sTemp3 As String
'    Dim sTemp4 As String
'    Dim sTemp5 As String
'    Dim sTemp6 As String
'    Dim LbList As Variant
'
'    'Store the list in an array for sorting
'    LbList = Me.ListBox1.List
'
'
'    For i = LBound(LbList, 1) To UBound(LbList, 1)
'        For j = i + 1 To UBound(LbList, 1)
'
'            'Bubble sort the array on the fourth value
'            If LbList(i, 2) > LbList(j, 2) Or LbList(i, 3) > LbList(j, 3) Then
'
'                'skip the first column
'                sTemp = LbList(i, 0)
'                LbList(i, 0) = LbList(j, 0)
'                LbList(j, 0) = sTemp
'
'                'skip the second column
'                sTemp2 = LbList(i, 1)
'                LbList(i, 1) = LbList(j, 1)
'                LbList(j, 1) = sTemp2
'
'                'skip the third column
'                 sTemp3 = LbList(i, 2)
'                LbList(i, 2) = LbList(j, 2)
'                LbList(j, 2) = sTemp3
'
'                'skip the fourth column
'                sTemp4 = LbList(i, 3)
'                LbList(i, 3) = LbList(j, 3)
'                LbList(j, 3) = sTemp4
'
'
'            End If
'        Next j
'    Next i
'
'    'Remove the contents of the listbox
'    Me.ListBox1.Clear
'
'    'Repopulate with the sorted list
'    Me.ListBox1.List = LbList
'

End Sub

Sub Combo(ByVal Tablo As String)

        If Tablo = "" Then Tablo = "[Sheet1$A:J]"
    
    On Error Resume Next
    
    ComboBox3.Column = con.Execute("select distinct F3 from (" & Tablo & ")").GetRows
    ComboBox4.Column = con.Execute("select distinct F4 from (" & Tablo & ")").GetRows

End Sub

Private Sub CommandButton2_Click()


    'ADD record
    
    Dim sh As Worksheet
    Set sh = Workbooks("all Conditions.xlsm").Sheets("Sheet1")
    
    Dim Last_Row As Long
    Last_Row = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    Workbooks("all Conditions.xlsm").Worksheets("Sheet1").Activate
    
    '======================= Validation =========================
    
        If Me.TextBox7.Value = "" Then
        
            MsgBox "Please enter the Customer", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox8.Value = "" Then
        
            MsgBox "Please enter the Brand", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox4.Value = "" Then
        
            MsgBox "Please enter Transport", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox3.Value = "" Then
        
            MsgBox "Please enter Hendling", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox1.Value = "" Then
        
            MsgBox "Please enter ADD", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox2.Value = "" Then
        
            MsgBox "Please enter Discount", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox5.Value = "" Then
        
            MsgBox "Please enter Price condition", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox6.Value = "" Then
        
            MsgBox "Please enter Valid from date", vbCritical
            
            Exit Sub
            
        End If
    
    '===================================================
    
    sh.Range("A" & Last_Row + 1).Value = "=IF(B" & Last_Row + 1 & "="""","""",ROW()-1)"
    sh.Range("B" & Last_Row + 1).Value = Now
    sh.Range("C" & Last_Row + 1).Value = Me.TextBox7.Value
    sh.Range("D" & Last_Row + 1).Value = Me.TextBox8.Value
    sh.Range("E" & Last_Row + 1).Value = Me.TextBox4.Value
    sh.Range("F" & Last_Row + 1).Value = Me.TextBox3.Value
    sh.Range("G" & Last_Row + 1).Value = Me.TextBox1.Value
    sh.Range("H" & Last_Row + 1).Value = Me.TextBox2.Value
    sh.Range("I" & Last_Row + 1).Value = Me.TextBox5.Value
    sh.Range("J" & Last_Row + 1).Value = Me.TextBox6.Value
    
    Me.TextBox7.Value = ""
    Me.TextBox8.Value = ""
    Me.TextBox4.Value = ""
    Me.TextBox3.Value = ""
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    Me.TextBox5.Value = ""
    Me.TextBox6.Value = ""
    
        Call Listbox
    
            MsgBox "The Data Was Added."

End Sub

Private Sub CommandButton3_Click()

    'UPDATE record

Workbooks("all Conditions.xlsm").Worksheets("Sheet1").Activate
   
        If Me.TextBox7.Value = "" Then
    
              MsgBox "Please select a record to update", vbCritical
            
            Exit Sub
        
        End If

    Dim sh As Worksheet
    Set sh = Workbooks("all Conditions.xlsm").Sheets("Sheet1")
    
    Dim Selected_Row As Long
    Selected_Row = Application.WorksheetFunction.Match(Me.ListBox1.List(Me.ListBox1.ListIndex, 0), sh.Range("A:A"), 0)
    
    
    '======================= Validation =========================
    
        If Me.TextBox8.Value = "" Then
        
            MsgBox "Please enter the Customer", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox4.Value = "" Then
        
            MsgBox "Please enter the Brand", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox4.Value = "" Then
        
            MsgBox "Please enter Transport", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox3.Value = "" Then
        
            MsgBox "Please enter Hendling", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox1.Value = "" Then
        
            MsgBox "Please enter ADD", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox2.Value = "" Then
        
            MsgBox "Please enter Discount", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox5.Value = "" Then
        
            MsgBox "Please enter Price condition", vbCritical
            
            Exit Sub
            
        End If
        
        If Me.TextBox6.Value = "" Then
        
            MsgBox "Please enter Valid from date", vbCritical
            
            Exit Sub
            
        End If
    
    '===================================================
    sh.Range("A" & Selected_Row).EntireRow.Copy Destination:=Sheets("OldRecords").Range("A" & Rows.Count).End(xlUp).Offset(1)
    Sheets("OldRecords").Range("K" & Rows.Count).End(xlUp).Offset(1).Value = Now
    
    sh.Range("B" & Selected_Row).Value = Now
    sh.Range("C" & Selected_Row).Value = Me.TextBox7.Value
    sh.Range("D" & Selected_Row).Value = Me.TextBox8.Value
    sh.Range("E" & Selected_Row).Value = Me.TextBox4.Value
    sh.Range("F" & Selected_Row).Value = Me.TextBox3.Value
    sh.Range("G" & Selected_Row).Value = Me.TextBox1.Value
    sh.Range("H" & Selected_Row).Value = Me.TextBox2.Value
    sh.Range("I" & Selected_Row).Value = Me.TextBox5.Value
    sh.Range("J" & Selected_Row).Value = Me.TextBox6.Value
    
    Me.TextBox7.Value = ""
    Me.TextBox8.Value = ""
    Me.TextBox4.Value = ""
    Me.TextBox3.Value = ""
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    Me.TextBox5.Value = ""
    Me.TextBox6.Value = ""
    
        Call Listbox
        
            MsgBox "The Data Was Updated."

End Sub

Private Sub CommandButton5_Click()

    'DEL record
    
    Workbooks("all Conditions.xlsm").Worksheets("Sheet1").Activate
           
        If Me.ListBox1.ListIndex < 0 Then
        
            MsgBox "Please select a record to delete", vbCritical
            
            Exit Sub
            
        End If
        
    On Error GoTo MyErrorHandler:
    
    Dim sh As Worksheet
    Set sh = Workbooks("all Conditions.xlsm").Sheets("Sheet1")
    
    Dim Selected_Row As Long
    Selected_Row = Application.WorksheetFunction.Match(Me.ListBox1.List(Me.ListBox1.ListIndex, 0), sh.Range("A:A"), 0)
    
    sh.Range("A" & Selected_Row).EntireRow.Copy Destination:=Sheets("OldRecords").Range("A" & Rows.Count).End(xlUp).Offset(1)
    Sheets("OldRecords").Range("K" & Rows.Count).End(xlUp).Offset(1).Value = Now
       
    sh.Range("A" & Selected_Row).EntireRow.Delete
        

    
MyErrorHandler:

    If Err.Number = 1004 Then
    
        MsgBox "Please enter valid record to delete."

    Else
    
        MsgBox "The record was deleted."

    End If
        
                Call Listbox

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, _
  CloseMode As Integer)
  
    If CloseMode = vbFormControlMenu Then
    
      Cancel = True
      MsgBox "Please use the Close button!"
      
    End If
  
End Sub
Private Sub CommandButton4_Click()

    'OK
    
    Unload Me
    Workbooks("all Conditions.xlsm").Close savechanges:=True

End Sub
'
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    'OK
    
    If Me.ListBox1.List(Me.ListBox1.ListIndex, 0) <> "" Then
    
        Me.TextBox7.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
        Me.TextBox8.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 3)
        Me.TextBox4.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 4)
        Me.TextBox3.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 5)
        Me.TextBox1.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 6)
        Me.TextBox2.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 7)
        Me.TextBox5.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 8)
        Me.TextBox6.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 9)
        
    End If


End Sub
