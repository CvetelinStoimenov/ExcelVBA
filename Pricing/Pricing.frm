VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} myfrm 
   Caption         =   "Цена"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7920
   OleObjectBlob   =   "Pricing.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "myfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

    Me.ComboBox3.DropDown

End Sub

Private Sub ComboBox4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Me.ComboBox4.DropDown

End Sub

Private Sub ComboBox3_Change()

    If Not ComboBox3.Text = "" Then
    
        Call Listbox
        Call Combo(sql)
        
    End If
    
End Sub

Private Sub ComboBox4_Change()

    If Not ComboBox4.Text = "" Then
    
        Call Listbox
        Call Combo(sql)
        
    End If

End Sub

Private Sub CommandButton2_Click()

                ListBox1.RowSource = ""

    Set con = Nothing
    
           ComboBox3 = Empty
           ComboBox4 = Empty
                     
            ComboBox3.Clear
            ComboBox4.Clear

    
    Call Userform_initialize
    


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
    

End Sub

Sub Combo(ByVal Tablo As String)

        If Tablo = "" Then Tablo = "[Sheet1$A:J]"
    
    On Error Resume Next

    ComboBox3.Column = con.Execute("select distinct F3 from (" & Tablo & ")").GetRows
    ComboBox4.Column = con.Execute("select distinct F4 from (" & Tablo & ")").GetRows
    
End Sub

Private Sub CommandButton1_Click()

Unload Me
Workbooks("all Conditions.xlsm").Close savechanges:=False

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, _
  CloseMode As Integer)
  
    If CloseMode = vbFormControlMenu Then
    
      Cancel = True
      MsgBox "Моля използвайте бутона затвори!"
      
    End If
  
End Sub
Private Sub Price_Click()
 
    Call Prices
        
End Sub


Sub Prices()

Dim Transport As Currency
Dim Handling As Currency
Dim ADD As Currency
Dim DISCOUNT As Currency
Dim NET As Currency
Dim BGN As Currency
Dim cel As Range
Dim selectedRange As Range
Dim Customer As String
Dim Brand As String

    Customer = ComboBox3.Value
    Brand = ComboBox4.Value
    BGN = 1.96

'        Workbooks("all Conditions.xlsm").Worksheets("Sheet1").Protect Password:="", _
'        UserInterfaceOnly:=True, AllowFiltering:=True
           
    Workbooks("all Conditions.xlsm").Worksheets("Sheet2").Range("B1").Value = Customer
    Workbooks("all Conditions.xlsm").Worksheets("Sheet2").Range("B2").Value = Brand
                  
    If Customer = "" Then
          
        MsgBox ("Моля изберете Клиент.")
          
    Else
      
        If Brand = "" Then
            
            MsgBox ("Моля изберете Бранд.")
                  
        Else
        
        On Error Resume Next
            Transport = Workbooks("all Conditions.xlsm").Worksheets("Sheet2").Range("B3").Value
        On Error Resume Next
            Handling = Workbooks("all Conditions.xlsm").Worksheets("Sheet2").Range("B4").Value
        On Error Resume Next
            ADD = Workbooks("all Conditions.xlsm").Worksheets("Sheet2").Range("B5").Value
        On Error Resume Next
            DISCOUNT = Workbooks("all Conditions.xlsm").Worksheets("Sheet2").Range("B6").Value

    
            Set selectedRange = Application.Selection
            
                For Each cel In selectedRange.Cells
                    
                
                    NET = cel.Offset(0, -1).Value
    
                        If Customer = "KOSER" Then
    
                            cel = Round(((NET * 100 / (100 - Transport)) * 100 / (100 - Handling) * 100 / (100 - ADD) * BGN) * 100 / (100 - DISCOUNT), 2)
    
                        Else
    
                            cel = Round(((NET * 100 / (100 - Transport)) * 100 / (100 - Handling) * 100 / (100 - ADD)) * 100 / (100 - DISCOUNT), 2)
    
                        End If
    
                Next cel
            
    End If
       
        End If
        
    With Me.ListBox1
    
        .ColumnCount = 6
        .ColumnWidths = "60,60,60,60,60,60"
        .RowSource = "('[all Conditions.xlsm]Sheet2'!A12:F13)"
    
    End With

    
           ComboBox4 = Empty
           
            ComboBox4.Clear
    
    Call Userform_initialize
       
    Call ComboBox3_Change
    

End Sub


