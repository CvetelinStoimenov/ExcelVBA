VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   OleObjectBlob   =   "BrandsUserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Customer_Exit(ByVal Cancel As MSForms.ReturnBoolean)


End Sub


Private Sub CloseButton_Click()

'Close the Workbook

Unload Me
Workbooks("BRANDS.xlsm").Close SaveChanges:=False

End Sub


Private Sub Label2_Click()

End Sub

Private Sub Price_Click()
 
    Call Brands
        
End Sub


Private Sub UserForm_Initialize()

'Makes a drop down menu with listed items

With Brand

    .AddItem "ATE"
    .AddItem "BERU"
    .AddItem "BILSTEIN"
    .AddItem "BOSCH"
    .AddItem "BREMBO DISCS"
    .AddItem "BREMBO BRAKE PADS"
    .AddItem "CONTITECH"
    .AddItem "CORTECO"
    .AddItem "DAYCO"
    .AddItem "DELPHI"
    .AddItem "DENSO"
    .AddItem "DOLZ"
    .AddItem "FAI"
    .AddItem "GATES"
    .AddItem "HELLA"
    .AddItem "HUCO"
    .AddItem "KYB"
    .AddItem "LEM"
    .AddItem "LOBRO"
    .AddItem "MANN"
    .AddItem "NARVA"
    .AddItem "NGK"
    .AddItem "NRF"
    .AddItem "PHILIPS"
    .AddItem "REINZ"
    .AddItem "RUVILLE"
    .AddItem "SACHS"
    .AddItem "SCHAEFFLER"
    .AddItem "SKF"
    .AddItem "SWAG"
    .AddItem "TEXTAR"
    .AddItem "TRW"
    .AddItem "WAHLER"
    .AddItem "VARTA"
    .AddItem "PURFLUX"
    .AddItem "PIERBURG"
    .AddItem "VALEO"

    
End With


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, _
  CloseMode As Integer)
  
  'Returns a textbox when the user click on the userform close 'X' at the top of the userform
  
  If CloseMode = vbFormControlMenu Then
  
    Cancel = True
    
    MsgBox "Моля използвайте бутона затвори!"
    
  End If
  
End Sub
Private Sub UserForm_Click()

End Sub

Sub Brands()

'Applies formulas according to the brand chosen by the user.

Dim cel As Range
Dim selectedRange As Range
Set selectedRange = Application.Selection
 
    If Brand = "ATE" Then
    
         For Each cel In selectedRange.Cells
            
            cel = "=CONCATENATE(""ATE"","" "",LEFT(RC[-1],2),""."",MID(RC[-1],4,4),""-"",MID(RC[-1],9,4),""."",MID(RC[-1],14,1))"
        
        Next cel
        
    End If

    If Brand = "BERU" Then
    
         For Each cel In selectedRange.Cells
      
           cel = "=IFERROR(CONCATENATE(""BERU"","" "",LEFT(RC[-1],(MIN(FIND({0,1,2,3,4,5,6,7,8,9},RC[-1]&""0123456789"")))-1), "" "",RIGHT(RC[-1],LEN(RC[-1])-(MIN(FIND({0,1,2,3,4,5,6,7,8,9},RC[-1]&""0123456789"")))+1)),"""")"
        
        Next cel

    End If
    
    If Brand = "BILSTEIN" Then
    
         For Each cel In selectedRange.Cells
      
           cel = "=CONCATENATE(""BILSTEIN"","" "",LEFT(RC[-1],2),""-"",RIGHT(RC[-1],6))"
        
        Next cel

    End If

    
    If Brand = "BOSCH" Then
    
         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""BOSCH"","" "",LEFT(RC[-1],1),"" "",MID(RC[-1],2,3),"" "",MID(RC[-1],5,3),"" "",MID(RC[-1],8,3))"

        Next cel
        
    End If

    If Brand = "BREMBO DISCS" Then
    
         For Each cel In selectedRange.Cells
         
            cel = "=CONCATENATE(""BREMBO"","" "",LEFT(RC[-1],2),""."",MID(RC[-1],3,4),""."",MID(RC[-1],7,2))"
              
         Next cel
         
    End If
    
    If Brand = "BREMBO BRAKE PADS" Then
    
         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""BREMBO"","" "",LEFT(RC[-1],1),"" "",MID(RC[-1],2,2),"" "",MID(RC[-1],4,3))"
            
        Next cel
      
    End If

    If Brand = "CONTITECH" Then
    
         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""CONTITECH"","" "",RC[-1])"
                     
         Next cel
      
    End If

    If Brand = "CORTECO" Then
    
         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""CORTECO"","" "",RC[-1])"
                     
         Next cel
      
    End If
    
    If Brand = "DAYCO" Then

         For Each cel In selectedRange.Cells
         
            cel = "=CONCATENATE(""DAYCO"","" "",RC[-1])"
                     
        Next cel

    End If
    
    If Brand = "DELPHI" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""DELPHI"","" "",RC[-1])"
                
        Next cel

    End If

    If Brand = "DENSO" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""DENSO"","" "",RC[-1])"

        Next cel

    End If

    If Brand = "DOLZ" Then

          For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""DOLZ"","" "",RC[-1])"
 
        Next cel

    End If

    If Brand = "FAI" Then

        For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""FAI"","" "",RC[-1])"
             
         Next cel

    End If

    If Brand = "GATES" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""GATES"","" "",RC[-1])"
               
        Next cel
    
    End If
    
    If Brand = "HELLA" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""HELLA"","" "",LEFT(RC[-1],3),"" "",MID(RC[-1],4,3),"" "",MID(RC[-1],7,3),""-"",MID(RC[-1],10,3))"
                    
        Next cel

    End If

    If Brand = "HUCO" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""HUCO"","" "",RC[-1])"
         
        Next cel
      
    End If
      
    If Brand = "KYB" Then
           
         For Each cel In selectedRange.Cells
         
            cel = "=CONCATENATE(""KYB"","" "",RC[-1])"
        
        Next cel

    End If
    
    If Brand = "LEM" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""LEM"","" "",LEFT(RC[-1],5),"" "",RIGHT(RC[-1],2))"
         
         Next cel

    End If
    
    If Brand = "LOBRO" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""LOBRO"","" "",RC[-1])"
           
        Next cel

    End If
    
    If Brand = "MANN" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""MANN"","" "",RC[-1])"
           
        Next cel

    End If
    
    If Brand = "NARVA" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""NARVA"","" "",RC[-1])"
           
        Next cel
      
    End If

    If Brand = "NGK" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""NGK"","" "",TRIM(RIGHT(SUBSTITUTE(RC[-1],"" "",REPT("" "",255)),255)))"
           
        Next cel

    End If
    
    If Brand = "NRF" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""NRF"","" "",RC[-1])"
           
        Next cel

    End If
    
    If Brand = "PHILIPS" Then

          For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""PHILIPS"","" "",RC[-1])"
           
         Next cel
      
    End If

    If Brand = "REINZ" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""REINZ"","" "",RC[-1])"
           
         Next cel

    End If
    
    If Brand = "RUVILLE" Then

          For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""RUVILLE"","" "",RC[-1])"
           
          Next cel

    End If
    
    If Brand = "SACHS" Then

         For Each cel In selectedRange.Cells
       
            cel = "=IF(LEN(RC[-1])=6,CONCATENATE(""SACHS"","" "",LEFT(RC[-1],3),"" "",RIGHT(RC[-1],3)),CONCATENATE(""SACHS"","" "",LEFT(RC[-1],4),"" "",MID(RC[-1],5,3),"" "",MID(RC[-1],8,3)))"
           
        Next cel

    End If
    
    If Brand = "SCHAEFFLER" Then

          For Each cel In selectedRange.Cells
       
            cel = "=IF(ISTEXT(RC[-1]),CONCATENATE(""FAG"","" "",RC[-1]),IF(LEFT(RC[-1],3)=""713"",CONCATENATE(""FAG"","" "",LEFT(RC[-1],3),"" "",MID(RC[-1],4,4),"" "",MID(RC[-1],8,2)),IF(LEFT(RC[-1],2)=""53"",CONCATENATE(""INA"","" "",LEFT(RC[-1],3),"" "",MID(RC[-1],4,4),"" "",MID(RC[-1],8,2)),IF(LEFT(RC[-1],2)=""42"",CONCATENATE(""INA"","" "",LEFT(RC[-1],3),"" "",MID(RC[-1],4,4),"" "",MID(RC[-1],8,2)),CONCATENATE(""LUK"","" "",LEFT(RC[-1],3),"" "",MID(RC[-1],4,4),"" "",MID(RC[-1],8,2))))))"
           
         Next cel

    End If
    
    If Brand = "SKF" Then

          For Each cel In selectedRange.Cells
       
            cel = "=IF(ISERROR(FIND("" "",RC[-1],1)),CONCATENATE(""SKF"","" "",LEFT(RC[-1],(MIN(FIND({0,1,2,3,4,5,6,7,8,9},RC[-1]&""0123456789"")))-1), "" "",RIGHT(RC[-1],LEN(RC[-1])-(MIN(FIND({0,1,2,3,4,5,6,7,8,9},RC[-1]&""0123456789"")))+1)),CONCATENATE(""SKF"","" "",RC[-1]))"
           
        Next cel

    End If
    
    If Brand = "SWAG" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""SWAG"","" "",LEFT(RC[-1],2),"" "",MID(RC[-1],3,2),"" "",MID(RC[-1],5,4))"
           
         Next cel

    End If
    
    If Brand = "TEXTAR" Then

           For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""TEXTAR"","" "",RC[-1])"
           
         Next cel

    End If
    
    If Brand = "TRW" Then

          For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""TRW"","" "",RC[-1])"
           
        Next cel

    End If
    
    If Brand = "WAHLER" Then

          For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""WAHLER"","" "",RC[-1])"
           
         Next cel

    End If
    
    If Brand = "VARTA" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""VARTA"","" "",RC[-1])"
           
         Next cel

    End If
    
    If Brand = "PURFLUX" Then

         For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""PURFLUX"","" "",RC[-1])"
           
         Next cel

    End If
    
    If Brand = "PIERBURG" Then

         For Each cel In selectedRange.Cells
       
            cel = "=IF(ISERROR(FIND(""."",RC[-1],1)),CONCATENATE(""PIERBURG"","" "",LEFT(RC[-1],1),""."",MID(RC[-1],2,5),""."",MID(RC[-1],7,2),""."",MID(RC[-1],9,1)),CONCATENATE(""PIERBURG"","" "",RC[-1]))"
           
        Next cel
        
    End If
    
    If Brand = "VALEO" Then

          For Each cel In selectedRange.Cells
       
            cel = "=CONCATENATE(""VALEO"","" "",TEXT(RC[-1],""000000""))"
           
          Next cel

    End If
    
End Sub
