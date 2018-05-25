VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Field Item Calculator"
   ClientHeight    =   10080
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8004
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
    
    If Controls("CheckBox1").Value = True Then
        TextBox5.Locked = False
    ElseIf Controls("CheckBox1").Value = False Then
        TextBox5.Locked = True
        TextBox5.Value = 0
    End If
    
End Sub

Private Sub cmbItem_Change()
    TextBoxesEqns
End Sub

Private Sub CommandButton1_Click()

    If cmbItem.Text = "Select an item..." Then
        MsgBox "Select an item from the dropdown list"
        Exit Sub
    End If
    
    'columns: item, qty, cost
    With ListBox1
        i = .ListCount
        ListBox1.ColumnCount = 3
        .ColumnWidths = "180;80;80"
        .AddItem
        .List(i, 0) = cmbItem.Text
        .List(i, 1) = TextBox5.Value
        .List(i, 2) = TextBox6.Value
    End With
    
    Dim TotalCost As Double
    TotalCost = 0
    With ListBox1
        If .ListCount > 1 Then
            For n = 0 To .ListCount - 1
                TotalCost = TotalCost + .List(n, 2)
            Next
        ElseIf .ListCount = 1 Then
            TotalCost = .List(0, 2)
        End If
    End With
    
    TextBox4.Value = TotalCost
    
End Sub


Private Sub CommandButton2_Click()
          
    Dim NewCost As Double
    If ListBox1.ListIndex >= 0 Then
        NewCost = TextBox4.Value - ListBox1.List(ListBox1.ListIndex, 2)
        TextBox4.Value = NewCost
        ListBox1.RemoveItem ListBox1.ListIndex
    End If
    
End Sub

Private Sub TextBox1_Change()
    TextBoxesEqns
End Sub



Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub
Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub TextBox5_Change()
    
    If TextBox5.Value = "" Then
        TextBox6.Value = 0
        Exit Sub
    End If
    TextBoxesEqns
    
End Sub

Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub TextBox2_Change()
    TextBoxesEqns
End Sub



Private Sub TextBox3_Change()
    TextBoxesEqns
End Sub

Private Sub TextBox7_Change()
    TextBoxesEqns
End Sub

Private Sub UserForm_Activate()

    cmbItem.Text = "Select an item..."
    cmbItem.AddItem ("Boxes of rags")
    cmbItem.AddItem ("Brake Cleaner")
    cmbItem.AddItem ("Cases of Water")
    cmbItem.AddItem ("Contact Cleaner")
    cmbItem.AddItem ("Impact Gloves")
    cmbItem.AddItem ("Latex Gloves")
    cmbItem.AddItem ("O-Rings (217)")
    cmbItem.AddItem ("O-Rings (218)")
    cmbItem.AddItem ("O-Rings (222)")
    cmbItem.AddItem ("O-Rings (230)")
    cmbItem.AddItem ("O-Rings (231)")
    cmbItem.AddItem ("Paper")
    cmbItem.AddItem ("Printer Ink")
    cmbItem.AddItem ("Scotch Locks (2-way)")
    cmbItem.AddItem ("Scotch Locks (3-way)")
    cmbItem.AddItem ("Setting Tool Oil")
    cmbItem.AddItem ("Trash Bags")
    cmbItem.AddItem ("WD-40")
    
    cmbItem.AddItem ("Electrical Tape???") 'yes
    cmbItem.AddItem ("Wire???")     'yes
    cmbItem.AddItem ("Grease???")   'yes
    cmbItem.AddItem ("Peanut Butter???")    'yes
    
    TextBox4.Value = 0

End Sub


Private Sub TextBoxesEqns()
    Dim Total As Double
    Dim chartindex As Integer
    
    Total = 0
    
    
    If cmbItem.Text = "Boxes of rags" Then
        If Len(TextBox1.Value) > 0 Then Total = CDbl(Int(CDbl(TextBox1.Value) / CDbl(20)) + 1)
        
    ElseIf cmbItem.Text = "Brake Cleaner" Then
        If Len(TextBox1.Value) > 0 Then Total = CDbl(Int(CDbl(TextBox1.Value) / CDbl(5)) + 1)
        
    ElseIf cmbItem.Text = "Cases of Water" Then
        If Len(TextBox3.Value) > 0 And Len(TextBox2.Value) > 0 Then Total = CDbl(Int(CDbl(TextBox3.Value) * CDbl(TextBox2.Value) / CDbl(3)) + 1)
        
    ElseIf cmbItem.Text = "Contact Cleaner" Then
        If Len(TextBox1.Value) > 0 Then Total = CDbl(Int(CDbl(TextBox1.Value) / CDbl(5)) + 1)
        
    ElseIf cmbItem.Text = "Impact Gloves" Then
        If Len(TextBox1.Value) > 0 And Len(TextBox3.Value) > 0 Then Total = CDbl(Int((CDbl(TextBox1.Value) * CDbl(TextBox3.Value) / CDbl(20))) + 1)
        
    ElseIf cmbItem.Text = "Latex Gloves" Then
        If Len(TextBox1.Value) > 0 And Len(TextBox3.Value) > 0 Then Total = CDbl(Int(CDbl(TextBox1.Value) * CDbl(TextBox3.Value) / CDbl(50)) + 1)
        
    ElseIf cmbItem.Text = "O-Rings (217)" Then
        If Len(TextBox1.Value) > 0 Then Total = CDbl(TextBox1.Value) * CDbl(2)
        
    ElseIf cmbItem.Text = "O-Rings (218)" Then
        If Len(TextBox1.Value) > 0 Then Total = CDbl(TextBox1.Value) * CDbl(2)
        
    ElseIf cmbItem.Text = "O-Rings (222)" Then
        If Len(TextBox1.Value) > 0 Then Total = CDbl(TextBox1.Value) * CDbl(2)
        
    ElseIf cmbItem.Text = "O-Rings (230)" Then
        If Len(TextBox1.Value) > 0 Then Total = CDbl(TextBox1.Value) * CDbl(8)
        
    ElseIf cmbItem.Text = "O-Rings (231)" Then
        If Len(TextBox1.Value) > 0 Then Total = CDbl(TextBox1.Value) * CDbl(8)
        
    ElseIf cmbItem.Text = "Paper" Then
        Total = CDbl(1)
        
    ElseIf cmbItem.Text = "Printer Ink" Then
        Total = CDbl(1)
        
    ElseIf cmbItem.Text = "Scotch Locks (2-way)" Then
        Total = CDbl(2)
        
    ElseIf cmbItem.Text = "Scotch Locks (3-way)" Then
        Total = CDbl(2)
        
    ElseIf cmbItem.Text = "Setting Tool Oil" Then
        If Len(TextBox1.Value) > 0 Then Total = CDbl(Int(CDbl(TextBox1.Value) / CDbl(5)) + 1)
        
    ElseIf cmbItem.Text = "Trash Bags" Then
        If Len(TextBox2.Value) > 0 Then Total = CDbl(Int(CDbl(6) * CDbl(TextBox2.Value) / CDbl(100)))
        
    ElseIf cmbItem.Text = "WD-40" Then
        If Len(TextBox1.Value) > 0 Then Total = CDbl(Int(CDbl(TextBox1.Value) / CDbl(5)) + 1)
        
    End If
    
    If Controls("CheckBox1").Value = False Then
        TextBox5.Value = Total
    End If
    
    TextBox6.Value = CDbl(0)
    
    If Len(TextBox7.Value) > 0 Then
        TextBox6.Value = CDbl(TextBox7.Value) * CDbl(TextBox5.Value)
    End If
    
End Sub

Private Function NumericOnly(ByVal KeyAscii As MSForms.ReturnInteger) As MSForms.ReturnInteger
    Dim Key As MSForms.ReturnInteger
    Select Case KeyAscii
        Case 46, 48 To 57 ' Accept only decimal "." and numbers [0-9]
            Set Key = KeyAscii
        Case Else
            KeyAscii = 0 ' Minor bug earlier
            Set Key = KeyAscii
    End Select
    Set NumericOnly = Key
End Function


