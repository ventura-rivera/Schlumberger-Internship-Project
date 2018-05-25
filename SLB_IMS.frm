VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SLB_IMS 
   Caption         =   "SLB Inventory Management System"
   ClientHeight    =   10584
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   20364
   OleObjectBlob   =   "SLB_IMS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SLB_IMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()

    If Controls("Checkbox1").Value = False Then
        Controls("CheckBox9").Value = False
    End If
    
End Sub

Private Sub CheckBox10_Click()

    If Controls("CheckBox10").Value = True Then
        ItemPlot.Text = "All Items"
        SpendingPlot.Text = "All Items"
        ItemPlot.Enabled = False
        ItemPlot.BackColor = &HE0E0E0
        SpendingPlot.Enabled = False
        SpendingPlot.BackColor = &HE0E0E0
    ElseIf Controls("CheckBox10").Value = False Then
        ItemPlot.Enabled = True
        ItemPlot.BackColor = &H80000005
        SpendingPlot.Enabled = True
        SpendingPlot.BackColor = &H80000005
    End If

End Sub

Private Sub CheckBox2_Click()

    If Controls("Checkbox2").Value = False Then
        Controls("CheckBox9").Value = False
    End If

End Sub

Private Sub CheckBox3_Click()

    If Controls("Checkbox3").Value = False Then
        Controls("CheckBox9").Value = False
    End If

End Sub

Private Sub CheckBox4_Click()

    If Controls("Checkbox4").Value = False Then
        Controls("CheckBox9").Value = False
    End If

End Sub

Private Sub CheckBox5_Click()

    If Controls("Checkbox5").Value = False Then
        Controls("CheckBox9").Value = False
    End If

End Sub

Private Sub CheckBox6_Click()

    If Controls("Checkbox6").Value = False Then
        Controls("CheckBox9").Value = False
    End If

End Sub

Private Sub CheckBox7_Click()

    If Controls("Checkbox7").Value = False Then
        Controls("CheckBox9").Value = False
    End If

End Sub

Private Sub CheckBox8_Click()

    If Controls("Checkbox8").Value = False Then
        Controls("CheckBox9").Value = False
    End If

End Sub

Private Sub CheckBox9_Click()

    Dim countcheck As Integer
    countcheck = 0
    For i = 1 To 8
        If Controls("CheckBox" & i).Value = True Then
            countcheck = countcheck + 1
        End If
    Next
        
    
    
    If Controls("CheckBox9").Value = True Then
        For i = 1 To 8
            Controls("CheckBox" & i).Value = True
        Next
    ElseIf Controls("CheckBox9").Value = False And countcheck = 8 Then
        For i = 1 To 8
            Controls("CheckBox" & i).Value = False
        Next
    End If
        
    
    
End Sub

Private Sub CommandButton2_Click()
    
    If Truck.Text = "Select a truck..." Then
        MsgBox "Select a truck from the dropdown list"
        Exit Sub
    End If
    
    Dim LastRow As Long
    Dim rng As Range
    
    If Truck.Text = "PTX1" Then
        sh = "PTX1"
    ElseIf Truck.Text = "PTX2" Then
        sh = "PTX2"
    ElseIf Truck.Text = "PTX3" Then
        sh = "PTX3"
    ElseIf Truck.Text = "PTX4" Then
        sh = "PTX4"
    ElseIf Truck.Text = "PTX5" Then
        sh = "PTX5"
    ElseIf Truck.Text = "PTX6" Then
        sh = "PTX6"
    ElseIf Truck.Text = "PTX7" Then
        sh = "PTX7"
    ElseIf Truck.Text = "PTX8" Then
        sh = "PTX8"
    ElseIf Truck.Text = "All Trucks" Then
        sh = "All"
    Else
        MsgBox "Select a valid truck from the dropdown list"
        Exit Sub
    End If
    
    Set rng = Sheets(sh).Cells
    LastRow = Last(1, rng)
    ListBox1.RowSource = sh & "!A2:G" & LastRow
    

End Sub

Private Sub CommandButton3_Click()

    UserForm1.Show

End Sub

Private Sub CommandButton4_Click()

    UserForm2.Show

End Sub


''''''''''''''''''''''''''''
'''''''''PLOT STUFF'''''''''
''''''''''''''''''''''''''''


Private Sub CommandButton5_Click()
    
    Dim count As Integer
    count = 0
    For i = 1 To 8
        If Controls("CheckBox" & i).Value = True Then
            count = 1
        End If
    Next
    
    If count = 0 Then
        MsgBox "Select one or more trucks with the checkboxes"
        Exit Sub
    End If
    
    If DTPicker1.Value > DTPicker2.Value Then
        MsgBox "Your start date is past your end date"
        Exit Sub
    End If
    
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""
    TextBox5.Value = ""
    'TextBox4.Value = ""
    'TextBox6.Value = ""
    
    UserForm3.Show

End Sub


Private Sub CommandButton6_Click()
    
    Dim count As Integer
    count = 0
    For i = 1 To 8
        If Controls("CheckBox" & i).Value = True Then
            count = 1
        End If
    Next
    
    If count = 0 Then
        MsgBox "Select one or more trucks with the checkboxes"
        Exit Sub
    End If
        
    If DTPicker1.Value > DTPicker2.Value Then
        MsgBox "Your start date is past your end date"
        Exit Sub
    End If
    
    
    
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""
    TextBox5.Value = ""
    'TextBox4.Value = ""
    'TextBox6.Value = ""
    
    UserForm4.Show

End Sub




'''''''''''''''''''''''''''''
''''''''END PLOT STUFF'''''''
'''''''''''''''''''''''''''''


Private Sub UserForm_Activate()

    Truck.Text = "Select a truck..."
    Truck.AddItem ("PTX1")
    Truck.AddItem ("PTX2")
    Truck.AddItem ("PTX3")
    Truck.AddItem ("PTX4")
    Truck.AddItem ("PTX5")
    Truck.AddItem ("PTX6")
    Truck.AddItem ("PTX7")
    Truck.AddItem ("PTX8")
    
    ItemPlot.Text = "All Items"
    ItemPlot.AddItem ("Boxes of Rags")
    ItemPlot.AddItem ("Brake Cleaner")
    ItemPlot.AddItem ("Cases of Water")
    ItemPlot.AddItem ("Boxes of Rags")
    ItemPlot.AddItem ("Contact Cleaner")
    ItemPlot.AddItem ("Latex Gloves")
    ItemPlot.AddItem ("O-Rings (217)")
    ItemPlot.AddItem ("O-Rings (218)")
    ItemPlot.AddItem ("O-Rings (222)")
    ItemPlot.AddItem ("O-Rings (230)")
    ItemPlot.AddItem ("O-Rings (231)")
    ItemPlot.AddItem ("Paper")
    ItemPlot.AddItem ("Printer Ink")
    ItemPlot.AddItem ("Scotch Locks (2-way)")
    ItemPlot.AddItem ("Scotch Locks (3-way)")
    ItemPlot.AddItem ("Setting Tool Oil")
    ItemPlot.AddItem ("Trash Bags")
    ItemPlot.AddItem ("WD-40")
    
    ItemPlot.AddItem ("Electrical Tape???")
    ItemPlot.AddItem ("Wire???")
    ItemPlot.AddItem ("Switches???")
    ItemPlot.AddItem ("Explosives???")
    ItemPlot.AddItem ("Grease???")
    ItemPlot.AddItem ("Peanut Butter???")
    
    SpendingPlot.Text = "All Items"
    SpendingPlot.AddItem ("Boxes of Rags")
    SpendingPlot.AddItem ("Brake Cleaner")
    SpendingPlot.AddItem ("Cases of Water")
    SpendingPlot.AddItem ("Contact Cleaner")
    SpendingPlot.AddItem ("Impact Gloves")
    SpendingPlot.AddItem ("Latex Gloves")
    SpendingPlot.AddItem ("O-Rings (217)")
    SpendingPlot.AddItem ("O-Rings (218)")
    SpendingPlot.AddItem ("O-Rings (222)")
    SpendingPlot.AddItem ("O-Rings (230)")
    SpendingPlot.AddItem ("O-Rings (231)")
    SpendingPlot.AddItem ("Paper")
    SpendingPlot.AddItem ("Printer Ink")
    SpendingPlot.AddItem ("Scotch Locks (2-way)")
    SpendingPlot.AddItem ("Scotch Locks (3-way)")
    SpendingPlot.AddItem ("Setting Tool Oil")
    SpendingPlot.AddItem ("Trash Bags")
    SpendingPlot.AddItem ("WD-40")
    
    SpendingPlot.AddItem ("Electrical Tape???")
    SpendingPlot.AddItem ("Wire???")
    SpendingPlot.AddItem ("Switches???")
    SpendingPlot.AddItem ("Explosives???")
    SpendingPlot.AddItem ("Grease???")
    SpendingPlot.AddItem ("Peanut Butter???")
    
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    

End Sub

Function Last(choice As Long, rng As Range)
'Ron de Bruin, 5 May 2008
' 1 = last row
' 2 = last column
' 3 = last cell
    Dim lrw As Long
    Dim lcol As Long

    Select Case choice

    Case 1:
        On Error Resume Next
        Last = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
        On Error GoTo 0

    Case 2:
        On Error Resume Next
        Last = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

    Case 3:
        On Error Resume Next
        lrw = rng.Find(What:="*", _
                       After:=rng.Cells(1), _
                       Lookat:=xlPart, _
                       LookIn:=xlFormulas, _
                       SearchOrder:=xlByRows, _
                       SearchDirection:=xlPrevious, _
                       MatchCase:=False).Row
        On Error GoTo 0

        On Error Resume Next
        lcol = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

        On Error Resume Next
        Last = rng.Parent.Cells(lrw, lcol).Address(False, False)
        If Err.Number > 0 Then
            Last = rng.Cells(1).Address(False, False)
            Err.Clear
        End If
        On Error GoTo 0

    End Select
End Function

