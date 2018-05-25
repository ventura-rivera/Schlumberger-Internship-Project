VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Explosives Field Count"
   ClientHeight    =   11160
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9252
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    
    If TextBox1.Text = "" Or TextBox2.Text = "" Then
        MsgBox "Provide the inputs"
        Exit Sub
    End If
    
    CommandButton1.Caption = "Include Input"
    
    Dim DTnum As Double
    Dim CPSTnum As Double
    Dim PCnum As Double
    
    DTnum = TextBox1.Value * TextBox2.Value
    CPSTnum = TextBox2.Value
    PCnum = TextBox2.Value
    
    With ListBox1
        If .ListCount = 0 Then
            .ColumnCount = 2
            .AddItem ("Detonators")
            .AddItem ("CPSTs")
            .AddItem ("Power Charges")
            .AddItem
            .List(0, 1) = DTnum
            .List(1, 1) = CPSTnum
            .List(2, 1) = PCnum
        Else
            DTnum = DTnum + CDbl(.List(0, 1))
            CPSTnum = CPSTnum + CDbl(.List(1, 1))
            PCnum = PCnum + CDbl(.List(2, 1))
            .List(0, 1) = DTnum
            .List(1, 1) = CPSTnum
            .List(2, 1) = PCnum
        End If
    End With
    
    

End Sub

Private Sub CommandButton2_Click()

    ListBox1.Clear
    CommandButton1.Caption = "Start Input"
    
    TextBox1.Text = ""
    TextBox2.Text = ""
    
End Sub

Private Sub CommandButton3_Click()
    
    If TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
        MsgBox "Provide the inputs"
        Exit Sub
    End If
    
    CommandButton3.Caption = "Include Input"
    
    Dim DTnum As Double
    Dim CPSTnum As Double
    Dim PCnum As Double
    
    Dim DTsub As Double
    Dim CPSTsub As Double
    Dim PCsub As Double
    
    DTsub = TextBox3.Value * TextBox7.Value
    CPSTsub = TextBox7.Value
    PCsub = TextBox7.Value
    
    With ListBox2
        If .ListCount = 0 Then
            DTnum = TextBox4.Value
            CPSTnum = TextBox5.Value
            PCnum = TextBox6.Value
            
            .ColumnCount = 2
            .AddItem ("Detonators")
            .AddItem ("CPSTs")
            .AddItem ("Power Charges")
            .AddItem
            .List(0, 1) = DTnum - DTsub
            .List(1, 1) = CPSTnum - CPSTsub
            .List(2, 1) = PCnum - PCsub
            TextBox4.Enabled = False
            TextBox4.BackColor = &HE0E0E0
            TextBox5.Enabled = False
            TextBox5.BackColor = &HE0E0E0
            TextBox6.Enabled = False
            TextBox6.BackColor = &HE0E0E0
        Else
            .List(0, 1) = CDbl(.List(0, 1)) - DTsub
            .List(1, 1) = CDbl(.List(1, 1)) - CPSTsub
            .List(2, 1) = CDbl(.List(2, 1)) - PCsub
        End If
    End With
    
    
End Sub

Private Sub CommandButton4_Click()
    
    ListBox2.Clear
    TextBox4.Enabled = True
    TextBox4.BackColor = &H80000005
    TextBox5.Enabled = True
    TextBox5.BackColor = &H80000005
    TextBox6.Enabled = True
    TextBox6.BackColor = &H80000005
    
    TextBox4.Text = ""
    TextBox5.Text = ""
    TextBox6.Text = ""
    TextBox3.Text = ""
    TextBox7.Text = ""
    
    CommandButton3.Caption = "Start Input"

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

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub
Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub
Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub
Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub UserForm_Click()

End Sub
