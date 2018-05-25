VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   9756
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   16668
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Dim MyChart As Chart
    Dim ChartData As Range
    Dim ChartName As String
    
    ''''''''''''''''''''''''''''''''''''''''
    ''''''''''''THE FOR STARTS''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''
    
    Dim sh As String
    Dim sh_first As String
    Dim good As Integer
    Dim n As Integer
    Dim spending_good As Integer
    
    found_item = 0
    good = 0
    For n = 1 To 8
        If SLB_IMS.Controls("CheckBox" & n).Value = True Then
            
            good = good + 1
            sh = "PTX" & CStr(n)
            If good = 1 Then sh_first = "PTX" & CStr(n)
            Sheets(sh).Activate
    
            Dim LastRow As Long
            Dim rng As Range
            Set rng = ActiveSheet.Cells
            LastRow = Last(1, rng)
            
                
            If good = 1 Then
                Set ChartData = ActiveSheet.Range("E2:E" & CStr(LastRow))
                ChartName = sh
                Application.ScreenUpdating = False
                Set MyChart = ActiveSheet.Shapes.AddChart(xlXYScatterLines).Chart
                MyChart.SeriesCollection.NewSeries
                MyChart.SeriesCollection(good).Name = ChartName
                MyChart.SeriesCollection(good).Values = ChartData
                MyChart.SeriesCollection(good).XValues = ActiveSheet.Range("C2:C" & CStr(LastRow))
            End If
                
            rngXData = ActiveSheet.Range("C2:C" & CStr(LastRow))
            rngYData = ChartData
        
            Dim countTemp As Integer
            Dim count As Integer
            Dim countsum As Range
            Dim i As Integer
            Dim j As Integer
            
            If good = 1 Then
                For i = MyChart.SeriesCollection.count To good Step -1
                    MyChart.SeriesCollection(i).Delete
                Next
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''FOR LOOP - WITH CONDITION - STARTS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            countTemp = 1
            If SLB_IMS.SpendingPlot.Text <> "All Items" Then
                For j = 2 To LastRow
                    If rngXData(j - 1, 1) >= SLB_IMS.DTPicker1.Value And rngXData(j - 1, 1) <= SLB_IMS.DTPicker2.Value And SLB_IMS.SpendingPlot.Text = ActiveSheet.Range("A" & j) Then
                        found_item = 1
                        count = 0
                        For i = j + 1 To LastRow
                            If ActiveSheet.Range("C" & j) = ActiveSheet.Range("C" & i) Then
                                count = 1
                                Exit For
                            End If
                        Next
            
                        If count = 0 Then
                            countTemp = countTemp + 1
                            ActiveSheet.Range("H" & j) = 0
                            For k = 2 To j - 1
                                If ActiveSheet.Range("C" & j) = ActiveSheet.Range("C" & k) And SLB_IMS.SpendingPlot.Text = ActiveSheet.Range("A" & k) Then
                                    ActiveSheet.Range("H" & j) = ActiveSheet.Range("H" & j) + ActiveSheet.Range("E" & k)
                                End If
                            Next
                            ActiveSheet.Range("G" & j) = CDate(ActiveSheet.Range("C" & j))
                            ActiveSheet.Range("H" & j) = ActiveSheet.Range("H" & j) + ActiveSheet.Range("E" & j)
                            
                        End If
                    End If
                Next
            
            Else
                For j = 2 To LastRow
                    If rngXData(j - 1, 1) >= SLB_IMS.DTPicker1.Value And rngXData(j - 1, 1) <= SLB_IMS.DTPicker2.Value Then
            
                        count = 0
                        For i = j + 1 To LastRow
                            If ActiveSheet.Range("C" & j) = ActiveSheet.Range("C" & i) Then
                                count = 1
                                Exit For
                            End If
                        Next
            
                        If count = 0 Then
                            countTemp = countTemp + 1
                            ActiveSheet.Range("H" & j) = 0
                            For k = 2 To j - 1
                                If ActiveSheet.Range("C" & j) = ActiveSheet.Range("C" & k) Then
                                    ActiveSheet.Range("H" & j) = ActiveSheet.Range("H" & j) + ActiveSheet.Range("E" & k)
                                End If
                            Next
                            ActiveSheet.Range("G" & j) = CDate(ActiveSheet.Range("C" & j))
                            ActiveSheet.Range("H" & j) = ActiveSheet.Range("H" & j) + ActiveSheet.Range("E" & j)
                            
                        End If
                    End If
                Next
            End If
            
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''FOR LOOP - WITH CONDITION - ENDS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
        
            Range("G2:H" & LastRow).Select
            ActiveWorkbook.Worksheets(sh).Sort.SortFields.Clear
            ActiveWorkbook.Worksheets(sh).Sort.SortFields.Add Key:=Range("G2:G" & LastRow), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ActiveWorkbook.Worksheets(sh).Sort
                .SetRange Range("G2:H" & LastRow)
                .Header = xlGuess
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        
            With MyChart.SeriesCollection.NewSeries
                .XValues = ActiveSheet.Range("G2:G" & countTemp)
                .Values = ActiveSheet.Range("H2:H" & countTemp)
            End With
            
            With MyChart
                 'X axis name
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Date"
                 'y-axis name
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Amount Spent ($)"
                .SeriesCollection(good).Name = sh
            End With
        End If
    Next
    
    Sheets(sh_first).Activate
    ActiveSheet.ChartObjects(1).Activate
    ActiveSheet.ChartObjects(1).Activate
    ActiveSheet.Shapes(1).Height = 444
    ActiveSheet.Shapes(1).Width = 720
    ActiveSheet.ChartObjects(1).Activate
    
    Dim imageName As String
    imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"

    MyChart.Export Filename:=imageName, FilterName:="GIF"

    ActiveSheet.ChartObjects(1).Delete

    Application.ScreenUpdating = True

    UserForm4.Image1.Picture = LoadPicture(imageName)
    

    If good = 1 And SLB_IMS.SpendingPlot.Text = "All Items" Then
        SLB_IMS.TextBox1.Value = CDbl(Round(CDbl(Application.sum(Sheets(sh).Range("H2:H" & countTemp))) / CDbl(countTemp - 1), 2))
        SLB_IMS.TextBox2.Value = CDbl(Round(CDbl((ActiveSheet.Range("H" & countTemp) - ActiveSheet.Range("H2")) / CDbl((ActiveSheet.Range("G" & countTemp) - ActiveSheet.Range("G2")))), 2))
        SLB_IMS.TextBox3.Value = CDbl(Application.Max(Sheets(sh).Range("H2:H" & countTemp)))
        SLB_IMS.TextBox5.Value = CDbl(Application.Min(Sheets(sh).Range("H2:H" & countTemp)))
    '    Dim icount As Variant
    '    Columns("G:G").Select
    '    Application.CutCopyMode = False
    '    Selection.NumberFormat = "m/d/yyyy"
    '    ' for the date if i wanna go for that?
    '    icount = Application.Match(Application.Max(Sheets(sh).Range("H2:H" & countTemp)), Sheets(sh).Range("H2:H" & countTemp), 0)
    '    SLB_IMS.TextBox4.Value = CDate(Sheets(sh).Range("G" & CInt(icount)))
    '    Dim dDate As Date
    '    SLB_IMS.TextBox4.Text = Format(dDate, "MM/dd/yyyy")
        
    ElseIf SLB_IMS.SpendingPlot.Text <> "All Items" And found_item = 1 And good = 1 Then
        SLB_IMS.TextBox1.Value = CDbl(Round(CDbl(Application.sum(Sheets(sh).Range("H2:H" & countTemp))) / CDbl(countTemp - 1), 2))
        SLB_IMS.TextBox2.Value = CDbl(Round(CDbl((ActiveSheet.Range("H" & countTemp) - ActiveSheet.Range("H2")) / CDbl((ActiveSheet.Range("G" & countTemp) - ActiveSheet.Range("G2")))), 2))
        SLB_IMS.TextBox3.Value = CDbl(Application.Max(Sheets(sh).Range("H2:H" & countTemp)))
        SLB_IMS.TextBox5.Value = CDbl(Application.Min(Sheets(sh).Range("H2:H" & countTemp)))
        
    End If
    
    Dim x As Integer
    For x = 1 To 8
        Sheets("PTX" & x).Activate
        Columns("G:H").Select
        Selection.Delete Shift:=xlToLeft
    Next
    
    ''''''''''''''''''''''''''''''''''''''''
    ''''''''''''THE FOR ENDS''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''

    
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

