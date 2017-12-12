'MUS'

Sub GenSample()
'
' GenSample Macro
' Macro recorded 11/28/2003 by Michelle McEacharn
'
' Keyboard Shortcut: Ctrl+s
'

Sheets("SamplePrint").Select

' To Declare range variables and set start values
  
    Dim rngCell1 As Range, numCell As Range

    Set rngCell1 = Range("AA1")
    Set numCell = Range("AA2")

    Columns("A:AC").AutoFit
    Columns("A:AC").Select
    Selection.ClearContents
    
    Counter = 0

'To determine sampling interval
    
    Sheets("Worksheet").Select
    PopSize = Range("b9")
    SmpSize = Range("g9")
    Interval = PopSize / SmpSize
    Lower = 0
    
'To determine random starting point

    Sheets("SamplePrint").Select
    Randomize
    Beg = Int((Interval - Lower + 1) * Rnd + Lower)
    rngCell1.Value = Beg
    Set rngCell1 = rngCell1.Offset(rowoffset:=1)
 
 
 'To generate sample dollars
   
    Counter = 1
    Do Until Beg = PopSize Or Beg > PopSize Or Counter = SmpSize
        Beg = Beg + Interval
        rngCell1.Value = Beg
        Set rngCell1 = rngCell1.Offset(rowoffset:=1)
        Counter = Counter + 1
    Loop


'To clear out and set up print worksheet
   
   Sheets("SamplePrint").Select
    Columns("A:L").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = True
        .MergeCells = False
    End With
    


'To number the sample dollars
 
    Sheets("SamplePrint").Select
    Range("Z1").Value = 1
        'Set numCell = Range("AA2")
        CntCell = 1
        Do Until numCell.Value = ""
            CntCell = CntCell + 1
            numCell.Offset(columnoffset:=-1).Value = CntCell
            Set numCell = numCell.Offset(rowoffset:=1)
        Loop
    
    
'To copy sample dollars to print worksheet
         
            
    'Copies first 50 dollar items
        
        Sheets("SamplePrint").Select
        Range("Z1:AA50").Select
        Selection.Copy
        Range("A2").Select
        ActiveSheet.Paste
        Range("A1").Value = "#"
        Range("B1").Value = "Sample Dollar"
        Range("C1").Value = "Item"
        Range("D1").Value = "Error Found?"
    
    'Copies sample items 51 - 100 if needed
    
        If CntCell > 50 Then
            Range("Z51:AA100").Select
            Selection.Copy
            Range("F2").Select
            ActiveSheet.Paste
            Range("F1").Value = "#"
            Range("G1").Value = "Sample Dollar"
            Range("H1").Value = "Item"
            Range("I1").Value = "Error Found?"
        End If
    
    'Copies sample items 101 - 150 if needed
        
        If CntCell > 100 Then
            Range("Z101:AA150").Select
            Selection.Copy
            Range("A53").Select
            ActiveSheet.Paste
            Range("A52").Value = "#"
            Range("B52").Value = "Sample Dollar"
            Range("C52").Value = "Item"
            Range("D52").Value = "Error Found?"
        End If
    
    'Copies sample items 151 - 200 if needed
        
        If CntCell > 150 Then
            Range("Z151:AA200").Select
            Selection.Copy
            Range("F53").Select
            ActiveSheet.Paste
            Range("F52").Value = "#"
            Range("G52").Value = "Sample Dollar"
            Range("H52").Value = "Item"
            Range("I52").Value = "Error Found?"
        End If
    
    
    'Copies sample items 201 - 250 if needed
        
        If CntCell > 200 Then
            Range("Z201:AA250").Select
            Selection.Copy
            Range("A104").Select
            ActiveSheet.Paste
            Range("A103").Value = "#"
            Range("B103").Value = "Sample Dollar"
            Range("C103").Value = "Item"
            Range("D103").Value = "Error Found?"
        End If
    
        
    'Copies sample items 251 - 300 if needed
        
        If CntCell > 250 Then
            Range("Z251:AA300").Select
            Selection.Copy
            Range("F104").Select
            ActiveSheet.Paste
            Range("F103").Value = "#"
            Range("G103").Value = "Sample Dollar"
            Range("H103").Value = "Item"
            Range("I103").Value = "Error Found?"
        End If


    'Copies sample items 301 - 350 if needed
    
        If CntCell > 300 Then
            Range("Z301:AA350").Select
            Selection.Copy
            Range("A155").Select
            ActiveSheet.Paste
            Range("A154").Value = "#"
            Range("B154").Value = "Sample Dollar"
            Range("C154").Value = "Item"
            Range("D154").Value = "Error Found?"
        End If
    
        
    'Copies sample items 351 - 400 if needed
        
        If CntCell > 350 Then
            Range("Z351:AA400").Select
            Selection.Copy
            Range("F155").Select
            ActiveSheet.Paste
            Range("F154").Value = "#"
            Range("G154").Value = "Sample Dollar"
            Range("H154").Value = "Item"
            Range("I154").Value = "Error Found?"
        End If
   
    'Copies sample items 401 - 450 if needed
    
        If CntCell > 400 Then
            Range("Z401:AA450").Select
            Selection.Copy
            Range("A206").Select
            ActiveSheet.Paste
            Range("A205").Value = "#"
            Range("B205").Value = "Sample Dollar"
            Range("C205").Value = "Item"
            Range("D205").Value = "Error Found?"
        End If
    
        
    'Copies sample items 451 - 500 if needed
        
        If CntCell > 450 Then
            Range("Z451:AA500").Select
            Selection.Copy
            Range("F206").Select
            ActiveSheet.Paste
            Range("F205").Value = "#"
            Range("G205").Value = "Sample Dollar"
            Range("H205").Value = "Item"
            Range("I205").Value = "Error Found?"
        End If
        
'To format print page

    Range("B:B,G:G").Select
    Range("G1").Activate
    Selection.Style = "Currency"
    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    Range("C:C,H:H").Select
    Range("H1").Activate
    Selection.ColumnWidth = 11
    Range("D:D,I:I").Select
    Range("I1").Activate
    Selection.ColumnWidth = 14
    Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .MergeCells = False
    End With
    Columns("F:F").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .MergeCells = False
    End With
  

'To clear and set print sheet to focal point
 
    MsgBox "Processing is now complete.  You may print the worksheet by using normal Excel print commands", vbOKOnly, "Finished!"
    Columns("Y:AC").Select
    Selection.ClearContents
    Range("A1").Select
    

End Sub

'MUS'
Sub Results()
'
' Results Macro
' Macro recorded 10/6/2003 by Bruce Wampler
'
' Keyboard Shortcut: Ctrl+r
'
   

    Sheets("Worksheet").Select
    ActiveSheet.Unprotect
    Range("B9:B12").Select
    Selection.Copy
    Sheets("Data Entry").Select
    Range("D5").Select
    ActiveSheet.Paste
    Sheets("Worksheet").Select
    Range("B15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Data Entry").Select
    Range("D9").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("B13:C32").Clear
    Calculate
    MsgBox "Be sure to read and follow the instructions on Row 3 of this worksheet.", vbOKOnly, "Warning!"
    
'    ActiveSheet.Unprotect
'    Calculate
'    Range("B21:D30").Select
'    Range("D21").Activate
'    Selection.Sort Key1:=Range("D21"), Order1:=xlDescending, Header:=xlGuess _
'        , OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
'    Range("B37:D46").Select
'    Range("D37").Activate
'    Selection.Sort Key1:=Range("D21"), Order1:=xlDescending, Header:=xlGuess _
'        , OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

'    Range("B52").Select

 '   ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    

End Sub