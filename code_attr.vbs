Sub PullSample()
'
' PullSample Macro
' Macro recorded 1/25/2001 by Michelle McEacharn
'
' Keyboard Shortcut: Ctrl+s
'
  
Sheets("SamplePrint").Select

' To Declare range variables and set start value
  
Dim rngCell1 As Range, rngCell2 As Range
Dim numCell As Range, extCell As Range

Set rngCell1 = Range("AA1")
Set numCell = Range("AA2")
Set rngCell2 = Range("AC1")
Set extCell = Range("AC2")


'To clear out and set up print worksheet
   
    Sheets("SamplePrint").Select
    Columns("A:AC").AutoFit
    Columns("A:AC").Select
    Selection.ClearContents
    
    Columns("A:L").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    

' To get variable values

    Sheets("SampleCalc").Select
    Range("B14").Copy
    Sheets("SamplePrint").Select
    Range("Y1").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Number = Range("Y1")
    
    MsgBox "Please answer the next three questions.  Processing will then take a few seconds.  A message will appear when the processing is complete.", vbOKOnly, "Random Number Generation Start"
         
           
    Extra = InputBox("If you would like to generate any extra sample items, type the number below and click 'Ok'", "Extra Sample Items")
        If Extra = "" Then
            Extra = 0
        End If
        If Extra < 0 Then
            MsgBox "You must enter a number greater than 0 to generate any extra random numbers.", vbCritical, "Error"
            Sheets("SampleCalc").Select
            Range("A1").Select
            Exit Sub
        End If
        If Extra > 30 Then
            MsgBox "You may not request more than 30 extra sample items.", vbCritical, "Error"
            Extra = InputBox("Please reenter a number between 0 and 30.", "Extra Sample Items")
                If Extra = "" Then
                    Extra = 0
                End If
                If Extra < 0 Or Extra > 30 Then
                    Sheets("SampleCalc").Select
                    Range("A1").Select
                    Exit Sub
                End If
        End If
    Extra = CInt(Extra)
    
    
    Lower = InputBox("What is the lowest value in the population from which you are pulling your sample?", "Lowest Item Number")
        If Lower = "" Or Lower = 0 Or Lower < 0 Then
            MsgBox "You must enter a number greater than 0 for lowest sample item.", vbCritical, "Error"
            Sheets("SampleCalc").Select
            Range("A1").Select
            Exit Sub
        End If
    Lower = CDbl(Lower)
    
    
    Upper = InputBox("What is the highest value in the population from which you are pulling your sample?", "Highest Item Number")
        If Upper = "" Then
            Upper = 0
            MsgBox "You must enter a number greater than the lowest possible sample item number", vbCritical, "Error"
            Sheets("SampleCalc").Select
            Range("A1").Select
            Exit Sub
        End If
    Upper = CDbl(Upper)
        If Upper < Lower Then
            MsgBox "You must enter a number greater than the lowest possible sample number", vbCritical, "Error"
            Sheets("SampleCalc").Select
            Range("A1").Select
            Exit Sub
        End If
    
    
   
    'To generate and sort random numbers
    
    
    Sheets("SamplePrint").Select
    'ActiveSheet.Calculate
    Counter = 1
    Randomize
    For Counter = 1 To Number
        rngCell1.Value = Int((Upper - Lower + 1) * Rnd + Lower)
        Set rngCell1 = rngCell1.Offset(rowoffset:=1)
    Next
        
    Columns("AA:AA").Select
    Selection.Sort Key1:=Range("AA1"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        
    ' To number the random numbers
        
    Range("Z1").Value = 1
        'Set numCell = Range("AA2")
        CntCell = 1
        Do Until numCell.Value = ""
            CntCell = CntCell + 1
            numCell.Offset(columnoffset:=-1).Value = CntCell
            Set numCell = numCell.Offset(rowoffset:=1)
        Loop
    
    
    'To generate extra random numbers
    
    Counter = 1
    For Counter = 1 To Extra
        rngCell2.Value = Int((Upper - Lower + 1) * Rnd + Lower)
        Set rngCell2 = rngCell2.Offset(rowoffset:=1)
    Next
    
    ' To number the extra numbers
    
    If Extra > 0 Then
        Range("AB1").Value = "Extra"
        'Set extCell = Range("AC2")
        Do Until extCell.Value = ""
            extCell.Offset(columnoffset:=-1).Value = "Extra"
            Set extCell = extCell.Offset(rowoffset:=1)
        Loop
    End If
    
    
    ' To copy random numbers to print worksheet
         
            
        'Copies first 50 numbers
        
    Range("Z1:AA50").Select
    Selection.Copy
    Range("A2").Select
    ActiveSheet.Paste
    Range("A1").Value = "Number"
    Range("B1").Value = "Sample Item"
    Range("C1").Value = "Deviation?"
    Range("D1").Value = "Note Number"
    
    
        
    
        'Copies next 50 numbers if needed
            
    If Number > 50 Then
        Range("Z51:AA100").Select
        Selection.Copy
        Range("F2").Select
        ActiveSheet.Paste
        Range("F1").Value = "Number"
        Range("G1").Value = "Sample Item"
        Range("H1").Value = "Deviation?"
        Range("I1").Value = "Note Number"
    End If
    
        'Copies next 50 numbers if needed
        
    If Number > 100 Then
        Range("Z101:AA150").Select
        Selection.Copy
        Range("A53").Select
        ActiveSheet.Paste
        Range("A52").Value = "Number"
        Range("B52").Value = "Sample Item"
        Range("C52").Value = "Deviation?"
        Range("D52").Value = "Note Number"
    
  End If
    
        'Copies next 50 numbers if needed
        
    If Number > 150 Then
        Range("Z151:AA200").Select
        Selection.Copy
        Range("F53").Select
        ActiveSheet.Paste
        Range("F52").Value = "Number"
        Range("G52").Value = "Sample Item"
        Range("H52").Value = "Deviation?"
        Range("I52").Value = "Note Number"
    
    End If
    
    
        'Copies next 50 numbers if needed
        
    If Number > 200 Then
        Range("Z201:AA250").Select
        Selection.Copy
        Range("A104").Select
        ActiveSheet.Paste
        Range("A103").Value = "Number"
        Range("B103").Value = "Sample Item"
        Range("C103").Value = "Deviation?"
        Range("D103").Value = "Note Number"
    
  End If
    
        
        'Copies last 50 numbers if needed
        
    If Number > 250 Then
        Range("Z251:AA300").Select
        Selection.Copy
        Range("F104").Select
        ActiveSheet.Paste
        Range("F103").Value = "Number"
        Range("G103").Value = "Sample Item"
        Range("H103").Value = "Deviation?"
        Range("I103").Value = "Note Number"
    
   End If


    
    
 'To copy extra numbers to print worksheet
 
 
 If Extra > 0 Then
 
    Range("AB1:AC30").Select
    Selection.Copy
    
    If Number > 250 Then
        Range("A155").Select
        ActiveSheet.Paste
        Range("A154").Value = "Number"
        Range("B154").Value = "Sample Item"
        Range("C154").Value = "Deviation?"
        Range("D154").Value = "Note Number"
    ElseIf Number > 200 Then
        Range("F104").Select
        ActiveSheet.Paste
        Range("F103").Value = "Number"
        Range("G103").Value = "Sample Item"
        Range("H103").Value = "Deviation?"
        Range("I103").Value = "Note Number"
    ElseIf Number > 150 Then
        Range("A104").Select
        ActiveSheet.Paste
        Range("A103").Value = "Number"
        Range("B103").Value = "Sample Item"
        Range("C103").Value = "Deviation?"
        Range("D103").Value = "Note Number"
    ElseIf Number > 100 Then
        Range("F53").Select
        ActiveSheet.Paste
        Range("F52").Value = "Number"
        Range("G52").Value = "Sample Item"
        Range("H52").Value = "Deviation?"
        Range("I52").Value = "Note Number"
    ElseIf Number > 50 Then
        Range("A53").Select
        ActiveSheet.Paste
        Range("A52").Value = "Number"
        Range("B52").Value = "Sample Item"
        Range("C52").Value = "Deviation?"
        Range("D52").Value = "Note Number"
    Else
        Range("F2").Select
        ActiveSheet.Paste
        Range("F1").Value = "Number"
        Range("G1").Value = "Sample Item"
        Range("H1").Value = "Deviation?"
        Range("I1").Value = "Note Number"
    End If
  
  End If
  
    
    
 ' To clear and set print sheet to focal point
 
    MsgBox "Processing is now complete.  You may print the worksheet by using normal Excel print commands", vbOKOnly, "Finished!"
    Columns("Y:AC").Select
    Selection.ClearContents
    Range("A1").Select
    

End Sub
