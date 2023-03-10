Sub dynamic_dates_shift()
'
' spojnosc_pol Makro
' 
'
    Windows("hs_kopiaDanych_czlonkowie.xlsx").Activate
    
    Columns("W:W").EntireColumn.AutoFit
    Columns("X:X").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("X:X").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("W2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'moze text to columns zastapic usunieciem bledu
    Selection.TextToColumns Destination:=Range("W2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("X1").Select
    ActiveCell.FormulaR1C1 = ""
    Windows("bc_kopiaDanych_czlonkowie.xlsx").Activate
    
    'zaznacza caly arkusz
    '
    Set lastRbcKopia = Cells(Rows.Count, "AE").End(xlUp)
    Set R = Range("AE1", lastRbcKopia)
    'Selection.AutoFill Destination:=Range("N2", Cells(lastR.Row, "N"))
    Range("AE1").Select
    Range("AE1", Cells(lastRbcKopia.Row, "AE")).Select
    Selection.Copy
    Windows("hs_kopiaDanych_czlonkowie.xlsx").Activate
    
    Range("X1").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    
    'blad wprowadzanie nie mam pewnosci
    'moze text to columns zastapic usunieciem bledu
    'byl teksto jako kolumny nie widze dalszego sensu skoro dobry format
   
    
   
    Range("X1").Select
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "wp"
    
    Dim LastX As Long
    LastX = Range("X" & Rows.Count).End(xlUp).Row
    
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],R2C24:R" & LastX & "C24,1,FALSE)" '2337 swap for var
    Range("Z2").Select
    Range("Y2").Select
    
    
    Set lastRhsKopia = Cells(Rows.Count, "A").End(xlUp)
    Set R = Range("A1", lastRhsKopia)
    'Selection.AutoFill Destination:=Range("N2", Cells(lastR.Row, "N"))
    
    Selection.AutoFill Destination:=Range("Y2", Cells(lastRhsKopia.Row, "Y"))
    Range("Y2:Y2337").Select
    
    
    Range("Y2").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1", Cells(lastRhsKopia.Row, "AE")).AutoFilter Field:=25, Criteria1:="#N/D"
    ActiveSheet.Range("$A$1", Cells(lastRhsKopia.Row, "AE")).AutoFilter Field:=23, Criteria1:="<>"

    Columns("A:A").Select
    Selection.Copy
    'wklejenie do bylych czlonkow
    Windows("import byli Członkowie.xlsx").Activate
    
    Range("A1").Select
    ActiveSheet.Paste
    Range("F50:G50").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("F2").Select
    
    Set lastRbyli = Cells(Rows.Count, "A").End(xlUp)
    Set R = Range("A1", lastRbyli)
    'Selection.AutoFill Destination:=Range("N2", Cells(lastR.Row, "N"))
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2", Cells(lastRbyli.Row, "F")) 'zakres do konca
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2", Cells(lastRbyli.Row, "G"))
    Range("F49:G50").Select
    Windows("hs_kopiaDanych_czlonkowie.xlsx").Activate
    Range("O385").Select
    Windows("bc_kopiaDanych_czlonkowie.xlsx").Activate
    
    
    Range("AM1").Select
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1", Cells(lastRbcKopia.Row, "AJ")).AutoFilter Field:=14, Criteria1:=Array( _
        "Aktywne", "Opóźnienie", "Zbliżające się przedłużenie"), Operator:= _
        xlFilterValues
        
        
    Dim currentMonth As Integer
    Dim currentYear As Integer
    Dim filterCriteria As Variant
    Dim arrForMonths(10) As Integer
    Dim iteration As Integer
    Dim inputForMonths As Integer
    Dim properIteration As Integer
    
    inputForMonths = 1
    properIteration = 0
    
    currentMonth = Month(Date)
    currentYear = Year(Date)
    
    For iteration = 0 To 11
        If iteration <> currentMonth Then
            arrForMonths(properIteration) = inputForMonths
            MsgBox (properIteration & arrForMonths(properIteration))
            properIteration = properIteration + 1
        End If
        inputForMonths = inputForMonths + 1
    Next
    
    
   
   'varaiable with flexible previous years
   filterCriteria = Array(0, "12/1/" & currentYear + 3, 0, "12/1/" & currentYear + 2, 0, "12/1/" & currentYear + 1, _
        1, arrForMonths(0) & "/1/" & currentYear, 1, arrForMonths(1) & "/1/" & currentYear, 1, arrForMonths(2) & "/1/" & currentYear, 1, arrForMonths(3) & "/1/" & currentYear, _
        1, arrForMonths(4) & "/1/" & currentYear, 1, arrForMonths(5) & "/1/" & currentYear, 1, arrForMonths(6) & "/1/" & currentYear, 1, arrForMonths(7) & "/1/" & currentYear, _
        1, arrForMonths(8) & "/1/" & currentYear, 1, arrForMonths(9) & "/1/" & currentYear, 1, arrForMonths(10) & "/1/" & currentYear, 0, "12/1/" & currentYear - 1, _
        0, "12/1/" & currentYear - 2, 0, "12/1/" & currentYear - 3)
        
        
    'filter that moves along with time
    ActiveSheet.Range("$A$1", Cells(lastRbcKopia.Row, "AJ")).AutoFilter Field:=18, Operator:= _
        xlFilterValues, Criteria2:=filterCriteria
        
    'static dates filter
       ' ActiveSheet.Range("$A$1", Cells(lastRbcKopia.Row, "AJ")).AutoFilter Field:=18, Operator:= _
      '  xlFilterValues, Criteria2:=Array(0, "12/1/2026", 0, "12/1/2025", 0, "12/1/2024", 1, _
     '   "1/1/2023", 1, "2/1/2023", 1, "3/4/2023", 1, "5/1/2023", 1, "6/1/2023", 1, "7/1/2023", 1, _
     '   "8/1/2023", 1, "9/1/2023", 1, "10/1/2023", 1, "11/1/2023", 1, "12/1/2023", 0, "6/1/2021", _
     '   0, "5/1/2019")
        'nastepny 4/1/2023
        
    Range("A1", Cells(lastRbcKopia.Row, "AJ")).Select
    Selection.Copy
    Windows("przygotowanie_czlonkowie.xlsx").Activate
    ActiveSheet.Paste
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Windows("bc_kopiaDanych_czlonkowie.xlsx").Activate
    Range("L2331").Select
    Windows("hs_kopiaDanych_czlonkowie.xlsx").Activate
    Cells.Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Cells.Select
    Selection.Delete Shift:=xlUp
    Windows("bc_kopiaDanych_czlonkowie.xlsx").Activate
    Cells.Select
    Selection.Delete Shift:=xlUp
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("B6").Select
    Windows("hs_kopiaDanych_czlonkowie.xlsx").Activate
    Range("L7").Select
    'Windows("czlonkowie_makro.xlsm").Activate

    Windows("przygotowanie_czlonkowie.xlsx").Activate


    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    
    
    
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Range("G1").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "A"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "B"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "C"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "D"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "E"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "F"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "G"
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    Range("I2209").Select
    ActiveCell.FormulaR1C1 = "SUN"
    Range("I2207").Select
    Selection.End(xlUp).Select
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "H"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "I"
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "J"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "K"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "L"
    Columns("N:N").Select
    Selection.Delete Shift:=xlToLeft
    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Range("O1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "M"
    Range("P1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "N"
    Range("Q1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "O"
    Range("R1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "P"
    Range("S1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "R"
    Range("T1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "S"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "Tak"
    Range("T3").Select
    Range("T2").Select
    
    
    Range("O2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "T"
    Range("O2").Select
    
    Set lastRprzygotowaniehs = Cells(Rows.Count, "A").End(xlUp)
    Set R = Range("A1", lastRprzygotowaniehs)
    'Selection.AutoFill Destination:=Range("N2", Cells(lastRprzygotowaniehs.Row, "N"))
    
    Selection.AutoFill Destination:=Range("O2", Cells(lastRprzygotowaniehs.Row, "O")), Type:=xlFillDefault 'zakres do konca
    Range("O2", Cells(lastRprzygotowaniehs.Row, "O")).Select
    Range("O2214").Select
    Selection.End(xlUp).Select
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "U"
    'Windows("Zeszyt6").Activate
    Rows("1:1").Select
    Columns("P:P").EntireColumn.AutoFit
    
    'Windows("czlonkowie_makro_import.xlsm").Activate
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("Q2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Tak"
    Range("R2").Select
    Range("R2").Select
    'Windows("czlonkowie_makro_import.xlsm").Activate
    Range("P2").Select
    Selection.AutoFill Destination:=Range("P2", Cells(lastRprzygotowaniehs.Row, "P")) ' zmiana zakresu na koncowy
    Range("P2", Cells(lastRprzygotowaniehs.Row, "P")).Select
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2", Cells(lastRprzygotowaniehs.Row, "Q")) ' zmiana akresu na koncowy
    Range("Q2", Cells(lastRprzygotowaniehs.Row, "Q")).Select
    Range("T2").Select
    Selection.AutoFill Destination:=Range("T2", Cells(lastRprzygotowaniehs.Row, "T")) ' zmiana zakresu na koncowy
    Range("T2", Cells(lastRprzygotowaniehs.Row, "T")).Select
    Range("Q2").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("R2").Select
    
    'Windows("czlonkowie_przygotowanie.xlsx").Activate
    Sheets("grupa+region").Select
    Dim LastA As Long
    LastA = Range("A" & Rows.Count).End(xlUp).Row
    Sheets("Arkusz1").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-9],'grupa+region'!R2C1:R" & LastA & "C3,3,FALSE)"
        '"=VLOOKUP(RC[-9],'grupa+region'!R2C1:R114C3,3,FALSE)"
    Range("R2").Select
    Selection.AutoFill Destination:=Range("R2", Cells(lastRprzygotowaniehs.Row, "R")) 'zakres do konca
    Range("R2", Cells(lastRprzygotowaniehs.Row, "R")).Select
    Range("S2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-10],'grupa+region'!R2C1:R" & LastA & "C6,6,FALSE)"
        '"=VLOOKUP(R[-9]C[8],'grupa+region'!R2C1:R114C6,6,FALSE)"
        '"=VLOOKUP(RC[-10],'grupa+region'!R2C1:R114C5,5,FALSE)"
    Range("S2").Select
    Selection.AutoFill Destination:=Range("S2", Cells(lastRprzygotowaniehs.Row, "S")) ' zakres do konca
    Range("S2", Cells(lastRprzygotowaniehs.Row, "S")).Select
    'Windows("czlonkowie_makro_import.xlsm").Activate
    
    'Windows("czlonkowie_hs_przygotowanie.xlsx").Activate
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("K2213").Select
    Selection.End(xlUp).Select
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-1],9)"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2", Cells(lastRprzygotowaniehs.Row, "L")) ' zakres do konca
    Range("L2", Cells(lastRprzygotowaniehs.Row, "L")).Select
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=""+48""&RC[-1]"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2", Cells(lastRprzygotowaniehs.Row, "M")) ' zakres do konca
    Range("M2", Cells(lastRprzygotowaniehs.Row, "M")).Select
    Selection.Copy
    Range("K2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L1").Select
    Columns("K:K").ColumnWidth = 8.56
    Columns("L:M").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("K2").Select
    'Windows("czlonkowie_makro_import.xlsm").Activate
       
    
    'Windows("czlonkowie_hs_przygotowanie.xlsx").Activate
    Cells.Select
    Selection.Copy
    Windows("import aktualizacja bazy Członków.xlsx").Activate
    Range("A1").Select
    'Range("A1").PasteSpecial Paste:=xlPasteValues
    'ActiveSheet.Paste
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("A1").PasteSpecial Paste:=xlPasteValues
    
    Range("U9").Select
    Columns("T:T").EntireColumn.AutoFit
    Columns("G:G").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "yyyy-mm-dd;@"
    Columns("H:H").Select
    Selection.NumberFormat = "yyyy-mm-dd;@"
    Range("J11").Select
    Application.WindowState = xlNormal
    'Windows("czlonkowie_makro_import.xlsm").Activate
     Windows("przygotowanie_czlonkowie.xlsx").Activate
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("B3").Select

    
    'closing workbooks
    Workbooks("bc_kopiaDanych_czlonkowie.xlsx").Close SaveChanges:=True
    Workbooks("hs_kopiaDanych_czlonkowie.xlsx").Close SaveChanges:=True
    Workbooks("przygotowanie_czlonkowie.xlsx").Close SaveChanges:=True
    Workbooks("import byli Członkowie.xlsx").Close SaveChanges:=True
    Workbooks("import aktualizacja bazy Członków.xlsx").Close SaveChanges:=True
    Workbooks("czlonkowie_makro.xlsm").Close SaveChanges:=True
    
    
End Sub
