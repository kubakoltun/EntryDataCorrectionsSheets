Sub obrobienie_pol_dla_aktualizacji()
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
    
    
    'szukam starych branz
    Dim lR As Long
    lR = Cells(Rows.Count, "E").End(xlUp).Row
    For Each i In Range("E2:E" & lR).Cells
       If i.Value = "Architekt" Or i.Value = "Architektura i Inżynieria" Or i.Value = "Architektura" Or i.Value = "Architekt krajobrazu" Or i.Value = "Architektura i Inżynieria (inne)" Or i.Value = "Architektura Vaastu" Or i.Value = "Architektura wnętrz" Or i.Value = "Feng Shui" Or i.Value = "Inspektor" Or i.Value = "Inżynier budowlany" Or i.Value = "Pielęgnacja trawników" Or i.Value = "Serwis terenów zielonych" Or i.Value = "Usługi architektoniczne" Or i.Value = "Usługi drzewne" Then
            i.Value = "Architektura i Inżynieria"
            
       ElseIf i.Value = "Auta / Auto detailing" Or i.Value = "Auta i Motocykle" Or i.Value = "Auta / Części i akcesoria" Or i.Value = "Auta / Naprawy" Or i.Value = "Auta / Sprzedaż" Or i.Value = "Auta / Warsztat blacharski" Or i.Value = "Auta / Wypożyczalnia / Leasing" Or i.Value = "Auta i Motocykle (inne)" Or i.Value = "Dealer pojazdów komercyjnych (do przewozu towarów lub osób)" Or i.Value = "Instruktor jazdy" Or i.Value = "Sprzedaż / wymiana opon" Or i.Value = "Sprzedaż / wymiana opon" Or i.Value = "Stacja paliw" Or i.Value = "Szyby samochodowe" Then
            i.Value = "Auta i Motocykle"
  
       ElseIf i.Value = "Bezpieczeństwo (inne)" Or i.Value = "Bezpieczeństwo" Or i.Value = "Ochrona" Or i.Value = "Bezpieczeństwo pracy" Or i.Value = "Kamery / Telewizja przemysłowa" Or i.Value = "Ochrona przeciwpożarowa" Or i.Value = "Pracownicy ochrony" Or i.Value = "Systemy zabezpieczeń" Or i.Value = "Ślusarz" Or i.Value = "Usługi detektywistyczne / Detektyw" Then
            i.Value = "Bezpieczeństwo"
        
       ElseIf i.Value = "Agregaty prądotwórcze" Or i.Value = "Balkon/Weranda" Or i.Value = "Baseny, spa i sauny" Or i.Value = "Blaty / lady" Or i.Value = "Bramy garażowe" Or i.Value = "Budowa kominków i piecyków" Or i.Value = "Budowa kuchni" Or i.Value = "Budownictwo" Or i.Value = "Budownictwo (inne)" Or i.Value = "Budownictwo domów jednorodzinnych" Or i.Value = "Budowniczy / Generalny wykonawca" Or i.Value = "Cement / Beton" Or i.Value = "Cieśla" Or i.Value = "Dekoracje okienne" Or i.Value = "Drenaż" Or i.Value = "Elektryk" Or i.Value = "Energia odnawialna" Or i.Value = "Glazurnik" Or i.Value = "Hydroizolacja - uszczelnianie" Or i.Value = "Instalacje elektryczne" Or i.Value = "Instalacje wodno - kanalizacyjna" Or i.Value = "Inżynier ciepłownictwa" Or i.Value = "Komercyjne projektowanie wnętrz" Or i.Value = "Kowal" Or i.Value = "Malarz" Or i.Value = "Murarz / kamieniarz" Or i.Value = "Mycie ciśnieniowe" Or i.Value = "Obróbka metali" Or i.Value = "Odbudowa" Then
            i.Value = "Budownictwo"
            
       ElseIf i.Value = "Ogrodzenia" Or i.Value = "Budownictwo" Or i.Value = "Usługi środowiskowe" Or i.Value = "Okiennice i markizy" Or i.Value = "Okiennice i markizy" Or i.Value = "Okna i drzwi" Or i.Value = "Płyty gipsowo-kartonowe" Or i.Value = "Podłogi" Or i.Value = "Pokrycia dachowe i rynny" Or i.Value = "Powłoki ochronne / uszczelnienia" Or i.Value = "Projektowanie wnętrz" Or i.Value = "Renowacje / przebudowy" Or i.Value = "Roboty ziemne" Or i.Value = "Rozbiórki budowlane" Or i.Value = "Stolarz meblowy" Or i.Value = "Systemy ogrzewania, wentylacji i klimatyzacji" Or i.Value = "Systemy septyczne" Or i.Value = "Szkło" Or i.Value = "Tynkarz" Or i.Value = "Usługi energetyczne" Or i.Value = "Windy" Or i.Value = "Zarządzanie projektami budowlanymi" Or i.Value = "Złota rączka" Or i.Value = "Zwalczanie szkodników" Or i.Value = "Remontowo - budowlana" Then
            i.Value = "Budownictwo"
   
       ElseIf i.Value = "Broker biznesowy" Or i.Value = "Doradztwo biznesowe" Or i.Value = "Doradca biznesowy" Or i.Value = "Doradztwo biznesowe (inne)" Or i.Value = "Doradztwo energetyczne" Or i.Value = "Konsultant biznesowy" Or i.Value = "Konsultant biznesowy ds. małych przedsiębiorstw" Or i.Value = "Konsultant biznesowy ds. organizacji i procesów" Or i.Value = "Konsultant biznesowy ds. restrukturyzacji" Or i.Value = "Konsultant biznesowy ds. zarządzania" Or i.Value = "Konsultant biznesowy ds. zarządzania jakością" Or i.Value = "Profesjonalny Organizator" Or i.Value = "Specjalista ds. różnorodności, równości i integracji" Then
            i.Value = "Doradztwo biznesowe"
            
       ElseIf i.Value = "Broker ubezpieczeniowy" Or i.Value = "Doradca finansowy" Or i.Value = "Emerytury / Renty" Or i.Value = "Finanse i Ubezpieczenia (inne)" Or i.Value = "Finansowanie aktywów" Or i.Value = "Finansowanie biznesu" Or i.Value = "Fundusze inwestycyjne" Or i.Value = "Inwestycje finansowe" Or i.Value = "Karty kredytowe / usługi handlowe" Or i.Value = "Kolekcjoner" Or i.Value = "Komercyjne usługi bankowe" Or i.Value = "Kredyty hipoteczne" Or i.Value = "Kredyty indywidualne" Or i.Value = "Kredyty mieszkaniowe" Or i.Value = "Kredyty na budowę nieruchomości" Or i.Value = "Makler giełdowy" Or i.Value = "Naprawa historii kredytowej" Or i.Value = "Pełnomocnik finansowy" Or i.Value = "Pożyczki komercyjne" Or i.Value = "Rzeczoznawca ubezpieczeniowy" Or i.Value = "Sekretarz spółki" Or i.Value = "Syndyk masy upadłościowej" Or i.Value = "Ubezpieczenia majątkowe i wypadkowe" Or i.Value = "Ubezpieczenia na życie" Or i.Value = "Ubezpieczenia zdrowotne" Or i.Value = "Ubezpieczenie dodatkowe" Then
            i.Value = "Finanse i Ubezpieczenia"
  
       ElseIf i.Value = "Ubezpieczenie komercyjne" Or i.Value = "Finanse i Ubezpieczenia" Or i.Value = "Ubezpieczenia" Or i.Value = "Finanse" Or i.Value = "Usługi bankowe" Or i.Value = "Waluty zagraniczne" Or i.Value = "Zakupy grupowe" Or i.Value = "Zarządzanie aktywami" Then
            i.Value = "Finanse i Ubezpieczenia"
            
       ElseIf i.Value = "Catering" Or i.Value = "Gastronomia" Or i.Value = "Cukiernik" Or i.Value = "Gastronomia" Or i.Value = "Gastronomia (inne)" Or i.Value = "Piekarz" Or i.Value = "Restaurator" Or i.Value = "Serwis napojów" Or i.Value = "Sprzedawca win / Wino" Then
            i.Value = "Gastronomia"
            
       ElseIf i.Value = "Agencja zatrudnienia" Or i.Value = "HR i Rekrutacja" Or i.Value = "HR i Rekrutacja (inne)" Or i.Value = "Konsultant ds. prawa pracy" Or i.Value = "Rekruter" Or i.Value = "Usługi administracyjne" Or i.Value = "Wirtualny asystent" Or i.Value = "Zasoby ludzkie" Or i.Value = "HR" Then
            i.Value = "HR i Rekrutacja"
            
       ElseIf i.Value = "Bezpieczeństwo danych" Or i.Value = "Komputery i Programowanie" Or i.Value = "Deweloper aplikacji" Or i.Value = "IT i sieci" Or i.Value = "Komputery i Programowanie (inne)" Or i.Value = "Konsultant IT" Or i.Value = "Kursy komputerowe" Or i.Value = "Oprogramowanie ERP" Or i.Value = "Oprogramowanie komputerowe" Or i.Value = "Programista" Or i.Value = "Sprzedawca komputerów" Or i.Value = "Usługi w chmurze" Or i.Value = "Technologia informatyczna" Or i.Value = "Komputery" Or i.Value = "Programowanie" Then
            i.Value = "Komputery i Programowanie"
            
       ElseIf i.Value = "Agencja reklamowa" Or i.Value = "Marketing i Reklama" Or i.Value = "Branding" Or i.Value = "Copywriter / Pisarz" Or i.Value = "Drukarnia" Or i.Value = "Drukarnia cyfrowa" Or i.Value = "Drukarnia offsetowa" Or i.Value = "Drukarnia wielkoformatowa" Or i.Value = "Fotograf" Or i.Value = "Fotograf komercyjny" Or i.Value = "Grafik" Or i.Value = "Inteligentny dom" Or i.Value = "Kamerzysta / producent filmowy" Or i.Value = "Konsultant marketingowy" Or i.Value = "Marketing cyfrowy" Or i.Value = "Marketing relacji" Or i.Value = "Pozycjonowanie" Or i.Value = "Produkty promocyjne / reklamowe" Or i.Value = "Projektowanie stron" Or i.Value = "Public Relation" Or i.Value = "Reklama drukowana" Or i.Value = "Reklama i Marketing (inne)" Or i.Value = "Reklama radiowa" Or i.Value = "Reklama telewizyjna" Or i.Value = "Social Media" Or i.Value = "Szyldy" Or i.Value = "Tworzenie stron internetowych" Or i.Value = "Usługi medialne" Or i.Value = "Wydawca" Or i.Value = "Zdobienia" Then
            i.Value = "Marketing i Reklama"
            
       ElseIf i.Value = "Naprawa komputerów" Or i.Value = "Naprawy i renowacje" Or i.Value = "Naprawa maszyn i urządzeń" Or i.Value = "Naprawa mebli / tapicerka" Or i.Value = "Naprawa urządzeń" Or i.Value = "Naprawy i renowacje (inne)" Then
            i.Value = "Naprawy i renowacje"
            
       ElseIf i.Value = "Administrator apartamentów" Or i.Value = "Agencja nieruchomości" Or i.Value = "Agent ds. zakupów" Or i.Value = "Czyszczenie dywanów, tapicerki" Or i.Value = "Czyszczenie okien" Or i.Value = "Inspektor budowlany" Or i.Value = "Inspektor Budynków Mieszkalnych" Or i.Value = "Inwestycje w nieruchomości" Or i.Value = "Komercyjny serwis sprzątający" Or i.Value = "Konsultant ds. planowania nieruchomości" Or i.Value = "Nieruchomości (inne)" Or i.Value = "Nieruchomości komercyjne" Or i.Value = "Przygotowanie nieruchomości na sprzedaż" Or i.Value = "Rozwój nieruchomości" Or i.Value = "Serwis sprzątający" Or i.Value = "Sprzedawca energii elektrycznej i gazu" Or i.Value = "Usługi tytułowe" Or i.Value = "Utrzymanie / opieka nad nieruchomościami" Or i.Value = "Wycena nieruchomości" Or i.Value = "Wynajem nieruchomości" Or i.Value = "Zarządzanie i gospodarowanie odpadami" Or i.Value = "Zarządzanie nieruchomościami" Or i.Value = "Zwolnienie z podatku od nieruchomości" Then
            i.Value = "Nieruchomości"
            
       ElseIf i.Value = "Nieruchomości" Then
            i.Value = "Nieruchomości"
            
       ElseIf i.Value = "Givers Gain®" Or i.Value = "Obsługa przedsiębiorstw" Or i.Value = "Izba / Stowarzyszenie" Or i.Value = "Obsługa przedsiębiorstw (inne)" Or i.Value = "Organizacje non-profit / fundraising" Then
            i.Value = "Obsługa przedsiębiorstw"

       ElseIf i.Value = "Agent biura podróży" Or i.Value = "Podróże" Or i.Value = "Podróże (inne)" Or i.Value = "Sprzedaż biletów" Or i.Value = "Wycieczki / przewodnik" Then
            i.Value = "Podróże"
            
       ElseIf i.Value = "Audytor" Or i.Value = "Certyfikowany księgowy" Or i.Value = "Czynności notarialne" Or i.Value = "Doradca podatkowy" Or i.Value = "Funkcjonariusz organów ścigania" Or i.Value = "Księgowość" Or i.Value = "Mediator" Or i.Value = "Notariusz" Or i.Value = "Plan obsługi prawnej" Or i.Value = "Planowanie przestrzenne" Or i.Value = "Prawa osób starszych" Or i.Value = "Prawnik" Or i.Value = "Prawo biznesowe" Or i.Value = "Prawo cywilne" Or i.Value = "Prawo i Rachunkowość (inne)" Or i.Value = "Prawo imigracyjne" Or i.Value = "Prawo karne" Or i.Value = "Prawo nieruchomości" Or i.Value = "Prawo odszkodowawcze" Or i.Value = "Prawo podatkowe" Or i.Value = "Prawo rodzinne" Or i.Value = "Prawo upadłościowe" Or i.Value = "Prawo własności intelektualnej" Or i.Value = "Prawo zdrowotne" Or i.Value = "Rejestrowanie przedsiębiorstw za granicą" Or i.Value = "Testamenty / Umowy powiernicze" Or i.Value = "Usługi księgowe" Or i.Value = "Usługi rządowe i samorządowe" Then
            i.Value = "Prawo i Rachunkowość"
   
       ElseIf i.Value = "Zakładanie spółek" Or i.Value = "Prawo i Rachunkowość" Or i.Value = "Prawo" Or i.Value = "Rachunkowość" Or i.Value = "Zatrudnienie / prawo pracy" Then
            i.Value = "Prawo i Rachunkowość"
            
       ElseIf i.Value = "Drewno i korek (z wyjątkiem mebli)" Or i.Value = "Produkcja" Or i.Value = "Przemysł" Or i.Value = "Komputery, Elektronika i Optyka" Or i.Value = "Nośniki informacji" Or i.Value = "Odzież" Or i.Value = "Pakowanie" Or i.Value = "Papier i produkty papierowe" Or i.Value = "Pojazdy silnikowe" Or i.Value = "Producent farb" Or i.Value = "Producent leków" Or i.Value = "Producent maszyn i urządzeń" Or i.Value = "Producent mebli" Or i.Value = "Producent napojów" Or i.Value = "Producent podłóg" Or i.Value = "Producent produktów gumowych i plastikowych" Or i.Value = "Producent sprzętu elektrycznego" Or i.Value = "Producent stali" Or i.Value = "Produkcja (inne)" Or i.Value = "Produkcja metali" Or i.Value = "Produkcja oświetlenia" Or i.Value = "Produkty chemiczne" Or i.Value = "Produkty naftowe" Or i.Value = "Produkty skórzane" Or i.Value = "Produkty żywieniowe" Or i.Value = "Sprzęt transportowy" Or i.Value = "Surowce niemetaliczne" Or i.Value = "Tekstylia" Or i.Value = "Wyroby tytoniowe" Then
            i.Value = "Produkcja"
  
       ElseIf i.Value = "Agronom" Or i.Value = "Rolnictwo" Or i.Value = "Rolnictwo (inne)" Then
            i.Value = "Rolnictwo"
            
       ElseIf i.Value = "Sport i rekreacja (inne)" Or i.Value = "Sport i rekreacja" Or i.Value = "Sztuki walki" Or i.Value = "Trener jogi / Pilates / Qi-gong" Then
            i.Value = "Sport i rekreacja"
            
       ElseIf i.Value = "Akcesoria komputerowe" Or i.Value = "Jubilerstwo" Or i.Value = "Akcesoria łazienkowe" Or i.Value = "Antykwariusz" Or i.Value = "Artykuły biurowe" Or i.Value = "Diamenty" Or i.Value = "Drwal" Or i.Value = "Florysta" Or i.Value = "Jubiler" Or i.Value = "Kina domowe" Or i.Value = "Materace" Or i.Value = "Meble biurowe" Or i.Value = "Odzież na zamówienie / krawiec" Or i.Value = "Paliwa" Or i.Value = "Płytki ceramiczne, kafelki" Or i.Value = "Pokrycie ścian" Or i.Value = "Prezenty / Upominki" Or i.Value = "Produkty czyszczące" Or i.Value = "Produkty środowiskowe" Or i.Value = "Projektant biżuterii" Or i.Value = "Sprzedawca dzieł sztuki/Właściciel galerii" Or i.Value = "Sprzedawca elektroniki" Or i.Value = "Sprzedawca odzieży i akcesoriów" Or i.Value = "Sprzedawca sportowy" Or i.Value = "Sprzedaż (inne)" Or i.Value = "Sprzedaż drzwi" Or i.Value = "Sprzedaż farb" Or i.Value = "Sprzedaż materiałów budowlanych" Or i.Value = "Sprzedaż mebli" Then
            i.Value = "Sprzedaż"
            
       ElseIf i.Value = "Sprzedaż oświetlenia" Or i.Value = "Sprzedaż" Or i.Value = "Sprzedaż podłóg" Or i.Value = "Sprzedaż sprzętu elektrycznego" Or i.Value = "Sprzedaż umundurowania" Or i.Value = "Sprzedaż urządzeń" Or i.Value = "Sprzedaż wyrobów budowlanych" Or i.Value = "Sprzęt biurowy / maszyny" Or i.Value = "Systemy wodne" Or i.Value = "UPS / Inverter" Or i.Value = "Wyposażenie domu" Then
            i.Value = "Sprzedaż"
            
       ElseIf i.Value = "Centrum szkoleniowe" Or i.Value = "Szkolenia i coaching" Or i.Value = "Szkolenia biznesowe / trener" Or i.Value = "Szkolenia i coaching (inne)" Or i.Value = "Szkolenia sprzedażowe / Coaching" Or i.Value = "Terapeuta życiowy" Or i.Value = "Trener komunikacji" Or i.Value = "Trener przywództwa" Or i.Value = "Trener zarządzania" Or i.Value = "Usługi edukacyjne / Korepetycje" Then
            i.Value = "Szkolenia i coaching"
            
       ElseIf i.Value = "Artysta" Or i.Value = "DJ" Or i.Value = "Sztuka i Rozrywka" Or i.Value = "Komik" Or i.Value = "Muzyk" Or i.Value = "Sztuka i Rozrywka (inne)" Then
            i.Value = "Sztuka i Rozrywka"
            
       ElseIf i.Value = "Produkty / usługi telekomunikacyjne" Or i.Value = "Telekomunikacja" Or i.Value = "Telekomunikacja (inne)" Or i.Value = "Telekomunikacja mobilna" Then
            i.Value = "Telekomunikacja"
            
       ElseIf i.Value = "Kurier" Or i.Value = "Transport i wysyłka" Or i.Value = "Transport" Or i.Value = "Wysyłka" Or i.Value = "Transport / limuzyny" Or i.Value = "Transport i wysyłka (inne)" Or i.Value = "Transport komercyjny" Or i.Value = "Usługi frachtowe" Or i.Value = "Usługi pocztowe" Or i.Value = "Usługi przeprowadzkowe" Then
            i.Value = "Transport i wysyłka"
            
       ElseIf i.Value = "Call Center" Or i.Value = "Usługi eventowe i biznesowe" Or i.Value = "Event Manager / Marketingowiec" Or i.Value = "Eventy" Or i.Value = "Hotel" Or i.Value = "Imprezy firmowe" Or i.Value = "Planowanie imprez, wydarzeń" Or i.Value = "Technik - audio, wideo" Or i.Value = "Tłumacz / Usługi językowe" Or i.Value = "Usługi biurowe" Or i.Value = "Usługi eventowe i biznesowe (inne)" Or i.Value = "Wynajem pomieszczeń eventowych" Or i.Value = "Wypożyczalnia sprzętu na eventy" Then
            i.Value = "Usługi eventowe i biznesowe"
            
       ElseIf i.Value = "Astrolog" Or i.Value = "Usługi osobiste" Or i.Value = "Dostawca usług" Or i.Value = "Fryzjer" Or i.Value = "Konsultant ds. kolorów i stylu" Or i.Value = "Kosmetyki / Pielęgnacja skóry" Or i.Value = "Organizator ślubów" Or i.Value = "Pralnia chemiczna / Pralnia" Or i.Value = "Salon / Spa" Or i.Value = "Usługi osobiste (inne)" Or i.Value = "Usługi pogrzebowe" Then
            i.Value = "Usługi osobiste"
            
       ElseIf i.Value = "Akupunktura" Or i.Value = "Zdrowie i Uroda" Or i.Value = "Dentysta" Or i.Value = "Dietetyk" Or i.Value = "Doradca zdrowego stylu życia" Or i.Value = "Farmaceuta" Or i.Value = "Fizjoterapeuta" Or i.Value = "Hipnoterapeuta" Or i.Value = "Hospicjum" Or i.Value = "Kręgarz" Or i.Value = "Lekarz" Or i.Value = "Masażysta" Or i.Value = "Medycyna alternatywna" Or i.Value = "Naturopaci" Or i.Value = "Okulista / Optyk" Or i.Value = "Olejki eteryczne" Or i.Value = "Opieka domowa" Or i.Value = "Ortodonta" Or i.Value = "Osteopata" Or i.Value = "Produkty zdrowie i uroda" Or i.Value = "Psycholog / Terapeuta" Or i.Value = "Siłownia / Klub" Or i.Value = "Słuch / audiologia" Or i.Value = "Sofrolog" Or i.Value = "Suplementy diety" Or i.Value = "Trener osobisty - fitness" Or i.Value = "Usługi medyczne" Or i.Value = "Usługi zdrowie i uroda" Or i.Value = "Zdrowie i Uroda (inne)" Then
            i.Value = "Zdrowie i Uroda"
            
       ElseIf i.Value = "Akwarium/Ryby" Or i.Value = "Zwierzęta" Or i.Value = "Lekarz weterynarii" Or i.Value = "Opieka nad zwierzętami domowymi" Or i.Value = "Pielęgnacja zwierząt" Or i.Value = "Trener psów" Or i.Value = "Zwierzęta (inne)" Or i.Value = "Żywność dla zwierząt" Then
            i.Value = "Zwierzęta"
            
       Else
            i.Value = "Nie znaleziono na podstawie specjalizacji"
       End If
    Next
    
    
    
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
