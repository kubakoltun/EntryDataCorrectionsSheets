Sub export_preparation_assigning_category()

    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Range("L1").Select
    
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "A"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "B"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = ""
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "C"
    Range("N1").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "D"
    Range("O1").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "E"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "F"
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "G"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "H"
 
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "J"
    Range("M2").Select
    
    
    Dim Var As Integer
    Var = 2
    Dim lRk As Long
    lRk = Cells(Rows.Count, "H").End(xlUp).Row
    For Each i In Range("H2:H" & lRk).Cells
        If InStr(i.Value, "Vale1") > 0 Then
            'i.Offset(0, 1).Value = "Certain Value1"
            Range("M" & Var).Value = "Certain Value2" 
        Else
            'i.Offset(0, 1).Value = "Vale2"
            Range("M" & Var).Value = "Certain Value2"
        End If
        Var = Var + 1
    Next
    
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("O2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-1]=""Field"",RC[-1]=""Repeat""), ""Cfr"",""UnCfr"")"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-9],'grupa+owner'!C[-16]:C[-14],3,0)"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = ""
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],'grupa+owner'!C[-15]:C[-14],2,0)"
    Range("P3").Select
    Columns("P:P").EntireColumn.AutoFit
    Range("K2:L2").Select
    
    Set lastRAuFv = Cells(Rows.Count, "E").End(xlUp)
    Set R = Range("E1", lastRAuFv)
    
    Selection.AutoFill Destination:=Range("K2", Cells(lastRAuFv.Row, "L")), Type:=xlFillDefault
    Range("K2:L2251").Select
    Range("O2:Q2").Select
    Selection.AutoFill Destination:=Range("O2", Cells(lastRAuFv.Row, "Q"))
    Range("O2:Q2251").Select
    Range("O2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("N2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("N2").Select
    Range("O2").Select
    Application.CutCopyMode = False
    Range("O2").Select
    Range("O1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("O:O").Select
    Columns("F:F").Select
    Columns("F:F").Select
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G2").Select
    ActiveCell.FormulaR1C1 = ""
    Columns("F:F").Select
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-1],9)"
    Range("H2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=""+48""&RC[-1]"
    Range("H2").Select
    
    Set lastRAuF = Cells(Rows.Count, "E").End(xlUp)
    Set R = Range("E1", lastRAuF)
    
    Selection.AutoFill Destination:=Range("H2", Cells(lastRAuF.Row, "H"))
    Range("H2:H2251").Select
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2", Cells(lastRAuF.Row, "G"))
    Range("G2:G2251").Select
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("G:G").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "M"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "N"
    Range("K2").Select
    
    Columns("E:E").Select
    ActiveSheet.Range("$A:$P").RemoveDuplicates Columns:=5, Header:=xlYes
    
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "O"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "P"
    Range("D1").Select
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "R"
    Range("F2").Select
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "S"
    Range("J2").Select
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "T"
    
    'Range("Q2").Select
    'ActiveCell.FormulaR1C1 = _
      '  "=VLOOKUP(RC[-15],'grupa+owner'!C[-15]:C[-14],2,0)"
    
    Columns("I:I").Select
    Selection.NumberFormat = "yyyy-mm-dd;@"
    
    'dodanie pola branzy 
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "U"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "V"
    Range("E2").Select
    
    'assigning category
    
    Dim lR As Long
    lR = Cells(Rows.Count, "D").End(xlUp).Row
    For Each i In Range("D2:D" & lR).Cells
       If i.Value = "Architekt" Or i.Value = "Architekt krajobrazu" Or i.Value = "Architektura i Inżynieria (inne)" Then
            i.Offset(0, 1).Value = "Architektura i Inżynieria"
            
       ElseIf i.Value = "Auta / Auto detailing" Or i.Value = "Auta / Części i akcesoria" Then
            i.Offset(0, 1).Value = "Auta i Motocykle"
  
       ElseIf i.Value = "Bezpieczeństwo (inne)" Or i.Value = "Bezpieczeństwo pracy" Then
            i.Offset(0, 1).Value = "Bezpieczeństwo"
            
       ElseIf i.Value = "Professional" Then
            i.Offset(0, 1).Value = "Professional"
            
       ElseIf i.Value = "Agregaty prądotwórcze" Or i.Value = "Balkon/Weranda" Or i.Value = "Baseny, spa i sauny" Or i.Value = "Blaty / lady" Then
            i.Offset(0, 1).Value = "Budownictwo"
   
       ElseIf i.Value = "Broker biznesowy" Or i.Value = "Doradca biznesowy" Or i.Value = "Doradztwo biznesowe (inne)" Then
            i.Offset(0, 1).Value = "Doradztwo biznesowe"
            
       ElseIf i.Value = "Broker ubezpieczeniowy" Or i.Value = "Doradca finansowy" Then
            i.Offset(0, 1).Value = "Finanse i Ubezpieczenia"
            
       ElseIf i.Value = "Catering" Or i.Value = "Cukiernik" Or i.Value = "Gastronomia" Then
            i.Offset(0, 1).Value = "Gastronomia"
            
       ElseIf i.Value = "Agencja zatrudnienia" Or i.Value = "HR i Rekrutacja (inne)" Or i.Value = "Konsultant ds. prawa pracy" Then
            i.Offset(0, 1).Value = "HR i Rekrutacja"
            
       ElseIf i.Value = "Agencja reklamowa" Or i.Value = "Branding" Or i.Value = "Copywriter / Pisarz" Then
            i.Offset(0, 1).Value = "Marketing i Reklama"
            
       ElseIf i.Value = "Naprawa komputerów" Or i.Value = "Naprawa maszyn i urządzeń" Then
            i.Offset(0, 1).Value = "Naprawy i renowacje"
            
       ElseIf i.Value = "Administrator apartamentów" Or i.Value = "Agencja nieruchomości" Then
            i.Offset(0, 1).Value = "Nieruchomości"

       ElseIf i.Value = "Agent biura podróży" Or i.Value = "Podróże (inne)" Or i.Value = "Sprzedaż biletów" Then
            i.Offset(0, 1).Value = "Podróże"
   
       ElseIf i.Value = "Zakładanie spółek" Or i.Value = "Zatrudnienie / prawo pracy" Then
            i.Offset(0, 1).Value = "Prawo i Rachunkowość"
            
       ElseIf i.Value = "Drewno i korek (z wyjątkiem mebli)" Or i.Value = "Komputery, Elektronika i Optyka" Then
            i.Offset(0, 1).Value = "Produkcja"
  
       ElseIf i.Value = "Agronom" Or i.Value = "Rolnictwo (inne)" Then
            i.Offset(0, 1).Value = "Rolnictwo"
            
       ElseIf i.Value = "Sport i rekreacja (inne)" Or i.Value = "Sztuki walki" Then
            i.Offset(0, 1).Value = "Sport i rekreacja"
            
       ElseIf i.Value = "Akcesoria komputerowe" Or i.Value = "Akcesoria łazienkowe" Then
            i.Offset(0, 1).Value = "Sprzedaż"
            
       ElseIf i.Value = "Sprzedaż oświetlenia" Or i.Value = "Sprzedaż podłóg" Then
            i.Offset(0, 1).Value = "Sprzedaż"
            
       ElseIf i.Value = "Centrum szkoleniowe" Or i.Value = "Szkolenia biznesowe / trener" Or i.Value = "Szkolenia i coaching (inne)" Then
            i.Offset(0, 1).Value = "Szkolenia i coaching"
            
       ElseIf i.Value = "Artysta" Or i.Value = "Muzyk" Or i.Value = "Sztuka i Rozrywka (inne)" Then
            i.Offset(0, 1).Value = "Sztuka i Rozrywka"
            
       ElseIf i.Value = "Produkty / usługi telekomunikacyjne" Or i.Value = "Telekomunikacja (inne)" Then
            i.Offset(0, 1).Value = "Telekomunikacja"
            
       ElseIf i.Value = "Kurier" Or i.Value = "Transport / limuzyny" Then
            i.Offset(0, 1).Value = "Transport i wysyłka"
            
       ElseIf i.Value = "Call Center" Or i.Value = "Event Manager / Marketingowiec" Or i.Value = "Eventy" Then
            i.Offset(0, 1).Value = "Usługi eventowe i biznesowe"
            
       ElseIf i.Value = "Astrolog" Or i.Value = "Dostawca usług" Then
            i.Offset(0, 1).Value = "Usługi osobiste"
            
       ElseIf i.Value = "Akupunktura" Or i.Value = "Dentysta" Or i.Value = "Dietetyk" Then
            i.Offset(0, 1).Value = "Zdrowie i Uroda"
            
       ElseIf i.Value = "Akwarium/Ryby" Or i.Value = "Lekarz weterynarii" Then
            i.Offset(0, 1).Value = "Zwierzęta"
            
       Else
            i.Offset(0, 1).Value = "Nie znaleziono na podstawie specjalizacji"
       End If
    Next
    
    
    
    Columns("J:J").Select
    Selection.NumberFormat = "yyyy-mm-dd;@"
    Range("H3").Select
    
          
    Columns("F:F").Select
    Selection.Replace What:=".MERGE", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
        
    'email verification and correction
    Dim shortReg As String
    Dim email As String
    Dim lastRowEmail As Long
    lastRowEmail = Cells(Rows.Count, "F").End(xlUp).Row
    For Each i In Range("F2:F" & lastRowEmail).Cells
        email = Trim(i.Value)
        lastDotPos = InStrRev(email, ".")
        shortReg = Right(email, Len(email) - lastDotPos)
        If InStr(shortReg, "com") > 0 Or InStr(shortReg, "co") > 0 Then
        ''If shortReg = "com" Then
            If shortReg <> "com" Then
               i.Value = Left(email, lastDotPos) & "com"
            End If
            
         ElseIf InStr(shortReg, "pl") > 0 Or InStr(shortReg, "p") > 0 Or InStr(shortReg, "l") > 0 Then
            If shortReg <> "pl" Then
               i.Value = Left(email, lastDotPos) & "pl"
            End If
        End If
    Next
    
    
    Set lastRdE = Cells(Rows.Count, "E").End(xlUp)
    Set R = Range("E1", lastRdE)
    
    
    Workbooks("obecni_przygotowanie_eksportu.xlsx").Worksheets("Arkusz4").Range("A1", Cells(lastRdE.Row, "Q")).Copy
  
    Workbooks("import obecnych gości.xlsx").Worksheets("Arkusz2").Range("A1").PasteSpecial Paste:=xlPasteValues
    
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    
    Application.WindowState = xlNormal
    Windows("import obecnych gości.xlsx").Activate
    Columns("J:J").Select
    Selection.NumberFormat = "yyyy-mm-dd;@"
    
    Windows("obecni_przygotowanie_eksportu.xlsx").Activate
    Cells.Select
    Selection.Delete Shift:=xlUp
         
    
    'closing workbooks
    ActiveWorkbook.Close SaveChanges:=True
    Workbooks("import obecnych gości.xlsx").Close SaveChanges:=True
'   Workbooks("obecni_przygotowanie_eksportu.xlsx").Close SaveChanges:=True 'out of range
    
    Workbooks("obecni_makro.xlsm").Close SaveChanges:=True
    
End Sub
