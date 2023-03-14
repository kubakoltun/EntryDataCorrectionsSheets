Public shapeState As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
'lista wieloelementowa z mozliwoscia usuniecia wartosci

If Not Intersect(ActiveCell, Range("J:J")) Is Nothing Or Not Intersect(ActiveCell, Range("M:M")) Is Nothing Then
            On Error GoTo 0
        Dim xRng As Range
        Dim xValue1 As String
        Dim xValue2 As StringPublic shapeState As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
'lista wieloelementowa z mozliwoscia usuniecia wartosci

If Not Intersect(ActiveCell, Range("J:J")) Is Nothing Or Not Intersect(ActiveCell, Range("M:M")) Is Nothing Then
            On Error GoTo 0
        Dim xRng As Range
        Dim xValue1 As String
        Dim xValue2 As String
        Dim semiColonCnt As Integer
        If Target.Count > 1 Then Exit Sub
        On Error Resume Next
        'Set xRng = Cells.SpecialCells(xlCellTypeAllValidation)
        Set xRng = Cells.Range("I6:I10")
        If xRng Is Nothing Then Exit Sub
        Application.EnableEvents = False
        If Application.Intersect(Target, xRng) Then
        xValue2 = Target.Value
        Application.Undo
        xValue1 = Target.Value
        Target.Value = xValue2
        If xValue1 <> "" Then
        If xValue2 <> "" Then
        If xValue1 = xValue2 Or xValue1 = xValue2 & "," Or xValue1 = xValue2 & ", " Then
        xValue1 = Replace(xValue1, ", ", "")
        xValue1 = Replace(xValue1, ",", "")
        Target.Value = xValue1
        ElseIf InStr(1, xValue1, ", " & xValue2) Then
        xValue1 = Replace(xValue1, xValue2, "")
        Target.Value = xValue1
        ElseIf InStr(1, xValue1, xValue2 & ",") Then
        xValue1 = Replace(xValue1, xValue2, "")
        Target.Value = xValue1
        Else
        Target.Value = xValue1 & ", " & xValue2
        End If
        Target.Value = Replace(Target.Value, ",,", ",")
        Target.Value = Replace(Target.Value, ", ,", ",")
        If InStr(1, Target.Value, "; ") = 1 Then
        Target.Value = Replace(Target.Value, ", ", "", 1, 1)
        End If
        If InStr(1, Target.Value, ";") = 1 Then
        Target.Value = Replace(Target.Value, ",", "", 1, 1)
        End If
        semiColonCnt = 0
        
        For I = 1 To Len(Target.Value)
        If InStr(I, Target.Value, ",") Then
        semiColonCnt = semiColonCnt + 1
        End If
        Next I
        If semiColonCnt = 1 Then
        Target.Value = Replace(Target.Value, ", ", "")
        Target.Value = Replace(Target.Value, ",", "")
        End If
        End If
        End If
        End If
        Application.EnableEvents = True
    End If

End Sub



Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'ksztalt wyswietlany przy kazdym kliknieciu

Set curCell = ActiveCell

If Not Intersect(ActiveCell, Range("I:I")) Is Nothing Or Not Intersect(ActiveCell, Range("J:J")) Is Nothing Then
            On Error GoTo 0
            If curCell.Value = "NIE" Or curCell.Offset(0, -1).Value = "NIE" Then
                If curCell.Offset(0, -1).Value = "NIE" Then
                    curCell.Offset(0, 0).Value = "nd."
                    curCell.Offset(0, 1).Value = "nd."
                Else
                    curCell.Offset(0, 1).Value = "nd."
                    curCell.Offset(0, 2).Value = "nd."
                End If
            End If
            
            
End If

If Not Intersect(ActiveCell, Range("J:J")) Is Nothing Or Not Intersect(ActiveCell, Range("K:K")) Is Nothing Then
            If curCell.Offset(0, 0).Value <> "" And curCell.Offset(0, 3).Value = "" And Intersect(ActiveCell, Range("K:K")) Is Nothing And curCell.Offset(0, 0).Value <> "nd." Then
                curCell.Offset(0, 3).Value = curCell.Offset(0, 0).Value
            End If
            
            If curCell.Offset(0, -1).Value <> "" And curCell.Offset(0, 2).Value = "" And Not Intersect(ActiveCell, Range("K:K")) Is Nothing And curCell.Offset(0, 0).Value <> "nd." Then
                curCell.Offset(0, 2).Value = curCell.Offset(0, -1).Value
            End If
End If

If Not Intersect(ActiveCell, Range("O:O")) Is Nothing Or Not Intersect(ActiveCell, Range("R:R")) Is Nothing Then
            On Error GoTo 0


    If ActiveSheet.Shapes("Prostokat1").TextFrame2.TextRange.Characters.Text = "Kliknij, aby wprowadzić wybrane podfoldery" Or ActiveSheet.Shapes("Prostokat1").TextFrame2.TextRange.Characters.Text = "Kliknij, aby wprowadzić wybrane tablice Trello" Then
       ActiveSheet.ListBox2.Visible = True
    End If
    
    If Not Intersect(ActiveCell, Range("O1:O5")) Is Nothing Or Not Intersect(ActiveCell, Range("R1:R5")) Is Nothing Then
        ActiveSheet.ListBox2.Visible = False
    End If

     With ActiveSheet.Shapes("Prostokat1")
     If Not Intersect(ActiveCell, Range("O1:O5")) Is Nothing Or Not Intersect(ActiveCell, Range("R1:R5")) Is Nothing Then
        .Visible = False
    Else
        .Visible = True
    End If
        .Top = Target.Offset(0).Top

        .Left = Target.Offset(, 1).Left

    End With
    'ActiveSheet.ListBox2.Visible = True
    
    
    
    Dim validInput As Boolean
    Dim cell As Range
    Dim lR As Long
    
    validInput = False
    'MsgBox (Worksheets("baza").Range("F2").Value)
    'blok przypisywania podfolderow dla O:O - podfoldery
  If Not Intersect(ActiveCell, Range("O:O")) Is Nothing Then
    
                                                                                                            If InStr(curCell.Offset(0, -2), "Value1") > 0 Then 'Worksheets("baza").Range("H2").Value
    
        lR = Worksheets("baza").Cells(Rows.Count, "Z").End(xlUp).Row
        
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("Z2:Z" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("Z2:Z" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H2").Value) > 0 And Worksheets("baza").Range("H2").Value <> "" Then
        
        lR = Worksheets("baza").Cells(Rows.Count, "I").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("I2:I" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("I2:I" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H3").Value) > 0 And Worksheets("baza").Range("H3").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "J").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("J2:J" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("J2:J" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H4").Value) > 0 And Worksheets("baza").Range("H4").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "K").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("K2:K" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("K2:K" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H5").Value) > 0 And Worksheets("baza").Range("H5").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "L").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("L2:L" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("L2:L" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H6").Value) > 0 And Worksheets("baza").Range("H6").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "M").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("M2:M" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("M2:M" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H7").Value) > 0 And Worksheets("baza").Range("H7").Value <> "" Then
        
        lR = Worksheets("baza").Cells(Rows.Count, "N").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("N2:N" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("N2:N" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H8").Value) > 0 And Worksheets("baza").Range("H8").Value <> "" Then
   
        lR = Worksheets("baza").Cells(Rows.Count, "O").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("O2:O" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("O2:O" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H9").Value) > 0 And Worksheets("baza").Range("H9").Value <> "" Then
        
        lR = Worksheets("baza").Cells(Rows.Count, "P").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("P2:P" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("P2:P" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H10").Value) > 0 And Worksheets("baza").Range("H10").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "Q").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("Q2:Q" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("Q2:Q" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H11").Value) > 0 And Worksheets("baza").Range("H11").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "R").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("R2:R" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("R2:R" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H12").Value) > 0 And Worksheets("baza").Range("H12").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "S").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("S2:S" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("S2:S" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H13").Value) > 0 And Worksheets("baza").Range("H13").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "T").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("T2:T" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("T2:T" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H14").Value) > 0 And Worksheets("baza").Range("H14").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "U").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("U2:U" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("U2:U" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H15").Value) > 0 And Worksheets("baza").Range("H15").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "V").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("V2:V" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("V2:V" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H16").Value) > 0 And Worksheets("baza").Range("H16").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "W").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("W2:W" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("W2:W" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H17").Value) > 0 And Worksheets("baza").Range("H17").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AB").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AB2:AB" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AB2:AB" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H18").Value) > 0 And Worksheets("baza").Range("H18").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AC").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AC2:AC" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AC2:AC" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H19").Value) > 0 And Worksheets("baza").Range("H19").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AD").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AD2:AD" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AD2:AD" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H20").Value) > 0 And Worksheets("baza").Range("H20").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AE").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AE2:AE" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AE2:AE" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H21").Value) > 0 And Worksheets("baza").Range("H21").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AF").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AF2:AF" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AF2:AF" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H22").Value) > 0 And Worksheets("baza").Range("H22").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AG").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AG2:AG" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AG2:AG" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H23").Value) > 0 And Worksheets("baza").Range("H23").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AH").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AH2:AH" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AH2:AH" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H24").Value) > 0 And Worksheets("baza").Range("H24").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AI").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AI2:AI" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AI2:AI" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H25").Value) > 0 And Worksheets("baza").Range("H25").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AJ").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AJ2:AJ" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AJ2:AJ" & lR).Value
        End If
        validInput = True
    End If
    
    If validInput = False Then
        ListBox2.Clear
    End If
 End If
 
 'Dim validInput As Boolean
    'Dim cell As Range
    Dim lR1 As Long
    
    validInput = False
 
 If Not Intersect(ActiveCell, Range("R:R")) Is Nothing Then
    
    If InStr(curCell.Offset(0, -1), "TAK") > 0 Or InStr(curCell.Offset(0, -1), "NIE") > 0 Then
        lR1 = Worksheets("baza").Cells(Rows.Count, "X").End(xlUp).Row
   
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("X2:X" & lR1)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("X2:X" & lR1).Value
        End If
        validInput = True
    End If
    If validInput = False Then
        ListBox2.Clear
    End If
    
 End If
    
    'ListBox2.List = Worksheets("baza").Range("I2:I7").Value
    'curCell.Offset(0, -2).Select
    
    ListBox2.Height = 108
    ListBox2.Width = 220
    
    ListBox2.Top = Target.Offset(0).Top
    ListBox2.Left = Target.Offset(, 2).Left
    
Else
    ActiveSheet.Shapes("Prostokat1").Visible = False
    ActiveSheet.ListBox2.Visible = False
        
    
End If

End Sub

Private Sub Worksheet_SelectionChangeL(ByVal Target As Excel.Range)

        On Error GoTo 0

        With Cells(Windows(1).ScrollRow, Windows(1).ScrollColumn)

            ListBox2.Top = Target.Offset(0).Top

            ListBox2.Left = Target.Offset(, 2).Left

       End With

End Sub


        Dim semiColonCnt As Integer
        If Target.Count > 1 Then Exit Sub
        On Error Resume Next
        'Set xRng = Cells.SpecialCells(xlCellTypeAllValidation)
        Set xRng = Cells.Range("I6:I10")
        If xRng Is Nothing Then Exit Sub
        Application.EnableEvents = False
        If Application.Intersect(Target, xRng) Then
        xValue2 = Target.Value
        Application.Undo
        xValue1 = Target.Value
        Target.Value = xValue2
        If xValue1 <> "" Then
        If xValue2 <> "" Then
        If xValue1 = xValue2 Or xValue1 = xValue2 & "," Or xValue1 = xValue2 & ", " Then
        xValue1 = Replace(xValue1, ", ", "")
        xValue1 = Replace(xValue1, ",", "")
        Target.Value = xValue1
        ElseIf InStr(1, xValue1, ", " & xValue2) Then
        xValue1 = Replace(xValue1, xValue2, "")
        Target.Value = xValue1
        ElseIf InStr(1, xValue1, xValue2 & ",") Then
        xValue1 = Replace(xValue1, xValue2, "")
        Target.Value = xValue1
        Else
        Target.Value = xValue1 & ", " & xValue2
        End If
        Target.Value = Replace(Target.Value, ",,", ",")
        Target.Value = Replace(Target.Value, ", ,", ",")
        If InStr(1, Target.Value, "; ") = 1 Then
        Target.Value = Replace(Target.Value, ", ", "", 1, 1)
        End If
        If InStr(1, Target.Value, ";") = 1 Then
        Target.Value = Replace(Target.Value, ",", "", 1, 1)
        End If
        semiColonCnt = 0
        
        For I = 1 To Len(Target.Value)
        If InStr(I, Target.Value, ",") Then
        semiColonCnt = semiColonCnt + 1
        End If
        Next I
        If semiColonCnt = 1 Then
        Target.Value = Replace(Target.Value, ", ", "")
        Target.Value = Replace(Target.Value, ",", "")
        End If
        End If
        End If
        End If
        Application.EnableEvents = True
    End If

End Sub



Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'ksztalt wyswietlany przy kazdym kliknieciu

Set curCell = ActiveCell

If Not Intersect(ActiveCell, Range("I:I")) Is Nothing Or Not Intersect(ActiveCell, Range("J:J")) Is Nothing Then
            On Error GoTo 0
            If curCell.Value = "NIE" Or curCell.Offset(0, -1).Value = "NIE" Then
                If curCell.Offset(0, -1).Value = "NIE" Then
                    curCell.Offset(0, 0).Value = "nd."
                    curCell.Offset(0, 1).Value = "nd."
                Else
                    curCell.Offset(0, 1).Value = "nd."
                    curCell.Offset(0, 2).Value = "nd."
                End If
            End If
            
            
End If

If Not Intersect(ActiveCell, Range("J:J")) Is Nothing Or Not Intersect(ActiveCell, Range("K:K")) Is Nothing Then
            If curCell.Offset(0, 0).Value <> "" And curCell.Offset(0, 3).Value = "" And Intersect(ActiveCell, Range("K:K")) Is Nothing And curCell.Offset(0, 0).Value <> "nd." Then
                curCell.Offset(0, 3).Value = curCell.Offset(0, 0).Value
            End If
            
            If curCell.Offset(0, -1).Value <> "" And curCell.Offset(0, 2).Value = "" And Not Intersect(ActiveCell, Range("K:K")) Is Nothing And curCell.Offset(0, 0).Value <> "nd." Then
                curCell.Offset(0, 2).Value = curCell.Offset(0, -1).Value
            End If
End If

If Not Intersect(ActiveCell, Range("O:O")) Is Nothing Or Not Intersect(ActiveCell, Range("R:R")) Is Nothing Then
            On Error GoTo 0


    If ActiveSheet.Shapes("Prostokat1").TextFrame2.TextRange.Characters.Text = "Kliknij, aby wprowadzić wybrane podfoldery" Or ActiveSheet.Shapes("Prostokat1").TextFrame2.TextRange.Characters.Text = "Kliknij, aby wprowadzić wybrane tablice Trello" Then
       ActiveSheet.ListBox2.Visible = True
    End If
    
    If Not Intersect(ActiveCell, Range("O1:O5")) Is Nothing Or Not Intersect(ActiveCell, Range("R1:R5")) Is Nothing Then
        ActiveSheet.ListBox2.Visible = False
    End If

     With ActiveSheet.Shapes("Prostokat1")
     If Not Intersect(ActiveCell, Range("O1:O5")) Is Nothing Or Not Intersect(ActiveCell, Range("R1:R5")) Is Nothing Then
        .Visible = False
    Else
        .Visible = True
    End If
        .Top = Target.Offset(0).Top

        .Left = Target.Offset(, 1).Left

    End With
    'ActiveSheet.ListBox2.Visible = True
    
    
    
    Dim validInput As Boolean
    Dim cell As Range
    Dim lR As Long
    
    validInput = False
    'MsgBox (Worksheets("baza").Range("F2").Value)
    'blok przypisywania podfolderow dla O:O - podfoldery
  If Not Intersect(ActiveCell, Range("O:O")) Is Nothing Then
    
If InStr(curCell.Offset(0, -2), "Value1") > 0 Then 'Worksheets("baza").Range("H2").Value
    
        lR = Worksheets("baza").Cells(Rows.Count, "Z").End(xlUp).Row
        
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("Z2:Z" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("Z2:Z" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H2").Value) > 0 And Worksheets("baza").Range("H2").Value <> "" Then
        
        lR = Worksheets("baza").Cells(Rows.Count, "I").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("I2:I" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("I2:I" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H3").Value) > 0 And Worksheets("baza").Range("H3").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "J").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("J2:J" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("J2:J" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H4").Value) > 0 And Worksheets("baza").Range("H4").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "K").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("K2:K" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("K2:K" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H5").Value) > 0 And Worksheets("baza").Range("H5").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "L").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("L2:L" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("L2:L" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H6").Value) > 0 And Worksheets("baza").Range("H6").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "M").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("M2:M" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("M2:M" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H7").Value) > 0 And Worksheets("baza").Range("H7").Value <> "" Then
        
        lR = Worksheets("baza").Cells(Rows.Count, "N").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("N2:N" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("N2:N" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H8").Value) > 0 And Worksheets("baza").Range("H8").Value <> "" Then
   
        lR = Worksheets("baza").Cells(Rows.Count, "O").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("O2:O" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("O2:O" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H9").Value) > 0 And Worksheets("baza").Range("H9").Value <> "" Then
        
        lR = Worksheets("baza").Cells(Rows.Count, "P").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("P2:P" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("P2:P" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H10").Value) > 0 And Worksheets("baza").Range("H10").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "Q").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("Q2:Q" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("Q2:Q" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H11").Value) > 0 And Worksheets("baza").Range("H11").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "R").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("R2:R" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("R2:R" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H12").Value) > 0 And Worksheets("baza").Range("H12").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "S").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("S2:S" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("S2:S" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H13").Value) > 0 And Worksheets("baza").Range("H13").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "T").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("T2:T" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("T2:T" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H14").Value) > 0 And Worksheets("baza").Range("H14").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "U").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("U2:U" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("U2:U" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H15").Value) > 0 And Worksheets("baza").Range("H15").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "V").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("V2:V" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("V2:V" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H16").Value) > 0 And Worksheets("baza").Range("H16").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "W").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("W2:W" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("W2:W" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H17").Value) > 0 And Worksheets("baza").Range("H17").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AB").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AB2:AB" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AB2:AB" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H18").Value) > 0 And Worksheets("baza").Range("H18").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AC").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AC2:AC" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AC2:AC" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H19").Value) > 0 And Worksheets("baza").Range("H19").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AD").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AD2:AD" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AD2:AD" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H20").Value) > 0 And Worksheets("baza").Range("H20").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AE").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AE2:AE" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AE2:AE" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H21").Value) > 0 And Worksheets("baza").Range("H21").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AF").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AF2:AF" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AF2:AF" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H22").Value) > 0 And Worksheets("baza").Range("H22").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AG").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AG2:AG" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AG2:AG" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H23").Value) > 0 And Worksheets("baza").Range("H23").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AH").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AH2:AH" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AH2:AH" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H24").Value) > 0 And Worksheets("baza").Range("H24").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AI").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AI2:AI" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AI2:AI" & lR).Value
        End If
        validInput = True
    End If
    
    If InStr(curCell.Offset(0, -2), Worksheets("baza").Range("H25").Value) > 0 And Worksheets("baza").Range("H25").Value <> "" Then
        lR = Worksheets("baza").Cells(Rows.Count, "AJ").End(xlUp).Row
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("AJ2:AJ" & lR)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("AJ2:AJ" & lR).Value
        End If
        validInput = True
    End If
    
    If validInput = False Then
        ListBox2.Clear
    End If
 End If
 
 'Dim validInput As Boolean
    'Dim cell As Range
    Dim lR1 As Long
    
    validInput = False
 
 If Not Intersect(ActiveCell, Range("R:R")) Is Nothing Then
    
    If InStr(curCell.Offset(0, -1), "TAK") > 0 Or InStr(curCell.Offset(0, -1), "NIE") > 0 Then
        lR1 = Worksheets("baza").Cells(Rows.Count, "X").End(xlUp).Row
   
        If validInput = True Then
            For Each cell In Worksheets("baza").Range("X2:X" & lR1)
                ListBox2.AddItem cell.Value
            Next
        Else
            ListBox2.List = Worksheets("baza").Range("X2:X" & lR1).Value
        End If
        validInput = True
    End If
    If validInput = False Then
        ListBox2.Clear
    End If
    
 End If
    
    'ListBox2.List = Worksheets("baza").Range("I2:I7").Value
    'curCell.Offset(0, -2).Select
    
    ListBox2.Height = 108
    ListBox2.Width = 220
    
    ListBox2.Top = Target.Offset(0).Top
    ListBox2.Left = Target.Offset(, 2).Left
    
Else
    ActiveSheet.Shapes("Prostokat1").Visible = False
    ActiveSheet.ListBox2.Visible = False
        
    
End If

End Sub

Private Sub Worksheet_SelectionChangeL(ByVal Target As Excel.Range)

        On Error GoTo 0

        With Cells(Windows(1).ScrollRow, Windows(1).ScrollColumn)

            ListBox2.Top = Target.Offset(0).Top

            ListBox2.Left = Target.Offset(, 2).Left

       End With

End Sub

