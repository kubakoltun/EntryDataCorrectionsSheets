Public shapeState As Boolean


Sub Prostokat1_Click()
'otworzenie listy po kliknieciu na wskazany ksztalt
'wproawzdenie i odznaczenie wartosci listy

Dim xSelShp As Shape, xSelLst As Variant, I, J As Integer
Dim xV As String
Dim curCell As Range
Set xSelShp = ActiveSheet.Shapes(Application.Caller)
Set curCell = ActiveCell

Set xLstBox = ActiveSheet.ListBox2

If xLstBox.Visible = False Then
    xLstBox.Visible = True
   
    'bigger
    xSelShp.LockAspectRatio = msoFalse
    xSelShp.Height = 68
    xSelShp.Width = 105
    xSelShp.Placement = xlMove
    xSelShp.LockAspectRatio = msoTrue
    xSelShp.TextFrame.Characters.Font.ColorIndex = 2
    xSelShp.Fill.ForeColor.RGB = RGB(64, 188, 92)
    
    If Not Intersect(curCell, Range("R:R")) Is Nothing Then
        xSelShp.TextFrame2.TextRange.Characters.Text = "Kliknij, aby wprowadzić wybrane tablice Trello"
    Else
        xSelShp.TextFrame2.TextRange.Characters.Text = "Kliknij, aby wprowadzić wybrane podfoldery"
    End If
    
    shapeState = True
    xSelShp.TextFrame.Characters.Font.Italic = False
    xSelShp.Line.Weight = 1
    xSelShp.Line.ForeColor.RGB = RGB(64, 188, 92)
    'xSelShp.AutoShapeType = msoShapeRoundedRectangle
    
    xStr = ""
    xStr = curCell.Value
    
    If xStr <> "" Then
         xArr = Split(xStr, ";")
    For I = xLstBox.ListCount - 1 To 0 Step -1
        xV = xLstBox.List(I)
        For J = 0 To UBound(xArr)
            If xArr(J) = xV Then
              xLstBox.Selected(I) = True
              Exit For
            End If
        Next
    Next I
    End If
Else
    xLstBox.Visible = False
    'smaller
    xSelShp.LockAspectRatio = msoFalse
    xSelShp.Height = 40
    xSelShp.Width = 70
    xSelShp.Placement = xlMove
    xSelShp.LockAspectRatio = msoTrue
    xSelShp.TextFrame.Characters.Font.Italic = True
    xSelShp.Line.Weight = 1.5
    xSelShp.Line.ForeColor.RGB = RGB(0, 0, 0)
    'xSelShp.AutoShapeType = msoShapeRoundedRectangle
    xSelShp.TextFrame.Characters.Font.ColorIndex = 1
    'bezowy z excela 256, 220, 124
    xSelShp.Fill.ForeColor.RGB = RGB(8, 172, 244)
    
    If Not Intersect(curCell, Range("R:R")) Is Nothing Then
        xSelShp.TextFrame2.TextRange.Characters.Text = "Wybierz tablice"
    Else
        xSelShp.TextFrame2.TextRange.Characters.Text = "Wybierz podfoldery"
    End If
    
    shapeState = False
    For I = xLstBox.ListCount - 1 To 0 Step -1
        If xLstBox.Selected(I) = True Then
        xSelLst = xLstBox.List(I) & "; " & xSelLst
        xLstBox.Selected(I) = False
        End If
    Next I
    If xSelLst <> "" Then
        If Not Intersect(curCell, Range("O:O")) Is Nothing Or Not Intersect(curCell, Range("R:R")) Is Nothing Then
            curCell = Mid(xSelLst, 1, Len(xSelLst) - 1)
        End If
    Else
        If Not Intersect(curCell, Range("O:O")) Is Nothing Or Not Intersect(curCell, Range("R:R")) Is Nothing Then
            curCell = ""
        End If
    End If
End If
End Sub

