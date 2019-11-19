Option Base 0

Type Color
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Global playing As Boolean
Global generation1(100, 100) As Boolean
Global generation2(100, 100) As Boolean
Global aliveCells As Integer
Global deadCells As Integer
Global generation As Integer
Global aliveCellColor As Color
Global deadCellColor As Color

Global zoneDeJeu As Range
Global generationCell As Range
Global aliveCellsCell As Range
Global deadCellsCell As Range


Sub main()
    Set zoneDeJeu = Range("F1:DA100")
    Set generationCell = Range("D16")
    Set aliveCellsCell = Range("D19")
    Set deadCellsCell = Range("D21")
    aliveCellColor.Red = 119
    aliveCellColor.Green = 117
    aliveCellColor.Blue = 255
    deadCellColor.Red = 215
    deadCellColor.Green = 214
    deadCellColor.Blue = 255
    If playing = True Then
        Call changeGeneration
        Call updateTable
    End If
End Sub

Sub generateRandom()
    Dim randomizedNumber As Integer
    For y = 0 To UBound(generation1, 2)
        For x = 0 To UBound(generation1, 1)
            randomizedNumber = Int(5 * Rnd) + 1
            If randomizedNumber = 1 Then
                generation1(x, y) = True
                generation2(x, y) = False
            Else
                generation1(x, y) = False
                generation2(x, y) = False
            End If
        Next
    Next
    Call updateTable
End Sub

Sub generateEmpty()
    For y = 0 To UBound(generation1, 2)
        For x = 0 To UBound(generation1, 1)
            generation1(x, y) = False
            generation2(x, y) = False
        Next
    Next
    Call updateTable
End Sub

Sub changeGeneration()
    Dim coords(2) As Integer
    For y = 0 To UBound(generation1, 2)
     For x = 0 To UBound(generation1, 1)
            coords(0) = x
            coords(1) = y
            generation2(x, y) = checkAlive(coords)
        Next
    Next
    For y = 0 To UBound(generation1, 2)
        For x = 0 To UBound(generation1, 1)
            generation1(x, y) = generation2(x, y)
            generation2(x, y) = False
        Next
    Next
    generation = generation + 1
    Call updateTable
End Sub

Function checkAlive(coords() As Integer) As Boolean
    Dim alive As Boolean
    Dim neighboursAlive As Integer
    
    If coords(0) > 0 Then
        If coords(1) > 0 Then
            If generation1(coords(0) - 1, coords(1) - 1) = True Then
                neighboursAlive = neighboursAlive + 1
            End If
        End If
        If coords(1) < UBound(generation1, 2) - 1 Then
            If generation1(coords(0) - 1, coords(1) + 1) = True Then
                neighboursAlive = neighboursAlive + 1
            End If
        End If
        If generation1(coords(0) - 1, coords(1)) = True Then
            neighboursAlive = neighboursAlive + 1
        End If
    End If
    
    If coords(0) < UBound(generation1, 1) - 1 Then
        If coords(1) > 0 Then
            If generation1(coords(0) + 1, coords(1) - 1) = True Then
                neighboursAlive = neighboursAlive + 1
            End If
        End If
        If coords(1) < UBound(generation1, 2) - 1 Then
            If generation1(coords(0) + 1, coords(1) + 1) = True Then
                neighboursAlive = neighboursAlive + 1
            End If
        End If
        If generation1(coords(0) + 1, coords(1)) = True Then
            neighboursAlive = neighboursAlive + 1
        End If
    End If
    
    If coords(1) > 0 Then
        If generation1(coords(0), coords(1) - 1) = True Then
            neighboursAlive = neighboursAlive + 1
        End If
    End If
    
    If coords(1) < UBound(generation1, 2) - 1 Then
        If generation1(coords(0), coords(1) + 1) = True Then
            neighboursAlive = neighboursAlive + 1
        End If
    End If
    
    If neighboursAlive = 3 Then
        alive = True
    ElseIf neighboursAlive = 2 Then
        alive = generation1(coords(0), coords(1))
    ElseIf neighboursAlive < 2 Or neighboursAlive > 3 Then
        alive = False
    End If
    checkAlive = alive
    Exit Function
End Function

Sub updateTable()
    aliveCells = 0
    Set zoneDeJeu = Range("F1:DA100")
    Set generationCell = Range("D10")
    Set aliveCellsCell = Range("D13")
    Set deadCellsCell = Range("D16")
    aliveCellColor.Red = 119
    aliveCellColor.Green = 117
    aliveCellColor.Blue = 255
    deadCellColor.Red = 215
    deadCellColor.Green = 214
    deadCellColor.Blue = 255
    For Each c In zoneDeJeu
        If generation1(c.Row, c.Column - 5) = True Then
            c.Value = 1
            'c.Select
            'With selection.Interior
            '    .Color = RGB(aliveCellColor.Red, aliveCellColor.Green, aliveCellColor.Blue)
            'End With
            'With selection.Font
            '    .Color = RGB(255, 255, 255)
            'End With
        Else
            c.Value = 0
            'c.Select
            'With selection.Interior
            '    .Color = RGB(deadCellColor.Red, deadCellColor.Green, deadCellColor.Blue)
            'End With
            'With selection.Font
            '    .Color = RGB(aliveCellColor.Red, aliveCellColor.Green, aliveCellColor.Blue)
            'End With
        End If
        If generation1(c.Row, c.Column - 5) = True Then
            aliveCells = aliveCells + 1
        End If
    Next c
    generationCell.Value = generation
    aliveCellsCell.Value = aliveCells
    deadCellsCell.Value = 10000 - aliveCells
End Sub

Sub load()
    Dim import As Range
    Set import = Worksheets("Import").Range("F1:DA100")
    For Each c In import
        generation1(c.Row, c.Column - 5) = c.Value
    Next c
    Call updateTable
End Sub
