Attribute VB_Name = "Main"
Public Board() As Boolean
Public NextBoard() As Boolean

Public boardsize As Integer
Public generations As Integer
Public generation As Integer
Public endless As Boolean

Public alive As Integer
Public dead As Integer

Public lastX As Integer
Public lastY As Integer

Public mousedown As Boolean



Public Sub Render() 'responsible for rendering the array
'for now, Game.Print to render
Game.Cls
For Y = 1 To UBound(Board, 1)
Game.CurrentX = 0.375
    For X = 1 To UBound(Board, 2)
        If Board(X, Y) = True Then
            Game.ForeColor = vbBlue
            Game.Print "O";
        Else
            Game.ForeColor = vbRed
            Game.Print "X";
        End If
        Game.CurrentX = (X * 1.375 * 1) + 0.375
    Next X
    Game.Print vbCrLf
    Game.CurrentY = Y * 1.625 * 0.45
    Next Y
End Sub

Public Sub FullPopulate(value As Boolean) ' for dev purposes, will fill the array with a value
For Y = 1 To UBound(Board, 1)
    For X = 1 To UBound(Board, 2)
        Board(X, Y) = value
        NextBoard(X, Y) = value
    Next X
Next Y
End Sub

Public Sub ClickBoard(X As Single, Y As Single)
whichX = Round(((X) / UBound(Board, 1)) * (22 * (boardsize / 30))) '22 is arbitrary...chosen cause it works, scales to size
whichY = Round((Y / UBound(Board, 2)) * (41 * (boardsize / 30))) '41 is arbitrary...chosen cause it works, scales to size
'make sure there is no 0 value
If whichX = 0 Then
    whichX = 1
End If
If whichY = 0 Then
    whichY = 1
End If
    'change the value of the clicked place
If (whichX <> lastX) Or (whichY <> lastY) Then
    If (whichX <= boardsize) And (whichY <= boardsize) Then
        'clicked in board bounds
        Board(whichX, whichY) = Not Board(whichX, whichY)
        Call Render
        
        If Board(whichX, whichY) = False Then
            dead = dead + 1
            alive = alive - 1
        Else
            dead = dead - 1
            alive = alive + 1
        End If
        Call DisplayStats
    End If
    lastX = whichX
    lastY = whichY
End If
End Sub

Public Sub CalculateChange()
'configure the new board, saving results to the new board
For Y = 1 To UBound(Board, 1)
    For X = 1 To UBound(Board, 2)
        'this is each induvidual cell, its number of living neighbours must be calculated
        neighbours = 0
        For RY = (Y - 1) To (Y + 1)
            For RX = (X - 1) To (X + 1)
                'this for loop iterates through surrounding cells
                If ((RY <> Y) Or (RX <> X)) And ((RY >= 1) And RY <= 30) And ((RX >= 1) And RX <= 30) Then 'this makes sure cell is valid to be checked, that is, not the original cell and is within bounds
                    If Board(RX, RY) = True Then ' if living neighbour
                        neighbours = neighbours + 1
                    End If
                End If
            Next RX
        Next RY
        If Board(X, Y) = True Then 'cell is alive
            If neighbours < 2 Or neighbours > 3 Then
                NextBoard(X, Y) = False 'Rule 1: if there are less than 2 or more than 3 live neighbours, kill cell
            End If
        Else 'cell is dead
            If neighbours = 3 Then
                NextBoard(X, Y) = True 'Rule 1: if a dead cell has 3 live neighbours, cell comes alive
            End If
        End If
    Next X
Next Y

'transfer newboard values to actual board
alive = 0
dead = boardsize ^ 2
For Y = 1 To UBound(Board, 1)
    For X = 1 To UBound(Board, 2)
        Board(X, Y) = NextBoard(X, Y)
        If Board(X, Y) = True Then
            alive = alive + 1
            dead = dead - 1
        End If
        
    Next X
Next Y
Call Render
End Sub

Public Sub DisplayStats()
    Game.lblGeneration.Caption = "Generation: " + CStr(generation)
    Game.lblAlive.Caption = "Alive: " + CStr(alive)
    Game.lblDead.Caption = "Dead: " + CStr(dead)
End Sub
