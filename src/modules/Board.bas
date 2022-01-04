Attribute VB_Name = "Board"
' Copyright 2021 Alejandro D.
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
Option Explicit

Public Sub setTilesProperties(firstTileAddress As String, Optional prevBoardMines As Variant)
    ' Callers: Logic.tileLeftMouseUp
    ' Sets the Mine property of the corresponding tiles-
    ' sets the Number property of the tiles according to
    ' the number of adjacent mines
    
    Dim mines As Collection, mineAddr As Variant
    Dim numbers() As Variant
    Dim i As Integer, j As Integer
    
    ' When replay use last game Mines collection
    If IsMissing(prevBoardMines) Then
        Set mines = generateMines(firstTileAddress)
    Else
        Set mines = prevBoardMines
    End If
    
    ' Save the mines to use in replay
    Set Game.mines = mines
    
    numbers = calculateNumbers(mines)
    
    ' Set Mine property
    For Each mineAddr In mines
        Game.tiles(mineAddr).mine = True
    Next mineAddr
    
    ' Set Number property
    For i = 0 To Game.boardY - 1
        For j = 0 To Game.boardX - 1
            If IsEmpty(numbers(i, j)) Then
                Game.tiles(addr(i, j)).Number = 0
            Else
                Game.tiles(addr(i, j)).Number = numbers(i, j)
            End If
        Next j
    Next i
End Sub

Private Function generateMines(firstAddress As String) As Collection
    ' Returns collection of RANDOMNLY generate mine addresses
    
    Dim i As Integer, column As Integer, row As Integer
    Dim mines As Collection
    Set mines = New Collection
    
    For i = 0 To Game.numberOfMines - 1
GetRandom:
        Call Randomize
        column = Int(rnd() * Game.boardY)
        Call Randomize
        row = Int(rnd() * Game.boardX)
        ' Re-generate a mine address if the address is already in the collection or
        ' the address matches the first-cliked tile address
        If Not (alreadyInMines(mines, addr(column, row)) Or (addr(column, row) = firstAddress)) Then
            mines.add addr(column, row)
        Else
            GoTo GetRandom
        End If
    Next i
    Set generateMines = mines
End Function

Private Function calculateNumbers(mines As Collection) As Variant()
    ' Returns a 2-dimensional array of numbers
    ' corresponding to the number of adjacent mines
    ' the array index's themselves serve as the tiles address's
    
    Dim mine As Variant
    Dim mineAddr As String
    Dim numbers() As Variant: ReDim numbers(0 To Game.boardY - 1, 0 To Game.boardX - 1)
    Dim i As Integer, Y As Integer, X As Integer
    For Each mine In mines
        mineAddr = CStr(mine)
        For i = 0 To 7
            Y = inx(mineAddr).Y
            X = inx(mineAddr).X
            If ((X + Moves(i).X < Game.boardX And X + Moves(i).X >= 0) And (Y + Moves(i).Y < Game.boardY And Y + Moves(i).Y >= 0)) Then
                numbers(Y + Moves(i).Y, X + Moves(i).X) = numbers(Y + Moves(i).Y, X + Moves(i).X) + 1
            End If
        Next i
    Next mine
    calculateNumbers = numbers
End Function

Public Sub handleWin()
    ' Stop the in-game timer
    Call Time.stopTimer
    
    ' Place yellow-flags in unflagged mines
    Dim t As MinesweeperTile
    For Each t In Game.tiles
        If t.state = Hidden And t.mine Then t.Image.Picture = Sprite.useSprite("uf")
    Next t
    
    ' Remove the transparent image to disable input
    Call boardForm.controls.Remove("imgTransp")
    Call Menu.setMenuPicture("woke")
    
    ' Handle stats
    Call Stats.onWin
    
    If Game.Settings.ShowStatsOnGameEnd Then
        endForm.Caption = "Win"
        endForm.show
    End If
End Sub

Public Sub handleLose(tile As MinesweeperTile)
    ' Stop de in-game timer
    Call Time.stopTimer
    
    ' Place crossed mines on incorrectly-
    ' placed flags
    Dim t As MinesweeperTile
    For Each t In Game.tiles
        If tile.address = t.address Then GoTo NextT
        If t.state = Flagged And Not t.mine Then
            t.Image.Picture = Sprite.useSprite("notmine")
        ElseIf t.mine Then
            t.Image.Picture = Sprite.useSprite("mine")
        End If
NextT:
    Next t
    
    ' Remove the transparent image to disable input
    Call boardForm.controls.Remove("imgTransp")
    Call Menu.setMenuPicture("oops")
    Call Stats.onLose
    
    ' Set endForm caption
    If Game.Settings.ShowStatsOnGameEnd Then
        endForm.Caption = "Lose"
        endForm.show
    End If
End Sub

Public Sub reset()
    Dim t As MinesweeperTile
    Dim boardX As Integer, boardY As Integer, numberOfMines As Integer
    Dim FormTop As Single, FormLeft As Single, diff As Difficulty
    
    ' Save the previous values
    boardX = Game.boardX
    boardY = Game.boardY
    numberOfMines = Game.numberOfMines
    diff = Game.Difficulty
    
    ' Save the previous boardForm position
    FormTop = boardForm.top
    FormLeft = boardForm.left
    
    ' Kill all the objects
    Unload boardForm
    For Each t In Game.tiles
        Set t = Nothing
    Next t
    Set Game = Nothing
    
    ' Restart
    Call Macros.start(boardX, boardY, numberOfMines, diff, FormLeft, FormTop)
End Sub

Public Sub replay()
    Dim t As MinesweeperTile
    Dim boardX As Integer, boardY As Integer, numberOfMines As Integer
    Dim FormTop As Single, FormLeft As Single, diff As Difficulty, mines As Collection
    
    ' Save the previous properties
    boardX = Game.boardX
    boardY = Game.boardY
    numberOfMines = Game.numberOfMines
    diff = Game.Difficulty
    Set mines = Game.mines
    
    ' Save the previous boardForm position
    FormTop = boardForm.top
    FormLeft = boardForm.left
    
    ' Kill all the objects
    Unload boardForm
    For Each t In Game.tiles
        Set t = Nothing
    Next t
    Set Game = Nothing
    
    ' Restart
    Call Macros.start(boardX, boardY, numberOfMines, diff, FormLeft, FormTop, mines)
End Sub

Public Sub uncoverAll()
    ' TEST: used for testing
    Dim tile As MinesweeperTile
    For Each tile In Game.tiles
        If tile.mine Then
            tile.Image.Picture = Sprite.useSprite("mine")
        Else
            tile.Image.Picture = Sprite.useSprite(CStr(tile.Number))
        End If
        tile.state = Revealed
    Next tile
End Sub

Private Function generateMines_(firstAddress As String) As Collection
    ' TEST: USED FOR TESTING THE 3BV ALGORITHM
    ' 3BV should return 39
    Set generateMines_ = New Collection
    generateMines_.add ("A0")
    generateMines_.add ("A1")
    generateMines_.add ("A6")
    generateMines_.add ("A9")
    generateMines_.add ("B0")
    generateMines_.add ("B1")
    generateMines_.add ("B6")
    generateMines_.add ("B12")
    generateMines_.add ("C0")
    generateMines_.add ("C5")
    
    generateMines_.add ("C8")
    generateMines_.add ("E12")
    generateMines_.add ("E13")
    generateMines_.add ("F13")
    generateMines_.add ("G2")
    generateMines_.add ("G6")
    generateMines_.add ("G7")
    generateMines_.add ("H5")
    generateMines_.add ("H11")
    generateMines_.add ("I5")
    
    generateMines_.add ("I6")
    generateMines_.add ("I7")
    generateMines_.add ("I13")
    generateMines_.add ("J12")
    generateMines_.add ("J5")
    generateMines_.add ("J6")
    generateMines_.add ("J7")
    generateMines_.add ("K4")
    generateMines_.add ("K6")
    
    generateMines_.add ("L2")
    generateMines_.add ("L3")
    generateMines_.add ("L6")
    generateMines_.add ("L11")
    generateMines_.add ("M2")
    generateMines_.add ("M12")
    generateMines_.add ("M15")
    generateMines_.add ("N0")
    generateMines_.add ("N5")
    generateMines_.add ("N9")
    generateMines_.add ("O2")
End Function
