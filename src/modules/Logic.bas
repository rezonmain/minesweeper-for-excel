Attribute VB_Name = "Logic"
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

Dim leftDown As Boolean, rightDown As Boolean, chord As Boolean

Public Sub tileLeftMouseDown(tile As MinesweeperTile)

    ' Set the menu button picture to the wow face
    Call Menu.setMenuPicture("wow")
    
    leftDown = True
    
    ' Start chording if both MB are pressed
    If leftDown And rightDown Then
        chord = True
        Call Chording.startChording(tile)
        Exit Sub
    End If
    
    ' Handle left mouse-down according
    ' to current tile state
    Select Case tile.state
    Case Hidden
        tile.Image.Picture = Sprite.useSprite("0")
    Case Revealed
        Exit Sub
    Case Flagged
        Exit Sub
    End Select
End Sub

Public Sub tileLeftMouseUp(tile As MinesweeperTile)

    ' Complete the menu button 'animation'
    Call Menu.setMenuPicture("ok")
    
    ' Check if its player's first click
    If Not Game.FirstClick Then
        ' Generate mines and set numbers
        If Game.IsReplay Then
            Call Board.setTilesProperties(tile.address, Game.mines)
        Else
            Call Board.setTilesProperties(tile.address)
        End If
        
        Game.Stats.BBBV = Stats.get3BV
        Time.startTimer
        
        Game.FirstClick = True
    End If
    
    ' End chording
    If chord Then
        leftDown = False
        rightDown = False
        chord = False
        Call Chording.endChording(tile)
        Exit Sub
    End If
    
    leftDown = False
    
    ' Handle left mouse-up according to-
    ' active tile state
    Select Case tile.state
    Case Hidden
        ' If player hits mine
        Stats.ilefts (True)
        If tile.mine Then
            tile.Image.Picture = Sprite.useSprite("end")
            tile.state = Revealed
            Call Game.end_(0, tile)
        Else
            tile.state = Revealed
            Call Adjacent.floodFill(tile, Game.tiles, isZero(tile))
            If Utils.testWin Then Call Game.end_(1, tile): Exit Sub
        End If
    Case Revealed
        Exit Sub
    Case Flagged
        Exit Sub
    End Select
End Sub

Public Sub tileRightMouseUp(tile As MinesweeperTile)
    Call Menu.setMenuPicture("ok")
    
    If chord Then
        leftDown = False
        rightDown = False
        chord = False
        Call Chording.endChording(tile)
    End If
    rightDown = False
End Sub

Public Sub tileRightMouseDown(tile As MinesweeperTile)
    rightDown = True
    
    If leftDown And rightDown Then
        chord = True
        Call Chording.startChording(tile)
        Exit Sub
    End If
    
    Select Case tile.state
    Case Hidden
        tile.state = Flagged
        tile.Image.Picture = Sprite.useSprite("flag")
        Game.NumberOfFlags = Game.NumberOfFlags + 1
        Call Digits.setFlagCounter(Game.numberOfMines - Game.NumberOfFlags)
        If tile.mine Then Stats.irights (True)
    Case Revealed
        Exit Sub
    Case Flagged
        tile.state = Hidden
        tile.Image.Picture = Sprite.useSprite("t")
        Game.NumberOfFlags = Game.NumberOfFlags - 1
        Call Digits.setFlagCounter(Game.numberOfMines - Game.NumberOfFlags)
    End Select
    Stats.irights
End Sub

Public Sub tileMouseDownChange(active As MinesweeperTile, old As MinesweeperTile)
    ' Handle logic when mouse is dragged
    
    If old Is Nothing Then Exit Sub
    If active.address = old.address Then Exit Sub
    
    If chord Then
        Call Chording.moveChording(active, old)
        Exit Sub
    End If
        
    If rightDown Then Exit Sub
    
    If leftDown Then
        If old.state = Hidden Then old.Image.Picture = Sprite.useSprite("t")
        If active.state = Hidden Then active.Image.Picture = Sprite.useSprite("0")
    End If
End Sub
