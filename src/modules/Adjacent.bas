Attribute VB_Name = "Adjacent"
' Copyright 2022 Alejandro D.
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

Public Sub floodFill(tile_ As MinesweeperTile, ByRef tiles As Collection, zero As Boolean)
    ' Recursively uncovers the minesweeper tiles-
    ' that have a 0 as a number. The floodfill is-
    ' limited by tiles that have a number > 0-
    ' or a mine or already revealed tile.
    
    Dim X As Integer, Y As Integer, i As Integer
    Dim tile As MinesweeperTile
    
    ' Get the current tile address as an integer index to use with the numbers array
    Y = inx(tile_.address).Y
    X = inx(tile_.address).X
    
    ' Set the tile picture to its corresponding number-picture
    If Not tile_.Image Is Nothing Then
        tile_.Image.Picture = Sprite.useSprite(CStr(tile_.Number))
    End If
    
    ' Set it to revealed so as not to call it again
    tile_.state = Revealed
    
    ' Uncover adjacent tiles only if current tile has zero as number
    If zero Then
        ' Cycle through the moves array
        For i = 0 To 7
            ' Check if adjacent address will not be out of bounds
            If (X + Moves(i).X >= 0 And X + Moves(i).X < Game.boardX) And (Y + Moves(i).Y >= 0 And Y + Moves(i).Y < Game.boardY) Then
                ' Set adjacent tile
                Set tile = tiles(addr(Y + Moves(i).Y, X + Moves(i).X))
                
                ' Skip recursive call if tile has already been revealed
                If Not ((tile.state = Revealed) Or (tile.mine) Or (tile.state = Flagged)) Then
                    Call floodFill(tile, tiles, isZero(tile))
                End If
            End If
        Next i
    End If
End Sub

Public Sub setAdjacentTiles(tile As MinesweeperTile)
    ' Visually set adjacent tiles
    
    Dim adj As Collection
    Dim t As MinesweeperTile
    Set adj = getAdjacentTiles(tile.address)
    adj.add tile
    For Each t In adj
        If t.state = Hidden Then
            t.Image.Picture = Sprite.useSprite("0")
        End If
    Next t
End Sub

Public Sub unsetAdjacentTiles(tile As MinesweeperTile)
    ' Visually unset adjacent tiles.
    
    Dim adj As Collection
    Dim t As MinesweeperTile
    Set adj = getAdjacentTiles(tile.address)
    adj.add tile
    For Each t In adj
        If t.state = Hidden Then
            t.Image.Picture = Sprite.useSprite("t")
        End If
    Next t
End Sub

Public Function getAdjacentTiles(address As String) As Collection
    ' Returns a collection with the adjacent tiles of-
    ' the provided tile object address.
    
    Dim Y As Integer, X As Integer, i As Integer
    Set getAdjacentTiles = New Collection
    Y = inx(address).Y
    X = inx(address).X
    For i = 0 To 7
        If ((X + Moves(i).X < Game.boardX And X + Moves(i).X >= 0) And (Y + Moves(i).Y < Game.boardY And Y + Moves(i).Y >= 0)) Then
            getAdjacentTiles.add Game.tiles(addr(Y + Moves(i).Y, X + Moves(i).X))
        End If
    Next i
End Function

Public Function getNumberOfAdjacentFlags(tile As MinesweeperTile) As Integer
    ' Returns number of adjacent flagged tiles of provided tile object.
    
    Dim Y As Integer, X As Integer, i As Integer
    getNumberOfAdjacentFlags = 0
    Y = inx(tile.address).Y
    X = inx(tile.address).X
    For i = 0 To 7
        If ((X + Moves(i).X < Game.boardX And X + Moves(i).X >= 0) And (Y + Moves(i).Y < Game.boardY And Y + Moves(i).Y >= 0)) Then
            If Game.tiles(addr(Y + Moves(i).Y, X + Moves(i).X)).state = Flagged Then
                getNumberOfAdjacentFlags = getNumberOfAdjacentFlags + 1
            End If
        End If
    Next i
End Function
