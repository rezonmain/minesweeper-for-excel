Attribute VB_Name = "Chording"
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

Public Sub startChording(tile As MinesweeperTile)
    ' Visually show/set the adjacent tiles
    Call Adjacent.setAdjacentTiles(tile)
    Stats.ichords
End Sub

Public Sub endChording(tile As MinesweeperTile)
    Dim adj As Collection
    Dim t As MinesweeperTile
    Dim effectiveChord As Boolean
    Set adj = Adjacent.getAdjacentTiles(tile.address)
    adj.add tile
    ' Reveal tiles if tile.number is the same as the number of adjacent flags
    If ((tile.state = Revealed) And (tile.Number = Adjacent.getNumberOfAdjacentFlags(tile))) Then
        For Each t In adj
        
            ' If non flagged tile has a mine end game
            If t.mine And Not t.state = Flagged Then
                t.Image.Picture = Sprite.useSprite("end")
                t.state = Revealed
                Call Game.end_(0, tile)
                Exit Sub
            End If
            
            ' Otherwise reveal the tiles
            If t.state = Hidden Then
                effectiveChord = True
                t.state = Revealed
                Call Adjacent.floodFill(t, Game.tiles, isZero(t))
                If Utils.testWin Then Call Game.end_(1, tile): Stats.ichords (True): Exit Sub
            End If
        Next t
        
    Else
        ' Otherwise just complete the "animation"
        Call Adjacent.unsetAdjacentTiles(tile)
    End If
    If effectiveChord Then Stats.ichords (True): effectiveChord = False
End Sub

Public Sub moveChording(active As MinesweeperTile, old As MinesweeperTile)
    Call Adjacent.unsetAdjacentTiles(old)
    Call Adjacent.setAdjacentTiles(active)
End Sub
