Attribute VB_Name = "Vars"
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

Global Const DATA_SHEET As String = "_minesweeperdata_"
Global Const MAX_BOARD_SIZE As Integer = 32
Global Const MIN_BOARD_SIZE As Integer = 9

Global BTN_SIZE As Integer
Global TILES_TO_REVEAL As Integer
Global FormLeft As Single
Global FormTop As Single

Global Ev As Collection
Global Game As MinesweeperGame
Global Moves() As Move

Type Move
    X As Integer
    Y As Integer
End Type

Public Sub setVars(game_ As MinesweeperGame)
    ' Set global vars use throughout the modules
    
    BTN_SIZE = game_.Settings.TileSize
    TILES_TO_REVEAL = game_.boardX * game_.boardY - game_.numberOfMines
    Call initMoves
    
    ' Global game object
    Set Game = game_
    
    Set Ev = New Collection
    
    ' Prevent spinning cursor from showing
    Application.Cursor = xlNorthwestArrow
End Sub

Private Sub initMoves()
    ' Initializes the Moves array -
    ' used for doing adjacent tile operations.
    
    ReDim Moves(0 To 7)
    Dim m As Move
    m.X = 1: m.Y = 0: Moves(0) = m ' R
    m.X = -1: m.Y = 0: Moves(1) = m ' L
    m.X = 0: m.Y = -1: Moves(2) = m ' U
    m.X = 0: m.Y = 1: Moves(3) = m ' D
    m.X = 1: m.Y = -1: Moves(4) = m ' R-U
    m.X = 1: m.Y = 1: Moves(5) = m ' R-D
    m.X = -1: m.Y = -1: Moves(6) = m ' L-U
    m.X = -1: m.Y = 1: Moves(7) = m ' L-D
End Sub
