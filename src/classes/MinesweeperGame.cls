VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MinesweeperGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Enum Difficulty
    beginner = 0
    intermediate = 1
    expert = 2
    custom = 3
End Enum

Private NumberOfMines_ As Integer
Private BoardX_ As Integer
Private BoardY_ As Integer
Private Tiles_ As Collection
Private NumberOfFlags_ As Integer
Private FirstClick_ As Boolean
Private FormLeft_ As Variant
Private FormTop_ As Variant
Private Difficulty_ As Difficulty
Private IsReplay_ As Boolean
Private Mines_ As Collection

Public stats As MinesweeperStats
Public Settings As MinesweeperSettings

Public Property Get numberOfMines() As Integer
    numberOfMines = NumberOfMines_
End Property

Public Property Get boardX() As Integer
    boardX = BoardX_
End Property

Public Property Get boardY() As Integer
    boardY = BoardY_
End Property

Public Property Get tiles() As Collection
    Set tiles = Tiles_
End Property

Public Property Get NumberOfFlags() As Integer
    NumberOfFlags = NumberOfFlags_
End Property

Public Property Get FirstClick() As Boolean
    FirstClick = FirstClick_
End Property

Public Property Get FormLeft() As Variant
    FormLeft = FormLeft_
End Property

Public Property Get FormTop() As Variant
    FormTop = FormTop_
End Property

Public Property Get Difficulty() As Difficulty
    Difficulty = Difficulty_
End Property

Public Property Get IsReplay() As Boolean
    IsReplay = IsReplay_
End Property

Public Property Get mines() As Collection
    Set mines = Mines_
End Property

Public Property Let numberOfMines(value As Integer)
    NumberOfMines_ = value
End Property

Public Property Let boardX(value As Integer)
    BoardX_ = value
End Property

Public Property Let boardY(value As Integer)
    BoardY_ = value
End Property

Public Property Set tiles(value As Collection)
    Set Tiles_ = value
End Property

Public Property Let NumberOfFlags(value As Integer)
    NumberOfFlags_ = value
End Property

Public Property Let FirstClick(value As Boolean)
    FirstClick_ = value
End Property

Public Property Let FormLeft(value As Variant)
    FormLeft_ = value
End Property

Public Property Let FormTop(value As Variant)
    FormTop_ = value
End Property

Public Property Let Difficulty(value As Difficulty)
    Difficulty_ = value
End Property

Public Property Let IsReplay(value As Boolean)
    IsReplay_ = value
End Property

Public Property Set mines(value As Collection)
    Set Mines_ = value
End Property

Private Sub Class_Initialize()
    Set Tiles_ = New Collection
    Set stats = New MinesweeperStats
    Set Settings = New MinesweeperSettings
    Set Mines_ = New Collection
    FirstClick_ = False
    NumberOfFlags_ = 0
    IsReplay_ = False
End Sub

Public Sub end_(state As Integer, tile As MinesweeperTile)
    Select Case state
    Case 0 ' player revealed a mine
        Call Board.handleLose(tile)
    Case 1 ' player won
        Board.handleWin
    End Select
End Sub


