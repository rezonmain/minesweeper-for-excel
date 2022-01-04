Attribute VB_Name = "Macros"
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

Public Sub start(boardX As Integer, boardY As Integer, numberOfMines As Integer, diff As Difficulty, _
                Optional FormLeft As Variant, Optional FormTop As Variant, Optional mines As Variant)
                
    ' Initializes the game object
    Dim game_ As MinesweeperGame
    Set game_ = New MinesweeperGame
    
    game_.boardX = boardX
    game_.boardY = boardY
    game_.numberOfMines = numberOfMines
    game_.Difficulty = diff
    game_.FormLeft = FormLeft
    game_.FormTop = FormTop
    
    ' Mines is provided when board is replayed
    If Not IsMissing(mines) Then
        Set game_.mines = mines
        game_.IsReplay = True
    End If
    
    Call init.init(game_)
End Sub

Public Sub beginner_()
    Call start(9, 9, 10, beginner, Data.getValue("lastFormLeft"), Data.getValue("lastFormTop"))
End Sub

Public Sub intermediate_()
    Call start(16, 16, 40, intermediate, Data.getValue("lastFormLeft"), Data.getValue("lastFormTop"))
End Sub

Public Sub expert_()
    Call start(30, 16, 99, expert, Data.getValue("lastFormLeft"), Data.getValue("lastFormTop"))
End Sub

Public Sub openMenuForm(ribbonButton As IRibbonControl)
    menuForm.show vbModeless
End Sub

Public Sub openCustomForm()
    customForm.show
End Sub

Public Sub openStatsForm()
    statsForm.show
End Sub

Public Sub openSettingsForm()
    settingsForm.show
End Sub

Public Sub exitWorkbook()
    ThisWorkbook.Save
    Application.Quit
End Sub
