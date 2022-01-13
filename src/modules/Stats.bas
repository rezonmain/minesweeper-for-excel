Attribute VB_Name = "Stats"
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

Public Function setRecords()
    If Game.IsReplay And (Not Game.Settings.recordsOnReplay) Then Exit Function
    ' Returns which records were set
    Set setRecords = New Dictionary
    Select Case Game.Difficulty
    Case 0
        If Game.Stats.Time < Int(Data.getValue("beginnerTime")) Then
            Call Data.writeValue("beginnerTime", Game.Stats.Time, True)
            Call Data.writeValue("lastGameTime", Game.Stats.Time, True)
        End If
        
        If Game.Stats.BBBVS > Data.getValue("beginner3BV/s") Then
            Call writeValue("beginner3BV/s", Game.Stats.BBBVS, True)
            Call Data.writeValue("lastGame3BV/s", Game.Stats.Time, True)
        End If
    Case 1
        If Game.Stats.Time < Int(Data.getValue("intermediateTime")) Then
            Call Data.writeValue("intermediateTime", Game.Stats.Time, True)
            Call Data.writeValue("lastGameTime", Game.Stats.Time, True)
        End If
        
        If Game.Stats.BBBVS > Data.getValue("intermediate3BV/s") Then
            Call writeValue("intermediate3BV/s", Game.Stats.BBBVS, True)
            Call Data.writeValue("lastGame3BV/s", Game.Stats.Time, True)
        End If
    Case 2
        If Game.Stats.Time < Int(Data.getValue("expertTime")) Then
            Call Data.writeValue("expertTime", Game.Stats.Time, True)
            Call Data.writeValue("lastGameTime", Game.Stats.Time, True)
        End If
        
        If Game.Stats.BBBVS > Data.getValue("expert3BV/s") Then
            Call writeValue("expert3BV/s", Game.Stats.BBBVS, True)
            Call Data.writeValue("lastGame3BV/s", Game.Stats.Time, True)
        End If
    End Select
End Function

Public Function getNewRecords() As Dictionary
    Set getNewRecords = New Dictionary
    Dim cell As Range
    With ThisWorkbook.Sheets(DATA_SHEET)
        For Each cell In .Range("F1:F" & CStr(findRow("*")))
            If cell.value = "New" Then
                getNewRecords.add .Range("A" & cell.row).value, True
            End If
            cell.value = ""
        Next cell
    End With
End Function

Public Sub onLose()
    setTimeStat
    Game.Stats.CompletedBBBV = getCompleted3BV
    If Not Game.Stats.Time = 0 Then
        Game.Stats.BBBVS = Game.Stats.BBBV / Game.Stats.Time
    Else
        Game.Stats.BBBVS = 0
    End If
    
    Game.Stats.setPropertiesDict
    incrementLose
    writeLastGameStats
End Sub

Public Sub onWin()
    setTimeStat
    Game.Stats.CompletedBBBV = getCompleted3BV
    Game.Stats.BBBVS = Game.Stats.BBBV / Game.Stats.Time
    Game.Stats.setPropertiesDict
    incrementWin
    setRecords
    writeLastGameStats
End Sub

Private Sub incrementWin()
    Call Data.writeValue("gamesWon", Data.getValue("gamesWon") + 1)
End Sub

Private Sub incrementLose()
    Call Data.writeValue("gamesLost", Data.getValue("gamesLost") + 1)
End Sub

Public Sub ichords(Optional effective As Boolean)
    If effective Then Game.Stats.EffectiveChords = Game.Stats.EffectiveChords + 1: Exit Sub
    Game.Stats.Chords = Game.Stats.Chords + 1
End Sub

Public Sub ilefts(Optional effective As Boolean)
    If effective Then Game.Stats.EffectiveLeftClicks = Game.Stats.EffectiveLeftClicks + 1
    Game.Stats.LeftClicks = Game.Stats.LeftClicks + 1
End Sub

Public Sub irights(Optional effective As Boolean)
    If effective Then Game.Stats.EffectiveRightClicks = Game.Stats.EffectiveRightClicks + 1: Exit Sub
    Game.Stats.RightClicks = Game.Stats.RightClicks + 1
End Sub

Private Sub setTimeStat()
    Game.Stats.Time = Time.getStatSeconds
End Sub

Public Function get3BV() As Integer
    ' Returns the 3BV of the current board
    
    Dim tiles As Collection
    Dim t As MinesweeperTile
    Dim tNew As MinesweeperTile
    Set tiles = New Collection
    
    ' Copy the Game.tiles collection to a new collection
    For Each t In Game.tiles
        Set tNew = New MinesweeperTile
        With tNew
            .address = t.address
            .mine = t.mine
            .Number = t.Number
            .state = t.state
        End With
        tiles.add tNew, tNew.address
    Next t
    
    ' Reveal tiles using floodFill
    For Each t In tiles
        If t.Number = 0 And (Not t.mine) Then
            If t.state = Revealed Then GoTo Continue
            get3BV = get3BV + 1
            Call Adjacent.floodFill(t, tiles, isZero(t))
        End If
Continue:
    Next t
    ' Count the remaining tiles that don't get revealed with floodFill
    For Each t In tiles
        If t.state = Hidden And (Not t.mine) Then get3BV = get3BV + 1
    Next t
End Function

Public Function getCompleted3BV() As Integer
    getCompleted3BV = Game.Stats.BBBV - get3BV
End Function

Public Sub writeLastGameStats()
    Dim stat As Variant
    For Each stat In Game.Stats.Properties
        writeValue "lastGame" & stat, Game.Stats.Properties(stat)
    Next stat
End Sub

