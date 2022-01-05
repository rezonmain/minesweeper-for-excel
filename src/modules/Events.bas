Attribute VB_Name = "Events"
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

Dim activeTile As MinesweeperTile, oldTile As MinesweeperTile

Public Function handleMouseDown(Button As Integer)
    Select Case Button
    Case 1 ' LM
        Call Logic.tileLeftMouseDown(activeTile)
    Case 2 ' RM
        Call Logic.tileRightMouseDown(activeTile)
    Case 4 ' MM
        Exit Function
    End Select
End Function

Public Function handleMouseUp(Button As Integer)
    Select Case Button
    Case 1 ' LM
        Call Logic.tileLeftMouseUp(activeTile)
    Case 2 ' RM
        Call Logic.tileRightMouseUp(activeTile)
    Case 4 ' MM
        ' DEBUG
        Debug.Print activeTile.address & ": State:" & activeTile.state, "Number:" & activeTile.Number, "Mine:" & activeTile.mine
    End Select
End Function

Public Function handleMouseMove(X As Single, Y As Single, Button As Integer)
    Set oldTile = activeTile
    Set activeTile = Utils.getActiveTile(X, Y)
    Call Logic.tileMouseDownChange(activeTile, oldTile)
End Function
