Attribute VB_Name = "Utils"
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

Public Function addr(column As Integer, row As Integer) As String
    ' Returns a String address like "columnRow" Ex: A12
    addr = chr(column + 65) & CStr(row)
End Function

Public Function inx(address As String) As Move
    ' Returns an address as a Move type
    inx.Y = Asc(left(address, 1)) - 65
    inx.X = Int(right(address, Len(address) - 1))
End Function

Public Function alreadyInMines(mines As Collection, addr As String) As Boolean
    ' Returns true if mine address is already in Mines collection
    ' false otherwise
    Dim m As Variant
    For Each m In mines
        If addr = m Then
            alreadyInMines = True
            Exit Function
        End If
    Next m
    alreadyInMines = False
End Function

Public Function getActiveTile(X As Single, Y As Single) As MinesweeperTile
    ' Returns the currently active (hovered over) tile object.
    
    ' BTN_SIZE is the size of the tile-image,
    ' by doing int(x / BTN_SIZE) you get an int that is the index of the tile
    
    Dim tempX As Integer, tempY As Integer, limitX As Integer, limitY As Integer
    
    limitX = Game.boardX - 1
    limitY = Game.boardY - 1
    
    tempX = Int(X / BTN_SIZE)
    If tempX < 0 Then tempX = 0
    If tempX > limitX Then tempX = limitX
    
    tempY = Int(Y / BTN_SIZE)
    If tempY < 0 Then tempY = 0
    If tempY > limitY Then tempY = limitY
    
    Set getActiveTile = Game.tiles(addr(tempY, tempX))
End Function

Public Function isZero(tile As MinesweeperTile) As Boolean
    If tile.Number = 0 Then
        isZero = True
    Else
        isZero = False
    End If
End Function

Public Function isLoaded(name As String) As Boolean
    ' Returns true if form with name "name" is loaded
    
    Dim i As Integer
    For i = 0 To VBA.UserForms.Count - 1
        isLoaded = UserForms(i).name = name
        If isLoaded Then Exit Function
    Next i
    isLoaded = False
End Function

Public Function getNumeric(str As String) As Variant
    Dim numericStr As String
    Dim i As Integer
    For i = 1 To Len(str)
        If IsNumeric(Mid(str, i, 1)) Then numericStr = numericStr & Mid(str, i, 1)
    Next i
    If numericStr = "" Then getNumeric = Empty: Exit Function
    getNumeric = Int(numericStr)
End Function

Public Function testWin() As Boolean
    Dim revealedTiles As Integer
    Dim t As MinesweeperTile
    For Each t In Game.tiles
        If t.state = Revealed Then revealedTiles = revealedTiles + 1
    Next t
    If revealedTiles = TILES_TO_REVEAL Then testWin = True
End Function

Public Sub cleanUp()
    If Time.timer Then Time.stopTimer
    Application.Cursor = xlDefault
    Data.writeValue "lastFormTop", boardForm.top
    Data.writeValue "lastFormLeft", boardForm.left
End Sub

Private Sub recoverView()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""RIBBON"", TRUE)"
    Application.DisplayFormulaBar = True
    Application.ActiveWindow.DisplayHeadings = True
    Application.ActiveWindow.DisplayGridlines = True
    Application.ActiveWindow.DisplayWorkbookTabs = True
End Sub

Public Function getDiffName(diff As Difficulty) As String
    Dim arr As Variant
    arr = Array("Beginner", "Intermediate", "Expert", "Custom")
    getDiffName = arr(diff)
End Function

' FIX:? I think boards render differently on different versions of Excel
Public Sub sizeForm(frm As Object)
    Dim Y_OFFSET As Integer
    Dim X_OFFSET As Integer
    
    If Application.Version >= 16 Then
        Y_OFFSET = 30
        X_OFFSET = 12
    Else
        Y_OFFSET = 20
        X_OFFSET = 4
    End If
    
    Const PADDING As Integer = 6
    
    Dim ctl As Control, rmCtrl As Control, bmCtrl As Control
    Dim rightMost As Single, bottomMost As Single

    ' Get right most and bottom most controls-
    ' in form
    For Each ctl In frm.Controls
        If ctl.Visible Then
            With ctl
                If .left + .width > rightMost Then: rightMost = .left + .width
                If .top + .height > bottomMost Then: bottomMost = .top + .height
            End With
        End If
    Next ctl
    
    ' Set width and height according to right most and bottom most
    frm.width = rightMost + X_OFFSET + PADDING
    frm.height = bottomMost + Y_OFFSET + PADDING
End Sub

Public Function isSheet(sheet As String) As Boolean
    ' Returns true if sheet with name "sheet" exists
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.name = sheet Then isSheet = True: Exit Function
    Next sh
End Function
