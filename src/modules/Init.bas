Attribute VB_Name = "Init"
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

Public Sub init(game_ As MinesweeperGame)
    Vars.setVars game_
    clearNewRecords
    configureBoard
    addBorders
    addTiles
    addImage
    addMenuBtn
    addDigits
    Utils.sizeForm boardForm
    boardForm.Show
End Sub

Private Sub configureBoard()
    ' Callers: Init.init
    ' Configures form size to accomodate board size
    
    Dim initialY As Integer, initialX As Integer
    Dim borderOffsetY As Integer, borderOffsetX As Integer
    
    If IsEmpty(Game.FormTop) Then
        ' Use system default positioning
        boardForm.StartUpPosition = 2
    Else
        boardForm.StartUpPosition = 0
        boardForm.top = Game.FormTop
        boardForm.left = Game.FormLeft
    End If
    
    Select Case Game.Settings.theme
    Case Default
        boardForm.BackColor = &HE0E0E0
    Case Dark
        boardForm.BackColor = &H404040
    End Select
    
    boardForm.Caption = "Minesweeper"
    
    If Game.Difficulty = custom Then boardForm.Caption = boardForm.Caption & " - Custom"
    
    If Game.IsReplay Then boardForm.Caption = boardForm.Caption & " - Replay"
End Sub

Private Sub addBorders()
    ' Adds the "sunken" effect border around the tiles for default theme
    ' Dark theme is just flat
    Dim img As Image
    
    Dim left_ As Long, height_ As Long, width_ As Long, top_ As Long
    top_ = 64
    left_ = 6
    height_ = Game.boardY * BTN_SIZE + 6
    width_ = Game.boardX * BTN_SIZE + 6
    Set img = boardForm.Controls.add("Forms.Image.1", "imgBb")
    With img
        .top = top_
        .left = left_
        .height = height_
        .width = width_
        .BorderStyle = fmBorderStyleNone
        .Picture = Sprite.useSprite("border")
        .PictureSizeMode = fmPictureSizeModeStretch
    End With
    
    ' Use frame for the heading border
    With boardForm.frmHeading
        .top = 8
        .left = left_
        .height = 50
        .width = width_
        .SpecialEffect = fmSpecialEffectSunken
        .Caption = ""
        .BackColor = &HE0E0E0
        If Game.Settings.theme = Dark Then
            .SpecialEffect = fmSpecialEffectFlat
            .BackColor = &H404040
        End If
        
    End With
End Sub

Private Sub addTiles()
    ' Callers: Init.init
    ' Creates tile grid and adds it to form-
    ' initializes minesweeperTiles objects
    
    Dim tile As MinesweeperTile
    Dim top_ As Long, left_ As Long
    Dim i As Integer, j As Integer
    
    ' Tiles starting position in form:
    top_ = 67
    left_ = 9
    
    For i = 0 To Game.boardY - 1
        For j = 0 To Game.boardX - 1
            Set tile = New MinesweeperTile: tile.address = addr(i, j)
            Set tile.Image = boardForm.Controls.add("Forms.Image.1", addr(i, j))
            With tile.Image
                .height = BTN_SIZE
                .width = BTN_SIZE
                .left = left_
                .top = top_
                
                ' increment left_ position for next tile
                left_ = left_ + BTN_SIZE
                
                .PictureSizeMode = fmPictureSizeModeStretch
                .Picture = Sprite.useSprite("t")
                .BorderStyle = fmBorderStyleNone
            End With
            Game.tiles.add tile, tile.address
            
        Next j
        top_ = top_ + BTN_SIZE
        left_ = 9
    Next i
End Sub

Private Sub addImage()
    ' Callers: Init.init
    ' Adds a transparent img on top of tiles for-
    ' mouse position tracking, and mouse event
    ' Sets imgEvents object
    
    Dim imgEvents As ImageEvents
    Dim img As MSForms.Image
    Set img = boardForm.Controls.add("Forms.Image.1", "imgTransp")
    With img
        .top = 67
        .left = 8
        .height = Game.boardY * BTN_SIZE
        .width = Game.boardX * BTN_SIZE
        .BorderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
    End With
    Set imgEvents = New ImageEvents
    Set imgEvents.img = img
    Vars.Ev.add imgEvents
End Sub

Private Sub addMenuBtn()
    ' Callers: Init.init
    ' Adds menu button to form.
    ' Sets btnEvents object
    
    Dim img As MSForms.Image
    Dim btn As MSForms.CommandButton
    Dim lbl As MSForms.Label
    Dim btnEvents As ButtonEvents
    Dim menuBtnSize As Single
    menuBtnSize = Vars.BTN_SIZE * 1.7
    
    Set img = boardForm.frmHeading.Controls.add("Forms.Image.1", "btnMenu")
    With img
        .top = (boardForm.frmHeading.height / 2) - (menuBtnSize / 2)
        .left = (boardForm.frmHeading.width / 2) - (menuBtnSize / 2) - 2.5
        .width = menuBtnSize
        .height = menuBtnSize
        .Picture = Sprite.useSprite("ok")
        .BorderStyle = fmBorderStyleNone
        .PictureSizeMode = fmPictureSizeModeStretch
        
    End With
    Set btnEvents = New ButtonEvents
    Set btnEvents.btn = img
    Vars.Ev.add btnEvents
    
    ' Create a non-visible button to steal focus from menubtn
    Set btn = boardForm.Controls.add("Forms.CommandButton.1", "btnIStealFocus")
    With btn
        .top = 0
        .left = 0
        .width = 0
        .height = 0
        .TabIndex = 0
        .SetFocus
    End With
End Sub

Private Sub addDigits()
    ' Adds 7 segment counters on board header
    
    Dim img As MSForms.Image
    Dim fra As MSForms.Frame
    Dim se As Integer
    Dim width_ As Single, height_ As Single, i As Integer
    Dim left_ As Single
    height_ = Vars.BTN_SIZE * 1.7
    width_ = height_ / 1.7
    left_ = 2
    Select Case Game.Settings.theme
        Case Default
            se = 2
        Case Dark
            se = 0
    End Select
    Set fra = boardForm.frmHeading.Controls.add("Forms.Frame.1")
    With fra
        .top = (boardForm.frmHeading.height / 2) - (height_ / 2) - 1
        .left = left_ - 1
        .height = height_ + 3
        .width = width_ * 3 + 3
        .Caption = ""
        .SpecialEffect = se
        left_ = 0
        For i = 1 To 3
            Set img = .Controls.add("Forms.Image.1", "dig" & CStr(i))
            With img
                .top = 0
                .left = left_
                .height = height_
                .width = width_
                .Picture = Sprite.useSprite("n0")
                .PictureSizeMode = fmPictureSizeModeStretch
                .BorderStyle = fmBorderStyleNone
            End With
            left_ = left_ + width_
        Next i
    End With
    
    left_ = Game.boardX * BTN_SIZE - (width_ * 3) - 1
    Set fra = boardForm.frmHeading.Controls.add("Forms.Frame.1")
    With fra
        .top = (boardForm.frmHeading.height / 2) - (height_ / 2) - 1
        .left = left_ + 1
        .height = height_ + 3
        .width = width_ * 3 + 3
        .Caption = ""
        .SpecialEffect = se
        left_ = width_ * 2
        For i = 1 To 3
            Set img = .Controls.add("Forms.Image.1", "dig" & CStr(i + 3))
                With img
                .top = 0
                .left = left_
                .height = height_
                .width = width_
                .Picture = Sprite.useSprite("n0")
                .PictureSizeMode = fmPictureSizeModeStretch
                .BorderStyle = fmBorderStyleNone
            End With
            left_ = left_ - width_
        Next i
    End With
    Call Digits.setFlagCounter(Game.numberOfMines)
End Sub

Private Sub clearNewRecords()
    Stats.getNewRecords
End Sub
