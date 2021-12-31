VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} settingsForm 
   Caption         =   "Settings"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3345
   OleObjectBlob   =   "settingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "settingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Dim ctrls As Collection

Private Sub btnApply_Click()
    Call clearBoard
    Call generateBoard(settingsForm.spnSize, settingsForm.cmbTheme.value)
    settingsForm.btnApply.Enabled = False
End Sub

Private Sub btnClearData_Click()
    Dim res As Long
    res = MsgBox(Prompt:="Are you sure you want to clear all game stats?", Buttons:=vbYesNo + vbDefaultButton2 + vbExclamation)
    Select Case res
    Case vbYes
        Data.clearStats
    Case vbNo
        Exit Sub
    End Select
End Sub

Private Sub UserForm_Initialize()
    Randomize
    With settingsForm
        .spnSize.value = getValue("TileSize")
        .cmbTheme.ColumnWidths = "0;20"
        .cmbTheme.AddItem ("0")
        .cmbTheme.List(0, 1) = "Default"
        .cmbTheme.AddItem ("1")
        .cmbTheme.List(1, 1) = "Dark"
        .cmbTheme.value = getValue("Theme")
        .xStatsOnEnd.value = getValue("ShowStatsOnGameEnd")
        .xRecordsOnReplay.value = getValue("RecordsOnReplay")
        .lblSize.Caption = CStr(.spnSize.value)
        .btnApply.Enabled = False
        
        Call generateBoard(.spnSize.value, getValue("Theme"))
        .width = 175.75 + (.spnSize.value * 9) + 6
        .btnSave.Enabled = False
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If (settingsForm.btnSave.Enabled) Then
        Dim res As Integer
        Select Case MsgBox(Prompt:="Exit without saving settings?", Buttons:=vbYesNo + vbExclamation + vbDefaultButton2, Title:="Minesweeper")
        Case vbYes
            Unload Me
        Case vbNo
            Cancel = 1
            Exit Sub
        End Select
    End If
End Sub

Private Sub btnDefault_Click()
    Dim res As Integer
    Select Case MsgBox("Are you sure you want to restore settings to their default values?", Buttons:=vbYesNo + vbExclamation + vbDefaultButton2, Title:="Confirm")
    Case vbYes
        Data.clearSettings
        clearBoard
        With settingsForm
            .spnSize.value = getValue("TileSize")
            .xStatsOnEnd.value = getValue("ShowStatsOnGameEnd")
            .cmbTheme.value = getValue("Theme")
            .xRecordsOnReplay = getValue("RecordsOnReplay")
            .lblSize.Caption = CStr(.spnSize.value)
            Call generateBoard(.spnSize.value, .cmbTheme.value)
            .width = 175.75 + (.spnSize.value * 9) + 6
            .btnSave.Enabled = False
        End With
    Case vbNo
        Exit Sub
    End Select
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Dim statsOnEnd As Boolean, theme As theme, boardSize As Integer
    Dim recordsOnReplay As Boolean
    With settingsForm
        statsOnEnd = .xStatsOnEnd.value
        theme = .cmbTheme.value
        boardSize = .spnSize.value
        recordsOnReplay = .xRecordsOnReplay.value
        .btnSave.Enabled = False
    End With
    Call Data.writeValue("Theme", theme)
    Call Data.writeValue("ShowStatsOnGameEnd", statsOnEnd)
    Call Data.writeValue("TileSize", boardSize)
    Call Data.writeValue("RecordsOnReplay", recordsOnReplay)
End Sub

Private Sub cmbTheme_Change()
    settingsForm.btnSave.Enabled = True
    settingsForm.btnApply.Enabled = True
    settingsForm.btnApply.SetFocus
End Sub

Private Sub spnSize_SpinDown()
    With settingsForm
        .lblSize.Caption = CStr(.spnSize.value)
        Call clearBoard
        Call generateBoard(.spnSize.value, .cmbTheme.value)
        .btnSave.Enabled = True
    End With
End Sub

Private Sub spnSize_SpinUp()
    With settingsForm
        .lblSize.Caption = CStr(.spnSize.value)
        Call clearBoard
        Call generateBoard(.spnSize.value, .cmbTheme.value)
        .btnSave.Enabled = True
    End With
End Sub

Private Sub generateBoard(boardSize As Integer, theme As theme)
    Application.ScreenUpdating = False
    Set ctrls = New Collection
    Dim i As Integer, j As Integer, heightLimit As Single
    Dim boardHeight As Single
    Call addBorder(boardSize, theme)
    Call addTiles(boardSize, theme)
    Call addMenuBtn(boardSize, theme)
    Call addDigits(boardSize, theme)
    
    settingsForm.width = 175.75 + (boardSize * 9) + 6
    heightLimit = 280.5
    boardHeight = 92 + (boardSize * 9)
    If boardHeight > heightLimit Then
        settingsForm.height = boardHeight + 3
    Else
        settingsForm.height = heightLimit
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub addTiles(boardSize As Integer, theme As theme)
    Dim top As Single, left As Single
    Dim i As Integer, j As Integer
    Dim rs As Collection, r As Single
    Dim img As Image
    
    Set rs = New Collection
    left = settingsForm.frmGeneral.width + 12 + 2
    top = 67
    For i = 0 To 8
        For j = 0 To 8
            Set img = settingsForm.controls.add("Forms.Image.1", addr(i, j))
            With img
                .top = top
                .left = left
                .width = boardSize
                .height = boardSize
                r = rnd() * 10
                If r <= 3 And rs.Count < 9 Then
                    r = rnd() * 9
                    If r >= 5 Then
                        .Picture = Sprite.settingsSprite(rs.Count, theme)
                        rs.add r
                    ElseIf r >= 3 And r < 5 Then
                        .Picture = Sprite.settingsSprite("mine", theme)
                    ElseIf r < 3 Then
                        .Picture = Sprite.settingsSprite("flag", theme)
                    End If
                Else
                    .Picture = Sprite.settingsSprite("t", theme)
                End If
                .PictureSizeMode = fmPictureSizeModeStretch
                .BorderStyle = fmBorderStyleNone
            End With
            ctrls.add img, img.name
            left = left + boardSize
        Next j
        left = settingsForm.frmGeneral.width + 12 + 2
        top = top + boardSize
    Next i
End Sub

Private Sub addBorder(boardSize As Integer, theme As theme)
        ' Adds the "sunken" effect border around the tiles
    Dim img As Image
    
    Dim left_ As Long, height_ As Long, width_ As Long, top_ As Long
    top_ = 64
    left_ = settingsForm.frmGeneral.width + 12
    height_ = 9 * boardSize + 6
    width_ = 9 * boardSize + 6
    Set img = settingsForm.controls.add("Forms.Image.1", "border")
    With img
        .top = top_
        .left = left_
        .height = height_
        .width = width_
        .BorderStyle = fmBorderStyleNone
        .Picture = Sprite.settingsSprite("border", theme)
        .PictureSizeMode = fmPictureSizeModeStretch
    End With
    ctrls.add img, img.name
    
    ' Use frame for the heading border
    With settingsForm.frmHeading
        .top = 6
        .left = left_
        .height = 50
        .width = width_
        .SpecialEffect = fmSpecialEffectSunken
        .Caption = ""
        Select Case theme
        Case Default
            .BackColor = &HE0E0E0
        Case Dark
            .BackColor = &H404040
            .SpecialEffect = fmSpecialEffectFlat
        End Select
    End With
End Sub

Private Sub addMenuBtn(boardSize As Integer, theme As theme)
    ' Callers: Init.init
    ' Adds menu button to form.
    ' Sets btnEvents object
    
    Dim img As MSForms.Image
    Dim btn As MSForms.CommandButton
    Dim lbl As MSForms.Label
    Dim btnEvents As ButtonEvents
    Dim menuBtnSize As Single
    menuBtnSize = boardSize * 1.5
    
    Set img = settingsForm.frmHeading.controls.add("Forms.Image.1", "btnMenu")
    With img
        .top = (settingsForm.frmHeading.height / 2) - (menuBtnSize / 2)
        .left = (settingsForm.frmHeading.width / 2) - (menuBtnSize / 2) - 2.5
        .width = menuBtnSize
        .height = menuBtnSize
        .Picture = Sprite.settingsSprite("ok", theme)
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectRaised
        .PictureSizeMode = fmPictureSizeModeStretch
    End With
    ctrls.add img, img.name

End Sub

Private Sub addDigits(boardSize As Integer, theme As theme)
    ' Adds 7 segment counters on board header
    
    Dim img As MSForms.Image
    Dim width_ As Single, height_ As Single, i As Integer
    Dim left_ As Single
    height_ = boardSize * 1.5
    width_ = height_ / 1.5
    left_ = 2
    
    For i = 1 To 3
        Set img = settingsForm.frmHeading.controls.add("Forms.Image.1", "dig" & CStr(i))
        With img
            .top = (settingsForm.frmHeading.height / 2) - (height_ / 2)
            .left = left_
            .height = height_
            .width = width_
            .Picture = Sprite.settingsSprite("n" & i, theme)
            .PictureSizeMode = fmPictureSizeModeStretch
            .BorderStyle = fmBorderStyleNone
        End With
        left_ = left_ + width_
        ctrls.add img, img.name
    Next i
    
    left_ = 9 * boardSize - width_ + 1
    For i = 1 To 3
        Set img = settingsForm.frmHeading.controls.add("Forms.Image.1", "dig" & CStr(i + 3))
            With img
            .top = (settingsForm.frmHeading.height / 2) - (height_ / 2)
            .left = left_
            .height = height_
            .width = width_
            .Picture = Sprite.settingsSprite("n" & 3 - i + 1, theme)
            .PictureSizeMode = fmPictureSizeModeStretch
            .BorderStyle = fmBorderStyleNone
        End With
        left_ = left_ - width_
        ctrls.add img, img.name
    Next i
End Sub

Private Sub clearBoard()
    Dim img As Variant
    For Each img In ctrls
        settingsForm.controls.Remove img.name
        Set img = Nothing
    Next img
End Sub

Private Sub xRecordsOnReplay_Click()
    settingsForm.btnSave.Enabled = True
End Sub

Private Sub xStatsOnEnd_Click()
    settingsForm.btnSave.Enabled = True
End Sub
