VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} endForm 
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4350
   OleObjectBlob   =   "endForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "endForm"
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

Private Sub populateForm()
    Dim lastGameStats As Collection
    Dim top As Single, left As Single, stat As Variant
    Dim rowSpacing As Integer, btnOffset As Integer, formOffset As Integer
    Dim nameColumnSpacing As Integer, valueColumnSpacing As Integer, fontSize As Integer
    Dim lbl As MSForms.Label
    Dim records As Dictionary
    Dim i As Integer, j As Integer
    
    fontSize = 10
    rowSpacing = fontSize + 2
    btnOffset = fontSize
    formOffset = fontSize * 3.5
    nameColumnSpacing = fontSize * 11
    valueColumnSpacing = fontSize * 5
    
    Set lastGameStats = getLastGameStats
    Set records = stats.getNewRecords
    
    top = 5
    left = 2
    
    For i = 1 To lastGameStats.Count
        stat = lastGameStats(i)
        
        ' Don't create a label for the tooltip
        For j = 0 To UBound(stat) - 1
            Set lbl = endForm.frmLastGame.controls.add("Forms.label.1")
            With lbl
                .left = left
                .top = top
                .Caption = CStr(stat(j))
                .width = 130
                .Font.Size = 10
                If j = 0 Then .ControlTipText = stat(3)
                
                ' Alternate colors
                If (i Mod 2 = 0) Then
                    .ForeColor = &H80000012
                Else
                    .ForeColor = &H80000015
                End If
                
                If records.Item(stat(4)) Then
                    .ForeColor = &HFF
                End If
            End With
            
            ' Set correct spacing
            If j = 0 Then
                left = left + nameColumnSpacing
            Else
                left = left + valueColumnSpacing
            End If
        Next j
        left = 2
        top = top + rowSpacing
    Next i
    endForm.frmLastGame.height = (lastGameStats.Count * rowSpacing) + rowSpacing
    endForm.height = btnStartGame.height + frmLastGame.height + formOffset
    btnStartGame.top = endForm.frmLastGame.height + 3
    btnExit.top = btnStartGame.top + (btnStartGame.height / 2) + 3
    btnReplay.top = btnStartGame.top
End Sub

Private Function getLastGameStats() As Collection
    Dim lastGameStatsRow As Long, endRow As Long
    Dim name As String, value As String, unit As String, toolTip As String, varName As String
    Dim cell As Range
    Set getLastGameStats = New Collection
    
    lastGameStatsRow = Data.findRow("LAST_GAME_STATS")
    endRow = Data.findRow("*")
    
    With ThisWorkbook.Sheets(DATA_SHEET)
        For Each cell In .Range("D" & lastGameStatsRow + 1 & ":D" & endRow)
            name = cell.value
            value = .Range("B" & cell.row).value
            unit = .Range("C" & cell.row).value
            toolTip = .Range("E" & cell.row).value
            varName = .Range("A" & cell.row).value
            getLastGameStats.add Array(name, value, unit, toolTip, varName)
        Next cell
    End With
End Function

Private Sub btnExit_Click()
    Unload Me
    Unload boardForm
End Sub

Private Sub btnReplay_Click()
    Unload Me
    Board.replay
End Sub

Private Sub btnStartGame_Click()
    Unload Me
    Board.reset
End Sub

Private Sub UserForm_Initialize()
    populateForm
    endForm.StartUpPosition = 0
    endForm.top = boardForm.top
    endForm.left = boardForm.left + boardForm.width - 2
    ' always keep the from in view
    If (endForm.left + boardForm.width + 2) > System.width * (3 / 4) Then
        endForm.left = boardForm.left - boardForm.width - 15
    End If
End Sub
