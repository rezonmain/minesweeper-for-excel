VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} statsForm 
   Caption         =   "Stats"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4350
   OleObjectBlob   =   "statsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "statsForm"
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

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' test if data sheet exists
    getValue ("gamesLost")
    Call populateForm
End Sub

Private Sub populateForm()
    Dim playerRecords As Collection, lastGameStats As Collection, records As Dictionary
    Dim top As Single, left As Single, stat As Variant
    Dim rowSpacing As Integer, btnOffset As Integer, formOffset As Integer
    Dim nameColumnSpacing As Integer, valueColumnSpacing As Integer, fontSize As Integer
    Dim lbl As MSForms.Label
    Dim i As Integer, j As Integer
    
    fontSize = 10
    rowSpacing = fontSize + 2
    btnOffset = fontSize
    formOffset = fontSize * 3.5
    nameColumnSpacing = fontSize * 11
    valueColumnSpacing = fontSize * 5
    statsForm.frmRecords.top = 0
    
    Set playerRecords = getPlayerRecords
    Set lastGameStats = getLastGameStats
    
    Set records = stats.getNewRecords
    
    top = 5
    left = 2
    
    For i = 1 To playerRecords.Count
        stat = playerRecords(i)
        
        For j = 0 To UBound(stat) - 1
            Set lbl = statsForm.frmRecords.controls.add("Forms.Label.1")
            With lbl
                .left = left
                .top = top
                .Caption = CStr(stat(j))
                .width = 120
                .Font.Size = 10
                If j = 0 Then .ControlTipText = stat(3)
                
                ' Alternate colors
                If (i Mod 2 = 0) Then
                    .ForeColor = &H80000015
                Else
                    .ForeColor = &H80000012
                End If
                
                If records.Item(stat(4)) Then
                    .ForeColor = &HFF
                End If
            End With
            
            If j = 0 Then
                left = left + nameColumnSpacing
            Else
                left = left + valueColumnSpacing
            End If
        Next j
        left = 2
        top = top + rowSpacing
    Next i
    statsForm.frmRecords.height = (playerRecords.Count * rowSpacing) + rowSpacing
    statsForm.frmLastGame.top = frmRecords.height + 5
    
    top = 5
    left = 2
    
    For i = 1 To lastGameStats.Count
        stat = lastGameStats(i)
        
        ' Dont create a label for the tooltip
        For j = 0 To UBound(stat) - 1
            Set lbl = statsForm.frmLastGame.controls.add("Forms.label.1")
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
    statsForm.frmLastGame.height = (lastGameStats.Count * rowSpacing) + rowSpacing
    btnExit.top = frmRecords.height + frmLastGame.height + btnOffset
    statsForm.height = frmRecords.height + btnExit.height + frmLastGame.height + formOffset
End Sub


Private Function getPlayerRecords() As Collection
    Dim playerRecordsRow As Long, lastGameRow As Long
    Dim name As String, value As String, unit As String, toolTip As String, varName As String
    Dim cell As Range
    Set getPlayerRecords = New Collection
    
    playerRecordsRow = Data.findRow("PLAYER_RECORDS")
    lastGameRow = Data.findRow("LAST_GAME_STATS")
    With ThisWorkbook.Sheets(DATA_SHEET)
        For Each cell In .Range("D" & playerRecordsRow + 1 & ":D" & lastGameRow - 1)
            name = cell.value
            value = .Range("B" & cell.row).value
            unit = .Range("C" & cell.row).value
            toolTip = .Range("E" & cell.row).value
            varName = .Range("A" & cell.row).value
            getPlayerRecords.add Array(name, value, unit, toolTip, varName)
        Next cell
    End With
End Function

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
