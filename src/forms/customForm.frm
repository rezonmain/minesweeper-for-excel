VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} customForm 
   Caption         =   "Custom"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3000
   OleObjectBlob   =   "customForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "customForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnPlayCustom_Click()

    ' Validate number of mines
    If CInt(txtMines.value) > CInt(txtColumns.value) * CInt(txtRows.value) - 1 Then
        MsgBox Prompt:="Max number of mines for this configuration is " & CStr(CInt(txtColumns.value) * CInt(txtRows.value) - 1), Title:="Error"
        Exit Sub
    End If
    
    ' Save the players custom setting
    writeValue "lastCustomRows", CInt(txtRows.value)
    writeValue "lastCustomColumns", CInt(txtColumns.value)
    writeValue "lastCustomMines", CInt(txtMines.value)
    
    ' Start the custom board game
    Macros.start CInt(txtRows.value), CInt(txtColumns.value), CInt(txtMines.value), custom, Data.getValue("lastFormLeft"), Data.getValue("lastFormTop")
End Sub

Private Sub txtColumns_Change()

    ' Ignore non numeric values and limit board size to 32
    txtColumns.value = CStr(getNumeric(txtColumns.value))
    If Not txtColumns.value = "" Then
        If CInt(txtColumns.value) > MAX_BOARD_SIZE Then txtColumns.value = CStr(MAX_BOARD_SIZE)
    End If
End Sub

Private Sub txtColumns_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    ' Limit board size minimum to 9
    If Not txtColumns.value = "" Then
        If CInt(txtColumns.value) < MIN_BOARD_SIZE Then txtColumns.value = CStr(MIN_BOARD_SIZE)
    End If
End Sub

Private Sub txtRows_Change()
    txtRows.value = CStr(getNumeric(txtRows.value))
    If Not txtRows.value = "" Then
        If CInt(txtRows.value) > MAX_BOARD_SIZE Then txtRows.value = CStr(MAX_BOARD_SIZE)
    End If
End Sub

Private Sub txtRows_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not txtRows.value = "" Then
        txtRows.value = CStr(CInt(txtRows.value))
        If CInt(txtRows.value) < 9 Then txtRows.value = CStr(9)
    End If
End Sub

Private Sub txtMines_Change()

    ' Ignore non numeric values
    txtMines.value = CStr(getNumeric(txtMines.value))
    If Not txtMines.value = "" Then
        txtMines.value = CStr(CInt(txtMines.value))
    End If
End Sub

Private Sub UserForm_Initialize()
    
    ' Get the last used custom settings the player entered
    customForm.txtRows.value = Data.getValue("lastCustomRows")
    If IsEmpty(customForm.txtRows.value) Then customForm.txtRows.value = ""
    
    customForm.txtColumns.value = Data.getValue("lastCustomColumns")
    If IsEmpty(customForm.txtColumns.value) Then customForm.txtColumns.value = ""
    
    customForm.txtMines.value = Data.getValue("lastCustomMines")
    If IsEmpty(customForm.txtMines.value) Then customForm.txtMines.value = ""
    
    ' Set start up position for customForm
    customForm.StartUpPosition = 0
    customForm.top = Application.top + 148
    customForm.left = Application.left + 205
    Utils.sizeForm Me
End Sub

