VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} menuForm 
   Caption         =   "Minesweeper"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4245
   OleObjectBlob   =   "menuForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "menuForm"
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

Private Sub btnBeginner_Click()
    Macros.beginner_
End Sub

Private Sub btnExit_Click()
    ThisWorkbook.Save
    Unload Me
End Sub

Private Sub btnIntermediate_Click()
    Macros.intermediate_
End Sub

Private Sub btnExpert_Click()
    Macros.expert_
End Sub

Private Sub btnCustom_Click()
    Macros.openCustomForm
End Sub

Private Sub btnStats_Click()
    Macros.openStatsForm
End Sub

Private Sub btnSettings_Click()
    Macros.openSettingsForm
End Sub


