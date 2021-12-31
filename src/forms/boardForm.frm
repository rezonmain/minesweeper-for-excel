VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} boardForm 
   Caption         =   "Minesweeper"
   ClientHeight    =   150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   1950
   OleObjectBlob   =   "boardForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "boardForm"
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
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Utils.cleanUp
End Sub
