VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MinesweeperTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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

Enum States
    Hidden = 0
    Revealed = 1
    Flagged = 2
End Enum

Private Address_ As String
Private State_ As Integer
Private Number_ As Integer
Private Mine_ As Boolean
Public Image As MSForms.Image

Public Property Get address() As String
    address = Address_
End Property

Public Property Get state() As States
    state = State_
End Property

Public Property Get Number() As Integer
    Number = Number_
End Property

Public Property Get mine() As Boolean
    mine = Mine_
End Property

Public Property Let address(value As String)
    Address_ = value
End Property

Public Property Let state(value As States)
    State_ = value
End Property

Public Property Let Number(value As Integer)
    Number_ = value
End Property

Public Property Let mine(value As Boolean)
    Mine_ = value
End Property

Private Sub Class_Initialize()
    State_ = Hidden
    Number_ = 0
    Mine_ = False
End Sub
