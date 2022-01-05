VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MinesweeperSettings"
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

Enum theme
    Default = 0
    Dark = 1
End Enum

Private Theme_ As theme
Private ShowStatsOnGameEnd_ As Boolean
Private RecordsOnReplay_ As Boolean
Private TileSize_ As Single

Private Sub Class_Initialize()
    Theme_ = getValue("Theme")
    ShowStatsOnGameEnd_ = getValue("ShowStatsOnGameEnd")
    TileSize_ = getValue("TileSize")
    RecordsOnReplay_ = getValue("RecordsOnReplay")
End Sub

Public Property Get theme() As theme
    theme = Theme_
End Property

Public Property Get ShowStatsOnGameEnd() As Boolean
    ShowStatsOnGameEnd = ShowStatsOnGameEnd_
End Property

Public Property Get recordsOnReplay() As Boolean
    recordsOnReplay = RecordsOnReplay_
End Property

Public Property Get TileSize() As Single
    TileSize = TileSize_
End Property
