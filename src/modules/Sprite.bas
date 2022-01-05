Attribute VB_Name = "Sprite"
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

Public Function useSprite(name As String) As Object
    ' Returns a picture object used for setting the -
    ' picture property of the Image control for the tiles
    
    Dim themeStr As String
    Select Case getValue("Theme")
    Case Default
        themeStr = "def"
    Case Dark
        themeStr = "drk"
    End Select
    
    Set useSprite = assetsForm.Controls(themeStr & name).Picture
End Function

Public Function settingsSprite(name As String, theme As theme) As Object
    Dim themeStr As String
    Select Case theme
    
    Case Default
        themeStr = "def"
    Case Dark
        themeStr = "drk"
    End Select
    
    Set settingsSprite = assetsForm.Controls(themeStr & name).Picture
End Function

Public Function useThisSprite(name As String) As Object
    Set useThisSprite = assetsForm.Controls(name).Picture
End Function

Public Function useSprite_(name As String) As Object
    ' Returns a picture object used for setting the -
    ' picture property of the Image control for the tiles
    
    Dim themeStr As String
    Select Case Game.Settings.theme
    Case Default
        themeStr = "def"
    Case Dark
        themeStr = "drk"
    End Select
    
    Set useSprite = ThisWorkbook.Worksheets("assets").OLEObjects(themeStr & name).Object.Picture
End Function

Public Function settingsSprite_(name As String, theme As theme) As Object
    Dim themeStr As String
    Select Case theme
    
    Case Default
        themeStr = "def"
    Case Dark
        themeStr = "drk"
    End Select
    
    Set settingsSprite = ThisWorkbook.Worksheets("assets").OLEObjects(themeStr & name).Object.Picture

End Function
