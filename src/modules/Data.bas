Attribute VB_Name = "Data"
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

Enum DataError
    EMPTY_VALUE = 0
    NO_DATA_SHEET = 1
    FIELD_NOT_FOUND = 2
    NO_ERROR = 3
End Enum

Public Sub writeValue(field As String, value As Variant, Optional newRecord As Variant)
    ' Save the provided value to field in DATA_SHEET
    Dim row As Long, valid As DataError
    
    ' Validate the operation
    row = findRow(field)
    valid = validateOperation(field, row)
    
    Select Case valid
    Case NO_DATA_SHEET
        buildData (False)
    Case FIELD_NOT_FOUND
        buildData (True)
    End Select
    
    ' If validateOperation returns NO_ERROR write the value
    With ThisWorkbook.Sheets(DATA_SHEET)
        .Range("B" & CStr(findRow(field))).value = value
        
        ' Write 'New' on new record so stats forms can highlight it with red
        If Not IsMissing(newRecord) Then .Range("F" & CStr(findRow(field))).value = "New"
    End With
End Sub

Public Function getValue(field As String) As Variant
    ' Return the value of the provided field
    Dim row As Long, valid As DataError, res As Variant
    
    row = findRow(field)
    valid = validateOperation(field, row)
    
    Select Case valid
    Case NO_DATA_SHEET
        buildData (False)
    Case FIELD_NOT_FOUND
        buildData (True)
    End Select
    
    res = ThisWorkbook.Sheets(DATA_SHEET).Range("B" & CStr(findRow(field))).value
    
    ' If field is empty return the default value for that field
    If IsEmpty(res) Or res = "" Then
        getValue = useDefault(field)
        Exit Function
    End If
    
    getValue = res
End Function

Public Function findRow(stringToMatch As String) As Long
    ' Returns Long with row number of first rows that matches the string
    ' if string to match is not found returns 0
    
    On Error Resume Next
    Dim lRow As Long
    lRow = ThisWorkbook.Sheets(DATA_SHEET).Cells.Find(what:=stringToMatch, _
        LookAt:=xlWhole, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=True).row
    findRow = lRow
End Function

Private Function validateOperation(field As String, row As Long) As DataError
    Dim res As Variant
    
    ' If no DATA_SHEET is found
    If Not Utils.isSheet(DATA_SHEET) Then
        validateOperation = NO_DATA_SHEET
        Exit Function
    End If
    
    ' If field is not found
    If row = 0 Then
        validateOperation = FIELD_NOT_FOUND
        Exit Function
    End If
    
    validateOperation = NO_ERROR
End Function

Public Sub clearSettings()
    Data.writeValue "Theme", 0
    Data.writeValue "ShowStatsOnGameEnd", False
    Data.writeValue "RecordsOnReplay", False
    Data.writeValue "TileSize", 20
End Sub

Public Sub clearStats()
    Call Data.writeValue("gamesLost", 0)
    Call Data.writeValue("gamesWon", 0)
    Call Data.writeValue("beginnerTime", 999)
    Call Data.writeValue("intermediateTime", 999)
    Call Data.writeValue("expertTime", 999)
    Call Data.writeValue("beginner3BV/s", 0)
    Call Data.writeValue("intermediate3BV/s", 0)
    Call Data.writeValue("expert3BV/s", 0)
    Call Data.writeValue("lastGame3BV", "")
    Call Data.writeValue("lastGame3BV/s", "")
    Call Data.writeValue("lastGameTime", "")
    Call Data.writeValue("lastGameRightClicks", "")
    Call Data.writeValue("lastGameLeftClicks", "")
    Call Data.writeValue("lastGameChords", "")
    Call Data.writeValue("lastGameCompleted3BV", "")
    Call Data.writeValue("lastGameEffectiveLeftClicks", "")
    Call Data.writeValue("lastGameEffectiveRightClicks", "")
    Call Data.writeValue("lastGameEffectiveChords", "")
End Sub

Public Sub clearAllData()
    Data.writeDefualtSettings
    Data.clearStats
End Sub

Private Function useDefault(field As String) As Variant
    ' Returns the default value of the provided field
    
    Dim def As Collection, arr() As Variant
    Set def = getDefaultData
    arr = def(field)
    
    ' Value is in the second position of the array
    useDefault = arr(1)
End Function

Public Sub buildData(dataSheet As Boolean)
    ' In case there is an error getting a value
    ' rebuild the DATA_SHEET with default values
    ' if no DATA_SHEET exists, create it
    
    Dim def As Collection, arr() As Variant
    Dim i As Integer, j As Integer
    Set def = getDefaultData
    
    Application.ScreenUpdating = False
    If Not dataSheet Then
        ThisWorkbook.Sheets.add.name = DATA_SHEET
    End If
    
    With ThisWorkbook.Sheets(DATA_SHEET)
        .Range("A:A").NumberFormat = "Text"
        .Range("A:A").ColumnWidth = 28.29
        .Range("B:B").NumberFormat = "General"
        .Range("B:B").ColumnWidth = 12.29
        .Range("C:C").NumberFormat = "Text"
        .Range("C:C").ColumnWidth = 10.71
        .Range("D:D").NumberFormat = "Text"
        .Range("D:D").ColumnWidth = 18.14
        .Range("E:E").NumberFormat = "Text"
        .Range("E:E").ColumnWidth = 57.43

        For i = 1 To def.Count
            arr = def(i)
            For j = 0 To UBound(arr)
                .Cells(i, j + 1).value = arr(j)
            Next j
        Next i
        .Visible = False
    End With
    
    Application.ScreenUpdating = True
End Sub

Public Function getDefaultData() As Collection
    Set getDefaultData = New Collection
    ' varName, defualtValue, unit, displayName, description
    getDefaultData.add Array("VARNAME", "VALUE", "UNIT", "DISPLAYNAME", "DESCRIPTION"), "header"
    getDefaultData.add Array("lastFormTop", Empty, "single", "", ""), "lastFormTop"
    getDefaultData.add Array("lastFormLeft", Empty, "single", "", ""), "lastFormLeft"
    getDefaultData.add Array("lastCustomRows", Empty, "int", "", ""), "lastCustomRows"
    getDefaultData.add Array("lastCustomColumns", Empty, "int", "", ""), "lastCustomColumns"
    getDefaultData.add Array("lastCustomMines", Empty, "int", "", ""), "lastCustomMines"
    getDefaultData.add Array("PLAYER_SETTINGS", "", "", "", ""), "playerSettingsHeader"
    getDefaultData.add Array("Theme", 0, "enum", "", ""), "Theme"
    getDefaultData.add Array("ShowStatsOnGameEnd", False, "boolean", "", ""), "ShowStatsOnGameEnd"
    getDefaultData.add Array("RecordsOnReplay", False, "boolean", "", ""), "RecordsOnReplay"
    getDefaultData.add Array("TileSize", 20, "int", "", ""), "TileSize"
    getDefaultData.add Array("PLAYER_RECORDS", "", "", "", ""), "playerRecordHeader"
    getDefaultData.add Array("gamesLost", 0, "games", "Games Lost", ""), "gamesLost"
    getDefaultData.add Array("gamesWon", 0, "games", "Games Won", ""), "gamesWon"
    getDefaultData.add Array("beginnerTime", 999, "seconds", "Beginner Time", "Lowest time done in beginner difficulty"), "beginnerTime"
    getDefaultData.add Array("intermediateTime", 999, "seconds", "Intermediate Time", "Lowest time done in intermediate difficulty"), "intermediateTime"
    getDefaultData.add Array("expertTime", 999, "seconds", "Expert Time", "Lowest time done in expert difficulty"), "expertTime"
    getDefaultData.add Array("beginner3BV/s", 0, "3BV/s", "Beginner 3BV/s", "Highest 3BV/s done in beginner difficulty"), "beginner3BV/s"
    getDefaultData.add Array("intermediate3BV/s", 0, "3BV/s", "Intermediate 3BV/s", "Highest 3BV/s done in intermediate difficulty"), "intermediate3BV/s"
    getDefaultData.add Array("expert3BV/s", 0, "3BV/s", "Expert 3BV/s", "Highest 3BV/s done in expert difficulty"), "expert3BV/s"
    getDefaultData.add Array("LAST_GAME_STATS", "", "", "", ""), "lastGameStatsHeader"
    getDefaultData.add Array("lastGameDifficulty", Empty, "", "Difficulty", ""), "lastGameDifficulty"
    getDefaultData.add Array("lastGameTime", 999, "seconds", "Time", "Time passed from first click to game end"), "lastGameTime"
    getDefaultData.add Array("lastGame3BV", 0, "3BV", "3BV", "3BV is the minimun amount of clicks required to clear a board"), "lastGame3BV"
    getDefaultData.add Array("lastGameCompleted3BV", 0, "3BV", "Completed 3BV", "Completed 3BV is the amount of uncovering done on game end"), "lastGameCompleted3BV"
    getDefaultData.add Array("lastGame3BV/s", 0, "3BV/s", "3BV/s", "3BV/s is the completed 3BV over the game time"), "lastGame3BV/s"
    getDefaultData.add Array("lastGameLeftClicks", 0, "clicks", "Left Clicks", "Total amount of left clicks"), "lastGameLeftClicks"
    getDefaultData.add Array("lastGameRightClicks", 0, "clicks", "Right Clicks", "Total amount of right clicks"), "lastGameRightClicks"
    getDefaultData.add Array("lastGameChords", 0, "chords", "Chords", "Total amount of chords done"), "lastGameChords"
    getDefaultData.add Array("lastGameEffectiveLeftClicks", 0, "clicks", "Eff. Left Clicks", "Not wasted left clicks"), "lastGameEffectiveLeftClicks"
    getDefaultData.add Array("lastGameEffectiveRightClicks", 0, "clicks", "Eff. Right Clicks", "Right clicks done on a tile with a mine"), "lastGameEffectiveRightClicks"
    getDefaultData.add Array("lastGameEffectiveChords", 0, "chords", "Eff. Chords", "Not wasted chords done"), "lastGameEffectiveChords"
End Function

