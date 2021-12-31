Attribute VB_Name = "Digits"
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

Public Sub setTimeCounter(Time As Integer)
    ' Set the digits-image to its
    ' corresponding time value
    
    Call setDigits(Time, 1)
End Sub

Public Sub setFlagCounter(NumberOfFlags As Integer)
    ' Set the flag counter digit-image
    ' to its corresponding value
    
    Call setDigits(NumberOfFlags, 0)
End Sub

Private Sub setDigits(value As Integer, counter As Integer)
    Dim numbers() As Variant
    Dim str As String
    str = CStr(value)
    
    ' Set the flag counter:
    If counter = 0 Then
        If value < 0 Then
            If value < -99 Then Call helpMe: Exit Sub
            If Len(str) <> 3 Then str = left(str, 1) & "0" & right(str, 1)
            boardForm.controls("dig3").Picture = Sprite.useSprite("n" & right(str, 1))
            boardForm.controls("dig2").Picture = Sprite.useSprite("n" & Mid(str, 2, 1))
            boardForm.controls("dig1").Picture = Sprite.useSprite("minus")
        Else
            Do While Len(str) < 3
                str = "0" & str
            Loop
            boardForm.controls("dig3").Picture = Sprite.useSprite("n" & right(str, 1))
            boardForm.controls("dig2").Picture = Sprite.useSprite("n" & Mid(str, 2, 1))
            boardForm.controls("dig1").Picture = Sprite.useSprite("n" & left(str, 1))
        End If
    End If
    
    ' Set the time counter
    If counter = 1 Then
    
        ' Preppend 0's
        Do While Len(str) < 3
            str = "0" & str
        Loop
        If counter > 999 Then Call helpMe: Exit Sub
        boardForm.controls("dig6").Picture = Sprite.useSprite("n" & left(str, 1))
        boardForm.controls("dig5").Picture = Sprite.useSprite("n" & Mid(str, 2, 1))
        boardForm.controls("dig4").Picture = Sprite.useSprite("n" & right(str, 1))
    End If
End Sub

Private Sub helpMe()
    boardForm.controls("dig1").Picture = Sprite.useSprite("nh")
    boardForm.controls("dig2").Picture = Sprite.useSprite("nl")
    boardForm.controls("dig3").Picture = Sprite.useSprite("np")
End Sub
