Attribute VB_Name = "Time"
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

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private seconds_ As Integer
Private statTimeStart As Long
Private statTimeEnd As Long

Public timer As Boolean
Public seconds As Integer

Public Sub startTimer()
    statTimeStart = GetTickCount
    timer = True
    seconds_ = -1
    updateTime
End Sub

Public Sub stopTimer()
    On Error Resume Next
    statTimeEnd = GetTickCount
    seconds = seconds_
    seconds_ = -1
    timer = False
    Application.OnTime Now + TimeValue("00:00:01"), "updateTime", Schedule:=False
End Sub

Public Sub updateTime()
    On Error Resume Next
    If timer Then
        Application.OnTime Now + TimeValue("00:00:01"), "updateTime", Schedule:=True
        seconds_ = seconds_ + 1
        Call Digits.setTimeCounter(seconds_)
    End If
End Sub

Public Function getStatSeconds() As Double
    getStatSeconds = (statTimeEnd - statTimeStart) / 1000
End Function
