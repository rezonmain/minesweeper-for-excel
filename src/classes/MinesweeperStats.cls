VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MinesweeperStats"
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

Private LeftClicks_ As Integer
Private EffectiveLeftClicks_ As Integer
Private RightClicks_ As Integer
Private EffectiveRightClicks_ As Integer
Private Chords_ As Integer
Private EffectiveChords_ As Integer
Private BBBV_ As Integer
Private CompletedBBBV_ As Integer
Private Time_ As Double
Private BBBVS_ As Double
Public Properties As Dictionary

Public Property Get LeftClicks() As Integer
    LeftClicks = LeftClicks_
End Property

Public Property Get EffectiveLeftClicks() As Integer
    EffectiveLeftClicks = EffectiveLeftClicks_
End Property

Public Property Get RightClicks() As Integer
    RightClicks = RightClicks_
End Property

Public Property Get EffectiveRightClicks() As Integer
    EffectiveRightClicks = EffectiveRightClicks_
End Property

Public Property Get Chords() As Integer
    Chords = Chords_
End Property

Public Property Get EffectiveChords() As Integer
    EffectiveChords = EffectiveChords_
End Property

Public Property Get BBBV() As Integer
    BBBV = BBBV_
End Property

Public Property Get CompletedBBBV() As Integer
    CompletedBBBV = CompletedBBBV_
End Property

Public Property Get Time() As Double
    Time = Time_
End Property

Public Property Get BBBVS() As Double
    BBBVS = BBBVS_
End Property


Public Property Let LeftClicks(value As Integer)
    LeftClicks_ = value
End Property

Public Property Let EffectiveLeftClicks(value As Integer)
    EffectiveLeftClicks_ = value
End Property

Public Property Let RightClicks(value As Integer)
    RightClicks_ = value
End Property

Public Property Let EffectiveRightClicks(value As Integer)
    EffectiveRightClicks_ = value
End Property

Public Property Let Chords(value As Integer)
    Chords_ = value
End Property

Public Property Let EffectiveChords(value As Integer)
    EffectiveChords_ = value
End Property

Public Property Let BBBV(value As Integer)
    BBBV_ = value
End Property

Public Property Let CompletedBBBV(value As Integer)
    CompletedBBBV_ = value
End Property

Public Property Let Time(value As Double)
    Time_ = value
End Property

Public Property Let BBBVS(value As Double)
    BBBVS_ = Format(value, "##0.000")
End Property

Private Sub Class_Initialize()
    LeftClicks_ = 0
    EffectiveLeftClicks_ = 0
    RightClicks_ = 0
    EffectiveRightClicks_ = 0
    Chords_ = 0
    EffectiveChords_ = 0
    BBBV_ = 0
    CompletedBBBV_ = 0
    Time_ = 999
End Sub

Public Sub setPropertiesDict()
    Set Properties = New Dictionary
    Properties.add "LeftClicks", LeftClicks_
    Properties.add "EffectiveLeftClicks", EffectiveLeftClicks_
    Properties.add "RightClicks", RightClicks_
    Properties.add "EffectiveRightClicks", EffectiveRightClicks_
    Properties.add "Chords", Chords_
    Properties.add "EffectiveChords", EffectiveChords_
    Properties.add "3BV", BBBV_
    Properties.add "Completed3BV", CompletedBBBV_
    Properties.add "Time", Time_
    Properties.add "3BV/s", BBBVS_
    Properties.add "Difficulty", Utils.getDiffName(Game.Difficulty)
End Sub

