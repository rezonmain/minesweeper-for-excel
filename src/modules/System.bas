Attribute VB_Name = "System"
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

Private Declare PtrSafe Function GetSystemMetrics32 Lib "User32" _
    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

Public Function width() As Long
    width = GetSystemMetrics32(0) ' width in points
    
End Function

Public Function height() As Long
    height = GetSystemMetrics32(1)
End Function


