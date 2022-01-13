Attribute VB_Name = "Project"
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

' Development automatization functions
Private Sub export(workbookPath As String)
    Dim wb As Workbook
    Dim module As VBIDE.VBComponent
    Dim path As String
    Dim fs As FileSystemObject
    Set wb = Workbooks.Open(workbookPath)
    Set fs = New FileSystemObject
    
    For Each module In wb.VBProject.VBComponents
        path = ThisWorkbook.path & "\src"
        Select Case module.Type
        Case vbext_ct_ClassModule
            path = path & "\classes\" & module.name & ".cls"
            addLicence module.CodeModule
            module.export path
        
        Case vbext_ct_StdModule
            path = path & "\modules\" & module.name & ".bas"
            addLicence module.CodeModule
            module.export path
        
        Case vbext_ct_MSForm
            path = path & "\forms\" & module.name & ".frm"
            addLicence module.CodeModule
            module.export path
            
        Case Else
            GoTo Continue
        End Select
        
        
Continue:
    Next module
    wb.Save
End Sub

Private Sub loadAssetsToForm()
    ' ***VBA IS DUMB, DELETE assetsForm, THEN SAVE AND RUN***
    
    Const frmName As String = "assetsForm"
    Dim path As String, top As Single, left As Single
    Dim fs As FileSystemObject, dir As folder, file_ As file
    Dim frm As Object
    Dim comp As VBIDE.VBComponent
    Dim img As MSForms.Image
    
    With ThisWorkbook.VBProject
        Set frm = .VBComponents.add(vbext_ct_MSForm)
    End With

    With frm
        .Properties("Name") = frmName
        .Properties("Width") = 100
        .Properties("Height") = 5
        .Properties("Caption") = "_assets"
    End With
    
    top = 0
    left = 0
    path = ActiveWorkbook.path & "\assets"
    
    Set fs = New FileSystemObject
    Set dir = fs.GetFolder(path)
    
    For Each file_ In dir.Files
        If Not (fs.GetExtensionName(file_.path) = "bmp" Or fs.GetExtensionName(file_.path) = "emf") Then GoTo Continue
        Set img = frm.Designer.Controls.add("Forms.Image.1", fs.GetBaseName(file_.name), True)
        With img
            .Picture = LoadPicture(file_.path)
            .width = 5
            .height = 5
            .left = left
            .top = top
            .PictureSizeMode = fmPictureSizeModeStretch
            .BorderStyle = fmBorderStyleNone
        End With
        left = left + 5
        If left >= 100 Then
            top = top + 5
            frm.Properties("Height") = frm.Properties("Height") + 5
            left = 0
        End If
Continue:
    Next file_
End Sub

Public Sub build()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim path As String
    
    path = buildExcelFile
    export path
    removeDevModules path
    buildAddIn
    
    MsgBox "Build succesful"
End Sub

Private Function buildExcelFile() As String
    Dim path As String
    Dim cmp As VBComponent
    Dim sh As Worksheet
    Dim wb As Workbook
    Dim fs As FileSystemObject
    
    Set fs = New FileSystemObject
    
    path = ThisWorkbook.path & "\dist"
    If Not fs.FolderExists(path) Then fs.CreateFolder (path)
    
    path = ThisWorkbook.path & "\dist\Minesweeper.xlsm"
    If fs.FileExists(path) Then Kill path
    
    ThisWorkbook.SaveCopyAs path
    Set wb = Workbooks.Open(path)
    With wb
        ' Remove the data sheet
        If Utils.isSheet(DATA_SHEET) Then
            Set sh = .Worksheets(DATA_SHEET)
            sh.Delete
        End If
    End With
    wb.Save
    wb.Close
    buildExcelFile = path
    
End Function

Private Sub buildAddIn()
    Dim path As String, root As String
    Dim fil As Object
    Dim cmp As VBComponent
    Dim app As Object
    Dim fs As FileSystemObject
    Dim sh As Worksheet
    Dim wb As Workbook
    
   
    Set fs = New FileSystemObject
    
    path = ThisWorkbook.path & "\dist\Minesweeper.xlam"
    If fs.FileExists(path) Then fs.DeleteFile (path)

    ' Open 'built' excel file
    path = ThisWorkbook.path & "\dist\Minesweeper.xlsm"
    Set wb = Workbooks.Open(path)
    
    ' Save it as zip to add .rels and customUI.xml files
    path = ThisWorkbook.path & "\ribbon\ui.xlsm.zip"
    wb.SaveCopyAs path
    wb.Close
    
    root = ThisWorkbook.path
    Set app = CreateObject("Shell.Application")
    
    ' Replace .rels file
    With app.Namespace(path & "\_rels\")
        For Each fil In .Items
            If fil.name = ".rels" Then
                Call app.Namespace(CVar(root)).MoveHere(fil)
            End If
        Next fil
        Application.Wait Now + TimeValue("00:00:02")
        Kill root & "\.rels"
        .copyHere app.Namespace(root & "\ribbon\_rels\").Items
        Application.Wait Now + TimeValue("00:00:02")
    End With
    
    ' Add customUI folder and file
    With app.Namespace(path & "\")
        .copyHere app.Namespace(root & "\ribbon\customUI").Items
        Application.Wait Now + TimeValue("00:00:02")
    End With
    
    ' Remove the .zip extension
    path = ThisWorkbook.path & "\ribbon\ui.xlsm"
    Call fs.CopyFile(ThisWorkbook.path & "\ribbon\ui.xlsm.zip", path)
    
    ' Open and save as .xlam file
    Set wb = Workbooks.Open(path)
    
    ' Remove Game sheet
    wb.Worksheets.add
    wb.Worksheets("Game").Delete
    
    wb.SaveAs Filename:=ThisWorkbook.path & "\dist\Minesweeper", FileFormat:=xlOpenXMLAddIn
    wb.Save
    wb.Close
    
    ' Delete leftover file
    Kill ThisWorkbook.path & "/ribbon/ui.xlsm.zip"
End Sub

Private Sub disableAlertsOnSave(wb As Workbook)
    Dim cmp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Dim i As Long

    With wb
        Set cmp = .VBProject.VBComponents("ThisWorkbook")
        Set codeMod = cmp.CodeModule

        With codeMod
            i = .CreateEventProc("BeforeSave", "Workbook")
            i = i + 1
            .InsertLines i, "Application.DisplayAlerts = False"
            i = 4
            .InsertLines i, "Application.DisplayAlerts = False"
        End With
    End With
End Sub

Private Sub addLicence(code As VBIDE.CodeModule)
    Dim fileNum As Integer
    Dim lin As String, str As String
    Dim i As Long
    Const comment As String = "' "
    i = 1
    
    fileNum = FreeFile()
    Open ThisWorkbook.path & "/LICENCE" For Input As #fileNum
    
    While Not EOF(fileNum)
        Line Input #fileNum, lin
        With code
            str = comment & lin
            .InsertLines i, str
            i = i + 1
        End With
    Wend
End Sub

Private Sub removeDevModules(path As String)
    Dim wb As Workbook
    Dim cmp As VBComponent
    Set wb = Workbooks.Open(path)
    With wb
        ' Remove the project module, not needed in distribution build
        Set cmp = .VBProject.VBComponents("Project")
        .VBProject.VBComponents.Remove cmp
    End With
    wb.Save
    wb.Close
End Sub


