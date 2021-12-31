Attribute VB_Name = "Project"
Option Explicit

' Development automatization functions
Private Sub export()
    Dim module As VBIDE.VBComponent
    Dim path As String
    
    For Each module In ThisWorkbook.VBProject.VBComponents
        path = ActiveWorkbook.path & "\src"
        Select Case module.Type
        Case vbext_ct_ClassModule
            path = path & "\classes\" & module.name & ".cls"
        
        Case vbext_ct_StdModule
            path = path & "\modules\" & module.name & ".bas"
        
        Case vbext_ct_MSForm
            path = path & "\forms\" & module.name & ".frm"
            
        Case Else
            GoTo Continue
        End Select
        module.export path
        
Continue:
    Next module
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
        If Not fs.GetExtensionName(file_.path) = "bmp" Then GoTo Continue
        Set img = frm.Designer.controls.add("Forms.Image.1", fs.GetBaseName(file_.name), True)
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
    Call export
    Call buildExcelFile
    Call buildAddIn
    MsgBox "Build succesful"
End Sub

Private Sub buildExcelFile()
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
        ' Remove the project module, not needed in distribution build
        Set cmp = .VBProject.VBComponents("Project")
        .VBProject.VBComponents.Remove cmp
        
        ' Remove the data sheet
        Set sh = .Worksheets(DATA_SHEET)
        If Not sh Is Nothing Then
            sh.Delete
            
        End If
        'wb.RemoveDocumentInformation xlRDIAll
        'wb.RemovePersonalInformation = True
    End With
    'disableAlertsOnSave wb
    wb.Save
    wb.Close
    
End Sub

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


