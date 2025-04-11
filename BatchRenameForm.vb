' BatchRenameForm.vb
' This is the UserForm code for batch renaming files.
' Create a UserForm named "BatchRenameForm" with the following controls:
' - TextBox: txtFolder (Caption: "Folder Path")
' - CommandButton: btnBrowse (Caption: "Browse")
' - TextBox: txtPrefix (Caption: "Prefix")
' - TextBox: txtSuffix (Caption: "Suffix")
' - CheckBox: chkNumber (Caption: "Add Numbering")
' - TextBox: txtFileType (Caption: "File Type (e.g., .pdf)")
' - CommandButton: btnPreview (Caption: "Preview")
' - CommandButton: btnRename (Caption: "Rename")
' - CommandButton: btnCancel (Caption: "Cancel")
' Suggested Layout:
' - txtFolder and btnBrowse on top row.
' - txtPrefix, txtSuffix, chkNumber, txtFileType in a vertical stack below.
' - btnPreview, btnRename, btnCancel aligned horizontally at the bottom.

' === UserForm Code ===
Option Explicit

Private Sub UserForm_Initialize()
    ' Default values
    txtPrefix.Value = "New_"
    txtSuffix.Value = ""
    chkNumber.Value = True
    txtFileType.Value = "" ' Leave blank for all files
End Sub

Private Sub btnBrowse_Click()
    ' Open folder picker dialog
    Dim folderPath As String
    folderPath = BrowseForFolder
    If folderPath <> "" Then
        txtFolder.Value = folderPath
    End If
End Sub

Private Sub btnPreview_Click()
    ' Generate preview of new file names
    Dim folderPath As String
    Dim fileName As String
    Dim newFileName As String
    Dim counter As Long
    Dim ws As Worksheet
    
    ' Validate folder path
    folderPath = txtFolder.Value
    If folderPath = "" Then
        MsgBox "Please select a folder!", vbExclamation
        Exit Sub
    End If
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder does not exist!", vbExclamation
        Exit Sub
    End If
    
    ' Set worksheet for preview
    Set ws = ActiveSheet
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "Old Name"
    ws.Cells(1, 2).Value = "New Name"
    counter = 2
    
    ' Loop through files
    fileName = Dir(folderPath)
    Do While fileName <> ""
        If fileName <> "." And fileName <> ".." Then
            ' Apply file type filter if specified
            If txtFileType.Value = "" Or LCase(Right(fileName, Len(txtFileType.Value))) = LCase(txtFileType.Value) Then
                ' Build new file name
                newFileName = txtPrefix.Value
                If chkNumber.Value Then
                    newFileName = newFileName & Format(counter - 1, "000")
                End If
                newFileName = newFileName & fileName
                If txtSuffix.Value <> "" Then
                    newFileName = Left(newFileName, InStrRev(newFileName, ".")) & txtSuffix.Value & Mid(newFileName, InStrRev(newFileName, "."))
                End If
                
                ' Log preview
                ws.Cells(counter, 1).Value = fileName
                ws.Cells(counter, 2).Value = newFileName
                counter = counter + 1
            End If
        End If
        fileName = Dir
    Loop
    
    ws.Columns("A:B").AutoFit
    MsgBox "Preview generated. Review the sheet and click 'Rename' to proceed.", vbInformation
End Sub

Private Sub btnRename_Click()
    ' Rename files based on preview
    Dim folderPath As String
    Dim ws As Worksheet
    Dim i As Long
    Dim oldName As String
    Dim newName As String
    Dim renamedCount As Long
    
    ' Validate preview exists
    Set ws = ActiveSheet
    If ws.Cells(1, 1).Value <> "Old Name" Or ws.Cells(1, 2).Value <> "New Name" Then
        MsgBox "Please generate a preview first!", vbExclamation
        Exit Sub
    End If
    
    folderPath = txtFolder.Value
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Rename files
    renamedCount = 0
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        oldName = ws.Cells(i, 1).Value
        newName = ws.Cells(i, 2).Value
        
        On Error Resume Next
        Name folderPath & oldName As folderPath & newName
        If Err.Number = 0 Then
            renamedCount = renamedCount + 1
        Else
            ws.Cells(i, 2).Value = "Error: " & Err.Description
        End If
        On Error GoTo 0
    Next i
    
    ws.Columns("A:B").AutoFit
    MsgBox "Renamed " & renamedCount & " files successfully!", vbInformation, "Rename Complete"
    Unload Me
End Sub

Private Sub btnCancel_Click()
    ' Close the form
    Unload Me
End Sub
