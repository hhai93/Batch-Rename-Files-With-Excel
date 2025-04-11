' Module: BatchRename
Option Explicit

Sub BatchRename()
    ' Launch the UserForm
    With New BatchRenameForm
        .Show
    End With
End Sub

' Helper function to browse for folder
Function BrowseForFolder() As String
    Dim shell As Object
    Dim folder As Object
    
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, "Select a folder to rename files", 0)
    
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If
End Function
