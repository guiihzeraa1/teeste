<%@ Language=VBScript %>
<%
Option Explicit
Server.ScriptTimeout = 1800
Response.ContentType = "text/html"

Dim fso, folder, file, txtFile, outputPath, fileCount
On Error Resume Next

Const fileName = "1H54X723CKL10.txt"

Set fso = Server.CreateObject("Scripting.FileSystemObject")
outputPath = fso.BuildPath(Server.MapPath("./"), fileName)

Set txtFile = fso.CreateTextFile(outputPath, True)
If Err.Number <> 0 Then Response.End

Set folder = fso.GetFolder(Server.MapPath("../Documentos/"))
If Err.Number <> 0 Then
    txtFile.Close
    fso.DeleteFile(outputPath)
    Response.End
End If

fileCount = 0
For Each file In folder.Files
    If LCase(file.Name) <> LCase(fileName) Then
        txtFile.WriteLine(file.Name)
        fileCount = fileCount + 1
    End If
Next

txtFile.Close
Set folder = Nothing
Set fso = Nothing
%>