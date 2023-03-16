Attribute VB_Name = "Module1"
Sub Loop_through_file_names()
 Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim folderPath As String
    Dim rowNum As Long
    
    ' Get the folder path from cell F1
    folderPath = Range("F1").Value
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(folderPath)
    
    ' Start writing the file names from row 2
    rowNum = 2
    
    For Each objFile In objFolder.Files
        ' Write the file name to the current row
        Range("A" & rowNum).Value = objFile.Name
        rowNum = rowNum + 1 ' Move to the next row
    Next objFile
    
    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing
    
    MsgBox ("All files from Target folder has been looped through, files names are displayed on Column A.")

End Sub
Sub Rename_Files()


    Dim sourceFolder As String
    Dim oldName As String
    Dim newName As String
    Dim i As Integer
    
    ' Set the source folder path from cell F1
    sourceFolder = Range("F1").Value
    
    ' Loop through the rows in Column A
    For i = 2 To Range("A" & Rows.Count).End(xlUp).Row
        
        ' Get the old name from Column A and new name from Column B
        oldName = sourceFolder & "\" & Range("A" & i).Value
        newName = sourceFolder & "\" & Range("B" & i).Value
        
        ' Check if the old file exists and rename it
        If Dir(oldName) <> "" Then
            Name oldName As newName
        End If
        
        
    Next i
    
    MsgBox ("All file names has been updated to their new file names.")
    
End Sub
Sub Clean_Content()

 Range("A2:B999").ClearContents
 
 MsgBox ("Cleaned!")
 
End Sub

Sub Upload_Files_to_GDrive()

    ' Define variables
    Dim objHTTP As New MSXML2.XMLHTTP60
    Dim objFSO As New Scripting.FileSystemObject
    Dim strFolderPath As String
    Dim strBoundary As String
    Dim strRequestHeader As String
    Dim strRequestBody As String
    Dim strResponse As String
    Dim lngFileSize As Long
    Dim strParentFolderId As String
    Dim objFolder As Scripting.Folder
    Dim objFile As Scripting.File
    
    ' Set variables
    strFolderPath = Range("F1").Value ' Replace with the cell address of your folder path
    strBoundary = "----Boundary" & Format(Now, "yyyyMMddHHmmss")
    strParentFolderId = Range("F2").Value ' Replace with the cell address of your parent folder ID
    
    ' Build request header
    strRequestHeader = "Content-Type: multipart/related; boundary=" & strBoundary & vbCrLf
    strRequestHeader = strRequestHeader & "Authorization: Bearer " & GetAccessToken() & vbCrLf ' Replace with your access token
    
    ' Loop through files in folder
    Set objFolder = objFSO.GetFolder(strFolderPath)
    For Each objFile In objFolder.Files
        
        ' Build request body
        strRequestBody = "--" & strBoundary & vbCrLf
        strRequestBody = strRequestBody & "Content-Type: application/json; charset=UTF-8" & vbCrLf & vbCrLf
        strRequestBody = strRequestBody & "{""name"": """ & objFile.Name & """, ""parents"": [""" & strParentFolderId & """]}" & vbCrLf
        strRequestBody = strRequestBody & "--" & strBoundary & vbCrLf
        strRequestBody = strRequestBody & "Content-Type: " & GetMimeType(objFile.Path) & vbCrLf
        strRequestBody = strRequestBody & "Content-Transfer-Encoding: base64" & vbCrLf
        strRequestBody = strRequestBody & vbCrLf & GetFileContentBase64(objFile.Path) & vbCrLf
        strRequestBody = strRequestBody & "--" & strBoundary & "--"
        
        ' Set request options
        With objHTTP
            .Open "POST", "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", False
            .setRequestHeader "Content-Type", "multipart/related; boundary=" & strBoundary
            .setRequestHeader "Authorization", "Bearer " & GetAccessToken()
            .send strRequestBody
        End With
        
        ' Get response
        strResponse = objHTTP.responseText
        Debug.Print strResponse
        
    Next objFile

End Sub
