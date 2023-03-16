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


