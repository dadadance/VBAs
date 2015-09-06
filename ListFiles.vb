
Sub ListFiles()

    'Set a reference to Microsoft Scripting Runtime by using
    'Tools > References in the Visual Basic Editor (Alt+F11)
    
    'Declare the variables
    Dim objFSO As Scripting.FileSystemObject
    Dim objTopFolder As Scripting.Folder
    Dim strTopFolderName As String
    
    'Insert the headers for Columns A through F
    Range("A1").Value = "File Name"
    Range("B1").Value = "File Size"
    Range("C1").Value = "File Type"
    Range("D1").Value = "Date Created"
    Range("E1").Value = "Date Last Accessed"
    Range("F1").Value = "Date Last Modified"
    Range("G1").Value = "ParentFolder"
    Range("H1").Value = "Path"
    Range("I1").Value = "ShortName"
    Range("J1").Value = "ShortPath"
    Range("K1").Value = "Attributes"
    
    
    'Assign the top folder to a variable
    strTopFolderName = "D:\Drive\FAO\ENPARD\14.00 Conditionalities\2015\Specific Conditions\2.2. Extension\02.12. Data Collection Exercise (Annex 2)\2014 (annex 12)"
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Get the top folder
    Set objTopFolder = objFSO.GetFolder(strTopFolderName)
    
    'Call the RecursiveFolder routine
    Call RecursiveFolder(objTopFolder, True)
    
    'Change the width of the columns to achieve the best fit
    Columns.AutoFit
    
End Sub

Sub RecursiveFolder(objFolder As Scripting.Folder, IncludeSubFolders As Boolean)

    'Declare the variables
    Dim objFile As Scripting.File
    Dim objSubFolder As Scripting.Folder
    Dim NextRow As Long
    
    'Find the next available row
    NextRow = Cells(Rows.count, "A").End(xlUp).Row + 1
    
    'Loop through each file in the folder
    For Each objFile In objFolder.Files
        Cells(NextRow, "A").Value = objFile.Name
        Cells(NextRow, "B").Value = objFile.Size
        Cells(NextRow, "C").Value = objFile.Type
        Cells(NextRow, "D").Value = objFile.DateCreated
        Cells(NextRow, "E").Value = objFile.DateLastAccessed
        Cells(NextRow, "F").Value = objFile.DateLastModified
        Cells(NextRow, "G").Value = objFile.ParentFolder
        Cells(NextRow, "H").Value = objFile.path
        Cells(NextRow, "I").Value = objFile.ShortName
        Cells(NextRow, "J").Value = objFile.ShortPath
        Cells(NextRow, "K").Value = objFile.Attributes
        NextRow = NextRow + 1
    Next objFile
    
    'Loop through files in the subfolders
    If IncludeSubFolders Then
        For Each objSubFolder In objFolder.SubFolders
            Call RecursiveFolder(objSubFolder, True)
        Next objSubFolder
    End If
    
End Sub


