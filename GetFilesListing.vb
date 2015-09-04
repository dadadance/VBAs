Option Explicit
'the first row with data 
Const ROW_FIRST As Integer = 5 

'This is an event handler. It exectues when the user 
'presses the run button 
Private Sub btnGet_Click()
'determines if the user selects a directory 
'from the folder dialog 
Dim intResult As Integer 
'the path selected by the user from the 
'folder dialog 
Dim strPath As String 
'Filesystem object 
Dim objFSO As Object
'the current number of rows 
Dim intCountRows As Integer 
Application.FileDialog(msoFileDialogFolderPicker).Title = _
"Select a Path"
'the dialog is displayed to the user 
intResult = Application.FileDialog( _
msoFileDialogFolderPicker).Show
'checks if user has cancled the dialog 
If intResult <> 0 Then 
    strPath = Application.FileDialog(msoFileDialogFolderPicker _ 
    ).SelectedItems(1) 
    'Create an instance of the FileSystemObject 
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    
    'loops through each file in the directory and prints their 
    'names and path 
    intCountRows = GetAllFiles(strPath, ROW_FIRST, objFSO) 
    'loops through all the files and folder in the input path 
    Call GetAllFolders(strPath, objFSO, intCountRows) 
End If 
End Sub 

''' 
'This function prints the name and path of all the files 
'in the directory strPath 
'strPath: The path to get the list of files from 
'intRow: The current row to start printing the file names 
'in 
'objFSO: A Scripting.FileSystem object. 
Private Function GetAllFiles(ByVal strPath As String, _
ByVal intRow As Integer, ByRef objFSO As Object) As Integer 
Dim objFolder As Object
Dim objFile As Object
Dim i As Integer 
i = intRow - ROW_FIRST + 1
Set objFolder = objFSO.GetFolder(strPath)
For Each objFile In objFolder.Files
        'print file name 
        Cells(i + ROW_FIRST - 1, 1) = objFile.Name 
        'print file path 
        Cells(i + ROW_FIRST - 1, 2) = objFile.Path 
        i = i + 1 
Next objFile
GetAllFiles = i + ROW_FIRST - 1
End Function

''' 
'This function loops through all the folders in the 
'input path. It makes a call to the GetAllFiles 
'function. It also makes a recursive call to itself 
'strFolder: The folder to loop through 
'objFSO: A Scripting.FileSystem object 
'intRow: The current row to print the file data on 
Private Sub GetAllFolders(ByVal strFolder As String, _
ByRef objFSO As Object, ByRef intRow As Integer) 
Dim objFolder As Object
Dim objSubFolder As Object

'Get the folder object 
Set objFolder = objFSO.GetFolder(strFolder)
'loops through each file in the directory and 
'prints their names and path 
For Each objSubFolder In objFolder.subfolders
    intRow = GetAllFiles(objSubFolder.Path, _ 
        intRow, objFSO) 
    'recursive call to to itsself 
    Call GetAllFolders(objSubFolder.Path, _ 
        objFSO, intRow) 
Next objSubFolder
End Sub 