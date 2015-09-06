'Force the explicit delcaration of variables
Option Explicit

Sub Kall()
Dim FolderPath As String, path As String, count As Integer, Filename As String
Dim paZ As String

paZ = "C:\Users\Administrator\Desktop\2014 (annex 12)\Animal Husbandry"
  
'paZ = paZ & "\*.xls"
'
'Filename = Dir(paZ)
'
'Do While Filename <> ""
'       count = count + 1
'       Debug.Print Filename
'        Filename = Dir()
'Loop
'
'Debug.Print count

Call MergeAllWorkbooks(paZ)

End Sub


Sub MergeAllWorkbooks(myPathReceived As String)
    Dim myPath As String, FilesInPath As String
    Dim MyFiles() As String
    Dim SourceRcount As Long, FNum As Long
    Dim mybook As Workbook, BaseWks As Worksheet
    Dim sourceRange As Range, destrange As Range
    Dim rnum As Long, CalcMode As Long

    ' Change this to the path\folder location of your files.
    myPath = myPathReceived
    Debug.Print "myPath: " & myPath
    
    ' Add a slash at the end of the path if needed.
    If Right(myPath, 1) <> "\" Then
        myPath = myPath & "\"
    End If

    ' If there are no Excel files in the folder, exit.
    FilesInPath = Dir(myPath & "*.xls*")
    If FilesInPath = "" Then
        MsgBox "No files found"
        Exit Sub
    End If

    ' Fill the myFiles array with the list of Excel files
    ' in the search folder.
    FNum = 0
    Do While FilesInPath <> ""
        FNum = FNum + 1
        ReDim Preserve MyFiles(1 To FNum)
        MyFiles(FNum) = FilesInPath
        FilesInPath = Dir()
    Loop

    ' Set various application properties.
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    ' Add a new workbook with one sheet.
    'Set BaseWks = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    
    'Add wksheet to end with read file name
    Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "MySheet"
    rnum = 1
    
    ' Loop through all files in the myFiles array.
    If FNum > 0 Then
        For FNum = LBound(MyFiles) To UBound(MyFiles)
            Set mybook = Nothing
            On Error Resume Next
            Set mybook = Workbooks.Open(myPath & MyFiles(FNum))
            On Error GoTo 0

            If Not mybook Is Nothing Then
                On Error Resume Next

                ' Change this range to fit your own needs.
                With mybook.Worksheets(1)
                    Set sourceRange = .Range("A1:C1")
                End With

                If Err.Number > 0 Then
                    Err.Clear
                    Set sourceRange = Nothing
                Else
                    ' If source range uses all columns then
                    ' skip this file.
                    If sourceRange.Columns.count >= ThisWorkbook.ActiveSheet.Columns.count Then
                        Set sourceRange = Nothing
                    End If
                End If
                On Error GoTo 0

                If Not sourceRange Is Nothing Then

                    SourceRcount = sourceRange.Rows.count

                    If rnum + SourceRcount >= ThisWorkbook.ActiveSheet.Rows.count Then
                        MsgBox "There are not enough rows in the target worksheet."
                        ThisWorkbook.ActiveSheet.Columns.AutoFit
                        mybook.Close savechanges:=False
                        GoTo ExitTheSub
                    Else

                        ' Copy the file name in column A.
                        With sourceRange
                            ThisWorkbook.ActiveSheet.Cells(rnum, "A"). _
                                    Resize(.Rows.count).Value = MyFiles(FNum)
                        End With

                        ' Set the destination range.
                        Set destrange = ThisWorkbook.ActiveSheet.Range("B" & rnum)

                        ' Copy the values from the source range
                        ' to the destination range.
                        With sourceRange
                            Set destrange = destrange. _
                                            Resize(.Rows.count, .Columns.count)
                        End With
                        destrange.Value = sourceRange.Value

                        rnum = rnum + SourceRcount
                    End If
                End If
                mybook.Close savechanges:=False
            End If

        Next FNum
        ThisWorkbook.ActiveSheet.Columns.AutoFit
    End If

ExitTheSub:
    ' Restore the application properties.
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub

