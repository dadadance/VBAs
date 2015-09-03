Sub loopShmoop(c1 As Integer, c2 As Integer, likeType As String)
Dim wk As Worksheet
Dim sumWk As Worksheet
Dim zType As String

Application.ScreenUpdating = True

Set sumWk = Sheets("summary")

For Each wk In ThisWorkbook.Sheets
    
    If wk.Name Like "#.2" & "*" & likeType Then
        'wk.Select
        If c1 = 1 Then
            zkult = wk.Range("a2").Value
        Else
            zkult = wk.Range("g2").Value
        End If
        
        If likeType = "b" Then
            zType = "Family holdings"
        Else
            zType = "Agriculural enterprise"
        End If
        

        'Range("a4").Select
        
        For c = c1 To c2
        If wk.Range("G2").Value = "" And c1 = 7 Then GoTo HOP
            For r = 4 To 76
                'Cells(r, c).Select
                If wk.Cells(r, 8) <> "n" Then
                    If wk.Cells(r, c) <> "" Then
                        Call lastRowSub
                        'wk.Select
                        'copy paste kult1 to summ sheet
                        'If zkult = "Mulberry" Then MsgBox "Whoa!!"
                        sumWk.Cells(lastRow, 1) = zkult
                        'copy paste municipality to summ sheet
                        If wk.Cells(r, 13) <> "" Then
                            sumWk.Cells(lastRow, 3) = wk.Cells(r, 13)
                        Else
                            sumWk.Cells(lastRow, 3) = wk.Cells(r, 7)
                        End If
                        'copy paste holding type to summ sheet
                        sumWk.Cells(lastRow, 4) = zType
                        'copy paste key to summ sheet
                        sumWk.Cells(lastRow, 5) = wk.Cells(3, c)
                        'copy paste qtty to summ sheet
                        sumWk.Cells(lastRow, 6) = wk.Cells(r, c)
                        'Debug.Print zkult & " | " & wk.Cells(r, 13) & _
                        " | " & wk.Cells(3, c) & " | " & wk.Cells(r, c)
                        'sumWk.Select
                    End If
                    'wk.Select
                End If
              Next r
            Next c
    End If
HOP:
Next wk


Application.ScreenUpdating = True
End Sub
