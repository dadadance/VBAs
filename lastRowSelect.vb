Sub lastRowSub()
Dim wkR As Worksheet
Set wkR = Sheets("summary")
    wkR.Select
    lastRow = wkR.Cells(wkR.Rows.Count, "A").End(xlUp).Row + 1
End Sub
